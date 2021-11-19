
import os
import io
import zipfile
import xml.etree.ElementTree as ET

from enum import Enum
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from collections import deque


from btypes import RelationshipType, ContentType, XMLNSName


def norm_path(path, return_segs=False, return_path=True):
    if not path:
        path = ''
    elif isinstance(path, PartRelationship):
        path = path.target
    elif isinstance(path, PartInfo):
        path = path.path

    psegs = deque()
    for seg in path.split('/'):
        if not seg or seg == '.':
            continue
        elif seg == '..':
            if len(psegs):
                psegs.pop()
        else:
            psegs.append(seg)
    if return_path and return_segs:
        return '/'.join(psegs), psegs
    else:
        return psegs if return_segs else '/'.join(psegs)


class PartInfo:
    def __init__(self, path, content_type, relationships):
        self.path = path
        self.content_type = content_type
        self.relationships = relationships
    
    def get_relationship(self, by_name, value):
        if by_name == 'Id':
            by_name = 'rid'
        elif by_name == 'Type':
            by_name = 'rtype'
        elif by_name == 'Target':
            by_name = 'target'
        
        if hasattr(value, 'value'):
            value = value.value
        
        for r in self.relationships:
            v = getattr(r, by_name)
            if (v.value if hasattr(v, 'value') else v) == value:
                return r
        
        return None


class PartRelationship:
    
    @staticmethod
    def from_xml(stream):
        ns = {'': XMLNSName.RELATIONSHIPS.value}
        
        root = ET.parse(stream)
        result = []
        for el in root.findall('.//Relationship', ns):
            result.append(PartRelationship(el.get('Id'), RelationshipType.resolve(el.get('Type')), el.get('Target')))
        return result
    
    def __init__(self, rid, rtype, target):
        self.rid = rid
        self.rtype = rtype
        self.target = target
    


class ZipOfficeOpenXMLPackage:
    def __init__(self, fname):
        self.fname = fname
        self.extract_dir = None
    
    
    def get_part_info(self, path=None):
        path, psegs = norm_path(path)
        
        if not path:
            try:
                return PartInfo('/', None, PartRelationship.from_xml(self.open_part('_rels/.rels')))
            except FileNotFoundError:
                return PartInfo('/', None, [])
        else:
            if not self.exists(path):
                return None
            
            content_type = self.determine_content_type(path)
            
            psegs.insert(-1, '_rels')
            psegs[-1] = f'{psegs[-1]}.rels'
            rel_path = '/'.join(psegs)
            try:
                return PartInfo(path, content_type, PartRelationship.from_xml(self.open_part(rel_path)))
            except FileNotFoundError:
                return PartInfo(part, content_type, [])
    
    def determine_content_type(self, path):
        path = norm_path(path.casefold())
        
        # Open Media Stream
        try:
            root = ET.parse(self.open_part('[Content_Types].xml'))
        except FileNotFoundError:
            raise ValueError(f'Unable to determine content type: {path} (Media stream does not exist)')
        
        ns = {'': XMLNSName.CONTENT_TYPES.value}
        
        # Check Overrides
        for override in root.iterfind('.//Override', ns):
            if override.get('PartName').casefold() == path:
                return ContentType.resolve(override.get('ContentType'))
        
        # Check Defaults
        ext_i = path.rfind('.')
        if ext_i == -1:
            raise ValueError(f'Unable to determine content type: {path} (No extension, no override)')
        ext = path[ext_i + 1:]
        
        for default in root.iterfind('.//Default', ns):
            if default.get('Extension').casefold() == ext:
                return ContentType.resolve(default.get('ContentType'))
                
        raise ValueError(f'Unable to determine content type: {path} (No matching override or default)')
    
    
    def exists(self, path):
        path = norm_path(path)
        
        if self.extract_dir:
            try:
                os.stat(os.path.join(self.extract_dir, path.replace('/', os.sep)))
            except FileNotFoundError:
                return False
            else:
                return True
        else:
            try:
                with ZipFile(self.fname) as f:
                    return path in f.namelist()
            except FileNotFoundError:
                return False
    
    def open_part(self, path, mode='r'):
        path = norm_path(path)
        
        if self.extract_dir:
            return open(os.path.join(self.extract_dir, path.replace('/', os.sep)), mode)
        elif mode in ('w', 'a', 'x'):
            self.extract_temp()
            return open(os.path.join(self.extract_dir, path.replace('/', os.sep)), mode)
        else:
            with ZipFile(self.fname, mode) as f:
                try:
                    return f.open(path)
                except KeyError:
                    raise FileNotFoundError(f'Item not contained in archive: {path}')
    
    
    def extract_temp(self):
        if self.extract_dir:
            return
        
        extract_dir = self.extract_dir = TemporaryDirectory()
        try:
            with ZipFile(self.fname, 'r') as f:
                f.extractall(extract_dir)
        except FileNotFoundError:
            pass
    
    def close(self):
        if self.extract_dir:
            self.extract_dir.cleanup()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
        