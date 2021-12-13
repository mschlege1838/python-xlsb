
import os
import io
import zipfile
import xml.etree.ElementTree as ET

from enum import Enum
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED
from collections import deque

from btypes import RelationshipType, ContentType


def rec_zip(fname, src_dir):
    zf = ZipFile(fname, 'w', ZIP_DEFLATED, compresslevel=6)
    
    if not str(src_dir).endswith(os.sep):
        src_dir = str(src_dir) + os.sep
    
    def rec_helper(cur):
        for e in os.scandir(cur):
            if e.is_dir():
                rec_helper(e)
            else:
                with zf.open(e.path.replace(str(src_dir), ''), 'w') as dest, open(e, 'rb') as src:
                    dest.write(src.read())
    
    rec_helper(src_dir)
    zf.close()
    

def norm_path(path, return_segs=False, return_path=True, leading_slash=True):
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
    
    if return_segs and not return_path:
        return psegs
    
    path = f"/{'/'.join(psegs)}" if leading_slash else '/'.join(psegs)
    
    if return_path and return_segs:
        return path, psegs
    else:
        return path


class XMLNSName(Enum):
    RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types'


class PartInfo:
    def __init__(self, path, content_type, relationships):
        self.path = path
        self.content_type = content_type
        self.relationships = relationships
    
    def get_rel(self, by_name, value):
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
    
    def __str__(self):
        return f'{self.path} ({self.content_type})'

class PartRelationship:
    def __init__(self, rid, rtype, target, raw_target):
        self.rid = rid
        self.rtype = rtype
        self.target = target
        self.raw_target = raw_target
    
    def __str__(self):
        return f'{self.rid}: {self.rtype} -> {self.target}'
    

class PartRelationshipsPart:
    
    @staticmethod
    def resolve_query(by_name, value):
        if by_name == 'rid':
            by_name = 'Id'
        elif by_name == 'rtype':
            by_name = 'Type'
        elif by_name == 'target':
            by_name = 'Target'
        
        if by_name not in ('Id', 'Type', 'Target'):
            raise ValueError(f'Relationship attributes are Id, Type, and Target: {value}')
        
        if hasattr(value, 'value'):
            value = value.value
        
        return by_name, value
    
    @staticmethod
    def resolve_rel(root, el):
        raw_target = el.get('Target')
        if raw_target.startswith('/') or not root:
            target = raw_target
        elif root == '/':
            target = f'/{raw_target}'
        else:
            target = f'{root}/{raw_target}'
        return PartRelationship(el.get('Id'), RelationshipType.resolve(el.get('Type')), target, raw_target)

    
    def __init__(self, root, xtree, file_handle=None):
        self.root = root
        self.xtree = xtree
        self.file_handle = file_handle
    
    
    def get_rel(self, by_name, value):
        by_name, value = PartRelationshipsPart.resolve_query(by_name, value)
        ns = {'': XMLNSName.RELATIONSHIPS.value}
        
        el = self.xtree.find(f'.//Relationship[@{by_name}="{value}"]', ns)
        if el is None:
            return None
        
        return PartRelationshipsPart.resolve_rel(self.root, el)
    
    def remove_rel(self, by_name, value):
        by_name, value = PartRelationshipsPart.resolve_query(by_name, value)
        ns = {'': XMLNSName.RELATIONSHIPS.value}
        
        xtree = self.xtree
        el = xtree.find(f'.//Relationship[@{by_name}="{value}"]', ns)
        if el is not None:
            xtree.getroot().remove(el)
    
    def add_rel(self, rtype, target, rid=None):
        ns = {'': XMLNSName.RELATIONSHIPS.value}
        
        xtree = self.xtree
        
        if rid is None:
            i = 1
            while True:
                if xtree.find(f'.//Relationship[@Id="rId{i}"]', ns) is None:
                    break
                i += 1
            rid = f'rId{i}'
        
        if hasattr(rtype, 'value'):
            rtype = rtype.value
        
        ET.SubElement(xtree.getroot(), f'{{{XMLNSName.RELATIONSHIPS.value}}}Relationship', {'Id': rid, 'Type': rtype, 'Target': target})
        return rid
    
    def close(self):
        file_handle = self.file_handle
        if not file_handle:
            return
            
        ET.register_namespace('', XMLNSName.RELATIONSHIPS.value)
        self.xtree.write(file_handle, 'UTF-8', True)
        file_handle.close()
    
    
    def __iter__(self):
        ns = {'': XMLNSName.RELATIONSHIPS.value}
        root = self.root
        for el in self.xtree.iterfind('.//Relationship', ns):
            yield PartRelationshipsPart.resolve_rel(root, el)
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
    
    
    

class ZipOfficeOpenXMLPackage:
    def __init__(self, fname, overwrite=False):
        self.fname = fname
        self.extract_dir = None
        self.on_close_hooks = None
        
        exists = True
        try:
            os.stat(fname)
        except FileNotFoundError:
            exists = False
        
        if not exists or overwrite:
            extract_dir = self.extract_dir = TemporaryDirectory()
            with open(os.path.join(extract_dir.name, '[Content_Types].xml'), 'w', encoding='utf-8') as f:
                f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>')

    def register_on_close_hook(self, closeable):
        on_close_hooks = self.on_close_hooks
        if on_close_hooks is None:
            on_close_hooks = self.on_close_hooks = []
        on_close_hooks.add(closeable)
    
    def get_part_info(self, path=None):
        path, psegs = norm_path(path, True, True)
        
        if path == '/':
            try:
                with self.open_part_relationships(path) as rels:
                    return PartInfo('/', None, iter(rels))
            except FileNotFoundError:
                return PartInfo('/', None, [])
        else:
            if not self.exists(path):
                return None
            
            content_type = self._determine_content_type(path)
            
            try:
                with self.open_part_relationships(path) as rels:
                    return PartInfo(path, content_type, list(rels))
            except FileNotFoundError:
                return PartInfo(path, content_type, [])
    
    
    def exists(self, path):
        path = norm_path(path, leading_slash=False)
        
        if self.extract_dir:
            try:
                os.stat(os.path.join(self.extract_dir.name, path.replace('/', os.sep)))
            except FileNotFoundError:
                return False
            else:
                return True
        else:
            with ZipFile(self.fname) as f:
                return path in f.namelist()
    
    def open_part_relationships(self, path=None, mode='r'):
        psegs = norm_path(path, True, False)
        
        root = f"/{'/'.join(list(psegs)[:-1])}"
        
        if len(psegs):
            psegs.insert(-1, '_rels')
            psegs[-1] = f'{psegs[-1]}.rels'
            rel_path = '/'.join(psegs)
        else:
            rel_path = '_rels/.rels'
        
        if not self.exists(rel_path) and mode == 'w':
            with self.open_part(rel_path, 'w') as f:
                f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        
        with self.open_part(rel_path) as f:
            xtree = ET.parse(f)
        
        return PartRelationshipsPart(root, xtree, self.open_part(rel_path, 'w') if mode == 'w' else None)
    
    def open_part(self, path, mode='r', content_type=None, update_content_type=False):
        path = norm_path(path, leading_slash=False)
        if content_type and hasattr(content_type, 'value'):
            content_type = content_type.value
        
        if mode != 'r' and mode != 'w':
            raise ValueError(f'Mode must be "r" or "w": {mode}')
        
        exists = self.exists(path)
        if mode == 'r' and not exists:
            raise FileNotFoundError(f'Item not contained in package: {path}')
        
        if mode == 'w':
            if not content_type:
                content_type = self._determine_content_type(path, False)
            
            if not exists and not content_type and not path.endswith('.rels'):
                raise ValueError(f'Content Type must be provided with new paths: {path}')
            if path != '[Content_Types].xml' and not path.endswith('.rels'):
                self._add_content_type(path, content_type, update_content_type)
        
        if mode == 'w' and not self.extract_dir:
            self._extract_temp()
        
        if self.extract_dir:
            target = os.path.join(self.extract_dir.name, path.replace('/', os.sep))
            target_dir = os.path.split(target)[0]
            if target_dir:
                os.makedirs(target_dir, exist_ok=True)
            return open(target, f'{mode}b')
        else:
            with ZipFile(self.fname, mode) as f:
                return f.open(path)
    
    
    def close(self):
        on_close_hooks = self.on_close_hooks
        if on_close_hooks:
            for closeable in on_close_hooks:
                closeable.close()
        
        extract_dir = self.extract_dir
        if extract_dir:
            rec_zip(self.fname, extract_dir.name)
            extract_dir.cleanup()
    
    def _determine_content_type(self, path, err_if_none=True):
        ns = {'': XMLNSName.CONTENT_TYPES.value}
        path = path.casefold()
        
        # Open Media Stream
        try:
            xtree = ET.parse(self.open_part('[Content_Types].xml'))
        except FileNotFoundError:
            raise ValueError(f'Unable to determine content type: {path} (Media stream does not exist)')
        
        # Check Overrides
        for override in xtree.iterfind('.//Override', ns):
            if override.get('PartName').casefold() == path:
                return ContentType.resolve(override.get('ContentType'))
        
        # Check Defaults
        ext_i = path.rfind('.')
        if ext_i == -1:
            raise ValueError(f'Unable to determine content type: {path} (No extension, no override)')
        ext = path[ext_i + 1:]
        
        for default in xtree.iterfind('.//Default', ns):
            if default.get('Extension').casefold() == ext:
                return ContentType.resolve(default.get('ContentType'))
            
        if err_if_none:
            raise ValueError(f'Unable to determine content type: {path} (No matching override or default)')
        else:
            return None
    
    def _add_content_type(self, path, content_type, update_override=False):
        ns = {'': XMLNSName.CONTENT_TYPES.value}
        ET.register_namespace('', XMLNSName.CONTENT_TYPES.value)
        path = f'/{path.casefold()}'
        content_type = (content_type.value if hasattr(content_type, 'value') else content_type)
        if content_type:
            content_type = content_type.casefold()
        
        
        # Open Media Stream
        try:
            xtree = ET.parse(self.open_part('[Content_Types].xml'))
        except FileNotFoundError:
            raise ValueError(f'Unable to determine content type: {path} (Media stream does not exist)')
        
        # Defaults
        ext_i = path.rfind('.')
        if ext_i != -1:
            ext = path[ext_i + 1:]
            for default in xtree.iterfind('.//Default', ns):
                if default.get('Extension').casefold() == ext:
                    # If extension and content_type match an existing Default, nothing to do.
                    if default.get('ContentType').casefold() == content_type:
                        return
                    # If no match, add Override
                    break
            else:
                # If no Default element matches the extension, add Default with extension.
                ET.SubElement(xtree.getroot(), f'{{{XMLNSName.CONTENT_TYPES.value}}}Default', {'Extension': ext, 'ContentType': content_type})
                with self.open_part('[Content_Types].xml', 'w') as f:
                    xtree.write(f, 'UTF-8', True)
                return
        
        
        # Overrides
        for override in xtree.iterfind('.//Override', ns):
            if override.get('PartName').casefold() == path:
                # If override already exists with same path, but different content type, error or update.
                if override.get('ContentType').casefold() != content_type:
                    if update_override:
                        override.set('ContentType', content_type)
                        with self.open_part('[Content_Types].xml', 'w') as f:
                            xtree.write(f, 'UTF-8', True)
                        return
                    else:
                        raise ValueError(f'An override with the path {path} already exists with a different ContentType: {override.get("ContentType")}; content_type: {content_type}')
                # Otherwise, noting to do.
                return
        
        # Add Override if nothing else above matches.
        ET.SubElement(xtree.getroot(), f'{{{XMLNSName.CONTENT_TYPES.value}}}Override', {'PartName': path, 'ContentType': content_type})
        with self.open_part('[Content_Types].xml', 'w') as f:
            xtree.write(f, 'UTF-8', True)
    
    
    def _extract_temp(self):
        if self.extract_dir:
            return
        
        extract_dir = self.extract_dir = TemporaryDirectory()
        with ZipFile(self.fname, 'r') as f:
            f.extractall(extract_dir.name)
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
        