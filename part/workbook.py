
import struct
from enum import Enum

from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository, RecordDescriptor


class WorkbookPart:
    @staticmethod
    def read(stream, for_update=False):
        rprocessor = RecordProcessor.resolve(stream)
        
        # Begin
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginBook:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginBook)
        
        repository = RecordRepository(for_update)
        
        # Skip 1
        r = rprocessor.skip_until(BinaryRecordType.BrtBeginBundleShs, repository=repository)
        repository.push_current()
        
        # Sheet References
        sheet_refs = []
        r = rprocessor.read_descriptor()
        while r.rtype == BinaryRecordType.BrtBundleSh:
            sheet_refs.append(BundledSheet.read(rprocessor))
            r = rprocessor.read_descriptor()
        
        if r.rtype != BinaryRecordType.BrtEndBundleShs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtEndBundleShs)
        
        
        # Skip 2
        rprocessor.skip_until(BinaryRecordType.BrtEndBook, repository=repository)
        repository.push_current()
        
        return WorkbookPart(sheet_refs, repository=repository)
    
    def __init__(self, sheet_refs, *, repository):
        self.sheet_refs = sheet_refs
        self.repository = repository
    
    def write(self, stream):
        repository = self.repository
        rprocessor = RecordProcessor.resolve(stream)
        
        # Begin
        RecordDescriptor(BinaryRecordType.BrtBeginBook).write(rprocessor)
        
        # Skip 1
        repository.write_poll(rprocessor)
        
        # Sheet References
        RecordDescriptor(BinaryRecordType.BrtBeginBundleShs).write(rprocessor)
        for sheet_ref in self.sheet_refs:
            RecordDescriptor(BinaryRecordType.BrtBundleSh, len(sheet_ref)).write(rprocessor)
            sheet_ref.write(rprocessor)
        RecordDescriptor(BinaryRecordType.BrtEndBundleShs).write(rprocessor)
        
        # Skip 2
        repository.write_poll(rprocessor)
        
        # End
        RecordDescriptor(BinaryRecordType.BrtEndBook).write(rprocessor)
    
    def __len__(self):
        result = len(self.repository)
        for sheet_ref in self.sheet_refs:
            result += len(sheet_ref)
        return result
        
    

class HiddenState(Enum):
    VISIBLE = 0
    HIDDEN = 1
    VERY_HIDDEN = 2
    
class BundledSheet:
    
    @staticmethod
    def read(rprocessor):
        hs_state = struct.unpack('<I', rprocessor.read(4))[0]
        i_tab_id = struct.unpack('<I', rprocessor.read(4))[0]
        rel_id = rprocessor.read_xl_w_string()
        sheet_name = rprocessor.read_xl_w_string(False)
        return BundledSheet(HiddenState(hs_state), i_tab_id, rel_id, sheet_name)
    
    @staticmethod
    def check_sheet_name(name):
        for reserved in ('\0', '\u0003', ':', '\\', '*', '?', '/', '[', ']'):
            if reserved in name:
                raise ValueError(f'Sheet name cannot contain "{reserved}": {name}')
        if name.startswith("'") or name.endswith("'"):
            raise ValueError(f'Sheet name cannot start or end with "\'": {name}')
    
    def __init__(self, hidden_state, tab_id, rel_id, sheet_name):
        self.hidden_state = hidden_state
        self.tab_id = tab_id
        RecordProcessor.check_rel_id(rel_id)
        self._rel_id = rel_id
        BundledSheet.check_sheet_name(sheet_name)
        self._sheet_name = sheet_name
    
    @property
    def rel_id(self):
        return self._rel_id
    
    @rel_id.setter
    def rel_id(self, value):
        RecordProcessor.check_rel_id(value)
        self._rel_id = value
    
    @property
    def sheet_name(self):
        return self._sheet_name
    
    @sheet_name.setter
    def sheet_name(self, value):
        BundledSheet.check_sheet_name(value)
        self._sheet_name = value
    
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        rprocessor.write(struct.pack('<I', self.hidden_state.value))
        rprocessor.write(struct.pack('<I', self.tab_id))
        rprocessor.write_xl_w_string(self._rel_id)
        rprocessor.write_xl_w_string(self._sheet_name, False)
    
    def __len__(self):
        return 8 + RecordProcessor.len_xl_w_string(self.rel_id) + RecordProcessor.len_xl_w_string(self.sheet_name, False)