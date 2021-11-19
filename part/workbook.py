
import struct
from enum import Enum

from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor


class WorkbookPart:
    @staticmethod
    def read(stream):
        rprocessor = stream if isinstance(stream, RecordProcessor) else RecordProcessor(stream)
    
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginBook:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginBook)
        
        while r.rtype != BinaryRecordType.BrtBeginBundleShs:
            r.skip(rprocessor)
            r = rprocessor.read_descriptor()
        
        sheet_refs = []
        r = rprocessor.read_descriptor()
        while r.rtype == BinaryRecordType.BrtBundleSh:
            sheet_refs.append(BundledSheet.read(rprocessor))
            r = rprocessor.read_descriptor()
        
        if r.rtype != BinaryRecordType.BrtEndBundleShs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtEndBundleShs)
        
        return WorkbookPart(sheet_refs)
    
    def __init__(self, sheet_refs):
        self.sheet_refs = sheet_refs

class HiddenState(Enum):
    VISIBLE = 0
    HIDDEN = 1
    VERY_HIDDEN = 2
    
class BundledSheet:
    
    @staticmethod
    def read(rprocessor):
        hs_state = struct.unpack('<I', rprocessor.read(4))[0]
        i_tab_id = struct.unpack('<I', rprocessor.read(4))[0]
        rel_id = rprocessor.read_xl_nullable_w_string()
        sheet_name = rprocessor.read_xl_nullable_w_string()
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
        self.rel_id = rel_id
        BundledSheet.check_sheet_name(sheet_name)
        self._sheet_name = sheet_name
    
    @property
    def sheet_name(self):
        return self._sheet_name
    
    @sheet_name.setter
    def sheet_name(self, value):
        BundledSheet.check_sheet_name(value)
        self._sheet_name = value