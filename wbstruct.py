
from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException


class WorkbookStruct:
    @staticmethod
    def read(rprocessor):
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginBook:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginBook)
        
        while r.rtype != BinaryRecordType.BrtBeginBundleShs:
            r.skip()
            r = rprocessor.read_descriptor()
        
        sheet_refs = []
        r = rprocessor.read_descriptor()
        while r.rtype == BinaryRecordType.BrtBundleSh:
            sheet_refs.append(BundledSheet.read(rprocessor))
            r = rprocessor.read_descriptor()
        
        if r.rtype != BinaryRecordType.BrtEndBundleShs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtEndBundleShs)
        
        return WorkbookStruct(sheet_refs)
    
    def __init__(self, sheet_refs):
        self.sheet_refs = sheet_refs


class BundledSheet:
    
    @staticmethod
    def read(rprocessor):
        hs_state = struct.unpack('<I', rprocessor.read(4))[0]
        i_tab_id = struct.unpack('<I', rprocessor.read(4))[0]
        rel_id = rprocessor.read_xl_nullable_w_string()
        sheet_name = rprocessor.read_xl_nullable_w_string()
    
    def __init__(self, hidden, tab_id, rel_id, sheet_name):
        self.hidden = hidden
        self.tab_id = tab_id
        self.rel_id = rel_id
        self.sheet_name = sheet_name
    
