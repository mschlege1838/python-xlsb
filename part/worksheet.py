
import struct
from collections import deque

from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository, RecordDescriptor

class WorksheetPart
    @staticmethod
    def read(stream, for_update=False):
        rprocessor = stream if isinstance(stream, RecordProcessor) else RecordProcessor(stream)
        if for_update:
            repository = RecordRepository()
        
        
        # Begin
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginSheet:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginSheet)
        
        # Skip 1
        r = RecordDescriptor.skip_until(rprocessor, repository, BinaryRecordType.BrtWsDim, BinaryRecordType.BeginColInfos, BinaryRecordType.BrtBeginSheetData, push_current=False)
        if r.rtype == BinaryRecordType.BrtWsDim:
            r = RecordDescriptor.skip_until(BinaryRecordType.BeginColInfos, BinaryRecordType.BrtBeginSheetData)
        repository.push_current()
        
        # Col Info
        col_info = []
        if r.rtype == BinaryRecordType.BrtBeginColInfos:
            r = rprocessor.read_descriptor()
            while r.rtype == BinaryRecordType.BrtColInfo:
                col_info.append(ColInfo.read(rprocessor))
                r = r.read_descriptor()
            if r.rtype != BinaryRecordType.BrtEndColInfos:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtEndColInfos)
        
        # Skip 3
        r = RecordDescriptor.skip_until(rprocessor, repository, BinaryRecordType.BrtBeginSheetData)
        repository.push_current()
        
        
        
        

class ColInfo:
    @staticmethod
    def read(rprocessor):
        col_first = struct.unpack('<i', rprocessor.read(4))[0]
        col_last = struct.unpack('<i', rprocessor.read(4))[0]
        col_idx = struct.unpack('<I', rprocessor.read(4))[0]
        ixfe = struct.unpack('<I', rprocessor.read(4))[0]
        
        flags_1 = rprocessor.read(1)
        f_hidden = flags_1 & 0x80
        f_user_set = flags_1 & 0x40
        f_best_fit = flags_1 & 0x20
        f_phonetic = flags_1 & 0x10
        reserved_1 = flags_1 & 0x0f
        
        flags_2 = rprocessor.read(1)
        i_out_level = (flags_2 & 0xe0) >> 5
        unused = flags_2 & 0x10
        f_collapsed = flags_2 & 0x08
        reserved_2 = flags_2 & 0x07
        
        return ColInfo(col_first, col_last, col_idx, ixfe, bool(f_hidden), bool(f_user_set), bool(f_best_fit), 
                bool(f_phonetic), i_out_level, bool(f_collapsed))
    
    def __init__(self, col_first, col_last, std_digit_width, style_index, hidden, user_width_set, best_fit_set,
            phonetic, outline_level, outline_collapsed):
        self.col_first = col_first
        self.col_first = col_first
        self.col_last = col_last
        self.std_digit_width = std_digit_width
        self.style_index = style_index
        self.hidden = hidden
        self.user_width_set = user_width_set
        self.best_fit_set = best_fit_set
        self.phonetic = phonetic
        self.outline_level = outline_level
        self.outline_collapsed = outline_collapsed