
import struct
from enum import Enum

import btypes
from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository


class SharedStringsPart:
    @staticmethod
    def validate_str_count(value):
        if value > 0x7fffffff:
            raise ValueError(f'Shared string counts must be less than or equal to {0x7fffffff}: {value}')
    
    @staticmethod
    def read(stream, for_update=False):
        rprocessor = RecordProcessor.resolve(stream)
        repository = RecordRepository(for_update)
        
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginSst:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginSst)
        
        cst_total = struct.unpack('<I', rprocessor.read(4))
        cst_unique = struct.unpack('<I', rprocessor.read(4))
        SharedStringsPart.validate_str_count(cst_unique)
        
        items = []
        for i in range(cst_unique):
            r = rprocessor.read_descriptor()
            if r.rtype != BinaryRecordType.BrtSSTItem:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtSSTItem)
            items.append(RichStr.read(rprocessor))
        
        rprocessor.skip_until(BinaryRecordType.BrtEndSst, repository=repository)
        
        return SharedStringsPart(cst_total, items, repository=repository)
    
    
    def __init__(self, referencce_count, items, *, repository):
        SharedStringsPart.validate_str_count(referencce_count)
        SharedStringsPart.validate_str_count(len(items))
        
        self.referencce_count = referencce_count
        self.items = items
    

class RichStr:
    @staticmethod
    def validate_str_run(value):
        if value > 0x7fff:
            raise ValueError(f'String run count must be less than or equal to {0x7fff}: {value}')
    
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        flags_1 = rprocessor.read(1)
        f_rich_str = flags_1 & 0x01
        f_ext_str = flags_1 & 0x02
        unused_1 = flags_1 & 0xfc
        
        _str = rprocessor.read_xl_w_string(False)
        
        rgs_str_run = None
        if f_rich_str:
            dw_size_str_run = struct.unpack('<I', rprocessor.read(4))[0]
            SharedStringsPart.validate_str_run(dw_size_str_run)
            
            rgs_str_run = []
            prev = None
            for i in range(dw_size_str_run):    
                run = StrRun.read(rprocessor)
                if run.start_index >= len(_str):
                    raise ValueError(f'String run start index must be less than the number of characters in the string: {run.start_index}'
                                        f'; Len={len(_str)}; Val={_str}')
                if prev and run.start_index <= prev.start_index:
                    raise ValueError(f'Each string run must have a start index greater than the previous: {run.start_index}'
                                        f'; Prev={previous.start_index}; Run Index={i}')
                rgs_str_run.append(run)
                prev = run
        
        
        phonetic_str = None
        rgs_ph_run = None
        if f_ext_str:
            phonetic_str = rprocessor.read_xl_w_string(False)
            
            dw_phonetic_run = struct.unpack('<I', rprocessor.read(4))[0]
            SharedStringsPart.validate_str_run(dw_phonetic_run)
            
            rgs_ph_run = []
            prev = None
            for i in range(dw_phonetic_run):
                run = PhoneticRun.read(rprocessor)
                if run.ph_start_index > len(phonetic_str):
                    raise ValueError(f'Phonetic string run start index must be less than the numbers of characters in the phonetic string: {run.ph_start_index}'
                                        f'; Len={len(phonetic_str)}; Val={phonetic_str}')
                if prev and run.ph_start_index <= prev.ph_start_index:
                    raise ValueError(f'Each phonetic string run must have a start index greater than the previous: {run.ph_start_index}'
                                        f'; Prev={previous.ph_start_index}; Run Index={i}')
        
        return RichStr
    
    def __init__(self, val, runs=None, phonetic_val=None, phonetic_runs=None):
        self.val = val
        self.runs = runs
        self.phonetic_val = phonetic_val
        self.phonetic_runs = phonetic_runs
    

class StrRun:
    @staticmethod
    def read(stream):
        ich, ifnt = struct.unpack('<H', stream.read(4))
        return StrRun(ich, ifnt)
    
    def __init__(self, start_index, font_index):
        self.start_index = start_index
        self.font_index = font_index


class KatakanaType(Enum):
    NARROW = 0x00
    WIDE = 0x01
    HIRAGANA = 0x02
    AS_ENTERED = 0x03

class AlignmentType(Enum):
    LEFT_ALL = 0x00
    LEFT_EACH = 0x01
    CENTER_EACH = 0x02
    JUSTIFIED = 0x03

class PhoneticRun:
    @staticmethod
    def read(stream):
        ich_first, ich_mom, cch_mom, i_fnt, flags = struct.unpack('<H', stream.read(10))
        ph_type = flags & 0x03
        alc_h = flags & 0x0c >> 2
        unused_1 = flags & 0xfff0
        return PhoneticRun(ich_first, ich_mom, cch_mom, i_fnt, KatakanaType(ph_type), AlignmentType(alc_h))

    def __init__(self, ph_start_index, str_start_index, str_count, font_index, katakana_type, alignment_type):
        self.ph_start_index = ph_start_index
        self.str_start_index = str_start_index
        self.str_count = str_count
        self.font_index = font_index
        self.katakana_type = katakana_type
        self.alignment_type = alignment_type
        
        