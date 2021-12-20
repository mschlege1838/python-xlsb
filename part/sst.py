
import struct
from enum import Enum

import btypes
from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository, RecordDescriptor


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
        
        cst_total, cst_unique = struct.unpack('<II', rprocessor.read(8))
        SharedStringsPart.validate_str_count(cst_unique)
        
        items = []
        for i in range(cst_unique):
            r = rprocessor.read_descriptor()
            if r.rtype != BinaryRecordType.BrtSSTItem:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtSSTItem)
            items.append(RichStr.read(rprocessor))
        
        rprocessor.skip_until(BinaryRecordType.BrtEndSst, repository=repository)
        repository.push_current()
        
        return SharedStringsPart(cst_total, items, repository=repository)
    
    
    def __init__(self, reference_count, items, *, repository):
        SharedStringsPart.validate_str_count(reference_count)
        SharedStringsPart.validate_str_count(len(items))
        
        self.reference_count = reference_count
        self.items = items
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        items = self.items
        
        RecordDescriptor(BinaryRecordType.BrtBeginSst, 8).write(rprocessor)
        rprocessor.write(struct.pack('<I', self.reference_count))
        rprocessor.write(struct.pack('<I', len(items)))
        
        for item in items:
            RecordDescriptor(BinaryRecordType.BrtSSTItem, len(item)).write(rprocessor)
            item.write(rprocessor)
        
        if self.repository:
            self.repository.write_poll(rprocessor)
        
        RecordDescriptor(BinaryRecordType.BrtEndSst).write(rprocessor)
    
    def __getitem__(self, i):
        return self.items[i]
    

class RichStr:
    @staticmethod
    def validate_str_run_count(value):
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
            SharedStringsPart.validate_str_run_count(dw_size_str_run)
            
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
            SharedStringsPart.validate_str_run_count(dw_phonetic_run)
            
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
        
        return RichStr(_str, rgs_str_run, phonetic_str, rgs_ph_run)
    
    def __init__(self, val, runs=None, phonetic_val=None, phonetic_runs=None):
        if runs:
            SharedStringsPart.validate_str_run_count(len(runs))
        if phonetic_runs:
            SharedStringsPart.validate_str_run_count(len(phonetic_runs))
        
        self.val = val
        self.runs = runs
        self.phonetic_val = phonetic_val
        self.phonetic_runs = phonetic_runs
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        phonetic_val = self.phonetic_val
        
        flags = 0
        if self.runs:
            flags |= 0x01
        if self.phonetic_val:
            flags |= 0x02
        rprocessor.write(flags)
        
        rprocessor.write_xl_w_string(self.val, False)
        
        runs = self.runs
        if runs:
            RichStr.validate_str_run_count(len(runs))
            rprocessor.write(struct.pack('<I', len(runs)))
            for run in runs:
                run.write(rprocessor)
        
        if phonetic_val:
            rprocessor.write_xl_w_string(phonetic_val, False)
            
            phonetic_runs = self.phonetic_runs
            RichStr.validate_str_run_count(len(phonetic_runs))
            rprocessor.write(struct.pack('<I', len(phonetic_runs)))
            for run in phonetic_runs:
                run.write(rprocessor)
    
    def __len__(self):
        result = 1 + 2 * len(self.val) + 4 + (sum(len(v) for v in self.runs) if self.runs else 0)
        phonetic_val = self.phonetic_val
        if phonetic_val:
            return result + 2 * len(phonetic_val) + 4 + (sum(len(v) for v in self.phonetic_runs) if self.phonetic_runs else 0)
        else:
            return result
    
    

class StrRun:
    @staticmethod
    def read(stream):
        ich, ifnt = struct.unpack('<HH', stream.read(4))
        return StrRun(ich, ifnt)
    
    def __init__(self, start_index, font_index):
        self.start_index = start_index
        self.font_index = font_index
    
    def write(self, stream):
        stream.write(struct.pack('<H', self.start_index))
        stream.write(struct.pack('<H', self.font_index))
    
    def __len__(self):
        return 4


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
        ich_first, ich_mom, cch_mom, i_fnt, flags = struct.unpack('<HHHHH', stream.read(10))
        ph_type = flags & 0x03
        alc_h = (flags & 0x0c) >> 2
        unused_1 = flags & 0xfff0
        return PhoneticRun(ich_first, ich_mom, cch_mom, i_fnt, KatakanaType(ph_type), AlignmentType(alc_h))

    def __init__(self, ph_start_index, str_start_index, str_count, font_index, katakana_type, alignment_type):
        self.ph_start_index = ph_start_index
        self.str_start_index = str_start_index
        self.str_count = str_count
        self.font_index = font_index
        self.katakana_type = katakana_type
        self.alignment_type = alignment_type
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        rprocessor.write(struct.pack('<H', self.ph_start_index))
        rprocessor.write(struct.pack('<H', self.str_start_index))
        rprocessor.write(struct.pack('<H', self.str_count))
        rprocessor.write(struct.pack('<H', self.font_index))
        
        flags = 0
        flags |= self.katakana_type.value
        flags |= (self.alignment_type.value << 2)
        rprocessor.write(flags)
    
    def __len__(self):
        return 10
        