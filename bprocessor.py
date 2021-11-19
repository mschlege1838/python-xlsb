
import io
import struct
from collections import deque
from enum import Enum, auto

from btypes import BinaryRecordType, FutureRecordType, AlternateContentRecordType


# Exceptions
class UnexpectedEOFException(Exception):
    pass


class UnexpectedRecordException(Exception):
    def __init__(self, record, *expected):
        super().__init__(f'Unexpected record: {record}; Expected: {", ".join(expected)}')
        self.record = record
        self.expected = expected


# Types
class RecordDescriptor:
    def __init__(self, rtype, size):
        self.rtype = rtype
        self.size = size
    
    def write(self, processor):
        rtype = self.rtype.value
        if rtype <= 0x7f:
            processor.write(rtype)
        else:
            d = struct.pack('<H', rtype)
            processor.write(0x80 | d[0])
            processor.write((d[1] << 1) | ((d[0] & 0x80) >> 7))
        
        dsize = struct.pack('<I', self.size)
        for i in range(3):
            has_next = dsize[i + 1] != 0
            processor.write(0x80 | (dsize[i] & 0x7f) if has_next else (dsize[i] & 0x7f))
            if not has_next:
                break
    
    def skip(self, target):
        target.seek(self.size, io.SEEK_CUR)
    
    def __str__(self):
        return f'{self.rtype} ({self.size} bytes)'



class RecordReadState(Enum):
    FUTURE_RECORD = auto()
    ALT_CONTENT = auto()


class RecordProcessor:
    def __init__(self, stream):
        self.stream = stream
        self.read_stack = deque()
        self.single_buf = bytearray(1)
    
    def read_descriptor(self):
        stack = self.read_stack
        
        # Read Type Number
        d = self.read(1)
        rtype_num = (((self.read(1) & 0x7f) << 7) | d & 0x7f) if d & 0x80 else d & 0x7f
        
        # Determine Type
        if not len(stack):
            rtype = BinaryRecordType(rtype_num)
        else:
            try:
                rtype = BinaryRecordType(rtype_num)
            except ValueError:
                state = stack[-1]
                if state == RecordReadState.FUTURE_RECORD:
                    rtype = FutureRecordType(rtype_num)
                elif state == RecordReadState.ALT_CONTENT:
                    rtype = AlternateContentRecordType(rtype_num)
                else:
                    raise ValueError(state)
        
        
        # Evaluate State
        if rtype == BinaryRecordType.BrtFRTBegin:
            stack.append(RecordReadState.FUTURE_RECORD)
        elif rtype == BinaryRecordType.BrtACBegin:
            stack.append(RecordReadState.ALT_CONTENT)
        elif rtype in (BinaryRecordType.BrtFRTEnd, BinaryRecordType.BrtACEnd):
            if not len(stack):
                raise UnexpectedRecordException(rtype, f'NOT ({BinaryRecordType.BrtFRTEnd}, {BinaryRecordType.BrtACEnd})')
            stack.pop()
        
        
        # Determine Size
        size = 0
        for i in range(4):
            d = self.read(1)
            size = ((d & 0x7f) << (7 * i)) | size
            if not d & 0x80:
                break
        
        return RecordDescriptor(rtype, size)
    
    
    def read(self, size):
        if size <= 0:
            raise ValueError(size)
        
        stream = self.stream
        if size == 1:
            single_buf = self.single_buf
            ct = stream.readinto(single_buf)
            if ct == 0:
                raise UnexpectedEOFException()
            return single_buf[0]
        else:
            return stream.read(size)
    
    def write(self, data):
        stream = self.stream
        if isinstance(data, int):
            single_buf = self.single_buf
            single_buf[0] = data
            stream.write(single_buf)
        else:
            stream.write(data)
    
    def seek(self, n, whence=io.SEEK_SET):
        self.stream.seek(n, whence)