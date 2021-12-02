
import io
import struct
from collections import deque
from enum import Enum, auto
from tempfile import TemporaryFile

from btypes import BinaryRecordType, FutureRecordType, AlternateContentRecordType


# Exceptions
class UnexpectedEOFException(Exception):
    pass


class UnexpectedRecordException(Exception):
    def __init__(self, record, *expected):
        super().__init__(f'Unexpected record: {record}; Expected: {", ".join(str(i) for i in expected)}')
        self.record = record
        self.expected = expected

class UnexpectedRecordSizeException(Exception):
    def __init__(self, read_size, record):
        super().__init__(f'Unexpected read size for record: {read_size}; {record}')

# Types
class RecordDescriptor:
    
    @staticmethod
    def iter_parts(stream):
        rprocessor = RecordProcessor.resolve(stream)
        try:
            while True:
                r = rprocessor.read_descriptor()
                yield r, stream.read(r.size)
        except UnexpectedEOFException:
            pass
    
    def __init__(self, rtype, size=0):
        self.rtype = rtype
        self.size = size
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
    
        rtype = self.rtype.value
        if rtype <= 0x7f:
            rprocessor.write(rtype)
        else:
            d = struct.pack('<H', rtype)
            rprocessor.write(0x80 | d[0])
            rprocessor.write((d[1] << 1) | ((d[0] & 0x80) >> 7))
        
        dsize = struct.pack('<I', self.size)
        count = 0
        for i, val in enumerate(dsize):
            if val != 0:
                count = i
        
        for i in range(count + 1):
            o = dsize[i] if i == 0 else ((dsize[i] << 1) | ((dsize[i - 1] & 0x80) >> 7))
            rprocessor.write((0x80 | o) if i < count - 1 else (o & 0x7f))
    
    def skip(self, target, repository=None):
        if repository and repository.for_update:
            data = target.read(self.size, single_as_int=False)
            # print(self, data)
            repository.store(self, data)
        else:
            target.seek(self.size, io.SEEK_CUR)

    
    def __str__(self):
        return f'{self.rtype} ({self.size} bytes)'



class RecordReadState(Enum):
    FUTURE_RECORD = auto()
    ALT_CONTENT = auto()


class RecordProcessor:
    
    empty_data = bytes(0)
    
    @staticmethod
    def check_rel_id(value):
        if not value:
            return
        if len(value) > 255:
            raise ValueError(f'Character length must be less than 255: {value}')
        if '\0' in value:
            raise ValueError(f'Must not contain the NUL character: {value}')
    
    @staticmethod
    def len_xl_w_string(value, nullable=True):
        if value is None:
            if nullable:
                return 4
            else:
                if len(value) >= 0xffffffff:
                    raise ValueError(f'Character length must be less than {0xffffffff}: {value}')
        
        return 4 + len(value) * 2
    
    @staticmethod
    def resolve(stream):
        return stream if isinstance(stream, RecordProcessor) else RecordProcessor(stream)
    
    
    def __init__(self, stream):
        self.stream = stream
        self.read_stack = deque()
        self.single_buf = bytearray(1)
    
    def skip_until(self, *until_lst, repository=None, skip_last=False):
        while True:
            r = self.read_descriptor()
            if r.rtype in until_lst:
                if skip_last:
                    r.skip(self, repository)
                break
            r.skip(self, repository)
        return self.read_descriptor() if skip_last else r
    
    def skip_while(self, *while_lst, current=None, repository=None):
        while True:
            r = current if current else self.read_descriptor()
            current = None
            if r.rtype not in while_lst:
                break
            r.skip(self, repository)
        return r
    
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
                    raise ValueError(rtype_num)
        
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
    
    def read_xl_w_string(self, nullable=True):
        cch_characters = self.read(4)
        if cch_characters == 0xffffffff:
            if nullable:
                return None
            else:
                raise ValueError(f'Character length must be less than {0xffffffff}.')
        
        cch_len = struct.unpack('<I', cch_characters)[0] * 2
        rgch_data = self.read(cch_len).decode('utf-16le')
        
        return rgch_data
    
    def write_xl_w_string(self, value, nullable=True):
        if nullable:
            if value is None:
                self.write(struct.pack('<I', 0xffffffff))
                return
        else:
            if len(value) >= 0xffffffff:
                raise ValueError(f'Character length must be less than {0xffffffff}: {value}')
        
        self.write(struct.pack('<I', len(value)))
        self.write(value.encode('utf-16le'))

    
    def read(self, size, *, single_as_int=True):
        if size < 0:
            raise ValueError(size)
        if size == 0:
            return RecordProcessor.empty_data
        
        stream = self.stream
        if size == 1 and single_as_int:
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
    



class RecordCopy:
    def __init__(self, descriptor, data_len, data_file):
        self.descriptor = descriptor
        self.data_len = data_len
        self.data_file = data_file
    
    def write_to(self, stream):
        self.descriptor.write(stream)
        if self.data_len:
            stream.write(self.data_file.read(self.data_len))
    
    def __len__(self):
        return self.data_len

class RecordRepository:
    def __init__(self, for_update):
        self.for_update = for_update
        if for_update:
            self.queue = deque()
            self.current = []
            self.f = TemporaryFile()
    
    def store(self, descriptor, data):
        if not self.for_update:
            return
        f = self.f
        if data:
            f.write(data)
        self.current.append(RecordCopy(descriptor, len(data), f))
    
    def push_current(self):
        if not self.for_update:
            return
        self.queue.append(self.current)
        self.current = []
    
    def begin_write(self):
        if not self.for_update:
            return
        self.f.seek(0)
    
    def write_poll(self, stream):
        if not self.for_update:
            return
        for item in self.queue.popleft():
            item.write_to(stream)
    
    def close(self):    
        if not self.for_update:
            return
        self.f.close()
    
    def __len__(self):
        if not self.for_update:
            return 0
        result = 0
        for l in self.queue:
            for i in l:
                result += len(i)
        return result
    
    def __bool__(self):
        return True
    
    