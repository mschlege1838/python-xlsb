
import struct

import btypes
from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository, RecordDescriptor


class SharedStringsPart:
    @staticmethod
    def read(stream):
        rprocessor = RecordDescriptor.resolve(stream)
        
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginSst:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginSst)
        
        