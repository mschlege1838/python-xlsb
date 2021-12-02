
import struct

from bprocessor import RecordProcessor, UnexpectedRecordException, RecordDescriptor
from btypes import BinaryRecordType


class ACUid:
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        ac = AlternatContent.read(rprocessor)
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtUid:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtUid)
        
        data = stream.read(r.size)
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtACEnd:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtACEnd)
        
        return ACUid(ac, data)
    
    def __init__(self, ac, data):
        self.ac = ac
        self.data = data
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        ac = self.ac
        RecordDescriptor(BinaryRecordType.BrtACBegin, len(ac)).write(rprocessor)
        ac.write(rprocessor)
        
        RecordDescriptor(BinaryRecordType.BrtUid, len(self)).write(rprocessor)
        rprocessor.write(self.data)
        
        RecordDescriptor(BinaryRecordType.BrtACEnd).write(rprocessor)
    
    def __len__(self):
        return len(self.data)
    
    


class AlternatContent:
    @staticmethod
    def read(stream):
        c_ver = struct.unpack('<H', stream.read(2))[0]
        
        product_versions = []
        for i in range(c_ver):
            product_versions.append(ACProductVersion.read(stream))
        
        return AlternatContent(product_versions)
    
    def __init__(self, product_versions):
        self.product_versions = product_versions
    
    def write(self, stream):
        product_versions = self.product_versions
        stream.write(struct.pack('<H', len(product_versions)))
        for product_version in product_versions:
            product_version.write(stream)
    
    def __len__(self):
        return 2 + sum(len(v) for v in self.product_versions)

class ACProductVersion:
    @staticmethod
    def read(stream):
        file_version, flags = struct.unpack('<HH', stream.read(4))
        file_product = flags & 0x7fff
        file_extension = flags & 0x8000
        return ACProductVersion(file_version, file_product, bool(file_extension))
    
    def __init__(self, version, product, forward_compatiblity):
        self.version = version
        self.product = product
        self.forward_compatiblity = forward_compatiblity
    
    def write(self, stream):
        flags = self.product
        if self.forward_compatiblity:
            flags |= 0x8000
        stream.write(struct.pack('<HH', self.version, flags))
    
    def __len__(self):
        return 4