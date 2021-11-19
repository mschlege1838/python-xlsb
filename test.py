
import sys
import os

from bprocessor import RecordProcessor
from btypes import BinaryRecordType


import struct

import io


def fuck(b):
    res = []
    for i in b:
        res.append(f'{i:08b}')
    return ''.join(res)


fname = sys.argv[1]
fsize = os.stat(fname).st_size

with open(fname, 'rb') as f:
    rprocessor = RecordProcessor(f)
    
    r = rprocessor.read_descriptor()
    if r.rtype != BinaryRecordType.BrtBeginBook:
        raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginBook)
    
    while r.rtype != BinaryRecordType.BrtBeginBundleShs:
        r.skip(rprocessor)
        r = rprocessor.read_descriptor()
    
    
    r = rprocessor.read_descriptor()
    while r.rtype == BinaryRecordType.BrtBundleSh:
        hs_state = struct.unpack('<I', rprocessor.read(4))[0]
        i_tab_id = struct.unpack('<I', rprocessor.read(4))[0]
        rel_id = rprocessor.read_xl_nullable_w_string()
        str_name = rprocessor.read_xl_nullable_w_string()
        print(hs_state, i_tab_id, rel_id, str_name)
        
        r = rprocessor.read_descriptor()
    