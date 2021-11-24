
import struct

from bprocessor import RecordProcessor


class RichString:
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        flags_1 = rprocessor.read(1)
        f_rich_str = flags_1 & 0x80
        f_ext_str = flags_1 & 0x40
        unused_1 = flags_1 & 0x3f
        
        v_str = rprocessor.read_xl_w_string(False)
        
        if f_rich_str:
            dw_size_str_run = 