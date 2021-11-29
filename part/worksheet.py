
import math
import struct

import btypes
from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository, RecordDescriptor


error_lookup = {
    0x00: '#NULL!',
    0x07: '#DIV/0!',
    0x0F: '#VALUE!',
    0x17: '#REF!',
    0x1D: '#NAME?',
    0x24: '#NUM!',
    0x2A: '#N/A',
    0x2B: '#GETTING_DATA'
}
    
error_rlookup = dict((v, k) for k, v in error_lookup.items())


def check_poll(self, stream):
    repository = self.repository
    if repository:
        repository.write_poll(stream)


class WorksheetPart:
    @staticmethod
    def read(stream, for_update=False):
        rprocessor = RecordProcessor.resolve(stream)
        repository = RecordRepository(for_update)
        
        # Begin
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginSheet:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginSheet)
        
        # Skip 1
        r = rprocessor.read_descriptor()
        if r.rtype == BinaryRecordType.BrtWsProp:
            r.skip(rprocessor, repository)
            r = rprocessor.read_descriptor()
        repository.push_current()
        
        
        # Sheet Dimension
        sheet_dimension = None
        if r.rtype == BinaryRecordType.BrtWsDim:
            sheet_dimension = SheetDimension.read(rprocessor)
            r = rprocessor.read_descriptor()
        
        
        # Skip 2
        if r.rtype == BinaryRecordType.BrtBeginWsViews:
            r.skip(rprocessor, repository)
            r = rprocessor.skip_until(BinaryRecordType.BrtEndWsViews, repository=repository, skip_last=True)
        if r.rtype == BinaryRecordType.BrtACBegin:
            r.skip(rprocessor, repository)
            r = rprocessor.skip_until(BinaryRecordType.BrtACEnd, repository=repository, skip_last=True)
        if r.rtype == BinaryRecordType.BrtWsFmtInfo:
            r.skip(rprocessor, repository)
            r = rprocessor.read_descriptor()
        repository.push_current()
        
        # Col Info
        col_info = []
        if r.rtype == BinaryRecordType.BrtBeginColInfos:
            r = rprocessor.read_descriptor()
            while r.rtype == BinaryRecordType.BrtColInfo:
                col_info.append(ColInfo.read(rprocessor))
                r = rprocessor.read_descriptor()
            if r.rtype != BinaryRecordType.BrtEndColInfos:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtEndColInfos)
            r = rprocessor.read_descriptor()
        
        
        # Cell Table
        if r.rtype != BinaryRecordType.BrtBeginSheetData:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginSheetData)
        
        # Rows
        rows = []
        r = rprocessor.read_descriptor()
        rows_done = False
        while not rows_done:
            # Skip 4
            if r.rtype == BinaryRecordType.BrtACBegin:
                r.skip(rprocessor, repository)
                r = rprocessor.skip_until(BinaryRecordType.BrtACEnd, repository=repository, skip_last=True)
            repository.push_current()
            
            if r.rtype == BinaryRecordType.BrtEndSheetData:
                rows_done = True
                break
            
            if r.rtype != BinaryRecordType.BrtRowHdr:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtRowHdr)
            row_header = RowHeader.read(rprocessor)
            
            # Cells
            cells = []
            r = rprocessor.read_descriptor()
            while True:
                if r.rtype == BinaryRecordType.BrtEndSheetData:
                    rows_done = True
                    break
                
                if r.rtype in (BinaryRecordType.BrtRowHdr, BinaryRecordType.BrtACBegin):
                    break
                
                # Skip 5
                r = rprocessor.skip_while(BinaryRecordType.BrtCellMeta, BinaryRecordType.BrtValueMeta, repository=repository, current=r)
                repository.push_current()
                
                if r.rtype == BinaryRecordType.BrtCellBlank:
                    cells.append(BlankCell.read(rprocessor, repository=repository))
                elif r.rtype == BinaryRecordType.BrtCellRk:
                    cells.append(RkCell.read(rprocessor, repository=repository))
                elif r.rtype == BinaryRecordType.BrtCellError:
                    cells.append(ErrorCell.read(rprocessor, repository=repository))
                elif r.rtype == BinaryRecordType.BrtCellBool:
                    cells.append(BoolCell.read(rprocessor, repository=repository))
                elif r.rtype == BinaryRecordType.BrtCellReal:
                    cells.append(RealCell.read(rprocessor, repository=repository))
                elif r.rtype == BinaryRecordType.BrtCellIsst:
                    cells.append(SharedStringCell.read(rprocessor, repository=repository))
                elif r.rtype == BinaryRecordType.BrtCellSt:
                    cells.append(InlineStringCell.read(rprocessor, repository=repository))
                elif r.rtype == BinaryRecordType.BrtFmlaString:
                    raise Exception('Not Implemented')
                elif r.rtype == BinaryRecordType.BrtFmlaNum:
                    raise Exception('Not Implemented')
                elif r.rtype == BinaryRecordType.BrtFmlaBool:
                    raise Exception('Not Implemented')
                elif r.rtype == BinaryRecordType.BrtFmlaError:
                    raise Exception('Not Implemented')
                elif r.rtype == BinaryRecordType.BrtShrFmla:
                    raise Exception('Not Implemented')
                elif r.rtype == BinaryRecordType.BrtArrFmla:
                    raise Exception('Not Implemented')
                elif r.rtype == BinaryRecordType.BrtTable:
                    # TODO BrtTable
                    raise Exception('Not Implemented')
                    
                    # Skip 6
                    r = rprocessor.skip_while(BinaryRecordType.BrtCellMeta, BinaryRecordType.BrtValueMeta, repository=repository, skip_last=True)
                    repository.push_current()
                    
                    if r.rtype == BinaryRecordType.BrtCellRk:
                        (RkCell.read(rprocessor, repository=repository))
                    if r.rtype == BinaryRecordType.BrtCellError:
                        (ErrorCell.read(rprocessor, repository=repository))
                    if r.rtype == BinaryRecordType.BrtCellBool:
                        (BoolCell.read(rprocessor, repository=repository))
                    if r.rtype == BinaryRecordType.BrtCellReal:
                        (RealCell.read(rprocessor, repository=repository))
                    if r.rtype == BinaryRecordType.BrtCellSt:
                        (InlineStringCell.read(rprocessor, repository=repository))
                    else:
                        raise UnexpectedRecordException(BinaryRecordType.BrtCellRk, BinaryRecordType.BrtCellError, BinaryRecordType.BrtCellBool,
                                BinaryRecordType.BrtCellReal, BinaryRecordType.BrtCellSt)
                    
                else:
                    raise UnexpectedRecordException(r, BinaryRecordType.BrtCellBlank, BinaryRecordType.BrtCellRk, BinaryRecordType.BrtCellError, BinaryRecordType.BrtCellBool,
                            BinaryRecordType.BrtCellReal, BinaryRecordType.BrtCellIsst, BinaryRecordType.BrtCellSt, BinaryRecordType.BrtFmlaString, BinaryRecordType.BrtFmlaNum,
                            BinaryRecordType.BrtFmlaBool, BinaryRecordType.BrtFmlaError, BinaryRecordType.BrtShrFmla, BinaryRecordType.BrtArrFmla, BinaryRecordType.BrtTable)
                
                r = rprocessor.read_descriptor()
                
                
            rows.append(Row(row_header, cells, repository=repository))
        
        
        # Skip 7
        rprocessor.skip_until(BinaryRecordType.BrtEndSheet, repository=repository)
        repository.push_current()
        
        return WorksheetPart(sheet_dimension, col_info, rows, repository=repository)


    def __init__(self, sheet_dimension, col_info, rows, *, repository=None):
        self.sheet_dimension = sheet_dimension if sheet_dimension else SheetDimension(0, 0, 0, 0)
        self.col_info = col_info
        self.rows = rows
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        repository = self.repository
        repository.begin_write()
        
        # Begin
        RecordDescriptor(BinaryRecordType.BrtBeginSheet).write(rprocessor)
        
        # Skip 1
        if repository:
            repository.write_poll(rprocessor)
        
        # Sheet Dimension
        sheet_dimension = self.sheet_dimension
        sheet_dimension.refresh_from(self.rows)
        sheet_dimension.write(rprocessor)
        
        # Skip 2
        if repository:
            repository.write_poll(rprocessor)
        
        # Col Info
        col_info = self.col_info
        if col_info:
            RecordDescriptor(BinaryRecordType.BrtBeginColInfos).write(rprocessor)
            for ci in col_info:
                ci.write(rprocessor)
            RecordDescriptor(BinaryRecordType.BrtEndColInfos).write(rprocessor)
        
        # Cell Table
        RecordDescriptor(BinaryRecordType.BrtBeginSheetData).write(rprocessor)
        for row in self.rows:
            # Skips 4, 5, (6) in impl
            row.write(rprocessor)
        RecordDescriptor(BinaryRecordType.BrtEndSheetData).write(rprocessor)
        
        # Skip 7
        if repository:
            repository.write_poll(rprocessor)
        
        RecordDescriptor(BinaryRecordType.BrtEndSheet).write(rprocessor)
        repository.close()

class Cell:
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        column = struct.unpack('<i', rprocessor.read(4))[0]
        i_style_ref = struct.unpack('<i', rprocessor.read(3) + bytes(1))[0]
        
        flags = rprocessor.read(1)
        show_phonetic_info = flags & 0x01
        reserved = flags & 0xfe
        
        return Cell(column, i_style_ref, bool(show_phonetic_info))
    
    def __init__(self, column, style_index, show_phonetic_info):
        self.column = column
        self.style_index = style_index
        self.show_phonetic_info = show_phonetic_info
    
    def write(self, stream):
        stream.write(struct.pack('<i', self.column))
        stream.write(struct.pack('<i', self.style_index)[:3])
        stream.write(bytes((1,)) if self.show_phonetic_info else bytes(1))
    
    def __len__(self):
        return 8
        
class BlankCell:
    @staticmethod
    def read(stream, *, repository=None):
        return BlankCell(Cell.read(stream), repository=repository)
    
    def __init__(self, cell, *, repository=None):
        self.cell = cell
        self.repository = repository
    
    @property
    def value(self):
        return None
    
    def write(self, stream):
        check_poll(self, stream)
        RecordDescriptor(BinaryRecordType.BrtCellBlank, len(self)).write(stream)
        self.cell.write(stream)
    
    def __len__(self):
        return len(self.cell)


class RkCell:
    @staticmethod
    def convert_real(num):
        packed = struct.pack('<d', num)
        raw = struct.unpack('<Q', packed)[0]
        raw_truncated = raw & 0xfffffffc00000000
        repacked = struct.pack('<Q', raw_truncated)
        return struct.unpack('<d', repacked)[0], repacked[4:]
    
    @staticmethod
    def read(stream, *, repository=None):
        cell = Cell.read(stream)
        
        r = stream.read(4)
        rbuf = bytearray(r)
        f_x100 = rbuf[0] & 0x01
        f_int = rbuf[0] & 0x02
        rbuf[0] = rbuf[0] & 0xfc
        
        
        if f_int:   
            num = struct.unpack('<i', rbuf)[0] >> 2
        else:
            num = struct.unpack('<d', bytes(4) + rbuf)[0]
        
        if f_x100:
            num /= 100
        
        return RkCell(cell, num, repository=repository)
    
    def __init__(self, cell, num, *, repository=None):
        self.cell = cell
        self.num = num
        self.repository = repository
    
    @property
    def value(self):
        return self.num
    
    def write(self, stream):
        check_poll(self, stream)
        RecordDescriptor(BinaryRecordType.BrtCellRk, len(self)).write(stream)
        self.cell.write(stream)
        
        num = self.num
        flags = 0
        
        converted, data = RkCell.convert_real(num)
        if converted != num:
            converted, data = RkCell.convert_real(num * 100)
            flags |= 0x01
            if converted / 100 != num:
                converted = int(num * 100) & 0xfffffffc
                flags |= 0x02
                data = struct.pack('<i', converted)
                if converted / 100 != num:
                    raise ValueError(f'Number out of RkCell range. Use RealCell: {num}')
        
        data = bytearray(data)
        data[0] |= flags
        stream.write(data)
        
    
    def __len__(self):
        return len(self.cell) + 4

class ErrorCell:
    @staticmethod
    def read(stream, *, repository=None):
        rprocessor = RecordProcessor.resolve(stream)
        
        cell = Cell.read(rprocessor)
        b_error = rprocessor.read(1)
        
        return ErrorCell(cell, b_error, repository=repository)
        
    
    def __init__(self, cell, error_number, *, repository=None):
        self.cell = cell
        self.error_number = error_number
        self.repository = repository
    
    @property
    def value(self):
        return self.error
    
    @property
    def error(self):
        return error_lookup[self.error_number]
    
    @error.setter
    def error(self, value):
        self.error_number = error_rlookup[value]
    
    def write(self, stream):
        check_poll(self, stream)
        RecordDescriptor(BinaryRecordType.BrtCellError, len(self)).write(stream)
        self.cell.write(stream)
        RecordProcessor.resolve(stream).write(self.error_number)
    
    def __len__(self):
        return len(self.cell) + 1
    
    


class BoolCell:
    @staticmethod
    def read(stream, *, repository=None):
        rprocessor = RecordProcessor.resolve(stream)
        
        cell = Cell.read(stream)
        f_bool = stream.read(1)
        
        return BoolCell(cell, bool(f_bool), repository=repository)
        
    
    def __init__(self, cell, val, *, repository=None):
        self.cell = cell
        self.val = val
        self.repository = repository
    
    @property
    def value(self):
        return self.val
    
    def write(self, stream):
        check_poll(self, stream)
        RecordDescriptor(BinaryRecordType.BrtCellBool, len(self)).write(stream)
        self.cell.write(stream)
        RecordProcessor.resolve(stream).write(1 if self.val else 0)
    
    def __len__(self):
        return len(self.cell) + 1


class RealCell:
    @staticmethod
    def validate_xnum(value):
        if math.isinf(value) or math.isnan(value):
            raise ValueError(f'Xnums cannot be Infinity or NaN: {value}')
    
    @staticmethod
    def read(stream, *, repository=None):
        cell = Cell.read(stream)
        xnum = struct.unpack('<d', stream.read(8))[0]
        return RealCell(cell, xnum, repository=repository)
    
    def __init__(self, cell, num, *, repository=None):
        RealCell.validate_xnum(num)
        self.cell = cell
        self.num = num
        self.repository = repository
    
    @property
    def value(self):
        return self.num
    
    def write(self, stream):
        check_poll(self, stream)
        RecordDescriptor(BinaryRecordType.BrtCellReal, len(self)).write(stream)
        self.cell.write(stream)
        num = self.num
        RealCell.validate_xnum(num)
        stream.write(struct.pack('<d', num))
    
    def __len__(self):
        return len(self.cell) + 8


class SharedStringCell:
    @staticmethod
    def read(stream, *, repository=None):
        cell = Cell.read(stream)
        isst = struct.unpack('<I', stream.read(4))[0]
        return SharedStringCell(cell, isst, repository=repository)


    def __init__(self, cell, str_index, *, repository=None):
        self.cell = cell
        self.str_index = str_index
        self.repository = repository
    
    @property
    def value(self):
        return f'<SharedIndex> {self.str_index}'
    
    def write(self, stream):
        check_poll(self, stream)
        RecordDescriptor(BinaryRecordType.BrtCellIsst, len(self)).write(stream)
        self.cell.write(stream)
        stream.write(struct.pack('<I', self.str_index))
    
    def __len__(self):
        return len(self.cell) + 4
    
    


class InlineStringCell:
    @staticmethod
    def check_value(value):
        if len(value) > 32767:
            raise ValueError(f'Inline strings must be less than or equal to 32767 characters: {len(value)}')
        
    @staticmethod
    def read(stream, *, repository=None):
        rprocessor = RecordProcessor.resolve(stream)
        cell = Cell.read(rprocessor)
        val = rprocessor.read_xl_w_string(False)
        return InlineStringCell(cell, val, repository=repository)
    
    def __init__(self, cell, val, *, repository=None):
        InlineStringCell.check_value(val)
        self.cell = cell
        self.val = val
        self.repository = repository
    
    @property
    def value(self):
        return self.val
    
    def write(self, stream):
        check_poll(self, stream)
        RecordDescriptor(BinaryRecordType.BrtCellSt, len(self)).write(stream)
        self.cell.write(stream)
        RecordProcessor.resolve(stream).write_xl_w_string(self.val, False)
    
    def __len__(self):
        return len(self.cell) + len(self.val) * 2


class Row:
    def __init__(self, header, cells, *, repository=None):
        self.header = header
        self.cells = cells
        self.repository = repository
    
    @property
    def col_range(self):
        col_spans = self.header.col_spans
        return min(col_spans, key=lambda v: v.col_first).col_first, max(col_spans, key=lambda v: v.col_last).col_last
    
    def write(self, stream):
        check_poll(self, stream)
        
        self.header.write(stream)
        for cell in self.cells:
            cell.write(stream)

class RowHeader:
    @staticmethod
    def validate_row_height(value):
        if value > 0x2000:
            raise ValueError(f'Row height must be less than or equal to {0x2000}: {value}')
    
    @staticmethod
    def validate_ccol_span(value):
        if value > 16:
            raise ValueError(f'Colspan must be less than or equal to 16: {value}')
    
    @staticmethod
    def validate_outline_level(value):
        if value > 0x07:
            raise ValueError(f'Outline level must be less than or equal to {0x07}: {value}')
    
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        rw = struct.unpack('<i', rprocessor.read(4))[0]
        ixfe = struct.unpack('<I', rprocessor.read(4))[0]
        miy_row = struct.unpack('<H', rprocessor.read(2))[0]
        
        flags_1 = rprocessor.read(1)
        extra_asc = flags_1 & 0x01
        extra_dsc = flags_1 & 0x02
        reserved_1 = flags_1 & 0xfc
        
        flags_2 = rprocessor.read(1)
        i_out_level = flags_2 & 0x07
        f_collapsed = flags_2 & 0x08
        f_dy_zero = flags_2 & 0x10
        f_unsynced = flags_2 & 0x20
        f_ghost_dirty = flags_2 & 0x40
        f_reserved = flags_2 & 0x80
        
        flags_3 = rprocessor.read(1)
        f_ph_show = flags_3 & 0x01
        reserved_2 = flags_3 & 0xfe
        
        ccol_span = struct.unpack('<I', rprocessor.read(4))[0]
        RowHeader.validate_ccol_span(ccol_span)
        
        col_spans = []
        for i in range(ccol_span):
            col_spans.append(ColumnSpan.read(rprocessor))
        
        return RowHeader(rw, ixfe, miy_row, bool(extra_asc), bool(extra_dsc), i_out_level, bool(f_collapsed), bool(f_dy_zero),
                bool(f_unsynced), bool(f_ghost_dirty), bool(f_ph_show), col_spans)
    
    def __init__(self, row_index, style_index, row_height, allocate_asc_padding, allocate_desc_padding, outline_level, outline_collapsed,
            hidden, manual_height, ghost_dirty, has_phonetic_guide, col_spans):
        RowHeader.validate_outline_level(outline_level)
        RowHeader.validate_ccol_span(len(col_spans))
        
        self.row_index = row_index
        self.style_index = style_index
        self.row_height = row_height
        self.allocate_asc_padding = allocate_asc_padding
        self.allocate_desc_padding = allocate_desc_padding
        self.outline_level = outline_level
        self.outline_collapsed = outline_collapsed
        self.hidden = hidden
        self.manual_height = manual_height
        self.ghost_dirty = ghost_dirty
        self.has_phonetic_guide = has_phonetic_guide
        self.col_spans = col_spans
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        RecordDescriptor(BinaryRecordType.BrtRowHdr, len(self)).write(rprocessor)
        
        rprocessor.write(struct.pack('<i', self.row_index))
        rprocessor.write(struct.pack('<I', self.style_index))
        rprocessor.write(struct.pack('<H', self.row_height))
        
        flags = 0
        if self.allocate_asc_padding:
            flags |= 0x01
        if self.allocate_desc_padding:
            flags |= 0x02
        rprocessor.write(flags)
        
        flags = self.outline_level
        RowHeader.validate_outline_level(flags)
        if self.outline_collapsed:
            flags |= 0x08
        if self.hidden:
            flags |= 0x10
        if self.manual_height:
            flags |= 0x20
        if self.ghost_dirty:
            flags |= 0x40
        rprocessor.write(flags)
        
        flags = 0
        if self.has_phonetic_guide:
            flags |= 0x01
        rprocessor.write(flags)
        
        
        col_spans = self.col_spans
        RowHeader.validate_ccol_span(len(col_spans))
        rprocessor.write(struct.pack('<I', len(col_spans)))
        for sp in col_spans:
            sp.write(rprocessor)
        
    
    def __len__(self):
        return 17 + sum(len(sp) for sp in self.col_spans)


class ColumnSpan:
    @staticmethod
    def read(stream):
        col_mic = struct.unpack('<i', stream.read(4))[0]
        col_last = struct.unpack('<i', stream.read(4))[0]
        return ColumnSpan(col_mic, col_last)
    
    def __init__(self, col_first, col_last):
        btypes.validate_col(col_first)
        btypes.validate_col(col_last)
        self.col_first = col_first
        self.col_last = col_last
    
    def write(self, stream):
        stream.write(struct.pack('<i', self.col_first))
        stream.write(struct.pack('<i', self.col_last))
    
    def __len__(self):
        return 8
    


class SheetDimension:
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        rw_first = struct.unpack('<i', stream.read(4))[0]
        rw_last = struct.unpack('<i', stream.read(4))[0]
        col_first = struct.unpack('<i', stream.read(4))[0]
        col_last = struct.unpack('<i', stream.read(4))[0]
        
        return SheetDimension(rw_first, rw_last, col_first, col_last)
    
    def __init__(self, row_first, row_last, col_first, col_last):
        self.row_first = row_first
        self.row_last = row_last
        self.col_first = col_first
        self.col_last = col_last
        self.validate()
    
    def refresh_from(self, rows):
        row_first = row_last = col_first = col_last = None
        for row in rows:
            if row_first is None or row.header.row_index < row_first:
                row_first = row.header.row_index
            if row_last is None or row.header.row_index > row_last:
                row_last = row.header.row_index
            for cell in row.cells:
                if col_first is None or cell.cell.column < col_first:
                    col_first = cell.cell.column
                if col_last is None or cell.cell.column > col_last:
                    col_last = cell.cell.column
        
        self.row_first = row_first
        self.row_last = row_last
        self.col_first = col_first
        self.col_last = col_last
        self.validate()
    
    def validate(self):
        row_first = self.row_first
        row_last = self.row_last
        col_first = self.col_first
        col_last = self.col_last
        
        btypes.validate_rw(row_first)
        btypes.validate_rw(row_last)
        if row_last < row_first:
            raise ValueError(f'Last row cannot be less than the first: {row_last}; First={row_first}')
        btypes.validate_col(col_first)
        btypes.validate_col(col_last)
        if col_last < col_first:
            raise ValueError(f'Last column cannot be less than the first: {col_last}; First={col_first}')
    
    def write(self, stream):
        self.validate()
        RecordDescriptor(BinaryRecordType.BrtWsDim, len(self)).write(stream)
        stream.write(struct.pack('<i', self.row_first))
        stream.write(struct.pack('<i', self.row_last))
        stream.write(struct.pack('<i', self.col_first))
        stream.write(struct.pack('<i', self.col_last))
    
    def __len__(self):
        return 16

class ColInfo:
    @staticmethod
    def read(rprocessor):
        col_first = struct.unpack('<i', rprocessor.read(4))[0]
        col_last = struct.unpack('<i', rprocessor.read(4))[0]
        col_idx = struct.unpack('<I', rprocessor.read(4))[0]
        ixfe = struct.unpack('<I', rprocessor.read(4))[0]
        
        flags_1 = rprocessor.read(1)
        f_hidden = flags_1 & 0x01
        f_user_set = flags_1 & 0x02
        f_best_fit = flags_1 & 0x04
        f_phonetic = flags_1 & 0x08
        reserved_1 = flags_1 & 0xf0
        
        flags_2 = rprocessor.read(1)
        i_out_level = flags_2 & 0x07
        unused = flags_2 & 0x08
        f_collapsed = flags_2 & 0x10
        reserved_2 = flags_2 & 0xe0
        
        return ColInfo(col_first, col_last, col_idx, ixfe, bool(f_hidden), bool(f_user_set), bool(f_best_fit), 
                bool(f_phonetic), i_out_level, bool(f_collapsed))
    
    def __init__(self, col_first, col_last, col_width, style_index, hidden, user_width_set, best_fit_set,
            phonetic, outline_level, outline_collapsed):
        self.col_first = col_first
        self.col_last = col_last
        self.col_width = col_width
        self.style_index = style_index
        self.hidden = hidden
        self.user_width_set = user_width_set
        self.best_fit_set = best_fit_set
        self.phonetic = phonetic
        self.outline_level = outline_level
        self.outline_collapsed = outline_collapsed
        self.validate()
    
    def write(self, stream):
        self.validate()
        
        rprocessor = RecordProcessor.resolve(stream)
        
        RecordDescriptor(BinaryRecordType.BrtColInfo, len(self)).write(stream)
        
        rprocessor.write(struct.pack('<i', self.col_first))
        rprocessor.write(struct.pack('<i', self.col_last))
        rprocessor.write(struct.pack('<I', self.col_width))
        rprocessor.write(struct.pack('<I', self.style_index))
        
        flags = 0
        if self.hidden:
            flags |= 0x01
        if self.user_width_set:
            flags |= 0x02
        if self.best_fit_set:
            flags |= 0x04
        if self.phonetic:
            flags |= 0x08
        rprocessor.write(flags)
        
        flags = self.outline_level
        if self.outline_collapsed:
            flags |= 0x10
        rprocessor.write(flags)
        
    
    def validate(self): 
        col_first = self.col_first
        col_last = self.col_last
        
        btypes.validate_col(col_first)
        btypes.validate_col(col_last)
        if col_last < col_first:
            raise ValueError(f'Last column cannot be less than first: {col_last}; First={col_first}')
        
        if self.col_width > 65535:
            raise ValueError(f'Column width must be less than or equal to 65535: {self.col_width}')
        
        if self.outline_level > 7:
            raise ValueError(f'Outline level must be less than or equal to 7: {self.outline_level}')
    
    def __len__(self):
        return 18