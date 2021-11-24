
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
            r = rprocessor.skip_until(BinaryRecordType.BrtEndWsViews, repository=repository, skip_last=True)
        repository.push_current()
        
        # Skip 3
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
                r = r.read_descriptor()
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
            
            if r.rtype != BinaryRecordType.BrtRowHdr:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtRowHdr)
            row_header = RowHeader.read(rprocessor)
            
            # Cells
            cells = []
            r = rprocessor.read_descriptor()
            while True:
                # Skip 5
                r = rprocessor.skip_while(BinaryRecordType.BrtCellMeta, BinaryRecordType.BrtValueMeta, repository=repository, current=r)
                repository.push_current()
                
                if r.rtype == BinaryRecordType.BrtEndSheetData:
                    rows_done = True
                    break
                
                if r.rtype in (BinaryRecordType.BrtRowHdr, BinaryRecordType.BrtACBegin):
                    break
                
                if r.rtype == BinaryRecordType.BrtCellBlank:
                    cells.append(BlankCell.read(rprocessor))
                elif r.rtype == BinaryRecordType.BrtCellRk:
                    cells.append(RkCell.read(rprocessor))
                elif r.rtype == BinaryRecordType.BrtCellError:
                    cells.append(ErrorCell.read(rprocessor))
                elif r.rtype == BinaryRecordType.BrtCellBool:
                    cells.append(BoolCell.read(rprocessor))
                elif r.rtype == BinaryRecordType.BrtCellReal:
                    cells.append(RealCell.read(rprocessor))
                elif r.rtype == BinaryRecordType.BrtCellIsst:
                    cells.append(SharedStringCell.read(rprocessor))
                elif r.rtype == BinaryRecordType.BrtCellSt:
                    cells.append(InlineStringCell.read(rprocessor))
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
                        (RkCell.read(rprocessor))
                    if r.rtype == BinaryRecordType.BrtCellError:
                        (ErrorCell.read(rprocessor))
                    if r.rtype == BinaryRecordType.BrtCellBool:
                        (BoolCell.read(rprocessor))
                    if r.rtype == BinaryRecordType.BrtCellReal:
                        (RealCell.read(rprocessor))
                    if r.rtype == BinaryRecordType.BrtCellSt:
                        (InlineStringCell.read(rprocessor))
                    else:
                        raise UnexpectedRecordException(BinaryRecordType.BrtCellRk, BinaryRecordType.BrtCellError, BinaryRecordType.BrtCellBool,
                                BinaryRecordType.BrtCellReal, BinaryRecordType.BrtCellSt)
                    
                else:
                    raise UnexpectedRecordException(r, BinaryRecordType.BrtCellBlank, BinaryRecordType.BrtCellRk, BinaryRecordType.BrtCellError, BinaryRecordType.BrtCellBool,
                            BinaryRecordType.BrtCellReal, BinaryRecordType.BrtCellIsst, BinaryRecordType.BrtCellSt, BinaryRecordType.BrtFmlaString, BinaryRecordType.BrtFmlaNum,
                            BinaryRecordType.BrtFmlaBool, BinaryRecordType.BrtFmlaError, BinaryRecordType.BrtShrFmla, BinaryRecordType.BrtArrFmla, BinaryRecordType.BrtTable)
                
                r = rprocessor.read_descriptor()
                
                
            rows.append(Row(row_header, cells))
        
        
        rprocessor.skip_until(BinaryRecordType.BrtEndSheet, repository=repository)
        
        return WorksheetPart(sheet_dimension, col_info, rows, repository=repository)


    def __init__(self, sheet_dimension, col_info, rows, *, repository=None):
        self.sheet_dimension = sheet_dimension
        self.col_info = col_info
        self.rows = rows
        self.repository = repository


class Cell:
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        column = struct.unpack('<i', rprocessor.read(4))[0]
        i_style_ref = struct.unpack('<i', rprocessor.read(3) + bytes(1))[0]
        
        flags = rprocessor.read(1)
        show_phonetic_info = flags & 0x80
        reserved = flags & 0x7f
        
        return Cell(column, i_style_ref, bool(show_phonetic_info))
    
    def __init__(self, column, style_index, show_phonetic_info):
        self.column = column
        self.style_index = style_index
        self.show_phonetic_info = show_phonetic_info

class BlankCell:
    @staticmethod
    def read(stream):
        return BlankCell(Cell.read(stream))
    
    def __init__(self, cell):
        self.cell = cell


class RkCell:
    @staticmethod
    def read(stream):
        cell = Cell.read(stream)
        
        rbuf = bytearray(stream.read(4))
        f_x100 = rbuf[0] & 0x01
        f_int = rbuf[0] & 0x02
        rbuf[0] = rbuf[0] & 0xfc
        
        if f_int:   
            num = struct.unpack('<i', rbuf)[0] >> 2
        else:
            num = struct.unpack('<d', bytes(4) + rbuf)[0]
        
        if f_x100:
            num /= 100
        
        return RkCell(cell, num)
    
    def __init__(self, cell, num):
        self.cell = cell
        self.num = num
    
    @property
    def value(self):
        return self.num

class ErrorCell:
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        cell = Cell.read(rprocessor)
        b_error = rprocessor.read(1)
        
        return ErrorCell(cell, b_error)
        
    
    def __init__(self, cell, error_number):
        self.cell = cell
        self.error_number = error_number
    
    @property
    def value(self):
        return self.error
    
    @property
    def error(self):
        return error_lookup[self.error_number]
    
    @error.setter
    def error(self, value):
        self.error_number = error_rlookup[value]


class BoolCell:
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        cell = Cell.read(stream)
        f_bool = stream.read(1)
        
        return BoolCell(cell, bool(f_bool))
        
    
    def __init__(self, cell, val):
        self.cell = cell
        self.val = val
    
    @property
    def value(self):
        return self.val


class RealCell:
    @staticmethod
    def read(stream):
        cell = Cell.read(stream)
        xnum = struct.unpack('<d', stream.read(8))[0]
        return RealCell(cell, xnum)
    
    def __init__(self, cell, num):
        self.cell = cell
        self.num = num
    
    @property
    def value(self):
        return self.num


class SharedStringCell:
    @staticmethod
    def read(stream):
        cell = Cell.read(stream)
        isst = struct.unpack('<I', stream.read(4))[0]
        return SharedStringCell(cell, isst)


    def __init__(self, cell, str_index):
        self.cell = cell
        self.str_index = str_index
    
    @property
    def value(self):
        return f'<SharedIndex> {self.str_index}'


class InlineStringCell:
    @staticmethod
    def check_value(value):
        if len(value) > 32767:
            raise ValueError(f'Inline strings must be less than or equal to 32767 characters: {len(value)}')
        
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        cell = Cell.read(rprocessor)
        val = rprocessor.read_xl_w_string(False)
        return InlineStringCell(cell, val)
    
    def __init__(self, cell, val):
        InlineStringCell.check_value(val)
        self.cell = cell
        self.val = val
    
    @property
    def value(self):
        return self.val


class Row:
    def __init__(self, header, cells):
        self.header = header
        self.cells = cells

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
        btypes.validate_rw(row_first)
        btypes.validate_rw(row_last)
        if row_last < row_first:
            raise ValueError(f'Last row cannot be less than the first: {row_last}; First={row_first}')
        btypes.validate_col(col_first)
        btypes.validate_col(col_last)
        if col_last < col_first:
            raise ValueError(f'Last column cannot be less than the first: {col_last}; First={col_first}')
        
        self.row_first = row_first
        self.row_last = row_last
        self.col_first = col_first
        self.col_last = col_last
    
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