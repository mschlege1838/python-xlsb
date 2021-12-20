

import struct

from btypes import BinaryRecordType, HorizontalAlignmentType, VerticalAlignmentType, ReadingOrderType, XFProperty, ColorType, \
        PaletteColor, ThemeColor, SubscriptType, UnderlineType, FontFamilyType, CharacterSetType, FontSchemeType, FillType, GradientType, \
        BorderType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository, RecordDescriptor
from part import ACUid



num_format_lu_all_langs = {
    0: ''
    , 1: '0'
    , 2: '0.00'
    , 3: '#,##0'
    , 4: '#,##0.00'
    , 9: '0%'
    , 10: '0.00%'
    , 11: '0.00E+00'
    , 12: '# ?/?'
    , 13: '# ??/??'
    , 14: 'mm-dd-yy'
    , 15: 'd-mmm-yy'
    , 16: 'd-mmm'
    , 17: 'mmm-yy'
    , 18: 'h:mm AM/PM'
    , 19: 'h:mm:ss AM/PM'
    , 20: 'h:mm'
    , 21: 'h:mm:ss'
    , 22: 'm/d/yy h:mm'
    , 37: '#,##0 ;(#,##0)'
    , 38: '#,##0 ;[Red](#,##0)'
    , 39: '#,##0.00;(#,##0.00)'
    , 40: '#,##0.00;[Red](#,##0.00)'
    , 45: 'mm:ss'
    , 46: '[h]:mm:ss'
    , 47: 'mmss.0'
    , 48: '##0.0E+0'
    , 49: '@'
}

class StylesheetPart:
    @staticmethod
    def create_default():
        formats = []
        fonts = FontList([Font.create_default()])
        fills = [Fill.create_default()]
        borders = [Border.create_default()]
        style_xfs = [CellXF.create_default_named()]
        cell_xfs = [CellXF.create_default_inline()]
        styles = [StyleInfo.create_default()]
        return StylesheetPart(formats, fonts, fills, borders, style_xfs, cell_xfs, styles)

    @staticmethod
    def read(stream, for_update=False):
        rprocessor = RecordProcessor.resolve(stream)
        repository = RecordRepository(for_update)
        
        # Begin
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginStyleSheet:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginStyleSheet)
        
        r = rprocessor.read_descriptor()
        
        # Formats
        formats = []
        if r.rtype == BinaryRecordType.BrtBeginFmts:
            cfmts = struct.unpack('<I', rprocessor.read(4))[0]
            for i in range(cfmts):
                r = rprocessor.read_descriptor()
                # AC 2016 Formatting (Cheating by skipping w/repository)
                # Skip1
                if r.rtype == BinaryRecordType.BrtACBegin:
                    r.skip(rprocessor, repository)
                    r = rprocessor.read_descriptor()
                repository.push_current()
                
                if r.rtype != BinaryRecordType.BrtFmt:
                    raise UnexpectedRecordException(r, BinaryRecordType.BrtFmt)
                formats.append(NumberFormat.read(rprocessor, repository=repository))
                
                # Skip2
                r = rprocessor.read_descriptor()
                if r.rtype == BinaryRecordType.BrtACEnd:
                    r.skip(rprocessor, repository)
                    r = rprocessor.read_descriptor()
                repository.push_current()
            
            if r.rtype != BinaryRecordType.BrtEndFmts:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtEndFmts)
            r = rprocessor.read_descriptor()
        
        # Fonts
        if r.rtype == BinaryRecordType.BrtBeginFonts:
            fonts = FontList(repository=repository)
            cfonts = struct.unpack('<I', rprocessor.read(4))[0]
            for i in range(cfonts):
                r = rprocessor.read_descriptor()
                if r.rtype != BinaryRecordType.BrtFont:
                    raise UnexpectedRecordException(r, BinaryRecordType.BrtFont)
                fonts.append(Font.read(rprocessor))
            r = rprocessor.read_descriptor()
            
            # KnownFonts
            # Skip3
            if r.rtype == BinaryRecordType.BrtACBegin:
                r.skip(rprocessor, repository)
                r = rprocessor.skip_until(BinaryRecordType.BrtACEnd, repository=repository, skip_last=True)
            repository.push_current()
            
            if r.rtype != BinaryRecordType.BrtEndFonts:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtEndFonts)
            r = rprocessor.read_descriptor()
        else:
            fonts = FontList()
        
        # Fills
        fills = []
        if r.rtype == BinaryRecordType.BrtBeginFills:
            cfills = struct.unpack('<I', rprocessor.read(4))[0]
            for i in range(cfills):
                r = rprocessor.read_descriptor()
                if r.rtype != BinaryRecordType.BrtFill:
                    raise UnexpectedRecordException(r, BinaryRecordType.BrtFill)
                fills.append(Fill.read(rprocessor))
            
            r = rprocessor.read_descriptor()
            if r.rtype != BinaryRecordType.BrtEndFills:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtEndFills)
            r = rprocessor.read_descriptor()
        
        
        # Borders
        borders = []
        if r.rtype == BinaryRecordType.BrtBeginBorders:
            cborders = struct.unpack('<I', rprocessor.read(4))[0]
            for i in range(cborders):
                r = rprocessor.read_descriptor()
                if r.rtype != BinaryRecordType.BrtBorder:
                    raise UnexpectedRecordException(r, BinaryRecordType.BrtBorder)
                borders.append(Border.read(rprocessor))
            
            r = rprocessor.read_descriptor()
            if r.rtype != BinaryRecordType.BrtEndBorders:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtEndBorders)
            r = rprocessor.read_descriptor()
        
        
        # Cell Style XFs
        if r.rtype != BinaryRecordType.BrtBeginCellStyleXFs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginCellStyleXFs)
            
        cxfs = struct.unpack('<I', rprocessor.read(4))[0]
        style_xfs = []
        r = rprocessor.read_descriptor()
        for i in range(cxfs):
            if r.rtype == BinaryRecordType.BrtEndCellStyleXFs:
                break
            if r.rtype != BinaryRecordType.BrtXF:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtXF)
            
            style_xfs.append(CellXF.read(rprocessor, repository=repository))
            # Skip5
            r = rprocessor.skip_until(BinaryRecordType.BrtXF, BinaryRecordType.BrtEndCellStyleXFs, repository=repository)
            repository.push_current()
        
        if r.rtype != BinaryRecordType.BrtEndCellStyleXFs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtEndCellStyleXFs)
        r = rprocessor.read_descriptor()
        
        
        # Cell XFs
        if r.rtype != BinaryRecordType.BrtBeginCellXFs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginCellXFs)
            
        cxfs = struct.unpack('<I', rprocessor.read(4))[0]
        cell_xfs = []
        r = rprocessor.read_descriptor()
        for i in range(cxfs):
            if r.rtype == BinaryRecordType.BrtEndCellXFs:
                break
            if r.rtype != BinaryRecordType.BrtXF:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtXF)
            
            cell_xfs.append(CellXF.read(rprocessor, repository=repository))
            # Skip5
            r = rprocessor.skip_until(BinaryRecordType.BrtXF, BinaryRecordType.BrtEndCellXFs, repository=repository)
            repository.push_current()
        
        if r.rtype != BinaryRecordType.BrtEndCellXFs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtEndCellXFs)
        r = rprocessor.read_descriptor()
        
        
        # Styles
        if r.rtype != BinaryRecordType.BrtBeginStyles:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginStyles)
        
        styles = []
        cstyles = struct.unpack('<I', rprocessor.read(4))[0]
        r = rprocessor.read_descriptor()
        for i in range(cstyles):
            acuid = None
            if r.rtype == BinaryRecordType.BrtACBegin:
                acuid = ACUid.read(rprocessor)
                r = rprocessor.read_descriptor()
            
            if r.rtype != BinaryRecordType.BrtStyle:
                raise UnexpectedRecordException(r, BinaryRecordType.BrtStyle)
            
            # Skip6
            styles.append(StyleInfo.read(rprocessor, acuid, repository=repository))
            r = rprocessor.skip_until(BinaryRecordType.BrtStyle, BinaryRecordType.BrtACBegin, BinaryRecordType.BrtEndStyles, repository=repository)
            repository.push_current()
        
        if r.rtype != BinaryRecordType.BrtEndStyles:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtEndStyles)
        
        
        # Skip7
        r = rprocessor.skip_until(BinaryRecordType.BrtEndStyleSheet, repository=repository)
        repository.push_current()
        
        return StylesheetPart(formats, fonts, fills, borders, style_xfs, cell_xfs, styles, repository=repository)
        
    
    def __init__(self, formats, fonts, fills, borders, style_xfs, cell_xfs, styles, *, repository=None):
        self.formats = formats
        self.fonts = fonts
        self.fills = fills
        self.borders = borders
        self.style_xfs = style_xfs
        self.cell_xfs = cell_xfs
        self.styles = styles
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        repository = self.repository
        if repository:
            repository.begin_write()
        
        RecordDescriptor(BinaryRecordType.BrtBeginStyleSheet).write(rprocessor)
        
        formats = self.formats
        if len(formats):
            RecordDescriptor(BinaryRecordType.BrtBeginFmts, 4).write(rprocessor)
            rprocessor.write(struct.pack('<I', len(formats)))
            for fmt in formats:
                fmt.write(rprocessor)
            RecordDescriptor(BinaryRecordType.BrtEndFmts).write(rprocessor)
        
        self.fonts.write(rprocessor)
        
        fills = self.fills
        if len(fills):
            RecordDescriptor(BinaryRecordType.BrtBeginFills, 4).write(rprocessor)
            rprocessor.write(struct.pack('<I', len(fills)))
            for fill in fills:
                fill.write(rprocessor)
            RecordDescriptor(BinaryRecordType.BrtEndFills).write(rprocessor)
        
        borders = self.borders
        if len(borders):
            RecordDescriptor(BinaryRecordType.BrtBeginBorders, 4).write(rprocessor)
            rprocessor.write(struct.pack('<I', len(borders)))
            for border in borders:
                border.write(rprocessor)
            RecordDescriptor(BinaryRecordType.BrtEndBorders).write(rprocessor)
        
        style_xfs = self.style_xfs
        RecordDescriptor(BinaryRecordType.BrtBeginCellStyleXFs, 4).write(rprocessor)
        rprocessor.write(struct.pack('<I', len(style_xfs)))
        for style_xf in style_xfs:
            style_xf.write(rprocessor)
        RecordDescriptor(BinaryRecordType.BrtEndCellStyleXFs).write(rprocessor)
        
        cell_xfs = self.cell_xfs
        RecordDescriptor(BinaryRecordType.BrtBeginCellXFs, 4).write(rprocessor)
        rprocessor.write(struct.pack('<I', len(cell_xfs)))
        for cell_xf in cell_xfs:
            cell_xf.write(rprocessor)
        RecordDescriptor(BinaryRecordType.BrtEndCellXFs).write(rprocessor)
        
        
        styles = self.styles
        RecordDescriptor(BinaryRecordType.BrtBeginStyles, 4).write(rprocessor)
        rprocessor.write(struct.pack('<I', len(styles)))
        for style in styles:
            style.write(rprocessor)
        RecordDescriptor(BinaryRecordType.BrtEndStyles).write(rprocessor)
        
        
        # Skip7
        if repository:
            repository.write_poll(rprocessor)
        
        RecordDescriptor(BinaryRecordType.BrtEndStyleSheet).write(rprocessor)


class Color:
    @staticmethod
    def from_rgba(red, green, blue, alpha):
        return Color(ColorType.RGBA, 0, 0, True, red, green, blue, alpha)
    
    @staticmethod
    def from_rgba_packed(rgba_int32):
        red, green, blue, alpha = struct.pack('>I', rgba_int32)
        return Color.from_rgba(red, green, blue, alpha)
    
    @staticmethod
    def read(stream):
        flags, index = stream.read(2)
        f_valid_rgb = flags & 0x01
        x_color_type = (flags & 0xfe) >> 1
        
        n_tint_and_shade = struct.unpack('<h', stream.read(2))[0]
        b_red, b_green, b_blue, b_alpha = stream.read(4)
        
        color_type = ColorType(x_color_type)
        if color_type == ColorType.INDEX:
            index = PaletteColor(index)
        elif color_type == ColorType.THEME:
            index = ThemeColor(index)
        
        return Color(color_type, n_tint_and_shade, index, bool(f_valid_rgb), b_red, b_green, b_blue, b_alpha)
        
    def __init__(self, color_type, shade_amount, color_index, valid_rgba, red, green, blue, alpha):
        self.color_type = color_type
        self.shade_amount = shade_amount
        self.color_index = color_index
        self.valid_rgba = valid_rgba
        self._red = red
        self._green = green
        self._blue = blue
        self._alpha = alpha
    
    
    @property
    def red(self):
        return self._red
    
    @red.setter
    def red(self, value):
        self.color_type = ColorType.RGBA
        self._red = value
    
    @property
    def green(self):
        return self._green
    
    @green.setter
    def green(self, value):
        self.color_type = ColorType.RGBA
        self._green = value
    
    @property
    def blue(self):
        return self._blue
    
    @blue.setter
    def blue(self, value):
        self.color_type = ColorType.RGBA
        self._blue = value
    
    @property
    def alpha(self):
        return self._alpha
    
    @alpha.setter
    def alpha(self, value):
        self.color_type = ColorType.RGBA
        self._alpha = value
    
    
    def set_palette(self, palette_color):
        self.color_type = ColorType.PALETTE
        self.color_index = palette_color
        rgba = palette_color.get_rgba()
        if rgba is not None:
            self._set_rgba_int32(rgba)
        else:
            self.valid_rgba = False
    
    def set_theme(self, theme_color):
        self.color_type = ColorType.THEME
        self.color_index = theme_color
        self.valid_rgba = False
    
    
    
    def set_rgba(self, red, green, blue, alpha):
        self.color_type = ColorType.RGBA
        self.red = red
        self.green = green
        self.blue = blue
        self.alpha = alpha
    
    def set_rgba_packed(self, rgba_int32):
        self.color_type = ColorType.RGBA
        self._set_rgba_int32(rgba_int32)
    
    
    
    
    def _set_rgba_int32(self, rgba_int32):
        self.valid_rgba = True
        self.red, self.green, self.blue, self.alpha = struct.pack('>I', rgba_int32)
    
    
    
    def write(self, stream, write_header=False):
        rprocessor = RecordProcessor.resolve(stream)
        
        if write_header:
            RecordDescriptor(BinaryRecordType.BrtColor, len(self)).write(rprocessor)
        
        color_type = self.color_type
        if color_type == ColorType.RGBA or self.valid_rgba:
            rprocessor.write((color_type.value << 1) | 0x01)
        else:
            rprocessor.write(color_type.value << 1)
        
        color_index = self.color_index
        rprocessor.write(color_index.value if isinstance(color_index, (PaletteColor, ThemeColor)) else color_index)
        
        rprocessor.write(struct.pack('<h', self.shade_amount))
        rprocessor.write(bytes((self.red, self.green, self.blue, self.alpha)))
    
    def __str__(self):
        return f'Color: {self.color_type}; i: {self.color_index}; ARGB (valid): {self.alpha}, {self.red}, {self.green}, {self.blue} ({self.valid_rgba}); Shade: {self.shade_amount}'
    
    def __len__(self):
        return 8


class CellXF:
    @staticmethod
    def create_default_named():
        return CellXF(0xffff, 0, 0, 0, 0, 0, 0, HorizontalAlignmentType.GENERAL, VerticalAlignmentType.BOTTOM, False, False,
                False, False, ReadingOrderType.CONTEXT_DEPENDENT, True, False, False, False, 0x0000)
    
    @staticmethod
    def create_default_inline():
        return CellXF(0, 0, 0, 0, 0, 0, 0, HorizontalAlignmentType.GENERAL, VerticalAlignmentType.BOTTOM, False, False,
                False, False, ReadingOrderType.CONTEXT_DEPENDENT, True, False, False, False, 0x0000)
    
    @staticmethod
    def read(stream, *, repository=None):
        ixfe_parent, i_font, i_fmt, i_fill, ix_border = struct.unpack('<HHHHH', stream.read(10))
        trot, indent, flags_1, flags_2 = stream.read(4)
        
        alc = flags_1 & 0x07
        alcv = (flags_1 & 0x38) >> 3
        f_wrap = flags_1 & 0x40
        f_just_last = flags_1 & 0x80
        
        f_shrink_to_fit = flags_2 & 0x01
        f_merge_cell = flags_2 & 0x02
        i_reading_order = (flags_2 & 0x0c) >> 2
        f_locked = flags_2 & 0x10
        f_hidden = flags_2 & 0x20
        f_sx_button = flags_2 & 0x40
        f_123_prefix = flags_2 & 0x80
        
        flags_3 = struct.unpack('<H', stream.read(2))[0]
        xf_grbit_atr = flags_3 & 0x003f
        unused = flags_3 & 0xffc0
        
        return CellXF(ixfe_parent, i_font, i_fmt, i_fill, ix_border, trot, indent, HorizontalAlignmentType(alc), VerticalAlignmentType(alcv),
                bool(f_wrap), bool(f_just_last), bool(f_shrink_to_fit), bool(f_merge_cell), ReadingOrderType(i_reading_order), bool(f_locked), bool(f_hidden),
                bool(f_sx_button), bool(f_123_prefix), xf_grbit_atr, repository=repository)
    
    
    def __init__(self, parent_index, format_id, font_index, fill_index, border_index, text_trot,
            indent, horizontal_align_type, vertical_align_type, wrap_text, justify_on_last_line, shrink_to_fit,
            merged, reading_order_type, locked, hidden, has_pivot_table_dropdown, single_quote_prefix, gr_bit, *, repository=None):
        self.parent_index = parent_index
        self.format_id = format_id
        self.font_index = font_index
        self.fill_index = fill_index
        self.border_index = border_index
        self.text_trot = text_trot
        self.indent = indent
        self.horizontal_align_type = horizontal_align_type
        self.vertical_align_type = vertical_align_type
        self.wrap_text = wrap_text
        self.justify_on_last_line = justify_on_last_line
        self.shrink_to_fit = shrink_to_fit
        self.merged = merged
        self.reading_order_type = reading_order_type
        self.locked = locked
        self.hidden = hidden
        self.has_pivot_table_dropdown = has_pivot_table_dropdown
        self.single_quote_prefix = single_quote_prefix
        self.gr_bit = gr_bit
        self.repository = repository
    
    @property
    def is_style(self):
        return self.parent_index == 0xffff
    
    def should_ignore(self, xf_property):
        return self.is_style and self.gr_bit_val(xf_property)
    
    def set_ignore(self, xf_property, ignore):
        if not self.is_style:
            raise ValueError('CellXF is a (normal) Cell XF.')
        self.set_gr_bit_val(ignore)
    
    def should_persist(self, xf_property):
        return not self.is_style and self.gr_bit & xf_property.value
    
    def set_persist(self, xf_property, persist):
        if self.is_style:
            raise ValueError('CellXF is a Cell Style XF.')
        self.set_gr_bit_val(persist)
    
    def gr_bit_val(self, xf_property):
        return bool(self.gr_bit & xf_property.value)
    
    def set_gr_bit_val(self, xf_property, value):
        if value:
            self.gr_bit |= xf_property.value
        else:
            self.gr_bit &= xf_property.value ^ 0xffff
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        RecordDescriptor(BinaryRecordType.BrtXF, len(self)).write(rprocessor)
        
        rprocessor.write(struct.pack('<HHHHH', self.parent_index, self.format_id, self.font_index, self.fill_index, self.border_index))
        
        flags_1 = self.horizontal_align_type.value
        flags_1 |= self.vertical_align_type.value << 3
        if self.wrap_text:
            flags_1 |= 0x40
        if self.justify_on_last_line:
            flags_1 |= 0x80
        
        flags_2 = 0
        if self.shrink_to_fit:
            flags_2 |= 0x01
        if self.merged:
            flags_2 |= 0x02
        flags_2 |= self.reading_order_type.value << 2
        if self.locked:
            flags_2 |= 0x10
        if self.hidden:
            flags_2 |= 0x20
        if self.has_pivot_table_dropdown:
            flags_2 |= 0x40
        if self.single_quote_prefix:
            flags_2 |= 0x80
        
        rprocessor.write(bytes((self.text_trot, self.indent, flags_1, flags_2)))
        rprocessor.write(struct.pack('<H', self.gr_bit))
        
        # Skip5
        repository = self.repository
        if repository:
           repository.write_poll(rprocessor) 
        
    
    def __len__(self):
        return 16

    def __str__(self):
        result = [f'XF: {"Named Style" if self.is_style else "Inline Style"}']
        result.append(f'    parent_index: {self.parent_index} ({hex(self.parent_index)})')
        result.append(f'    format_id: {self.format_id} (gr_bit flag: {self.gr_bit_val(XFProperty.FMT)})')
        result.append(f'    font_index: {self.font_index} (gr_bit flag: {self.gr_bit_val(XFProperty.FONT)})')
        result.append(f'    fill_index: {self.fill_index} (gr_bit flag: {self.gr_bit_val(XFProperty.FILL)})')
        result.append(f'    border_index: {self.border_index} (gr_bit flag: {self.gr_bit_val(XFProperty.BORDER)})')
        result.append(f'    text_trot: {self.text_trot} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    indent: {self.indent} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    horizontal_align_type: {self.horizontal_align_type} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    vertical_align_type: {self.vertical_align_type} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    wrap_text: {self.wrap_text} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    justify_on_last_line: {self.justify_on_last_line} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    shrink_to_fit: {self.shrink_to_fit} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    merged: {self.merged} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    reading_order_type: {self.reading_order_type} (gr_bit flag: {self.gr_bit_val(XFProperty.ALIGNMENT)})')
        result.append(f'    locked: {self.locked} (gr_bit flag: {self.gr_bit_val(XFProperty.PROTECTION)})')
        result.append(f'    hidden: {self.hidden} (gr_bit flag: {self.gr_bit_val(XFProperty.PROTECTION)})')
        result.append(f'    has_pivot_table_dropdown: {self.has_pivot_table_dropdown}')
        result.append(f'    single_quote_prefix: {self.single_quote_prefix}')
        result.append(f'    (gr_bit): {hex(self.gr_bit)}')
        return '\n'.join(result)


class NumberFormat:
    @staticmethod
    def read(stream, *, repository=None):
        ifmt = struct.unpack('<H', stream.read(2))[0]
        st_fmt_code = RecordProcessor.resolve(stream).read_xl_w_string(False)
        return NumberFormat(ifmt, st_fmt_code, repository=repository)
    
    def __init__(self, format_id, format_str, *, repository=None):
        
        
        self.format_id = format_id
        self.format_str = format_str
        self.repository = repository
    
    def write(self, stream):
        format_id = self.format_id
        if not (5 <= format_id <= 8 or 23 <= format_id <= 26 or 41 <= format_id <= 44 or 63 <= format_id <= 66 or 164 <= format_id <= 382):
            raise ValueError(f'Format ID must be between one of the following ranges [5, 8] [23, 26] [41, 44] [63, 66] [164, 382]: {format_id}')
        
        format_str = self.format_str
        if not (1 <= len(format_str) <= 255):
            raise ValueError(f'Format string must between 1 and 255 characters: {format_str} ({len(format_str)} characters)')
            
        rprocessor = RecordProcessor.resolve(stream)
        repository = self.repository
        
        # Skip1
        if repository:
            repository.write_poll(rprocessor)
        
        RecordDescriptor(BinaryRecordType.BrtFmt, len(self)).write(rprocessor)
        rprocessor.write(struct.pack('<H', format_id))
        rprocessor.write_xl_w_string(format_str, False)
        
        # Skip2
        if repository:
            repository.write_poll(rprocessor)
    
    
    
    def __len__(self):
        return 2 + RecordProcessor.len_xl_w_string(self.format_str)




class FontList:
    def __init__(self, fonts=None, *, repository=None):
        self.fonts = [] if fonts is None else fonts
        self.repository = repository
    
    def append(self, font):
        self.fonts.append(font)
    
    def write(self, stream):
        fonts = self.fonts
        if len(fonts):
            rprocessor = RecordProcessor.resolve(stream)
            repository = self.repository
            
            RecordDescriptor(BinaryRecordType.BrtBeginFonts, 4).write(rprocessor)
            rprocessor.write(struct.pack('<I', len(fonts)))
            for font in fonts:
                font.write(rprocessor)
            
            # Skip3
            if repository:
                repository.write_poll(rprocessor)
            
            RecordDescriptor(BinaryRecordType.BrtEndFonts).write(rprocessor)
    
    def __getitem__(self, key):
        return self.fonts[key]
    
    def __iter__(self):
        return iter(self.fonts)
    
    def __len__(self):
        return len(self.fonts)

class Font:
    @staticmethod
    def create_default():
        return Font(height=220, italic=False, strikeout=False, outline_only=False, shadow=False, condense=False, extend=False, weight=400,
            subscript_type=SubscriptType.NONE, underline_type=UnderlineType.NONE, family=FontFamilyType.SWISS, char_set_type=CharacterSetType.ANSI_CHARSET,
            color=Color(ColorType.THEME, 0, ThemeColor.LT_1, True, 0, 0, 0, 255), font_scheme=FontSchemeType.MINOR, name='Calibri')
    
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        dy_height, gr_bit, bls, sss = struct.unpack('<HHHH', rprocessor.read(8))
        
        unused_1 = gr_bit & 0x0001
        f_italic = gr_bit & 0x0002
        unused_2 = gr_bit & 0x0004
        f_strikeout = gr_bit & 0x0008
        f_outline = gr_bit & 0x0010
        f_shadow = gr_bit & 0x0020
        f_condense = gr_bit & 0x0040
        f_extend = gr_bit & 0x0080
        unused_3 = gr_bit & 0xff00
        
        uls, b_family, b_char_set, unused = rprocessor.read(4)
        
        brt_color = Color.read(rprocessor)
        
        b_font_scheme = rprocessor.read(1)
        
        name = rprocessor.read_xl_w_string(False)
        
        return Font(dy_height, bool(f_italic), bool(f_strikeout), bool(f_outline), bool(f_shadow), bool(f_condense),
                bool(f_extend), bls, SubscriptType(sss), UnderlineType(uls), FontFamilyType(b_family), CharacterSetType(b_char_set),
                brt_color, FontSchemeType(b_font_scheme), name)
    
    def __init__(self, height, italic, strikeout, outline_only, shadow, condense, extend, weight, subscript_type, underline_type, family, 
            char_set_type, color, font_scheme, name):
        self.height = height
        self.italic = italic
        self.strikeout = strikeout
        self.outline_only = outline_only
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.weight = weight
        self.subscript_type = subscript_type
        self.underline_type = underline_type
        self.family = family
        self.char_set_type = char_set_type
        self.color = color
        self.font_scheme = font_scheme
        self.name = name
    
    def __str__(self):
        result = [f'Font: {self.name}']
        result.append(f'    height: {self.height}')
        result.append(f'    italic: {self.italic}')
        result.append(f'    strikeout: {self.strikeout}')
        result.append(f'    outline_only: {self.outline_only}')
        result.append(f'    shadow: {self.shadow}')
        result.append(f'    condense: {self.condense}')
        result.append(f'    extend: {self.extend}')
        result.append(f'    weight: {self.weight}')
        result.append(f'    subscript_type: {self.subscript_type}')
        result.append(f'    underline_type: {self.underline_type}')
        result.append(f'    family: {self.family}')
        result.append(f'    char_set_type: {self.char_set_type}')
        result.append(f'    color: {self.color}')
        result.append(f'    font_scheme: {self.font_scheme}')
        return '\n'.join(result)
    
    def write(self, stream):
        weight = self.weight
        if not 0x0190 <= weight <= 0x03e8
        
        rprocessor = RecordProcessor.resolve(stream)
        
        RecordDescriptor(BinaryRecordType.BrtFont, len(self)).write(rprocessor)
        
        gr_bit = 0
        if self.italic:
            gr_bit |= 0x0002
        if self.strikeout:
            gr_bit |= 0x0008
        if self.outline_only:
            gr_bit |= 0x0010
        if self.shadow:
            gr_bit |= 0x0020
        if self.condense:
            gr_bit |= 0x0040
        if self.extend:
            gr_bit |= 0x0080
        
        rprocessor.write(struct.pack('<HHHH', self.height, gr_bit, weight, self.subscript_type.value))
        rprocessor.write(bytes((self.underline_type.value, self.family.value, self.char_set_type.value, 0)))
        self.color.write(rprocessor, False)
        rprocessor.write(self.font_scheme.value)
        rprocessor.write_xl_w_string(self.name, False)
        
    
    def __len__(self):
        return 21 + RecordProcessor.len_xl_w_string(self.name)

class StyleInfo:
    @staticmethod
    def create_default():
        return StyleInfo(None, 0, True, False, False, 0, 0, 'Normal')

    @staticmethod
    def read(stream, acuid, *, repository=None):
        rprocessor = RecordProcessor.resolve(stream)
        
        ixf, gr_bit_obj1 = struct.unpack('<IH', rprocessor.read(6))
        f_built_in = gr_bit_obj1 & 0x0001
        f_hidden = gr_bit_obj1 & 0x0002
        f_custom = gr_bit_obj1 & 0x0004
        unused = gr_bit_obj1 & 0xfff8
        
        i_sty_built_in, i_level = rprocessor.read(2)
        
        st_name = rprocessor.read_xl_w_string()
        return StyleInfo(acuid, ixf, bool(f_built_in), bool(f_hidden), bool(f_custom), i_sty_built_in, i_level, st_name, repository=repository)
    
    def __init__(self, acuid, xf_index, built_in, hidden, built_in_customized, built_in_style_id, built_in_level, name, *, repository=None):
        self.acuid = acuid
        self.xf_index = xf_index
        self.built_in = built_in
        self.hidden = hidden
        self.built_in_customized = built_in_customized
        self.built_in_style_id = built_in_style_id
        self.built_in_level = built_in_level
        self.name = name
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        acuid = self.acuid
        if acuid:
            acuid.write(rprocessor)
        
        RecordDescriptor(BinaryRecordType.BrtStyle, len(self)).write(rprocessor)
        
        gr_bit_obj1 = 0
        if self.built_in:
            gr_bit_obj1 |= 0x0001
        if self.hidden:
            gr_bit_obj1 |= 0x0002
        if self.built_in_customized:
            gr_bit_obj1 |= 0x0004
        
        rprocessor.write(struct.pack('<IH', self.xf_index, gr_bit_obj1))
        rprocessor.write(bytes((self.built_in_style_id, self.built_in_level)))
        rprocessor.write_xl_w_string(self.name)
        
        # Skip6
        repository = self.repository
        if repository:
            repository.write_poll(rprocessor)
    
    
    def __len__(self):
        return 8 + RecordProcessor.len_xl_w_string(self.name)
    
    def __str__(self):
        result = [f'Style Info: {self.name}']
        result.append(f'    xf_index: {self.xf_index}')
        result.append(f'    built_in: {self.built_in}')
        result.append(f'    hidden: {self.hidden}')
        result.append(f'    built_in_customized: {self.built_in_customized}')
        result.append(f'    built_in_style_id: {self.built_in_style_id}')
        result.append(f'    built_in_level: {self.built_in_level}')
        return '\n'.join(result)

class Fill:
    @staticmethod
    def create_default():
        return Fill(fill_type=FillType.NONE, foreground_color=Color(ColorType.PALETTE, 0, PaletteColor.icvForeground, True, 0, 0, 0, 255), 
            background_color=Color(ColorType.PALETTE, 0, PaletteColor.icvBackground, True, 255, 255, 255, 255), gradient_type=GradientType.LINEAR, gradient_angle=0.0,
            gradient_fill_left=0.0, gradient_fill_right=0.0, gradient_fill_top=0.0, gradient_fill_bottom=0.0, gradient_stops=[])
    
    @staticmethod
    def read(stream):
        fls = struct.unpack('<I', stream.read(4))[0]
        
        brt_color_fore = Color.read(stream)
        brt_color_back = Color.read(stream)
        
        i_gradient_type, xnum_degree, xnum_fill_to_left, xnum_fill_to_right, xnum_fill_to_top, \
                xnum_fill_to_bottom, c_num_stop = struct.unpack('<IdddddI', stream.read(48))
        
        xfill_gradient_stops = []
        for i in range(c_num_stop):
            brt_color = Color.read(stream)
            xnum_position = struct.unpack('<d', stream.read(8))[0]
        
        return Fill(FillType(fls), brt_color_fore, brt_color_back, GradientType(i_gradient_type), xnum_degree, xnum_fill_to_left, xnum_fill_to_right,
                xnum_fill_to_top, xnum_fill_to_bottom, xfill_gradient_stops)
    
    def __init__(self, fill_type, foreground_color, background_color, gradient_type, gradient_angle, gradient_fill_left, gradient_fill_right, 
            gradient_fill_top, gradient_fill_bottom, gradient_stops):
        self.fill_type = fill_type
        self.foreground_color = foreground_color
        self.background_color = background_color
        self.gradient_type = gradient_type
        self.gradient_angle = gradient_angle
        self.gradient_fill_left = gradient_fill_left
        self.gradient_fill_right = gradient_fill_right
        self.gradient_fill_top = gradient_fill_top
        self.gradient_fill_bottom = gradient_fill_bottom
        self.gradient_stops = gradient_stops
    
    def write(self, stream):
        RecordDescriptor(BinaryRecordType.BrtFill, len(self)).write(stream)
        
        stream.write(struct.pack('<I', self.fill_type.value))
        self.foreground_color.write(stream)
        self.background_color.write(stream)
        
        gradient_stops = self.gradient_stops
        stream.write(struct.pack('<IdddddI', self.gradient_type.value, self.gradient_angle, self.gradient_fill_left, self.gradient_fill_right,
                self.gradient_fill_top, self.gradient_fill_bottom, len(gradient_stops)))
        
        for gradient_stop in gradient_stops:
            gradient_stop.write(stream)
    
    def __len__(self):
        return 4 + len(self.foreground_color) + len(self.background_color) + 48 + sum(len(i) for i in self.gradient_stops)
    
    def __str__(self):
        result = [f'Fill: {self.fill_type}']
        result.append(f'    foreground_color: {self.foreground_color}')
        result.append(f'    background_color: {self.background_color}')
        result.append(f'    gradient_type: {self.gradient_type}')
        result.append(f'    gradient_angle: {self.gradient_angle}')
        result.append(f'    gradient_fill_left: {self.gradient_fill_left}')
        result.append(f'    gradient_fill_right: {self.gradient_fill_right}')
        result.append(f'    gradient_fill_top: {self.gradient_fill_top}')
        result.append(f'    gradient_fill_bottom: {self.gradient_fill_bottom}')
        result.append(f'    gradient_stops: {", ".join(self.gradient_stops)}')
        return '\n'.join(result)

class GradientStop:
    @staticmethod
    def read(stream):
        brt_color = Color.read(stream)
        xnum_position = struct.unpack('<d', stream.read(8))[0]
        return GradientStop(brt_color, xnum_position)

    def __init__(self, color, position):
        self.color = color
        self.position = position
    
    def write(self, stream):
        self.color.write(stream)
        stream.write(struct.pack('<d', self.position))
    
    def __len__(self):
        return len(self.color) + 8
    
    def __str__(self):
        return f'(color: {self.color}, position: {self.position}'

class BorderDefinition:
    @staticmethod
    def read(stream):
        dg, reserved = stream.read(2)
        brt_color = Color.read(stream)
        
        return BorderDefinition(BorderType(dg), brt_color)
    
    def __init__(self, border_type, color):
        self.border_type = border_type
        self.color = color
    
    def write(self, stream):
        stream.write(bytes((self.border_type.value, 0)))
        self.color.write(stream)
    
    def __str__(self):
        return f'(border_type: {self.border_type}, color: {self.color})'

class Border:
    @staticmethod
    def create_default():
        return Border(has_diagonal_down=False, has_diagonal_up=False, top=BorderDefinition(BorderType.NONE, Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)), 
            bottom=BorderDefinition(BorderType.NONE, Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)),
            left=BorderDefinition(BorderType.NONE, Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)),
            right=BorderDefinition(BorderType.NONE, Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)),
            diagonal=BorderDefinition(BorderType.NONE, Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)))
    
    @staticmethod
    def read(stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        flags = rprocessor.read(1)
        f_bdr_diag_down = flags & 0x01
        f_bdr_diag_up = flags & 0x02
        reserved = 0xfc
        
        bxlf_top = BorderDefinition.read(stream)
        bxlf_bottom = BorderDefinition.read(stream)
        bxlf_left = BorderDefinition.read(stream)
        bxlf_right = BorderDefinition.read(stream)
        bxlf_diag = BorderDefinition.read(stream)
        
        return Border(bool(f_bdr_diag_down), bool(f_bdr_diag_up), bxlf_top, bxlf_bottom, bxlf_left, bxlf_right, bxlf_diag)
    
    def __init__(self, has_diagonal_down, has_diagonal_up, top, bottom, left, right, diagonal):
        self.has_diagonal_down = has_diagonal_down
        self.has_diagonal_up = has_diagonal_up
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.diagonal = diagonal
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        RecordDescriptor(BinaryRecordType.BrtBorder, len(self)).write(rprocessor)
        
        flags = 0
        if self.has_diagonal_down:
            flags |= 0x01
        if self.has_diagonal_up:
            flags |= 0x02
        rprocessor.write(flags)
        
        self.top.write(stream)
        self.bottom.write(stream)
        self.left.write(stream)
        self.right.write(stream)
        self.diagonal.write(stream)
    
    def __len__(self):
        return 51
    
    def __str__(self):
        result = ['Border:']
        result.append(f'    has_diagonal_down: {self.has_diagonal_down}')
        result.append(f'    has_diagonal_up: {self.has_diagonal_up}')
        result.append(f'    top: {self.top}')
        result.append(f'    bottom: {self.bottom}')
        result.append(f'    left: {self.left}')
        result.append(f'    right: {self.right}')
        result.append(f'    diagonal: {self.diagonal}')
        return '\n'.join(result)


