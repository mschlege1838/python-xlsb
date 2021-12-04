

import struct

from btypes import BinaryRecordType, HorizontalAlignmentType, VerticalAlignmentType, ReadingOrderType, XFProperty, ColorType, \
        IndexedColor, ThemeColor, SubscriptType, UnderlineType, FontFamilyType, CharacterSetType, FontSchemeType
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
        
        # Fonts
        r = rprocessor.read_descriptor()
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
        else:
            fonts = FontList()
        
        
        # Skip4
        r = rprocessor.skip_until(BinaryRecordType.BrtBeginCellStyleXFs, repository=repository)
        repository.push_current()
        
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
        
        return StylesheetPart(formats, fonts, style_xfs, cell_xfs, styles, repository=repository)
        
    
    def __init__(self, formats, fonts, style_xfs, cell_xfs, styles, *, repository=None):
        self.formats = formats
        self.fonts = fonts
        self.style_xfs = style_xfs
        self.cell_xfs = cell_xfs
        self.styles = styles
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        repository = self.repository
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
        
        # Skip4
        if repository:
            repository.write_poll(rprocessor)
        
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

class CellXF:
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
        return self.is_style and self.gr_bit & xf_property.value
    
    def set_ignore(self, xf_property, ignore):
        if not self.is_style:
            raise ValueError('CellXF is a (normal) Cell XF.')
        if ignore:
            self.gr_bit |= xf_property.value
        else:
            self.gr_bit &= xf_property.value ^ 0xffff
    
    def should_persist(self, xf_property):
        return not self.is_style and self.gr_bit & xf_property.value
    
    def set_persist(self, xf_property, persist):
        if self.is_style:
            raise ValueError('CellXF is a Cell Style XF.')
        if persist:
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

class NumberFormat:
    @staticmethod
    def read(stream, *, repository=None):
        ifmt = struct.unpack('<H', stream.read(2))[0]
        st_fmt_code = RecordProcessor.resolve(stream).read_xl_w_string(False)
        return NumberFormat(ifmt, st_fmt_code, repository=repository)
    
    def __init__(self, format_id, format_str, *, repository=None):
        if not (5 <= format_id <= 8 or 23 <= format_id <= 26 or 41 <= format_id <= 44 or 63 <= format_id <= 66 or 164 <= format_id <= 382):
            raise ValueError(f'Format ID must be between one of the following ranges [5, 8] [23, 26] [41, 44] [63, 66] [164, 382]: {format_id}')
        if not (1 <= len(format_str) <= 255):
            raise ValueError(f'Format string must between 1 and 255 characters: {format_str} ({len(format_str)} characters)')
        
        self.format_id = format_id
        self.format_str = format_str
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        repository = self.repository
        
        # Skip1
        if repository:
            repository.write_poll(rprocessor)
        
        RecordDescriptor(BinaryRecordType.BrtFmt, len(self)).write(rprocessor)
        rprocessor.write(struct.pack('<H', self.format_id))
        rprocessor.write_xl_w_string(self.format_str, False)
        
        # Skip2
        if repository:
            repository.write_poll(rprocessor)
    
    def __len__(self):
        return 2 + RecordProcessor.len_xl_w_string(self.format_str)




class FontList:
    def __init__(self, *, repository=None):
        self.fonts = []
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
        
        
        

class Font:
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
    
    def __init__(self, height, italic, strikeout, outline, shadow, condense, extend, weight,
            subscript_type, underline_type, family, char_set_type, color, font_scheme, name):
        self.height = height
        self.italic = italic
        self.strikeout = strikeout
        self.outline = outline
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
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        RecordDescriptor(BinaryRecordType.BrtFont, len(self)).write(rprocessor)
        
        gr_bit = 0
        if self.italic:
            gr_bit |= 0x0002
        if self.strikeout:
            gr_bit |= 0x0008
        if self.outline:
            gr_bit |= 0x0010
        if self.shadow:
            gr_bit |= 0x0020
        if self.condense:
            gr_bit |= 0x0040
        if self.extend:
            gr_bit |= 0x0080
        
        rprocessor.write(struct.pack('<HHHH', self.height, gr_bit, self.weight, self.subscript_type.value))
        rprocessor.write(bytes((self.underline_type.value, self.family.value, self.char_set_type.value, 0)))
        self.color.write(rprocessor, False)
        rprocessor.write(self.font_scheme.value)
        rprocessor.write_xl_w_string(self.name, False)
        
    
    def __len__(self):
        return 21 + RecordProcessor.len_xl_w_string(self.name)




class Color:
    @staticmethod
    def read(stream):
        flags, index = stream.read(2)
        f_valid_rgb = flags & 0x01
        x_color_type = (flags & 0xfe) >> 1
        
        n_tint_and_shade = struct.unpack('<h', stream.read(2))[0]
        b_red, b_green, b_blue, b_alpha = stream.read(4)
        
        color_type = ColorType(x_color_type)
        if color_type == ColorType.INDEX:
            index = IndexedColor(index)
        elif color_type == ColorType.THEME:
            index = ThemeColor(index)
        
        return Color(color_type, n_tint_and_shade, index, bool(f_valid_rgb), b_red, b_green, b_blue, b_alpha)
        
    def __init__(self, color_type, shade_amount, color_index, valid_argb, red, green, blue, alpha):
        self.color_type = color_type
        self.shade_amount = shade_amount
        self.color_index = color_index
        self.valid_argb = valid_argb
        self.red = red
        self.green = green
        self.blue = blue
        self.alpha = alpha
    
    def write(self, stream, write_header):
        rprocessor = RecordProcessor.resolve(stream)
        
        if write_header:
            RecordDescriptor(BinaryRecordType.BrtColor, len(self)).write(rprocessor)
        
        color_type = self.color_type
        if color_type == ColorType.ARGB or self.valid_argb:
            rprocessor.write((color_type.value << 1) | 0x01)
        else:
            rprocessor.write(color_type.value << 1)
        
        rprocessor.write(self.color_index.value)
        
        rprocessor.write(struct.pack('<h', self.shade_amount))
        rprocessor.write(bytes((self.red, self.green, self.blue, self.alpha)))

class StyleInfo:
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