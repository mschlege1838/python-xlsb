
from enum import Enum

from btypes import HorizontalAlignmentType, VerticalAlignmentType, ReadingOrderType, XFProperty, ColorType, \
        PaletteColor, ThemeColor, SubscriptType, UnderlineType, FontFamilyType, CharacterSetType, FontSchemeType, FillType, GradientType, \
        BorderType

from part.styles import Color, num_format_lu_all_langs
from part.worksheet import SharedStringCell


class FontWeight(Enum):
    NORMAL = 400
    MEDIUM = 500
    SEMI_BOLD = 600
    BOLD = 700
    EXTRABOLD = 800
    HEAVY = 900
    MAX = 1000
    
    @staticmethod
    def resolve_i(font_weight_i):
        itr = iter(FontWeight)
        current = next(itr)
        for subsequent in itr:
            if current.value <= font_weight_i < subsequent.value:
                return current
            current = subsequent
        return FontWeight.MAX

class Styleable:
    def __init__(self, format_part=None, font_part=None, border_part=None, fill_part=None, xf_part=None):
        self._initialized = self._dirty = False
        
        if format_part:
            self.number_format = format_part.format_str
        elif xf_part and xf_part.format_id in num_format_lu_all_langs:
            self.number_format = num_format_lu_all_langs[xf_part.format_id]
        else:
            self.number_format = ''
        
        self.font_height = font_part.height if font_part else 220
        self.italic = font_part.italic if font_part else False
        self.strikeout = font_part.strikeout if font_part else False
        self.font_weight_i = font_part.weight if font_part else 400
        self.subscript_type = font_part.subscript_type if font_part else SubscriptType.NONE
        self.underline_type = font_part.underline_type if font_part else UnderlineType.NONE
        self.font_family = font_part.family if font_part else FontFamilyType.SWISS
        self.charset = font_part.char_set_type if font_part else CharacterSetType.ANSI_CHARSET
        self.text_color = font_part.color if font_part else Color(ColorType.THEME, 0, ThemeColor.LT_1, True, 0, 0, 0, 255)
        self.font_scheme = font_part.font_scheme if font_part else FontSchemeType.MINOR
        self.font_name = font_part.name if font_part else 'Calibri'
        self.fill_type = fill_part.fill_type if fill_part else FillType.NONE
        self.fill_fg_color = fill_part.foreground_color if fill_part else Color(ColorType.INDEX, 0, PaletteColor.icvForeground, True, 0, 0, 0, 255)
        self.fill_bg_color = fill_part.background_color if fill_part else Color(ColorType.INDEX, 0, PaletteColor.icvBackground, True, 255, 255, 255, 255)
        self.fill_grad_type = fill_part.gradient_type if fill_part else GradientType.LINEAR
        self.fill_grad_angle = fill_part.gradient_angle if fill_part else 0.0
        self.fill_grad_left = fill_part.gradient_fill_left if fill_part else 0.0
        self.fill_grad_right = fill_part.gradient_fill_right if fill_part else 0.0
        self.fill_grad_top = fill_part.gradient_fill_top if fill_part else 0.0
        self.fill_grad_bottom = fill_part.gradient_fill_bottom if fill_part else 0.0
        self.fill_grad_stops = list(fill_part.gradient_stops) if fill_part else []
        
        for prefix, pname in (('border_top_', 'top'), ('border_bottom_', 'bottom'), ('border_right_', 'right'), ('border_left_', 'left'), ('border_diag_', 'diagonal')):
            setattr(self, f'{prefix}_type', getattr(border_part, pname).border_type if border_part else BorderType.NONE)
            setattr(self, f'{prefix}_color', getattr(border_part, pname).color if border_part else Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0))

        self.has_diag_down_border = border_part.has_diagonal_down if border_part else False
        self.has_diag_up_border = border_part.has_diagonal_up if border_part else False
        self.text_trot = xf_part.text_trot if xf_part else 0
        self.indent = xf_part.indent if xf_part else 0
        self.horizontal_alignment = xf_part.horizontal_alignment if xf_part else HorizontalAlignmentType.GENERAL
        self.vertical_alignment = xf_part.vertical_alignment if xf_part else VerticalAlignmentType.BOTTOM
        self.shrink_to_fit = xf_part.shrink_to_fit if xf_part else False
        self.reading_order_type = xf_part.reading_order_type if xf_part else ReadingOrderType.CONTEXT_DEPENDENT
        
        self.format_part = format_part
        self.font_part = font_part
        self.border_part = border_part
        self.fill_part = fill_part
        self.xf_part = xf_part
        
        self._initialized = True
    
    @property
    def font_weight(self):
        return FontWeight.resolve(self.font_weight_i)
    
    @font_weight.setter
    def font_weight(self, value):
        self.font_weight_i = value.value
    
    @property
    def part_list(self):
        return ['format_part', 'font_part', 'border_part', 'fill_part', 'xf_part']
    
    def __setattr__(self, name, value):
        object.__setattr__(self, name, value)
        
        if name in ('_initialized', '_dirty'):
            return
        
        if self._initialized and not self._dirty:
            part_list = self.part_list
            if name in part_list:
                return
            
            for pname in part_list:
                object.__setattr__(self, pname, None)
            
            object.__setattr__(self, '_dirty', True)


class Column(Styleable):
    def __init__(self, column_index, col_info_part=None, format_part=None, font_part=None, border_part=None, fill_part=None, xf_part=None):
        pass

class Cell(Styleable):
    def __init__(self, row_header_part=None, cell_part=None, shared_strings_part=None, format_part=None, font_part=None, border_part=None, fill_part=None, xf_part=None):
        self.row_index = None
        self.column_index = None
        self.value = None
        
        rich_str_part = self.rich_str_part = None
        self.cell_part = cell_part
        
        if isinstance(cell_part, SharedStringCell):
            if not shared_strings_part:
                raise ValueError('Shared strings part must be provided for SharedStringCells.')
            rich_str_part = self.rich_str_part = shared_strings_part[cell_part.str_index]
            self.value = rich_str_part.val
        elif cell_part:
            self.value = cell.value
        
        if cell_part:
            if not row_header_part:
                raise ValueError('Row header must be provided with cell part.')
            self.row_index = row_header_part.row_index
            self.column_index = cell_part.header.column
        
        super().__init__(format_part, font_part, border_part, fill_part, xf_part)
        
    
    @property
    def part_list(self):
        return super().part_list + ['cell_part', 'rich_str_part']

