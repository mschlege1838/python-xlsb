
from btypes import HorizontalAlignmentType, VerticalAlignmentType, ReadingOrderType, XFProperty, ColorType, \
        PaletteColor, ThemeColor, SubscriptType, UnderlineType, FontFamilyType, CharacterSetType, FontSchemeType, FillType, GradientType, \
        BorderType

from part.styles import Color

class Styleable:
    def __init__(self):
        self.number_format = ''
        self.font_height = 220
        self.italic = False
        self.strikeout = False
        self.font_weight = 400
        self.subscript_type = SubscriptType.NONE
        self.underline_type = UnderlineType.NONE
        self.font_family = FontFamilyType.SWISS
        self.charset = CharacterSetType.ANSI_CHARSET
        self.text_color = Color(ColorType.THEME, 0, ThemeColor.LT_1, True, 0, 0, 0, 255)
        self.font_scheme = FontSchemeType.MINOR
        self.font_name = 'Calibri'
        self.fill_type = FillType.NONE
        self.fill_fg_color = Color(ColorType.INDEX, 0, PaletteColor.icvForeground, True, 0, 0, 0, 255)
        self.fill_bg_color = Color(ColorType.INDEX, 0, PaletteColor.icvBackground, True, 255, 255, 255, 255)
        self.fill_grad_type = GradientType.LINEAR
        self.fill_grad_angle = 0.0
        self.fill_grad_left = 0.0
        self.fill_grad_right = 0.0
        self.fill_grad_top = 0.0
        self.fill_grad_bottom = 0.0
        self.fill_grad_stops = []
        self.border_top_type = BorderType.NONE
        self.border_top_color = Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)
        self.border_bottom_type = BorderType.NONE
        self.border_bottom_color = Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)
        self.border_left_type = BorderType.NONE
        self.border_left_color = Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)
        self.border_right_type = BorderType.NONE
        self.border_right_color = Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)
        self.border_diag_type = BorderType.NONE
        self.border_diag_color = Color(ColorType.AUTO, 0, 0, True, 0, 0, 0, 0)
        self.has_diag_down_border = False
        self.has_diag_up_border = False
        self.text_trot = 0
        self.indent = 0
        self.horizontal_alignment = HorizontalAlignmentType.GENERAL
        self.vertical_alignment = VerticalAlignmentType.BOTTOM
        self.shrink_to_fit = False
        self.reading_order_type = ReadingOrderType.CONTEXT_DEPENDENT


class Workbook:
    def __init__(self):
        self.sheets = {}