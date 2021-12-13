
import struct
from enum import Enum

from btypes import BinaryRecordType
from bprocessor import UnexpectedRecordException, RecordProcessor, RecordRepository, RecordDescriptor


class WorkbookPart:
    @staticmethod
    def read(stream, for_update=False):
        rprocessor = RecordProcessor.resolve(stream)
        
        # Begin
        r = rprocessor.read_descriptor()
        if r.rtype != BinaryRecordType.BrtBeginBook:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtBeginBook)
        
        repository = RecordRepository(for_update)
        
        # Skip 1
        r = rprocessor.skip_until(BinaryRecordType.BrtWbProp, BinaryRecordType.BrtBeginBundleShs, repository=repository)
        repository.push_current()
        
        # Workbook Properties
        props = None
        if r.rtype == BinaryRecordType.BrtWbProp:
            props = WorkbookProperties.read(rprocessor, repository=repository)
            # Skip 3
            r = rprocessor.skip_until(BinaryRecordType.BrtBeginBundleShs, repository=repository)
            repository.push_current()
        
        # Sheet References
        sheet_refs = []
        r = rprocessor.read_descriptor()
        while r.rtype == BinaryRecordType.BrtBundleSh:
            sheet_refs.append(BundledSheet.read(rprocessor))
            r = rprocessor.read_descriptor()
        
        if r.rtype != BinaryRecordType.BrtEndBundleShs:
            raise UnexpectedRecordException(r, BinaryRecordType.BrtEndBundleShs)
        
        
        # Skip 2
        rprocessor.skip_until(BinaryRecordType.BrtEndBook, repository=repository)
        repository.push_current()
        
        return WorkbookPart(props, sheet_refs, repository=repository)
    
    def __init__(self, props, sheet_refs, *, repository=None):
        self.props = props
        self.sheet_refs = sheet_refs
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        repository = self.repository
        if repository:
            repository.begin_write()
        
        
        # Begin
        RecordDescriptor(BinaryRecordType.BrtBeginBook).write(rprocessor)
        
        # Skip 1
        if repository:
            repository.write_poll(rprocessor)
        
        props = self.props
        if props:
            props.write(rprocessor)
        
        # Sheet References
        RecordDescriptor(BinaryRecordType.BrtBeginBundleShs).write(rprocessor)
        for sheet_ref in self.sheet_refs:
            RecordDescriptor(BinaryRecordType.BrtBundleSh, len(sheet_ref)).write(rprocessor)
            sheet_ref.write(rprocessor)
        RecordDescriptor(BinaryRecordType.BrtEndBundleShs).write(rprocessor)
        
        # Skip 2
        if repository:
            repository.write_poll(rprocessor)
        
        # End
        RecordDescriptor(BinaryRecordType.BrtEndBook).write(rprocessor)
        if repository:
            repository.close()
        
    def __len__(self):
        result = len(self.repository)
        for sheet_ref in self.sheet_refs:
            result += len(sheet_ref)
        return result
        
    

class HiddenState(Enum):
    VISIBLE = 0
    HIDDEN = 1
    VERY_HIDDEN = 2
    
class BundledSheet:
    
    @staticmethod
    def read(rprocessor):
        hs_state = struct.unpack('<I', rprocessor.read(4))[0]
        i_tab_id = struct.unpack('<I', rprocessor.read(4))[0]
        rel_id = rprocessor.read_xl_w_string()
        sheet_name = rprocessor.read_xl_w_string(False)
        return BundledSheet(HiddenState(hs_state), i_tab_id, rel_id, sheet_name)
    
    @staticmethod
    def check_sheet_name(name):
        for reserved in ('\0', '\u0003', ':', '\\', '*', '?', '/', '[', ']'):
            if reserved in name:
                raise ValueError(f'Sheet name cannot contain "{reserved}": {name}')
        if name.startswith("'") or name.endswith("'"):
            raise ValueError(f'Sheet name cannot start or end with "\'": {name}')
    
    def __init__(self, hidden_state, tab_id, rel_id, sheet_name):
        self.hidden_state = hidden_state
        self.tab_id = tab_id
        RecordProcessor.check_rel_id(rel_id)
        self._rel_id = rel_id
        BundledSheet.check_sheet_name(sheet_name)
        self._sheet_name = sheet_name
    
    @property
    def rel_id(self):
        return self._rel_id
    
    @rel_id.setter
    def rel_id(self, value):
        RecordProcessor.check_rel_id(value)
        self._rel_id = value
    
    @property
    def sheet_name(self):
        return self._sheet_name
    
    @sheet_name.setter
    def sheet_name(self, value):
        BundledSheet.check_sheet_name(value)
        self._sheet_name = value
    
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        rprocessor.write(struct.pack('<I', self.hidden_state.value))
        rprocessor.write(struct.pack('<I', self.tab_id))
        rprocessor.write_xl_w_string(self._rel_id)
        rprocessor.write_xl_w_string(self._sheet_name, False)
    
    def __len__(self):
        return 8 + RecordProcessor.len_xl_w_string(self.rel_id) + RecordProcessor.len_xl_w_string(self.sheet_name, False)


class UpdateLinksBehavior(Enum):
    APP_SPECIFIC = 0x0
    MANUAL_UPDATE = 0x1
    AUTO_UPDATE = 0x2

class PublishType(Enum):
    SHEET_LEVEL = 0x0
    ITEM_LEVEL = 0x1

class ShapeDisplayType(Enum):
    VISIBLE = 0x0
    PLACEHOLDERS = 0x1
    INVISIBLE = 0x2



class WorkbookProperties:
    @staticmethod
    def create_default():
        return WorkbookProperties(False, False, False, False, True, False, False, UpdateLinksBehavior.APP_SPECIFIC, False, PublishType.SHEET_LEVEL,
                False, ShapeDisplayType.VISIBLE, False, True, False, 166925, '')
    
    @staticmethod
    def read(stream, *, repository=None):
        rprocessor = RecordProcessor.resolve(stream)
        
        flags, theme_version = struct.unpack('<II', rprocessor.read(8))
        
        f_1904 = flags & 0x00000001
        reserved_1 = flags & 0x00000002
        f_hide_border_unsel_lists = flags & 0x00000004
        f_filter_privacy = flags & 0x00000008
        f_bugged_user_about_solution = flags & 0x00000010
        f_show_ink_annotation = flags & 0x00000020
        f_backup = flags & 0x00000040
        f_no_save_sup = flags & 0x00000080
        update_links = (flags & 0x00000300) >> 8
        f_hide_pivot_table_flist = flags & 0x00000400
        f_published_book_items = flags & 0x00000800
        f_check_compat = flags & 0x00001000
        md_dsp_obj = (flags & 0x00006000) >> 13
        f_show_pivot_chart_filter = flags & 0x00008000
        f_auto_compress_picture = flags & 0x00010000
        reserved_2 = flags & 0x00020000
        f_refresh_all = flags & 0x00040000
        unused = flags & 0xfff80000
        
        str_name = rprocessor.read_xl_w_string(False)
        
        return WorkbookProperties(bool(f_1904), bool(f_hide_border_unsel_lists), bool(f_filter_privacy), bool(f_bugged_user_about_solution), bool(f_show_ink_annotation),
                bool(f_backup), bool(f_no_save_sup), UpdateLinksBehavior(update_links), bool(f_hide_pivot_table_flist), PublishType(f_published_book_items), bool(f_check_compat),
                ShapeDisplayType(md_dsp_obj), bool(f_show_pivot_chart_filter), bool(f_auto_compress_picture), bool(f_refresh_all), theme_version, str_name, repository=repository)
    
    def __init__(self, uses_legacy_date_format, hide_borders_for_inactive_tables, has_pii, request_smart_doc_warning, show_ink_comments, enable_backup,
            cache_external_link_values, update_links_behavior, hide_pivot_table_fields, publish_type, check_compatiblity, shape_display_type, show_pivot_chart_filter,
            auto_compress_pictures, refresh_all_on_open, theme_version, vba_name, *, repository=None):
        self.uses_legacy_date_format = uses_legacy_date_format
        self.hide_borders_for_inactive_tables = hide_borders_for_inactive_tables
        self.has_pii = has_pii
        self.request_smart_doc_warning = request_smart_doc_warning
        self.show_ink_comments = show_ink_comments
        self.enable_backup = enable_backup
        self.cache_external_link_values = cache_external_link_values
        self.update_links_behavior = update_links_behavior
        self.hide_pivot_table_fields = hide_pivot_table_fields
        self.publish_type = publish_type
        self.check_compatiblity = check_compatiblity
        self.shape_display_type = shape_display_type
        self.show_pivot_chart_filter = show_pivot_chart_filter
        self.auto_compress_pictures = auto_compress_pictures
        self.refresh_all_on_open = refresh_all_on_open
        self.theme_version = theme_version
        self.vba_name = vba_name
        self.repository = repository
    
    def write(self, stream):
        rprocessor = RecordProcessor.resolve(stream)
        
        RecordDescriptor(BinaryRecordType.BrtWbProp, len(self)).write(rprocessor)
        
        flags = 0x00000000
        if self.uses_legacy_date_format:
            flags |= 0x00000001
        if self.hide_borders_for_inactive_tables:
            flags |= 0x00000004
        if self.has_pii:
            flags |= 0x00000008
        if self.request_smart_doc_warning:
            flags |= 0x00000010
        if self.show_ink_comments:
            flags |= 0x00000020
        if self.enable_backup:
            flags |= 0x00000040
        if self.cache_external_link_values:
            flags |= 0x00000080
        flags |= (self.update_links_behavior.value << 8)
        if self.hide_pivot_table_fields:
            flags |= 0x00000400
        flags |= (self.publish_type.value << 11)
        if self.check_compatiblity:
            flags |= 0x00001000
        flags |= (self.shape_display_type.value << 13)
        if self.show_pivot_chart_filter:
            flags |= 0x00008000
        if self.auto_compress_pictures:
            flags |= 0x00010000
        if self.refresh_all_on_open:
            flags |= 0x00040000
        
        rprocessor.write(struct.pack('<II', flags, self.theme_version))
        rprocessor.write_xl_w_string(self.vba_name, False)
        
        # Skip 3
        repository = self.repository
        if repository:
            repository.write_poll(rprocessor)
    
    def __len__(self):
        return 8 + RecordProcessor.len_xl_w_string(self.vba_name, False)
    
    def __str__(self):
        result = ['Workbook Properties']
        result.append(f'    uses_legacy_date_format: {self.uses_legacy_date_format}')
        result.append(f'    hide_borders_for_inactive_tables: {self.hide_borders_for_inactive_tables}')
        result.append(f'    has_pii: {self.has_pii}')
        result.append(f'    request_smart_doc_warning: {self.request_smart_doc_warning}')
        result.append(f'    show_ink_comments: {self.show_ink_comments}')
        result.append(f'    enable_backup: {self.enable_backup}')
        result.append(f'    cache_external_link_values: {self.cache_external_link_values}')
        result.append(f'    update_links_behavior: {self.update_links_behavior}')
        result.append(f'    hide_pivot_table_fields: {self.hide_pivot_table_fields}')
        result.append(f'    publish_type: {self.publish_type}')
        result.append(f'    check_compatiblity: {self.check_compatiblity}')
        result.append(f'    shape_display_type: {self.shape_display_type}')
        result.append(f'    show_pivot_chart_filter: {self.show_pivot_chart_filter}')
        result.append(f'    auto_compress_pictures: {self.auto_compress_pictures}')
        result.append(f'    refresh_all_on_open: {self.refresh_all_on_open}')
        result.append(f'    theme_version: {self.theme_version}')
        result.append(f'    vba_name: {self.vba_name}')
        return '\n'.join(result)