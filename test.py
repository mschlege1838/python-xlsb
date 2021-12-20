


import sys




if sys.argv[1] == 'r':
    from ooxmlpkg import ZipOfficeOpenXMLPackage
    from btypes import RelationshipType
    from part.workbook import WorkbookPart
    from part.sst import SharedStringsPart
    from part.worksheet import WorksheetPart, SharedStringCell
    
    with ZipOfficeOpenXMLPackage(sys.argv[2]) as pkg:
        pkg_info = pkg.get_part_info()
        
        wb_info = pkg.get_part_info(pkg_info.get_rel('Type', RelationshipType.WORKBOOK))
        with pkg.open_part(wb_info) as f:
            wb = WorkbookPart.read(f)
        with pkg.open_part(wb_info.get_rel('Type', RelationshipType.SHARED_STRINGS)) as f:
            sst = SharedStringsPart.read(f)
        
        for sheet_ref in wb.sheet_refs:
            print(sheet_ref.sheet_name)
            print('-------------------------------------------')
            with pkg.open_part(wb_info.get_rel('Id', sheet_ref.rel_id)) as f:
                sh = WorksheetPart.read(f)
                for row in sh.rows:
                    min_col, max_col = row.col_range
                    vals = list(None for i in range(max_col - min_col + 1))
                    
                    for cell in row.cells:
                        target = cell.header.column - min_col
                        if isinstance(cell, SharedStringCell):
                            vals[target] = sst.items[cell.str_index].val
                        else:
                            vals[target] = cell.value
                    print('\t'.join(str(v) for v in vals))
            print('-------------------------------------------')
            print()

elif sys.argv[1] == 'i':
    from ooxmlpkg import ZipOfficeOpenXMLPackage
    from bprocessor import RecordDescriptor
    from btypes import ContentType
    
    def iter_part(pkg, part_info):
        if isinstance(part_info.content_type, ContentType) and part_info.content_type not in (ContentType.CALCULATION_CHAIN, ContentType.THEME):
            print(part_info.path)
            print('-------------------------------------------')
            with pkg.open_part(part_info) as f:
                for part, data in RecordDescriptor.iter_parts(f):
                    print(part, data)
            print('-------------------------------------------')
            print()
            
        for rel in part_info.relationships:
            iter_part(pkg, pkg.get_part_info(rel))
    
    with ZipOfficeOpenXMLPackage(sys.argv[2]) as pkg:
        iter_part(pkg, pkg.get_part_info())
        


elif sys.argv[1] == 'w':
    from part.styles import StylesheetPart
    from part.workbook import WorkbookPart, BundledSheet, HiddenState
    from part.worksheet import WorksheetPart, ColInfo
    from ooxmlpkg import ZipOfficeOpenXMLPackage
    from btypes import RelationshipType, ContentType
    
    with ZipOfficeOpenXMLPackage(sys.argv[2], True) as pkg:
        ws = WorksheetPart.create_default()
        stylesheet = StylesheetPart.create_default()
        
        with pkg.open_part('/xl/worksheets/sheet1.bin', 'w', ContentType.WORKSHEET) as f:
            ws.write(f)
        with pkg.open_part('/xl/styles.bin', 'w', ContentType.STYLES) as f:
            stylesheet.write(f)
        
        with pkg.open_part_relationships('/xl/workbook.bin', 'w') as rels:
            sheet_rid = rels.add_rel(RelationshipType.WORKSHEET, 'worksheets/sheet1.bin')
        
        wb = WorkbookPart(None, [BundledSheet(HiddenState.VISIBLE, 1, sheet_rid, 'Sheet1')])
        with pkg.open_part('/xl/workbook.bin', 'w', ContentType.WORKBOOK) as f:
            wb.write(f)
        
        with pkg.open_part_relationships('/xl/workbook.bin', 'w') as rels:
            rels.add_rel(RelationshipType.STYLES, 'styles.bin')
        with pkg.open_part_relationships('/', 'w') as rels:
            rels.add_rel(RelationshipType.WORKBOOK, 'xl/workbook.bin')
    

elif sys.argv[1] == 'c':
    import os.path
    from bprocessor import RecordDescriptor
    
    tbl = []
    with open(sys.argv[2], 'rb') as f:
        for part, data in RecordDescriptor.iter_parts(f):
            tbl.append([(hex(f.tell()), f'{part}: {data}'), ''])
    
    
    tbl_i = 0
    with open(sys.argv[3], 'rb') as f:
        for part, data in RecordDescriptor.iter_parts(f):
            if tbl_i < len(tbl):
                tbl[tbl_i][1] = (hex(f.tell()), f'{part}: {data}')
            else:
                tbl.append(['', (hex(f.tell()), f'{part}: {data}')])
            tbl_i += 1
    
    with open(f'{os.path.split(sys.argv[2])[1]}-{os.path.split(sys.argv[3])[1]}.html', 'w') as f:
        f.write('<!DOCTYPE html>\n')
        f.write(f'<html><body><table><thead><tr><th>{sys.argv[2]}</th><th>{sys.argv[3]}</th></tr></thead><tbody>')
        for c1, c2 in tbl:
            st = 'background-color:yellow' if str(c1) != str(c2) else ''
            f.write(f'<tr><td style="{st}">{c1}</td><td style="{st}">{c2}</td></tr>')
        f.write('</tbody></table></body></html>')

elif sys.argv[1] == 'z':
    from ooxmlpkg import rec_zip
    
    rec_zip(sys.argv[2] + '.xlsb', sys.argv[2])

elif sys.argv[1] == 't':
    import shutil
    from ooxmlpkg import ZipOfficeOpenXMLPackage
    from btypes import RelationshipType
    from bprocessor import RecordDescriptor
    
    
    shutil.copy2(sys.argv[3], sys.argv[4])
    
    with ZipOfficeOpenXMLPackage(sys.argv[4]) as pkg:
        wb_ref = pkg.get_part_info().get_rel('Type', RelationshipType.WORKBOOK)
        with pkg.open_part(wb_ref) as f:
            parts = list(RecordDescriptor.iter_parts(f))
        
        if sys.argv[2] == 'r':
            for i, part in enumerate(parts):
                desc, data = part
                print(i, desc, data)
        elif sys.argv[2] == 'm':
            rm_targets = set(int(i) for i in sys.argv[5:])
            print(rm_targets)
            to_write = []
            for i, part in enumerate(parts):
                if i not in rm_targets:
                    to_write.append(part)
                else:
                    print(part[0], part[1])
            with pkg.open_part(wb_ref, 'w') as f:
                for desc, data in to_write:
                    desc.write(f)
                    f.write(data)
    
    