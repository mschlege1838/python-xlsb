

import sys
import struct

from ooxmlpkg import ZipOfficeOpenXMLPackage
from part.worksheet import WorksheetPart

from bprocessor import RecordProcessor


with ZipOfficeOpenXMLPackage(sys.argv[1]) as pkg:
    with pkg.open_part('xl/worksheets/sheet1.bin') as f:
        ws = WorksheetPart.read(f)
        for row in ws.rows:
            for cell in row.cells:
                print(cell.value, end='\t')
        
        # rp = RecordProcessor(f)
        # while True:
            # r = rp.read_descriptor()
            # print(r)
            # r.skip(rp)

