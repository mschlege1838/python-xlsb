
import sys
import os

from bprocessor import RecordProcessor


fname = sys.argv[1]
fsize = os.stat(fname).st_size

with open(fname, 'rb') as f:
    rreader = RecordProcessor(f)
    while f.tell() < fsize:
        r = rreader.read_descriptor()
        print(r)
        r.skip(rreader)