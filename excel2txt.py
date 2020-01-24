#!/usr/bin/env python3
"""
Author : Ken Youens-Clark <kyclark@gmail.com>
Date   : 2019-09-26
Purpose: Convert Excel files to delimited text
"""

import argparse
import os
import re
import sys
from openpyxl import load_workbook
from typing import Any, Optional


# --------------------------------------------------
def get_args():
    """Get command-line arguments"""

    parser = argparse.ArgumentParser(
        description='Convert Excel files to delimited text',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)

    parser.add_argument('file',
                        metavar='FILE',
                        type=argparse.FileType('r'),
                        nargs='+',
                        help='Input Excel file(s)')

    parser.add_argument('-o',
                        '--outdir',
                        help='Output directory',
                        metavar='str',
                        type=str,
                        default=os.getcwd())

    parser.add_argument('-d',
                        '--delimiter',
                        help='Delimiter for output file',
                        metavar='str',
                        type=str,
                        default='\t')

    parser.add_argument('-D',
                        '--mkdirs',
                        help='Create directories for output files',
                        action='store_true')

    parser.add_argument('-n',
                        '--normalize',
                        help='Normalize headers',
                        action='store_true')

    args = parser.parse_args()

    if not os.path.isabs(args.outdir):
        args.outdir = os.path.abspath(args.outdir)

    return args


# --------------------------------------------------
def main():
    """Make a jazz noise here"""

    args = get_args()

    if not os.path.isdir(args.outdir):
        os.makedirs(args.outdir)

    for i, fh in enumerate(args.file, start=1):
        print(f'{i:3}: {fh.name}')
        if not process(fh, args):
            print(f'Something amiss with {fh.name}', file=sys.stderr)

    print('Done.')
    return 0


# --------------------------------------------------
def process(fh, args):
    """Process a file"""

    file = fh.name
    fh.close()
    basename, _ = os.path.splitext(os.path.basename(file))
    wb = load_workbook(file)
    delimiter = args.delimiter

    for ws in wb:
        ws_name = normalize(ws.title)
        if not ws_name:
            continue

        out_file = '__'.join([basename, ws_name]) + '.txt'
        out_dir = args.outdir
        if args.mkdirs:
            out_dir = os.path.join(out_dir, basename)
            if not os.path.isdir(out_dir):
                os.makedirs(out_dir)

        out_path = os.path.join(out_dir, out_file)

        with open(out_path, 'wt') as out_fh:
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            headers = []
            for i, row in enumerate(rows):
                if i == 0:
                    headers = list(row)
                    while headers and headers[-1] is None:
                        headers.pop()
                    row = list(map(normalize, headers))

                if headers:
                    data = row[:len(headers) - 1]
                    if all(map(lambda x: x is None, data)):
                        break

                    out_fh.write(delimiter.join(map(cell_norm, data)) + '\n')

        # Remove empty worksheets
        if os.path.getsize(out_path) == 0:
            os.remove(out_path)

    return True


# --------------------------------------------------
def normalize(s: Optional[str]) -> str:
    """Remove funky bits from strings"""

    return '' if s is None else re.sub('[^a-z0-9_]', '',
                                       re.sub(r'[\s-]+', '_', s.lower()))


# --------------------------------------------------
def test_normalize():
    """Test normalize"""

    assert normalize(None) == ''
    assert normalize('FOO') == 'foo'
    assert normalize('FOO  BAR') == 'foo_bar'
    assert normalize('foo-b*!a,r') == 'foo_bar'


# --------------------------------------------------
def cell_norm(val: Any) -> str:
    """Normalize None/etc"""

    return '' if val is None else str(val)


# --------------------------------------------------
def test_cell_norm():
    """Test cell_norm"""

    assert cell_norm(None) == ''
    assert cell_norm(400) == '400'
    assert cell_norm(400.3) == '400.3'
    assert cell_norm('foo') == 'foo'


# --------------------------------------------------
if __name__ == '__main__':
    main()
