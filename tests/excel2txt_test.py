#!/usr/bin/env python3
"""tests for excel2txt.py"""

import csv
import os
import random
import re
import string
from excel2txt import normalize, cell_norm
from subprocess import getstatusoutput
from shutil import rmtree

PRG = './excel2txt.py'
INPUT1 = './tests/test1.xlsx'
INPUT2 = './tests/Test 2.xlsx'


# --------------------------------------------------
def test_normalize():
    """Test normalize"""

    assert normalize(None) == ''
    assert normalize('') == ''
    assert normalize('FOO') == 'foo'
    assert normalize('FOO  BAR') == 'foo_bar'
    assert normalize('foo-b*!a,r') == 'foobar'
    assert normalize('Foo Bar') == 'foo_bar'
    assert normalize('Foo / Bar') == 'foo_bar'
    assert normalize('Foo (Bar)') == 'foo_bar'
    assert normalize('FooBarBAZ') == 'foo_bar_baz'


# --------------------------------------------------
def test_cell_norm():
    """Test cell_norm"""

    assert cell_norm(None) == ''
    assert cell_norm(400) == '400'
    assert cell_norm(400.3) == '400.3'
    assert cell_norm('foo') == 'foo'


# --------------------------------------------------
def test_exists():
    """exists"""

    assert os.path.isfile(PRG)


# --------------------------------------------------
def test_usage():
    """ Prints sage """

    for flag in ['-h', '--help']:
        rv, out = getstatusoutput(f'{PRG} {flag}')
        assert rv == 0
        assert out.lower().startswith('usage')


# --------------------------------------------------
def test_bad_file():
    """ Dies on bad file """

    bad = random_string()
    rv, out = getstatusoutput(f'{PRG} {bad}')
    assert rv != 0
    assert out.lower().startswith('usage')
    assert re.search(f"No such file or directory: '{bad}'", out)


# --------------------------------------------------
def test_input1():
    """ Input1 file """

    out_file = './test1__sheet1.txt'

    if os.path.isfile(out_file):
        os.remove(out_file)

    try:
        rv, out = getstatusoutput(f'{PRG} "{INPUT1}"')
        assert rv == 0
        lines = out.splitlines()
        assert lines[-1].strip() == f'Done, see output in "{os.getcwd()}".'
        assert os.path.isfile(out_file)

        reader = csv.DictReader(open(out_file), delimiter='\t')
        assert reader.fieldnames == ['name', 'rank', 'serial_number']
        data = list(reader)
        assert len(data) == 2
        assert data[0]['name'] == 'Ed'
        assert data[1]['rank'] == 'Major'

    finally:
        if os.path.isfile(out_file):
            os.remove(out_file)


# --------------------------------------------------
def test_input2():
    """ Input2 file """

    out_dir = f'./{random_string()}'
    if os.path.isdir(out_dir):
        rmtree(out_dir)

    try:
        rv, out = getstatusoutput(f'{PRG} -d "," -o {out_dir} -n "{INPUT2}"')
        assert rv == 0
        lines = out.splitlines()
        msg = f'Done, see output in "{os.path.abspath(out_dir)}".'
        assert lines[-1].strip() == msg
        out_file = f'./{out_dir}/test_2__sheet1.csv'
        assert os.path.isfile(out_file)

        reader = csv.DictReader(open(out_file), delimiter=',')
        assert reader.fieldnames == ['ice_cream_flavor', 'peoples_rank']
        data = list(reader)
        assert len(data) == 3
        assert data[0]['ice_cream_flavor'] == 'chocolate'
        assert data[-1]['peoples_rank'] == '3'

    finally:
        if os.path.isdir(out_dir):
            rmtree(out_dir)


# --------------------------------------------------
def test_multiple_inputs():
    """ More than one input file """

    out_dir = f'./{random_string()}'
    if os.path.isdir(out_dir):
        rmtree(out_dir)

    try:
        delimiter = random.choice(';|:')
        cmd = (f'{PRG} --delimiter "{delimiter}" --outdir {out_dir} '
               f'--normalize --mkdirs "{INPUT1}" "{INPUT2}"')
        rv, out = getstatusoutput(cmd)
        assert rv == 0

        lines = out.splitlines()
        msg = f'Done, see output in "{os.path.abspath(out_dir)}".'
        assert lines[-1].strip() == msg
        out_files = [
            os.path.join(out_dir, 'test1', 'test1__sheet1.txt'),
            os.path.join(out_dir, 'test_2', 'test_2__sheet1.txt')
        ]

        assert all(map(os.path.isfile, out_files))

        reader1 = csv.DictReader(open(out_files[0]), delimiter=delimiter)
        assert reader1.fieldnames == ['name', 'rank', 'serial_number']
        data1 = list(reader1)
        assert len(data1) == 2
        assert data1[0]['name'] == 'Ed'
        assert data1[1]['rank'] == 'Major'

        reader2 = csv.DictReader(open(out_files[1]), delimiter=delimiter)
        assert reader2.fieldnames == ['ice_cream_flavor', 'peoples_rank']
        data2 = list(reader2)
        assert len(data2) == 3
        assert data2[0]['ice_cream_flavor'] == 'chocolate'
        assert data2[-1]['peoples_rank'] == '3'

    finally:
        if os.path.isdir(out_dir):
            rmtree(out_dir)


# --------------------------------------------------
def random_string():
    """generate a random string"""

    k = random.randint(5, 10)
    return ''.join(random.choices(string.ascii_letters + string.digits, k=k))
