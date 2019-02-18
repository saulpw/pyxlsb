import os
import sys
import xml.etree.ElementTree as ElementTree
from . import recordtypes as rt
from .formula import Formula
from .recordreader import RecordReader
from itertools import zip_longest
from visidata import Sheet, asyncthread, vd, Progress, Column, options

if sys.version_info > (3,):
    xrange = range


class Cell(object):
    __slots__ = ('r', 'c', 'value', '_formula')

    def __init__(self, r, c, v, f=None):
        self.r = r
        self.c = c
        self.value = v
        self._formula = f

    def __repr__(self):
        return 'Cell(r={}, c={}, v={}, f={})'.format(self.r, self.c, self.value, self._formula)

    @property
    def v(self):
        return self.value

    @property
    def formula(self):
        if self._formula is None:
            return None
        elif not isinstance(self._formula, Formula):
            self._formula = Formula.parse(self._formula)
        return self._formula


class Worksheet(Sheet):
    def __init__(self, workbook, name, fp, rels_fp=None):
        super().__init__(name, workbook=workbook)
        self._fp = fp
        self._rels_fp = rels_fp

    @asyncthread
    def reload(self):
        self.columns = list()
        self.rows = list()

        iterrows = self.iterload()
        hdrs = [next(iterrows) for i in range(options.header)]
        if hdrs:
            for i in range(len(hdrs[0])):
                self.addColumn(Column('', getter=lambda c,r,i=i: r[i].v))
            self.setColNames(hdrs)

        for row in iterrows:
            for i in range(len(self.columns), len(row)):  # no-op if same
                self.addColumn(Column('', getter=lambda c,r,i=i: r[i].v))

            self.addRow(row)

    def iterload(self):
        self.dimension = None
        self.hyperlinks = dict()

        row = None
        row_num = -1
        sparse = True  # options.keep_empty_rows

        for rectype, rec in RecordReader(self._fp):
            if rectype == rt.WS_DIM:
                self.dimension = rec
            elif rectype == rt.COL_INFO:
                self.addColumn(Column('', width=rt.width, style=rt.style, expr=rt.customWidth))
            elif rectype == rt.BEGIN_SHEET_DATA:
                self.parse_rels(self._rels_fp)
            elif rectype == rt.H_LINK:
                for r in xrange(rec.h):
                    for c in xrange(rec.w):
                        self.hyperlinks[rec.r + r, rec.c + c] = rec.rId

            elif rectype == rt.ROW_HDR and rec.r != row_num:
                if row is not None:
                    yield row
                while not sparse and row_num < rec.r - 1:
                    row_num += 1
                    yield [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
                row_num = rec.r
                row = [Cell(row_num, i, None, None) for i in xrange(self.dimension.c + self.dimension.w)]
            elif rectype == rt.CELL_ISST:
                row[rec.c] = Cell(row_num, rec.c, self.workbook.get_shared_string(rec.v), rec.f)
            elif rectype >= rt.CELL_BLANK and rectype <= rt.FMLA_ERROR:
                row[rec.c] = Cell(row_num, rec.c, rec.v, rec.f)
            elif rectype == rt.END_SHEET_DATA:
                if row is not None:
                    yield row
                break

    @asyncthread
    def parse_rels(self, rels_fp):
        self.rels = dict()
        doc = ElementTree.parse(rels_fp)
        for el in doc.getroot():
            self.rels[el.attrib['Id']] = el.attrib['Target']

    def close(self):
        self._fp.close()
        if self._rels_fp is not None:
            self._rels_fp.close()
