import os
from . import recordtypes as rt
from .recordreader import RecordReader


class Styles(object):
    def __init__(self, fp):
        self._fp = fp
        self._parse()

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def _parse(self):
        self._colors = list()
        self._dxfs = list()
        self._table_styles = list()
        self._fills = list()
        self._fonts = list()
        self._borders = list()
        self._cell_xfs = list()
        self._cell_styles = list()
        self._cell_style_xfs = list()

        for rectype, rec in RecordReader(self._fp):
            # TODO
            if rectype == rt.END_STYLE_SHEET:
                break

    def get_style(self, idx):
        return None

    def close(self):
        self._fp.close()
