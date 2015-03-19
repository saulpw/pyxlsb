import biff12
from reader import BIFF12Reader

class Worksheet(object):
  def __init__(self, fp, stringtable=None):
    super(Worksheet, self).__init__()
    self._reader = BIFF12Reader(fp=fp)
    self._stringtable = stringtable
    self._parse()

  def __enter__(self):
    return self

  def __exit__(self, type, value, traceback):
    self.close()

  def _parse(self):
    for item in self._reader:
      if item[0] == biff12.SHEETDATA:
        break

  def rows(self):
    row_num = 0
    row = []
    for item in self._reader:
      if item[0] == biff12.ROW and item[1].r != row_num:
        while row_num + 1 < item[1].r:
          yield []
          row_num += 1
        yield row
        row_num = item[1].r
        row = []
      elif item[0] >= biff12.BLANK and item[0] <= biff12.FORMULA_BOOLERR:
        while len(row) <= item[1].c:
          row.append(None)
        if item[0] == biff12.STRING and not self._stringtable is None:
          row[item[1].c] = self._stringtable[item[1].v]
        else:
          row[item[1].c] = item[1].v
      elif item[0] == biff12.SHEETDATA_END:
        break
    if row:
      yield row

  def close(self):
    self._reader.close()