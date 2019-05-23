from xlrd import open_workbook
from xlutils.copy import copy

rb = open_workbook("devis_t.xls")
wb = copy(rb)

s = wb.get_sheet(0)
s.write(11,6,"test")
wb.save("devis_t.xls")