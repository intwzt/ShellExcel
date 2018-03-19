from openpyxl import Workbook

from Common import double_thin_top_border
from Tools import style_range

wb = Workbook()
ws = wb.active
style_range(ws, "B2:F2", double_thin_top_border)
wb.save("styled.xlsx")
