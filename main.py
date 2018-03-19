from openpyxl import Workbook

from Common import double_thin_top_border, side_pattern, side_style
from Tools import style_range, set_border

wb = Workbook()
ws = wb.active
style_range(ws, "B2:F4", double_thin_top_border)
set_border(ws, "H2:M4", side_pattern[side_style[3]])
wb.save("styled.xlsx")
