from openpyxl import Workbook

from Common import double_thin_top_border
from Tools import style_range, coordinate_transfer

wb = Workbook()
ws = wb.active

coordinate = coordinate_transfer(3, 27) + ':' + coordinate_transfer(5, 29)
print coordinate
style_range(ws, coordinate, double_thin_top_border)
wb.save("styled.xlsx")
