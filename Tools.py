# coding=utf-8
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment


def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
    top = Border(top=border.top)
    left = Border(left=border.left)
    right = Border(right=border.right)
    bottom = Border(bottom=border.bottom)

    first_cell = ws[cell_range.split(":")[0]]
    if alignment:
        ws.merge_cells(cell_range)
        first_cell.alignment = alignment

    rows = ws[cell_range]
    if font:
        first_cell.font = font

    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom

    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
        if fill:
            for c in row:
                c.fill = fill


def decimal2letter(x):
    result = ''
    while int(x / 26):
        result += chr(ord('@') + x % 26)
        x /= 26
    result += chr(ord('@') + x % 26)
    return result[::-1]


def coordinate_transfer(x, y):
    return decimal2letter(y) + str(x)
