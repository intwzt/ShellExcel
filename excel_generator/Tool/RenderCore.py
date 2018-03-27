# coding=utf-8
from openpyxl.styles import numbers

from Tools import coordinate_transfer
from excel_generator.Common import border_pattern, fill_pattern, font_pattern, alignment_pattern


class RenderCore:
    def __init__(self, ws):
        self.ws = ws

    def _set_formula(self, x, y, formula):
        self.ws[coordinate_transfer(x, y)] = formula

    @staticmethod
    def _style_factory(c, cube):
        border = None if cube.style.border is None else border_pattern[cube.style.border]
        fill = None if cube.style.fill is None else fill_pattern[cube.style.fill]
        font = None if cube.style.font is None else font_pattern[cube.style.font]
        al = None if cube.style.al is None else alignment_pattern[cube.style.al]
        c.border = border
        c.fill = fill
        c.font = font
        c.alignment = al

    @staticmethod
    def _builtin_style(c, cube):
        if cube.number_format is not None:
            c.number_format = numbers.builtin_format_code(cube.number_format)

    def write_cube_to_book(self, x, y, cube):
        c = self.ws.cell(x, y, cube.value)
        if cube.formula is not None:
            self._set_formula(x, y, cube.formula)
        # write style
        # get cell range
        self._style_factory(c, cube)
        self._builtin_style(c, cube)


