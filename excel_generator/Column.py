# coding=utf-8
from Cube import Cube
from Style import Style
from excel_generator.Common import bg_color, column_type, role, column_last_row, formula_type, side_style


class Column:
    def __init__(self, num, col_type, col_role, number_format=None, last_row=column_last_row[4], formula_type=None):
        self.count = num
        self.container = []
        self.col_type = col_type
        self.col_role = col_role
        self.last_row = last_row
        self.number_format = number_format
        self.formula_type = formula_type
        for i in range(num):
            self.container.append(self._set_style())
        # add last row
        if self.last_row != column_last_row[4]:
            self.container.append(self._get_last_row())
        self._check_number_format()

    def _set_style(self):
        if self.col_type == column_type[1] and self.col_role == role[1] and self.formula_type is None:
            return Cube(bg_color[3])
        elif self.col_type == column_type[1] and self.col_role == role[2] and self.formula_type is None:
            return Cube(bg_color[4])
        elif self.col_type == column_type[2] and self.formula_type is None:
            return Cube(bg_color[2])
        elif self.formula_type is not None:
            return Cube(bg_color[1])
        else:
            return Cube(bg_color[4])

    def _check_number_format(self):
        if self.number_format is not None:
            for item in self.container:
                item.set_number_format(self.number_format)

    def _get_last_row(self):
        if self.last_row == column_last_row[1]:
            return Cube(bg_color[4])
        elif self.last_row == column_last_row[2]:
            column_sum_style = Style(bg_color=bg_color[4], border=side_style[3])
            return Cube(bg_color[4], value='N/A', style=column_sum_style)
        elif self.last_row == column_last_row[3]:
            column_sum_style = Style(bg_color=bg_color[1], border=side_style[3])
            return Cube(bg_color[1], formula=formula_type[1], style=column_sum_style)
        else:
            pass

    def set_column_value(self, source):
        source_len = len(source)
        if source_len != self.count:
            print ('target column do not match source')
            return
        for i in range(source_len):
            self.container[i].set_value(source[i])
