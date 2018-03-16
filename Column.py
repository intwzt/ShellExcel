# coding=utf-8
from Cube import Cube
from Common import bg_color, column_type, role, column_last_row, formula_type


class Column:
    def __init__(self, num, col_type, col_role, last_row=column_last_row[4]):
        self.count = num
        self.container = []
        self.col_type = col_type
        self.col_role = col_role
        self.last_row = last_row
        for i in range(num):
            self.container.append(self._set_style())

    def _set_style(self):
        if self.col_type == column_type[1] and self.col_role == role[1]:
            return Cube(bg_color[3])
        elif self.col_type == column_type[1] and self.col_role == role[2]:
            return Cube(bg_color[4])
        elif self.col_type == column_type[2]:
            return Cube(bg_color[2])
        else:
            return Cube(bg_color[4])

    def _get_last_row(self):
        if self.last_row == column_last_row[1]:
            return Cube(bg_color[4])
        elif self.last_row == column_last_row[2]:
            return Cube(bg_color[4], value='N/A')
        elif self.last_row == column_last_row[3]:
            return Cube(bg_color[1], formula=formula_type[1])
        else:
            pass

    def set_column_value(self, source):
        source_len = len(source)
        if source_len != self.count:
            print 'target column do not match source'
            return
        for i in range(source_len):
            self.container[i].set_value(source[i])
        # add last row
        if self.last_row != column_last_row[4]:
            self.container.append(self._get_last_row())
