# coding=utf-8
from Cube import Cube
from Common import bg_color, column_type, role


class Column:
    def __init__(self, num, col_type, col_role):
        self.count = num
        self.container = []
        self.col_type = col_type
        self.col_role = col_role
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

    def set_column_value(self, source):
        source_len = len(source)
        if source_len != self.count:
            print 'target column do not match source'
            return
        for i in range(source_len):
            self.container[i].set_value(source[i])
