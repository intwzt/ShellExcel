# coding=utf-8
from Style import Style


class Cube:
    def __init__(self, bg_color=None, value=None, formula=None, style=None, number_format=None):
        if style is None:
            self.style = Style(bg_color)
        else:
            self.style = style
        self.formula = formula
        self.value = value
        self.number_format = number_format

    def set_style(self, style):
        self.style = style

    def set_value(self, value):
        self.value = value

    def set_formula(self, formula):
        self.formula = formula

    def set_number_format(self, number_format):
        self.number_format = number_format
