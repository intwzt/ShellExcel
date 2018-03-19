# coding=utf-8
from Style import Style


class Cube:
    def __init__(self, bg_color=None, value=None, formula=None, style=None):
        if style is None:
            self.style = Style(bg_color)
            self.formula = formula
            self.value = value
        else:
            self.style = style
            self.formula = formula
            self.value = value

    def set_value(self, value):
        self.value = value

    def set_formula(self, formula):
        self.formula = formula
