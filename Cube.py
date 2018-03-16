# coding=utf-8
from Style import Style


class Cube:
    def __init__(self, bg_color, value=None, formula=None):
        self.style = Style(bg_color)
        self.formula = formula
        self.value = value

    def set_value(self, value):
        self.value = value

    def set_formula(self, formula):
        self.formula = formula
