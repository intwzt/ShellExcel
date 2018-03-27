# coding=utf-8

from excel_generator.Common import side_style, font_style, alignment


class Style:
    def __init__(self, bg_color=None, border=side_style[1], font=font_style[1], al=alignment[3]):
        self.fill = bg_color
        self.border = border
        self.font = font
        self.al = al
