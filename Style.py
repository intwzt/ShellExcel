# coding=utf-8

from Common import side_style, font_style, alignment


class Style:
    def __init__(self, bg_color, border=side_style[1], font=font_style[1], al=alignment[2]):
        self.fill = bg_color
        self.border = border
        self.font = font
        self.al = al
