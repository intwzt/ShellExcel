# coding=utf-8
from Style import Style
from Common import bg_color, side_style

common1_matrix = [
    [None, None],
    [None, 'Product'],
    ['PCMO', 'WT'],
    ['PCMO', 'Ultra'],
    ['PCMO', 'HX8'],
    ['PCMO', 'HX7'],
    ['PCMO', 'HX6'],
    ['PCMO', 'HX5'],
    ['PCMO', 'HX3'],
    ['PCMO', 'HX2'],
    ['PCMO', 'Other'],
    ['CRTO', 'R6'],
    ['CRTO', 'R5'],
    ['CRTO', 'R4 Plus'],
    ['CRTO', 'R4'],
    ['CRTO', 'R3'],
    ['CRTO', 'R2'],
    ['CRTO', 'Gadus'],
    ['CRTO', 'Spirax'],
    ['CRTO', 'Other'],
    ['Sum Total', None]
]

common1_need_merge = [
    {'coordinate': [2, 0, 10, 0], 'style': Style(bg_color[4])},
    {'coordinate': [11, 0, 19, 0], 'style': Style(bg_color[4])},
    {'coordinate': [20, 0, 20, 1], 'style': Style(bg_color[4], border=side_style[3])}
]

common_header1 = {'matrix': common1_matrix, 'merge': common1_need_merge}
