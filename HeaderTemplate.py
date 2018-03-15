# coding=utf-8
from Style import Style
from Common import bg_color

common1_matrix = [
    [None, None],
    [None, 'Product'],
    ['PCMO', 'WT'],
    ['PCMO', 'Ultra'],
    ['PCMO', 'HX8'],
    ['PCMO', 'HX7'],
    ['PCMO', 'HX6'],
    ['PCMO', 'HX5'],
    ['PCMO', 'HX4'],
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
    ['Sum Total']
]

common1_need_merge = [
    {'coordinate': [2, 1, 10, 1], 'style': Style(bg_color[4])},
    {'coordinate': [11, 1, 19, 1], 'style': Style(bg_color[4])},
    {'coordinate': [20, 1, 20, 2], 'style': Style(bg_color[4])}
]

common_header1 = {'matrix': common1_matrix, 'merge': common1_need_merge}
