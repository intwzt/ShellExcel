# coding=utf-8


from excel_generator.Style import Style
# if you need add style for single gird, add it as need_merge cell
from excel_generator.Common import bg_color, alignment, side_style, font_style, number_format

# RSM and SD

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

common1_header_product = ['WT', 'Ultra', 'HX8', 'HX7', 'HX6', 'HX5', 'HX3', 'HX2', 'Other',
                          'R6', 'R5', 'R4 Plus', 'R4', 'R3', 'R2', 'Gadus', 'Spirax', 'Other']

common1_need_merge = [
    {'coordinate': [2, 0, 10, 0], 'style': Style(bg_color[4], al=alignment[5])},
    {'coordinate': [11, 0, 19, 0], 'style': Style(bg_color[4], al=alignment[5])},
    {'coordinate': [20, 0, 20, 1],
     'style': Style(bg_color[4], border=side_style[3], font=font_style[2], al=alignment[5])},

    {'coordinate': [1, 1, 1, 1], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [2, 1, 2, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [3, 1, 3, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [4, 1, 4, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [5, 1, 5, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [6, 1, 6, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [7, 1, 7, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [8, 1, 8, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [9, 1, 9, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [10, 1, 10, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [11, 1, 11, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [12, 1, 12, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [13, 1, 13, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [14, 1, 14, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [15, 1, 15, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [16, 1, 16, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [17, 1, 17, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [18, 1, 18, 1], 'style': Style(bg_color[4], al=alignment[1])},
    {'coordinate': [19, 1, 19, 1], 'style': Style(bg_color[4], al=alignment[1])}
]

common_header1 = {'matrix': common1_matrix, 'merge': common1_need_merge, 'product': common1_header_product,
                  'row': 21, 'col': 2, 'owner': 1, 'a_column': 7, 'owner_width': 3, 'follower_width': 6}

# SGM

common_header_sgm_1_0 = [
    ['SD', 'RSM', 'Ref Target KL', 'RSM&SD Submitted Target KL', 'Target Volume KL',
     'Target C3 $', 'Target Proceed $']
]

common_header_sgm_1_0_merge = [
    {'coordinate': [0, 0, 0, 0], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 1, 0, 1], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])}
]

common_header_sgm_1_1 = [
    ['SD', 'RSM', 'Province', 'City', 'LE KL', 'Market size KL(this year)', 'Market Share %',
     'Market size KL(last year)', 'Market Growth %', 'Platform', 'Market Share Score',
     'Market Growth Score', 'Platform Score', 'Market Share Score(0.75)',
     'Market Growth Score(0.15)',
     'Platform Score(0.1)', 'Total Score', 'Increase %', 'Ref Target KL', 'Target KL']
]

common_header_sgm_1_1_merge = [
    {'coordinate': [0, 0, 0, 0], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 1, 0, 1], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 2, 0, 2], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 3, 0, 3], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])}
]

common_header_sgm_1_1_formula = [0, 0, 0, 0, 0, 0, 1, 0, 1, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, -1]
common_header_sgm_1_1_total = [0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
common_header_sgm_1_1_number = [None, None, None, None, None, None, number_format['percent'], None,
                                number_format['percent'], None,
                                None, None, None, None, None, None, None, number_format['percent'], None, None]

common_header_sgm_2_0 = [
    ['SD', 'RSM', 'Province', 'City', 'Ref Volume KL', 'Ref C3 $', 'Ref Proceed $',
     'Target Volume KL', 'Target C3 $', 'Target Proceed $']
]

common_header_sgm_2_1 = [
    ['SD', 'RSM', 'Province', 'City', 'UC3 $', 'UC3 $', 'UC3 $', 'UC3 $', 'UC3 $',
     'UC3 $', 'UC3 $', 'UC3 $', 'UC3 $', 'UNP $', 'UNP $', 'UNP $', 'UNP $', 'UNP $',
     'UNP $', 'UNP $', 'UNP $', 'UNP $', 'Portfolio %', 'Portfolio %', 'Portfolio %',
     'Portfolio %', 'Portfolio %', 'Portfolio %', 'Portfolio %', 'Portfolio %', 'Portfolio %'],
    ['SD', 'RSM', 'Province', 'City', 'WT', 'Ultra', 'HX8', 'HX7', 'HX6', 'HX5', 'HX3', 'HX2', 'Other',
     'WT', 'Ultra', 'HX8', 'HX7', 'HX6', 'HX5', 'HX3', 'HX2', 'Other',
     'WT', 'Ultra', 'HX8', 'HX7', 'HX6', 'HX5', 'HX3', 'HX2', 'Other']
]
common_header_sgm_2_1_merge = [
    {'coordinate': [0, 0, 1, 0], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 1, 1, 1], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 2, 1, 2], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 3, 1, 3], 'style': Style(bg_color[4], font=font_style[2], al=alignment[1])},
    {'coordinate': [0, 4, 0, 12], 'style': Style(bg_color[4], font=font_style[2], al=alignment[2])},
    {'coordinate': [0, 13, 1, 21], 'style': Style(bg_color[4], font=font_style[2], al=alignment[2])},
    {'coordinate': [0, 22, 1, 30], 'style': Style(bg_color[4], font=font_style[2], al=alignment[2])}
]

common_header_sgm_1 = {
    0: {'data': common_header_sgm_1_0, 'scale': [1, 7], 'merge': common_header_sgm_1_0_merge, 'formula': None,
        'number_format': None},
    1: {'data': common_header_sgm_1_1, 'scale': [1, 20], 'merge': common_header_sgm_1_1_merge,
        'formula': common_header_sgm_1_1_formula, 'number_format': common_header_sgm_1_1_number,
        'total': common_header_sgm_1_1_total}
}

common_header_sgm_2 = {
    0: {'data': common_header_sgm_2_0, 'scale': [1, 10], 'merge': None, 'formula': None, 'number_format': None},
    1: {'data': common_header_sgm_2_1, 'scale': [2, 31], 'merge': common_header_sgm_2_1_merge, 'formula': [],
        'number_format': None}
}

header_index = {
    'RSM': common_header1,
    'SGM': {1: common_header_sgm_1,
            2: common_header_sgm_2,
            3: common_header_sgm_1,
            4: common_header_sgm_2}
}
