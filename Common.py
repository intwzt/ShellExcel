# coding=utf-8
from openpyxl.styles import PatternFill, Alignment, Side

role = {1: 'owner', 2: 'follower'}

bg_color = {1: 'gray', 2: 'yellow', 3: 'blue', 4: 'white'}

color_pattern = {bg_color[1]: PatternFill('solid', fgColor="e6e6e6"),
                 bg_color[2]: PatternFill('solid', fgColor='fff1ce'),
                 bg_color[3]: PatternFill('solid', fgColor='4674c1'),
                 bg_color[4]: PatternFill('solid', fgColor='ffffff')}

font_color = {1: 'red'}

formula = {1: 'None', 2: 'have'}

target_mapper = ['Target UC3 $', 'Target UNP $', 'Target Portfolio %']

ref_mapper = ['Ref UC3 $', 'Ref UNP $', 'Ref Portfolio %']

column_type = {1: 'target', 2: 'ref'}


side_style = {1: 'thin'}
side_pattern = {side_style[1]: Side(border_style="thin", color="000000")}

alignment = {1: 'left', 2: 'center', 3: 'right'}
alignment_pattern = {'center': Alignment(horizontal="center", vertical="center")}

