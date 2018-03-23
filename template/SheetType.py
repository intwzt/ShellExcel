# coding=utf-8
from HeaderTemplate import header_index
from TitleTemplate import title_index
from onepage.target.template.SheetPageTemplate import page_index

RSM = 'RSM'

sheet_type = {RSM: {'header': header_index[RSM], 'title': title_index[RSM],
                    'page': page_index[RSM], 'main_page': 1}}
