# coding=utf-8
from HeaderTemplate import header_index
from TitleTemplate import title_index
from template.SheetPageTemplate import page_index

RSM = 'RSM'
SGM = 'SGM'

sheet_type = {RSM: {'header': header_index[RSM], 'title': title_index[RSM], 'page': page_index[RSM]},
              SGM: {'header': header_index[SGM], 'title': title_index[SGM], 'page': page_index[SGM]}}
