# coding=utf-8

from excel_generator.Common import attach_column_type, target_mapper, ref_mapper
from template.SheetType import sheet_type


class PageParser:
    def __init__(self, ws, table_type, x=3, y=2):
        self.table_type = table_type
        self.origin = [x, y]
        self.main_origin = [x, y]
        self.current_ws = ws
        self.cal_main_table_coordinate()

    def cal_main_table_coordinate(self):
        count = 1
        flag = 0
        while flag < 2:
            flag += (1 if self.current_ws.cell(self.origin[0] + count, self.origin[1]).value is None else 0)
            count += 1
        self.main_origin = [self.origin[0] + count, self.origin[1]]

    def parse_attached_table(self):
        table_x = self.origin[0] + 1 + 1
        table_y = self.origin[1]
        data = []
        flag = 1
        count = 0
        while flag:
            row_data = {}
            for i in range(sheet_type[self.table_type]['header']['a_column']):
                if i == 0:
                    row_data['Name'] = self.current_ws.cell(table_x + count, table_y + i).value
                else:
                    row_data[attach_column_type[i - 1]] = self.current_ws.cell(table_x + count, table_y + i).value
            if row_data['Name'] != 'Sum Total':
                data.append(row_data)
            count += 1
            flag = 0 if self.current_ws.cell(table_x + count, table_y).value is None else 1
        return data

    def parse_main_table(self):
        owner_data = self._parse_owner()
        follower_data = self._parse_follower()
        follower_data.append(owner_data)
        return follower_data

    def _parse_owner(self):
        table_x = self.main_origin[0] + 1
        table_y = self.main_origin[1]
        width = sheet_type[self.table_type]['header']['owner_width']
        num_of_header_col = sheet_type[self.table_type]['header']['col']
        product_list = sheet_type[self.table_type]['header']['product']
        num_of_product = len(product_list)
        data = {'Name': self.current_ws.cell(table_x, table_y + num_of_header_col).value}
        for j in range(width):
            tmp = {}
            for i in range(num_of_product):
                tmp[product_list[i]] = self.current_ws.cell(table_x + 2 + i, table_y + num_of_header_col + j).value
            data[target_mapper[j]] = tmp
        return data

    def get_user_header(self):
        user_header = []
        for h in ref_mapper:
            user_header.append(h)
        for h in target_mapper:
            user_header.append(h)
        return user_header

    def _parse_follower(self):
        table_x = self.main_origin[0] + 1
        table_y = self.main_origin[1] + 2 + sheet_type[self.table_type]['header']['owner_width']
        width = sheet_type[self.table_type]['header']['follower_width']
        product_list = sheet_type[self.table_type]['header']['product']
        num_of_product = len(product_list)
        flag = 1
        count = 0
        data = []
        user_header = self.get_user_header()
        while flag:
            current_user = {'Name': self.current_ws.cell(table_x, table_y + count).value}
            for j in range(width):
                tmp = {}
                for i in range(num_of_product):
                    tmp[product_list[i]] = self.current_ws.cell(table_x + 2 + i, table_y + j + count).value
                current_user[user_header[j]] = tmp
            count += width
            flag = 0 if self.current_ws.cell(table_x, table_y + count).value is None else 1
            data.append(current_user)
        return data

