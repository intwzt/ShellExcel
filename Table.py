# coding=utf-8
from openpyxl import Workbook

from Common import color_pattern, bg_color, target_mapper, ref_mapper
from Cube import Cube


class Table:
    def __init__(self, ox, oy):
        self.owner = None
        self.follower = []
        self.header = None
        self.origin = [ox, oy]
        self.table_body = [ox, oy]

        self.wb = Workbook()
        self.ws = self.wb.active

    def assign_header(self, header):
        self.header = header
        self.table_body = [self.table_body[0], self.table_body[1] + header.y]

    def _check_cube_style(self, x, y):
        for item in self.header.merge:
            if item['coordinate'][0] == x and item['coordinate'][1] == y:
                return item['style']
        return None

    def _render_table_header(self):
        # render cell with style
        for i in range(self.header.x):
            for j in range(self.header.y):
                current_value = self.header.matrix[i][j]
                checked_style = self._check_cube_style(i, j)
                current_cube = Cube(bg_color[4], value=current_value) if checked_style is None else Cube(
                    checked_style.bg_color, value=current_value)
                self._write_cube_to_book(i+1, j+1, current_cube)

        # merge cell
        if self.header.merge is not None:
            for m in self.header.merge:
                self.ws.merge_cells(start_row=m['coordinate'][0] + self.origin[0],
                                    start_column=m['coordinate'][1] + self.origin[1],
                                    end_row=m['coordinate'][2] + self.origin[0],
                                    end_column=m['coordinate'][3] + self.origin[1])

    def assign_person(self, owner, follower=None):
        self.owner = owner
        if follower:
            for f in follower:
                self.follower.append(f)

    def _set_background(self, c, color):
        c.fill = color_pattern[color]

    def _set_style(self, c, style):
        self._set_background(c, style.bg_color)

    def _write_cube_to_book(self, x, y, cube):
        c = self.ws.cell(x, y, cube.value)
        self._set_style(c, cube.style)

    def print_info(self):
        self.owner.print_info()

    def save_workbook(self):
        self.wb.save('result.xlsx')

    def render(self):
        # render header
        self._render_table_header()
        # render body
        index_x = self.table_body[0]
        index_y = self.table_body[1]
        counter = 0
        num_of_column = 0 if self.owner is None else len(self.owner.target[target_mapper[0]].container)

        self._render_owner(counter, index_x, index_y, num_of_column)
        # modify index_x and reset counter
        index_y += 3
        counter = 0

        num_of_column = 0 if len(self.follower) is None else len(self.follower[0].target[target_mapper[0]].container)
        self._render_follower(counter, index_x, index_y, num_of_column)

    def _render_follower(self, counter, index_x, index_y, num_of_column):
        # write follower if exist
        if len(self.follower) > 0:
            # for all follower
            for f in range(len(self.follower)):
                inner_counter = 0
                # write person name
                for i in range(6):
                    self._write_cube_to_book(index_x + counter + inner_counter, index_y + i + f * 6,
                                             Cube(bg_color[4], self.follower[f].name))
                # merge name cells
                self.ws.merge_cells(start_row=index_x + counter + inner_counter,
                                    start_column=index_y + f * 6,
                                    end_row=index_x + counter + inner_counter,
                                    end_column=index_y + f * 6 + 6 - 1)

                inner_counter += 1
                # write column name
                for i in range(3):
                    self._write_cube_to_book(index_x + counter + inner_counter, index_y + i + f * 6,
                                             Cube(bg_color[4], ref_mapper[i]))
                for i in range(3, 6):
                    self._write_cube_to_book(index_x + counter + inner_counter, index_y + i + f * 6,
                                             Cube(bg_color[4], target_mapper[i - 3]))

                inner_counter += 1
                # write column data
                for i in range(num_of_column):
                    for j in range(3):
                        self._write_cube_to_book(index_x + counter + i + inner_counter, index_y + j + f * 6,
                                                 self.follower[f].ref[ref_mapper[j]].container[i])
                for i in range(num_of_column):
                    for j in range(3):
                        self._write_cube_to_book(index_x + counter + i + inner_counter, index_y + j + 3 + f * 6,
                                                 self.follower[f].target[target_mapper[j]].container[i])
            counter += 1

    def _render_owner(self, counter, index_x, index_y, num_of_column):
        # write owner if exist
        if self.owner:
            # write person name
            for i in range(3):
                self._write_cube_to_book(index_x + counter, index_y + i, Cube(bg_color[4], self.owner.name))
            # merge name cells
            self.ws.merge_cells(start_row=index_x + counter,
                                start_column=index_y,
                                end_row=index_x + counter,
                                end_column=index_y + 3 - 1)
            counter += 1
            # write column name
            for i in range(3):
                self._write_cube_to_book(index_x + counter, index_y + i, Cube(bg_color[4], target_mapper[i]))
            counter += 1
            # write column data
            for i in range(num_of_column):
                for j in range(3):
                    self._write_cube_to_book(index_x + counter + i, index_y + j,
                                             self.owner.target[target_mapper[j]].container[i])
