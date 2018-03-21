# coding=utf-8

from Column import Column
from Common import bg_color, formula_type, attach_column_type, column_type, role, target_mapper, font_style, alignment, \
    side_style
from Common import column_last_row, thick_border
from Cube import Cube
from Style import Style
from Tools import coordinate_transfer, style_range


class AttachedTable:
    def __init__(self, main_table, ox, oy, title=None, data=None):
        # add title row
        ox += 1
        self.main_table = main_table
        self.origin = [ox, oy]
        self.title = title
        self.container = {}
        self.name = []
        self.data = data
        self.ws = main_table.ws
        self.end_point = [ox, oy]

    def assign_data(self, data):
        self.data = data

    def assign_title(self, title):
        self.title = title

    def _cal_end_point(self):
        num_of_owner = 0 if self.main_table.owner is None else 1
        numb_of_follower = 0 if self.main_table.follower is None else len(self.main_table.follower)
        end_x = self.origin[0] + num_of_owner + numb_of_follower + 2 - 1
        end_y = self.origin[1] + 6
        self.end_point = [end_x, end_y]

    def _add_table_border(self):
        coordinate = coordinate_transfer(self.origin[0], self.origin[1]) + ':' + \
                     coordinate_transfer(self.end_point[0], self.end_point[1])
        style_range(self.ws, coordinate, thick_border)

    def _render_table_title(self):
        style = Style(bg_color[4], border=None, font=font_style[3], al=alignment[4])
        self.main_table.write_cube_to_book(self.origin[0] - 1, self.origin[1],
                                           Cube(bg_color[4], value=self.title, style=style))

    def _extract_name(self):
        if self.main_table.follower is not None:
            for f in self.main_table.follower:
                self.name.append(f.name)

        self.name.append('Sum Total')
        self.name.append(self.main_table.owner.name)

    def _init_column(self):
        # extract total number of people
        num = 0 if self.main_table.follower is None else len(self.main_table.follower)
        # ref column
        for i in range(2):
            col = Column(num, column_type[2], role[1], last_row=column_last_row[3])
            self.container[attach_column_type[i]] = col
        # init column
        self.container[attach_column_type[2]] = Column(num, column_type[2], role[1],
                                                       last_row=column_last_row[3],
                                                       number_format=10)

        self.container[attach_column_type[3]] = Column(num, column_type[1], role[2],
                                                       last_row=column_last_row[3])

        self.container[attach_column_type[4]] = Column(num, column_type[1], role[2],
                                                       last_row=column_last_row[3],
                                                       formula_type=formula_type[2])

        self.container[attach_column_type[5]] = Column(num, column_type[1], role[2],
                                                       last_row=column_last_row[3],
                                                       formula_type=formula_type[2])

    def _extract_column_data(self):
        self._cal_Target_C3()
        self._cal_Target_Proceed()

    def _get_ref_value(self, header_name, person_name):
        for f in self.main_table.follower:
            if f.name == person_name:
                return f.attached_column_name_mapper(header_name)
        print 'error!'
        return None

    def _set_data(self):
        # assign follower value
        num_of_follower = 0 if self.main_table.follower is None else len(self.main_table.follower)
        for i in range(3):
            for f in range(num_of_follower):
                header_name = attach_column_type[i]
                person_name = self.main_table.follower[f].name
                self.container[attach_column_type[i]].container[f].set_value(
                    self._get_ref_value(header_name, person_name))

            # set total formula
            total_coordinate = coordinate_transfer(self.origin[0] + 1, self.origin[1] + 1 + i) + ':' + \
                               coordinate_transfer(self.origin[0] + 1 + num_of_follower - 1, self.origin[1] + 1 + i)
            total_formula = "=SUM(" + total_coordinate + ")"
            self.container[attach_column_type[i]].container[-1].set_formula(total_formula)

            # assign if table has owner
            if self.main_table.owner is not None:
                self.container[attach_column_type[i]].container.append(Cube(bg_color=bg_color[4]))

        # add total and owner for Target Volume KL
        # assign total formula
        total_coordinate = coordinate_transfer(self.origin[0] + 1, self.origin[1] + 4) + ':' + \
                           coordinate_transfer(self.origin[0] + num_of_follower, self.origin[1] + 4)
        total_formula = "=SUM(" + total_coordinate + ")"
        self.container[attach_column_type[3]].container[-1].set_formula(total_formula)

        # assign if table has owner
        if self.main_table.owner is not None:
            self.container[attach_column_type[3]].container.append(Cube(bg_color=bg_color[3]))

    def _cal_Target_C3(self):
        main_table_body = self.main_table.table_body
        num_of_follower = 0 if self.main_table.follower is None else len(self.main_table.follower)
        num_of_row = len(self.main_table.follower[0].target[target_mapper[0]].container)
        for i in range(num_of_follower):
            # get base volume coordinate
            base_coordinate = coordinate_transfer(self.origin[0] + 1 + i, self.origin[1] + 4)

            first_column = coordinate_transfer(main_table_body[0] + 2, main_table_body[1] + 3 + i * 6 + 3) + \
                           ':' + coordinate_transfer(main_table_body[0] + 2 + num_of_row - 1,
                                                     main_table_body[1] + 3 + i * 6 + 3)
            second_column = coordinate_transfer(main_table_body[0] + 2, main_table_body[1] + 3 + i * 6 + 5) + \
                            ':' + coordinate_transfer(main_table_body[0] + 2 + num_of_row - 1,
                                                      main_table_body[1] + 3 + i * 6 + 5)
            formula = '=' + base_coordinate + '*' + 'SUMPRODUCT(' + first_column + ',' + second_column + ')'
            self.container[attach_column_type[4]].container[i].set_formula(formula)

        # assign total formula
        total_coordinate = coordinate_transfer(self.origin[0] + 1, self.origin[1] + 5) + ':' + \
                           coordinate_transfer(self.origin[0] + num_of_follower, self.origin[1] + 5)
        total_formula = "=SUM(" + total_coordinate + ")"
        self.container[attach_column_type[4]].container[-1].set_formula(total_formula)

        # assign if table has owner
        if self.main_table.owner is not None:
            self.container[attach_column_type[4]].container.append(Cube(bg_color=bg_color[3]))

    def _cal_Target_Proceed(self):
        main_table_body = self.main_table.table_body
        num_of_follower = 0 if self.main_table.follower is None else len(self.main_table.follower)
        num_of_row = len(self.main_table.follower[0].target[target_mapper[0]].container)
        for i in range(num_of_follower):
            # get base volume coordinate
            base_coordinate = coordinate_transfer(self.origin[0] + 1 + i, self.origin[1] + 4)

            first_column = coordinate_transfer(main_table_body[0] + 2, main_table_body[1] + 3 + i * 6 + 4) + \
                           ':' + coordinate_transfer(main_table_body[0] + 2 + num_of_row - 1,
                                                     main_table_body[1] + 3 + i * 6 + 4)
            second_column = coordinate_transfer(main_table_body[0] + 2, main_table_body[1] + 3 + i * 6 + 5) + \
                            ':' + coordinate_transfer(main_table_body[0] + 2 + num_of_row - 1,
                                                      main_table_body[1] + 3 + i * 6 + 5)
            formula = '=' + base_coordinate + '*' + 'SUMPRODUCT(' + first_column + ',' + second_column + ')'
            self.container[attach_column_type[5]].container[i].set_formula(formula)

        # assign total formula
        total_coordinate = coordinate_transfer(self.origin[0] + 1, self.origin[1] + 6) + ':' + \
                           coordinate_transfer(self.origin[0] + num_of_follower, self.origin[1] + 6)
        total_formula = "=SUM(" + total_coordinate + ")"
        self.container[attach_column_type[5]].container[-1].set_formula(total_formula)

        # assign if table has owner
        if self.main_table.owner is not None:
            self.container[attach_column_type[5]].container.append(Cube(bg_color=bg_color[3]))

    def _extract_data(self):
        self._init_column()
        self._set_data()
        self._extract_name()
        self._extract_column_data()

    def _get_current_cube(self, x, y, num):
        if y == 0:
            if x == num - 2:
                return self.container[attach_column_type[y]].container[x], Cube(value=self.name[x], style=Style(bg_color[4], font=font_style[2], border=side_style[3], al=alignment[1]))
            else:
                return self.container[attach_column_type[y]].container[x], Cube(value=self.name[x], style=Style(bg_color[4], font=font_style[2], al=alignment[1]))
        else:
            return self.container[attach_column_type[y]].container[x], None

    def _render_table_header(self):
        self.main_table.write_cube_to_book(self.origin[0], self.origin[1],
                                           Cube(value='Name', style=Style(bg_color[4], font=font_style[2], al=alignment[1])))
        for i in range(1, 7):
            self.main_table.write_cube_to_book(self.origin[0], self.origin[1] + i,
                                               Cube(value=attach_column_type[i - 1],
                                                    style=Style(bg_color[4], font=font_style[2], al=alignment[1])))

    def render(self):
        self._extract_data()
        self._render_table_title()
        self._render_table_header()
        num_of_row = len(self.name)
        for i in range(num_of_row):
            for j in range(len(attach_column_type)):
                col_cube, name_cube = self._get_current_cube(i, j, num_of_row)
                if name_cube is not None:
                    self.main_table.write_cube_to_book(i + self.origin[0] + 1, j + self.origin[1], name_cube)
                    self.main_table.write_cube_to_book(i + self.origin[0] + 1, j + self.origin[1] + 1, col_cube)
                else:
                    self.main_table.write_cube_to_book(i + self.origin[0] + 1, j + self.origin[1] + 1, col_cube)
        self._cal_end_point()
        self._add_table_border()
