# coding=utf-8
from MainTable import MainTable
from AttachedTable import AttachedTable
from Common import role
from Person import Person
from SheetType import sheet_type
from TableHeader import TableHeader


class Sheet:
    def __init__(self, data, table_type, x=2, y=2):
        self.origin = [x, y]
        self.main_origin = [x, y]
        self.owner = None
        self.follower = []
        self.table_type = table_type
        self.main_table = None
        self.attached_table = None
        for p in data:
            if p['role'] == role[1]:
                self.owner = Person(role[1], p)
            elif p['role'] == role[2]:
                self.follower.append(Person(role[2], p))
            else:
                print 'error role input'
                pass

    def _cal_main_table_coordinate(self):
        num_of_follower = len(self.follower)
        self.main_origin = [self.main_origin[0] + num_of_follower + 4 + 2, self.main_origin[1]]

    def render(self):
        self._cal_main_table_coordinate()

        main_table = MainTable(self.main_origin[0], self.main_origin[1])
        header = TableHeader(sheet_type[self.table_type]['header'])
        main_table.assign_title(sheet_type[self.table_type]['title'][2])
        main_table.assign_header(header)
        main_table.assign_person(self.follower, owner=self.owner)
        self.main_table = main_table
        main_table.render()

        attached_table = AttachedTable(main_table, self.origin[0], self.origin[1])
        attached_table.assign_data(self.follower)
        attached_table.assign_title(sheet_type[self.table_type]['title'][1])
        self.attached_table = attached_table
        attached_table.render()

    def save_sheet(self):
        self.main_table.save_workbook()
