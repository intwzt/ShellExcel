# coding=utf-8
from Common import role, column_type, target_mapper, ref_mapper, column_last_row, formula_type
from Column import Column


class Person:
    def __init__(self, person_role, data):
        person_role = person_role
        self.role = person_role
        self.name = ''
        self.ref = None

        # attached table value
        self.LE = None
        self.Market_Size = None
        self.Market_Share = None

        if person_role == role[2]:
            self.ref = {ref_mapper[0]: [], ref_mapper[1]: [], ref_mapper[2]: []}
        self.target = {target_mapper[0]: [], target_mapper[1]: [], target_mapper[2]: []}
        self.assign_all_value(data)

    # custom your rules
    def _check_last_column(self, key):
        if key == target_mapper[2] or key == ref_mapper[2]:
            return column_last_row[3]
        elif key in target_mapper or key in ref_mapper:
            return column_last_row[2]
        else:
            return column_last_row[4]

    def set_column_value(self, num, target_value, ref_value, name=None):
        self.name = name
        for k, v in target_value.iteritems():
            if k == target_mapper[2]:
                column = Column(num, column_type[1], self.role, last_row=self._check_last_column(k), number_format=10)
            else:
                column = Column(num, column_type[1], self.role, last_row=self._check_last_column(k))
            column.set_column_value(v)
            self.target[k] = column
        if self.role == role[1]:
            return
        if self.role == role[2]:
            for k, v in ref_value.iteritems():
                if k == ref_mapper[2]:
                    column = Column(num, column_type[2], self.role, last_row=self._check_last_column(k), number_format=10)
                else:
                    column = Column(num, column_type[2], self.role, last_row=self._check_last_column(k))
                column.set_column_value(v)
                self.ref[k] = column

    def assign_all_value(self, data):
        ref_value = None
        target_value = None
        if 'ref_value' in data:
            ref_value = data['ref_value']
        if 'target_value' in data:
            target_value = data['target_value']
        num = len(data['target_value'][target_mapper[0]])
        if self.role == role[2]:
            self.LE = data['LE KL']
            self.Market_Size = data['Market size KL']
            self.Market_Share = data['Market Share']
        self.set_column_value(num, target_value, ref_value, data['name'])

    def attached_column_name_mapper(self, col_name):
        if col_name == 'LE KL':
            return self.LE
        elif col_name == 'Market size KL':
            return self.Market_Size
        elif col_name == 'Market Share %':
            return self.Market_Share
        else:
            print 'error column name', col_name, 'for attached table'
            return None
