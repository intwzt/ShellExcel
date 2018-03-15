# coding=utf-8
from Person import Person
from Table import Table
from Common import role
from HeaderTemplate import common_header1
from TableHeader import TableHeader


if __name__ == '__main__':

    mock_array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 1, 2, 3, 4, 5, 6, 7, 8, 9]

    owner_data = {'name': '老大',
                  'target_value': {'Target UC3 $': mock_array, 'Target UNP $': mock_array,
                                   'Target Portfolio %': mock_array}}
    follow_data_1 = {'name': '张三',
                     'target_value': {'Target UC3 $': mock_array, 'Target UNP $': mock_array,
                                      'Target Portfolio %': mock_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array}}
    follow_data_2 = {'name': '李四',
                     'target_value': {'Target UC3 $': mock_array, 'Target UNP $': mock_array,
                                      'Target Portfolio %': mock_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array}}
    follow_data_3 = {'name': '王五',
                     'target_value': {'Target UC3 $': mock_array, 'Target UNP $': mock_array,
                                      'Target Portfolio %': mock_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array}}

    owner = Person(role[1], owner_data)
    f1 = Person(role[2], follow_data_1)
    f2 = Person(role[2], follow_data_2)
    f3 = Person(role[2], follow_data_3)

    follow = [f1, f2, f3]

    table = Table(1, 1)
    header = TableHeader(common_header1)
    table.assign_header(header)
    table.assign_person(owner, follow)
    table.render()
    table.save_workbook()


