# coding=utf-8
import random

from Person import Person
from MainTable import MainTable
from Common import role
from HeaderTemplate import common_header1
from TableHeader import TableHeader
from AttachedTable import AttachedTable


def get_random_array(num):
    res = []
    for i in range(num):
        res.append(random.randint(1, 9))
    return res


if __name__ == '__main__':

    num = 18

    mock_array = get_random_array(num)
    none_array = [None] * num

    owner_data = {'name': '老大',
                  'target_value': {'Target UC3 $': mock_array, 'Target UNP $': mock_array, 'Target Portfolio %': mock_array},}
    mock_array = get_random_array(num)
    follow_data_1 = {'name': '张三',
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array, 'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array},
                     'LE KL': 11, 'Market size KL': 12, 'Market Share': 13}

    mock_array = get_random_array(num)
    follow_data_2 = {'name': '李四',
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array, 'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array},
                     'LE KL': 10, 'Market size KL': 9, 'Market Share': 8}

    mock_array = get_random_array(num)
    follow_data_3 = {'name': '王五',
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array, 'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array},
                     'LE KL': 10, 'Market size KL': 10, 'Market Share': 10}

    owner = Person(role[1], owner_data)
    f1 = Person(role[2], follow_data_1)
    f2 = Person(role[2], follow_data_2)
    f3 = Person(role[2], follow_data_3)

    follow = [f1, f2, f3]

    main_table = MainTable(13, 2)
    header = TableHeader(common_header1)
    main_table.assign_title('2. 调整ICAM的Portfolio的UC3 UNP以及Portfolio%')
    main_table.assign_header(header)
    main_table.assign_person(follow, owner=owner)
    main_table.render()

    attached_table = AttachedTable(main_table, 3, 2)
    attached_table.assign_data(follow)
    attached_table.assign_title('1. 调整ICAM的Target Volume')
    attached_table.render()

    main_table.save_workbook()







