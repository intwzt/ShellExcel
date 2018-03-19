# coding=utf-8
import random

from Person import Person
from Table import Table
from Common import role
from HeaderTemplate import common_header1
from TableHeader import TableHeader


def get_random_array(num):
    res = []
    for i in range(num):
        res.append(random.randint(1, 100))
    return res


if __name__ == '__main__':

    num = 18

    mock_array = get_random_array(num)
    none_array = [None] * num

    owner_data = {'name': '老大',
                  'target_value': {'Target UC3 $': mock_array, 'Target UNP $': mock_array, 'Target Portfolio %': mock_array}}
    mock_array = get_random_array(num)
    follow_data_1 = {'name': '张三',
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array, 'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array}}

    mock_array = get_random_array(num)
    follow_data_2 = {'name': '李四',
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array, 'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array}}

    mock_array = get_random_array(num)
    follow_data_3 = {'name': '王五',
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array, 'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array}}

    owner = Person(role[1], owner_data)
    f1 = Person(role[2], follow_data_1)
    f2 = Person(role[2], follow_data_2)
    f3 = Person(role[2], follow_data_3)

    follow = [f1, f2, f3]

    table = Table(5, 3)
    header = TableHeader(common_header1)
    table.assign_title('2. 调整ICAM的Portfolio的UC3 UNP以及Portfolio%')
    table.assign_header(header)
    table.assign_person(owner, follow)
    table.render()
    table.save_workbook()







