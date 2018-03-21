# coding=utf-8
import random

from Sheet import Sheet
from Common import role
from SheetType import RSM


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
                  'role': role[1],
                  'target_value': {'Target UC3 $': mock_array, 'Target UNP $': mock_array,
                                   'Target Portfolio %': mock_array}, }
    mock_array = get_random_array(num)
    follow_data_1 = {'name': '张三',
                     'role': role[2],
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array,
                                      'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array},
                     'LE KL': 11, 'Market size KL': 12, 'Market Share': 13}

    mock_array = get_random_array(num)
    follow_data_2 = {'name': '李四',
                     'role': role[2],
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array,
                                      'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array},
                     'LE KL': 10, 'Market size KL': 9, 'Market Share': 8}

    mock_array = get_random_array(num)
    follow_data_3 = {'name': '王五',
                     'role': role[2],
                     'target_value': {'Target UC3 $': none_array, 'Target UNP $': none_array,
                                      'Target Portfolio %': none_array},
                     'ref_value': {'Ref UC3 $': mock_array, 'Ref UNP $': mock_array, 'Ref Portfolio %': mock_array},
                     'LE KL': 10, 'Market size KL': 10, 'Market Share': 10}

    if __name__ == '__main__':
        data = [owner_data, follow_data_1, follow_data_3, follow_data_3]
        sheet = Sheet(data, RSM)
        sheet.render()
        sheet.save_sheet()
