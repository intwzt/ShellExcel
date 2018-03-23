# coding=utf-8
from excel_parser.Parser import Parser
from template.SheetType import RSM

if __name__ == '__main__':
    filename = '../excel_files/test.xlsx'
    parse = Parser(filename, RSM)
    print (parse.parse_attached_table())
    data = parse.parse_main_table()
    for d in data:
        print (d)
