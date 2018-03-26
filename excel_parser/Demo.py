# coding=utf-8
from excel_parser.Parser import Parser
from template.SheetType import RSM

if __name__ == '__main__':
    filename = '../excel_files/test.xlsx'
    parser = Parser(filename, RSM)
    print parser.parse(2)
