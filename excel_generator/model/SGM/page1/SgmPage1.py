# coding=utf-8


class SgmPage1:
    def __init__(self, ws, title, header, data, x=3, y=2):
        self.ws = ws
        self.title = title
        self.header = header
        self.data = data
        self.origin = [x, y]

