# coding=utf-8


class TableHeader:
    def __init__(self, data):
        self.x = len(data['matrix'])
        self.y = len(data['matrix'][0])
        self.matrix = data['matrix']
        self.merge = None if 'merge' not in data else data['merge']


