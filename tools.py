# coding=utf-8


def decimal2letter(x):
    result = ''
    while int(x / 26):
        result += chr(ord('@') + x % 26)
        x /= 26
    result += chr(ord('@') + x % 26)
    return result[::-1]


def coordinate_transfer(x, y):
    return decimal2letter(y) + str(x)

