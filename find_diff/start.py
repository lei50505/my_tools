#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''doc'''

import time
from util import Book

def main():
    '''doc'''
    book = Book("in.xlsx")
    book.load()
    sheet1 = book.get_sheet("Sheet1")
    max_col = sheet1.get_max_col()
    max_row = sheet1.get_max_row()


    for col_index in range(1, max_col):
        for row_index in range(1, max_row):
            print(sheet1.get_cell(row_index, col_index).get_str_val())
    print("2秒后退出")
    time.sleep(2)

main()
