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

    out_book = Book("out.xlsx")
    out_book.create()
    out_sheet = out_book.get_active_sheet()
    for row_index in range(1, max_row + 1):
        for col_index in range(1, max_col + 1):
            print(sheet1.get_cell(row_index, col_index).get_str_val())
    print("2秒后退出")
    time.sleep(2)

main()
