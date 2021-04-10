#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time: 2020-11-15 18:52
# @Author: gaoxing
# @Email: 1511844263@qq.com
# @File: test01.py

import shutil
import os
import xlwings as xw
import re

# os.chdir("/Users/gonghongwei/Desktop/openpyxl")
def test():
    excel_file = 'D:\\office\\处理后数据\\202010超期明细.xlsx' # 处理后文件路径
    try:
        app = xw.App(visible=False, add_book=False)
        xls = app.books.open(excel_file)
    except:
        return
    sheet = xls.sheets(1)

    info = sheet.used_range
    info.columns.autofit() # 自适应列宽

    nrows = info.last_cell.row
    ncols = info.last_cell.column

    print(nrows)
    print(ncols)



    xls.app.quit()
test()



