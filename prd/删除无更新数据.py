#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2022-02-18 11:03
# @Author:xinggevip
# @File : 删除无更新数据.py
# @Software: PyCharm

import shutil
import os
import xlwings as xw
import re
import subprocess

def start():
    global wb1, wb2, danwei
    global xPositon
    global yPositon
    excel_file = 'C:\\Users\\Administrator\\Desktop\\待处理档案\\CAP-HR-03 员工档案表_1 (1).xls'  # 数据源
    out_file_path = 'C:\\Users\\Administrator\\Desktop\\待处理档案\\CAP-HR-03 员工档案表_1 (1)处理后.xls'  # 处理后文件地址
    headNum = 1

    # 1.复制模板文件
    shutil.copyfile(excel_file, out_file_path)  # oldfile只能是文件夹，newfile可以是文件，也可以是目标目录

    print('===============开始==================')
    try:
        app = xw.App(visible=False, add_book=False)
        wb1 = app.books.open(out_file_path)
    except:
        print("出现异常")
        wb1.save()
        wb1.app.quit()
        return

    sheet1 = wb1.sheets(1)
    info = sheet1.used_range
    nrows = info.last_cell.row
    ncols = info.last_cell.column

    print('一共' + str(nrows) + '行' + '   , ' + str(ncols) + '列')  # 行数
    # print(ncols)  # 列数

    # 从表头下一致遍历
    for i in range(headNum + 1,nrows + 1):
        for j in range(1,ncols + 1):
            print("(",i,",",j,")\t",end="")
            color = sheet1.range(i,j).color
            if color == None:
                sheet1.range(i,j).value = None
        print(end="\n")
    pass

    wb1.save()
    wb1.app.quit()
    print("运行结束")


if __name__ == '__main__':
    start()
    pass