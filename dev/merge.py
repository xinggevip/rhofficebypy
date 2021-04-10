#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2020-12-7 14:56
# @Author:xinggevip
# @File : merge.py
# @Software: PyCharm

import shutil
import os
import xlwings as xw
import re

app = xw.App(visible= False, add_book= False)
# app.display_alerts=False
# app.screen_updating=False

"""
功能目标：
1.多个工作簿中的工作表复制到一个工作簿
2.多个工作簿中的工作表合并为一个工作簿中的一个工作表
"""

newfile = 'C:\\Users\\Administrator\\Desktop\\处理后\\处理后.xlsx' # 处理后文件路径

data_dir = "C:\\Users\\Administrator\\Desktop\恒联超期\\"          # 待处理数据问价路径

"""获取文件列表"""
def getRawFileList(path):
    """-------------------------
    files,names=getRawFileList(raw_data_dir)
    files: ['datacn/dialog/one.txt', 'datacn/dialog/two.txt']
    names: ['one.txt', 'two.txt']
    ----------------------------"""
    files = []
    names = []
    for f in os.listdir(path):
        if not f.endswith("~") or not f == "":  # 返回指定的文件夹包含的文件或文件夹的名字的列表
            files.append(os.path.join(path, f))  # 把目录和文件名合成一个路径
            names.append(f)
    return files, names


def start():
    global wb
    try:
        # global wb2
        files, names = getRawFileList(data_dir)
        print("files:", files)
        print("names:", names)

        # 新建一个excel
        wb = app.books.add(r'C:\Users\Administrator\Desktop\处理后\处理后.xlsx')

        # 2.遍历文件插入到汇总
        for i in range(0, len(files)):

            pass


    except Exception as result:
        print("出现了异常=========================")
        print(result)
    finally:
        wb.app.quit()
        pass




start()
