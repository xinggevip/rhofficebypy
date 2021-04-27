#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-3-11 16:13
# @Author:xinggevip
# @File : baoxiaoMerge.py
# @Software: PyCharm

import shutil
import os
import xlwings as xw
import re
import subprocess

"""
把多个表合并成一个表
1. 得到文件列表
2. 复制第一个表
3. 遍历表
4. 复制行到新表
"""

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
    global wb1, wb2, danwei
    out_file_path = 'E:\\01--高星--\\01 工作文档\\03 超30天未报销事项\\20210310\\已合并\\固定资产合并后10.xlsx'
    # data_src_path = 'E:\\01--高星--\\01 工作文档\\03 超30天未报销事项\\20210310\\已接收\\'
    data_src_path = 'C:\\Users\\Administrator\\Desktop\\待处理标签\\合并前\\'

    open_path = 'E:\\01--高星--\\01 工作文档\\03 超30天未报销事项\\20210310\\已合并\\'
    head_num = 2

    files, names = getRawFileList(data_src_path)
    shutil.copyfile(files[0], out_file_path)

    try:
        app = xw.App(visible=False, add_book=False)
        wb1 = app.books.open(out_file_path)
    except:
        wb1.save()
        wb1.app.quit()
        return

    sheet1 = wb1.sheets(1)
    info1 = sheet1.used_range
    nrows1 = info1.last_cell.row
    ncols1 = info1.last_cell.column

    count = nrows1 + 1

    for i in range(1, len(files)):
        print(files[i])
        wb2 = app.books.open(files[i])
        sheet2 = wb2.sheets(1)
        info2 = sheet1.used_range
        nrows2 = info2.last_cell.row
        ncols2 = info2.last_cell.column

        for sub_tab_row in range(head_num + 1, nrows2 + 1):
            value = sheet2.range(sub_tab_row, 1).value
            if value != None and value != '':
                print(value)
                sheet2.api.Rows(sub_tab_row).Copy(sheet1.api.Rows(count))
                count = count + 1
            pass

        pass

    wb1.save()
    wb2.save()
    wb1.app.quit()

    subprocess.Popen('explorer ' + open_path)

    pass

start()