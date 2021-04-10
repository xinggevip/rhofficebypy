#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-3-1 15:49
# @Author:xinggevip
# @File : duban.py
# @Software: PyCharm

import shutil
import os
import xlwings as xw
import re
import subprocess

"""
把此文件夹下的文件设置自适应列宽，并在后面加上三个字段
"""

app = xw.App(visible= False, add_book= False)
# app.display_alerts=False
# app.screen_updating=False

data_dir = "E:\\01--高星--\\01 工作文档\\03 超30天督办事项\\20210301\\待处理\\"          # 待处理数据问价路径

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
    global wb1
    files, names = getRawFileList(data_dir)
    print("names:", names)

    try:
    # 遍历修改文件
        for i in range(0, len(files)):
            wb1 = app.books.open(files[i])
            ws1 = wb1.sheets(1)

            print(i + 1)
            print(names[i])

            ws1.range('H2').value = '处理情况'
            ws1.range('H2').color = (255,255,0)
            ws1.range('H2').api.Font.Bold = True
            ws1.range('H2').api.Font.Name = 'Arial'

            ws1.range('I2').value = '特殊流程未处理说明'
            ws1.range('I2').color = (255, 255, 0)
            ws1.range('I2').api.Font.Bold = True
            ws1.range('I2').api.Font.Name = 'Arial'

            ws1.range('J2').value = '备注'
            ws1.range('J2').color = (255, 255, 0)
            ws1.range('J2').api.Font.Bold = True
            ws1.range('J2').api.Font.Name = 'Arial'

            info = ws1.used_range
            info.columns.autofit()

            wb1.save()
            wb1.app.quit()

    except Exception as result:
        wb1.save()
        wb1.app.quit()
        print("出现了异常=========================")
        print(result)

    print("==============结束==============")

    subprocess.Popen('explorer ' + data_dir)

start()