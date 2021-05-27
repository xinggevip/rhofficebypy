#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-4-29 15:47
# @Author:xinggevip
# @File : xianjia.py
# @Software: PyCharm

#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-3-3 14:31
# @Author:xinggevip
# @File : baoxiao.py
# @Software: PyCharm

import shutil
import os
import xlwings as xw
import re
import subprocess

"""
汇总、拆分限价违规表
1. 复制数据源到指定地址
2. 打开待处理工作簿
3. 找出指定字段所在坐标。变量定义
3. 遍历去重后的单位列
4. 在表中复制单位数个表
5. 遍历表，查到和sheet表明不一致则删掉
"""


def start():
    global wb1, wb2
    global xPositon
    global yPositon
    excel_file = 'E:\\01--高星--\\01 工作文档\\待办\\车辆销售数据\\sql导出表格\\长城哈弗\\数据源\\长城哈弗新.xls'  # 数据源
    out_file_path = 'E:\\01--高星--\\01 工作文档\\待办\\车辆销售数据\\sql导出表格\\长城哈弗\\数据源\\长城哈弗新处理后.xls'         #处理后文件地址
    keyword = '单位'                                                                             #以此关键字划分文件
    addFileName = ''
    addKeyWord = []
    bgColor = [255, 255, 0]
    font = 'Calibri'
    blod = True
    headNum = 2

    print('===============开始==================')
    # 复制文件
    shutil.copyfile(excel_file, out_file_path)


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

    print('一共' + str(nrows1) + '行' + '   , ' + str(ncols1) + '列')
    subTabNum = len(wb1.sheets)
    print("该工作簿一共又" + str(subTabNum) + "个工作表")

    for i in range(1, headNum + 1):  # 遍历前两行
        for y in range(1, ncols1 + 1):  # 遍历最长列数
            temp = sheet1.range(i, y).value
            if temp != None:
                # print(str(i) + "     " + str(y) + "     =     " + temp)
                if temp == keyword:
                    print(str(i) + "     " + str(y) + "     =     " + temp)
                    xPositon = i
                    yPositon = y

    items = set()

    for i in range(xPositon + 1, nrows1 + 1):
        value = sheet1.range(i, yPositon).value
        print(str(i) + "    " + str(yPositon) + "    =    " + value)
        items.add(value)
        pass

    # 获得去查重后的单位列表
    # print(items)
    itemList = list(items)
    for i in range(0,len(itemList)):
        print(i,itemList[i])





    wb1.save()
    wb1.app.quit()


start()
