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
把一个文件根据单位拆分成多个文件
1. 打开原表
2. 找出单位字段所在坐标
3. 遍历单位列
4. 循环遍历创建文件
5. 遍历打开文件删除不相关的数据
"""


def start():
    global wb1, wb2, danwei
    global xPositon
    global yPositon
    excel_file = 'E:\\01--高星--\\01 工作文档\\03 超30天督办事项\\20220302\\workflowData_1.xls'  # 数据源
    out_file_path = 'E:\\01--高星--\\01 工作文档\\03 超30天督办事项\\20220302\\待处理\\'         #处理后文件地址
    keyword = '标题'                                                                             #以此关键字划分文件
    addFileName = '超30天未完结事项需反馈'
    addKeyWord = ['处理情况', '特殊流程未处理说明', '备注']
    bgColor = [255, 255, 0]
    font = 'Calibri'
    blod = True
    headNum = 2
    print('===============开始==================')
    try:
        app = xw.App(visible=False, add_book=False)
        wb1 = app.books.open(excel_file)
    except:
        wb1.save()
        wb1.app.quit()
        return

    sheet1 = wb1.sheets(1)
    info = sheet1.used_range
    nrows = info.last_cell.row
    ncols = info.last_cell.column

    print('一共' + str(nrows) + '行' + '   , ' + str(ncols) + '列')  # 行数
    # print(ncols)  # 列数

    for i in range(1, headNum + 1):  # 遍历前两行
        for y in range(1, ncols + 1):  # 遍历最长列数
            temp = sheet1.range(i, y).value

            if temp != None:
                if temp == keyword:
                    print(str(i) + "     " + str(y) + "     =     " + temp)
                    xPositon = i
                    yPositon = y

    items = set()

    for i in range(xPositon + 1, nrows + 1):
        value = sheet1.range(i, yPositon).value

        matchObj = re.match(r'.*-(.*)', value, re.M | re.I)

        if matchObj:
            print("  matchObj.group() : ", matchObj.group(1))
            value = matchObj.group(1)
            print(str(i) + "    " + str(yPositon) + "    =    " + value)
            items.add(value)
        else:
            print(value + "No match!!")


        pass

    # 获得去查重后的单位列表
    print(items)

    wb1.save()
    wb1.app.quit()

    print('\n===============1.获取关键字列表并去重已完成==================\n')


    # 遍历set复制excel文件并只保留
    for fileName in items:
        newfile = out_file_path + fileName + '.' + excel_file.split(".")[1]
        shutil.copyfile(excel_file, newfile)

        try:
            print(newfile)
            wb2 = app.books.open(newfile)
            sheet2 = wb2.sheets(1)
            #添加字段和样式
            for x in range(0, len(addKeyWord)):
                sheet2.range(headNum, ncols + 1 + x).value = addKeyWord[x]
                sheet2.range(headNum, ncols + 1 + x).color = (bgColor[0], bgColor[1], bgColor[2])
                sheet2.range(headNum, ncols + 1 + x).api.Font.Bold = blod
                sheet2.range(headNum, ncols + 1 + x).api.Font.Name = font

            #设置自适应列宽
            info = sheet2.used_range
            info.columns.autofit()

            #删除和文件名不一样的行
            for i in range(nrows, xPositon, -1):
                value = sheet2.range(i, yPositon).value
                # print('行  ' + str(i) + '   列' + str(yPositon) + '===' + value)

                if value != None:
                    # print(str(i) + "     " + str(y) + "     =     " + temp)

                    matchObj = re.match(r'.*-(.*)', value, re.M | re.I)

                    if matchObj:
                        print("  matchObj.group() : ", matchObj.group(1))
                        value = matchObj.group(1)
                    else:
                        print(value + "No match!!")

                if value == None:
                    sheet2.api.Rows(i).Delete()
                elif value != fileName:
                    print('行  ' + str(i) + '   列' + str(yPositon) + '===' + value + '删除此行')
                    sheet2.api.Rows(i).Delete()
                elif value == fileName:
                    print('行  ' + str(i) + '   列' + str(yPositon) + '===' + value)

            #删除第一列



            wb2.save()
            wb2.app.quit()
        except Exception as result:
            print('出现了异常')
            wb2.save()
            wb2.app.quit()
            print(result)
        pass
    print('\n===============2.拆分创建文件并添加字段和样式已完成==================\n')
    # 批量改后缀
    for fileName in items:
        oldFullFileName = out_file_path + fileName + '.' + excel_file.split(".")[1]
        newFileName = out_file_path + fileName + addFileName + '.' +  excel_file.split(".")[1]
        print(newFileName)
        os.rename(oldFullFileName, newFileName)

    print('\n===============3.批量修改文件名已完成==================\n')

    subprocess.Popen('explorer ' + out_file_path)

    print('\n===============4.处理完毕打开处理后的文件夹==================\n')



start()
