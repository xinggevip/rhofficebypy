#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-12-25 14:46
# @Author:xinggevip
# @File : yajin.py
# @Software: PyCharm


import pandas as pd
import time
import datetime
import os
import xlwings as xw

def start():
    # 读取表到pd
    data_path = 'D:\\office\\yajin\\data\\12月退押金90户.xls'
    tui_path = 'D:\\office\\yajin\\data\\保证金退费明细.xlsx'

    out_path = 'D:\\office\\yajin\\保证金退费明细处理后.xlsx'

    global xPositon
    global yPositon

    bgColor = [255, 255, 0]


    data = pd.read_excel(data_path, 'Sheet1', skiprows=2)
    print(data)
    tui = pd.read_excel(tui_path, '装修保证金退费', skiprows=1)
    print(tui)

    data['住址'] = data['楼栋'].astype('str') + '-' + data['房号'].astype('str')
    # print(data.columns.values)
    # print(data['楼栋'])
    print(data)
    tui = tui.dropna(subset = ['楼栋'])
    # tui['日期'] = pd.to_datetime(tui["日期"], errors='coerce')

    tui['住址'] = tui['楼栋'].astype('int').astype('str') + '-' + tui['户室'].astype('int').astype('str')
    print(data)

    tui['标黄'] = ''

    zhu_arr = data['住址'].tolist()
    for index, row in tui.iterrows():
        value1 = row['住址']
        if value1 in zhu_arr:
            if row['退费金额'] == 2000:
                tui.loc[index,'标黄'] = '是'
                print(tui.loc[index,:])

    # 输出
    tui.to_excel(out_path, sheet_name="装修保证金退费", index=False)

    # 标黄
    try:
        app = xw.App(visible=False, add_book=False)
        wb1 = app.books.open(out_path)
    except:
        wb1.save()
        wb1.app.quit()
        return

    sheet1 = wb1.sheets(1)
    info = sheet1.used_range
    nrows = info.last_cell.row
    ncols = info.last_cell.column

    print('一共' + str(nrows) + '行' + '   , ' + str(ncols) + '列')  # 行数


    for i in range(1, 1 + 1):  # 遍历前1行
        for y in range(1, ncols + 1):  # 遍历最长列数
            temp = sheet1.range(i, y).value
            if temp != None:
                # print(str(i) + "     " + str(y) + "     =     " + temp)
                if temp == "标黄":
                    print(str(i) + "     " + str(y) + "     =     " + temp)
                    xPositon = i
                    yPositon = y


    for i in range(2,nrows + 1):
        value = sheet1.range(i,yPositon).value
        print("正在处理第",i,"行")
        if value == '是':
            for y in range(1,ncols + 1):
                sheet1.range(i,y).color = (bgColor[0], bgColor[1], bgColor[2])

    wb1.save()
    wb1.app.quit()
    print("==========运行结束=========")




    print(tui)









    pass




if __name__ == '__main__':
    start()