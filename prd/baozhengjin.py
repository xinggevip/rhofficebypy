#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time: 2021-11-15 21:38
# @Author: gaoxing
# @Email: 1511844263@qq.com
# @File: baozhengjin.py

import shutil
import os
import traceback

import xlwings as xw
import re

def start():
    global wb1, wb2, app
    global xPositon
    global yPositon
    excel_file = 'D:\\office\\source\\保证金颓废明细data.xlsx'  # 数据源
    out_path = 'D:\\office\\source\\处理后.xlsx'

    print('===============开始==================')
    try:
        app = xw.App(visible=False, add_book=False)
        wb1 = app.books.open(excel_file)
        wb2 = app.books.open(out_path)
    except Exception as result:
        print('出现了异常')
        print(result)
        wb1.save()
        wb1.close()
        wb2.save()
        wb2.close()
        app.quit()
        return

    # 获取一共有几月

    try:
        sheet1 = wb1.sheets['Sheet1']
        sheet2 = wb2.sheets['模板']
        info = sheet1.used_range
        nrows = info.last_cell.row
        ncols = info.last_cell.column

        print('一共' + str(nrows) + '行' + '   , ' + str(ncols) + '列')

        items = set()

        month_json = {}


        for i in range(2,nrows + 1):
            value = sheet1.range(i, 1).value

            # print('type',type(value),'==============行', i,'值 ',value)
            if value == None or value == '':
                print('==============行',i,'，列1为空值')
                sheet1.range(1, i).value = '空'
            else:
                print(value)
                print(value.year,value.month,value.day,"行 ",i)

                # 如果月份已不存在，则创建，已存在则添加
                item = {
                }
                for y in range(1, ncols + 1):
                    value2 = sheet1.range(1, y).value
                    item[value2] = sheet1.range(i, y).value

                if str(value.month) + '月份' in month_json:
                    month_json[str(value.month) + '月份'].append(item)
                else:
                    month_json[str(value.month) + '月份'] = []
                    month_json[str(value.month) + '月份'].append(item)
            items.add(value.month)

        print(items)
        # print(month_json)

        for k,v in month_json.items():
            month_json[k] = list(filter(lambda x:x['退费金额'] != None,v))



        months = list(items)
        months = sorted(months, reverse=True)
        print(months)
        # 根据月份创建新文件
        for month in months:
        # list1 = [8]
        # for month in list1:
            sht3 = wb2.sheets.add(str(month) + '月份', after=sheet2)  # 新建工作表，放在sht工作表后面。
            num = len(month_json[str(month) + '月份'])
            print("num",num)
            for i in range(num):
                print("进入循环")
                if month_json[str(month) + '月份'][i]['退费金额'] != None:
                    # print("进入判断")
                    print("进入判断", i,'/',num,'/',str(month),'月份')

                    # 按行复制
                    for y in range(12):
                        sheet2.api.Rows(y+1).Copy(sht3.api.Rows(i * 11 + y + 1))

                    # sht3.range(i * 11 + 1,1).expand('table').value = sheet2.range('A1:N9').value

                    # 付款事由
                    # sht3.range(i * 11 + 5,2).value = str(month_json[str(month) + '月份'][i]['楼栋']) + "号楼" + "业主退装修保证金" + str(month_json[str(month) + '月份'][i]['退费金额']) + '元(收据号：'+ str(month_json[str(month) + '月份'][i]['收据号'])+')'
                    sht3.range(i * 11 + 5,2).value = str(month_json[str(month) + '月份'][i]['楼栋']) + "号楼" + "业主退装修保证金" + str(month_json[str(month) + '月份'][i]['退费金额']) + '元(收据号：'+ str(month_json[str(month) + '月份'][i]['收据号'])+')'
                    # 收款单位
                    # sht3.range(i * 11 + 6,2).value = month_json[str(month) + '月份'][i]['姓名']
                    sht3.range('C'+str(i * 11 + 6)).value = month_json[str(month) + '月份'][i]['姓名']
                    # 开户银行
                    # sht3.range(i * 11 + 7,2).value = month_json[str(month) + '月份'][i]['开户行']
                    sht3.range('C'+str(i * 11 + 7)).value = month_json[str(month) + '月份'][i]['开户行']
                    # 账号
                    # sht3.range(i * 11 + 7,4).value = month_json[str(month) + '月份'][i]['卡号']
                    sht3.range('K'+str(i * 11 + 7)).value = month_json[str(month) + '月份'][i]['卡号']
                    # 金额
                    # sht3.range(i * 11 + 8,3).value = month_json[str(month) + '月份'][i]['退费金额']
                    sht3.range('M'+str(i * 11 + 8)).value = month_json[str(month) + '月份'][i]['退费金额']

                    # 处理样式
                    # sht3.range(i * 11 + 1,14).api.merge()

    # 写入格式和数据
        wb1.save()
        wb1.close()
        wb2.save()
        wb2.close()
        app.quit()
        print("结束")
    except Exception as result:
        print('出现了异常')
        print(result)
        print('repr(e):\t', repr(result))
        wb1.save()
        wb1.close()
        wb2.save()
        wb2.close()
        app.quit()
        print("结束")





if __name__ == '__main__':
    start()
    pass