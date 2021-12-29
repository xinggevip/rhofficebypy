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
    global wb1, wb2, app, wb3
    global xPositon
    global yPositon
    excel_file = 'D:\\office\\source\\保证金颓废明细data.xlsx'  # 数据源
    out_path = 'D:\\office\\source\\处理后.xlsx'
    out_path2 = 'D:\\office\\source\\处理后\\'

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

        for i in month_json:
            print(i)


        months = list(items)
        months = sorted(months, reverse=True)
        print(months)
        # 根据月份创建新文件
        for month in months:
        # list1 = [8]
        # for month in list1:
            # 创建新工作簿
            wb3 = app.books.add()  # 这将创建一个新的工作簿
            wb3.save(out_path2 + str(month) + '月份' + '.xlsx')
            sheet4 = wb3.sheets(1)

            num = len(month_json[str(month) + '月份'])
            print("num",num)
            for i in range(num):
                print("进入循环")
                if month_json[str(month) + '月份'][i]['退费金额'] != None:
                    print("进入判断",month_json[str(month) + '月份'][i]['退费金额'],i)
                    # 拷贝表格
                    sheet2.api.Copy(Before=wb3.sheets(1).api)
                    wb3.sheets(1).name = month_json[str(month) + '月份'][i]['姓名']
                    sheet3 = wb3.sheets(1)

                    # sheet3.range(1, 1).expand('table').value = sheet2.range('A1:N9').value

                    # 付款事由
                    sheet3.range('B5').value = str(
                        month_json[str(month) + '月份'][i]['楼栋']) + "号楼" + "业主退装修保证金" + str(
                        month_json[str(month) + '月份'][i]['退费金额']) + '元(收据号：' + str(
                        month_json[str(month) + '月份'][i]['收据号']) + ')'
                    # 收款单位
                    sheet3.range('C6').value = month_json[str(month) + '月份'][i]['姓名']
                    # 开户银行
                    sheet3.range('C7').value = month_json[str(month) + '月份'][i]['开户行']
                    # 账号
                    sheet3.range('K7').value = month_json[str(month) + '月份'][i]['卡号']
                    # 金额
                    sheet3.range('M8').value = month_json[str(month) + '月份'][i]['退费金额']
            # sheet4.delete()
            wb3.save()
            wb3.close()

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