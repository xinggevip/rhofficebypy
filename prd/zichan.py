#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time: 2021-04-01 20:20
# @Author: gaoxing
# @Email: 1511844263@qq.com
# @File: zichan.py

import shutil
import os
import xlwings as xw
import re

""""
如果部门不存在则，设置为空
"""

zichan = "C:\\Users\\15118\\Desktop\\待处理\\除集团正常资产.xlsx"
bumen1 = "C:\\Users\\15118\\Desktop\\待处理\\departments_admin1.xls"
data_dir = "C:\\Users\\15118\\Desktop\待处理\\"


def start():
    global zichan_xls, bumen_xls
    try:
        app = xw.App(visible=False, add_book=False)
        zichan_xls = app.books.open(zichan)
        bumen_xls = app.books.open(bumen1)


        """"
        获取这两个表单的最大长度
        """
        zichan_xls_sheet = zichan_xls.sheets(1)
        bumen_xls_sheet = bumen_xls.sheets(1)

        zichan_xls_sheet_info = zichan_xls_sheet.used_range
        bumen_xls_sheet_info = bumen_xls_sheet.used_range

        nrows1 = zichan_xls_sheet_info.last_cell.row
        ncols1 = zichan_xls_sheet_info.last_cell.column

        nrows2 = bumen_xls_sheet_info.last_cell.row
        ncols2 = bumen_xls_sheet_info.last_cell.column

        print("资产表的行数", nrows1, "列数", ncols1)
        print("部门表行数", nrows2, "列数", ncols2)



        zichan_dep_arr = []
        admin_dep_arr = set()

        for i in range(2,nrows2 + 1):
            item1 = bumen_xls_sheet.range(i, 1).value
            item2 = bumen_xls_sheet.range(i, 2).value
            item3 = bumen_xls_sheet.range(i, 3).value
            admin_dep_arr.add(item1)
            admin_dep_arr.add(item2)
            admin_dep_arr.add(item3)
            print(item1, item2, item3)

        num1 = 1
        num2 = 1

        diff_bumen = []
        diff_danwei = []
        for i in range(2,nrows1 + 1):
            danwei = zichan_xls_sheet.range(i,12).value
            bumen = zichan_xls_sheet.range(i,13).value
            bianhao = zichan_xls_sheet.range(i, 1).value
            # zichan_dep_arr.append(item)
            item = {"bianhao":bianhao,"value":bumen}
            if bumen not in admin_dep_arr:
                print(num1)
                print(item)
                num1 = num1 + 1
                diff_bumen.append(item)
                zichan_xls_sheet.range(i, 13).value = ""
            if danwei not in admin_dep_arr:
                print(num2)
                print(item)
                num2 = num2 + 1
                diff_danwei.append(item)
                zichan_xls_sheet.range(i, 12).value = ""

        fileObject = open('C:\\Users\\15118\\Desktop\待处理\\result.txt', 'w')
        for ip in diff_bumen:
            fileObject.write(str(ip))
            fileObject.write('\n')

        fileObject.write("================================")
        for ip in diff_danwei:
            fileObject.write(str(ip))
            fileObject.write('\n')
        fileObject.close()

        zichan_xls.save()
        bumen_xls.save()
        zichan_xls.app.quit()

        os.startfile(data_dir)

        pass
    except Exception as result:
        print(result)
        zichan_xls.save()
        bumen_xls.save()
        zichan_xls.app.quit()

        pass
    pass


start()

