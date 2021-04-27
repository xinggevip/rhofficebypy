#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-4-26 16:04
# @Author:xinggevip
# @File : addJGap.py
# @Software: PyCharm

import shutil
import os
import xlwings as xw
import re

def test():
    excel_file = 'E:\\01--高星--\\01 工作文档\\待办\\车辆销售数据\\新建文件夹\\CAP-ZCXS-02 新车销售单_4.xls' # 待处理
    headNum1 = 3
    headNum2 = 2
    try:
        app = xw.App(visible=False, add_book=False)
        xls = app.books.open(excel_file)
    except:
        return

    try:
        mainSheet = xls.sheets['督办销售管理'] #主表

        info1 = mainSheet.used_range
        info1.columns.autofit() # 自适应列宽

        nrows1 = info1.last_cell.row
        ncols1 = info1.last_cell.column

        otherSheet = xls.sheets['group3']
        info2 = otherSheet.used_range
        info1.columns.autofit()  # 自适应列宽

        nrows2 = info2.last_cell.row
        ncols2 = info2.last_cell.column

        mainIndexArr = []
        positionArr = []
        for y in range(headNum1 + 1, nrows1):
            # print(mainSheet.range(y, 1).value)
            positionArr.append(y)
            mainIndexArr.append(mainSheet.range(y, 1).value)


        for i in range(headNum2 + 1,nrows2):
            value = otherSheet.range(i, 2).value
            index1 = otherSheet.range(i, 1).value
            # print(value)
            if(value != None and value != ""):
                if(value == "1.1GAP"):
                    # print(otherSheet.range(i, 6).value)
                    if(index1 in mainIndexArr):
                        for y in range(0, len(mainIndexArr)):
                            if index1 == mainIndexArr[y]:
                                print(mainIndexArr[y],positionArr[y],ncols1 + 1,otherSheet.range(i, 6).value)
                                mainSheet.range(positionArr[y], ncols1 + 1).value = otherSheet.range(i, 6).value

        xls.save()
        xls.app.quit()


        print(nrows1)
        print(ncols1)
        print(nrows2)
        print(ncols2)

    except Exception as result:
        print('出现了异常')
        xls.save()
        xls.app.quit()
        print(result)



test()