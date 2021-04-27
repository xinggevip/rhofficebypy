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
    headNum1 = 3 #主表的表头行数
    headNum2 = 2 #子表的表头行数
    jiaIdColum = 1 #主表假id列位置 和子表关联的id
    subJiaIdColm = 1 #子表假id列位置 和子表关联的id
    zhenIdColum = 3 #主表真id位置 需要写入到子表的id
    try:
        app = xw.App(visible=False, add_book=False)
        xls = app.books.open(excel_file)
    except:
        return

    try:
        '''
        1.拿到主表对象
        2.获取主表的rows，cloums
        3.逐行遍历主表、拿到假id的行数、值、真id值分别的数组-----字典
        4.获取工作表的个数
        5.从第二个工作表开始遍历
        6.拿到子表的行数、列数
        7.逐行遍历子表
        8.按照假id给子表在最后面添加id
        '''
        mainSheet = xls.sheets[0] #主表

        info1 = mainSheet.used_range
        info1.columns.autofit() # 自适应列宽

        nrows1 = info1.last_cell.row
        ncols1 = info1.last_cell.column
        zhenIdTitle = mainSheet.range(headNum1,zhenIdColum).value

        subTabNum = len(xls.sheets)

        print("该工作簿一共又"+str(subTabNum)+"个工作表")
        print("主表一共有"+str(nrows1)+"行")
        print("主表一共有" + str(ncols1) + "列")
        print("主表真id的值",zhenIdTitle)

        idDict = {}

        for y in range(headNum1 + 1, nrows1):
            idDict[mainSheet.range(y, jiaIdColum).value] = mainSheet.range(y, zhenIdColum).value

        # for (key,value) in idDict.items():
        #     print(key,value)

        for i in range(1,len(xls.sheets)):
            print("当前子表的索引为：",i)
            currentSubSheet = xls.sheets[i]
            info2 = currentSubSheet.used_range
            info2.columns.autofit()  # 自适应列宽
            nrows2 = info2.last_cell.row
            ncols2 = info2.last_cell.column
            currentSubSheet.range(headNum2, ncols2 + 1).value = zhenIdTitle
            for y in range(headNum2 + 1, nrows2):
                print("索引为",i,"的表","写入第",y,"行数据的值为",currentSubSheet.range(y, subJiaIdColm).value)
                currentJiaId = currentSubSheet.range(y, subJiaIdColm).value
                currentSubSheet.range(y,ncols2 + 1).value = idDict[currentJiaId]



        #
        # otherSheet = xls.sheets['group3']
        # info2 = otherSheet.used_range
        # info1.columns.autofit()  # 自适应列宽
        #
        # nrows2 = info2.last_cell.row
        # ncols2 = info2.last_cell.column
        #
        # mainIndexArr = []
        # positionArr = []
        # idArr = []
        # for y in range(headNum1 + 1, nrows1):
        #     # print(mainSheet.range(y, 1).value)
        #     positionArr.append(y) # 第几行
        #     mainIndexArr.append(mainSheet.range(y, 1).value)
        #
        #
        # for i in range(headNum2 + 1,nrows2):
        #     value = otherSheet.range(i, 2).value
        #     index1 = otherSheet.range(i, 1).value
        #     # print(value)
        #     if(value != None and value != ""):
        #         if(value == "1.1GAP"):
        #             # print(otherSheet.range(i, 6).value)
        #             if(index1 in mainIndexArr):
        #                 for y in range(0, len(mainIndexArr)):
        #                     if index1 == mainIndexArr[y]:
        #                         print(mainIndexArr[y],positionArr[y],ncols1 + 1,otherSheet.range(i, 6).value)
        #                         mainSheet.range(positionArr[y], ncols1 + 1).value = otherSheet.range(i, 6).value
        #
        xls.save()
        xls.app.quit()




    except Exception as result:
        print('出现了异常')
        xls.save()
        xls.app.quit()
        print(result)



test()