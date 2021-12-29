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
通用把一个文件根据单位拆分多个sheet
"""

def start(excel_file,out_file,huizongTitle):
    global wb1, wb2,app
    global xPositon
    global yPositon
    # excel_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\test\\拆分\\车辆信息管理5月原始数据.xlsx'  # 数据源
    # out_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\test\\拆分\\2021年5月瑞华集团公务车信息汇总表.xlsx'
    keywordList = ['现职单位']                                                                             #以此关键字划分文件
    # 上个月汇总表合计一行的关键字标识
    preHuZongHejiKey = "合计"
    # 原始表的表头行数
    headNum = 3
    print('===============开始==================')
    try:
        app = xw.App(visible=False, add_book=False)
        wb1 = app.books.open(excel_file)
    except Exception as result:
        print('出现了异常')
        print(result)
        wb1.save()
        wb1.close()
        app.quit()
        return

    try:
        sheet1 = wb1.sheets(1)
        info = sheet1.used_range
        nrows = info.last_cell.row
        ncols = info.last_cell.column
        wb2 = app.books.add()  # 这将创建一个新的工作簿
        sheet2 = wb2.sheets(1)

        print('一共' + str(nrows) + '行' + '   , ' + str(ncols) + '列')  # 行数
        # print(ncols)  # 列数

        for keyword in keywordList:

            for i in range(1, headNum + 1):  # 遍历前两行
                for y in range(1, ncols + 1):  # 遍历最长列数
                    temp = sheet1.range(i, y).value
                    if temp != None:
                        # print(str(i) + "     " + str(y) + "     =     " + temp)
                        if temp == keyword:
                            print(str(i) + "     " + str(y) + "     =     " + temp)
                            xPositon = i
                            yPositon = y

            items = set()

            for i in range(xPositon + 1, nrows + 1):
                value = sheet1.range(i, yPositon).value
                print(str(i) + "    " + str(yPositon) + "    =    " + value)
                items.add(value)
                pass

            # 获得去查重后的单位列表
            print(items)

            print('\n===============1.获取关键字列表并去重已完成==================\n')
            danweiArr = list(items)


            for sheetName in danweiArr:
                key = sheetName
                print(sheetName)
                sheetName = companyToSimple(sheetName)
                print(sheetName)
                sheet1.api.Copy(Before=wb2.sheets(1).api)
                wb2.sheets(1).name = sheetName
                # 取到当前插入的新Sheet
                sheet3 = wb2.sheets(1)
                # 删除和文件名不一样的行
                for i in range(nrows, xPositon, -1):
                    value = sheet3.range(i, yPositon).value
                    # print('行  ' + str(i) + '   列' + str(yPositon) + '===' + value)
                    if value == None:
                        sheet3.api.Rows(i).Delete()
                    elif value != key:
                        print('行  ' + str(i) + '   列' + str(yPositon) + '===' + value + '删除此行')
                        sheet3.api.Rows(i).Delete()
                    elif value == key:
                        print('行  ' + str(i) + '   列' + str(yPositon) + '===' + value + '保留此行')
                info = sheet3.used_range
                info.columns.autofit()
                pass

        sheet1.api.Copy(Before=wb2.sheets(1).api)
        wb2.sheets(1).name = "全部"
        sheet2.delete()
        
        # 自适应列宽
        info = wb2.sheets(1).used_range
        info.columns.autofit()

        
        wb1.save()
        wb2.save(out_file)
        wb1.close()
        wb2.close()
        app.quit()
    except Exception as result:
        print('出现了异常')
        print(result)
        wb1.save()
        wb2.save(out_file)
        wb1.close()
        wb2.close()
        app.quit()





# 将全称单位转换成简称
def companyToSimple(company):
    # 直接return
    goReturnList = ["集团本部","救援中心"]
    # 无规则return
    goReturnOfNoRuleList = [
        "瑞华机动车登记服务站",
        "河南瑞铭二手车",
        "河南耀泓汽车配件销售有限公司",
        "河南南泓仓储物流有限公司",
        "郑州南瑞汽车配件销售有限公司",
        "河南南泓汽车贸易有限公司",
        "巩义市德嘉汽车销售服务有限公司",
        "新密市瑞利汽车销售有限公司"
    ]

    # 直接return
    if company in goReturnList:
        return company
        pass
    # 无规则return
    if company in goReturnOfNoRuleList:
        if company == "瑞华机动车登记服务站":
            return "服务站"
        if company == "河南瑞铭二手车":
            return "二手车"
        if company == "河南耀泓汽车配件销售有限公司":
            return "耀泓"
        if company == "河南南泓仓储物流有限公司":
            return "南泓仓储"
        if company == "郑州南瑞汽车配件销售有限公司":
            return "南瑞"
        if company == "河南南泓汽车贸易有限公司":
            return "南泓汽贸"
        if company == "巩义市德嘉汽车销售服务有限公司":
            return "德嘉"
        if company == "新密市瑞利汽车销售有限公司":
            return "瑞利"
    # 有规则return
    matchObj = re.match(r'河南(.*?)汽车', company, re.M | re.I)

    if matchObj:
        return matchObj.group(1)
    else:
        print(company + "No match!!")
        return company

# 根据关键获取去重后的字段
def getSetData(wb,keywordList,headNum):
    global xPositon, yPositon
    sheet1 = wb.sheets["原始表"]
    info = sheet1.used_range
    nrows = info.last_cell.row
    ncols = info.last_cell.column
    resArr = []

    print('一共' + str(nrows) + '行' + '   , ' + str(ncols) + '列')  # 行数
    # print(ncols)  # 列数

    for keyword in keywordList:

        for i in range(1, headNum + 1):  # 遍历前两行
            for y in range(1, ncols + 1):  # 遍历最长列数
                temp = sheet1.range(i, y).value
                if temp != None:
                    # print(str(i) + "     " + str(y) + "     =     " + temp)
                    if temp == keyword:
                        print(str(i) + "     " + str(y) + "     =     " + temp)
                        xPositon = i
                        yPositon = y

        items = set()

        for i in range(xPositon + 1, nrows + 1):
            value = sheet1.range(i, yPositon).value
            print(str(i) + "    " + str(yPositon) + "    =    " + value)
            items.add(value)
            pass
            
        

        # 获得去查重后的单位列表
        print(items)

        print('\n===============1.获取关键字列表并去重已完成==================\n')
        dataArr = list(items)
        dictData = {}
        dictData['keyword'] = keyword
        dictData['dataArr'] = dataArr
        dictData['xPositon'] = xPositon
        dictData['yPositon'] = yPositon
        resArr.append(dictData)

    return resArr

if __name__ == '__main__':
    excel_file1 = 'C:\\Users\\Administrator\\Desktop\\服务顾问CAP-HR-03 员工档案表_1.xls'  # 数据源
    out_file1 = 'C:\\Users\\Administrator\\Desktop\\服务顾问CAP-HR-03 员工档案表_1_处理后.xls'
    huizongTitle1 = "2021年4月瑞华集团公务车信息汇总表"
    start(excel_file1,out_file1,huizongTitle1)

  


