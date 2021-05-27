#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time: 2021-05-16 14:29
# @Author: gaoxing
# @Email: 1511844263@qq.com
# @File: 公务车数量监测.py

import shutil
import os
import traceback

import xlwings as xw
import re
import subprocess

"""
功能：
    和上个月的车辆数量进行对比，将增加的车辆和减少的车辆信息汇总出来
步骤：
    1.打开这个月的车辆表和上个月的车辆表

"""


def start():
    global wb1,wb2,wb3,app,sheet3, xPositon, yPositon
    # 上个月数据
    pre_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\work\\汇总\\2021年4月份瑞华集团公务车信息汇总表.xlsx'
    # 这个月数据
    now_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\work\\拆分\车辆信息管理拆分后.xlsx'
    # 输出数据
    out_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\work\\汇总\\2021年5月份瑞华集团公务车信息汇总表.xlsx'
    # 以此关键字划分文件
    keywordList = ['使用单位', '二级类别']
    # 上个月汇总表合计一行的关键字标识
    preHuZongHejiKey = "合计"
    # 原始表的表头行数
    headNum = 2
    # 汇总表标题
    huizongTitle = "2021年5月瑞华集团公务车信息汇总表"
    # 单位 为了固定顺序而定义
    companyList = [
        '集团本部',
        '瑞华',
        '瑞铭',
        '瑞源',
        '丰俊',
        '服务站',
        '救援中心',
        '耀泓',
        '南泓仓储',
        '南泓物流',
        '瑞丰',
        '恒联',
        '瑞霖',
        '瑞宇',
        '瑞嘉',
        '瑞赫',
        '瑞欧',
        '瑞乾',
        '德嘉',
        '瑞利',
        '德骏',
        '瑞扬',
        '二手车',
        '南瑞'
    ]
    # 车辆类型为了固定顺序而定义
    typeList = [
        "高层配车",
        "工作车",
        "试驾车",
        "救援车",
        "业务车",

    ]




    # 打开这两个表，再创建一个新表
    try:
        app = xw.App(visible=False, add_book=False)

        wb1 = app.books.open(pre_file)
        wb2 = app.books.open(now_file)
        shutil.copyfile(now_file, out_file)
        wb3 = app.books.open(out_file)
        sheet3 = wb3.sheets(1)

        subTabNum = len(wb3.sheets)

        print("该工作簿一共又" + str(subTabNum) + "个工作表")

        # 拿到去重字段的数据
        resArr = getSetData(wb2,keywordList,headNum)
        print(resArr)

        # 创建汇总表
        huizong = wb3.sheets.add("汇总")
        huizong.range(1,1).value = huizongTitle
        huizong.range(2,1).value = "单位"
        danweiDict = {}
        typeDict = {}
        maxRow = 0
        maxCol = 0
        # 生成单位字典{'瑞源'：{'row':2,'col':1},...}
        for nrow in range(0,len(companyList)):
            currentDanwei = companyList[nrow]
            huizong.range(nrow + 3, 1).value = currentDanwei
            danweiDict[currentDanwei] = {}
            danweiDict[currentDanwei]['row'] = nrow + 3
            danweiDict[currentDanwei]['col'] = 1
            maxRow = nrow + 3
        # 生成车辆类型字典{'救援车'：{'row':2,'col':1},...}
        for ncol in range(0,len(typeList)):
            currentType = typeList[ncol]
            huizong.range(2,ncol + 2).value = currentType
            typeDict[currentType] = {}
            typeDict[currentType]['row'] = 2
            typeDict[currentType]['col'] = ncol + 2
            maxCol = ncol + 2

        sumNumRes = {}

        # 根据单位打开每个sheet，获取每个类型车辆的数量
        for danwei in resArr[0]['dataArr']:
            danwei = companyToSimple(danwei)
            currentDanweiSheet = wb3.sheets[danwei]
            xPositon = resArr[1]['xPositon']
            yPositon = resArr[1]['yPositon']
            info = currentDanweiSheet.used_range
            nrows = info.last_cell.row
            ncols = info.last_cell.column
            sumNumRes[danwei] = {}
            for type in resArr[1]['dataArr']:
                sumNumRes[danwei][type] = 0
            for row in range(headNum + 1,nrows + 1):
                currentType = currentDanweiSheet.range(row,yPositon).value
                sumNumRes[danwei][currentType] = sumNumRes[danwei][currentType] + 1

        print(sumNumRes)

        # 遍历单位写入汇总数据
        for key,value in danweiDict.items():
            for k,v in typeDict.items():
                if key not in sumNumRes.keys():
                    continue
                if sumNumRes[key][k] != 0:
                    huizong.range(danweiDict[key]['row'],typeDict[k]['col']).value = sumNumRes[key][k]

        print("现汇总表的行数和列数",maxRow,maxCol)

        huizong.range(2,maxCol + 1).value = "合计"
        huizong.range(maxRow + 1,1).value = "合计"

        for x in range(3, maxRow + 1):
            sum = 0
            for y in range(2,maxCol + 1):
                value = huizong.range(x,y).value
                if value == None:
                    value = 0
                sum = sum + value
            huizong.range(x,maxCol + 1).value = sum

        for y in range(2, maxCol + 2):
            sum = 0
            for x in range(3, maxRow + 1):
                value = huizong.range(x, y).value
                if value == None:
                    value = 0
                sum = sum + value
            huizong.range(maxRow + 1,y).value = sum


        # 读取上个月的表第一列 合计字段所在行
        preHuizong = wb1.sheets["汇总"]
        info1 = preHuizong.used_range
        nrows1 = info1.last_cell.row
        ncols1 = info1.last_cell.column

        preHejiRow = 0
        preHejiCol = 0
        for i in range(headNum + 1,nrows1 + 1):
            value = preHuizong.range(i,1).value
            if value == preHuZongHejiKey:
                preHejiRow = i
                preHejiCol = 1
                break

        print("上个月的合计字段所在行",preHejiRow)

        # 把上个月的合计数据写入到现表的合计
        huizong.range(maxRow + 3,1).value = "上月合计"
        for i in range(2,8):
            huizong.range(maxRow + 3,i).value = preHuizong.range(preHejiRow,i).value

        huizong.range(headNum, maxCol + 3).value = "上月合计"
        for i in range(headNum + 1,maxRow + 1):
            nowValue = huizong.range(i,1).value
            preValueRow = getKeyRow(preHuizong,1,nowValue)
            if preValueRow == None:
                continue
            huizong.range(i,maxCol + 3).value = preHuizong.range(preValueRow,7).value




        print("===============汇总表已完成==================")

        # 获取对比信息
        """
        
        
        
        """







        # wb1.save()
        # wb2.save()
        wb3.save()
        wb1.close()
        wb2.close()
        wb3.close()
        app.quit()
        pass
    except Exception as result:
        print('出现了异常')
        print(result)
        print('str(e):\t', result)
        print('repr(e):\t', repr(result))

        print('traceback.format_exc():\n%s' % traceback.format_exc())  # 字符串
        traceback.print_exc()  # 执行函数
        # wb1.save()
        # wb2.save()
        wb3.save(out_file)
        wb1.close()
        wb2.close()
        wb3.close()
        app.quit()


    pass
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
    sheet1 = wb.sheets(1)
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

# 找出指定关键字所在行
def getKeyRow(sheet,col,key):
    info = sheet.used_range
    nrows = info.last_cell.row
    ncols = info.last_cell.column
    for i in range(1,nrows + 1):
        value = sheet.range(i,col).value
        if value == key:
            return i
    return None

# 拿到指定表的说有sheets名字

if __name__ == '__main__':
    start();