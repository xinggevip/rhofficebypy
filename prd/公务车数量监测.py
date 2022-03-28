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
    pre_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\test\\拆分\\2022年1月瑞华集团公务车信息汇总表.xlsx'
    # 这个月数据
    now_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\test\\拆分\\2022年2月瑞华集团公务车信息汇总表.xlsx'
    # 输出数据
    out_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\test\\拆分\\2022年1月和2022年2月对比.xlsx'
    # 上个月汇总表合计一行的关键字标识
    preHuZongHejiKey = "合计"
    # 原始表的表头行数
    headNum = 2
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
        '南瑞',
        '瑞恒'
    ]
    # 车辆类型为了固定顺序而定义
    typeList = [
        "高层配车",
        "工作车",
        "试乘试驾车",
        "救援车",
        "业务车",

    ]

    # 定义只保留的字段
    delList = [
        "A车牌号",
        "使用单位",
        "品牌",
        "规格型号",
        "A车架号",
        "二级类别",
        "使用单位"
    ]

    # 定义不需要删除多余字段的表名列表
    noDelKeyTabNameList = ["更新车辆","移除车辆","新增车辆","汇总","原始表"]

    # 老资产编号列表
    oldDuibiNumArr = []
    # 新资产编号列表
    newDuibiNumArr = []
    # 资产编号关键字
    zichanKeyWord = "资产编号"
    # 监控变化的字段
    jiankongKeyWordArr = ["使用单位","使用部门","二级类别"]
    # 车辆类别字段
    carTypeKeyWord = "二级类别"
    # 变化检测后缀
    jiankongEndStr = "是否变化"


    # 打开这两个表，再创建一个新表
    try:
        app = xw.App(visible=False, add_book=False)

        wb1 = app.books.open(pre_file)
        wb2 = app.books.open(now_file)
        shutil.copyfile(now_file, out_file)
        wb3 = app.books.open(out_file)
        sheet3 = wb3.sheets(1)

        subTabNum = len(wb3.sheets)

        print("对比结果工作簿一共有" + str(subTabNum) + "个工作表")


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

        huizong = wb3.sheets["汇总"]
        info = huizong.used_range
        maxRow = info.last_cell.row
        maxCol = info.last_cell.column


        # 把上个月的合计数据写入到现表的合计
        huizong.range(maxRow + 1,1).value = "上月合计"
        for i in range(2,8):
            huizong.range(maxRow + 1,i).value = preHuizong.range(preHejiRow,i).value

        huizong.range(headNum, maxCol + 1).value = "上月合计"
        for i in range(headNum + 1,maxRow):
            nowValue = huizong.range(i,1).value
            preValueRow = getKeyRow(preHuizong,1,nowValue)
            if preValueRow == None:
                continue
            huizong.range(i,maxCol + 1).value = preHuizong.range(preValueRow,7).value

        print("===============汇总表已完成==================")

        # 读取旧表
        oldYuanshi = wb1.sheets["原始表"]
        oldInfo = oldYuanshi.used_range
        oldNrows = oldInfo.last_cell.row
        oldNcols = oldInfo.last_cell.column
        # 读取新表
        newYuanshi = wb2.sheets["原始表"]
        newInfo = newYuanshi.used_range
        newNrows = newInfo.last_cell.row
        newNcols = newInfo.last_cell.column

        # 找出两个表指定字段的坐标
        oldZichanKeyWordPos = getKeyPos(oldYuanshi,zichanKeyWord,headNum)
        newZichanKeyWordPos = getKeyPos(newYuanshi,zichanKeyWord,headNum)

        print("旧表y轴",oldZichanKeyWordPos['yPosition'])
        print("新表y轴",newZichanKeyWordPos['yPosition'])

        # 拿到旧表的资产编号列表
        for i in range(headNum + 1,oldNrows + 1):
            value = oldYuanshi.range(i,oldZichanKeyWordPos['yPosition']).value
            oldDuibiNumArr.append(value)
            print("旧表资产编号：",value)

        # 拿到新表的资产编号列表
        for i in range(headNum + 1,newNrows + 1):
            value = newYuanshi.range(i,newZichanKeyWordPos['yPosition']).value
            newDuibiNumArr.append(value)
            print("新表资产编号：", value)
        # 拿到新增资产，处置资产，无变化资产
        addDuibiNumArr,remDuibiNumArr,joinDuibiNumArr = findJoinData(oldDuibiNumArr,newDuibiNumArr)
        # 处置资产带行数
        remZichanPos = getColValuePos(oldYuanshi,headNum,oldNrows,oldZichanKeyWordPos['yPosition'],remDuibiNumArr)
        # 新增资产带行数
        addZichanPos = getColValuePos(newYuanshi,headNum,newNrows,newZichanKeyWordPos['yPosition'],addDuibiNumArr)
        # 无变化资产 旧表中的 带行数
        oldJoinZichanPos = getColValuePos(oldYuanshi,headNum,oldNrows,oldZichanKeyWordPos['yPosition'],joinDuibiNumArr)
        # 无变化资产 新表中的 带行数
        addJoinZichanPos = getColValuePos(newYuanshi,headNum,newNrows,newZichanKeyWordPos['yPosition'],joinDuibiNumArr)

        for key in oldJoinZichanPos:
            print("资产编号",key)
            print("无变化资产旧表编号行数",oldJoinZichanPos[key],"====","新   ",addJoinZichanPos[key])



        # 新增这三张表
        addSheet = wb3.sheets.add("新增车辆")
        remSheet = wb3.sheets.add("移除车辆")
        updateSheet = wb3.sheets.add("更新车辆")

        for i in range(1,headNum + 1):
            newYuanshi.api.Rows(i).Copy(addSheet.api.Rows(i))
            newYuanshi.api.Rows(i).Copy(remSheet.api.Rows(i))
            newYuanshi.api.Rows(i).Copy(updateSheet.api.Rows(i))


        print("遍历写入新增车辆")
        writeRows = headNum + 1
        for i in addDuibiNumArr:
            print("新增车辆", i)
            hang = addZichanPos[i]
            newYuanshi.api.Rows(hang).Copy(addSheet.api.Rows(writeRows))
            writeRows = writeRows + 1

        print("遍历写入移除车辆")
        writeRows = headNum + 1
        for i in remDuibiNumArr:
            print("移除车辆",i)
            hang = remZichanPos[i]
            oldYuanshi.api.Rows(hang).Copy(remSheet.api.Rows(writeRows))
            writeRows = writeRows + 1

        print("遍历写入无变化车辆")
        writeRows = headNum + 1
        for i in joinDuibiNumArr:
            print("无变化车辆", i)
            hang = addJoinZichanPos[i]
            newYuanshi.api.Rows(hang).Copy(updateSheet.api.Rows(writeRows))
            writeRows = writeRows + 1

        # 添加监控字段在后面
        # 获取指定字段的坐标
        # oldJiankongKeyPos = getKeyPosPro(oldYuanshi,jiankongKeyWordArr,headNum)

        updateSheetInfo = updateSheet.used_range
        updateSheetRows = updateSheetInfo.last_cell.row
        updateSheetCols = updateSheetInfo.last_cell.column

        addEndStrCount = 1
        for i in jiankongKeyWordArr:
            updateSheet.range(headNum,updateSheetCols + addEndStrCount).value = i + jiankongEndStr
            addEndStrCount = addEndStrCount + 1

        updateSheet.range(headNum,updateSheetCols + len(jiankongKeyWordArr) + 1).value = "变化详情"

        for i in range(headNum + 1,updateSheetRows + 1):
            # 拿到资产编号
            zichanNum = updateSheet.range(i,newZichanKeyWordPos['yPosition']).value
            print("新表的资产编号在第",i,"行",zichanNum)
            oldZichanRows = oldJoinZichanPos[zichanNum]
            print("旧表的资产编号在第",oldZichanRows,"行")

            oldDict = {}
            newDict = {}

            # 遍历旧资产的列
            for j in range(1,oldNcols + 1):
                value = oldYuanshi.range(oldZichanRows,j).value
                title = oldYuanshi.range(headNum,j).value
                oldDict[title] = value
                # print("旧表的资产编号在第", i, "行", zichanNum, "-----", "title", value)
            # 遍历新资产的列
            for j in range(1,newNcols + 1):
                value = updateSheet.range(i,j).value
                title = updateSheet.range(headNum,j).value
                # print("新表的资产编号在第", i, "行", zichanNum, "-----", "title", value)
                newDict[title] = value

            # 初始化字段
            strRes = ""
            allEq = True
            for index,jiankongKey in enumerate(jiankongKeyWordArr):

                print("旧",oldDict)
                print("新",newDict)



                if oldDict[jiankongKey] == None:
                    oldDict[jiankongKey] = ""
                if newDict[jiankongKey] == None:
                    newDict[jiankongKey] = ""


                if oldDict[jiankongKey] != newDict[jiankongKey]:
                    updateSheet.range(i,updateSheetCols + index + 1).value = "是"
                    strRes = strRes + jiankongKey +'：原值为"' + oldDict[jiankongKey] + '"，现值为"' + newDict[jiankongKey] + ';\n'
                    allEq = False
                else:
                    updateSheet.range(i, updateSheetCols + index + 1).value = "否"
                    pass

            if allEq:
                updateSheet.range(i, updateSheetCols + len(jiankongKeyWordArr) + 1).value = "无变化"
                pass
            else:
                updateSheet.range(i, updateSheetCols + len(jiankongKeyWordArr) + 1).value = strRes

        # 删除多余字段


        # 获取对比信息
        """
        
        
        
        """





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
        "新密市瑞利汽车销售有限公司",
        "郑州智领瑞华汽车销售服务有限公司"
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
        if company == "郑州智领瑞华汽车销售服务有限公司":
            return "智领瑞华"
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

def getKeyPos(sheet,keyword,headNum):
    info = sheet.used_range
    ncols = info.last_cell.column
    res = {}
    for i in range(1, headNum + 1):  # 遍历前两行
        for y in range(1, ncols + 1):  # 遍历最长列数
            temp = sheet.range(i, y).value
            if temp != None:
                # print(str(i) + "     " + str(y) + "     =     " + temp)
                if temp == keyword:
                    print(str(i) + "     " + str(y) + "     =     " + temp)
                    xPositon = i
                    yPositon = y
                    res["xPosition"] = xPositon
                    res["yPosition"] = yPositon

    return res
    pass

def getKeyPosPro(sheet,keywordList,headNum):
    info = sheet.used_range
    ncols = info.last_cell.column
    res = {}
    for key in keywordList:
        res[key] = {}
        for i in range(1, headNum + 1):  # 遍历前两行
            for y in range(1, ncols + 1):  # 遍历最长列数
                temp = sheet.range(i, y).value
                if temp != None:
                    # print(str(i) + "     " + str(y) + "     =     " + temp)
                    if temp == key:
                        print(str(i) + "     " + str(y) + "     =     " + temp)
                        xPositon = i
                        yPositon = y
                        res[key]["xPosition"] = xPositon
                        res[key]["yPosition"] = yPositon
        pass


    return res
    pass

# 找出新增数据，移除数据，共同数据
def findJoinData(list1,list2):
    joinList = []
    addList = []
    remList = []

    for i in list1:
        for j in list2:
            if i == j:
                joinList.append(i)
                print(i,"-------俩表都在，无变化")

    for b in (list1 + list2):
        if b not in joinList:
            if b not in list2:
                remList.append(b)
                print(b, "-------不在新表，已处置")
            if b not in list1:
                print(b, "-------不在旧表，新资产")
                addList.append(b)

    return addList,remList,joinList
    pass

# 传入表，表头行数，表最大行数，数据列数，数据  返回dataArr数的行数{RH20200529:1,...}
def getColValuePos(sheet,headNum,maxRow,nCol,dataArr):

    res = {}
    for i in range(headNum + 1,maxRow + 1):
        value = sheet.range(i,nCol).value
        if value in dataArr:
           res[value] = i

    return res
    pass

# 删除多余字段
def delCol(sheet,headNum,dataArr):
    """
    1. 获取sheet的列数
    2. 取表头的值 遍历判断是否在dataArr中，在则删除
    """
    sheetInfo = sheet.used_range
    sheetNrows = sheetInfo.last_cell.row
    sheetNcols = sheetInfo.last_cell.column
    sheetName = sheet.name
    for i in range(1,sheetNcols + 1):
        value = sheetInfo.range(headNum,i).value
        if value in dataArr:
            # 删除列
            print("删除",sheetName,"表中的",value,"列")
            sheet.api.Columns(i).Delete()
            pass
        pass
    pass

# 添加样式

# 添加序号列
def addNumCols(sheet,headNum):

    sheetInfo = sheet.used_range
    sheetNrows = sheetInfo.last_cell.row
    sheetNcols = sheetInfo.last_cell.column
    # 在第一列前插入插入一列
    sheet.api.Rows(1).Insert()
    sheet.range(headNum,1).value = "序号"
    count = 1
    for i in range(headNum + 1, sheetNrows + 1):
        sheet.range(i,1).value = count
        count = count + 1
    pass




# 拿到指定表的说有sheets名字

if __name__ == '__main__':
    start();