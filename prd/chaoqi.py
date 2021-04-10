#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time: 2020-11-15 17:21
# @Author: gaoxing
# @Email: 1511844263@qq.com
# @File: chaoqi.py
import shutil
import os
import xlwings as xw
import re
import subprocess

"""
文件自动汇总、计算已办和超期数据
"""

app = xw.App(visible= False, add_book= False)
# app.display_alerts=False
# app.screen_updating=False

oldfile = 'D:\\office\\模板数据\\202010超期明细.xlsx'   # 模板文件路径
newfile = 'D:\\office\\处理后数据\\202010超期明细.xlsx' # 处理后文件路径
out_file_path = 'D:\\office\\处理后数据\\'             #处理后文件夹路径

data_dir = "D:\\office\\待处理数据\\"          # 待处理数据文件路径

"""获取文件列表"""
def getRawFileList(path):
    """-------------------------
    files,names=getRawFileList(raw_data_dir)
    files: ['datacn/dialog/one.txt', 'datacn/dialog/two.txt']
    names: ['one.txt', 'two.txt']
    ----------------------------"""
    files = []
    names = []
    for f in os.listdir(path):
        if not f.endswith("~") or not f == "":  # 返回指定的文件夹包含的文件或文件夹的名字的列表
            files.append(os.path.join(path, f))  # 把目录和文件名合成一个路径
            names.append(f)
    return files, names


def start():
    global wb2, currSheetTitleAll
    try:
        # global wb2
        files, names = getRawFileList(data_dir)
        print("files:", files)
        print("names:", names)

        # 二手车
        escYiBan = 0
        escChaoQi = 0

        # 服务站
        fwzYiBan = 0
        fwzChaoQi = 0

        # 瑞丰
        ruiFengYiBan = 0
        ruiFengChaoQi = 0

        # 瑞霖
        ruiLinYiBan = 0
        ruiLinChaoqi = 0

        # 刘珂
        liuKeYiBan = 0
        liuKeChaoQi = 0

        # 李恒鹏
        liHengPengYiBan = 0
        liHengPengChaoQi = 0


        # 1.复制模板文件
        shutil.copyfile(oldfile, newfile)  # oldfile只能是文件夹，newfile可以是文件，也可以是目标目录
        print("1.复制好了================")

        # 2.遍历文件插入到汇总
        for i in range(0, len(files)):

            sheetName = names[i]
            line = names[i]

            matchObj = re.match(r'\d*(.*?)\d+', line, re.M | re.I)

            if matchObj:
                # print(i + 1,"  matchObj.group() : ",matchObj.group(1))
                sheetName = matchObj.group(1)
            else:
                print (names[i] + "No match!!")

            wb1 = app.books.open(files[i])
            wb2 = app.books.open(newfile)

            ws1 = wb1.sheets(1)
            # ws1.api.Copy(After=wb2.sheets(len(wb2.sheets)).api)
            # wb2.sheets(len(wb2.sheets)).name = sheetName

            wb2Len = len(wb2.sheets)
            print(wb2Len)

            ws1.api.Copy(Before=wb2.sheets(1).api)
            wb2.sheets(1).name = sheetName
            # print(wb2.sheets(len(wb2.sheets)).range('A2').value)


            # currentSheet = wb2.sheets(len(wb2.sheets))
            currentSheet = wb2.sheets(1)
            info = currentSheet.used_range
            nrows = info.last_cell.row  # 行数

            yiban, chaoqi = '',''


            if sheetName.find('刘珂') != -1:
                currentSheet.range('E' + str(nrows + 2)).formula = '=SUM(E3:' + 'E' + str(nrows + 1) + ')'
                currentSheet.range('G' + str(nrows + 2)).formula = '=SUM(G3:' + 'G' + str(nrows + 1) + ')'

                yiban = currentSheet.range('E' + str(nrows + 2)).value
                chaoqi = currentSheet.range('G' + str(nrows + 2)).value

                liuKeYiBan = yiban
                liuKeChaoQi = chaoqi
            elif sheetName.find('李恒鹏') != -1:
                currentSheet.range('E' + str(nrows + 2)).formula = '=SUM(E3:' + 'E' + str(nrows + 1) + ')'
                currentSheet.range('G' + str(nrows + 2)).formula = '=SUM(G3:' + 'G' + str(nrows + 1) + ')'

                yiban = currentSheet.range('E' + str(nrows + 2)).value
                chaoqi = currentSheet.range('G' + str(nrows + 2)).value

                liHengPengYiBan = yiban
                liHengPengChaoQi = chaoqi
            else:
                currentSheet.range('D' + str(nrows + 2)).formula = '=SUM(D3:' + 'D' + str(nrows + 1) + ')'
                currentSheet.range('F' + str(nrows + 2)).formula = '=SUM(F3:' + 'F' + str(nrows + 1) + ')'

                yiban = currentSheet.range('D' + str(nrows + 2)).value
                chaoqi = currentSheet.range('F' + str(nrows + 2)).value
                if sheetName.find('服务站') != -1:
                    fwzYiBan = yiban
                    fwzChaoQi = chaoqi
                elif sheetName.find('二手车') != -1:
                    escYiBan = yiban
                    escChaoQi = chaoqi
                elif sheetName.find('瑞丰') != -1:
                    ruiFengYiBan = yiban
                    ruiFengChaoQi = chaoqi
                elif sheetName.find('瑞霖') != -1:
                    # print(wb2.sheets(len(wb2.sheets)).range('A2').value)
                    ruiLinYiBan = yiban
                    ruiLinChaoqi = chaoqi


            print(sheetName, "   \t\t\t\t： 已办数量 = ",yiban, "   \t\t\t\t超期数量 = ", chaoqi)

            wb2.save()
            wb2.app.quit()


        print("2.汇总好了===============")



        wb2 = app.books.open(newfile)
        for i in range(0, len(wb2.sheets)):
            currSheetTitle = (wb2.sheets[i].name)[0:2]
            # print(currSheetTitle)
            # print(currSheet.range('A' + str(i + 1)).value)
            pass

        # huizongSheet = wb2.sheets(1)
        huizongSheet = wb2.sheets(len(wb2.sheets))
        info = huizongSheet.used_range
        nrows = info.last_cell.row

        print("------------")
        for i in range(0, nrows):
            # print(huizongSheet.range('A' + str(i + 1)).value)
            mystr = huizongSheet.range('A' + str(i + 1)).value

            isFind = False
            currSheetTitle = ''
            for j in range(0, len(wb2.sheets)):
                currSheetTitleAll = wb2.sheets[j].name
                currSheetTitle = (wb2.sheets[j].name)[0:2]
                # print(currSheetTitle)
                # print(currSheet.range('A' + str(i + 1)).value)
                if mystr.find(currSheetTitle) != -1:
                    isFind = True
                    break
                pass

            pass

            if isFind:

                # print('匹配到了')
                # print(currSheetTitle, "$$$$$$$$$$$$$$")
                print(mystr)

                # 1.获取行数
                sheet = wb2.sheets[currSheetTitleAll]
                info = sheet.used_range
                nrows = info.last_cell.row
                # 2.获取值
                yiban = sheet.range('D' + str(nrows)).value
                chaoqi = sheet.range('F' + str(nrows)).value

                # 3.赋值
                huizongSheet.range('B' + str(i + 1)).value = yiban
                huizongSheet.range('C' + str(i + 1)).value = chaoqi


            else:
                # print('没有匹配===========')
                # print(currSheetTitle,"***********")
                pass
            if mystr.find('服务站') != -1:
                huizongSheet.range('B' + str(i + 1)).value = escYiBan + fwzYiBan
                huizongSheet.range('C' + str(i + 1)).value = escChaoQi + fwzChaoQi
            # elif mystr.find('瑞丰') != -1:
            #     huizongSheet.range('B' + str(i + 1)).value = ruiFengYiBan - liHengPengYiBan
            #     huizongSheet.range('C' + str(i + 1)).value = ruiFengChaoQi - liHengPengChaoQi
            #     pass
            # elif mystr.find('瑞霖') != -1:
            #     huizongSheet.range('B' + str(i + 1)).value = ruiLinYiBan - liuKeYiBan
            #     huizongSheet.range('C' + str(i + 1)).value = ruiLinChaoqi - liuKeChaoQi
            #     pass


            # print("=================================")
        # for j in range(0, len(wb2.sheets)):
        #     currSheetTitle = (wb2.sheets[j].name)[0:2]
        #     if currSheetTitle == '德骏':
        #         print("找到了")
        #     print(currSheetTitle)

        pass



        wb2.save()
        wb2.app.quit()

        print("3. 计算好了")

        subprocess.Popen('explorer ' + out_file_path)

    except Exception as result:
        wb2.save()
        wb2.app.quit()
        print("出现了异常=========================")
        print(result)


start()
