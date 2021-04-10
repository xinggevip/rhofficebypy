#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-4-5 15:02
# @Author:xinggevip
# @File : chaoqijun.py
# @Software: PyCharm

"""
1. 手动输入需要平均的月份
2. 把复制一份模板到输出文件夹
3. 遍历文件夹把数据存到数组
4. 打开处理后文件并写入数据
5. 打开输出文件夹
"""

import shutil
import os
import xlwings as xw
import re
import subprocess

app = xw.App(visible= False, add_book= False)

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

def to_numarr(str_arr):
    for i in range(0,len(str_arr)):
        str_arr[i] = int(str_arr[i])
    pass

def start():
    global wb2, wb1
    try:
        # 定义存放处理后数据的目录
        chulihou_office_path = "F:\\newoffice\\"
        # 模板所在位置
        muban_file_path = "F:\\每个季度超期平均值\\模板\\超期明细平均值.xlsx"
        # 处理后打开位置
        open_file_path = "F:\\每个季度超期平均值\\处理后"

        names, files = getRawFileList(chulihou_office_path)

        print(names)

        months_text = input("请输入月份数字，用英文逗号隔开: ")
        # 文件前缀
        pre_file_name = input("请输入文件前缀：")
        month_arr = months_text.split(",")

        to_numarr(month_arr)

        pattern = re.compile(r'.*?\d{4}(\d{2}).*(?<!15)$')

        i_arr = []
        for i in range(0, len(names)):
            matchObj = pattern.match(names[i])

            if matchObj:
                # print(i + 1,"  matchObj.group() : ",matchObj.group(1))
                res = matchObj.group(1)
                print('{}匹配到了结果为: {}'.format(names[i], res))
                if int(res) in month_arr:
                    i_arr.append(i)
                    print('结果在输入数组中')
            else:
                print(names[i] + "   No match!!")
            pass

        # 复制一份模板到
        newfile = os.path.join(open_file_path, pre_file_name + "超期明细平均值.xlsx")
        shutil.copyfile(muban_file_path, newfile)

        wb1 = app.books.open(newfile)
        sheet1 = wb1.sheets("汇总")

        # 拿到处理后的文件路径
        yichuli_arr = []
        yichaoqi_arr = []
        for i in i_arr:
            chulihou_path = os.path.join(names[i], '处理后数据\\')
            aarr, barr = getRawFileList(chulihou_path)
            # 拿到每个文件的全路径了
            print(aarr[0])
            wb2 = app.books.open(aarr[0])
            sheet2 = wb2.sheets("汇总")
            yichuli_arr_sub = []
            yichaoqi_arr_sub = []
            for row in range(2, 22):
                danwei = sheet2.range(row, 1).value
                yichuli = sheet2.range(row, 2).value
                yichaoqi = sheet2.range(row, 3).value

                if yichuli == None or yichuli == '':
                    yichuli = 0

                if yichaoqi == None or yichaoqi == '':
                    yichaoqi = 0

                yichuli_arr_sub.append(yichuli)
                yichaoqi_arr_sub.append(yichaoqi)
                print(danwei, yichuli, yichaoqi)

            print("*" * 50)
            yichuli_arr.append(yichuli_arr_sub)
            yichaoqi_arr.append(yichaoqi_arr_sub)

        pass


        for row in range(2, 22):
            sum1 = 0
            sum2 = 0
            for i in range(0, len(yichuli_arr)):
                sum1 = sum1 +  yichuli_arr[i][row - 2]
                sum2 = sum2 + yichaoqi_arr[i][row - 2]
            sheet1.range(row, 2).value = sum1
            sheet1.range(row, 3).value = sum2

            print("sum1 = {} ,   sum2 = {}".format(sum1,sum2))

            pass





        wb1.save()
        wb2.save()
        wb1.app.quit()

        subprocess.Popen('explorer ' + open_file_path)
    except Exception as result:
        print(result)
        wb1.save()
        wb2.save()
        wb1.app.quit()
        pass
    pass




start()