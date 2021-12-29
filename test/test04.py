#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-12-16 13:43
# @Author:xinggevip
# @File : test04.py
# @Software: PyCharm

import pandas as pd
import time
import datetime
import numpy as np



path = "D:\\office\\各单位表十一\\"

# 创建一个空的excel文件
# nan_excle = pd.DataFrame(np.arange(12).reshape((3,4)))
# nan_excle.to_excel(path + filename,sheet_name="sheet2",index=False)

mainpath = 'D:\\office\\各单位表十一\\数据源\\CAP-HR-03 员工档案表_1.xls'  # 全部
data = pd.read_excel(mainpath, '员工档案数据管理', skiprows=2)

data = data[["现职单位","现职部门","姓名"]]

print(data["现职单位"].str[0:-2])

data.dropna(subset=["现职单位"], inplace=True)
listType = data['现职单位'].unique().tolist()
print(listType)

for index,sheet in enumerate(listType):
    # 写入到指定sheet
    table = data[data["现职单位"] == sheet]
    print(table)
    # 将table写入到指定文件夹
    file_name = str(index)  + sheet + '.xlsx'
    # table.to_excel(path + file_name, sheet_name=sheet, index=False)
    pass

