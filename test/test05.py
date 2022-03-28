#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2022-02-28 9:14
# @Author:xinggevip
# @File : test05.py
# @Software: PyCharm

import pandas as pd
import time
import datetime

def start():
    path = 'D:\\office\\各单位表十一\\数据源\\CAP-HR-03 员工档案表_1.xls'
    out_path = "D:\\office\\各单位表十一\\"

    data = pd.read_excel(path, '员工档案数据管理', skiprows=2)
    data.sort_values('现职单位', inplace=True)
    data.to_excel(out_path + "全集团人事月报测试.xlsx", sheet_name="sheet1", index=False)

    pass

if __name__ == '__main__':
    start()