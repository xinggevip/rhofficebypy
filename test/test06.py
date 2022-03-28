#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2022-03-04 15:13
# @Author:xinggevip
# @File : test06.py
# @Software: PyCharm

import pandas as pd
import time
import datetime


def path_deal(str):
    return str.replace("\\","\\\\")


# 数据源路径
old_table_path = 'D:\office\固定资产监控\数据源\资产数据管理1月.xlsx'
new_table_path = 'D:\office\固定资产监控\数据源\资产数据管理2月.xlsx'

# 表头行数
head_num = 2

# 监控字段
check_key_arr = ['资产状态', '使用单位', '存放位置']

# 读取的sheet名称
sheet_name = '固定资产信息表'

# 路径处理
old_table_path = path_deal(old_table_path)
new_table_path = path_deal(new_table_path)

# 读取excel
print("1.正在读取表格...")
start_time = time.time()
old_data = pd.read_excel(old_table_path, sheet_name, skiprows=head_num - 1)
new_data = pd.read_excel(new_table_path, sheet_name, skiprows=head_num - 1)
end_time = time.time()
print("读取完毕，用时", end_time - start_time, "秒")


for index, row in old_data.iterrows():
    old_data.append(row, ignore_index=True)

print(old_data)

