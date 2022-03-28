#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2022-03-03 9:35
# @Author:xinggevip
# @File : 固定资产监测.py
# @Software: PyCharm

import pandas as pd
import numpy as np
import time
import datetime

def start():
    # 数据源路径
    old_table_path = 'D:\office\固定资产监控\数据源\资产数据管理1.xlsx'
    new_table_path = 'D:\office\固定资产监控\数据源\资产数据管理2.xlsx'

    # 表头行数
    head_num = 2

    # 监控字段
    check_key_arr = ['资产状态','使用单位','存放位置']

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
    print("读取完毕，用时",end_time - start_time,"秒")

    old_data['资产编号'].astype(str)
    new_data['资产编号'].astype(str)

    old_data = old_data[old_data["一级类别"] == "车辆资产"]
    new_data = new_data[new_data["一级类别"] == "车辆资产"]

    # 拿到两张表的资产编号数组
    old_zichan_index_arr = old_data['资产编号'].tolist()
    new_zichan_index_arr = new_data['资产编号'].tolist()

    # 获取新增，处置，无变化资产编号数组
    print("2.正在获取新增、处置、无变化资产...")
    start_time = time.time()
    new_zichan_index_arr,del_zichan_index_arr,normal_zichan_index_arr = get_three_arr_with_yidong(old_zichan_index_arr,new_zichan_index_arr)

    end_time = time.time()
    print("获取完毕，用时", end_time - start_time, "秒")
    print('new_zichan_index_arr',new_zichan_index_arr)
    print('del_zichan_index_arr',del_zichan_index_arr)
    print('normal_zichan_index_arr',normal_zichan_index_arr)

    new_zichan_pd = new_data[new_data['资产编号'].isin(new_zichan_index_arr)].copy()
    new_zichan_pd['资产变更'] = '资产新增'
    print("新增资产=============================")
    print(new_zichan_pd)

    del_zichan_pd = old_data[old_data['资产编号'].isin(del_zichan_index_arr)].copy()
    del_zichan_pd['资产变更'] = '资产处置'
    print("处置资产=============================")
    print(del_zichan_pd)

    normal_zichan_pd = new_data[new_data['资产编号'].isin(normal_zichan_index_arr)].copy()
    normal_zichan_pd['资产变更'] = ''
    print("无变化资产=============================")
    print(normal_zichan_pd)

    # 根据keyarr监控字段变化
    print("3.监控字段变化...")
    start_time = time.time()
    jiankong_by_key_arr(old_data, normal_zichan_pd, check_key_arr)
    end_time = time.time()
    print("监控字段处理完毕，用时", end_time - start_time, "秒")

    # 使用单位发生变化 则复制一行为调出
    print("4.生成调出行...")
    start_time = time.time()
    normal_zichan_pd = diaochu_danwei_copy(old_data,normal_zichan_pd)
    print("调出行生成完毕，用时", end_time - start_time, "秒")

    biangeng(normal_zichan_pd,check_key_arr)

    normal_zichan_pd = normal_zichan_pd.append(new_zichan_pd, ignore_index=True)
    normal_zichan_pd = normal_zichan_pd.append(del_zichan_pd, ignore_index=True)

    col_list = ['资产变更','资产编号','资产名称','资产状态','品牌','规格型号','供货单位','购置日期','保修期至','使用单位','使用部门','存放位置']
    for key in check_key_arr:
        key = key + '是否变化'
        col_list.append(key)
    col_list.append('变化详情')
    normal_zichan_pd = normal_zichan_pd[col_list]

    normal_zichan_pd.fillna('', inplace=True)
    normal_zichan_pd = normal_zichan_pd.replace(np.nan, '', regex=True)
    normal_zichan_pd.dropna(subset=['资产变更'], inplace=True)

    normal_zichan_pd.to_excel('D:\\office\\固定资产监控\输出目录\\' + '全集团固定资产监控.xlsx', sheet_name="sheet1", index=False)
    print("运行结束")



def path_deal(str):
    return str.replace("\\","\\\\")

def get_three_arr_with_yidong(list1,list2):
    # 新增资产
    new_zichan_index_arr = []
    # 处置资产
    del_zichan_index_arr = []
    # 无变化资产
    normal_zichan_index_arr = []

    for i in list1:
        for j in list2:
            if i == j:
                normal_zichan_index_arr.append(i)

    for b in (list1 + list2):
        if b not in normal_zichan_index_arr:
            if b not in list2:
                del_zichan_index_arr.append(b)
            if b not in list1:
                new_zichan_index_arr.append(b)

    return new_zichan_index_arr,del_zichan_index_arr,normal_zichan_index_arr

def jiankong_by_key_arr(old_data,normal_zichan_pd,check_key_arr):
    for key in check_key_arr:
        key = key + '是否变化'
        normal_zichan_pd[key] = '否'
    normal_zichan_pd['变化详情'] = ''

    for index, row in normal_zichan_pd.iterrows():
        bianhao = row["资产编号"]
        index2 = old_data[old_data["资产编号"] == bianhao].index.tolist()[0]

        print('normal_zichan_pd 的资产编号 == ', bianhao)
        allEq = True
        for key in check_key_arr:
            # print('normal_zichan_pd.loc[index,key] === ',normal_zichan_pd.loc[index,key])
            # print('old_data.loc[index2,key] === ',old_data.loc[index2,key])
            flag_key = key + '是否变化'

            if str(normal_zichan_pd.loc[index,key]) != str(old_data.loc[index2,key]):
                allEq = False
                normal_zichan_pd.loc[index,flag_key] = '是'
                normal_zichan_pd.loc[index,'变化详情'] = normal_zichan_pd.loc[index,'变化详情'] + '"' + key + '"' + '原值为：' + str(old_data.loc[index2,key]) + ',现值为：' + str(normal_zichan_pd.loc[index,key]) + ';\n'
        if allEq == True:
            normal_zichan_pd.loc[index, '变化详情'] = '无变化'

    pass

def diaochu_danwei_copy(old_data,normal_zichan_pd):
    need_add_pd = pd.DataFrame()
    for index, row in normal_zichan_pd.iterrows():
        bianhao = row["资产编号"]

        bian_danwei = row['使用单位是否变化']

        if str(bian_danwei) == "是":
            normal_zichan_pd.loc[index,'资产变更'] = '资产调入'
            for index2, row2 in old_data.iterrows():
                if row2["资产编号"] == bianhao:
                    row2["资产变更"] = '资产调出'
                    need_add_pd = need_add_pd.append(row2, ignore_index=True)
                    print('need_add_pd',need_add_pd)

    # normal_zichan_pd = normal_zichan_pd.append(need_add_pd, ignore_index=True)
    normal_zichan_pd = pd.concat([normal_zichan_pd,need_add_pd], ignore_index=True)
    print(normal_zichan_pd)
    return normal_zichan_pd
def biangeng(normal_zichan_pd,check_key_arr):
    for index, row in normal_zichan_pd.iterrows():

        for key in check_key_arr:
            flag_key = key + '是否变化'
            if str(row["资产变更"]) == '':
                if str(row[flag_key]) ==  '是':
                    if flag_key == "资产状态是否变化":
                        normal_zichan_pd.loc[index,"资产变更"] = '状态调整'
                    elif flag_key == '存放位置是否变化':
                        normal_zichan_pd.loc[index,"资产变更"] = '位置转移'


    pass



if __name__ == '__main__':
    start()