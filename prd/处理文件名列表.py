#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-3-3 14:31
# @Author:xinggevip
# @File : baoxiao.py
# @Software: PyCharm

import shutil
import os
import re
import subprocess

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
        "郑州智领瑞华汽车销售服务有限公司",
    ]

    # 直接return
    if company in goReturnList:
        return company
        pass
    # 无规则return
    if company in goReturnOfNoRuleList:
        if "瑞华机动车登记服务站" in company:
            return "服务站"
        if "河南瑞铭二手车" in company:
            return "二手车"
        if "河南耀泓汽车配件销售有限公司" in company:
            return "耀泓"
        if "河南南泓仓储物流有限公司" in company:
            return "南泓仓储"
        if "郑州南瑞汽车配件销售有限公司" in company:
            return "南瑞"
        if "河南南泓汽车贸易有限公司" in company:
            return "南泓汽贸"
        if "巩义市德嘉汽车销售服务有限公司" in company:
            return "德嘉"
        if "新密市瑞利汽车销售有限公司" in company:
            return "瑞利"
        if "郑州智领瑞华汽车销售服务有限公司" in company:
            return "智领瑞华"
    # 有规则return
    matchObj = re.match(r'.*河南(.*?)汽车', company, re.M | re.I)

    if matchObj:
        return matchObj.group(1)
    else:
        print(company)
        print(company + " No match!!")
        return company
        
def detial_list_for_name(list1):
    res = []
    for item in list1:
        str = item.replace('超30天未完结事项需反馈.xls','')
        str = str.replace('超30天未报销事项查询_需反馈.xlsx','')
        res.append(str)
    return res        
        
def start(path):
    print(path)
    files, names = getRawFileList(path)
    # print(files)
    # print(names)
    names.sort()
    res = detial_list_for_name(names)
    print(res)
    
    for i in res:
        print(companyToSimple(i))
        
    print('=' * 30)    
    
    



if __name__ == '__main__':
    path1 = 'E:\\01--高星--\\01 工作文档\\03 超30天督办事项\\20220302\\待处理'
    path2 = 'E:\\01--高星--\\01 工作文档\\03 超30天未报销事项\\20220302\\待处理'
    
    start(path1)
    start(path2)


