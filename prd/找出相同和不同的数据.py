#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-3-30 9:40
# @Author:xinggevip
# @File : dangan.py
# @Software: PyCharm

"""
把cap3格式的档案转换成cap4的档案
1. 复制一份模板到输出目录
2. 打开待处理表和模板表
end.打开输出目录
"""

list1 = [
"RH20190213541",
"RH20190213186",
"RH20190213075",
"RH20190213034",
"RH20190213132",
"RH20190213184",
"RH20190213852",
"RH20190213620",
"RH20190213135",
"RH20190213303",
"RH20190213851",
"RH20190213161",
"RH20190213099",
"RH20201203011",
"RH20190213548",
"RH20190213544",
"RH20190213618",
"RH20190213547",
"RH20190213302",
"RH20190213134",
"RH20190213035",
"RH20190213543",
"RH20171104109",
"RH20190213619",
"RH20190213300",
"RH20171104110",
"RH20190213422",
"RH20190213101",
"RH20190213133",
"RH20190213853",
"RH20190213545",
"RH20190213074",
"RH20190213299",
"RH20190213162",
"RH20190213423",
"RH20190213297",
"RH20190213185",
"RH20190213304",
"RH20190213578",
"RH20190213301",
"RH20190213073",
"RH20190213546",
"RH20190319055",
"RH20190213298",
"RH20190213580",

]
list2 = [
"RH20201203011",
"RH20171104110",
"RH20171104109",
"RH20190213035",
"RH20190213034",
"RH20190213075",
"RH20190213074",
"RH20190213099",
"RH20190213101",
"RH20190213133",
"RH20190213135",
"RH20190213132",
"RH20190213134",
"RH20190213162",
"RH20190213186",
"RH20190213161",
"RH20190213184",
"RH20190213303",
"RH20190213300",
"RH20190213298",
"RH20190213301",
"RH20190213304",
"RH20190213297",
"RH20190213299",
"RH20190213423",
"RH20190213422",
"RH20190213548",
"RH20190213541",
"RH20190213580",
"RH20190213545",
"RH20190213543",
"RH20190213578",
"RH20190213546",
"RH20190213547",
"RH20190213544",
"RH20190213618",
"RH20190213620",
"RH20190213619",
"RH20190213853",
"RH20190213851",
"RH20190213852",
"RH20190319055",
"RH20190213302",
"RH20190213073",
"RH20190213185",
"RH20190213579",
]
list3 = []
list4 = []

for i in list1:
    for j in list2:
        if i == j:
            list3.append(i)

for b in (list1 + list2):
    if b not in list3:
        if b not in list2:
            print(b,"不在资产管理")
        if b not in list1:
            print(b,"不在360管理")
        list4.append(b)

print(list3)
print(list4)

