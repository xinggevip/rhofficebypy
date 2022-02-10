#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-3-30 9:40
# @Author:xinggevip
# @File : dangan.py
# @Software: PyCharm

"""
找出相同和不同的数据
"""

str1 = "调出"
str2 = "调入"

list1 = [
"步如飞",
"常昊",
"张林霞",
"苏凯旋",
"郭正建",
"祖婧",
"刘召辉"
]

list2 = [
"祖婧",
"梁宇",
"王梦科",
"常昊",
"赵敏",
"步如飞",
"刘召辉"
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
            print(b,"-------",str1,"有，",str2,"没有")
        if b not in list1:
            print(b,"-------",str2,"有，",str1,"没有")
        list4.append(b)

print(list3)
print(list4)

