#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-5-17 10:28
# @Author:xinggevip
# @File : 批量改子表字段名称.py
# @Software: PyCharm

import shutil
import os
import traceback

import xlwings as xw
import re
import subprocess


def start():
    pre_file = 'E:\\01--高星--\\01 工作文档\\18车辆数量监测\\work\\汇总\\2021年3月份瑞华集团公务车信息汇总表.xlsx'
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
        '南瑞'
    ]



    pass



if __name__ == '__main__':
    start()