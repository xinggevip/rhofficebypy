#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2021-3-6 13:59
# @Author:xinggevip
# @File : test03.py
# @Software: PyCharm

import re

# for i in range(10, 1 ,-1):
#     print(i)

def start():

    str = '德骏行政人事部-刘敏-物资采购申请-预计41020.00元-河南德骏2021年销售工装订购-2021-03-01-河南德骏汽车销售有限公司'

    matchObj = re.match(r'.*-(.*)', str, re.M | re.I)

    if matchObj:
        print("  matchObj.group() : ",matchObj.group(1))
        sheetName = matchObj.group(1)
    else:
        print (str + "No match!!")

if __name__ == '__main__':
    start()