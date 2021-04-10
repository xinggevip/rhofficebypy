#!/usr/bin/env python3
# -*- coding:utf-8 -*-
# @Time :2020-12-1 10:07
# @Author:xinggevip
# @File : dataUtils.py
# @Software: PyCharm

#!/usr/bin/python
#coding=UTF-8
import datetime

def getday(y=2017,m=8,d=15,n=0):
    the_date = datetime.datetime(y,m,d)
    result_date = the_date + datetime.timedelta(days=n)
    d = result_date.strftime('%Y-%m-%d')
    return d


# print(getday(2017,8,15,21)) #8月15日后21天
print(getday(2020,12,1,-30)) #9月1日前10天