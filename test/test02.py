import shutil
import os
import xlwings as xw
import re

app = xw.App(visible= False, add_book= False)

print('hello')

# 测试临时路径
path1 = 'D:\\office\\待处理数据\\1集团本部10月流程统计.xls'
path2 = 'D:\\office\\处理后数据\\202010超期明细.xlsx'

wb1 = app.books.open(path1)
wb2 = app.books.open(path2)



ws1 = wb1.sheets(1)
print(ws1.name)
ws1.api.Copy(After=wb2.sheets(1).api)
wb2.sheets(1).name = "测试"
print(wb2.sheets(1).name)

wb1.save()
wb1.app.quit()

wb2.save()
wb2.app.quit()
