# -*- coding: utf-8 -*-
"""
Created on Fri Jan 11 14:51:21 2019

@author: 杨镇宁
"""
import pandas as pd
import scipy.interpolate as spi

inputfile='C:/Users/杨镇宁/Desktop/work.xlsx'

#插值过程，采用样条插值，可有效避免拉格朗日插值出现的龙格现象。
missingData = pd.read_excel(inputfile,sheetname=1,header=None)   #将excel中数据读到dataframe
row=list(missingData.index)
row=row[2:len(row)]    #从第二行开始寻找缺省值
newrow=[]
newcol=[]
for i in missingData.columns:
    for j in row:
        if missingData[i].isnull()[j]:            #得保证离上下界最近的空值<50
            x = missingData[i][list(range(j - 50, j)) + list(range(j + 1, j + 51))]#前后各取50个
            x = x[x.notnull()]                    #剔除选取前后数据中空值
            ipo3=spi.splrep(x.index,list(x),k=1)  #横坐标数据必须按大小顺序排列，否则报错
            missingData[i][j] =spi.splev(j,ipo3)  #试了几种方法，发现k=1为均值插值，比较适合电力数据
            newrow +=[j]
            newcol +=[i]
missingData.to_excel('插值后数据1小时为单位.xls', header=None, index=False)

#修改插值后单元格颜色过程
import xlwt
import xlrd
from xlutils.copy import copy
file_name = '插值后数据1小时为单位.xls' #与上面一致
n=len(newrow)
styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour red;')  #红色
rb = xlrd.open_workbook(file_name)      #打开t.xls文件
ro = rb.sheets()[0]                     #读取表单0
wb = copy(rb)                           #利用xlutils.copy下的copy函数复制
ws = wb.get_sheet(0)                    #获取表单0
for i in range(n):                      #循环所有的行
    ws.write(newrow[i],newcol[i],ro.cell(newrow[i],newcol[i]).value,styleBlueBkg)#ro.cell(i, col)是一个dict
wb.save(file_name)
