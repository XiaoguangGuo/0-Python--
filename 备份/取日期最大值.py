# -*- coding: utf-8 -*-
import pandas as pd
import os

testdf =pd.read_excel(r'D:\\运营\\计划数据\\老站\\销售数据\\周销售数据test1.xlsx')
testdf2=pd.read_excel(r'D:\\运营\\计划数据\\老站\\销售数据\\周销售数据testtemp.xlsx')
#testdf2['当前']=""
#testdf2['测试周数']=""
a=max(testdf["日期"])
b=testdf['日期'].max()

#注意周数使用单引号
# ok: testdf['周数']=testdf["日期"]
testdf
print(a)
print(b)
#print(testdf['周数'])

testdf['日期'] = pd.to_datetime(testdf['日期'])
# testdf['当前']= pd.to_datetime(testdf["日期"].max())
maxtime=pd.to_datetime(testdf["日期"].max())
#testdf['当前']=pd.to_datetime("2020-1-4")
#print(testdf['当前'])
print(testdf['日期'])
#testdf['周数']=(testdf['日期']-testdf['当前']).dt.days

testdf['周数']=(maxtime-testdf['日期']).dt.days//7+1
#这个格式显示结果为数字
print(testdf['周数'])
#d=pd.to_datetime(testdf['日期'].max()
#print(pd.to_datetime(testdf['日期'].max())


# testdf['diff_time'] = (testdf['周数'] -testdf['日期']).values/np.timedelta64(1, 'h')
#print(testdf['diff_time'])

#testdf['测试周数']=testdf['日期']-testdf['当前']
#print("打印测试周数")
#print(testdf['测试周数'])
print(testdf)
                 
print(list(testdf))
print(list(testdf2))
if list(testdf)==list(testdf2):
    print("列名相同")
print (testdf.shape[1],testdf2.shape[1])
ru=testdf.shape[1]-testdf2.shape[1]
print(ru)


# 1. 判断列数是否相同

#如果列数相同，把表格中的表头换成目标表的。Append操作。
