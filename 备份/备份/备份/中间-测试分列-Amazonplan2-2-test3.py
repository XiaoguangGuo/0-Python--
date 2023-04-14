#测试分列 

# -*- coding: utf-8 -*-
import pandas as pd


df = pd.read_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx')
df2 = pd.DataFrame((x.split('_') for x in df['数据']),
                               columns=['国家', 'others'])
df["国家"]=df2[0]
df["其他"]=df2[1]
df2.to_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan3.xlsx',index=False)  # 保存 index=False是去掉序号index
df.to_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan5.xlsx',index=False)
