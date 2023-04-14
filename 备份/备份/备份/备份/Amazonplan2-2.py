

# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl 

df = pd.read_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx')
df2 = pd.DataFrame((x.split('_') for x in df['数据']),
                               columns=['国家', 'others'])
df2.to_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan3.xlsx', Columns=25,index=False)  # 保存 index=False是去掉序号index
