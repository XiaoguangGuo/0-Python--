

# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
import xlsxwriter

wertyyu = pd.read_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx')
 
    
df=wertyyu["数据"].str.split("_",expand=True)

wertyyu["国家"]=df[0]
wertyyu["其他"]=df[1]
print(wertyyu.head(5))
print(df.head(5))

wertyyu.to_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx',index=False)

