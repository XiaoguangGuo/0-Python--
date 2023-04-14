

# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
import xlsxwriter

df = pd.read_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx')
df2 = pd.DataFrame((x.split('_') for x in df['数据']),
                               columns=['R','S','T','U'])


with pd.ExcelWriter(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx',
                    mode='a',engine="openpyxl") as writer:  
    df2.to_excel(writer)
