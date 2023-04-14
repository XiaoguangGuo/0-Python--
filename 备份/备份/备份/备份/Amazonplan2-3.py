

# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl 
import pandas as pd
df = pd.read_excel(r'D:\运营\计划\学生数据表.xlsx')
df.head()
 

df2=pd.DataFrame((x.split('_')for x in df['数据']),
                  columns=['ASk','ADk','asd'])
                  
df2.to_excel(r'D:\运营\计划\学生数据表2.xlsx',index=False)

