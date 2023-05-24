
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil

newdate=input('输入最新日期y-m-d',) #输入最新一周的日期
maxtime=datetime.datetime.strptime(newdate,'%Y-%m-%d')
print(newdate)
print(maxtime)

filePath = r'D:\运营\1数据源\周SearchTerm数据'

f_name = os.listdir(filePath)


SeartchtermAll=pd.read_excel(r'D:\\运营\\2生成过程表\\Sponsored Products Search term report.xlsx')


for seartchtermfile in f_name:


    DFseartchtermfile=pd.read_excel(filePath+"\\"+str(seartchtermfile),engine="openpyxl").assign(Country=os.path.basename(seartchtermfile).split('_')[0])

    shutil.move(filePath+"\\"+str(seartchtermfile), r'D:/运营/1数据源/周SearchTerm数据')

    SeartchtermAll=pd.concat([SeartchtermAll,DFseartchtermfile])

    
SeartchtermAll["周数"]=(maxtime-SeartchtermAll["Start Date"]).dt.days//7+1


SeartchtermAll.to_excel(r'D:\\运营\\2生成过程表\\Sponsored Products Search term report2222.xlsx',engine="openpyxl",sheet_name="Sponsored Products Search term report",index=False)                                    

