
# -*- coding:utf-8 –*-
import sqlite3
import os
import pandas as pd


conn=sqlite3.connect('D:/运营/sqlite/Top1M.db')
filepath_Top1M=r'D:\\运营\\1数据源\\TopsearchTerms\\'

for file in os.listdir(filepath_Top1M):
    file1=file.split(".")[0]
    date_parts = file1.split("_")[6:9]
    formatted_date = "".join(date_parts)
    filename=filepath_Top1M+str(file)
    print(filename)
    df_Top1M=pd.read_csv(filepath_Top1M+str(file),header=1, encoding='latin1').assign(Country=os.path.basename(file).split('_')[0], 日期=formatted_date)

    df_Top1M.to_sql("Top1M", conn, if_exists='append', index=False)
    print(df_Top1M)
