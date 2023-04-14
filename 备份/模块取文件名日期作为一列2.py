import pandas as pd
import glob
import os
import csv
import xlrd
from xlutils.copy import copy
key =['US','CA','MX']
src_dir_path_sales=r'D:\运营\计划数据\老站\销售数据'

for file in os.listdir(src_dir_path_sales):
# files = glob.glob(r'D:/运营/计划数据/老站/销售数据/*.csv')
# print(files)
# basename  

# for fp in files:
    if key[0] in file:
        print(file)
        
    #例句   
        data_csv3 =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1]) 
        print(data_csv3)


        
       # 不能用这个方法： 会删掉之前的文件 data_csv3.to_excel(r'D:/运营/计划数据/老站/销售数据/Canada周销售数据.xlsx')
    elif key[1] in file:
        print"加拿大sales"
        data_csv3 =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1])
        # 不能用这个方法： 会删掉之前的文件 data_csv3.to_excel(r'D:/运营/计划数据/老站/销售数据/Canada周销售数据.xlsx')
    elif key[2] in file
        print"墨西哥sales"
        data_csv3 =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1])
    else
、        print"什么sales都没有"



def write_excel_xls_append
#import pandas as pd
#path1 = ‘D:/pandas1/读取文件.txt’
#read_data = pd.read_csv(path1, header=None,
#names = [‘gender’, ‘name’, ‘age’, ‘cellphone’,‘address’, ‘date’]
#, index_col= ‘date’
#, skiprows=[2,3])
#*index_col 为设置索引列，
#header 和 name一起使用来设置表头。否则会将第一行默认为表头
#skiprows 将某行跳过，不输出        
 
#with open(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))as f:
#data_csv3 = csv.reader(f)
    
#for row in reader：
#print(row)
#print(data_csv3)
#> cat some.csv 
#“114111”,“飞机,火车和汽车”,“50”,“BOOK”
 
#> python test.py 
# ['114111','Planes,火车和汽车','50','BOOK'] 

