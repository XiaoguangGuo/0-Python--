import pandas as pd
import glob
import os
key =['US','CA','MX']
src_dir_path_sales=r'D:\运营\计划数据\老站\销售数据'
# for file in os.listdir(src_dir_path_sales)
files = glob.glob(r'D:/运营/计划数据/老站/销售数据/*.csv')
print(files)
# basename
for fp in files:
    if key[0] in fp:
      print(fp)
    data_csv3 =pd.read_csv(fp).assign(日期=os.path.basename(fp).split('_')[1]) 
    print(data_csv3)
    data_csv3.to_excel(r'D:/运营/计划数据/老站/销售数据/test.xlsx')


    
