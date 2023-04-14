# -*- coding:utf-8 –*-
import pandas as pd
import os
src_dir_path_inventory=r'D:\运营\计划数据\Newcountries\当日库存'
for file in os.listdir(src_dir_path_inventory):
    
    print(os.listdir(src_dir_path_inventory))
    print("str"+str(file))
#data_csv = pd.read_csv(r'D:\\运营\\计划数据\\Newcountries\\当日库存\\NEW-FR_54253018925.csv')
data_csv = pd.read_csv(src_dir_path_inventory+'\\'+ str(file))

print(data_csv)
