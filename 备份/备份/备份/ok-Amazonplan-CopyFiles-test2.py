# -*- coding:utf-8 –*-
import os
import pandas as pd
src_dir_path_inventory=r'D:\运营\计划数据\老站\当日库存'

key =['US','CA','MX']
t=key[0]
print(t)
#获取原来库存文件的列名
data_inventory_US=pd.read_excel(r'D:\2019plan\当日Amazon库存.xlsx')
data_inventory_CA=pd.read_excel(r'D:\2019plan\Canada当前Amazon库存.xlsx')
data_inventory_MX=pd.read_excel(r'D:\2019plan\Mexico当日Amazon库存.xlsx')
                                
inventorycolumns_US=data_inventory_US.columns.tolist()
inventorycolumns_CA=data_inventory_CA.columns.tolist()
inventorycolumns_MX=data_inventory_MX.columns.tolist()
                                
print(inventorycolumns_US)

# 在文件夹里查找文件

for file in os.listdir(src_dir_path_inventory):
    print(os.listdir(src_dir_path_inventory))
    
    data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+ str(file))    # 读取以分        
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US库存")
       
         # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        print(data_csv)
        data_csv.columns=inventorycolumns_US                      
        data_csv.to_excel(r'D:\2019plan\当日Amazon库存.xlsx',sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        
    elif key[1]in file:
        print("有CA库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        data_csv.columns=inventorycolumns_CA    
                         
        data_csv.to_excel(r'D:\2019plan\Canada当前Amazon库存.xlsx', sheet_name="15828640259018099",startrow=0,header=True,index=False)

        
      
    elif key[2]in file:
        print("有MX库存")
        print(file)        
        #df_data.columns.tolist())
        
        data_csv.columns=inventorycolumns_MX  
                         
        data_csv.to_excel(r'D:\2019plan\Mexico当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)

        print(data_csv)
        
    else:
        print("什么库存文件都没有")

# 导入reStock

src_dir_path_restock=r'D:\运营\计划数据\老站\restock'
print(os.listdir(src_dir_path_restock))
for file in os.listdir(src_dir_path_restock):
    data_csv2 = pd.read_table(r'D:\\运营\\计划数据\\老站\\restock\\'+ str(file))    # 读取以分        
    if key[0] in file:
        print(file)
    # 执行语句
        print("有USrestock")
       
         # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        print(data_csv2)
                     
        data_csv2.to_excel(r'D:\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)
        
    elif key[1]in file:
        print("有CArestock")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        
                         
        data_csv2.to_excel(r'D:\2019plan\restock-report_CA.xlsx', sheet_name="REstock-CA",startrow=0,header=True,index=False)

        print(data_csv2)
      
    elif key[2]in file:
        print("有MXrestock")
        print(file)        
        #df_data.columns.tolist())
   
        data_csv2.to_excel(r'D:\2019plan\restock-report_MX.xlsx', sheet_name="restock-report_MX",startrow=0,header=True,index=False)

        print(data_csv2)
        
    else:
        print("什么restock文件都没有")

#复制销售数据
import glob, os
src_dir_path_sales=r'D:\运营\计划数据\老站\销售数据'
print(os.listdir(src_dir_path_sales))
data_csv2 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))

data_sales_US=pd.read_excel(r'D:\2019plan\周销售数据.xlsx')
data_sales_CA=pd.read_excel(r'D:\2019plan\Canada周销售数据.xlsx')
data_sales_MX=pd.read_excel(r'D:\2019plan\Mexico周销售数据.xlsx')
salescolumns_US=data_sales_US.columns.tolist()
salescolumns_CA=data_sales_CA.columns.tolist()
salescolumns_MX=data_sales_MX.columns.tolist()
#文件


for file in os.listdir(src_dir_path_sales):
     
    if key[0] in file
    data_csv3 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))    # 读取以分
    
    data['日期'] = ''
    # 把文件名的分列第一段写入日期

    files = glob.glob('files/*.csv')
    print(files)
# basename
    data_csv3 = pd.concat([pd.read_csv(fp).assign(New=os.path.basename(fp).split('_')[0]) for fp in files])

    
    

 


                


                      
         #  Print

                  #  For i in key
                      #if key[i] in file
                  #    print(key[i])

                    
