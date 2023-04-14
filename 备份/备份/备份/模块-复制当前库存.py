# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil



#导入库存

src_dir_path_inventory=r'D:\运营\计划数据\Newcountries\当日库存'

key =['US','CA','MX','UK','IT','DE','JP','ES','FR']
 
#获取原来库存文件的列名
data_inventory_US=pd.read_excel(r'D:\SailingstarFBA计划\当日Amazon库存.xlsx')
data_inventory_CA=pd.read_excel(r'D:\SailingstarFBA计划\Canada当日Amazon库存.xlsx')
data_inventory_UK=pd.read_excel(r'D:\SailingstarFBA计划\UK当日Amazon库存.xlsx')
data_inventory_IT=pd.read_excel(r'D:\SailingstarFBA计划\IT当日Amazon库存.xlsx')
data_inventory_DE=pd.read_excel(r'D:\SailingstarFBA计划\DE当日Amazon库存.xlsx')
data_inventory_MX=pd.read_excel(r'D:\SailingstarFBA计划\MX当日Amazon库存.xlsx')
data_inventory_JP=pd.read_excel(r'D:\SailingstarFBA计划\JP当日Amazon库存.xlsx')
                                
inventorycolumns_US=data_inventory_US.columns.tolist()
inventorycolumns_CA=data_inventory_CA.columns.tolist()
inventorycolumns_MX=data_inventory_MX.columns.tolist()
inventorycolumns_UK=data_inventory_UK.columns.tolist()
inventorycolumns_DE=data_inventory_DE.columns.tolist()
inventorycolumns_IT=data_inventory_IT.columns.tolist()
inventorycolumns_JP=data_inventory_JP.columns.tolist()

print(inventorycolumns_US)

# 遍历文件夹

for file in os.listdir(src_dir_path_inventory):
    
    print(os.listdir(src_dir_path_inventory))
    
    data_csv = pd.read_csv(r'D:\\运营\\计划数据\\Newcountries\\当日库存\\'+ str(file))
    # 读文件
#
    #US
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US库存")
       
         # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        print(data_csv)
        Print("比较列",data_csv.columns,data_inventory_US.columns)
        data_csv.columns=inventorycolumns_US                      
        data_csv.to_excel(r'D:\SailingstarFBA计划\当日Amazon库存.xlsx',sheet_name="US-new24374599305018570",startrow=0,header=True,index=False)
        
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
        
    elif key[1]in file:
        print("有CA库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("比较列",data_csv.columns,data_inventory_CA.columns)
        data_csv.columns=inventorycolumns_CA    
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\Canada当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
        
      
    elif key[2]in file:
        print("有MX库存")
        print(file)        
        #df_data.columns.tolist())
        print(data_csv)
        Print("bijiaolie",data_csv.columns,data_inventory_MX.columns)
        data_csv.columns=inventorycolumns_MX                      
        data_csv.to_excel(r'D:\SailingstarFBA计划\MX当日库存.xlsx',sheet_name="24493532708018574",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
        
    elif key[3]in file:
        print("有UK库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        print("bijiaolie",data_csv.columns,data_inventory_UK.columns)
        data_csv.columns=inventorycolumns_UK    
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\UK当日Amazon库存.xlsx', sheet_name="UK25372824608018570",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
    elif key[4]in file:
        print("有IT库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("bijiaolie",data_csv.columns,data_inventory_IT.columns)
        data_csv.columns=inventorycolumns_IT   
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\IT当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
    elif key[5]in file:
        print("有DE库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("bijiaolie",data_csv.columns,data_inventory_DE.columns)
        data_csv.columns=inventorycolumns_DE   
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\DE当日Amazon库存', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
    elif key[6]in file:
        print("有JP库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        print("bijiaolie",len(data_csv.columns),len(data_inventory_JP.columns))
        data_csv.columns=inventorycolumns_JP   
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\JP当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
    

        
        
    else:
        print("没有US,CA,MX,UK,IT,DE,JP当日库存")
   
