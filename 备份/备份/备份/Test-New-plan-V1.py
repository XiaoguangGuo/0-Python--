# -*- coding:utf-8 –*-
import os
import pandas as pd



#导入库存

src_dir_path_inventory=r'D:\运营\计划数据\Newcountries\当日库存'

key =['US','CA','MX','UK','IT','DE','JP']
 
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
        
    elif key[1]in file:
        print("有CA库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("比较列",data_csv.columns,data_inventory_CA.columns)
        data_csv.columns=inventorycolumns_CA    
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\Canada当日Amazon库存', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)

        
      
    elif key[2]in file:
        print("有MX库存")
        print(file)        
        #df_data.columns.tolist())
       print(data_csv)
        Print("bijiaolie",data_csv.columns,data_inventory_MX.columns)
        data_csv.columns=inventorycolumns_MX                      
        data_csv.to_excel(r'D:\SailingstarFBA计划\MX当日库存.xlsx',sheet_name="24493532708018574",startrow=0,header=True,index=False)

    elif key[3]in file:
        print("有UK库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("bijiaolie",data_csv.columns,data_inventory_UK.columns)
        data_csv.columns=inventorycolumns_UK    
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\UK当日Amazon库存', sheet_name="UK25372824608018570",startrow=0,header=True,index=False)

    elif key[4]in file:
        print("有IT库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("bijiaolie",data_csv.columns,data_inventory_IT.columns)
        data_csv.columns=inventorycolumns_IT   
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\IT当日Amazon库存', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)

    elif key[5]in file:
        print("有DE库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("bijiaolie",data_csv.columns,data_inventory_DE.columns)
        data_csv.columns=inventorycolumns_DE   
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\DE当日Amazon库存', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)

    elif key[6]in file:
        print("有JP库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        Print("bijiaolie",data_csv.columns,data_inventory_JP.columns)
        data_csv.columns=inventorycolumns_JP   
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\JP当日Amazon库存', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)

    

        
        
    else:
        print("什么库存文件都没有")


     shutil.move(r'D:\\运营\\计划数据\\Newcountries\\当日库存'+ str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')   
     #将文件转移到历史文件夹
     
# 导入reStock
sheetnamedic={ "US" : 'US-restock-report' , 'CA' : 'CA-restock-report_12-22-2020_09' , 'MX' : 'MX-restock-report_','UK':'UK-restock-report','IT':'UK-restock-report','DE':'UK-restock-report','JP':'JP-restock-report_12-22-2020_09' }
src_dir_path_restock=r'D:\运营\计划数据\Newcountries\restock'
print(os.listdir(src_dir_path_restock))
for file in os.listdir(src_dir_path_restock):
    data_csv2 = pd.read_table(r'D:\\运营\\计划数据\Newcountries\\restock'+ str(file))    # 读取以分        
    for i in range(len(key))
        if key[i] in file 
            print(key[i])
            print(str(key[i]))
            data_csv2.to_excel(r'D:\SailingstarFBA计划\'+str(key[i]'+"-restock-report.xlsx",sheet_name=str(sheetnamedic[key[i]]),startrow=0,header=True,index=False)        
            print("已导出"+str(key[i])+"-restock-report")
            break
         else
            print("查看resock目标文件库，缺key列表国家的目标文件"）
    shutil.move(r'D:\\运营\\计划数据\\Newcountries\\当日库存'+ str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/restock')   
         
        
#re导出完毕

  
    
# 复制TSV在途库存
shippedfiledic={ "US" : '在途库存' , 'CA' : 'Canada在途库存' , 'MX' : 'MX在途库存','UK':'UK在途库存 ','IT':'IT在途库存','DE':'DE在途库存','JP':'JP在途库存','ES':'ES在途库存' ,'FR':'FR在途库存'}
src_dir_path_shipped=r'D:\运营\计划数据\老站\在途库存'
#数据源文件目录
print(os.listdir(src_dir_path_shipped))


for file in os.listdir(src_dir_path_shipped):
    #遍历数据源文件
                  
    #旧程序data_shipped_US=pd.read_excel(r'D:\SailingstarFBA计划\在途库存.xlsx')
    #data_shipped_CA=pd.read_excel(r'D:\SailingstarFBA计划\Canada在途库存.xlsx')
    #data_shipped_MX=pd.read_excel(r'D:\SailingstarFBA计划\Mexico在途库存.xlsx')
    #salescolumns_US=data_shipped_US.columns.tolist()
    #salescolumns_CA=data_shipped_CA.columns.tolist()
    #salescolumns_MX=data_shipped_MX.columns.tolist()
    
    data_tsv5= pd.read_excel(r'D:\\运营\\计划数据\\老站\\在途库存\\'+ str(file))
    #读取源文件
    batchnumber= data_tsv5.iat[0,1]
    #读取批次号
    data_tsv5= pd.read_excel(r'D:\\运营\\计划数据\\老站\\在途库存\\'+ str(file),skiprows=8)
    #读取源文件去掉前8行；可以用去掉前8行重写
    data_tsv5["批次"]=batchnumber
    #加入批次号作为一列；可以用assign重写
    for i in range(len(key))
        
        if key[i] in file
                  
            data_shipped=pd.read_excel(r'D:\SailingstarFBA计划\'+str(shippedfiledic[key[i]]+'.xlsx')
            
            print(key[i])
            print(str(key[i]))
            data_shipped=data_shipped.append(data_tsv5,ignore_index=True)
            data_csv5.to_excel(r'D:\SailingstarFBA计划\'+str(shippedfiledic[key[i]]'+".xlsx",sheet_name="Sheet1",startrow=0,header=True,index=False)        
            print("已导出"+str(key[i])+"-restock-report")
            break
         #else
            #print("查看resock目标文件库，缺key列表国家的目标文件"）
    shutil.move(r'D:\\运营\\计划数据\\Newcountries\\当日库存'+ str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/restock')   
         











  
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US在途")

        data_tsv5['到货日期']=""
        data_tsv5['周数']=""
        print("lIESHU",data_tsv5.columns,salescolumns_US)
        data_tsv5.columns=salescolumns_US
        
    
        print(data_tsv5)
        data_shipped_US=data_shipped_US.append(data_tsv5,ignore_index=True)
                     
       #追加到在途计划 data_csv2.to_excel(r'D:\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)

        data_shipped_US.to_excel(r'D:\2019plan\在途库存.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
        print("US在途更新完成")
    #CA
    elif key[1]in file:
        print(file)
    # 执行语句
        print("有CA在途")   
         

        data_tsv5['到货日期']=""
        data_tsv5['周数']=""
        print("lIESHU",data_tsv5.columns,salescolumns_CA)
        data_tsv5.columns=salescolumns_CA
        
    
        print(data_tsv5)
        data_shipped_CA=data_shipped_CA.append(data_tsv5,ignore_index=True)
                     
       #追加到在途计划 data_csv2.to_excel(r'D:\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)

        data_shipped_CA.to_excel(r'D:\2019plan\Canada在途库存.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
        print("CA在途更新完成")
        
    elif key[2]in file:
        print("有MX在途")           
        data_tsv5['到货日期']=""
        data_tsv5['周数']=""
        print("在途库存列数比较",data_tsv5.columns,salescolumns_MX)
        data_tsv5.columns=salescolumns_MX
        
    
        print(data_tsv5)
        data_shipped_MX=data_shipped_MX.append(data_tsv5,ignore_index=True)
                     
       #追加到在途计划 data_csv2.to_excel(r'D:\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)

        data_shipped_MX.to_excel(r'D:\2019plan\Mexico在途库存.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
        print("MX在途更新完成")

    
    else:
        print("什么在途文件都没有")


       
