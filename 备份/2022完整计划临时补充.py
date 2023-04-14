
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil


print("以下为运行老站计划")
print("以下为运行老站计划")


# -*- coding:utf-8 –*-
import os
import pandas as pd



##以下为运行新站的计划
##以下为运行新站的计划
##以下为运行新站的计划
print("以下为运行新站的计划")
print("以下为运行新站的计划")
        
# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil



#导入库存

src_dir_path_inventory=r'D:\运营\计划数据\Newcountries\当日库存'

key =['NEW-US','CA','MX','UK','IT','DE','JP','ES','FR','HM-US']
 
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
    
    data_csv = pd.read_csv(r'D:\\运营\\计划数据\\Newcountries\\当日库存\\'+ str(file),encoding='Latin1')
    # 读文件
#
    #US
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US库存")
       
         # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        print(data_csv)
        print("比较列",data_csv.columns,data_inventory_US.columns)
        data_csv.columns=inventorycolumns_US                      
        data_csv.to_excel(r'D:\SailingstarFBA计划\当日Amazon库存.xlsx',sheet_name="US-new24374599305018570",startrow=0,header=True,index=False)
        
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
        
    elif key[1]in file:
        print("有CA库存")
 
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        print("比较列",data_csv.columns,data_inventory_CA.columns)
        data_csv.columns=inventorycolumns_CA    
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\Canada当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
        
      
    elif key[2]in file:
        print("有MX库存")
   
        #df_data.columns.tolist())
        print(data_csv)
        print("bijiaolie",data_csv.columns,data_inventory_MX.columns)
        data_csv.columns=inventorycolumns_MX                      
        data_csv.to_excel(r'D:\SailingstarFBA计划\MX当日库存.xlsx',sheet_name="24493532708018574",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
        
    elif key[3]in file:
        print("有UK库存")
  
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        print("bijiaolie",data_csv.columns,data_inventory_UK.columns)
        data_csv.columns=inventorycolumns_UK    
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\UK当日Amazon库存.xlsx', sheet_name="UK25372824608018570",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
    elif key[4]in file:
        print("有IT库存")
    
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        print("bijiaolie",data_csv.columns,data_inventory_IT.columns)
        data_csv.columns=inventorycolumns_IT  
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\IT当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
    elif key[5]in file:
        print("有DE库存")
        
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist()
        print("bijiaolie",data_csv.columns,data_inventory_DE.columns)
        data_csv.columns=inventorycolumns_DE   
                         
         # data_csv.to_excel(r'D:\SailingstarFBA计划\DE当日Amazon库存', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
    elif key[6]in file:
        print("有JP库存")
       
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        print("bijiaolie",len(data_csv.columns),len(data_inventory_JP.columns))
        data_csv.columns=inventorycolumns_JP   
                         
        data_csv.to_excel(r'D:\SailingstarFBA计划\JP当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        shutil.move('D:\\运营\\计划数据\\Newcountries\\当日库存\\'+str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/当日库存')
        #将文件转移到历史文件夹

        
        
    else:
        print("没有US,CA,MX,UK,IT,DE,JP当日库存")
   
    
     
# 导入reStock
sheetnamedic={ "US" : 'US-restock-report' , 'CA' : 'CA-restock-report_12-22-2020_09' , 'MX' : 'MX-restock-report_','UK':'UK-restock-report','IT':'UK-restock-report','DE':'UK-restock-report','JP':'JP-restock-report_12-22-2020_09' }
src_dir_path_restock=r'D:\运营\计划数据\Newcountries\restock'
print(os.listdir(src_dir_path_restock))
for file in os.listdir(src_dir_path_restock):
    data_csv2 = pd.read_table(r'D:\\运营\\计划数据\Newcountries\\restock\\'+ str(file))    # 读取以分        
    for i in range(len(key)):
        if key[i] in file: 
            print(key[i])
            print(str(key[i]))
            data_csv2.to_excel(r'D:\SailingstarFBA计划\ '+str(key[i])+"-restock-report.xlsx",sheet_name=str(sheetnamedic[key[i]]),startrow=0,header=True,index=False)
            print("已导出"+str(key[i])+"-restock-report")
            break
        else:
            print("查看resock目标文件库，缺key列表国家的目标文件")
    shutil.move(r'D:\\运营\\计划数据\\Newcountries\\restock\\'+ str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/restock')   
         
        
#re导出完毕

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import openpyxl
key =['US','CA','MX','UK','IT','DE','JP','ES','FR']
  
# 复制TSV在途库存
shippedfiledic={ 'US' : '在途库存' , 'CA' : 'Canada在途库存' , 'MX' : 'MX在途库存','UK':'UK在途库存','IT':'IT在途库存','DE':'DE在途库存','JP':'JP在途库存','ES':'ES在途库存' ,'FR':'FR在途库存'}
src_dir_path_shipped=r'D:\运营\计划数据\Newcountries\在途库存'
shippedfilepath=r'D:\\SailingstarFBA计划\\'
#数据源文件目录
print(os.listdir(src_dir_path_shipped))

a=0

for file in os.listdir(src_dir_path_shipped):
    #遍历数据源文件
                  
    #旧程序data_shipped_US=pd.read_excel(r'D:\SailingstarFBA计划\在途库存.xlsx')
    #data_shipped_CA=pd.read_excel(r'D:\SailingstarFBA计划\Canada在途库存.xlsx')
    #data_shipped_MX=pd.read_excel(r'D:\SailingstarFBA计划\Mexico在途库存.xlsx')
    #salescolumns_US=data_shipped_US.columns.tolist()
    #salescolumns_CA=data_shipped_CA.columns.tolist()
    #salescolumns_MX=data_shipped_MX.columns.tolist()
    data_tsv5 = pd.read_csv(r'D:/运营/计划数据/NewCountries/在途库存/'+ str(file), sep='\t',nrows =5)
    print(data_tsv5.iloc[0,1])
    batchnumber=data_tsv5.iloc[0,1]
    print(data_tsv5)
    print(batchnumber)
    data_tsv5= pd.read_csv(r'D:/运营/计划数据/NewCountries/在途库存/'+ str(file), sep='\t',header=7)
    print(data_tsv5)
 
    #batchnumber= data_tsv5.iat[0,1]
    #读取批次号
    if len(data_tsv5.columns)>10:
        x=[9,10]
        data_tsv5.drop(data_tsv5.columns[x], axis=1, inplace=True)
        print(data_tsv5)

    data_tsv5["批次"]=batchnumber
    data_tsv5['到货日期']=""
    data_tsv5['周数']=""
    print(data_tsv5)

    #读取源文件去掉前8行；可以用去掉前8行重写

    print(data_tsv5)
               
    #加入批次号作为一列；可以用assign重写
    b=0
    for i in range(len(key)):
    #遍历国家名字典
        if key[i] in file:
        #如果数据源文件名中包含国家
                    
            datashipped=pd.read_excel(shippedfilepath+ shippedfiledic[key[i]]+'.xlsx' )
            print(datashipped)
            #读取国家i的目标文件
            print(key[i])
                       
            data_tsv5.columns=datashipped.columns                            
            datashipped=datashipped.append(data_tsv5,ignore_index=True)
            datashipped.to_excel(r'D:\\SailingstarFBA计划\\'+shippedfiledic[key[i]]+'.xlsx',sheet_name="Sheet1",startrow=0,header=True,index=False)
            print("已导出"+str(key[i])+"在途库存导入")
            a+=1
            
            #将源数据文件加到目标文件                          
            #else
            #print("查看resock目标文件库，缺key列表国家的目标文件"）    
            break
            #一旦符合条件后面就不循环找了;实际就是找到目标文件中的第一个国家就跳出。
        else:
            b+=1
            
        if b==len(key):
            print("检查源文件")     
        
      
    

         
    
    shutil.move(r'D:\\运营\\计划数据\\Newcountries\\在途库存\\'+ str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/在途库存')
   
    
print("完成复制在途库存，完成了"+str(a)+"个在途库存导入")
      
#完成复制在途库存；可以写成一个def函数

#复制销售数据

samplecolumn=pd.read_excel('D:\\运营\\计划数据\\NewCountries\\samplecolumnUSMX.xlsx')
#取得墨西哥和美国的标准列名
USMXColumnlist=samplecolumn.columns.tolist()
key =['US','CA','MX','UK','IT','DE','JP','ES','FR']
#开始导入销售数据
#思路：1.遍历销售数据2. 导入file1加上国家，if 如果是美国，墨西哥（或者如果列数为固定的）则插入2个字段 ;将file1 的dataframe追加到 目标文件dataframe ;导出目标dataframe到excel
#思路：1.遍历销售数据2. 导入file1加上国家，if 如果是美国，墨西哥（或者如果列数为固定的）则插入2个字段 ;将file1 的dataframe追加到 目标文件中间件dataframe ;导出目中间件dataframe到目标dataframe到excel

salesfilepath = 'D:\\运营\\计划数据\\NewCountries\\销售数据\\'
Newallsales=pd.read_excel('D:/SailingstarFBA计划/NEW-ALL周销售数据.xlsx')
NewColumnlist= Newallsales.columns.tolist()
#取得目标文件的列名
               
a=0
for salesfile in os.listdir(salesfilepath):
    print(salesfile)
#遍历数据文件
    sourcedata=pd.read_csv(salesfilepath+str(salesfile)).assign(Country=os.path.basename(salesfile).split('_')[0], 日期=os.path.basename(salesfile).split('_')[1])
    sourcedata['日期'] = pd.to_datetime(sourcedata['日期'])
    print(sourcedata[['Country','日期']])
    for i in range(len(key)):
    #遍历国家名     
        if key[i] in salesfile:
                       
            if len(sourcedata.columns)==15:
            #如果列数=14
                #USMXColumnlist =["(Parent) ASIN","(Child) ASIN","Title","SKU","Sessions","Session Percentage","Page Views","Page Views Percentage","Buy Box Percentage","Units Ordered","Unit Session Percentage","Ordered Product Sales","Total Order Items","Country","日期"]
                #定义标准列名，给美国和MX赋值列名，reindex列名

                print(USMXColumnlist)
                 

                sourcedata.columns=USMXColumnlist
             
                #给数据源文件列名赋标准值
                
                print(NewColumnlist)
                sourcedata=sourcedata.reindex(columns=NewColumnlist)
                #给源文件插入全部缺失的列名，行值为0
                print(len(sourcedata.columns))
                #test.to_excel('D:/运营/计划数据/NewCountries/销售数据/temp.xlsx')
            elif len(sourcedata.columns)==19:
                sourcedata['周数']=""
                print(len(sourcedata.columns),len(NewColumnlist))
                sourcedata.columns=NewColumnlist
                 ##给数据源文件列名赋目标文件列名标准值
            else:
                  print("检查销售数据文件的列名")


            Newallsales=Newallsales.append(sourcedata,ignore_index=True)
            print(Newallsales)
             #追加数据到目标文件
            print("添加销售数据"+str(key[i]))
            break
    shutil.move(r'D:\\运营\\计划数据\\Newcountries\\销售数据\\'+ str(salesfile), 'D:/运营/HistoricalData/计划数据/Newcountries/销售数据')
    a+=1

#Newallsales["日期"] = Newallsales["日期"].dt.strftime("%Y-%m-%d")
#统一Newallsales的日期格式但是报错


maxtime2=pd.to_datetime(Newallsales['日期'].max())
  
print("最晚时间",maxtime2)
Newallsales['日期']=pd.to_datetime(Newallsales['日期'])
#获取最晚时间
Newallsales['周数']=(maxtime2-Newallsales['日期']).dt.days//7+1
#给周数赋值
Newallsales.to_excel(r'D:\SailingstarFBA计划\NEW-ALL周销售数据.xlsx',sheet_name="Sheet1",startrow=0,header=True,index=False)

print("导出目标文件到Excel，数量为"+str(a))
    
#销售数据复制完毕



