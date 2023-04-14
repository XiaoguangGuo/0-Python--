import pandas as pd
import os
import shutil
key =['US','CA','MX','UK','IT','DE','JP','ES','FR']
# 复制TSV在途库存
shippedfiledic={ "US" : '在途库存.xlsx' , 'CA' : 'Canada在途库存.xlsx' , 'MX' : 'MX在途库存.xlsx','UK':'UK在途库存.xlsx ','IT':'IT在途库存.xlsx','DE':'DE在途库存.xlsx','JP':'JP在途库存.xlsx','ES':'ES在途库存.xlsx' ,'FR':'FR在途库存.xlsx'}
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
    
    data_tsv5= pd.read_excel(r'D:\\运营\\计划数据\\Newcountries\\在途库存\\'+ str(file), header=0)
    #读取源文件TSV文件总是有问题需要先存成Excel
    batchnumber= data_tsv5.iat[0,1]
    #读取批次号
    data_tsv5= pd.read_excel(r'D:\\运营\\计划数据\\Newcountries\\在途库存\\'+ str(file),skiprows=8, header=0)
    data_tsv5['到货日期']=""
    data_tsv5['周数']=""


    #读取源文件去掉前8行；可以用去掉前8行重写
    data_tsv5["批次"]=batchnumber
               
    #加入批次号作为一列；可以用assign重写
    for i in range(len(key)):
    #遍历国家名字典
        if key[i] in file:
        #如果数据源文件名中包含国家
                    
            data_shipped=pd.read_excel(shippedfilepath+ shippedfiledic[key[i]] )
            #读取国家i的目标文件
            print(key[i])
            print(str(key[i]))
            data_tsv5.columns=data_shipped.columns                            
            data_shipped=data_shipped.append(data_tsv5,ignore_index=True)
            #将源数据文件加到目标文件                          
            #else
            #print("查看resock目标文件库，缺key列表国家的目标文件"）    
            break
            #一旦符合条件后面就不循环找了;实际就是找到目标文件中的第一个国家就跳出。
        else:
            print("检查源文件")
        
      
    data_shipped.to_excel(r'D:\SailingstarFBA计划\'+str(shippedfiledic[key[i]]'+".xlsx",sheet_name="Sheet1",startrow=0,header=True,index=False)

         
    print("已导出"+str(key[i])+"-restock-report")
    shutil.move(r'D:\\运营\\计划数据\\Newcountries\\在途库存\\'+ str(file), 'D:/运营/HistoricalData/计划数据/Newcountries/restock')
   
    a+=1
print("完成复制在途库存，完成了"+str(a)+"个在途库存导入")
      
#完成复制在途库存；可以写成一个def函数
