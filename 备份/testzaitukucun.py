# -*- coding:utf-8 –*-
import os
import pandas as pd

# 复制TSV在途库存
key =['US','CA','MX']
src_dir_path_shipped=r'D:\运营\计划数据\老站\在途库存'
print(os.listdir(src_dir_path_shipped))


for file in os.listdir(src_dir_path_shipped):
    
    data_shipped_US=pd.read_excel(r'D:\2019plan\在途库存.xlsx')
    data_shipped_CA=pd.read_excel(r'D:\2019plan\Canada在途库存.xlsx')
    data_shipped_MX=pd.read_excel(r'D:\2019plan\Mexico在途库存.xlsx')
    salescolumns_US=data_shipped_US.columns.tolist()
    salescolumns_CA=data_shipped_CA.columns.tolist()
    salescolumns_MX=data_shipped_MX.columns.tolist()

    
    data_tsv5 = pd.read_csv(r'D:/运营/计划数据/老站/在途库存/'+ str(file), sep='\t',nrows =5)
    print(data_tsv5.iloc[0,1])
    batchnumber=data_tsv5.iloc[0,1]
    print(data_tsv5)
    print(batchnumber)
    
    data_tsv5= pd.read_csv(r'D:/运营/计划数据/老站/在途库存/'+ str(file), sep='\t',header=6)
    print(data_tsv5)
    data_tsv5["批次"]=batchnumber
    
    data_tsv5['到货日期']=""
    data_tsv5['周数']=""
    print(data_tsv5)
    
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US在途")

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
