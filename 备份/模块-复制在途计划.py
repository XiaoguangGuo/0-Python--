#复制在途计划的模块

src_dir_path_shipped=r'D:\运营\计划数据\老站\在途计划'
print(os.listdir(src_dir_path_shipped))
data_shipped_US=pd.read_excel(r'D:\2019plan\在途计划.xlsx')
data_shipped_CA=pd.read_excel(r'D:\2019plan\Canada在途计划.xlsx')
data_shipped_MX=pd.read_excel(r'D:\2019plan\Mexico在途计划.xlsx')
salescolumns_US=data_sales_US.columns.tolist()
salescolumns_CA=data_sales_CA.columns.tolist()
salescolumns_MX=data_sales_MX.columns.tolist()

for file in os.listdir(src_dir_path_shipped):
    data_csv5= pd.read_csv(r'D:\\运营\\计划数据\\老站\\restock\\'+ str(file),sep='\t')    # 读取以分
    batchnumber= data_csv5.iat[2,1]
    data_csv5= pd.read_csv(r'D:\\运营\\计划数据\\老站\\restock\\'+ str(file),sep='\t'，skiprows=9)    # 读取以分        
    data_csv5["批次"]=batchnumber
  
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US在途")

    
        data_csv5.columns=salescolumns_US
        data_csv5['到货日期']=""
        data_csv5['周数']=""
    
        print(data_csv5)
        data_shipped_US=data_shipped_US.append(data_csv5,ignore_index=True)
                     
       #追加到在途计划 data_csv2.to_excel(r'D:\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)

        data_shipped_US.to_excel(r'D:\2019plan\在途计划.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
        print("US在途更新完成"）
    #CA
    elif key[1]in file:
        print("有CA在途")
        print(file） 
        data_csv5.columns=salescolumns_CA
        data_csv5['到货日期']=""
        data_csv5['周数']=""
        print(data_csv5)
        data_shipped_CA=data_shipped_CA.append(data_csv5,ignore_index=True)
                     
       #追加到在途计划 data_csv2.to_excel(r'D:\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)

        data_shipped_CA.to_excel(r'D:\2019plan\Canada在途计划.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
        print("CA在途更新完成"）
    elif key[2]in file:
        print("有MX在途")
        print(file)        
               
      
        data_csv5.columns=salescolumns_MX
        data_csv5['到货日期']=""
        data_csv5['周数']=""
        print(data_csv5)
        data_shipped_MX=data_shipped_CA.append(data_csv5,ignore_index=True)
                     
       #追加到在途计划 data_csv2.to_excel(r'D:\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)

        data_shipped_MX.to_excel(r'D:\2019plan\Mexico在途计划.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
        print("MX在途更新完成"）
   
       ,index=False)

        print(data_csv5)
        
    else:
        print("什么在途文件都没有")



                



