# 思路： 在文件夹中找各个国家的csv文件：循环。表格加日期和周数；如果列数相同，列名相同，就用目标表的列名赋值给原表。append到销售数据总表xlsx。
import glob, os
import pandas as pd

src_dir_path_sales=r'D:\运营\计划数据\老站\销售数据'
# 设置来源文件搜索目录
print(os.listdir(src_dir_path_sales))
key =['US','CA','MX']
#设置需要搜索的国家名字

# 以后做函数来简化程序def data_csv_open(file)
# def sourcesales_totargetsales(path,listofcountry,target_excel)未来做

data_sales_US=pd.read_excel(r'D:\2019plan\周销售数据.xlsx')
data_sales_CA=pd.read_excel(r'D:\2019plan\Canada周销售数据.xlsx')
data_sales_MX=pd.read_excel(r'D:\2019plan\Mexico周销售数据.xlsx')
#未来可以做一个文件名列表包含文件名和sheet名
salescolumns_US=data_sales_US.columns.tolist()
salescolumns_CA=data_sales_CA.columns.tolist()
salescolumns_MX=data_sales_MX.columns.tolist()
#取得目标文件的dataframe和列名


for file in os.listdir(src_dir_path_sales):
     
    if key[0] in file:
        data_csv3 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))
    # 打开原文件的dataframe
    
    #把文件名中的日期写进来
# basename
    #US
        data_csv_sales =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1])
    #加日期
        data_csv_sales['日期'] = pd.to_datetime( data_csv_sales['日期'])
        print( data_csv_sales['日期'])
        maxtime=pd.to_datetime(data_csv_sales["日期"].max())
        print(maxtime)
        data_csv_sales['周数']=(maxtime- data_csv_sales['日期']).dt.days//7+1
        
    #加周数
        ru= data_csv_sales.shape[1]-data_sales_US.shape[1]
    #比较列数
        if ru=0:
        #给列名赋值确保可以
            data_csv_sales.columns=salescolumns_US
            #做append
            data_sales_US=data_sales_US.append(data_csv.sales,ignore_index=True)
            data_sales_US.to_excel(r'D:\2019plan\周销售数据.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
            print("US销售数据更新完成"）
        else:
            print("请修改目标文件，以保证列数相同"）
   
              
    # CA
    
    elif key[1] in file:
        data_csv3 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))
    # 打开原文件的dataframe
    
    #把文件名中的日期写进来
   
        data_csv_sales =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1])
    #加日期
        data_csv_sales['日期'] = pd.to_datetime(data_csv_sales['日期'])
        print(data_csv_sales['日期'])
        maxtime=pd.to_datetime(data_csv_sales["日期"].max())
        print(maxtime)
        data_csv_sales['周数']=(maxtime-data_csv_sales['日期']).dt.days//7+1
  
    #加周数
        ru=data_csv_sales.shape[1]-data_sales_CA.shape[1]
        if ru=0:
    #给列名赋值确保可以
             data_csv_sales.columns=salescolumns_CA
    #做append
             data_sales_CA=data_sales_CA.append(data_csv.sales,ignore_index=True)
             data_sales_CA.to_excel(r'D:\2019plan\Canada周销售数据.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
             print("CA销售数据更新完成"）
            print("请修改目标文件，以保证列数相同"）
              
    # MX
       
    elif key[1] in file:
        data_csv3 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))
    # 打开原文件的dataframe
    
    #把文件名中的日期写进来
   
        data_csv_sales =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1])
    #加日期
        data_csv_sales['日期'] = pd.to_datetime(data_csv_sales['日期'])
        print(data_csv_sales['日期'])
        maxtime=pd.to_datetime(data_csv_sales["日期"].max())
        print(maxtime)
        data_csv_sales['周数']=(maxtime-data_csv_sales['日期']).dt.days//7+1
  
    #加周数
        ru=data_csv_sales.shape[1]-data_sales_MX.shape[1]
        if ru=0:
    #给列名赋值确保可以
            data_csv_sales.columns=salescolumns_MX
    #做append
            data_sales_CA=data_sales_CA.append(data_csv.sales,ignore_index=True)
            data_sales_CA.to_excel(r'D:\2019plan\Mexico周销售数据.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
            print("MX销售数据更新完成"）
        else:
        print("请修改目标文件，以保证列数相同"）
              
    else:
    print("什么销售文件都没有")
    

