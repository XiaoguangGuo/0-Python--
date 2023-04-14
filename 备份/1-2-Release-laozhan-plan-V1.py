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
    
    data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+ str(file),encoding="Latin1")    # 读取以encoding='Latin1'分        
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US库存")
       
         # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
                                                 
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
    data_csv2 = pd.read_table(r'D:\\运营\\计划数据\\老站\\restock\\'+ str(file),encoding="Latin1")    # 读取以分        
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

#复制销售数据 20210221模块待写入
src_dir_path_sales=r'D:\运营\计划数据\老站\销售数据'
# 设置来源文件搜索目录
print(os.listdir(src_dir_path_sales))
key =['US','CA','MX']
#设置需要搜索的国家名字

# 以后做函数来简化程序def data_csv_open(file)
# def sourcesales_totargetsales(path,listofcountry,target_excel)未来做


for file in os.listdir(src_dir_path_sales):
     
    data_sales_US=pd.read_excel(r'D:\2019plan\周销售数据.xlsx')
    data_sales_CA=pd.read_excel(r'D:\2019plan\Canada周销售数据.xlsx')
    data_sales_MX=pd.read_excel(r'D:\2019plan\Mexico周销售数据.xlsx')
#未来可以做一个文件名列表包含文件名和sheet名
    salescolumns_US=data_sales_US.columns.tolist()
    salescolumns_CA=data_sales_CA.columns.tolist()
    salescolumns_MX=data_sales_MX.columns.tolist()
#取得目标文件的dataframe和列名

    if key[0] in file:
        print("开始处理US数据")
    
   
   
        data_csv_sales =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1])
    #读取源数据加日期 把文件名中的日期写进来
        data_csv_sales['日期'] = pd.to_datetime(data_csv_sales['日期'])
        print(data_csv_sales['日期'])
        data_csv_sales['周数']=""
            
        print(data_csv_sales)
      
        ru=data_sales_US.columns.size-data_csv_sales.columns.size 
        
        if ru==0:
        #如果列数相同
            data_csv_sales.columns=salescolumns_US
            data_sales_US=data_sales_US.append(data_csv_sales,ignore_index=True)
        #做append将源数据合并到目标文件
            maxtime=pd.to_datetime(data_sales_US["日期"].max())
        #查目标文件的最晚日期
            print("最晚时间",maxtime)
            data_sales_US ['周数']=(maxtime-data_sales_US['日期']).dt.days//7+1
        #周数写到目标文件
        #在导出之前加周数
            data_sales_US.to_excel(r'D:\2019plan\周销售数据.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
            print("US销售数据更新完成")
        else:
            print("US销售数据未导出，请修改目标文件以保证列数相同")
            print("列数新下载数据文件和目标文件分别为：",data_csv_sales.columns.size,data_sales_CA.columns.size)
              

              
    # CA
    
    elif key[1] in file:
 
    
   
   
        data_csv_sales =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file),encoding="Latin1").assign(日期=os.path.basename(file).split('_')[1])
    #读取源数据加日期 把文件名中的日期写进来
        data_csv_sales['日期'] = pd.to_datetime(data_csv_sales['日期'])
        print(data_csv_sales['日期'])
        data_csv_sales['周数']=""
            
        print(data_csv_sales)
   
       
    
        
   
            
        ru=data_sales_CA.columns.size-data_csv_sales.columns.size 
        
        if ru==0:
        #如果列数相同
            data_csv_sales.columns=salescolumns_CA
            data_sales_CA=data_sales_CA.append(data_csv_sales,ignore_index=True)
        #做append将源数据合并到目标文件
            maxtime=pd.to_datetime(data_sales_CA["日期"].max())
        #查目标文件的最晚日期
            print("最晚时间",maxtime)
            data_sales_CA ['周数']=(maxtime-data_sales_CA['日期']).dt.days//7+1
        #周数写到目标文件
        #在导出之前加周数
            data_sales_CA.to_excel(r'D:\2019plan\Canada周销售数据.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
            print("CA销售数据更新完成")
        else:
            print("CA销售数据未导出，请修改目标文件以保证列数相同")
            print("列数新下载数据文件和目标文件分别为：",data_csv_sales.columns.size,data_sales_CA.columns.size)
              
    # MX

    elif key[2] in file:
        print("开始处理MX数据")
        # 不需要的 data_csv3 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))
    # 打开原文件的dataframe
     

   
        data_csv_sales =pd.read_csv(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file)).assign(日期=os.path.basename(file).split('_')[1])
    #加日期把文件名中的日期写进来
        data_csv_sales['日期'] = pd.to_datetime(data_csv_sales['日期'])
        print(data_csv_sales['日期'])
       
        data_csv_sales['周数']=""
  
    #加周数
        ru=data_csv_sales.columns.size-data_sales_MX.columns.size
        if ru==0:
    #给列名赋值确保可以
            data_csv_sales.columns=salescolumns_MX
    #做append
            data_sales_MX=data_sales_MX.append(data_csv_sales,ignore_index=True)
            maxtime=pd.to_datetime(data_sales_MX["日期"].max())
            print(maxtime)
            data_sales_MX['周数']=(maxtime-data_sales_MX['日期']).dt.days//7+1
            data_sales_MX.to_excel(r'D:\2019plan\Mexico周销售数据.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
            print("MX销售数据更新完成")
        else:
            print("请修改目标文件，以保证列数相同")
            print("列数新下载数据文件和目标文件分别为：",data_csv_sales.columns.size,data_sales_MX.columns.size) 
    else:
        print("什么销售文件都没有")
    

    
    
# 复制TSV在途库存

src_dir_path_shipped=r'D:\运营\计划数据\老站\在途库存'
print(os.listdir(src_dir_path_shipped))


for file in os.listdir(src_dir_path_shipped):
    
    data_shipped_US=pd.read_excel(r'D:\2019plan\在途库存.xlsx')
    data_shipped_CA=pd.read_excel(r'D:\2019plan\Canada在途库存.xlsx')
    data_shipped_MX=pd.read_excel(r'D:\2019plan\Mexico在途库存.xlsx')
    salescolumns_US=data_shipped_US.columns.tolist()
    salescolumns_CA=data_shipped_CA.columns.tolist()
    salescolumns_MX=data_shipped_MX.columns.tolist()
      
    data_tsv5= pd.read_csv(r'D:\\运营\\计划数据\\老站\\在途库存\\'+ str(file),sep='\t',nrows =5)    
    batchnumber= data_tsv5.iat[0,1]
    data_tsv5= pd.read_csv(r'D:\\运营\\计划数据\\老站\\在途库存\\'+ str(file),sep='\t',header=6)    # 读取以分        
    data_tsv5["批次"]=batchnumber
  
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

####################################以下为汇总新站的###############################################
###################################以下为汇总新站的###############################################

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime 

newdate=input('输入最新日期y-m-d',) #输入最新一周的日期
maxtime=datetime.datetime.strptime(newdate,'%Y-%m-%d')
print(maxtime)
maxtimeday=datetime.datetime.strptime(newdate,'%Y-%m-%d').date()
print(maxtimeday)

Product_Analyzepath = r'D:\\运营\\Product_Analyze产品分析\\'
print(Product_Analyzepath)
All_Product_Analyzefile=pd.read_excel(r'D:\\运营\All_Product_Analyzefile.xlsx',sheet_name="sheet1")
for Product_Analyzefile in os.listdir(Product_Analyzepath):
    print(Product_Analyzefile)
#遍历数据文件
    datestr=os.path.basename(Product_Analyzefile).split("_")[2]
    Product_Analyzefile_DF=pd.read_excel(Product_Analyzepath +str(Product_Analyzefile)).assign(日期=datestr[0:10])
    
    All_Product_Analyzefile["周数"]=1
    All_Product_Analyzefile=All_Product_Analyzefile.append(Product_Analyzefile_DF,ignore_index=True)
         
    shutil.move(Product_Analyzepath + str(Product_Analyzefile), 'D:/运营/HistoricalData/Product_Analyzefile')


All_Product_Analyzefile_Weeks=All_Product_Analyzefile[["ASIN","店铺名","站点"]].drop_duplicates()
All_Product_Analyzefile['日期'] = pd.to_datetime(All_Product_Analyzefile['日期'])
print(All_Product_Analyzefile[['日期']])

All_Product_Analyzefile['周数']=(maxtime-All_Product_Analyzefile['日期']).dt.days//7+1
    


max_week=All_Product_Analyzefile["周数"].max()
print(max_week)


for i in range(1,max_week):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
    All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile.loc[(All_Product_Analyzefile["周数"]==i)]

    if i==1:

        All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile_Weeks_i[["ASIN","店铺名","站点",'MSKU',"FBA可售","可售天数预估","标签","销量",'广告点击量','广告花费','广告订单量','毛利润']]
    else:
        All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile_Weeks_i[["ASIN","店铺名","站点","销量",'广告点击量','广告花费','广告订单量','毛利润']]
 
    All_Product_Analyzefile_Weeks_i.rename(columns = {'销量':'销量'+str(i), '广告点击量':'广告点击量'+str(i),'广告花费':'广告花费'+str(i),'广告订单量':'广告订单'+str(i),'毛利润':'毛利润'+str(i)}, inplace = True)

    #合并

    All_Product_Analyzefile_Weeks=pd.merge(All_Product_Analyzefile_Weeks,All_Product_Analyzefile_Weeks_i,on=["ASIN","店铺名","站点"] ,how="left")
    
writer=pd.ExcelWriter(r'D:\\运营\\All_Product_Analyzefile.xlsx')

#All_Product_Analyzefile.to_excel(writer,"All_Product_Analyzefile")
#All_Product_Analyzefile_Weeks.to_excel(writer,"All_Product_Analyzefile_Weeks")

                                 
All_Product_Analyzefile.to_excel(r'D:\\运营\\All_Product_Analyzefile.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\All_Product_Analyzefile_Weeks.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

All_Product_Analyzefile_Weeks=All_Product_Analyzefile_Weeks[["ASIN","店铺名","站点",'MSKU',"FBA可售","可售天数预估","标签","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","广告点击量1","广告点击量2","广告点击量3","广告点击量4","广告点击量5","广告点击量6","广告点击量7","广告点击量8","广告点击量9","广告点击量10","广告花费1", "广告花费2","广告花费3","广告花费4","广告花费5","广告花费6","广告花费7","广告花费8","广告花费9","广告花费10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]

All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\运行结果数据\\All_Product_Analyzefile_Weeks排序.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)






