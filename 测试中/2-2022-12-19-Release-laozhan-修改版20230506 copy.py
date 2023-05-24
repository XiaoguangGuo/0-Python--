# -*- coding:utf-8 –*-
import os
import pandas as pd
from datetime import datetime, timedelta


def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday
#将df中的日期列转换为周数并添加周数列
def update_week_numbers(df):


    last_saturday = find_last_saturday()
    print(last_saturday)

    # 检查输入 DataFrame 的列名中哪一个表示日期
    date_column = "日期" if "日期" in df.columns else "Date"

    df[date_column] = pd.to_datetime(df[date_column])
    df['周数'] = ((last_saturday - df[date_column]).dt.days // 7) + 1
    return df



src_dir_path_inventory=r'D:\运营\\1数据源\\计划数据\老站\当日库存'

key =['US','CA','MX']
t=key[0]

#获取原来库存文件的列名
data_inventory_US=pd.read_excel(r'D:\运营\2019plan\当日Amazon库存.xlsx')
data_inventory_CA=pd.read_excel(r'D:\运营\2019plan\Canada当前Amazon库存.xlsx')
data_inventory_MX=pd.read_excel(r'D:\运营\2019plan\Mexico当日Amazon库存.xlsx')
                                
inventorycolumns_US=data_inventory_US.columns.tolist()
inventorycolumns_CA=data_inventory_CA.columns.tolist()
inventorycolumns_MX=data_inventory_MX.columns.tolist()
                                
print(inventorycolumns_US)

# 在文件夹里查找文件

for file in os.listdir(src_dir_path_inventory):
    print(os.listdir(src_dir_path_inventory))
    
    data_csv = pd.read_csv(r'D:\\运营\\1数据源\\计划数据\\老站\\当日库存\\'+ str(file),encoding="Latin1")    # 读取以encoding='Latin1'分        
    if key[0] in file:
        print(file)
    # 执行语句
        print("有US库存")
       
         # 旧语句data_csv = pd.read_csv(r'D:\\运营\\1数据源\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
                                                 
        data_csv.columns=inventorycolumns_US                      
        data_csv.to_excel(r'D:\运营\2019plan\当日Amazon库存.xlsx',sheet_name="当前Amazon库存",startrow=0,header=True,index=False)
        
    elif key[1]in file:
        print("有CA库存")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\1数据源\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        data_csv.columns=inventorycolumns_CA    
                         
        data_csv.to_excel(r'D:\运营\2019plan\Canada当前Amazon库存.xlsx', sheet_name="15828640259018099",startrow=0,header=True,index=False)

        
      
    elif key[2]in file:
        print("有MX库存")
        print(file)        
        #df_data.columns.tolist())
        
        data_csv.columns=inventorycolumns_MX  
                         
        data_csv.to_excel(r'D:\运营\2019plan\Mexico当日Amazon库存.xlsx', sheet_name="当前Amazon库存",startrow=0,header=True,index=False)

        print(data_csv)
        
    else:
        print("什么库存文件都没有")

# 导入reStock

src_dir_path_restock=r'D:\运营\1数据源\计划数据\老站\restock'
print(os.listdir(src_dir_path_restock))
for file in os.listdir(src_dir_path_restock):
    data_csv2 = pd.read_table(r'D:\\运营\\1数据源\\计划数据\\老站\\restock\\'+ str(file),encoding="Latin1")    # 读取以分        
    if key[0] in file:
        print(file)
    # 执行语句
        print("有USrestock")
       
         # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        print(data_csv2)
                     
        data_csv2.to_excel(r'D:\运营\2019plan\restock-report.xlsx',sheet_name="restock-report",startrow=0,header=True,index=False)
        
    elif key[1]in file:
        print("有CArestock")
        print(file)
        # 旧语句data_csv = pd.read_csv(r'D:\\运营\\计划数据\\老站\\当日库存\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        
        #df_data.columns.tolist())
        
                         
        data_csv2.to_excel(r'D:\运营\2019plan\restock-report_CA.xlsx', sheet_name="REstock-CA",startrow=0,header=True,index=False)

        print(data_csv2)
      
    elif key[2]in file:
        print("有MXrestock")
        print(file)        
        #df_data.columns.tolist())
   
        data_csv2.to_excel(r'D:\运营\2019plan\restock-report_MX.xlsx', sheet_name="restock-report_MX",startrow=0,header=True,index=False)

        print(data_csv2)
        
    else:
        print("什么restock文件都没有")

#复制销售数据 20210221模块待写入
src_dir_path_sales=r'D:\运营\1数据源\计划数据\老站\销售数据'
# 设置来源文件搜索目录
print(os.listdir(src_dir_path_sales))
key =['US','CA','MX']
#设置需要搜索的国家名字


#读取D:\运营\2生成过程表\周销售数据总表.xlsx"


def process_sales_data(Country, target_file_path, src_dir_path_sales):
    print(f"开始处理 {Country} 数据")

    # 读取目标文件
    target_data = pd.read_excel(target_file_path)
    
    # 获取目标文件列名
    target_columns = target_data.columns.tolist()

    for file in os.listdir(src_dir_path_sales):
        file_first_part = file.split("_")[0]

      

        found_keyword=False
        if Country in file_first_part:
            found_keyword=True
            file_path = os.path.join(src_dir_path_sales, file)
            data_csv_sales = pd.read_csv(file_path).assign(日期=os.path.basename(file).split('_')[1])
            data_csv_sales['日期'] = pd.to_datetime(data_csv_sales['日期'])
            data_csv_sales['周数'] = ""
            target_data = pd.concat([target_data, data_csv_sales], ignore_index=True)

            maxtime = find_last_saturday()
            target_data['周数'] = (maxtime - target_data['日期']).dt.days // 7 + 1
            target_data.to_excel(target_file_path, sheet_name="Sheet1", startrow=0, header=True, index=False)
            print(f"{Country} 销售数据更新完成")
            break
        
    if not found_keyword:
        print(f"没找到：{Country} 销售数据")
       

src_dir_path_sales = r'D:\\运营\\1数据源\\计划数据\\老站\\销售数据\\'
process_sales_data("US", r"D:\运营\2019plan\周销售数据.xlsx", src_dir_path_sales)
process_sales_data("CA", r"D:\运营\2019plan\Canada周销售数据.xlsx", src_dir_path_sales)
process_sales_data("MX", r"D:\运营\2019plan\Mexico周销售数据.xlsx", src_dir_path_sales)   




all_sales_df=pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx',sheet_name="Sheet1")



for file in os.listdir(src_dir_path_sales):
        file_first_part = file.split("_")[0]
        for keyword in key:
            if keyword in file_first_part :

                file_path = os.path.join(src_dir_path_sales, file)
                data_csv_sales = pd.read_csv(file_path).assign(日期=os.path.basename(file).split('_')[1])
                data_csv_sales['日期'] = pd.to_datetime(data_csv_sales['日期'])
                data_csv_sales["Country"]="GV"+"-"+str(keyword)
                if keyword=="CA":
                    data_csv_sales=data_csv_sales.rename(columns = {
    "Sessions – Total – B2B": "Sessions - Total - B2B",
    "Session percentage - Total": "Session Percentage - Total",
    "Session Percentage – Total – B2B": "Session Percentage - Total - B2B",
    "Page views - Total": "Page Views - Total",
    "Page Views – Total – B2B": "Page Views - Total - B2B",
    "Page views percentage - Total": "Page Views Percentage - Total",
    "Page Views Percentage – Total – B2B": "Page Views Percentage - Total - B2B",
    "Featured Offer (Buy Box) Percentage – B2B": "Featured Offer (Buy Box) Percentage - B2B",
    "Units ordered": "Units Ordered",
    "Units ordered – B2B": "Units Ordered - B2B",
    "Unit session percentage": "Unit Session Percentage",
    "Units session percentage – B2B": "Unit Session Percentage - B2B",
    "Ordered product sales": "Ordered Product Sales",
    "Ordered product sales – B2B": "Ordered Product Sales - B2B",
    "Total order items": "Total Order Items",
    "Total order items – B2B": "Total Order Items - B2B"})

                #judge if the columns of data_csv_sales is a subset of all_sales_df
                if set(data_csv_sales.columns).issubset(set(all_sales_df.columns)):

                    all_sales_df = all_sales_df.reset_index(drop=True)
                    data_csv_sales = data_csv_sales.reset_index(drop=True)
                    print(data_csv_sales.columns)
                    input("请检查列名是否正确，按回车键继续")
                    all_sales_df = pd.concat([all_sales_df,data_csv_sales],axis=0,ignore_index=True)
                    break
                else:
                    print("列名有不符合的，需要修改")
                    input("请修改列名后按回车键继续")
                    
all_sales_df=update_week_numbers(all_sales_df)
print(all_sales_df)        
all_sales_df.to_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx', sheet_name="Sheet1", startrow=0, header=True, index=False)
 
# 复制TSV在途库存

src_dir_path_shipped=r'D:\运营\1数据源\计划数据\老站\在途库存'
print(os.listdir(src_dir_path_shipped))


for file in os.listdir(src_dir_path_shipped):
    
    data_shipped_US=pd.read_excel(r'D:\运营\2019plan\在途库存.xlsx')
    data_shipped_CA=pd.read_excel(r'D:\运营\2019plan\Canada在途库存.xlsx')
    data_shipped_MX=pd.read_excel(r'D:\运营\2019plan\Mexico在途库存.xlsx')
    salescolumns_US=data_shipped_US.columns.tolist()
    salescolumns_CA=data_shipped_CA.columns.tolist()
    salescolumns_MX=data_shipped_MX.columns.tolist()
      
    data_tsv5= pd.read_csv(r'D:\\运营\\1数据源\\计划数据\\老站\\在途库存\\'+ str(file),sep='\t',nrows =5)    
    batchnumber= data_tsv5.iat[0,1]
    data_tsv5= pd.read_csv(r'D:\\运营\\1数据源\\计划数据\\老站\\在途库存\\'+ str(file),sep='\t',header=6)    # 读取以分        
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

        data_shipped_US.to_excel(r'D:\运营\2019plan\在途库存.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
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

        data_shipped_CA.to_excel(r'D:\运营\2019plan\Canada在途库存.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
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

        data_shipped_MX.to_excel(r'D:\运营\2019plan\Mexico在途库存.xlsx', sheet_name="Sheet1",startrow=0,header=True,index=False)
        print("MX在途更新完成")

    
    else:
        print("什么在途文件都没有")


# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime 



input ("首先檢查三個在途計劃表,確認沒問題之後點回車")

##############################################################################################################################################################################

#讀取2019計劃的銷售表
Sales_US=pd.read_excel(r'D:\运营\2019plan\周销售数据.xlsx')

Sales_US["Country"]="GV-US"
Sales_US.rename(columns = {'(Child) ASIN':"Asin"}, inplace = True)
                
Stock_US=pd.read_excel(r'D:\运营\2019plan\当日Amazon库存.xlsx')
Stock_US.rename(columns = {'sku':"SKU",'asin':"Asin","afn-fulfillable-quantity":"Fufillable","afn-inbound-receiving-quantity":"Receiving","afn-reserved-quantity":"Reserved"}, inplace = True)
Stock_US["Country"]="GV-US"
Stock_US=Stock_US[["Country","Asin","SKU","Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]]


Intransit_us=pd.read_excel(r'D:\运营\2019plan\在途库存.xlsx')

Intransit_us["Country"]="GV-US"
 
Intransit_us.rename(columns = {'Merchant SKU':"SKU",'ASIN':"Asin"}, inplace = True)


#输出成excel表

Sales_CA=pd.read_excel(r'D:\运营\2019plan\Canada周销售数据.xlsx')
Sales_CA["Country"]="GV-CA"
Sales_CA.rename(columns = {'(Child) ASIN':"Asin","Units ordered":"Units Ordered"}, inplace = True)

                
Stock_CA=pd.read_excel(r'D:\运营\2019plan\Canada当前Amazon库存.xlsx')
Stock_CA["Country"]="GV-CA"
Stock_CA.rename(columns = {'sku':"SKU",'asin':"Asin","afn-fulfillable-quantity":"Fufillable","afn-inbound-receiving-quantity":"Receiving","afn-reserved-quantity":"Reserved"}, inplace = True)
Stock_CA=Stock_CA[["Country","Asin","SKU","Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]]


Intransit_ca=pd.read_excel(r'D:\运营\2019plan\Canada在途库存.xlsx')
Intransit_ca["Country"]="GV-CA"
Intransit_ca.rename(columns = {'Merchant SKU':"SKU",'ASIN':"Asin"}, inplace = True)

         
Sales_MX=pd.read_excel(r'D:\运营\2019plan\Mexico周销售数据.xlsx')
Sales_MX["Country"]="GV-MX"
Stock_MX=pd.read_excel(r'D:\运营\2019plan\Mexico当日Amazon库存.xlsx')
Stock_MX.rename(columns = {'sku':"SKU",'asin':"Asin"}, inplace = True)
Stock_MX.rename(columns = {'sku':"SKU",'asin':"Asin","afn-fulfillable-quantity":"Fufillable","afn-inbound-receiving-quantity":"Receiving","afn-reserved-quantity":"Reserved"}, inplace = True)
Stock_MX["Country"]="GV-MX"
Stock_MX=Stock_MX[["Country","Asin","SKU","Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]]                






Intransit_mx=pd.read_excel(r'D:\运营\2019plan\Mexico在途库存.xlsx')
Intransit_mx["Country"]="GV-MX"

                
Intransit_mx.rename(columns = {'Merchant SKU':"SKU",'ASIN':"Asin"}, inplace = True) 



Sales_All=pd.concat([Sales_US,Sales_CA,Sales_MX])
Sales_All.to_excel(r'D:\运营\2生成过程表\2023plan\Sales_all.xlsx')

Stock_All=pd.concat([Stock_US,Stock_CA,Stock_MX])
Intransit_All=pd.concat([Intransit_us,Intransit_ca,Intransit_mx])       

 
SKUAll_1=Stock_All[["Country","Asin","SKU"]].drop_duplicates()
SKUAll_2=Sales_All[["Country","Asin","SKU"]].drop_duplicates()
SKUAll=pd.concat([SKUAll_1,SKUAll_2])


max_week=100

Sales_Weeks=SKUAll_2

for i in range(1,max_week):
    
    Sales_Weeks_i=Sales_All.loc[(Sales_All["周数"]==i)]
   

    if i==1:

        Sales_Weeks_i=Sales_Weeks_i[["Country","Asin","Title","SKU","Units Ordered","Sessions - Total","Unit Session Percentage"]]

        Sales_Weeks_i.rename(columns = {"Units Ordered":str(i),"Sessions - Total":"Session"+str(i),"Unit Session Percentage":"Percentage"+str(i)}, inplace = True)

        print(Sales_Weeks_i)
       
    else:
        Sales_Weeks_i=Sales_Weeks_i[["Country","Asin","SKU","Units Ordered","Sessions - Total","Unit Session Percentage"]]
        print(Sales_Weeks_i)
       
        print(i)
        Sales_Weeks_i.rename(columns = {"Units Ordered":str(i),"Sessions - Total":"Session"+str(i),"Unit Session Percentage":"Percentage"+str(i)}, inplace = True)
    

    #合并

    Sales_Weeks=pd.merge(Sales_Weeks,Sales_Weeks_i,on=["Country","Asin","SKU"] ,how="left")
    Sales_Weeks.to_excel(r'D:\运营\2生成过程表\2023plan\Sales_Weeks.xlsx' ,index=False)  

max_week=11
Intransit_Weeks = Intransit_All[["Country","Asin","SKU"]].drop_duplicates()
for i in range(1,max_week):
  Intransit_All2=Intransit_All.groupby(["Country","Asin","SKU","周数"],as_index=False)[['Shipped']].agg('sum')
  Intransit_Weeks_i=Intransit_All2.loc[Intransit_All2["周数"]==i]
  if len(Intransit_Weeks_i)>0:


      
      Intransit_Weeks_i=Intransit_Weeks_i[["Country","Asin","SKU","Shipped"]]
      Intransit_Weeks_i.rename(columns = {"Shipped":"第"+str(i)+"周入库"}, inplace = True)
  Intransit_Weeks =pd.merge(Intransit_Weeks,Intransit_Weeks_i,on=["Country","Asin","SKU"] ,how="left")
  




PlanAll=pd.merge(SKUAll,Sales_Weeks,how="left", on=["Country","SKU","Asin"])

PlanAll.to_excel(r'D:\运营\2生成过程表\2023plan\PlanAllnew.xlsx' ,index=False)


PlanAll=pd.merge(PlanAll,Stock_All,how="left", on=["Country","SKU","Asin"])

PlanAll=pd.merge(PlanAll,Intransit_Weeks,how="left", on=["Country","SKU","Asin"])

PlanAll.fillna(0,inplace=True)
Listing=pd.read_excel(r'D:\运营\2019plan\Listing.xlsx',sheet_name="Listing")
Listing=Listing[["Country","SKU","大类","小类"]]

Price=pd.read_excel(r'D:\运营\2019plan\Listing.xlsx',sheet_name="Price")
Price=Price[["SKU","Price"]]
PlanAll=pd.merge(PlanAll,Listing,on=["Country","SKU" ] ,how="left")
PlanAll=pd.merge(PlanAll,Price,on=["SKU" ] ,how="left")
print(PlanAll)
    
    

WeekSalesIndex_Dic={"1st":0.2,"2nd":0.2,"3rd":0.1,"4th":0.1,"5th":0.1,"6th":0.1,"7th":0.1,"8th":0.1}

WeekSales=WeekSalesIndex_Dic["1st"]*PlanAll["1"]+WeekSalesIndex_Dic["2nd"]*PlanAll["2"]+WeekSalesIndex_Dic["3rd"]*PlanAll["3"]+WeekSalesIndex_Dic["4th"]*PlanAll["4"]+WeekSalesIndex_Dic["5th"]*PlanAll["5"]+WeekSalesIndex_Dic["6th"]*PlanAll["6"]+WeekSalesIndex_Dic["7th"]*PlanAll["7"]+WeekSalesIndex_Dic["8th"]*PlanAll["8"]



PlanAll["SELLING10"]=PlanAll["1"]+PlanAll["2"]+PlanAll["3"]+PlanAll["4"]+PlanAll["5"]+PlanAll["6"]+PlanAll["7"]+PlanAll["8"]+PlanAll["9"]+PlanAll["10"]
PlanAll["STOCKALL"]=PlanAll["Fufillable"]*1+PlanAll["Receiving"]*1+PlanAll["Reserved"]*1+PlanAll["afn-inbound-shipped-quantity"]
PlanAll["TotalAmount"]=PlanAll["STOCKALL"]*PlanAll["Price"]
PlanAll["Zhouzhuan10"]=10*PlanAll["STOCKALL"]/PlanAll["SELLING10"]
PlanAll["ZZ1"]=PlanAll["1"]-PlanAll["2"]

PlanAll["ZZ2"]=(PlanAll["1"]+PlanAll["2"]-PlanAll["3"]-PlanAll["4"])/2

PlanAll["For第2周销售的到货需求"]=WeekSales*2-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["Reserved"]

PlanAll["For第3周销售的到货需求"]=WeekSales*3-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]

PlanAll["For第4周销售的到货需求"]=WeekSales*4-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]

PlanAll["For第5周销售的到货需求"]=WeekSales*5-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]


PlanAll["For第6周销售的到货需求"]=WeekSales*6-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]

PlanAll["For第7周销售的到货需求"]=WeekSales*7-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]
PlanAll["For第8周销售的到货需求"]=WeekSales*8-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]
PlanAll["For第9周销售的到货需求"]=WeekSales*9-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]-PlanAll["第8周入库"] 

PlanAll["For第10周销售的到货需求"]=WeekSales*10-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]-PlanAll["第8周入库"]-PlanAll["第9周入库"] 
PlanAll["For第11周销售的到货需求"]=WeekSales*11-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]-PlanAll["第8周入库"]-PlanAll["第9周入库"]-PlanAll["第10周入库"]  
PlanAll["For第12周销售的到货需求"]=WeekSales*12-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]-PlanAll["第8周入库"]-PlanAll["第9周入库"]-PlanAll["第10周入库"] 
PlanAll["For第13周销售的到货需求"]=WeekSales*13-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]-PlanAll["第8周入库"]-PlanAll["第9周入库"]-PlanAll["第10周入库"]  
PlanAll["For第14周销售的到货需求"]=WeekSales*14-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]-PlanAll["第8周入库"]-PlanAll["第9周入库"]-PlanAll["第10周入库"]  

PlanAll["For第15周销售的到货需求"]=WeekSales*15-PlanAll["Fufillable"]-PlanAll["Receiving"]-PlanAll["第1周入库"]-PlanAll["第2周入库"]-PlanAll["Reserved"]-PlanAll["第3周入库"]-PlanAll["第4周入库"]-PlanAll["第5周入库"]-PlanAll["第6周入库"]-PlanAll["第7周入库"]-PlanAll["第8周入库"]-PlanAll["第9周入库"]-PlanAll["第10周入库"] 


PlanAll["Adjusted-Week2"]=PlanAll["ZZ2"]*0.7*2+PlanAll["For第2周销售的到货需求"]  
PlanAll["Adjusted-Week3"]=PlanAll["ZZ2"]*0.7*3+PlanAll["For第3周销售的到货需求"]

PlanAll["Adjusted-Week4"]=PlanAll["ZZ2"]*0.7*4+PlanAll["For第4周销售的到货需求"]

PlanAll["Adjusted-Week5"]=PlanAll["ZZ2"]*0.7*5+PlanAll["For第5周销售的到货需求"]

PlanAll["Adjusted-Week6"]=PlanAll["ZZ2"]*0.7*6+PlanAll["For第6周销售的到货需求"]

PlanAll["Adjusted-Week7"]=PlanAll["ZZ2"]*0.7*7+PlanAll["For第7周销售的到货需求"]
PlanAll["Adjusted-Week8"]=PlanAll["ZZ2"]*0.7*8+PlanAll["For第8周销售的到货需求"]
PlanAll["Adjusted-Week9"]=PlanAll["ZZ2"]*0.7*9+PlanAll["For第9周销售的到货需求"]

PlanAll["Adjusted-Week10"]=PlanAll["ZZ2"]*0.7*10+PlanAll["For第10周销售的到货需求"]

PlanAll["Adjusted-Week11"]=PlanAll["ZZ2"]*0.7*11+PlanAll["For第11周销售的到货需求"]
PlanAll["Adjusted-Week12"]=PlanAll["ZZ2"]*0.7*12+PlanAll["For第12周销售的到货需求"]
PlanAll["Adjusted-Week13"]=PlanAll["ZZ2"]*0.7*13+PlanAll["For第13周销售的到货需求"]
PlanAll["Adjusted-Week14"]=PlanAll["ZZ2"]*0.7*14+PlanAll["For第14周销售的到货需求"]
PlanAll["Adjusted-Week15"]=PlanAll["ZZ2"]*0.7*15+PlanAll["For第15周销售的到货需求"]




#SELECT "US" AS Country, 周销售数据_交叉表_SKU日期.SKU, 周销售数据_交叉表_SKU日期.[(Child)
#ASIN], listing.大类, listing.小类, listing.新品, listing.型号, listing.唯一中文名称,
#周销售数据_交叉表_SKU日期.Title之Last, 周销售数据_交叉表_SKU日期.[总计 Units Ordered],
#IIF([1]>0,[周Bulk广告数据汇总-US_交叉表加名字].广告1/[1],null) AS BILI1,
#([1]+[2]+[3]+[4]+[5]+[6]+[7]+[8]+[9]+[10]) AS SELLING10,
#([Fufillable]*1+[Receiving]*1+[Reserved]*1+[afn-inbound-shipped-quantity]*1)
#AS STOCKALL, IIF(SELLING10>0,(STOCKALL*10/SELLING10),Null) AS
#Zhouzhuan10, Productprice.Price,
#([Fufillable]*1+[Receiving]*1+[Reserved]*1+[afn-inbound-shipped-quantity]*1)*[Productprice]![Price]
#AS TotalAmount,
#([周Bulk广告数据汇总-US_交叉表加名字].广告1-[周Bulk广告数据汇总-US_交叉表加名字].广告2) AS GGZZ1,
#([1]-[2]) AS ZZ1, ([1]+[2]-[3]-[4])/2 AS ZZ2, IIf([周销售数据_交叉表_SKU日期].[1]
#Is Null,0,[周销售数据_交叉表_SKU日期].[1]) AS 1, IIf([周销售数据_交叉表_SKU日期].[2] Is
#Null,0,[周销售数据_交叉表_SKU日期].[2]) AS 2, IIf([周销售数据_交叉表_SKU日期].[3] Is
#Null,0,[周销售数据_交叉表_SKU日期].[3]) AS 3, IIf([周销售数据_交叉表_SKU日期].[4] Is
#Null,0,[周销售数据_交叉表_SKU日期].[4]) AS 4, IIf([周销售数据_交叉表_SKU日期].[5] Is
#Null,0,[周销售数据_交叉表_SKU日期].[5]) AS 5, IIf([周销售数据_交叉表_SKU日期].[6] Is
#Null,0,[周销售数据_交叉表_SKU日期].[6]) AS 6, IIf([周销售数据_交叉表_SKU日期].[7] Is
#Null,0,[周销售数据_交叉表_SKU日期].[7]) AS 7, IIf([周销售数据_交叉表_SKU日期].[8] Is
#Null,0,[周销售数据_交叉表_SKU日期].[8]) AS 8, IIf([周销售数据_交叉表_SKU日期].[9] Is
#Null,0,[周销售数据_交叉表_SKU日期].[9]) AS 9, IIf([周销售数据_交叉表_SKU日期].[10] Is
#Null,0,[周销售数据_交叉表_SKU日期].[10]) AS 10, [周Bulk广告数据汇总-US_交叉表加名字].广告1,
#[周Bulk广告数据汇总-US_交叉表加名字].广告2, [周Bulk广告数据汇总-US_交叉表加名字].广告3,
#[周Bulk广告数据汇总-US_交叉表加名字].广告4, [周Bulk广告数据汇总-US_交叉表加名字].广告5,
#[周Bulk广告数据汇总-US_交叉表加名字].广告6, [周Bulk广告数据汇总-US_交叉表加名字].广告7,
#[周Bulk广告数据汇总-US_交叉表加名字].广告8, [周Bulk广告数据汇总-US_交叉表加名字].广告9,
#[周Bulk广告数据汇总-US_交叉表加名字].广告10,
#([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4 AS 加权周平均销量,
#Nz([当日Amazon库存]![afn-fulfillable-quantity],0) AS Fufillable,
#Nz([当日Amazon库存]![afn-reserved-quantity],0) AS Reserved,
#当日Amazon库存.[afn-inbound-working-quantity],
#当日Amazon库存.[afn-inbound-shipped-quantity],
#Nz([当日Amazon库存]![afn-inbound-receiving-quantity],0) AS Receiving,
#Nz([在途库存_交叉表]![1],0) AS 第1周入库, Nz([在途库存_交叉表]![2],0) AS 第2周入库,
#Nz([在途库存_交叉表]![3],0) AS 第3周入库, Nz([在途库存_交叉表]![4],0) AS 第4周入库,
#Nz([在途库存_交叉表]![5],0) AS 第5周入库, Nz([在途库存_交叉表]![6],0) AS 第6周入库,
#Nz([在途库存_交叉表]![7],0) AS 第7周入库, Nz([在途库存_交叉表]![8],0) AS 第8周入库,
#Nz([在途库存_交叉表]![9],0) AS 第9周入库, Nz([在途库存_交叉表]![10],0) AS 第10周入库,
#Nz([在途库存_交叉表]![11],0) AS 第11周入库, Nz([在途库存_交叉表]![12],0) AS 第12周入库,
#Nz([在途库存_交叉表]![13],0) AS 第13周入库, Nz([在途库存_交叉表]![14],0) AS 第14周入库,
#Nz([在途库存_交叉表]![15],0) AS 第15周入库,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*2-[Fufillable]-[Receiving]-[第1周入库]-[Reserved]
#AS For第2周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*3-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[Reserved]
#AS For第3周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*4-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[Reserved]
#AS For第4周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*5-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[Reserved]
#AS For第5周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*6-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[Reserved]
#AS For第6周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*7-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]
#AS For第7周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*8-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]
#AS For第8周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*9-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]
#AS For第9周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*10-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]-[第9周入库]
#AS For第10周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*11-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]-[第9周入库]-[第10周入库]
#AS For第11周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*12-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]-[第9周入库]-[第10周入库]-[第11周入库]
#AS For第12周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*13-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]-[第9周入库]-[第10周入库]-[第11周入库]-[第12周入库]
#AS For第13周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*14-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]-[第9周入库]-[第10周入库]-[第11周入库]-[第12周入库]-[第13周入库]
#AS For第14周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*15-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]-[第9周入库]-[第10周入库]-[第11周入库]-[第12周入库]-[第13周入库]-[第14周入库]
#AS For第15周销售的到货需求,
#(([1]+[2]+[3]+[4])*0.6/4+([5]+[6]+[7]+[8])*0.4/4)*20-[Fufillable]-[Receiving]-[第1周入库]-[第2周入库]-[第3周入库]-[第4周入库]-[第5周入库]-[第6周入库]-[Reserved]-[第7周入库]-[第8周入库]-[第9周入库]-[第10周入库]-[第11周入库]-[第12周入库]-[第13周入库]-[第14周入库]
#AS For第20周销售的到货需求, [ZZ2]*0.7*2+[For第2周销售的到货需求] AS [Adjusted-Week2],
#[ZZ2]*0.7*3+[For第3周销售的到货需求] AS [Adjusted-Week3],
#[ZZ2]*0.7*4+[For第4周销售的到货需求] AS [Adjusted-Week4],
#[ZZ2]*0.7*5+[For第5周销售的到货需求] AS [Adjusted-Week5],
#[ZZ2]*0.7*2+[For第6周销售的到货需求] AS [Adjusted-Week6],
#[ZZ2]*0.7*7+[For第7周销售的到货需求] AS [Adjusted-Week7],
#[ZZ2]*0.7*8+[For第8周销售的到货需求] AS [Adjusted-Week8],
#[ZZ2]*0.7*9+[For第9周销售的到货需求] AS [Adjusted-Week9],
#[ZZ2]*0.7*10+[For第10周销售的到货需求] AS [Adjusted-Week10],
#[ZZ2]*0.7*11+[For第11周销售的到货需求] AS [Adjusted-Week11],
#[ZZ2]*0.7*12+[For第12周销售的到货需求] AS [Adjusted-Week12],
#[ZZ2]*0.7*13+[For第13周销售的到货需求] AS [Adjusted-Week13],
#[ZZ2]*0.7*14+[For第14周销售的到货需求] AS [Adjusted-Week14],
#[ZZ2]*0.7*15+[For第15周销售的到货需求] AS [Adjusted-Week15],
#[ZZ2]*0.7*20+[For第20周销售的到货需求] AS [Adjusted-Week20]
#FROM ((((周销售数据_交叉表_SKU日期 LEFT JOIN 当日Amazon库存 ON 周销售数据_交叉表_SKU日期.SKU=当日Amazon库存.sku) LEFT JOIN listing ON 周销售数据_交叉表_SKU日期.SKU=listing.[seller-sku]) LEFT JOIN 在途库存_交叉表 ON 周销售数据_交叉表_SKU日期.SKU=在途库存_交叉表.[Merchant SKU]) LEFT JOIN [周Bulk广告数据汇总-US_交叉表加名字] ON 周销售数据_交叉表_SKU日期.SKU=[周Bulk广告数据汇总-US_交叉表加名字].SKU之合计) LEFT JOIN Productprice ON 周销售数据_交叉表_SKU日期.SKU=Productprice.SKU;



CampaignSKU_Summary=pd.read_excel(r'D:\运营\2生成过程表\周bulk数据Summary.xlsx',sheet_name="SKU-WEEK")
CampaignSKU_Summary.rename(columns = {'Country':'Country'}, inplace = True)






CampaignSKU_Summary_biaotou=CampaignSKU_Summary[["Country","SKU"]].drop_duplicates()
print(CampaignSKU_Summary_biaotou)

for i in range(1,11):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
    CampaignSKU_Summary_i=CampaignSKU_Summary.loc[(CampaignSKU_Summary["周数"]==i)]
    
    CampaignSKU_Summary_i=CampaignSKU_Summary_i[["Country","SKU","Clicks","Orders","Spend"]]
    #更改列名

    CampaignSKU_Summary_i.rename(columns = {'Clicks':'广告Clicks'+str(i), 'Orders':'广告Orders'+str(i),'Spend':'广告'+str(i)}, inplace = True)

    PlanAll=pd.merge(PlanAll,CampaignSKU_Summary_i,on=["Country","SKU" ] ,how="left")
 
PlanAll["GGZZ1"]=PlanAll["广告1"]-PlanAll["广告2"]
PlanAll["BILI1"]=PlanAll["广告1"]/PlanAll["1"]
PlanAll=PlanAll.drop_duplicates()
PlanAll=PlanAll[["Country","SKU","Asin","Title","大类","小类","Price",	
                 "SELLING10","STOCKALL","TotalAmount","Zhouzhuan10","GGZZ1","BILI1",
                 "ZZ1","ZZ2","1","2","3","4","5","6","7","8","9","10",
                 "广告1","广告2","广告3","广告4","广告5","广告6","广告7","广告8",
                 "广告9","广告10","Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving","第1周入库","第2周入库","第3周入库","第4周入库","第5周入库","第6周入库","第7周入库","第8周入库","第9周入库","第10周入库","For第2周销售的到货需求","For第3周销售的到货需求","For第4周销售的到货需求","For第5周销售的到货需求","For第6周销售的到货需求","For第7周销售的到货需求","For第8周销售的到货需求","For第9周销售的到货需求",
                 "For第10周销售的到货需求","For第11周销售的到货需求","For第12周销售的到货需求","For第13周销售的到货需求","For第14周销售的到货需求","For第15周销售的到货需求","Adjusted-Week2","Adjusted-Week3","Adjusted-Week4","Adjusted-Week5","Adjusted-Week6","Adjusted-Week7","Adjusted-Week8","Adjusted-Week9","Adjusted-Week10","Adjusted-Week11","Adjusted-Week12","Adjusted-Week13","Adjusted-Week14","Adjusted-Week15","广告Clicks1","广告Orders1"	,"广告Clicks2","广告Orders2","广告Clicks3","广告Orders3","广告Clicks4","广告Orders4","广告Clicks5","广告Orders5","广告Clicks6","广告Orders6","广告Clicks7","广告Orders7","广告Clicks8","广告Orders8","广告Clicks9","广告Orders9","广告Clicks10","广告Orders10"]]

PlanAll.to_excel(r'D:\运营\2生成过程表\2023plan\plan.xlsx' ,index=False)       



