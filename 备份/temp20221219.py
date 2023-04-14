
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

Product_Analyzepath = r'D:\运营\1数据源\Product_Analyze产品分析\\'
print(Product_Analyzepath)
All_Product_Analyzefile=pd.read_excel(r'D:\\运营\2生成过程表\\All_Product_Analyzefile.xlsx',sheet_name="sheet1")
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
    
writer=pd.ExcelWriter(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile.xlsx')

#All_Product_Analyzefile.to_excel(writer,"All_Product_Analyzefile")
#All_Product_Analyzefile_Weeks.to_excel(writer,"All_Product_Analyzefile_Weeks")

                                 
All_Product_Analyzefile.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile_Weeks.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

All_Product_Analyzefile_Weeks=All_Product_Analyzefile_Weeks[["ASIN","店铺名","站点",'MSKU',"FBA可售","可售天数预估","标签","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","广告点击量1","广告点击量2","广告点击量3","广告点击量4","广告点击量5","广告点击量6","广告点击量7","广告点击量8","广告点击量9","广告点击量10","广告花费1", "广告花费2","广告花费3","广告花费4","广告花费5","广告花费6","广告花费7","广告花费8","广告花费9","广告花费10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]

All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile_Weeks排序.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)


