


# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime



USSellingData=pd.read_excel(r'D:\运营\2019计划\周销售数据.xlsx)
    
USSellingData["站点"]=US


max_week=10


for i in range(1,10):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
                            
    USSellingDataWeeks_i=USSellingData.loc[(All_Product_Analyzefile["周数"]==i)]

    if i==1:

        USSellingDataWeeks_i=USSellingDataWeeks_i[["站点","SKU","Units Ordered","Ordered Product Sales"]]
    else:
        All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile_Weeks_i[["ASIN","店铺名","站点","销量",'销售额','广告点击量','广告花费','广告订单量','毛利润']]
 
    All_Product_Analyzefile_Weeks_i.rename(columns = {'销量':'销量'+str(i), '销售额':'销售额'+str(i),'广告点击量':'广告点击量'+str(i),'广告花费':'广告'+str(i),'广告订单量':'广告订单'+str(i),'毛利润':'毛利润'+str(i)}, inplace = True)
    All_Product_Analyzefile_biaotou=pd.merge(All_Product_Analyzefile_biaotou,All_Product_Analyzefile_Weeks_i,on=["站点","ASIN"] ,how="left")


                            
price=pd.read_excel(r'D:\运营\数据源\Listing.xlsx',sheet_name="price")
USSellingData=pd.merge(USSellingData,price,on=["SKU"] ,how="left")

                            
All_Product_Analyzefile=pd.read_excel(r'D:\运营\运行结果数据\All_Product_Analyzefile.xlsx',sheet_name="sheet1")

#####：以下为草稿
#################am_gv=plan["销售额"].sum()am_sailingstar=All_Product_Analyzefile[[""销售额"],"周数"=1].sum()
#AM=am_gv+am_sailingstar


#Plan_gv_group=plan.groupby(["Country"],"xiaoshou").agg.("sum")
#All_Product_Analyzefile_group=All_Product_Analyzefile.(["Country"],"xiaoshou").agg.("sum")



#####：草稿结束


#筛选出老站计划近10周的销售额，订单数和广告额，广告订单数
plan=plan.drop_duplicates()
plan["销售额1"]=0.00001
plan["广告订单1"]=0



plan.rename(columns = {'COUNTRY':'站点','1':'销量1','2':'销量2','3':'销量3','4':'销量4','4':'销量4','5':'销量5','6':'销量6','7':'销量7','8':'销量8','9':'销量9','10':'销量10'},inplace=True)


print(plan)





plan=plan.groupby("站点").agg("sum")



#筛选出新站计划的10周内的数据。
All_Product_Analyzefile=All_Product_Analyzefile[All_Product_Analyzefile["周数"]<11]


#变成横的周数


max_week=All_Product_Analyzefile["周数"].max()+1
All_Product_Analyzefile_biaotou=All_Product_Analyzefile[["站点","ASIN"]].drop_duplicates()
print(max_week)


for i in range(1,max_week):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
    All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile.loc[(All_Product_Analyzefile["周数"]==i)]

    if i==1:

        All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile_Weeks_i[["ASIN","店铺名","站点",'MSKU',"FBA可售","可售天数预估","标签","销量",'销售额','广告点击量','广告花费','广告订单量','毛利润']]
    else:
        All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile_Weeks_i[["ASIN","店铺名","站点","销量",'销售额','广告点击量','广告花费','广告订单量','毛利润']]
 
    All_Product_Analyzefile_Weeks_i.rename(columns = {'销量':'销量'+str(i), '销售额':'销售额'+str(i),'广告点击量':'广告点击量'+str(i),'广告花费':'广告'+str(i),'广告订单量':'广告订单'+str(i),'毛利润':'毛利润'+str(i)}, inplace = True)
    All_Product_Analyzefile_biaotou=pd.merge(All_Product_Analyzefile_biaotou,All_Product_Analyzefile_Weeks_i,on=["站点","ASIN"] ,how="left")


    #合并
    print("??", All_Product_Analyzefile_biaotou)
All_Product_Analyzefile_biaotou=All_Product_Analyzefile_biaotou[["站点","毛利润1","毛利润2","毛利润3","毛利润4","毛利润5","毛利润6","毛利润7","毛利润8","毛利润9","毛利润10","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","销售额1","销售额2","销售额3","销售额4","销售额5","销售额6","销售额7","销售额8","销售额9","销售额10","广告1","广告2","广告3","广告4","广告5","广告6","广告7","广告8","广告9","广告10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]

#All_Product_Analyzefile_biaotou=All_Product_Analyzefile_Weeks_i.loc[:,["站点","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10"]]
All_Product_Analyzefile_biaotou.fillna(0)


All_Product_Analyzefile_biaotou=All_Product_Analyzefile_biaotou.groupby("站点").agg("sum")



print(All_Product_Analyzefile_biaotou)
AllCountry_Weeks= pd.concat([All_Product_Analyzefile_biaotou,plan])
AllCountry_Weeks=AllCountry_Weeks[["毛利润1","毛利润2","毛利润3","毛利润4","毛利润5","毛利润6","毛利润7","毛利润8","毛利润9","毛利润10","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","销售额1","销售额2","销售额3","销售额4","销售额5","销售额6","销售额7","销售额8","销售额9","销售额10","广告1","广告2","广告3","广告4","广告5","广告6","广告7","广告8","广告9","广告10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]
AllCountry_Weeks.to_excel(r'D:\\运营\\运行结果数据\\国家汇总.xlsx',sheet_name="sheet1",startrow=0,header=True,index=True)

