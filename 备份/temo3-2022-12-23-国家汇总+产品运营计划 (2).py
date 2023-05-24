####################################以下为汇总新站的###############################################
###################################以下为汇总新站的###############################################

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime 

newdate=input('输入最新日期y-m-d：新站汇总到周日',) #输入最新一周的日期
maxtime=datetime.datetime.strptime(newdate,'%Y-%m-%d')
print(maxtime)
maxtimeday=datetime.datetime.strptime(newdate,'%Y-%m-%d').date()
print(maxtimeday)

Product_Analyzepath = r'D:\\运营\\1数据源\\Product_Analyze产品分析\\'
print(Product_Analyzepath)
All_Product_Analyzefile=pd.read_excel(r'D:\\运营\2生成过程表\\All_Product_Analyzefile.xlsx',sheet_name=0)
for Product_Analyzefile in os.listdir(Product_Analyzepath):
    print(Product_Analyzefile)
#遍历数据文件
    datestr=os.path.basename(Product_Analyzefile).split("_")[2]
    Product_Analyzefile_DF=pd.read_excel(Product_Analyzepath +str(Product_Analyzefile)).assign(日期=datestr[0:10])
    
    All_Product_Analyzefile["周数"]=1
    All_Product_Analyzefile=All_Product_Analyzefile.append(Product_Analyzefile_DF,ignore_index=True)
         
    shutil.move(Product_Analyzepath + str(Product_Analyzefile), 'D:/运营/HistoricalData/Product_Analyzefile')


All_Product_Analyzefile_Weeks=All_Product_Analyzefile[["ASIN","店铺名","站点",'MSKU']].drop_duplicates()
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
        All_Product_Analyzefile_Weeks_i=All_Product_Analyzefile_Weeks_i[["ASIN","店铺名","站点",'MSKU',"销量",'广告点击量','广告花费','广告订单量','毛利润']]
 
    All_Product_Analyzefile_Weeks_i.rename(columns = {'销量':'销量'+str(i), '广告点击量':'广告点击量'+str(i),'广告花费':'广告花费'+str(i),'广告订单量':'广告订单'+str(i),'毛利润':'毛利润'+str(i)}, inplace = True)

    #合并

    All_Product_Analyzefile_Weeks=pd.merge(All_Product_Analyzefile_Weeks,All_Product_Analyzefile_Weeks_i,on=["ASIN","店铺名","站点",'MSKU'] ,how="left")
    
    All_Product_Analyzefile_Weeks=All_Product_Analyzefile_Weeks.drop_duplicates()
    All_Product_Analyzefile_Weeks=All_Product_Analyzefile_Weeks.dropna(axis=0,subset=['MSKU'])
#writer=pd.ExcelWriter(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile.xlsx')

#All_Product_Analyzefile.to_excel(writer,"All_Product_Analyzefile")
#All_Product_Analyzefile_Weeks.to_excel(writer,"All_Product_Analyzefile_Weeks")

                                 
All_Product_Analyzefile.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile_Weeks.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

All_Product_Analyzefile_Weeks=All_Product_Analyzefile_Weeks[["ASIN","店铺名","站点",'MSKU',"FBA可售","可售天数预估","标签","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","广告点击量1","广告点击量2","广告点击量3","广告点击量4","广告点击量5","广告点击量6","广告点击量7","广告点击量8","广告点击量9","广告点击量10","广告花费1", "广告花费2","广告花费3","广告花费4","广告花费5","广告花费6","广告花费7","广告花费8","广告花费9","广告花费10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]

All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile_Weeks排序.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)



#######################################################################################################################################################################################################################################################


# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime


exchangerate_20221217={"GV-US":1,"GV-CA":1.3701,"NEW-UK":0.8223,"NEW-JP":136.6790,"NEW-CA":1.3701,"NEW-IT":0.9457,"NEW-DE":0.9457,"NEW-ES":0.9457,"NEW-FR":0.9457,"NEW-US":1,"HM-US":1,"GV-MX":19.774,"NEW-MX":19.774}

plan1=pd.read_excel(r'D:\运营\1数据源\plan.xlsx',sheet_name=0)
All_Product_Analyzefile=pd.read_excel(r'D:\运营\2生成过程表\All_Product_Analyzefile.xlsx',sheet_name=0)
#####：以下为草稿
#################am_gv=plan["销售额"].sum()am_sailingstar=All_Product_Analyzefile[[""销售额"],"周数"=1].sum()
#AM=am_gv+am_sailingstar


#Plan_gv_group=plan.groupby(["Country"],"xiaoshou").agg.("sum")
#All_Product_Analyzefile_group=All_Product_Analyzefile.(["Country"],"xiaoshou").agg.("sum")



#####：草稿结束


#筛选出老站计划近10周的销售额，订单数和广告额，广告订单数
plan1=plan1.drop_duplicates()
plan1["销售额1"]=0.00001
plan1["广告订单1"]=0
plan1["毛利润"]=0

plan1.rename(columns = {'Country':'站点','1':'销量1','2':'销量2','3':'销量3','4':'销量4','4':'销量4','5':'销量5','6':'销量6','7':'销量7','8':'销量8','9':'销量9','10':'销量10'},inplace=True)


print(plan1)





plan1=plan1.groupby("站点").agg("sum")



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
All_Product_Analyzefile_biaotou=All_Product_Analyzefile_biaotou[["站点","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","销售额1","销售额2","销售额3","销售额4","销售额5","销售额6","销售额7","销售额8","销售额9","销售额10","广告1","广告2","广告3","广告4","广告5","广告6","广告7","广告8","广告9","广告10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]

#All_Product_Analyzefile_biaotou=All_Product_Analyzefile_Weeks_i.loc[:,["站点","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10"]]
All_Product_Analyzefile_biaotou.fillna(0)


All_Product_Analyzefile_biaotou=All_Product_Analyzefile_biaotou.groupby("站点").agg("sum")



print(All_Product_Analyzefile_biaotou)
AllCountry_Weeks= pd.concat([All_Product_Analyzefile_biaotou,plan1])
AllCountry_Weeks["销量1"].fillna(0,inplace=True)
AllCountry_Weeks["销量2"].fillna(0,inplace=True)
AllCountry_Weeks["销量3"].fillna(0,inplace=True)
AllCountry_Weeks["广告1"].fillna(0,inplace=True)
AllCountry_Weeks["广告2"].fillna(0,inplace=True)
AllCountry_Weeks["广告3"].fillna(0,inplace=True)
AllCountry_Weeks=AllCountry_Weeks[["销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","销售额1","销售额2","销售额3","销售额4","销售额5","销售额6","销售额7","销售额8","销售额9","销售额10","广告1","广告2","广告3","广告4","广告5","广告6","广告7","广告8","广告9","广告10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]
AllCountry_Weeks["本周销量增长"]=AllCountry_Weeks["销量1"]-AllCountry_Weeks["销量2"]
AllCountry_Weeks["本周广告增长"]=AllCountry_Weeks["广告1"]-AllCountry_Weeks["广告2"]


AllCountry_Weeks["本周广告销售比"]=AllCountry_Weeks["广告1"]/(AllCountry_Weeks.loc[AllCountry_Weeks["销量1"]>0,"销量1"])
AllCountry_Weeks["上周广告销售比"]=AllCountry_Weeks["广告2"]/(AllCountry_Weeks.loc[AllCountry_Weeks["销量2"]>0,"销量2"])


AllCountry_Weeks.loc[(AllCountry_Weeks["销量1"]==0)&(AllCountry_Weeks["广告1"]*1>0),"上周广告销售比"]=99999

AllCountry_Weeks.loc[(AllCountry_Weeks["销量2"]==0)&(AllCountry_Weeks["广告2"]*1>0),"上周广告销售比"]=99999

AllCountry_Weeks["本周广告销售比变化"]=AllCountry_Weeks["本周广告销售比"]-AllCountry_Weeks["上周广告销售比"]

AllCountry_Weeks.loc[(AllCountry_Weeks["广告1"]==0),"本周广告销售比"]=0
AllCountry_Weeks.loc[(AllCountry_Weeks["广告2"]==0),"上周广告销售比"]=0



AllCountry_Weeks=AllCountry_Weeks.reindex(columns=["本周销量增长","本周广告增长","本周广告销售比变化","本周广告销售比","上周广告销售比","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","销售额1","销售额2","销售额3","销售额4","销售额5","销售额6","销售额7","销售额8","销售额9","销售额10","广告1","广告2","广告3","广告4","广告5","广告6","广告7","广告8","广告9","广告10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"])

writer2=pd.ExcelWriter(r'D:\\运营\\3数据分析结果\\'+ "国家汇总.xlsx")


AllCountry_Weeks.to_excel(writer2,"CountriesSummary")

#两个报表合并
#Allcountry=conact(Plan_gv_group,)

#再groupby

#变成”国家“”第一周订单数“，“第一周金额””“第一周广告金额”“广告订单数”第n周订单数“，“第一周金额””“第一周广告金额”“广告订单数”

##########################################################################################################################################################################
######################################################################################################


# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime


CountrDic={"加拿大":"NEW-CA","美国":"NEW-US","英国":"NEW-UK","意大利":"NEW-IT","德国":"NEW-DE","法国":"NEW-FR","西班牙":"NEW-ES","日本":"NEW-JP","墨西哥":"NEW-MX"}

exchangerate_20221217={"GV-US":1,"GV-CA":1.3701,"NEW-UK":0.8223,"NEW-JP":136.6790,"NEW-CA":1.3701,"NEW-IT":0.9457,"NEW-DE":0.9457,"NEW-ES":0.9457,"NEW-FR":0.9457,"NEW-US":1,"HM-US":1,"GV-MX":19.774,"NEW-MX":19.774}

plan=pd.read_excel(r'D:\运营\1数据源\plan.xlsx',sheet_name=0)
plan["Country"].replace("CA","GV-CA",inplace=True)
plan["Country"].replace("US","GV-US",inplace=True)
plan["Country"].replace("MX","GV-MX",inplace=True)



SailingstarPlan=pd.read_excel(r'D:\运营\2生成过程表\All_Product_Analyzefile_Weeks排序.xlsx',sheet_name=0)


SailingstarPlan.rename(columns = {'站点':'Country','MSKU':'SKU','销量1':'1','销量2':'2','销量3':'3','销量4':'4','销量5':'5','销量6':'6','销量7':'7','销量8':'8','销量9':'9','销量10':'10'},inplace=True)

SailingstarPlan.rename(columns = {'广告花费1':'广告1','广告花费2':'广告2','广告花费3':'广告3','广告花费4':'广告4','广告花费5':'广告5','广告花费6':'广告6','广告花费7':'广告7','广告花费8':'广告8','广告花费9':'广告9','广告花费10':'广告10'},inplace=True)
SailingstarPlan.rename(columns = {'FBA可售':'Fufillable'},inplace=True)

SailingstarPlan=SailingstarPlan.loc[~SailingstarPlan["Country"].isnull()]

for countryname99 in SailingstarPlan["Country"].drop_duplicates().to_list():
    SailingstarPlan.loc[SailingstarPlan["Country"]==countryname99,'Country']=CountrDic[countryname99]

SailingstarPlan=SailingstarPlan.loc[~SailingstarPlan["Country"].isnull()]

plan=pd.concat([plan,SailingstarPlan],ignore_index=True)
plan["SKU"].astype(str)

CampaignWeek1=pd.read_excel(r'D:\运营\2生成过程表\周Bulk数据Summary.xlsx',sheet_name="SKU-Campaign-WEEK")
CampaignWeek1=CampaignWeek1[CampaignWeek1["周数"]==1]
print(CampaignWeek1.columns)
CampaignWeek1CampaignTotalCount=CampaignWeek1.groupby(["Country","SKU","Campaign Status"],as_index=False)["Campaign"].agg("count")#groupby(["Country","Campaign","周数"],as_index=False)

CampaignWeek1CampaignTotalCount.to_excel(r'D:\运营\2生成过程表\CampaignWeek1CampaignTotalCount.xlsx')

print(CampaignWeek1CampaignTotalCount.columns)


 
#CampaignWeek1["广告开启数量"]=CampaignWeek1[CampaignWeek1['Campaign Status']=="enabled"].groupby["Country","SKU"],as index=False)[["Campaign"]].count()

 
 

plan["广告开启数量"]=""
plan["广告浪费金额"]=""
plan["广告金额占比"]=""
plan["库存占比"]=""
plan["库存金额占比"]=""


plan["第1周入库"].fillna(0,inplace=True)
plan["Price"].fillna(0,inplace=True)

instock_List=["第1周入库","第2周入库","第3周入库","第4周入库","第5周入库","第6周入库","第7周入库","第8周入库","第9周入库","第10周入库","第11周入库","第12周入库","第13周入库","第14周入库"]

plan[instock_List]=plan[instock_List].fillna(0)
plan["Receiving"].fillna(0,inplace=True)
plan["Price"].fillna(0,inplace=True)
plan["TotalAmount"]=plan["Price"]*plan["STOCKALL"]
plan["广告1"].fillna(0,inplace=True)
plan["Exchangerate"]=""
plan_country_List=plan["Country"].drop_duplicates().to_list()
for plan_country in plan_country_List:
    print(plan_country)
    plan.loc[plan["Country"]==plan_country,"Exchangerate"]=exchangerate_20221217[plan_country]

    
    print(plan["Exchangerate"])

plan["广告1in美元"]=plan["广告1"]/plan["Exchangerate"]

for plan_country in plan_country_List:

    CampaignWeek1CampaignCountry_enabled=CampaignWeek1CampaignTotalCount.loc[(CampaignWeek1CampaignTotalCount["Country"]==plan_country)&(CampaignWeek1CampaignTotalCount["Campaign Status"]=="enabled")]

    CampaignWeek1CampaignCountry_enabled.to_excel(r'D:\运营\2生成过程表\CampaignWeek1CampaignTotalCount'+str(plan_country)+'.xlsx')
    
    #CampaignWeek1CampaignTotalCount.drop(labels=None,axis=0, index=None, columns=None, inplace=False) #在这里默认:axis=0,指删除index ...
   
    print(CampaignWeek1CampaignCountry_enabled)

    plan_country_SKU_list=plan.loc[plan["Country"]==plan_country,"SKU"].drop_duplicates().to_list()
    print(plan_country_SKU_list)
   
    CampaignWeek1WithEnabled_list=CampaignWeek1CampaignCountry_enabled.loc[(CampaignWeek1CampaignCountry_enabled["Country"]==plan_country)&(CampaignWeek1CampaignCountry_enabled["Campaign Status"]=="enabled"),"SKU"].drop_duplicates().to_list()

    #上面为找到开启状态的SKU的清单

    print("CampaignWeek1WithEnabled_list",CampaignWeek1WithEnabled_list)
    CampaignWeek1WithEnabled_listdf=pd.DataFrame(CampaignWeek1WithEnabled_list).drop_duplicates()
    CampaignWeek1WithEnabled_listdf.to_excel(r'D:\\运营\\2生成过程表\\'+str(plan_country)+"CampaignWeek1WithEnabled_list.xlsx")
    res=[]
    for i in CampaignWeek1WithEnabled_list:
        if i not in res:
            res.append(i)
   
    CampaignWeek1WithEnabled_list=res
    print("CampaignWeek1WithEnabled_list",CampaignWeek1WithEnabled_list)
   

    
    for plan_Country_SKU in plan_country_SKU_list:
 



        SKUchaifenAlllist=[]
        print(plan_Country_SKU)
               
        if ',' in  str(plan_Country_SKU):#是否包含逗号
            print("包含,",plan_Country_SKU)
            SKUchaifen99list=plan_Country_SKU.split(",")
            print(SKUchaifen99list)
            SKUchaifenAlllist+=SKUchaifen99list
            print(SKUchaifenAlllist)
            plan_enabled_sum=0
            for SKUchaifenAll_oi in SKUchaifenAlllist: 
                if SKUchaifenAll_oi in CampaignWeek1WithEnabled_list:
                    print("在清单中",SKUchaifenAll_oi)
                    
                    plan_enabled_oi=CampaignWeek1CampaignCountry_enabled.loc[(CampaignWeek1CampaignCountry_enabled["Country"]==plan_country)&(CampaignWeek1CampaignCountry_enabled["SKU"]==SKUchaifenAll_oi)&(CampaignWeek1CampaignCountry_enabled["Campaign Status"]=="enabled"),"Campaign"].values[0]
                    plan_enabled_sum+=plan_enabled_oi
                    print(plan_enabled_sum)
                    
                    plan.loc[(plan["Country"]==plan_Country_SKU)&(plan["SKU"]==plan_Country_SKU),"广告开启数量"]=plan_enabled_sum

        else:       
            print(plan_Country_SKU)
            if plan_Country_SKU in CampaignWeek1WithEnabled_list:
                print("不含逗号在清单",plan_Country_SKU)
                print("不含逗号在清单",CampaignWeek1WithEnabled_list)
                print(plan_country)
                
                xxenabled=CampaignWeek1CampaignCountry_enabled.loc[(CampaignWeek1CampaignCountry_enabled["Country"]==plan_country)&(CampaignWeek1CampaignCountry_enabled["SKU"]==plan_Country_SKU)&(CampaignWeek1CampaignCountry_enabled["Campaign Status"]=="enabled"),"Campaign"].values[0]

                plan.loc[(plan["Country"]==plan_country)&(plan["SKU"]==plan_Country_SKU),"广告开启数量"]=xxenabled
             

#################################################以下有问题#################################################################
  #################################################以上有问题#################################################################       
                                                                  
print(plan["广告1in美元"])
plan["皮质层标签"]=""
plan["行动方案"]=""
#平均每天销售一个的产品为重点产品
plan.loc[(plan["SELLING10"]>10),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 重点产品"
plan.loc[(plan["Adjusted-Week2"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要快递发货"
plan.loc[(plan["Adjusted-Week4"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要空运发货"
plan.loc[(plan["Adjusted-Week8"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要快海运发货"
plan.loc[(plan["Adjusted-Week15"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要慢海运发货"
plan.loc[(plan["GGZZ1"]>0.5),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告增长"
plan.loc[(plan["GGZZ1"]<-0.5),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告降低"
plan.loc[((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0)&((plan["STOCKALL"]/(plan["1"]+plan["2"]+plan["3"]+plan["4"])*4)<4),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 4周后缺货-提价"

plan.loc[(plan["广告1"]<0.5)| (plan["广告1"].isna())|(plan["广告1"]==""),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周无广告"

#每个订单花费广告费超过2美元就是效果差
plan.loc[((plan["Country"]=="US")|(plan["Country"]=="CA"))&( plan["BILI1"]>2 ),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告效果差"
plan.loc[((plan["Country"]=="US")|(plan["Country"]=="CA"))&( plan["BILI1"]<0.3 ),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 广告花费占比小"
#4周广告费比上四周订单超过2美元就是长期效果差

plan.loc[((plan["Country"]=="US")|(plan["Country"]=="CA"))&(((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0) & ((plan["广告1"]+plan["广告2"]+plan["广告3"]+plan["广告4"])/(plan["1"]+plan["2"]+plan["3"]+plan["4"])>2)),"皮质层标签" ] = plan["皮质层标签"].astype(str)+" 长期广告效果差"

#墨西哥广告超过40比索是效果差

plan.loc[(( plan["广告1"]/plan["Exchangerate"])-plan["1"])>21,"皮质层标签"] = plan["皮质层标签"].astype(str)+" 广告浪费大于21美金"

plan["广告浪费金额"]=(plan["广告1"]/plan["Exchangerate"])-plan["1"]

#plan.loc[ ((plan["广告1in美元"]-plan["1"])/(plan["广告1in美元"])>0.7),"皮质层标签"] = plan["皮质层标签"].where (plan["广告1in美元"]>0).astype(str)+"广告浪费比例大于70%"
#plan.loc[plan["广告1in美元"]>0,"广告1浪费比例"]=(plan["广告1in美元"]-plan["1"])/(plan["广告1in美元"])

plan.loc[((plan["Country"]=="MX")&( plan["BILI1"]>40 )),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告效果差"

plan.loc[(plan["Country"]=="MX")&(((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0) & ((plan["广告1"]+plan["广告2"]+plan["广告3"]+plan["广告4"])/(plan["1"]+plan["2"]+plan["3"]+plan["4"])>40)),"皮质层标签" ] = plan["皮质层标签"].astype(str)+" 长期广告效果差"

plan.loc[(plan["广告1"]>0) &(plan["1"]==0),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告效果差"

#All_Product_Analyzefile_LABEL=All_Product_Analyzefile[(All_Product_Analyzefile["销量"]<2)&(All_Product_Analyzefile["FBA可售"]>=50)&(All_Product_Analyzefile["周数"]==1)]
plan.loc[(plan["1"]<1),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周无销量"
plan.loc[(plan["2"]>0)& (plan["1"]/plan["2"]<0.5),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周销量暴跌"
############################
plan.loc[(plan["Fufillable"]<1),"皮质层标签"] = plan["皮质层标签"].astype(str)+"无库存"
plan.loc[(plan["Fufillable"]<5) &(plan["Fufillable"]>0),"皮质层标签"] = plan["皮质层标签"].astype(str)+"低库存"

plan.loc[(plan["Fufillable"]>10),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 规模以上库存"


plan.loc[plan["SELLING10"].isnull(),"SELLING10"]=(plan["1"]+plan["2"]+plan["3"]+plan["4"]+plan["5"]+plan["6"]+plan["7"]+plan["8"]+plan["9"]+plan["10"])
plan.loc[plan["STOCKALL"].isnull(),"STOCKALL"]=(plan["Fufillable"]+plan["第1周入库"]*1+plan["第2周入库"]*1+plan["第3周入库"]*1+plan["第4周入库"]*1+plan["第5周入库"]*1+plan["第6周入库"]*1+plan["第7周入库"]*1+plan["第8周入库"]*1+plan["第9周入库"]*1+plan["第10周入库"*1]+plan["第11周入库"]*1+plan["第12周入库"]*1+plan["第13周入库"]*1+plan["第14周入库"]*1)

plan.loc[(plan["Zhouzhuan10"].isnull())&(plan["SELLING10"]>0),"Zhouzhuan10"]=plan["STOCKALL"]*10/plan["SELLING10"]

plan.loc[(plan["Zhouzhuan10"]>50)| (plan["SELLING10"]==0)&(plan["STOCKALL"]>0),"皮质层标签"] = plan["皮质层标签"].astype(str)+"长期积压"#对新站老站都有效




plan.loc[(plan["Fufillable"]+plan["Receiving"]<10) &(plan["第1周入库"]+plan["第2周入库"]+plan["第3周入库"]==0),"皮质层标签"] = plan["皮质层标签"].astype(str)+"库存小于10且三周内无入库"

############################################################

plan.loc[plan["皮质层标签"].str.contains("提价")&plan["皮质层标签"].str.contains("本周广告增长"),"行动方案"]=plan["行动方案"].astype(str)+"减少广告"

plan.loc[plan["皮质层标签"].str.contains("无广告")&plan["皮质层标签"].str.contains("规模以上库存")&plan["皮质层标签"].str.contains("长期积压"),"行动方案"]=plan["行动方案"].astype(str)+ "2-1确保广告是开的"

plan.loc[(((plan["1"]+plan["2"]+plan["3"]+plan["4"])==0)&(plan["Fufillable"]>30))|((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0)&(((plan["Fufillable"]/(plan["1"]+plan["2"]+plan["3"]+plan["4"])*4)>4))&(plan["Fufillable"]>10),"行动方案"] = plan["行动方案"].astype(str)+"2-2确保广告是开的"



plan.loc[plan["皮质层标签"].str.contains("库存小于10且三周内无入库"),"行动方案"]=plan["行动方案"].astype(str)+"关闭广告" 




############################占比##########################

Country_Ad_Selling_Sum_All=plan.groupby("Country",as_index=False)[["1","2","3","4","5","6","7","8","9","10","广告1","广告2","广告3","广告4","广告5","广告6","广告7","广告8","广告9","广告10"]].agg("sum")
print(Country_Ad_Selling_Sum_All)
Country_Stockall_sum=plan.groupby("Country",as_index=False)[["STOCKALL","Fufillable","TotalAmount"]].agg("sum")
print(Country_Stockall_sum)

for country1 in plan["Country"].drop_duplicates().to_list():
    
    
 
    Country_Ad_Selling_Sum_sku_country_sum=Country_Ad_Selling_Sum_All.loc[(Country_Ad_Selling_Sum_All["Country"]==country1),"广告1"].values[0]
    print(Country_Ad_Selling_Sum_sku_country_sum)
    Country_Stockall_sum_STOCKALL=Country_Stockall_sum.loc[(Country_Stockall_sum["Country"]==country1),"STOCKALL"].values[0]
    Country_Stockall_sum_TotalAmount=Country_Stockall_sum.loc[(Country_Stockall_sum["Country"]==country1),"TotalAmount"].values[0]
    

    plan.loc[plan["Country"]==country1,"广告金额占比"]=plan["广告1"]/Country_Ad_Selling_Sum_sku_country_sum
    plan.loc[plan["Country"]==country1,"库存占比"]=plan["STOCKALL"]/Country_Stockall_sum_STOCKALL
    plan.loc[plan["Country"]==country1,"库存金额占比"]=plan["STOCKALL"]/Country_Stockall_sum_TotalAmount

#SailingstarPlan=




plan.to_excel(writer2,"ProductActions")
writer2.close()
 

