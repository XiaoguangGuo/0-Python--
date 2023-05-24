
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

plan=pd.read_excel(r'D:\运营\1数据源\plan.xlsx',sheet_name="Sheet1")
plan["Country"].replace("CA","GV-CA",inplace=True)
plan["Country"].replace("US","GV-US",inplace=True)
plan["Country"].replace("MX","GV-MX",inplace=True)



SailingstarPlan=pd.read_excel(r'D:\运营\2生成过程表\All_Product_Analyzefile_Weeks排序.xlsx',sheet_name="sheet1")


SailingstarPlan.rename(columns = {'站点':'Country','MSKU':'SKU','销量1':'1','销量2':'2','销量3':'3','销量4':'4','销量5':'5','销量6':'6','销量7':'7','销量8':'8','销量9':'9','销量10':'10'},inplace=True)

SailingstarPlan.rename(columns = {'广告花费1':'广告1','广告花费2':'广告2','广告花费3':'广告3','广告花费4':'广告4','广告花费5':'广告5','广告花费6':'广告6','广告花费7':'广告7','广告花费8':'广告8','广告花费9':'广告9','广告花费10':'广告10'},inplace=True)
SailingstarPlan.rename(columns = {'FBA可售':'Fufillable'},inplace=True)

SailingstarPlan=SailingstarPlan.loc[~SailingstarPlan["Country"].isnull()]

for countryname99 in SailingstarPlan["Country"].drop_duplicates().to_list():
    SailingstarPlan.loc[SailingstarPlan["Country"]==countryname99,'Country']=CountrDic[countryname99]

SailingstarPlan=SailingstarPlan.loc[~SailingstarPlan["Country"].isnull()]

plan=pd.concat([plan,SailingstarPlan],ignore_index=True)
plan["广告浪费金额"]=0
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
    plan.loc[plan["Country"]==plan_country,"Exchangerate"]=exchangerate_20221217[plan_country]

print(plan["Exchangerate"])

plan["广告1in美元"]=plan["广告1"]/plan["Exchangerate"]
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


writer=pd.ExcelWriter(r'D:\\运营\\3数据分析结果\\'+ "国家汇总.xlsx")

plan.to_excel(writer,"ProductActions")
writer.save()

