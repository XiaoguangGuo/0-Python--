

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime

CampaignSKU_Summary=pd.read_excel(r'D:/运营/2生成过程表/周bulk数据Summary.xlsx',sheet_name="SKU-Campaign-WEEK")

CampaignSKU_SummarySum=CampaignSKU_Summary.groupby(["Country","SKU","Campaign"],as_index=False)[["Impressions","Clicks","Spend","Orders","Total Units","Sales"]].agg('sum')



CampaignSKU_Summary["皮质层标签"]=" "

CampaignSKU_Summary["Zhouzhuanlv"]=CampaignSKU_Summary["Orders"]/CampaignSKU_Summary["Clicks"]

CampaignSKU_SummarySum["Zhouzhuanlv"]=CampaignSKU_SummarySum["Orders"]/CampaignSKU_SummarySum["Clicks"]


CampaignSKU_Summary.loc[(CampaignSKU_Summary["Clicks"]>0) &(CampaignSKU_Summary["Zhouzhuanlv"]>0.15),"皮质层标签"] = CampaignSKU_Summary["皮质层标签"].astype(str)+"好广告"


CampaignSKU_Summary.loc[(CampaignSKU_Summary["Clicks"]>0) &(CampaignSKU_Summary["Zhouzhuanlv"]<0.05),"皮质层标签"] = CampaignSKU_Summary["皮质层标签"].astype(str)+"差广告"

#CampaignSKU_Summary10=CampaignSKU_Summary.loc[(CampaignSKU_Summary["周数"]<5)&(CampaignSKU_Summary["Country"]=="GV-US")]

CampaignSKU_Summary_biaotou=CampaignSKU_Summary[["Country","SKU","Campaign"]].drop_duplicates()
CampaignSKU_Summary_biaotou=pd.merge(CampaignSKU_Summary_biaotou,CampaignSKU_SummarySum,on=["Country","SKU","Campaign"] ,how="left")
print(CampaignSKU_Summary_biaotou)

for i in range(1,20):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
    CampaignSKU_Summary_i=CampaignSKU_Summary.loc[(CampaignSKU_Summary["周数"]==i)]
    
    #CampaignSKU_Summary_i=CampaignSKU_Summary_i["Country","SKU","Campaign","Clicks","Orders"]
    #更改列名

    CampaignSKU_Summary_i.rename(columns = {'Clicks':'Clicks'+str(i), 'Orders':'Orders'+str(i),'Spend':'Spend'+str(i),'Impressions':'Impressions'+str(i)}, inplace = True)

    CampaignSKU_Summary_biaotou=pd.merge(CampaignSKU_Summary_biaotou,CampaignSKU_Summary_i,on=["Country","SKU","Campaign"] ,how="left")
    




#CampaignSKU_Summary_pivot10=CampaignSKU_Summary10.pivot_table(values=["Clicks","Orders"], index=['Country','SKU','Campaign'],columns="周数", aggfunc = 'sum', fill_value=None, margins=False, dropna=False,margins_name='All').reset_index() # 是否启用总计行/列# 值

print(CampaignSKU_Summary_biaotou)


CampaignSKU_Summary_biaotou.to_excel(r'D:\\运营\\2生成过程表\\CampaignSKU_Summary_biaotou.xlsx',sheet_name="sheet1",startrow=0,header=True,index=True)




