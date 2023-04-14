
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil

#######################以下生成Summary的程序##################################################################***************



print("以下生成Summary的程序")



#定义bulk数据汇总表所在路径
Allbulkpath='D:\\运营\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx',dtype = {'SKU':str})



AllbulkCampaign=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
#AllbulkCampaign1week=Allbulk[(Allbulk['Record Type']=="Campaign")&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')


AllbulkCampaignWEEK=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignKeywordWEEK=Allbulk.groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")
AllbulkSKUMax=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('max')

#以下为建立Camaign和SKU联系的程序
ALLbulkCampaignSKU=Allbulk[['Country','Campaign','SKU']]

ALLbulkCampaignSKU=ALLbulkCampaignSKU.drop_duplicates()
ALLbulkCampaignSKU=ALLbulkCampaignSKU.dropna(axis=0,how='any')
print(ALLbulkCampaignSKU)

CamaignSKUAgg=ALLbulkCampaignSKU.groupby(["Country","Campaign"],as_index=False).agg({'SKU':[",".join]})#追加的新的汇总comaignSKU
New_columns=['Country',"Campaign",'SKU']


AllbulkSKUMax_Campaign=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","SKU"],as_index=False)[['Spend']].agg('max')
AllbulkCampaign_Status=Allbulk.loc[(Allbulk['Record Type']=="Campaign")&(Allbulk['周数']==1),["Country","Campaign","Campaign Status"]]
 
CamaignSKUAgg.columns=New_columns




writer=pd.ExcelWriter(Allbulkpath+'周bulk数据Summary.xlsx')
AllbulkCampaign.to_excel(writer,"Campaign汇总")
AllbulkCampaignWEEK.to_excel(writer,"CampaignWEEK汇总")
AllbulkSKUWEEK.to_excel(writer,"SKU-WEEK")
AllbulkCampaignSKUWEEK.to_excel(writer,"SKU-Campaign-WEEK")
AllbulkCampaignKeywordWEEK.to_excel(writer,"Keyword-Campaign-WEEK")
AllbulkSKUMax_Campaign.to_excel(writer,"SKUMax-Campaign")
 
CamaignSKUAgg.to_excel(writer,"CamaignSKUAgg")#追加的新的汇总comaignSKU

writer.save()

######################################以下为做Biaotou周汇总###################################################


# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime

CampaignSKU_Summary=pd.read_excel(r'D:/运营/周bulk数据Summary.xlsx',sheet_name="SKU-Campaign-WEEK")




CampaignSKU_Summary["皮质层标签"]=" "

CampaignSKU_Summary["Zhouzhuanlv"]=CampaignSKU_Summary["Orders"]/CampaignSKU_Summary["Clicks"]




CampaignSKU_Summary.loc[(CampaignSKU_Summary["Clicks"]>0) &(CampaignSKU_Summary["Zhouzhuanlv"]>0.1),"皮质层标签"] = CampaignSKU_Summary["皮质层标签"].astype(str)+"好广告"


CampaignSKU_Summary.loc[(CampaignSKU_Summary["Clicks"]>0) &(CampaignSKU_Summary["Zhouzhuanlv"]<0.03),"皮质层标签"] = CampaignSKU_Summary["皮质层标签"].astype(str)+"差广告"

#CampaignSKU_Summary10=CampaignSKU_Summary.loc[(CampaignSKU_Summary["周数"]<5)&(CampaignSKU_Summary["Country"]=="GV-US")]

CampaignSKU_Summary_biaotou=CampaignSKU_Summary[["Country","SKU","Campaign"]].drop_duplicates()
print(CampaignSKU_Summary_biaotou)

for i in range(1,5):
    #CampaignSKU_Summary_i=CampaignSKU_Summary["Clicks","Orders"].loc[(CampaignSKU_Summary["周数"]==i)]
    CampaignSKU_Summary_i=CampaignSKU_Summary.loc[(CampaignSKU_Summary["周数"]==i)]
    
    #CampaignSKU_Summary_i=CampaignSKU_Summary_i["Country","SKU","Campaign","Clicks","Orders"]
    #更改列名

    CampaignSKU_Summary_i.rename(columns = {'Clicks':'Clicks'+str(i), 'Orders':'Orders'+str(i),'Spend':'Spend'+str(i),'Impressions':'Impressions'+str(i)}, inplace = True)

    CampaignSKU_Summary_biaotou=pd.merge(CampaignSKU_Summary_biaotou,CampaignSKU_Summary_i,on=["Country","SKU","Campaign"] ,how="left")
    




#CampaignSKU_Summary_pivot10=CampaignSKU_Summary10.pivot_table(values=["Clicks","Orders"], index=['Country','SKU','Campaign'],columns="周数", aggfunc = 'sum', fill_value=None, margins=False, dropna=False,margins_name='All').reset_index() # 是否启用总计行/列# 值

print(CampaignSKU_Summary_biaotou)


CampaignSKU_Summary_biaotou.to_excel(r'D:\\运营\\运行结果数据\\CampaignSKU_Summary_biaotou.xlsx',sheet_name="sheet1",startrow=0,header=True,index=True)


#########################################################################################



