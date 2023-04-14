#######################以下生成Summary的程序##################################################################***************


from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil

#在这之前要生成汇总表，并且把每个国家的bulk表备份到D:\\运营\\bulkoperationfiles\\
newdate=input('输入最新日期y-m-d',) #输入最新一周的日期
maxtime=datetime.datetime.strptime(newdate,'%Y-%m-%d')
print(newdate)
print(maxtime)

Allbulkpath='D:\\运营\\2生成过程表\\'  
#Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')
#Allbulkold1=pd.read_excel(r'D:\\运营\\1数据源\\周Bulk广告数据汇总表历史\\'+"周bulk广告数据汇总表_2022-8-27_2022-9-24.xlsx")
#Allbulkold2=pd.read_excel(r'D:\\运营\\1数据源\\周Bulk广告数据汇总表历史\\'+"周bulk广告数据汇总表_2022-5-28_2022-8-20.xlsx")



#Allbulk=pd.concat([Allbulk,Allbulkold1,Allbulkold2])
#Allbulk["周数"]=(Allbulk["日期"]-maxtime).dt.days//7+1


print("以下生成Summary的程序")



#定义bulk数据汇总表所在路径
Allbulkpath='D:\\运营\\2生成过程表\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx',dtype = {'SKU':str})



AllbulkCampaign=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","Campaign Targeting Type"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')

AllbulkCampaign_List=AllbulkCampaign["Campaign"].drop_duplicates().to_list()
AllbulkCampaign_Country_List=AllbulkCampaign["Country"].drop_duplicates().to_list()

#AllbulkCampaign1week=Allbulk[(Allbulk['Record Type']=="Campaign")&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')


AllbulkCampaignWEEK=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","周数","Campaign Status"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUTotal=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","Campaign Status"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')
AllbulkCampaignSKUWEEK["Campaign Targeting Type"]=""
AllbulkCampaignSKUTotal["Campaign Targeting Type"]=""

 
for i in range(len(AllbulkCampaignSKUTotal))
            


        
    campaininTotal_oi=AllbulkCampaignSKUTotal.iloc[[boi],[2]]
    CountryTotal_oi=AllbulkCampaignSKUTotal.iloc[[boi],[1]]
    campaigntype99=AllbulkCampaign.loc[(AllbulkCampaign["Country"]==CountryTotal_oi)&(AllbulkCampaign["Campaign"]==campaininTotal_oi,"Campaign Targeting Type"].values[0][0] 
    AllbulkCampaignSKUTotal.loc[(AllbulkCampaignSKUTotal["Campaign"]==campaininTotal_oi) &(AllbulkCampaignSKUTotal["Country"]==CountryTotal_oi),"Campaign Targeting Type"]=campaigntype99 



############################################################################################################################


AllbulkCampaignSKUTotal["zhuanhualv"]=AllbulkCampaignSKUTotal["Orders"]/AllbulkCampaignSKUTotal["Clicks"]

AllbulkCampaignSKUTotal["zhuanhualv_rank1"]=AllbulkCampaignSKUTotal.groupby(["Country","Campaign","SKU","Campaign Status","Campaign Targeting Type"],as_index=False)[["zhuanhualv"]].rank(ascending=0,method='max')
AllbulkCampaignSKUTotal["zhuanhualv_rank2"]=AllbulkCampaignSKUTotal.groupby(["Country","Campaign","SKU","Campaign Status","Campaign Targeting Type"],as_index=False)[["zhuanhualv"]].rank(ascending=0,method='dense')                                                                                       

AllbulkCampaignSKUTotalzhuanhualvMax=AllbulkCampaignSKUTotal.groupby(["Country","Campaign","SKU","Campaign Status","Campaign Targeting Type"],as_index=False)[["zhuanhualv"]].agg('max')

AllbulkCampaignSKUWEEK["Campaign Targeting Type"]=""
 
for i in range(len(AllbulkCampaignSKUWEEK))
    campaininTotal_week_oi=AllbulkCampaignSKUWEEK.iloc[[i],[2]]
    CountryTotal_week_i=AllbulkCampaignSKUWEEK.iloc[[i],[1]]
    CountryTotal_week_week_i=AllbulkCampaignSKUWEEK.iloc[[i],[4]]                          
                                       
    campaigntype_week=AllbulkCampaign.loc[(AllbulkCampaign["Country"]==CountryTotal_oi)&(AllbulkCampaign["Campaign"]==campaininTotal_oi,"Campaign Targeting Type"].values[0][0] 
    AllbulkCampaignSKUTotal.loc[(AllbulkCampaignSKUTotal["Campaign"]==campaigntype99) &(AllbulkCampaignSKUTotal["Country"]==CountryTotal_oi),"Campaign Targeting Type"]=campaigntype99 
 
AllbulkCampaignSKUWEEK["zhuanhualv"]=AllbulkCampaignSKUWEEK["Orders"]/AllbulkCampaignSKUWEEK["Clicks"]

AllbulkSKUMax=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].max()


                                                                              
AllbulkCampaignKeywordWEEK=Allbulk.groupby(["Country","Campaign","Keyword or Product Targeting","Ad Group","周数"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum")




#以下为建立Camaign和SKU联系的程序
ALLbulkCampaignSKU=Allbulk[['Country','Campaign','SKU']]

ALLbulkCampaignSKU=ALLbulkCampaignSKU.drop_duplicates()
ALLbulkCampaignSKU=ALLbulkCampaignSKU.dropna(axis=0,how='any')
print(ALLbulkCampaignSKU)

CamaignSKUAgg=ALLbulkCampaignSKU.groupby(["Country","Campaign"],as_index=False).agg({'SKU':[",".join]})#追加的新的汇总comaignSKU
New_columns=['Country',"Campaign",'SKU']


AllbulkSKUMax_Campaign=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","SKU","周数","Campaign Status"],as_index=False)[['Spend']].agg('max')
 
 
AllbulkCampaignSKUWEEK["Spend_Order"]=AllbulkCampaignSKUWEEK.groupby(["Country","Campaign","SKU","周数","Campaign Status","Campaign Targeting Type"],as_index=False)[['Spend']].rank(ascending=0,method='max')
AllbulkCampaignSKUWEEK["Click_Order"]=AllbulkCampaignSKUWEEK.groupby(["Country","Campaign","SKU","周数","Campaign Status","Campaign Targeting Type"],as_index=False)[['Clicks']].rank(ascending=0,method='max')
#SKUCampaign_zhouzhuqanlv_Max

                                                                              
CamaignSKUAgg.columns=New_columns




writer=pd.ExcelWriter(Allbulkpath+'周bulk数据Summary.xlsx')
AllbulkCampaign.to_excel(writer,"Campaign汇总")
AllbulkCampaignWEEK.to_excel(writer,"CampaignWEEK汇总")
AllbulkSKUWEEK.to_excel(writer,"SKU-WEEK")
AllbulkCampaignSKUWEEK.to_excel(writer,"SKU-Campaign-WEEK",index=False)
AllbulkCampaignKeywordWEEK.to_excel(writer,"Keyword-Campaign-WEEK")
AllbulkSKUMax_Campaign.to_excel(writer,"SKUMax-Campaign")
AllbulkCampaignSKUTotal.to_excel(writer,"AllSKUCampaign")
CamaignSKUAgg.to_excel(writer,"CamaignSKUAgg")#追加的新的汇总comaignSKU
AllbulkCampaignSKUTotalzhuanhualvMax.to_excel(writer,"CampaignSKUTotalzhuanhualvMax")
writer.close()


######################################以下为做Biaotou周汇总###################################################

