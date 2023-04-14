
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np


input("检查zongbiao后回车")

zhuanhualv_bad={"GV-US":0.05,"GV-CA":0.05,"NEW-UK":0.05,"NEW-JP":0.03,"NEW-CA":0.05,"NEW-IT":0.05,"NEW-DE":0.05,"NEW-ES":0.05,"NEW-FR":0.05,"NEW-US":0.05,"HM-US":0.05,"GV-MX":0.03,"NEW-MX":0.03,"HM-US":0.05}
bulkdatafilepath = 'D:\\运营\\1数据源\周bulk广告数据\\'

#底线策略
SearchTermAll=pd.read_excel(r'D:\\运营\\2生成过程表\\Sponsored Products Search term report.xlsx')

SearchTermAll["Clicks"].fillna(0,inplace=True)


SearchTermAll["Customer Search Term"].astype(str)

SearchTermAll_Sum=SearchTermAll.groupby(["COUNTRY","Campaign Name", "Ad Group Name","Customer Search Term"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")
SearchTermAllquan=SearchTermAll.groupby(["COUNTRY","Campaign Name", "Ad Group Name","Customer Search Term","Targeting","Match Type"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")

SearchTermAll_Sum.loc[SearchTermAll_Sum['Clicks']>0,"转化率"]=SearchTermAll_Sum["7 Day Total Orders (#)"]/SearchTermAll_Sum['Clicks']


SearchTermAll_Good=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]>=0.25)&(SearchTermAll_Sum["Clicks"]>=2)] #条件改为


SearchTermAll_Bad=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]<0.05)&(SearchTermAll_Sum["Clicks"]>=20)]


ProductActions=pd.read_excel(r'D:\\运营\\3数据分析结果\\国家汇总.xlsx',sheet_name="ProductActions")



#打开Bulkfile的第一种方法

Allbulkpath='D:\\运营\\2生成过程表\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')



#遍历SeachTermGood


SearchTermAll_Good_Country_list=SearchTermAll_Good["COUNTRY"].drop_duplicates().to_list()

Allbulk_Campaign_SKUSum=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('sum')#全部历史bulk的sku spend相加：找自动广告的sku

################################################################GOOD#########################################################################################################################

for countryname in SearchTermAll_Good_Country_list:   #遍历searchTemgood里的国家
    Allbulk_Campaign_SKUSum_Country=Allbulk_Campaign_SKUSum[Allbulk_Campaign_SKUSum["Country"]==countryname].reindex()

    CountrySKU_Close_List=ProductActions.loc[(ProductActions["COUNTRY"]==countryname)&(ProductActions["行动方案"].str.contains("关闭广告")),"SKU"].drop_duplicates().to_list()#这个国家要关闭的SKU的List

    AllCountryActions_Country_SKU_Close_List_nocomma_list=[] 
    for AllCountryActions_Country_SKU_Close_List_nocomma in CountrySKU_Close_List:
         
        print(AllCountryActions_Country_SKU_Close_List_nocomma)
        AllCountryActions_Country_SKU_Close_List_nocomma=AllCountryActions_Country_SKU_Close_List_nocomma
        if ',' in  str(AllCountryActions_Country_SKU_Close_List_nocomma):
            print("包含,",AllCountryActions_Country_SKU_Close_List_nocomma)
            chaifenlist=AllCountryActions_Country_SKU_Close_List_nocomma.split(",")
            print(chaifenlist)
            AllCountryActions_Country_SKU_Close_List_nocomma_list+=chaifenlist
            print(AllCountryActions_Country_SKU_Close_List_nocomma_list)
        else:
            print(AllCountryActions_Country_SKU_Close_List_nocomma)
            chaifen=AllCountryActions_Country_SKU_Close_List_nocomma
            AllCountryActions_Country_SKU_Close_List_nocomma_list+=[chaifen]
                        
                    
            print(chaifen)
                    
                    
    CountrySKU_Close_List=AllCountryActions_Country_SKU_Close_List_nocomma_list  # 确定了这个国家最终要关闭的清单，去掉逗号因素
    print(CountrySKU_Close_List)
                
                                                            
    
     
    
    n=0                                             
    for bulkdatafile in os.listdir(bulkdatafilepath): #找bulkfile对应的国家文件 

        Bulkfile_Country=bulkdatafile.split('_')[0]

        if Bulkfile_Country==countryname:
            SearchTermAll_Good_Country=SearchTermAll_Good[SearchTermAll_Good["COUNTRY"]==str(countryname)]
         
            
            print("SearchTermAll_Good_Country")
            bulkfile=pd.read_excel(bulkdatafilepath+str(bulkdatafile),engine="openpyxl",sheet_name=1)  #找到对应国家的bulkfile
            bulkfile["更改记录"]=""
            n+=1
            
            break
        
    if n==1:
         
        print("找到",bulkdatafile)
        Bulkfile_SearchTerm_Add=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_auto=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])

        Bulkfile_SearchTerm_Add_auto_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_manual=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_draft99=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_manual_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])

   

        for Stoi in range(0,len(SearchTermAll_Good_Country)):#遍历好词 的行#for A0
 
            Stoi_keyword=SearchTermAll_Good_Country.iloc[[Stoi],[3]].values[0][0]
            Stoi_keyword=str(Stoi_keyword)
            print("Stoi_keyword=",Stoi_keyword)
            Stoi_Campaign=SearchTermAll_Good_Country.iloc[[Stoi],[1]].values[0][0]#获取第n行的关键词字段的值，遍历
            print(Stoi_Campaign)
            Stoi_AdGroup=SearchTermAll_Good_Country.iloc[[Stoi],[2]].values[0][0]
        

                                                  
            averageprice=SearchTermAll_Good_Country.iloc[[Stoi],[6]].values[0][0]/SearchTermAll_Good_Country.iloc[[Stoi],[5]].values[0][0]
            print("averageprice",averageprice)
            Allbulk_Campaign_SKUSum_Country_Campaign=Allbulk_Campaign_SKUSum_Country[Allbulk_Campaign_SKUSum_Country["Campaign"]==Stoi_Campaign].reset_index(drop=True)
            print(Allbulk_Campaign_SKUSum_Country_Campaign)
            
    


            type1=bulkfile.loc[(bulkfile['Record Type']=="Campaign")&(bulkfile['Campaign']==Stoi_Campaign),"Campaign Targeting Type"]
            length_type=len(bulkfile.loc[(bulkfile['Record Type']=="Campaign")&(bulkfile['Campaign']==Stoi_Campaign),"Campaign Targeting Type"]) #好词对应Campaignd对应的类型是否存在？
            print(type1,length_type)
            if Stoi_keyword.isdigit(): #如果keyword是数据就不处理
                print("数字")  
                continue   
            if Stoi_keyword.startswith('b0'):
                print("暂不处理Asin")
                continue
            if length_type==0:
               print("这个好词对应的Camaign在bulkfile没有记录，生成新的Campaign") # F01之前有campaign现在没了?那应该找在ALL里面这个Camaign对应的SKU
               #生成一个新的campaign
               

               if len(Allbulk_Campaign_SKUSum_Country_Campaign)==0:
                   print("实在找不到Sku，循环下一个词")
                   continue
               
               a89=Allbulk_Campaign_SKUSum_Country_Campaign["Spend"].idxmax()
               bulkfile_Campaign_SKUMax66=Allbulk_Campaign_SKUSum_Country_Campaign.iloc[[a89],[3]].values[0][0] #找到历史上花费最大的SKU
               
               #bulkfile_Campaign_SKUMax66=Allbulk_Campaign_SKUMax.groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('max')#找到某个SKU花费最多的
               #bulkfile_Campaign_SKU_list66=bulkfile_Campaign_SKUMax66.loc[(bulkfile_Campaign_SKUMax66["Country"]==countryname)&(bulkfile_Campaign_SKUMax66["Campaign"]==Stoi_Campaign)&(bulkfile_Campaign_SKUMax66['Ad Group']==Stoi_AdGroup),"SKU"].to_list()
               #if len(bulkfile_Campaign_SKU_list66)==0:

               #bulkfile_Campaign_SKUMax66=bulkfile_Campaign_SKU_list66[0]
               #bulkfile_Campaign_SKUMax66=Allbulk_Campaign_SKUMax[(Allbulk_Campaign_SKUMax["Campaign"]==Stoi_Campaign) &(Allbulk_Campaign_SKUMax['Ad Group']==Stoi_AdGroup)&(Allbulk_Campaign_SKUMax['Spend'].max()),"SKU"].values[0][0])  #只要落在自动的SKU上
               print("找好词对应的SKU：bulkfile_Campaign_SKUMax66",bulkfile_Campaign_SKUMax66)

               pr66_sku="Pr4a"+"-"+str(bulkfile_Campaign_SKUMax66)
               if bulkfile_Campaign_SKUMax66 not in CountrySKU_Close_List:
                   
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Campaign","Campaign":pr66_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","更改记录":"新增广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad Group","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled","更改记录":"新增广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad ","Campaign":pr66_sku,"Ad Group":pr66_sku,"SKU":bulkfile_Campaign_SKUMax66,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","更改记录":"新增广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Keyword","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"新增广告"},ignore_index = True)

               if bulkfile_Campaign_SKUMax66  in CountrySKU_Close_List:    
                   
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Campaign","Campaign":pr66_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","更改记录":"关闭广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad Group","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled","更改记录":"关闭广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad ","Campaign":pr66_sku,"Ad Group":pr66_sku,"SKU":bulkfile_Campaign_SKUMax66,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","更改记录":"关闭广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Keyword","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"关闭广告"},ignore_index = True)
             
                  


               
            else:  #F01 有对应的Campaign

               
                Campaign_type_list=bulkfile.loc[(bulkfile['Record Type']=="Campaign")&(bulkfile['Campaign']==Stoi_Campaign),"Campaign Targeting Type"].to_list() #
                print(Campaign_type_list)
       
#####################################如果好词从Auto广告产生########################Auto#################################################################################################                
                if Campaign_type_list[0]=="Auto" :#Camaign是否为自动 #if:AA1
                    
                    print("这个好词来自自动广告",Stoi_Campaign)
                                                                                              
                
                #找在这个Campaign到历史花费最多的SKU
                    print("好词对应的广告从自动广告产生")

                    a90=Allbulk_Campaign_SKUSum_Country_Campaign["Spend"].idxmax()
                    bulkfile_Campaign_SKUMax2=Allbulk_Campaign_SKUSum_Country_Campaign.iloc[[a90],[3]].values[0][0]
               
                    
                    
                    print(bulkfile_Campaign_SKUMax2)                                                                     
            
                    pr4_sku="Pr4"+"-"+str(bulkfile_Campaign_SKUMax2)
                ##上面为找出这个自动广告中花费最多的SKU，后面新开的广告就用到这个SKU：bulkfile_Campaign_SKUMax2

                    
                    bulkfileSKU_Campaign_List88=bulkfile.loc[(bulkfile['Record Type']=="Ad")&(bulkfile['SKU']==bulkfile_Campaign_SKUMax2),"Campaign"].drop_duplicates().to_list()#找到这个SKU对应的bulkfile所有campaign #上面为这个花费最多的SKU对应的Bulkfile中的所有Campaign的list
                    print("bulkfile全部campaign list",bulkfileSKU_Campaign_List88)

                        
                    bulkfileSKU_Manual_Campaigndf=bulkfile.loc[(bulkfile["Campaign"].isin(bulkfileSKU_Campaign_List88)==True)&(bulkfile['Record Type']=="Campaign")&(bulkfile["Campaign Targeting Type"]=="Manual"),"Campaign"]
                #找到所有手动的 
                    bulkfileSKU_Manual_Campaign_List=bulkfileSKU_Manual_Campaigndf.to_list() #好词对应的手动广告的list
                    bulkfileSKU_Manual_Campaign_Length=len(bulkfileSKU_Manual_Campaign_List)
                    print("bulkfileSKU_Manual_Campaign_List",bulkfileSKU_Manual_Campaign_List)


                    
                    if bulkfileSKU_Manual_Campaign_Length==0: #if-A1
                        print("bulkfileSKU_Manual_Campaign_List长度为0,没有好词对应的手动广告 ")
                        Maxbid_Adgroup=0.99


                        if bulkfile_Campaign_SKUMax2 not in CountrySKU_Close_List:
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Campaign","Campaign":pr4_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","更改记录":"生成新广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled","更改记录":"生成新广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","更改记录":"生成新广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"生成新广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase","更改记录":"生成新广告"},ignore_index = True)        


                        elif bulkfile_Campaign_SKUMax2 in CountrySKU_Close_List:
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Campaign","Campaign":pr4_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","更改记录":"关闭广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled","更改记录":"关闭广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","更改记录":"关闭广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"关闭广告"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase","更改记录":"关闭广告"},ignore_index = True)        



                    else:#if-A1如果有好词对应的手动广告
                     
                        pr4number=0
                        for bulkfileSKU_Manual_Campaign in bulkfileSKU_Manual_Campaign_List:#遍历为这个花费最多的SKU对应的Bulkfile中SKU对应的所有手动Campaign的list# for A1
                            print("遍历手动Campaign")
                            
                        
#############################################################11                                                          
                            if bulkfileSKU_Manual_Campaign.startswith('Pr4'):#如果有专门为关键词累计广告开的Pr4开头的广告（也可以找手动广告，但为了清晰，不计划加到之前的人工手动广告）#ifB1
                                print("找到了Pr4开头的手动广告")
                        #要找到这个产品所在的Ad Group，指定第一个


                                bulkfileSKUManualCampaignAdGroup_List=bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile['Record Type']=="Ad Group"),"Ad Group"].drop_duplicates().to_list()
                                bulkfileSKUManualCampaignAdGroup_Assign=bulkfileSKUManualCampaignAdGroup_List[0]
                            
##################                    
                                if Stoi_keyword in (bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Keyword or Product Targeting"].drop_duplicates().to_list()):
                                #if C1
                    ###上：如果有这个词的exact广告
                                    print("找到了exact广告")

                                    if bulkfile_Campaign_SKUMax2 not in CountrySKU_Close_List:
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"更改记录"]="打开Campaign"
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")&bulkfile["SKU"]==bulkfile_Campaign_SKUMax2,"更改记录"]="打开Ad"
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")&bulkfile["SKU"]==bulkfile_Campaign_SKUMax2,"Status"]="enabled" #打开SKU的状态
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"更改记录"]="打开Keyword"
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Status"]="enabled"#打开关键词的状态

                        
                                        Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Campaign")]) #把这个Pr4广告的Campaign行加到变更文件Bulkfile_SearchTerm_Add_auto里              
                                        Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad Group")]) #把这个Pr4广告的Ad Group行加到变更文件Bulkfile_SearchTerm_Add_auto里
                                        Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")])#把这个Pr4广告的Ad行加到变更文件Bulkfile_SearchTerm_Add_auto里
                                        print(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])

                            
                                        Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])#把这个Pr4广告的keyword行加到变更文件Bulkfile_SearchTerm_Add_auto里
                            
##################                  

                                else:#没有这个词对应SKU  的exact #if C1
                                    print("没有找到exact手动广告")
#生成一个pr4放到 Bulkfile_SearchTerm_Add_auto_draft

                                    if bulkfile_Campaign_SKUMax2 not in CountrySKU_Close_List:

                            
                                 #Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Campaign","Campaign":bulkfileSKU_Manual_Campaign,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
                                # Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                                #Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                                        Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":bulkfileSKU_Manual_Campaign,"Ad Group":bulkfileSKUManualCampaignAdGroup_Assign,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"增加exact词"},ignore_index = True)
                        
                                    elif bulkfile_Campaign_SKUMax2  in CountrySKU_Close_List:
                                       
                                        Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":bulkfileSKU_Manual_Campaign,"Ad Group":bulkfileSKUManualCampaignAdGroup_Assign,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"paused","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"增加exact词关闭广告可上传"},ignore_index = True)




                
                
 



                    
                   
                                if Stoi_keyword in (bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Keyword or Product Targeting"].drop_duplicates().to_list()):
###########如果存在这个SKU的Pr4广告有这个词的phrase
                                    if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"更改记录"]="打开词组" 
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Status"]="enabled"
                        #把存在的phrase广告状态变成enabled
                        
                        #Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Campaign")])               
                        #Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad Group")])
                        #Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")])
                                        Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase")])
                        #把这个广告加到Bulkfile_SearchTerm_Add_auto
                        

###################################################################1
          
                                else:       #如果没有phrase
                                   
                                      print("没有phrase的手动广告")
                                      if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                                          Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":bulkfileSKU_Manual_Campaign,"Ad Group":bulkfileSKUManualCampaignAdGroup_Assign,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase","更改记录":"增加词组"},ignore_index = True)
                      
                                break #如果找到一个pr4就别循环找第二个了

                        

#############################################################################################################################


                            else:#如果没有Pr4广告 #if B1
                                pr4number+=1
                                print("本次List循环运行没有pr4广告,继续循环")

#######################如果循环完了，没有pr4:就生成一个##############################

                                print("pr4number",pr4number)
                                if pr4number==bulkfileSKU_Manual_Campaign_Length:  #A1全部循环完时
                                    print("没有pr4，生成一个")
                                    Aleadyadd=Bulkfile_SearchTerm_Add_auto_draft.loc[(Bulkfile_SearchTerm_Add_auto_draft["Campaign"]==pr4_sku)&(Bulkfile_SearchTerm_Add_auto_draft["Ad Group"]==pr4_sku),"Ad Group"].to_list()
                                    if len(Aleadyadd)>0:
                                        if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"增加exact"},ignore_index = True)
                                    else:                                    
                                        if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                                    
                                            Maxbid_Adgroup=0.99

                                    
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Campaign","Campaign":pr4_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","更改记录":"增加Campaign"},ignore_index = True)
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled","更改记录":"增加Ad Group"},ignore_index = True)
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","更改记录":"增加Ad"},ignore_index = True)
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"增加关键词"},ignore_index = True)
                                    
 
#####################################Auto###################################1end


########################################Manual####################如果来自手动广告:则判断在这个Stoi_Campaign中  ############################################################################################################################################################################2
                                                                                              
           
          
                else: #如果来自手动:则就追加到在这个Stoi_Campaign中 #if:AA1

                    a91=Allbulk_Campaign_SKUSum_Country_Campaign["Spend"].idxmax()

                    print(a91)
                    
                    bulkfile_Campaign_SKUMax33=Allbulk_Campaign_SKUSum_Country_Campaign.iloc[[a91],[3]].values[0][0]
                    print(bulkfile_Campaign_SKUMax33)
                     
                




                    
                    print("好词产生的广告来自手动广告")
#下面判断bulkfile中 含有Stoi_keyword的exacT
#################################################EXACT               
                #print("断点",bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"Keyword or Product Targeting"])
                #print("断点",bulkfile[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])
                #print("断点",bulkfile[(bulkfile["Campaign"]==Stoi_Campaign)])
                    if Stoi_keyword in (bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Ad Group"]==Stoi_AdGroup)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Keyword or Product Targeting"].drop_duplicates().to_list()):                                                                         
                    #if#MN-A
                 #如果已存在exact广告
                        print("这个手动广告包含有这个好词的exact")
                        if bulkfile_Campaign_SKUMax33  not in CountrySKU_Close_List:
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"更改记录"]="打开手动广告exact词"
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Campaign Status"]="enabled" #打开这个词，防止关闭
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Ad Group Status"]="enabled"#打开这个词，防止关闭
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Status"]="enabled"#打开这个词，防止关闭


                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"更改记录"]="打开Campaign"  
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。是否打开关闭广告由产品计划决定！！
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                        #bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad"),"Status"]="enabled" #打开所有SKU的状态
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Campaign")]) #追加Campaign              
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad Group")])#追加Ad Group
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad")])#追加Ad
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),{"更改记录":"添加好词"}])
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])#追加Keyword

                
                    else:#if#MN-A
 #如果没有exac####生成在产生这个词的Campaign中。                                               
                 #生成一个
                        print("这个手动广告不包含这个好词的exact")
                        if bulkfile_Campaign_SKUMax33  not in CountrySKU_Close_List:
                            Bulkfile_SearchTerm_Add_manual_draft=Bulkfile_SearchTerm_Add_manual_draft.append({"Record Type":"Keyword","Campaign":Stoi_Campaign,"Ad Group":Stoi_AdGroup,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"新增exact词"},ignore_index = True)



                    #Maxbid_Adgroup=0.99
                    #Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Campaign","Campaign":Stoi_Campaign,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
                 #生成Campaign 并追加
                    #Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Ad Group","Campaign":Stoi_Campaign,"Ad Group":Stoi_AdGroup, "Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                 #生成Ad Group 并追加
                    #Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Ad ","Campaign":Stoi_Campaign,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                 #生成Ad SKU 并追加
                    #Bulkfile_SearchTerm_Add_manual_draft=Bulkfile_SearchTerm_Add_manual_draft.append({"Record Type":"Keyword","Campaign":Stoi_Campaign,"Ad Group":Stoi_AdGroup,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)
                 #生成keyword 并追加

#################################################phrase               
                    if Stoi_keyword in (bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Ad Group"]==Stoi_AdGroup)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Keyword or Product Targeting"].drop_duplicates().to_list()):
                    #if#MN-B
         #如果存在一个phrase的广告
                        
                        print("这个手动广告包含有这个好词的phrase")
                        if bulkfile_Campaign_SKUMax33  not in CountrySKU_Close_List:



                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"更改记录"]="确保这个phrase是开的"
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Campaign Status"]="enabled" #打开campaign status
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Ad Group Status"]="enabled" #打开Ad group status
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Status"]="enabled" #打开Keyword status


                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"更改记录"]="打开Campaign"
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                    #bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad")&bulkfile["SKU"]==bulkfile_Campaign_SKUMax2,"Status"]="enabled" #打开SKU的状态



                            
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Campaign")]) #追加Campaign              
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad Group")])#追加Ad Group
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad")])#追加Ad
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase")])#追加Keyword

 
                    else:#如果不存在if#MN-B
                #写一行生成一行phrase
                        print("这个手动广告不包含这个好词的phrase")

                        if bulkfile_Campaign_SKUMax33  not in CountrySKU_Close_List:
                            Bulkfile_SearchTerm_Add_manual_draft=Bulkfile_SearchTerm_Add_manual_draft.append({"Record Type":"Keyword","Campaign":Stoi_Campaign,"Ad Group":Stoi_AdGroup,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase","更改记录":"新增phrase"},ignore_index = True)


        Bulkfile_SearchTerm_Add=pd.concat([Bulkfile_SearchTerm_Add_auto,Bulkfile_SearchTerm_Add_auto_draft,Bulkfile_SearchTerm_Add_manual,Bulkfile_SearchTerm_Add_manual_draft,Bulkfile_SearchTerm_Add,Bulkfile_SearchTerm_Add_draft99],ignore_index=True) #追加到Bulkfile_SearchTerm_Add                

             
         
        Bulkfile_SearchTerm_Add.to_excel(r'D:\\运营\\4行动表\\Bulkfile_SearchTerm\\'+'Good_'+str(datetime.date.today())+countryname+"Bulkfile_SearchTerm"+".xlsx",index=False)
             

                     #追加到Bulkfile_SearchTerm_Add 

                    
        
            


            
    
        
                                                                                              
    else:
        print("没有"+str(countryname)+"Bulkfile文件")
 

                                                  
################################################################Bad#########################################################################################################################
 
Bulkfile_SearchTerm_Add=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
Bulkfile_SearchTerm_Add_auto=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])

Bulkfile_SearchTerm_Add_auto_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
Bulkfile_SearchTerm_Add_manual=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
Bulkfile_SearchTerm_Add_manual_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])


SearchTermAll_Bad=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]<0.05)&(SearchTermAll_Sum["Clicks"]>=20)]
SearchTermAll_Bad_Country_list=SearchTermAll_Bad["COUNTRY"].drop_duplicates().to_list()


for countryname22 in SearchTermAll_Bad_Country_list:
    SearchBadwithCountry=SearchTermAll_Bad[SearchTermAll_Bad["COUNTRY"]==countryname22] 
    
    if len(SearchBadwithCountry)>0:    
        n=0                                             
        for bulkdatafile in os.listdir(bulkdatafilepath): #找bulkfile对应的国家文件

            Bulkfile_Country=bulkdatafile.split('_')[0]

            if Bulkfile_Country==countryname22:

                bulkfile=pd.read_excel(bulkdatafilepath+str(bulkdatafile),engine="openpyxl",sheet_name=1)
                n+=1
                break
        if n==1:#F01
            print(bulkdatafile,bulkfile)
            bulkfile["更改记录"]=""

   
            for oi78 in range(len(SearchBadwithCountry)):
            
                oi78_keyword=SearchBadwithCountry.iloc[[oi78],[3]].values[0][0]
                print(oi78_keyword)
            
                oi78_Campaign=SearchBadwithCountry.iloc[[oi78],[1]].values[0][0]#获取第n行的关键词字段的值，遍历
                print(oi78_Campaign)
               
                oi78_AdGroup=SearchBadwithCountry.iloc[[oi78],[2]].values[0][0]
            #下面找到bulkfile中对应的广告进行关闭
             #看是手动还是自动？


            # 如果是手动：
            # 如果是手动：
                campaignlist=bulkfile["Campaign"].drop_duplicates().to_list()
                
                if oi78_Campaign not in campaignlist:
                    print("这个词找不到Campaign")
                    
                    continue
                else:
             
                    campaign78type=bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Record Type"]=="Campaign"),"Campaign Targeting Type"].values[0]
                    print(campaign78type)
                    if campaign78type=="Manual":

                
                        testlist78=bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact")] 
                        if len(testlist78)>0:
                            print("存在exact")
                            bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact"),"更改记录"]="差词暂停"
                            bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact"),"Status"]="paused"
                            #print(bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact"),"更改记录"])



                            Bulkfile_SearchTerm_Add_manual=pd.concat([Bulkfile_SearchTerm_Add_manual,bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact")]],ignore_index=True)
                            #Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_Campaign)&(bulkfile["Match Type"]=="exact")])
                            print("执行添加exact行")
                        
                             
                            
            #添加negtive+
                        Bulkfile_SearchTerm_Add_manual_draft=Bulkfile_SearchTerm_Add_manual_draft.append({"Record Type":"Keyword","Campaign":oi78_Campaign,"Ad Group":oi78_AdGroup,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":oi78_keyword,"Match Type":"negative exact","更改记录":"增加否定词"},ignore_index = True)
                    elif campaign78type=="Auto":
                
                        Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":oi78_Campaign,"Ad Group":oi78_AdGroup,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":oi78_keyword,"Match Type":"negative exact","更改记录":"增加否定词"},ignore_index = True)
            


            Bulkfile_SearchTerm_Add78=pd.concat([Bulkfile_SearchTerm_Add_auto_draft,Bulkfile_SearchTerm_Add_manual_draft,Bulkfile_SearchTerm_Add_manual],ignore_index=True) #追加到Bulkfile_SearchTerm_Add  
            Bulkfile_SearchTerm_Add78.to_excel(r'D:\\运营\\4行动表\\Bulkfile_SearchTerm\\'+'Bad_'+str(datetime.date.today())+countryname22+"Bulkfile_SearchTerm"+".xlsx",index=False)   
                
 

        else:#F01
            print("没有"+str(countryname22)+"Bulkfile文件")    
 
