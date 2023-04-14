
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np

Country_Adgroup_Maxid={"NEW-JP":100,"GV-MX":20,"NEW-MX":20,"GV-US":0.99,"GV-CA":0.99,"NEW-US":0.99,"NEW-CA":0.99,"NEW-UK":0.99,"NEW-IT":0.99, "NEW-ES":0.99,"NEW-DE":0.99,"NEW-DE":0.99,"HM-US":0.99}
Country_Keyword_Maxbid={"NEW-JP":80,"GV-MX":10,"NEW-MX":10,"GV-US":0.99,"GV-CA":0.88,"NEW-CA":0.88,"NEW-UK":0.88,"NEW-IT":0.88, "NEW-ES":0.88,"NEW-DE":0.88,"NEW-DE":0.88,"HM-US":0.88}
Country_DailyBudget={"NEW-JP":300,"GV-MX":300,"NEW-MX":300,"GV-US":3,"GV-CA":3,"NEW-CA":3,"NEW-UK":3,"NEW-IT":3, "NEW-ES":3,"NEW-DE":3,"NEW-DE":3,"HM-US":3}
bulkdatafilepath = 'D:\\运营\\1数据源\周bulk广告数据\\'

AllCountryActions=pd.read_excel(r'D:\\运营\\3数据分析结果\\国家汇总.xlsx',sheet_name="ProductActions")

AllCountryActions_CountryList=AllCountryActions["COUNTRY"].drop_duplicates().to_list()

for AllCountryActions_Country in AllCountryActions_CountryList:
    print("现在处理"+AllCountryActions_Country)
    n=0                                             
    for bulkdatafile in os.listdir(bulkdatafilepath): #找bulkfile对应的国家文件

        Bulkfile_Country=bulkdatafile.split('_')[0]

        if Bulkfile_Country==AllCountryActions_Country:#if A1
            bulkfile_draft_1Country=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status", "Bidding strategy","Placement Type","Increase bids by placement","变更记录"]) 

            bulkfile=pd.read_excel(bulkdatafilepath+str(bulkdatafile),engine="openpyxl",sheet_name=1)
            bulkfile["变更记录"]=""
            bulkfile["之前状态"]=" "
            print("找到了要处理的文件",bulkfile)
            
#############################################################确保关闭##############################################################
            AllCountryActions_Country_SKU_Close_List=AllCountryActions.loc[(AllCountryActions["COUNTRY"]==AllCountryActions_Country)&(AllCountryActions["行动方案"].str.contains("关闭广告")),"SKU"].drop_duplicates().to_list()

            if len(AllCountryActions_Country_SKU_Close_List)>0:
                print("执行关闭广告"+str(AllCountryActions_Country_SKU_Close_List))

                
                for AllCountryActions_Country_SKU_Close in AllCountryActions_Country_SKU_Close_List:
                    print("执行关闭广告循环")

                    
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close),"变更记录"]="执行关闭广告"
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close)&(bulkfile["Status"]=="enabled"),"之前状态"]="之前为开" 
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close),"Status"]="paused"

                    
                    bulkfile_SKU_Campaign_List_Close=bulkfile.loc[(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close)&(bulkfile["Record Type"]=="Ad"),"Campaign"].drop_duplicates().to_list()
                                                                  
                    for bulkfile_SKU_Campaign_Close in bulkfile_SKU_Campaign_List_Close:
                        bulkfile_SKU_Campaign_Close_Adgroup_List=bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group"),"Ad Group"].drop_duplicates().to_list()
                        
                        for bulkfile_SKU_Campaign_Close_Adgroup in bulkfile_SKU_Campaign_Close_Adgroup_List:
                            bulkfile_SKU_Campaign_Close_Adgroup_Ad_List=bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup),"SKU"].drop_duplicates().to_list()
                            if len(bulkfile_SKU_Campaign_Close_Adgroup_Ad_List)==1:
                                "要关闭的Ad Group只有一个SKU"
                            
                                bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup)&(bulkfile["Ad Group Status"]=="enabled"),"之前状态"]="之前为开"
                                bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup),"变更记录"]="执行关闭：Ad Group中SKU唯一，Ad Group也关掉"
                                bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup),"Ad Group Status"]="paused"
                            elif len(bulkfile_SKU_Campaign_Close_Adgroup_Ad_List)>1:
                                bulkfile_SKU_Campaign_Close_AdGroup_Ad_enabledList=bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup)&(bulkfile["Status"]=="enabled"),"Status"].drop_duplicates().to_list()
                                if len(bulkfile_SKU_Campaign_Close_AdGroup_Ad_enabledList)==0:
                                    bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup)&(bulkfile["Ad Group Status"]=="enabled"),"之前状态"]="之前为开"
                                    bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup),"变更记录"]="执行关闭：Ad Group中SKU都关了，Ad Group也关掉"
                                    bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group"]==bulkfile_SKU_Campaign_Close_Adgroup),"Ad Group Status"]="paused"
         
                        
                        bulkfile_SKU_Campaign_Close_AdGroup_enabledList=bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Ad")&(bulkfile["Status"]=="enabled"),"Ad Group"].drop_duplicates().to_list()
                        if len(bulkfile_SKU_Campaign_Close_AdGroup_enabledList)==0:
                            
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Campaign"),"变更记录"]="执行关闭：Campaign中Ad Group 都关了，也关掉"
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Campaign Status"]=="enabled"),"之前状态"]="之前为开" 
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfile_SKU_Campaign_Close)&(bulkfile["Record Type"]=="Campaign"), "Campaign Status"]="paused"



            else:
                 print("这个国家没有确保关闭广告的SKU,跳过循环继续找其他国家")
            
    
            AllCountryActions_Country_SKU_List=AllCountryActions.loc[(AllCountryActions["COUNTRY"]==AllCountryActions_Country)&(AllCountryActions["行动方案"].str.contains("确保广告是开的")),"SKU"].drop_duplicates().to_list()
            if len(AllCountryActions_Country_SKU_List)>0:
                print("这个国家有要开广告的SKU，继续执行")
                
##################################################################################################3333

                for AllCountryActions_Country_SKU in AllCountryActions_Country_SKU_List:
                    pr5_SKU="Pr5"+"-"+str(AllCountryActions_Country_SKU)
                
             #开始遍历SKU
                    print("正在看的SKU",AllCountryActions_Country_SKU)
                       

                
                    SKU_AdGroup_List=bulkfile.loc[bulkfile["SKU"]==AllCountryActions_Country_SKU,"Ad Group"].drop_duplicates().to_list()
                
                    if len(SKU_AdGroup_List)==0:
                        
                        print("没有这个产品对应的Ad Group")
                        print("生成一个")
                        print("先生成一个自动")

                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Campaign","Campaign":"Auto"+str(pr5_SKU),"Campaign Daily Budget":3,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Auto","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","变更记录":"生成一个自动"},ignore_index = True)
        
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad Group","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Adgroup_Maxid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","变更记录":"生成一个自动"},ignore_index = True)
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"SKU":AllCountryActions_Country_SKU,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","变更记录":"生成一个自动"},ignore_index = True)
                    
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"close-match","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"loose-match","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"complements","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"substitutes","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)

                          ############# 生成自动广告



                        print("再生成一个手动广告")

                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Campaign","Campaign":pr5_SKU,"Campaign Daily Budget":Country_DailyBudget[Bulkfile_Country],"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","变更记录":"再生成一个手动"},ignore_index = True)
        
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad Group","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Adgroup_Maxid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","变更记录":"再生成一个手动"},ignore_index = True)
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"SKU":AllCountryActions_Country_SKU,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","变更记录":"再生成一个手动",},ignore_index = True)
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"手工加词","Match Type":"exact","变更记录":"再生成一个手动"},ignore_index = True)
                        bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"手工加词","Match Type":"phrase","变更记录":"再生成一个手动"},ignore_index = True)
                    
                    else:
                        print("有这个产品对应的Ad Group")
                        for SKU_AdGroup in SKU_AdGroup_List:
                            print("处理这个Ad Group",SKU_AdGroup)
                            SKU_AdGroup_Campaign_List=bulkfile.loc[(bulkfile["Ad Group"]==SKU_AdGroup)&(bulkfile["SKU"]==AllCountryActions_Country_SKU)&(bulkfile["Record Type"]=="Ad"),"Campaign"].drop_duplicates().to_list()
                            #这个Ad Group对应的所有Campaign
                            print(str(SKU_AdGroup)+"这个Ad Group对应的所有Campaign",SKU_AdGroup_Campaign_List)
                        
                            for SKU_AdGroup_Campaign in SKU_AdGroup_Campaign_List:#遍历所有这个Ad Group对应的所有Campaign
                        
                                SKU_AdGroup_Campaign_Type_list=bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["SKU"]==AllCountryActions_Country_SKU)&(bulkfile["Record Type"]=="Campaign"),"Match Type"].drop_duplicates().to_list()

                                SKU_AdGroup_Campaign_Adgroup_List=bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["SKU"]==AllCountryActions_Country_SKU)&(bulkfile["Record Type"]=="Ad"),"Ad Group"].drop_duplicates().to_list()
                                for SKU_AdGroup_Campaign_Adgroup in SKU_AdGroup_Campaign_Adgroup_List:
                                    print(SKU_AdGroup_Campaign_Adgroup)
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign"),"变更记录"]="打开所有广告"
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Campaign Status"]=="paused"),"之前状态"]="之前状态为paused"
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign"),"Campaign Status"]="enabled"
                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad Group"),"变更记录"]="打开所有广告"
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Campaign Status"]=="paused"),"之前状态"]="之前状态为paused"
                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad Group"),"Campaign Status"]="enabled"
                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad Group"),"变更记录"]="打开所有广告"
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Ad Group Status"]=="paused"),"之前状态"]="之前状态为paused"
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad Group"),"Ad Group Status"]="enabled"

                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"变更记录"]="打开所有广告"
                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Campaign Status"]=="paused"),"之前状态"]="之前状态为paused"
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Ad Group Status"]=="paused"),"之前状态"]="之前状态为paused"
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Status"]=="paused"),"之前状态"]="之前状态为paused"
                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Campaign Status"]="enabled"
                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Ad Group Status"]="enabled"
                                    
                                    
                                    bulkfile.loc[(bulkfile["Campaign"]==SKU_AdGroup_Campaign)&(bulkfile["Ad Group"]==SKU_AdGroup_Campaign_Adgroup)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Status"]="enabled"
                                    print("打开一切")  

                                
                          #检查广告是手动还是自动
                            SKU_AdGroup_Campaign_List=bulkfile.loc[(bulkfile["Ad Group"]==SKU_AdGroup)&(bulkfile["SKU"]==AllCountryActions_Country_SKU)&(bulkfile["Record Type"]=="Ad"),"Campaign"].drop_duplicates().to_list()        
                            SKU_AdGroup_Campaign_Type_list=bulkfile.loc[(bulkfile["Campaign"].isin(SKU_AdGroup_Campaign_List))&(bulkfile["Record Type"]=="Campaign"),"Match Type"].drop_duplicates().to_list()
                            if len(SKU_AdGroup_Campaign_Type_list)==1 and SKU_AdGroup_Campaign_Type_list[0]=="Manual":
                         
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Campaign","Campaign":pr5_SKU,"Campaign Daily Budget":3,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Auto","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
        
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad Group","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Adgroup_Maxid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"SKU":AllCountryActions_Country_SKU,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled"},ignore_index = True)
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"close-match","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"补充自动广告"},ignore_index = True)
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"loose-match","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"补充自动广告"},ignore_index = True)
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"complements","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"补充自动广告"},ignore_index = True)
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":pr5_SKU,"Ad Group":pr5_SKU,"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"substitutes","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"补充自动广告"},ignore_index = True)


                            elif (len(SKU_AdGroup_Campaign_Type_list)==1) and (SKU_AdGroup_Campaign_Type_list[0]=="Auto") :
                         
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Campaign","Campaign":"Auto"+str(pr5_SKU),"Campaign Daily Budget":Country_DailyBudget[Bulkfile_Country],"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
        
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad Group","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Adgroup_Maxid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                         
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"SKU":AllCountryActions_Country_SKU,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled"},ignore_index = True)
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"手动填写词","Match Type":"exact","变更记录":"补充手动广告"},ignore_index = True)
                                bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"手动填写词","Match Type":"phrase","变更记录":"补充手动广告"},ignore_index = True)
  

            
            else:
                print("这个国家"+str(Bulkfile_Country)+"没有要确保开的广告")
                
            bulkfile=pd.concat([bulkfile,bulkfile_draft_1Country],ignore_index=True)
            bulkfile=bulkfile.drop(columns=["Clicks","Spend","Orders","Sales","ACoS"])
            bulkfile=bulkfile.drop_duplicates()

            bulkfile.to_excel(r'D:\\运营\\4行动表\\调整SKU广告开关\\'+'_'+str(datetime.date.today())+Bulkfile_Country+"AdjustAdvertisement"+".xlsx",index=False)
              
            break    #A1  
        else:
            listdir=os.listdir(bulkdatafilepath)
            n+=1
            if n==len(listdir):
                print("没有对应国家的Bulkfile文件")
               
               
       
    
    
    
