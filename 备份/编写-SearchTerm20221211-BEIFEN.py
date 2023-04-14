
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np


bulkdatafilepath = 'D:\\运营\\周bulk广告数据\\'

#底线策略
SearchTermAll=pd.read_excel(r'D:\\运营\\运行结果数据\\Sponsored Products Search term report.xlsx')

SearchTermAll["Clicks"].fillna(0,inplace=True)



SearchTermAll_Sum=SearchTermAll.groupby(["COUNTRY","Campaign Name", "Ad Group Name","Customer Search Term"],as_index=False)[["Impressions","Clicks","Spend","7 Day Total Sales ","7 Day Total Orders (#)"]].agg("sum")


SearchTermAll_Sum.loc[SearchTermAll_Sum['Clicks']>0,"转化率"]=SearchTermAll_Sum["7 Day Total Orders (#)"]/SearchTermAll_Sum['Clicks']


SearchTermAll_Good=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]>0.2)]

SearchTermAll_Bad=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]<0.033)&(SearchTermAll_Sum["Clicks"]>30)]






#打开Bulkfile的第一种方法

Allbulkpath='D:\\运营\\'  
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')


Allbulk_Week1=Allbulk[Allbulk["周数"]==1]

#遍历SeachTermGood

Allbulk_Campaign_SKUMax=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('sum')


SearchTermAll_Good_Country_list=SearchTermAll_Good["COUNTRY"].drop_duplicates().to_list()



################################################################GOOD#########################################################################################################################

for countryname in SearchTermAll_Good_Country_list:   #遍历searchTemgood里的国家
    
    Bulkfile_SearchTerm_Add=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend",	"Orders","Total units",	"Sales	ACoS","Bidding strategy","Placement Type","Increase bids by placement"])
    
    SearchTermAll_Good_Country=SearchTermAll_Good[SearchTermAll_Good["COUNTRY"]==str(countryname)]
    n=0                                             
    for bulkdatafile in os.listdir(bulkdatafilepath): #找bulkfile对应的国家文件
        Bulkfile_Country=bulkdatafile.split('_')[0]

        if Bulkfile_Country==countryname:
            bulkfile=pd.read_excel(bulkdatafilepath+str(bulkdatafile),engine="openpyxl",sheet_name=1)
            n+=1
            break
    if n==1:
        print(bulkdatafile,bulkfile)
        Bulkfile_SearchTerm_Add_auto=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend",	"Orders","Total units",	"Sales	ACoS","Bidding strategy","Placement Type","Increase bids by placement"])

        Bulkfile_SearchTerm_Add_auto_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend",	"Orders","Total units",	"Sales	ACoS","Bidding strategy","Placement Type","Increase bids by placement"])
        Bulkfile_SearchTerm_Add_Manual=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend",	"Orders","Total units",	"Sales	ACoS","Bidding strategy","Placement Type","Increase bids by placement"])
        Bulkfile_SearchTerm_Add_Manual_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend",	"Orders","Total units",	"Sales	ACoS","Bidding strategy","Placement Type","Increase bids by placement"])



        for Stoi in range(0,len(SearchTermAll_Good_Country)):#在筛选国家后遍历行

    
            Stoi_keyword=SearchTermAll_Good_Country.iloc[[Stoi],[4]].values[0][0]
            Stoi_Campaign=SearchTermAll_Good_Country.iloc[[Stoi],[1]].values[0][0]#获取第n行的关键词字段的值，遍历
            Stoi_AdGroup=SearchTermAll_Good_Country.iloc[[Stoi],[2]].values[0][0]
        

                                                  
            avearageprice=SearchTermAll_Good_Country.iloc[[Stoi],[6]]/SearchTermAll_Good_Country.iloc[[Stoi],[5]]
                                                                                             

    

       
#####################################Auto###################################1
                                                                                              
            if bulkfile.loc[(bulkfile['Record Type']=="Campaign")&(bulkfile['Campaign']==Stoi_Campaign),"Campaign Targeting Type"].value[0][0]=="Auto" :#Camaign是否为自动
                #不要了Camaign_sku_List=bulkfile[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Ad Group"]==Stoi_AdGroup)&,bulkfile["Record Keytype"]=="Ad","SKU"].to_list()
                                                                                              
                
                #找到历史花费最多的SKU
                ##这个方法备用bulkfile_Campaign_SKUMax1=Allbulk_Campaign_SKUMax.groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('max')                                                                          
                bulkfile_Campaign_SKUMax2=Allbulk_Campaign_SKUMax[(Allbulk_Campaign_SKUMax["Campaign"]==Stoi_Campaign) &(Allbulk_Campaign_SKUMax['Record Type']=="Ad")&(Allbulk_Campaign_SKUMax['Ad Group']==Stoi_AdGroup)&(Allbulk_Campaign_SKUMax['Spend'].max()),"SKU"].value(0,0)                                                                   
                pr4_sku="Pr4"+"-"+str(bulkfile_Campaign_SKUMax2)


                    #上面为找出这个自动广告中花费最多的SKU
                bulkfileSKU_Manual_Campaign_List=bulkfile.loc[(bulkfile['Record Type']=="Ad")&(bulkfile['Ad Group']==Stoi_AdGroup)&(bulkfile['SKU']==bulkfile_Campaign_SKUMax2)&bulkfile['Campaign Targeting Type']=="Manual","Campaign"].to_list()
                    #上面为这个花费最多的SKU对应的Bulkfile中的所有Campaign的list
                
                for bulkfileSKU_Manual_Campaign in bulkfileSKU_Manual_Campaign_List:#遍历为这个花费最多的SKU对应的Bulkfile中的所有Campaign的list
                       
#############################################################11                                                          
                    if bulkfileSKU_Manual_Campaign.startswith('Pr4'):#如果有专门为关键词累计广告开的Pr4开头的广告（也可以找手动广告，但为了清晰，不计划加到之前的人工手动广告）
                    
##################                    
                        if Stoi_keyword in (bulkfile[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Keyword product targeting"].duplicates().to_list()):
                    ###上：如果有这个词的exact广告
                        
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")&bulkfile["SKU"]==bulkfile_Campaign_SKUMax2,"Status"]="enabled" #打开SKU的状态
                                        
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Status"]="enabled"#打开关键词的状态

                        
                            Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Campaign")]) #把这个Pr4广告的Campaign行加到变更文件Bulkfile_SearchTerm_Add_auto里              
                            Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad Group")]) #把这个Pr4广告的Ad Group行加到变更文件Bulkfile_SearchTerm_Add_auto里
                            Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")])#把这个Pr4广告的Ad行加到变更文件Bulkfile_SearchTerm_Add_auto里
                            Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])#把这个Pr4广告的keyword行加到变更文件Bulkfile_SearchTerm_Add_auto里


##################

                        else:#没有这个词对应SKU的Pr4 的exact

#生成一个pr4放到 Bulkfile_SearchTerm_Add_auto_draft
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Campaign","Campaign":pr4_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword ","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":avearageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)
           
#生成完毕





                
                
 



                    
                   
                        if Stoi_keyword in (bulkfile[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Keyword product targeting"].duplicates().to_list()):
###########如果存在这个SKU的Pr4广告有这个词的phrase

                                        
                            bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Status"]="enabled"
                        #把存在的phrase广告状态变成enabled
                        
                        #Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Campaign")])               
                        #Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad Group")])
                        #Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")])
                            Bulkfile_SearchTerm_Add_auto=Bulkfile_SearchTerm_Add_auto.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])
                        #把这个广告加到Bulkfile_SearchTerm_Add_auto
                        

###################################################################1
          
                        else:       #如果没有phrase
#生成一行phrase                    
                              Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword ","Campaign":bulkfileSKU_Manual_Campaign,"Ad Group":bulkfileSKU_Manual_Campaign,"Max Bid":avearageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase"},ignore_index = True)





                Bulkfile_SearchTerm_Add=pd.conact([Bulkfile_SearchTerm_Add,Bulkfile_SearchTerm_Add_auto],ignore_index=True) #追加到Bulkfile_SearchTerm_Add
                Bulkfile_SearchTerm_Add=pd.conact([Bulkfile_SearchTerm_Add,Bulkfile_SearchTerm_Add_auto_draft],ignore_index=True)
#####################################Auto###################################1end


########################################Manual#######################################2
                                                                                              
           
          
            else: #如果来自手动:则判断在这个Stoi_Campaign中    
#下面判断bulkfile中 含有Stoi_keyword的exacT
#################################################EXACT               
                if Stoi_keyword in (bulkfile[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Keyword product targeting"].duplicates().to_list()):                                                                         
                 #如果存在exact广告
                
                    bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Campus Status"]="enabled"
                    bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Ad Group Status"]="enabled"
                    bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Status"]="enabled"



                    bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。
                    bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                    bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")&bulkfile["SKU"]==bulkfile_Campaign_SKUMax2,"Status"]="enabled" #打开SKU的状态


                

                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Campaign")]) #追加Campaign              
                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad Group")])#追加Ad Group
                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")])#追加Ad
                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_amanual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])#追加Keyword

                
                else: #如果没有exac
                                               
                 #生成一个
                    Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual.append({"Record Type":"Keyword ","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":avearageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)



                    Maxbid_Adgroup=0.99
                    Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Campaign","Campaign":pr4_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
                 #生成Campaign 并追加
                    Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                 #生成Ad Group 并追加
                    Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                 #生成Ad SKU 并追加
                    Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Keyword ","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":avearageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)
                 #生成keyword 并追加

#################################################phrase               
                if Stoi_keyword in (bulkfile[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Keyword product targeting"].duplicates().to_list()):
         #如果存在一个phrase的广告
                    bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Campus Status"]="enabled" #打开campaign status
                    bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Ad Group Status"]="enabled" #打开Ad group status
                    bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Status"]="enabled" #打开Keyword status
                    bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。
                    bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                    bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")&bulkfile["SKU"]==bulkfile_Campaign_SKUMax2,"Status"]="enabled" #打开SKU的状态
                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Campaign")]) #追加Campaign              
                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad Group")])#追加Ad Group
                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")])#追加Ad
                    Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword product targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase")])#追加Keyword

 
                else:#如果不存在
                #写一行生成一行phrase
                    Bulkfile_SearchTerm_Add_Manual_draft=Bulkfile_SearchTerm_Add_Manual_draft.append({"Record Type":"Keyword ","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":avearageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase"},ignore_index = True)
                
                

                
            
                Bulkfile_SearchTerm_Add=pd.conact([Bulkfile_SearchTerm_Add,Bulkfile_SearchTerm_Add_manual],ignore_index=True)
                Bulkfile_SearchTerm_Add=pd.conact([Bulkfile_SearchTerm_Add,Bulkfile_SearchTerm_Add_manual_draft],ignore_index=True)
            
                #追加到Bulkfile_SearchTerm_Add 

                    
           
    
        Bulkfile_SearchTerm_Add.to_excel(r'D:\\运营\\分析结果\\Bulkfile_SearchTerm\\'+str(datetime.date.today())+countryname+"Bulkfile_SearchTerm",index=False)
                                                                                              
    else:
        print("没有"+str(countryname)+"Bulkfile文件")
 


                                                  
################################################################Bad#########################################################################################################################
                                                  
                                                  
                                                
