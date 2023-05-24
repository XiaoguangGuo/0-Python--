# 使用Oenyxl编写，没有使用pandas
### 第一部分汇总各国bulk广告数据到汇总表
#用于实际使用，从零开始导入10周最新的。创建新的文件，存到指定文件名。：测试OK。

# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np


ProductActions=pd.read_excel(r'D:\\运营\\3数据分析结果\\国家汇总.xlsx',sheet_name="ProductActions")
#在这之前要生成汇总表，并且把每个国家的bulk表备份到D:\\运营\\bulkoperationfiles\\
newdate=input('输入最新日期y-m-d',) #输入最新一周的日期
maxtime=datetime.datetime.strptime(newdate,'%Y-%m-%d')
print(newdate)
print(maxtime)

Allbulkpath='D:\\运营\\2生成过程表\\'  
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')
Allbulkold1=pd.read_excel(r'D:\\运营\\1数据源\\周Bulk广告数据汇总表历史\\'+"周bulk广告数据汇总表_2022-8-27_2022-9-24.xlsx")
Allbulkold2=pd.read_excel(r'D:\\运营\\1数据源\\周Bulk广告数据汇总表历史\\'+"周bulk广告数据汇总表_2022-5-28_2022-8-20.xlsx")



Allbulk=pd.concat([Allbulk,Allbulkold1,Allbulkold2])
Allbulk["周数"]=(maxtime-Allbulk["日期"]).dt.days//7+1


#！！！！筛选汇总表的数据---后续可以按日期>某个日期来筛选
AllbulkD5=Allbulk[(Allbulk['Keyword or Product Targeting']. notna())&(Allbulk['周数']<26)]#定义了26周汇总
print(AllbulkD5)

# AllbulkCampaignKeyword1Week=Allbulk[(Allbulk['Keyword or Product Targeting'].notna())&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')

#AbulkCampaignKeyword1week["zhuanhualv"]=AllbulkCampaignKeyword1week['Orders']/AllbulkCampaignKeyword1week['Clicks']


AllbulkCampaignKeyword=AllbulkD5.groupby(["Country","Campaign","Keyword or Product Targeting","Match Type","Ad Group"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum") 
#AllbulkCampaignKeyword所有周数小于AllbulkD5定义的历史数据汇总

print(AllbulkCampaignKeyword.columns)

maxrow=len(AllbulkCampaignKeyword)
#maxrow1Week=len(AllbulkCampaignKeyword1Week)
AllbulkCampaignKeyword["zhuanhualv"]=AllbulkCampaignKeyword['Orders']/AllbulkCampaignKeyword['Clicks']
print(maxrow)
print(AllbulkCampaignKeyword)

AllbulkCampaignKeyword.to_excel(Allbulkpath+'周bulk广告数据汇总表huizong.xlsx') #生成分析用的报表



#从国家-bulkopration遍历
bulkoperationfilespath = 'D:\\运营\\1数据源\\bulkoperationfiles\\' #待处理的bulk文件--下载下来作为上传的基础的
bulkoperationfilesnewpath='D:\\运营\\4行动表\\bulkoperationfilesNEW\\'#修改后的bulk文件目录
countrylist1=["NEW-JP","NEW-IT","GV-MX","NEW-MX","NEW-FR","NEW-DE","NEW-ES"]
bulkfilecountrylist=[]#国家空列表
for bulkoperationfile in os.listdir(bulkoperationfilespath):
    clickatempt=10
    clickenough1=12
    clickenough2=16
    #提取国家名
    bulkfileCountry=os.path.basename(bulkoperationfile).split('_')[0]#遍历国家名
  
    print(bulkfileCountry,type(bulkfileCountry))
  
    
    bulkfilecountrylist.append(bulkfileCountry) #国家名追加到成国家清单
   
    print(bulkfilecountrylist) 

    dfbulkfile=pd.read_excel(bulkoperationfilespath+str(bulkoperationfile),engine="openpyxl",sheet_name=1)#用PD读取待操作的某国bulk文件。
    #dfbulkfile.replace(",",".",inplace=True)#替换逗号为标点，这个操作可能无效。
    dfbulkfile["更新内容"]=""
    print(dfbulkfile)                            


    #变更后要做记录，下面这个模块就是处理这个部分。
    bulkrecordpath = 'D:\\运营\\2生成过程表\\bulk变更记录\\' #所有变更记录在国家的一张表 
    if not os.listdir(bulkrecordpath):
        print("文件夹为空")
    
        dfbulkmodifyrecord=pd.DataFrame()
        print(dfbulkmodifyrecord)
        dfbulkmodifyrecord=dfbulkmodifyrecord.append(dfbulkfile)#第一次没有变更记录就生成一张表
        print("新的记录表",dfbulkmodifyrecord)
    else: 
 
        print("文件夹是有文件",os.listdir(bulkrecordpath))
        #遍历记录文件夹
        for bulkmodifyrecordfile in os.listdir(bulkrecordpath):#
            print("os:",os.path.basename(bulkmodifyrecordfile).split('_')[0],"str:",str(os.path.basename(bulkmodifyrecordfile).split('_')[0]))
            
            if os.path.basename(bulkmodifyrecordfile).split('_')[0]==str(bulkfileCountry):
               yanzhnegfu=1
               bulkmodifyrecordfilefound=os.path.basename(bulkmodifyrecordfile) 
                
               print("有匹配os.path.basename(bulkmodifyrecordfile)")
               break
        if yanzhnegfu==1:
            print("文件夹内有变更记录")
            dfbulkmodifyrecord=pd.read_excel(bulkrecordpath+str(bulkmodifyrecordfilefound),engine="openpyxl",sheet_name=0)
            print(dfbulkmodifyrecord)
            dfbulkmodifyrecord=pd.concat([dfbulkmodifyrecord,dfbulkfile])#这一周有新增的内容加到变更记录表中
            dfbulkmodifyrecord.drop_duplicates(subset=["Record Type","Campaign","Campaign Targeting Type","Product Targeting ID","Ad Group","Keyword or Product Targeting","Product Targeting ID" ,"Match Type",  "SKU","Bidding strategy","Placement Type" ],inplace=True)
            #去除重复项
        else:
             
            print("文件夹内没有变更记录")
        
            dfbulkmodifyrecord=pd.DataFrame()
            print(dfbulkmodifyrecord)
            dfbulkmodifyrecord=dfbulkmodifyrecord.append(dfbulkfile)#生成一个记录表
            print("新的记录表",dfbulkmodifyrecord)
        
  
###############为记录表添加新列      
         
    Today= datetime.datetime.today().strftime('%Y-%m-%d')
    Columnnew1=Today+"_"+"最新状态"
    Columnnew2=Today+"_"+"更新内容"
    print(Today,Columnnew1,Columnnew2)
    dfbulkmodifyrecord[Columnnew1]=""
    dfbulkmodifyrecord[Columnnew2]=""
    print(dfbulkmodifyrecord.columns)
    
#########################
  
######定义各种判断条件：用筛选的方式
    
    #####可以自合df筛选条件的字符串


    #bulkcondiion_new01='AllbulkCampaignKeyword["Orders"]<(AllbulkCampaignKeyword["Clicks"]/clickenough1).apply(np.floor))'
    bulcountryconditionstr='AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"=="bulkfileCountry")'                            
   
    bulkcountrycondionclk1str='(AllbulkCampaignKeyword["Clicks"]>(clickatempt-1)) &(AllbulkCampaignKeyword["Clicks"]<(clickenough+1))]'
    bulkcountrycondionclk21str='AllbulkCampaignKeyword["Clicks"]>clickenough'
    bulkcountrycondionzhl01str= '(AllbulkCampaignKeyword["zhuanhualv"]==0]'
    bulkcountrycondionzhl02str= '(AllbulkCampaignKeyword["zhuanhualv"]>0.2]'
    bulkcountrycondionzhl032str='(AllbulkCampaignKeyword["zhuanhualv"]>0.5]'
    

    
    #bulkconditiondic={}#

####################################################################################################3
    
##1 在y筛选出第一种操作的对象条件表: #做了reindex
    if bulkfileCountry in countrylist1:
        

        
        bocondition_new01df=AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"]==str(bulkfileCountry)) & (AllbulkCampaignKeyword["Clicks"]>0)&(AllbulkCampaignKeyword["Orders"]<(AllbulkCampaignKeyword["Clicks"]/clickenough2).apply(np.floor))].reset_index(drop=True)
    #符合第一个条件是有点击且订单数>点击数/16.#取整替代方法尝试(AllbulkCampaignKeyword["Clicks"]/clickenough2).round(0) #.astype("int")
    else:
        bocondition_new01df=AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"]==str(bulkfileCountry)) & (AllbulkCampaignKeyword["Clicks"]>0)&(AllbulkCampaignKeyword["Orders"]<(AllbulkCampaignKeyword["Clicks"]/clickenough1).apply(np.floor))].reset_index(drop=True)   
                                                                                                             
   
    print("bocondition_new01df",bocondition_new01df)#打印符合条件1的列表
  
    
####从第一行开始遍历这个符合条件1的筛选表
                                                                                                    
    for boi in range(0,len(bocondition_new01df)):
                                                                                                    
        boi_keyword=bocondition_new01df.iloc[[boi],[2]].values[0][0]#获取第n行的关键词字段的值，遍历
        print(boi_keyword,boi)
         
        boi_campaign=bocondition_new01df.iloc[[boi],[1]].values[0][0]#获取第n行的关键词Campaign的值，遍历                                                                                            
        boi_ADgroup=bocondition_new01df.iloc[[boi],[4]].values[0][0]                                                                                             


#将对bulkfile中符合筛选表1的camaign和词进行Status批量进行修改成暂停pause

        
         
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Status']!="paused")&(dfbulkfile['Ad Group']==boi_ADgroup),"更新内容"]="paused"
                                            
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Status']!="paused")&(dfbulkfile['Ad Group']==boi_ADgroup),"Status"]="paused"
        
        print(dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Status']=="paused")&(dfbulkfile['Ad Group']==boi_ADgroup)])
           
        print("01:将bulkfile中这个campaign-词的Status进行修改")
#将bocondition02df修改记录写入记录表
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign)&(dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword),Columnnew1]="bocondition_new01df"
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign)&(dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword),Columnnew2]="paused"

        
        print("01:将bulkfile中这个campaign-词的修改记录" )


##############################################################################################################3

###周转率大于0.2的保持enabled
    bocondition05df=AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"]==bulkfileCountry) &(AllbulkCampaignKeyword["zhuanhualv"]>0.13)&(AllbulkCampaignKeyword["Clicks"]>0)]
    CountrySKU_Close_List=ProductActions.loc[(ProductActions["Country"]==bulkfileCountry)&(ProductActions["行动方案"].str.contains("关闭广告")),"SKU"].drop_duplicates().to_list()
    for boi in range(0,len(bocondition05df)):
           #定义要找的词和Campaign
        boi_Matchtype=bocondition05df.iloc[[boi],[3]].values[0][0]
        boi_keyword=bocondition05df.iloc[[boi],[2]].values[0][0]
        boi_campaign=bocondition05df.iloc[[boi],[1]].values[0][0]
        boi_ADgroup=bocondition05df.iloc[[boi],[4]].values[0][0]

        #keyword_campaign=bocondition05df.loc[(dfbulkfile["Campaign"]==boi_campaign)&(dfbulkfile["Ad Group"]==boi_ADgroup),"Keyword or Product Targeting"]

########################################################################################################################

        SKUlist77=dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile["Record Type"]=="Ad"),"SKU"].to_list()
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
        CountrySKU_Close_List=AllCountryActions_Country_SKU_Close_List_nocomma_list


        count77=0
        listingopen77=[]
        for sku77 in SKUlist77:
            if sku77 in CountrySKU_Close_List:
                count77+=1
            else:   
                listingopen77+=[sku77]
        if len(CountrySKU_Close_List)==count77:
            break
        elif len(CountrySKU_Close_List)>count77:
            countnew99=0
            for open_SKU_oi in listingopen77:
                countnew99+=1
                if countnew99==1:
            
 
        #对于只要WORD是ause的，打开word status  # 如是open的就不去改变
        
                    dfbulkfile.loc[((dfbulkfile['Record Type']=="Keyword")|(dfbulkfile['Record Type']=="Product Targeting"))&(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Match Type']==boi_Matchtype)&(dfbulkfile['Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[((dfbulkfile['Record Type']=="Keyword")|(dfbulkfile['Record Type']=="Product Targeting"))&(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Match Type']==boi_Matchtype)&(dfbulkfile['Status']=="paused"),"Status"]="enabled"                               
       
        

        
#对于campaign和adgroup有一个是pause的，原来其他的词status为Pause
        
                #keyword_campaign=bocondition05df[(bocondition05df["Campaign"]==boi_campaign)&(bocondition05df["Ad Group"]==boi_ADgroup)]
                #Keyword_list=keyword_campaign['Keyword or Product Targeting'].tolist()
                #keyword_campaign_matchtype=bocondition05df[(bocondition05df["Campaign"]==boi_campaign)&(bocondition05df["Ad Group"]==boi_ADgroup)&(bocondition05df["Keyword or Product Targeting"]==boi_keyword)]
                #keyword_Matchtype_list=keyword_campaign_matchtype["Match Type"]
 
                #dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile["Ad Group"]==boi_ADgroup)&((dfbulkfile["Campaign Status"]=="paused")|(dfbulkfile["Ad Group Status"]=="paused"))& (~(dfbulkfile['Keyword or Product Targeting'].isin(Keyword_list))) &(dfbulkfile["Status"]=="enabled"),"更新内容"]="转化率>0.2，开启：无关词关闭"                                    
                #dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile["Ad Group"]==boi_ADgroup)&((dfbulkfile["Campaign Status"]=="paused")|(dfbulkfile["Ad Group Status"]=="paused"))& (~(dfbulkfile['Keyword or Product Targeting'].isin(Keyword_list))) &(dfbulkfile["Status"]=="enabled"), "Status"]="paused"
                       
                #dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile["Ad Group"]==boi_ADgroup)&((dfbulkfile["Campaign Status"]=="paused")|(dfbulkfile["Ad Group Status"]=="paused"))& (dfbulkfile['Keyword or Product Targeting']==boi_keyword)&(~(dfbulkfile['Match Type'].isin(keyword_Matchtype_list)))&(dfbulkfile["Status"]=="enabled"),"更新内容"]="转化率>0.2，开启：同词不同Match关闭 "    
        
                #dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile["Ad Group"]==boi_ADgroup)&((dfbulkfile["Campaign Status"]=="paused")|(dfbulkfile["Ad Group Status"]=="paused"))& (dfbulkfile['Keyword or Product Targeting']==boi_keyword ) &(~(dfbulkfile['Match Type'].isin(keyword_Matchtype_list)))&(dfbulkfile["Status"]=="enabled"),"Status"]="paused"
                                                                                                                                                                                                                                          


####记录
        #dfbulkmodifyrecord.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkmodifyrecord["Ad Group"]==boi_ADgroup)&(dfbulkmodifyrecord['Keyword or Product Targeting']!=boi_keyword),Columnnew2]="转化率>0.2，开启：无关词关闭 "
                                    
        #dfbulkmodifyrecord.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkmodifyrecord["Ad Group"]==boi_ADgroup)&(dfbulkmodifyrecord['Keyword or Product Targeting']!=boi_keyword),Columnnew1]="bocondition05df"
                
        
        #dfbulkmodifyrecord.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkmodifyrecord["Ad Group"]==boi_ADgroup)&(dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Match Type']==boi_Matchtype),Columnnew1]="bocondition05df"
        #dfbulkmodifyrecord.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkmodifyrecord["Ad Group"]==boi_ADgroup)&(dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword)&(dfbulkfile['Match Type']==boi_Matchtype),Columnnew2]="转化率>0.2，开启"

  
################### 打开Ad enabled##################################################################################################################################################3
                #dfbulkfile.loc[(dfbulkfile["更新内容"]=="转化率>0.2，开启：同词不同Match关闭 ")|(dfbulkfile["更新内容"]=="转化率>0.2，开启：无关词关闭"),"Campaign Status"]="enabled"
                #dfbulkfile.loc[(dfbulkfile["更新内容"]=="转化率>0.2，开启：同词不同Match关闭 ")|(dfbulkfile["更新内容"]=="转化率>0.2，开启：无关词关闭"),"Ad Group Status"]="enabled"                                                              

##########################################################################################################################################################################################
 
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Campaign Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Campaign Status']=="paused"),"Campaign Status"]="enabled"
   
        
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Ad Group Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Ad Group Status']=="paused"),"Ad Group Status"]="enabled"
        
                                             
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Status']=="paused"),"Status"]="enabled"
        
         
        #将bocondition05df修改记录写入记录表
        #dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) &(dfbulkmodifyrecord["SKU"].isin(sku))&(dfbulkfile["Ad Group"]==boi_ADgroup)& (dfbulkmodifyrecord['Record Type']=="Ad"),Columnnew1]="bocondition05df"
        #dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) &(dfbulkmodifyrecord["SKU"].isin(sku))&(dfbulkfile["Ad Group"]==boi_ADgroup)& (dfbulkmodifyrecord['Record Type']=="Ad"),Columnnew2]="转化率>0.2，开启"


        
############################### # 打开Ad Group enabled######################################################################################################################################

                
 
        
 
                                             
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group")&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Campaign Status']=="paused"),"更新内容"]="转化率>0.2，开启"                                        
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group")&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Campaign Status']=="paused"),"Campaign Status"]="enabled"
                                             
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group")&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Ad Group Status']=="paused"),"更新内容"]="转化率>0.2，开启"                    
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group")&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Ad Group Status']=="paused"),"Ad Group Status"]="enabled"
        
                               
        
        

#将bocondition05df修改记录写入记录表
        #dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Ad Group")&(dfbulkfile["Ad Group"]==boi_ADgroup),Columnnew1]="bocondition05df"
        #dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Ad Group")&(dfbulkfile["Ad Group"]==boi_ADgroup),Columnnew2]="转化率>0.2，开启"

        
######################################### # 打开Campaign enabled#######################################################################################################33
 
 



  
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Campaign")&(dfbulkfile['Campaign Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Campaign")&(dfbulkfile['Campaign Status']=="paused"),"Campaign Status"]="enabled"  
 
        
        
        
#将bocondition05df修改记录写入记录表
        #dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Campaign")&(dfbulkfile['Campaign Status']=="paused"),Columnnew1]="bocondition05df"
        #dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Campaign")&(dfbulkfile['Campaign Status']=="paused"),Columnnew2]="转化率>0.2，开启"        






        
 ###################################################################################################################################################################333


                elif  countnew99>1:
                ################### 打开Ad enabled##################################################################################################################################################3
  
##########################################################################################################################################################################################
 
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Campaign Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Campaign Status']=="paused"),"Campaign Status"]="enabled"
   
        
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Ad Group Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Ad Group Status']=="paused"),"Ad Group Status"]="enabled"
        
                                             
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Status']=="paused"),"更新内容"]="转化率>0.2，开启"
                    dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad")&(dfbulkfile['SKU']==open_SKU_oi)&(dfbulkfile["Ad Group"]==boi_ADgroup)&(dfbulkfile['Status']=="paused"),"Status"]="enabled"
        



        
#########a将修改过的文件存入bulkoperationfilesnewpath
    dfbulkfile_notnull=dfbulkfile[(dfbulkfile['更新内容'].notna()) & (dfbulkfile['更新内容'] !="")]#只保留#只保留
    
    dfbulkfile_notnull.to_excel(bulkoperationfilesnewpath+"Spl-Re_"+Today+"_"+str(bulkoperationfile),index=False)#生成简化版
    
    dfbulkfile.to_excel(bulkoperationfilesnewpath+"Re_"+Today+"_"+str(bulkoperationfile),index=False)

    
#########a将更改记录文件存入bulkrecordpath    
    #dfbulkmodifyrecord.to_excel(bulkrecordpath+bulkfileCountry+"_"+"更改记录"+".xlsx",index=False)
########将周bulk广告数据里的文件移动到历史文件夹    
    #shutil.move(r'D:\\运营\\周bulk广告数据\\'+ str(bulkdatafile), r'D:\运营\HistoricalData\周bulk广告数据')

#########3将bulkoperation里的文件删除
    os.remove(bulkoperationfilespath+ str(bulkoperationfile))
########或者可以移动将bulkoperation里的文件移动到历史文件夹
    #shutil.move(bulkoperationfilespath+ str(bulkdatafile), r'D:\运营\HistoricalData\周bulk广告数据\bulkoperationfiles')
 
#本次修改主要把要打开的广告限制在一定范围
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np

Country_Adgroup_Maxid={"NEW-JP":100,"GV-MX":20,"NEW-MX":20,"GV-US":0.99,"GV-CA":0.99,"NEW-US":0.99,"NEW-CA":0.99,"NEW-UK":0.99,"NEW-IT":0.99, "NEW-ES":0.99,"NEW-DE":0.99,"NEW-FR":0.99,"HM-US":0.99}
Country_Keyword_Maxbid={"NEW-JP":80,"GV-MX":10,"NEW-MX":10,"GV-US":0.8,"GV-CA":0.8,"NEW-CA":0.8,"NEW-UK":0.8,"NEW-US":"0.8","NEW-IT":0.8, "NEW-ES":0.8,"NEW-DE":0.8,"HM-US":0.8,"NEW-FR":0.8}
Country_DailyBudget={"NEW-JP":300,"GV-MX":300,"NEW-MX":300,"GV-US":3,"GV-CA":3,"NEW-CA":3,"NEW-UK":3,"NEW-FR":3,"NEW-US":3,"NEW-IT":3, "NEW-ES":3,"NEW-DE":3,"HM-US":3}
bulkdatafilepath = 'D:\\运营\\1数据源\周bulk广告数据\\'

AllCountryActions=pd.read_excel(r'D:\\运营\\3数据分析结果\\国家汇总.xlsx',sheet_name="ProductActions")

AllCountryActions_CountryList=AllCountryActions["Country"].drop_duplicates().to_list()
AllCountryActions.dropna(subset=["SKU"],inplace=True)
AllCountryActions["SKU"].astype(str)

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
            print("找到了要处理的文件")
            
#############################################################确保关闭##############################################################
            AllCountryActions_Country_SKU_Close_List=AllCountryActions.loc[(AllCountryActions["Country"]==AllCountryActions_Country)&(AllCountryActions["行动方案"].str.contains("关闭广告")),"SKU"].drop_duplicates().to_list()
            print(AllCountryActions_Country_SKU_Close_List)
            AllCountryActions_Country_SKU_Close_List_nocomma_list=[]
            for AllCountryActions_Country_SKU_Close_List_nocomma in AllCountryActions_Country_SKU_Close_List:
                print(AllCountryActions_Country_SKU_Close_List_nocomma)
                 
                
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
                    
                    
            AllCountryActions_Country_SKU_Close_List=AllCountryActions_Country_SKU_Close_List_nocomma_list
            print(AllCountryActions_Country_SKU_Close_List)
                

            if len(AllCountryActions_Country_SKU_Close_List)>0:
                print("执行关闭广告"+str(AllCountryActions_Country_SKU_Close_List))

                
                for AllCountryActions_Country_SKU_Close in AllCountryActions_Country_SKU_Close_List:
                    print("执行关闭广告循环")

                    
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close),"变更记录"]="执行关闭广告"
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close)&(bulkfile["Status"]=="enabled"),"之前状态"]="之前为开" 
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close),"Status"]="paused"
                    #新增加
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close)&(bulkfile["Ad Group Status"]=="enabled"),"之前状态"]="之前为开" 
                    bulkfile.loc[(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU_Close)&(bulkfile["Ad Group Status"]=="paused"),"之前状态"]="之前为开"

                    
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
            
##################################################################################################################################确保打开###################################################################################################    

            BestSKUCampaign=pd.read_excel(r'D:\\运营\\2生成过程表\\周bulk数据Summary.xlsx',sheet_name="CampaignSKUTotalzhuanhualvMax")
            BestSKUCampaign["zhuanhualv"].fillna(0,inplace=True)
            print(BestSKUCampaign["zhuanhualv"])
         
            #BestSKUCampaignMax=BestSKUCampaign.groupby(["Country","SKU","Ad Group","Campaign Targeting Type"],as_index=False)[["zhuanhualv"]].agg('max')

            AllCountryActions_Country_SKU_Open_List=AllCountryActions.loc[(AllCountryActions["Country"]==AllCountryActions_Country)&(AllCountryActions["行动方案"].str.contains("确保广告是开的")),"SKU"].drop_duplicates().to_list()
           
           
            AllCountryActions_Country_SKU_Open_List_nocomma_list=[]
            for AllCountryActions_Country_SKU_Open_List_nocomma in AllCountryActions_Country_SKU_Open_List:
                print(AllCountryActions_Country_SKU_Open_List_nocomma)
                AllCountryActions_Country_SKU_Open_List_nocomma=AllCountryActions_Country_SKU_Open_List_nocomma
                if ',' in  AllCountryActions_Country_SKU_Open_List_nocomma:
                    print("包含,",AllCountryActions_Country_SKU_Open_List_nocomma)
                    chaifen66=AllCountryActions_Country_SKU_Open_List_nocomma.split(",")[0]
                    
                    AllCountryActions_Country_SKU_Open_List_nocomma_list+=[chaifen66]
            
                
                else:
                    print(AllCountryActions_Country_SKU_Open_List_nocomma)
                    chaifen=AllCountryActions_Country_SKU_Open_List_nocomma
                    AllCountryActions_Country_SKU_Open_List_nocomma_list+=[chaifen]
                        
                    
                    print(chaifen)
                    
                    
            AllCountryActions_Country_SKU_Open_List=AllCountryActions_Country_SKU_Open_List_nocomma_list  #AllCountryActions_Country_SKU_Open_List要开的SKU拆分
            print(AllCountryActions_Country_SKU_Open_List)

            if len(AllCountryActions_Country_SKU_Open_List)>0:
                print("这个国家有要开广告的SKU，继续执行")
                
##################################################################################################3333
                
                for AllCountryActions_Country_SKU in AllCountryActions_Country_SKU_Open_List: #遍历需要开通的SKU
                    pr5_SKU="Pr5"+"-"+str(AllCountryActions_Country_SKU)
                
             #开始遍历SKU
                    print("正在看的SKU",AllCountryActions_Country_SKU)
                       
                    Best52Auto_List=BestSKUCampaign.loc[(BestSKUCampaign["Country"]==AllCountryActions_Country)&(BestSKUCampaign["SKU"]==AllCountryActions_Country_SKU)&(BestSKUCampaign["Campaign Targeting Type"]=="Auto"),"Campaign"].drop_duplicates().to_list()#有历史表现最好的Auto
                    Best52Manual_List=BestSKUCampaign.loc[(BestSKUCampaign["Country"]==AllCountryActions_Country)&(BestSKUCampaign["SKU"]==AllCountryActions_Country_SKU)&(BestSKUCampaign["Campaign Targeting Type"]=="Manual"),"Campaign"].drop_duplicates().to_list()#有历史表现最好的Auto


                    #SKU_AdGroup_List=bulkfile.loc[bulkfile["SKU"]==AllCountryActions_Country_SKU,"Ad Group"].drop_duplicates().to_list()#这里要换成一个SKU的好的Campaign:1。转化率最高的2.
                    
                    if len(Best52Auto_List)==0 and len(Best52Manual_List)==0:
                        
                        print("没有这个产品对应的Campaign")
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


                        
                    elif len(Best52Manual_List)!=0:
                         
                           #手动里的MaxCamaign开启。

                        BestSKUCampaign55=BestSKUCampaign.loc[(BestSKUCampaign["Country"]==AllCountryActions_Country)&(BestSKUCampaign["SKU"]==AllCountryActions_Country_SKU)&(BestSKUCampaign["Campaign Targeting Type"]=="Manual")].reset_index(drop=True)
                        print(BestSKUCampaign55)
                        CampainMax55=BestSKUCampaign55.iloc[[BestSKUCampaign55['zhuanhualv'].idxmax()],[1]].values[0][0]
                        print(CampainMax55)
                    
                        #CampainMax55=BestSKUCampaign.loc[(BestSKUCampaignMax["Country"]==AllCountryActions_Country)&(BestSKUCampaign["SKU"]==AllCountryActions_Country_SKU)&(BestSKUCampaign["Campaign Targeting Type"]=="Manual")&(BestSKUCampaignMax["zhuanhualv"].max()),"Campaign"].values[0]
                        AdGroup55=BestSKUCampaign.loc[(BestSKUCampaign["Country"]==AllCountryActions_Country)&(BestSKUCampaign["SKU"]==AllCountryActions_Country_SKU)&(BestSKUCampaign["Campaign"]==CampainMax55)&(BestSKUCampaign["Campaign Targeting Type"]=="Manual"),"Ad Group"].values[0]      
                        print(AdGroup55)
                       

                           #现有开启的手动如果转化率>0.1则保留。
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Record Type"]=="Campaign"),"变更记录"]="打开Campaign"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Campaign Status"]=="paused"),"之前状态"]="之前状态为paused"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Record Type"]=="Campaign"),"Campaign Status"]="enabled"
                                    
                                    #下面打开Ad Group
                                    
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad Group"),"Campaign Status"]="enabled"
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad Group"),"变更记录"]="打开广告Ad Group"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group Status"]=="paused"),"之前状态"]="之前状态为paused"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad Group"),"Ad Group Status"]="enabled"

                                    #下面打开Ad  
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"变更记录"]="打开广告Ad"
                                    
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Campaign Status"]="enabled"
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Ad Group Status"]="enabled"
                                    
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax55)&(bulkfile["Ad Group"]==AdGroup55)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Status"]="enabled"
                        print("打开一切")                             
                           

                        if len(Best52Auto_List)==0:
                            
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Campaign","Campaign":"Auto"+str(pr5_SKU),"Campaign Daily Budget":3,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Auto","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","变更记录":"生成一个自动"},ignore_index = True)
        
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad Group","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Adgroup_Maxid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","变更记录":"生成一个自动"},ignore_index = True)
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"SKU":AllCountryActions_Country_SKU,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","变更记录":"生成一个自动"},ignore_index = True)
                    
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"close-match","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"loose-match","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"complements","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"substitutes","Product Targeting ID":"close-match","Match Type":"Targeting Expression Predefined","变更记录":"生成一个自动"},ignore_index = True)
               
 

                            
                    elif len(Best52Auto_List)!=0 :
                        #手动里的MaxCamaign开启。


                        BestSKUCampaign77=BestSKUCampaign.loc[(BestSKUCampaign["Country"]==AllCountryActions_Country)&(BestSKUCampaign["SKU"]==AllCountryActions_Country_SKU)&(BestSKUCampaign["Campaign Targeting Type"]=="Auto")].reset_index(drop=True)
           
                        print(BestSKUCampaign77['zhuanhualv'].idxmax())
                        CampainMax77=BestSKUCampaign77.iloc[[BestSKUCampaign77['zhuanhualv'].idxmax()],[1]].values[0][0] 
                        print("CampainMax77",CampainMax77)
                  
                        
                        AdGroup77=BestSKUCampaign.loc[(BestSKUCampaign["Country"]==AllCountryActions_Country)&(BestSKUCampaign["SKU"]==AllCountryActions_Country_SKU)&(BestSKUCampaign["Campaign"]==CampainMax77)&(BestSKUCampaign["Campaign Targeting Type"]=="Auto"),"Ad Group"].values[0]
                        print(AdGroup77)

                                                         
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Record Type"]=="Campaign"),"变更记录"]="打开Campaign"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Record Type"]=="Campaign")&(bulkfile["Campaign Status"]=="paused"),"之前状态"]="之前状态为paused"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Record Type"]=="Campaign"),"Campaign Status"]="enabled"
                                    
                        #下面打开Ad Group
                                    
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad Group"),"Campaign Status"]="enabled"
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad Group"),"变更记录"]="打开广告Ad Group"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad Group")&(bulkfile["Ad Group Status"]=="paused"),"之前状态"]="之前状态为paused"
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad Group"),"Ad Group Status"]="enabled"

                                    #下面打开Ad  
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"变更记录"]="打开广告Ad"
                                    
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Campaign Status"]="enabled"
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Ad Group Status"]="enabled"
                                    
                                    
                        bulkfile.loc[(bulkfile["Campaign"]==CampainMax77)&(bulkfile["Ad Group"]==AdGroup77)&(bulkfile["Record Type"]=="Ad")&(bulkfile["SKU"]==AllCountryActions_Country_SKU),"Status"]="enabled"
                        print("打开一切")     




                        if len(Best52Manual_List)==0:

                        
                              #生成一个Manual
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Campaign","Campaign":"Auto"+str(pr5_SKU),"Campaign Daily Budget":Country_DailyBudget[Bulkfile_Country],"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
        
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad Group","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Adgroup_Maxid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                         
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Ad ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"SKU":AllCountryActions_Country_SKU,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled"},ignore_index = True)
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"手动填写词","Match Type":"exact","变更记录":"补充手动广告"},ignore_index = True)
                            bulkfile_draft_1Country=bulkfile_draft_1Country.append({"Record Type":"Keyword ","Campaign":"Auto"+str(pr5_SKU),"Ad Group":"Auto"+str(pr5_SKU),"Max Bid":Country_Keyword_Maxbid[Bulkfile_Country],"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":"手动填写词","Match Type":"phrase","变更记录":"补充手动广告"},ignore_index = True)




                        
######################################################################################################以下可能不要##########################33                        
 

            
            else:
                print("这个国家"+str(Bulkfile_Country)+"没有要确保开的广告")
                
            bulkfile=pd.concat([bulkfile,bulkfile_draft_1Country],ignore_index=True)
            bulkfile=bulkfile.drop(columns=["Clicks","Spend","Orders","Sales","ACoS"])
            #bulkfile.dropna(subset=["变更记录"],inplace=True)
            #bulkfile.drop(bulkfile[bulkfile["变更记录"]==" "].index,inplace=True)
            bulkfile=bulkfile.drop_duplicates()

            bulkfile.to_excel(r'D:\\运营\\4行动表\\调整SKU广告开关\\'+'_'+str(datetime.date.today())+Bulkfile_Country+"AdjustAdvertisement"+".xlsx",index=False)
              
            break    #A1  
        else:
            listdir=os.listdir(bulkdatafilepath)
            n+=1
            if n==len(listdir):
                print("没有对应国家的Bulkfile文件")
               
