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

#在这之前要生成汇总表，并且把每个国家的bulk表备份到D:\\运营\\bulkoperationfiles\\

Allbulkpath='D:\\运营\\'  
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')


#！！！！筛选汇总表的数据---后续可以按日期>某个日期来筛选
AllbulkD5=Allbulk[(Allbulk['Keyword or Product Targeting'].notna())&(Allbulk['周数']<26)]#定义了26周汇总
print(AllbulkD5)

# AllbulkCampaignKeyword1Week=Allbulk[(Allbulk['Keyword or Product Targeting'].notna())&(Allbulk['周数']==1)].groupby(["Country","Campaign"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg('sum')

#AbulkCampaignKeyword1week["zhuanhualv"]=AllbulkCampaignKeyword1week['Orders']/AllbulkCampaignKeyword1week['Clicks']


AllbulkCampaignKeyword=AllbulkD5.groupby(["Country","Campaign","Keyword or Product Targeting"],as_index=False)[['Impressions','Clicks','Spend','Orders','Total Units','Sales']].agg("sum") 
#AllbulkCampaignKeyword所有周数小于AllbulkD5定义的历史数据汇总

#print(AllbulkCampaignKeyword)
maxrow=len(AllbulkCampaignKeyword)
#maxrow1Week=len(AllbulkCampaignKeyword1Week)
AllbulkCampaignKeyword["zhuanhualv"]=AllbulkCampaignKeyword['Orders']/AllbulkCampaignKeyword['Clicks']
print(maxrow)
print(AllbulkCampaignKeyword)

AllbulkCampaignKeyword.to_excel(Allbulkpath+'周bulk广告数据汇总表huizong.xlsx') #生成分析用的报表




#从国家-bulkopration遍历
bulkoperationfilespath = 'D:\\运营\\bulkoperationfiles\\' #待处理的bulk文件--下载下来作为上传的基础的
bulkoperationfilesnewpath='D:\\运营\\bulkoperationfilesNEW\\'#修改后的bulk文件目录
countrylist1=["NEW-JP","NEW-IT","GV-MX","NEW-MX","NEW-FR","NEW-DE","NEW-ES"]
bulkfilecountrylist=[]#国家空列表
for bulkoperationfile in os.listdir(bulkoperationfilespath):
    clickatempt=10
    clickenough1=16
    clickenough2=32
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
    bulkrecordpath = 'D:\\运营\\bulk变更记录\\' #所有变更记录在国家的一张表
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


    bulkcondiion_new01='AllbulkCampaignKeyword["Orders"]<(AllbulkCampaignKeyword["Clicks"]/clickenough1).apply(np.floor))'
    bulcountryconditionstr='AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"=="bulkfileCountry")'                            
   
    bulkcountrycondionclk1str='(AllbulkCampaignKeyword["Clicks"]>(clickatempt-1)) &(AllbulkCampaignKeyword["Clicks"]<(clickenough+1))]'
    bulkcountrycondionclk21str='AllbulkCampaignKeyword["Clicks"]>clickenough'
    bulkcountrycondionzhl01str= '(AllbulkCampaignKeyword["zhuanhualv"]==0]'
    bulkcountrycondionzhl02str= '(AllbulkCampaignKeyword["zhuanhualv"]>0.2]'
    bulkcountrycondionzhl032str='(AllbulkCampaignKeyword["zhuanhualv"]>0.5]'
    

    
    #bulkconditiondic={}#
########################################################################################################################################################

    
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
                                                                                                    


#将对bulkfile中符合筛选表的camaign和词进行Status批量进行修改成暂停
                                            
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword),"Status"]="paused"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword),"更新内容"]="paused"

        print("02:将bulkfile中这个campaign-词的Status进行修改")
#将bocondition02df修改记录写入记录表
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign)&(dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword),Columnnew1]="bocondition_new01df"
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign)&(dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword),Columnnew2]="paused"
        print("01:将bulkfile中这个campaign-词的修改记录" )

###周转率大约0.2的保持enabled
    bocondition05df=AllbulkCampaignKeyword[(AllbulkCampaignKeyword["Country"]==bulkfileCountry) &(AllbulkCampaignKeyword["zhuanhualv"]>0.2)]
    for boi in range(0,len(bocondition05df)):
           #定义要找的词和Campaign                                   
        boi_keyword=bocondition05df.iloc[[boi],[2]].values[0][0]
        boi_campaign=bocondition05df.iloc[[boi],[1]].values[0][0]
           # 打开Campaign enabled                                    
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Campaign"),"Campaign Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Campaign"),"更新内容"]="转化率>0.2，开启"
#将bocondition05df修改记录写入记录表
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Campaign"),Columnnew1]="bocondition05df"
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Campaign"),Columnnew2]="转化率>0.2，开启"
        
           # 打开Ad Group enabled
                                   
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group"),"Campaign Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group"),"更新内容"]="转化率>0.2，开启"                                   
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group"),"Ad Group Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad Group"),"更新内容"]="转化率>0.2，开启"

#将bocondition05df修改记录写入记录表
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Ad Group"),Columnnew1]="bocondition05df"
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Ad Group"),Columnnew2]="转化率>0.2，开启"


        
           # 打开Ad enabled                                    
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad"),"Campaign Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad"),"更新内容"]="转化率>0.2，开启"                                    
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad"),"Ad Group Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad"),"更新内容"]="转化率>0.2，开启"                                  
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad"),"Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Record Type']=="Ad"),"更新内容"]="转化率>0.2，开启"
#将bocondition05df修改记录写入记录表
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Ad"),Columnnew1]="bocondition05df"
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Record Type']=="Ad"),Columnnew2]="转化率>0.2，开启"
           #打开keyword enabled                                     
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword),"Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword),"Ad Group Status"]="enabled"
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword),"Campaign Status"]="enabled"
                                             
        dfbulkfile.loc[(dfbulkfile["Campaign"]==boi_campaign) & (dfbulkfile['Keyword or Product Targeting']==boi_keyword),"更新内容"]="转化率>0.2，开启"
#将bocondition05df修改记录写入记录表
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword),Columnnew1]="bocondition05df"
        dfbulkmodifyrecord.loc[(dfbulkmodifyrecord["Campaign"]==boi_campaign) & (dfbulkmodifyrecord['Keyword or Product Targeting']==boi_keyword),Columnnew2]="转化率>0.2，开启"
        
           ##?? 后续价格按推荐值调整                                   
        print("05:转化率》0.2的一律开启")


        
                                                   
#########a将修改过的文件存入bulkoperationfilesnewpath
    dfbulkfile_notnull=dfbulkfile[(dfbulkfile['更新内容'].notna()) & (dfbulkfile['更新内容'] !="")]#只保留#只保留
    
    dfbulkfile_notnull.to_excel(bulkoperationfilesnewpath+"Spl-Re_"+Today+"_"+str(bulkoperationfile),index=False)#生成简化版
    
    dfbulkfile.to_excel(bulkoperationfilesnewpath+"Re_"+Today+"_"+str(bulkoperationfile),index=False)

    
#########a将更改记录文件存入bulkrecordpath    
    dfbulkmodifyrecord.to_excel(bulkrecordpath+bulkfileCountry+"_"+"更改记录"+".xlsx",index=False)
########将周bulk广告数据里的文件移动到历史文件夹    
    #shutil.move(r'D:\\运营\\周bulk广告数据\\'+ str(bulkdatafile), r'D:\运营\HistoricalData\周bulk广告数据')

#########3将bulkoperation里的文件删除
    os.remove(bulkoperationfilespath+ str(bulkoperationfile))
########或者可以移动将bulkoperation里的文件移动到历史文件夹
    #shutil.move(bulkoperationfilespath+ str(bulkdatafile), r'D:\运营\HistoricalData\周bulk广告数据\bulkoperationfiles')
 
