# -*- coding: utf-8 -*-
import os
import xlrd
import xlsxwriter
import pandas as pd
import openpyxl
import xlwings
import shutil


filePath = r'D:\运营\1数据源\周SearchTerm数据'
#需合并的文件所在的文件夹路径

f_name = os.listdir(filePath)
#读取文件夹内所有文件名
#print(f_name)

source_xls= []
for i in f_name:
    source_xls.append(filePath + '\\' + i)
#将文件路径存储在列表中
#print(source_xls)

target_xls = r'D:\运营\1数据源\Adv.xlsx'
#合并后文件的路径


# 读取数据


data = []

for i in source_xls:

#用pandas打开
   # wb=pd.read_excel(i,engine='openpyxl')；
   #如果不能打开,除了xlrd的版本问题，还有可能是文件夹中有隐藏文件，要删除文件夹重新建立。
    wb = xlrd.open_workbook(i)
    for sheet in wb.sheets():
        for rownum in range(0,sheet.nrows): #从第yi行合并含标题
            print(rownum)
        #合并excel中的所有数据
        
            a = sheet.row_values(rownum)
            a.append(i.replace((filePath + '\\'), ''))
            #将文件名做为新的一列
            data.append(a)
            #data.append(sheet.row_values(rownum))
           
#print(data)
# 写入数据

workbook = xlsxwriter.Workbook(target_xls)
worksheet = workbook.add_worksheet()

font = workbook.add_format({"font_size":14})
for i in range(len(data)):
    for j in range(len(data[i])):
        worksheet.write(i, j, data[i][j], font)
        
worksheet.write('X1', '国家')
 
workbook.close()        

#分列
wertyyu = pd.read_excel(r'D:\运营\1数据源\Adv.xlsx')
 
#注意名称一定要有_,能分列；如果列为空就会报错。    
df=wertyyu["国家"].str.split("_",expand=True) 

wertyyu["国家"]=df[0]


print(wertyyu.head(5))
print(df.head(5))


#qudiao 除第一行
#通过.isin()，选取包含指定字符串"boy"的行

#wertyyu=wertyyu[~wertyyu.icol(0).isin(["asin"])]  
#通过~取反，选取不包含指定字符串"boy"的行
#wertyyu100=wertyyu[~wertyyu.iloc[: , 3].isin(["SKU"])]  
# 学习：删除指定列"sImagePath"=="wj"或者"sImagePath"=="/"的行数据
#df_checkimage = df_checkimage[~df_checkimage["sImagePath"].isin(["/","wj"])]

#去掉标题行，失败，实质是去掉了第一行数据
# wertyyu100.drop([0,0],inplace=True)

wertyyu.to_excel(r'D:\运营\1数据源\Adv.xlsx',index=False)


#复制一张表到另一张表

#复制一张表到另一张表
import sys

sys.path.append(r'D:\运营\0-Python程序')

#2、读取待复制的表格
xlpath = r'D:\运营\1数据源\Adv.xlsx'

xlworkbook = xlwings.Book(xlpath)
print(xlworkbook)

#3、读取待粘贴的表格
xlpath2 = r'D:\运营\2生成过程表\Sponsored Products Search term report.xlsx'
xlworkbook2 = xlwings.Book(xlpath2)

#3-1、找到最后一行的第一个单元格
rng = xlworkbook2.sheets("Sponsored Product Search Term R").range('A1').expand('table')

cell_index = str(rng.rows.count+1)

range1 = xlworkbook2.sheets("Sponsored Product Search Term R").range('A'+cell_index)

#3-2、按行复制数据到目标表格。
range1.value = xlworkbook.sheets("Sheet1").range('A1').expand('table').value
#不要第一行


xlworkbook.save()
xlworkbook.close()
xlworkbook2.save()
xlworkbook2.close()




shanchuduoyuhang=pd.read_excel(r'D:\\运营\\2生成过程表\\Sponsored Products Search term report.xlsx')
print(shanchuduoyuhang)
#qudiao 除第一行

shanchuduoyuhang=shanchuduoyuhang[~shanchuduoyuhang.iloc[:,0].isin(["日期"])]
shanchuduoyuhang=shanchuduoyuhang[~shanchuduoyuhang.iloc[:,0].isin(["Date"])]  


#通过~取反，选取不包含指定字符串"日期"的行：
#wertyyu100=wertyyu[~wertyyu.iloc[: , 3].isin(["SKU"])]  

#去掉标题行，失败，实质是去掉了第一行数据
# 此句测试shanchuduoyuhang.drop([0,0],inplace=True)
#  cuowu shanchuduoyuhang['Date'] = pd.to_datetime(shanchuduoyuhang['Date'])
#shanchuduoyuhang["Date"] = shanchuduoyuhang["Date"].dt.strftime("%Y-%m-%d")
#shanchuduoyuhang['Date'] = pd.to_datetime(shanchuduoyuhang.Date)
#shanchuduoyuhang.Date = pd.to_datetime(shanchuduoyuhang.Date)

shanchuduoyuhang["COUNTRY"].replace("CA","GV-CA",inplace=True)
shanchuduoyuhang["COUNTRY"].replace("US","GV-US",inplace=True)
shanchuduoyuhang["COUNTRY"].replace("MX","GV-MX",inplace=True)

shanchuduoyuhang.to_excel(r'D:\\运营\\2生成过程表\Sponsored Products Search term report.xlsx',sheet_name ="Sponsored Product Search Term R", index=False)

#####################################################以上是为了生成汇总表Sponsored Products Search term report#################################################################################################################

#################################################################################################################################################################################################################################
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


SearchTermAll_Good=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]>=0.2)&(SearchTermAll_Sum["Clicks"]>=5)] #条件改为


SearchTermAll_Bad=SearchTermAll_Sum[(SearchTermAll_Sum["转化率"]<0.05)&(SearchTermAll_Sum["Clicks"]>=20)]


ProductActions=pd.read_excel(r'D:\\运营\\3数据分析结果\\国家汇总.xlsx',sheet_name="ProductActions")



#打开Bulkfile的第一种方法

Allbulkpath='D:\\运营\\2生成过程表\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')



#遍历SeachTermGood


SearchTermAll_Good_Country_list=SearchTermAll_Good["COUNTRY"].drop_duplicates().to_list()

Allbulk_Campaign_SKUMax=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('sum')#全部历史bulk的sku spend相加：找自动广告的sku

################################################################GOOD#########################################################################################################################

for countryname in SearchTermAll_Good_Country_list:   #遍历searchTemgood里的国家


    CountrySKU_Close_List=ProductActions.loc[(ProductActions["COUNTRY"]==countryname)&(ProductActions["行动方案"].str.contains("关闭广告")),"SKU"].drop_duplicates().to_list()

    AllCountryActions_Country_SKU_Close_List_nocomma_list=[]
    for AllCountryActions_Country_SKU_Close_List_nocomma in CountrySKU_Close_List:
        print(AllCountryActions_Country_SKU_Close_List_nocomma)
        AllCountryActions_Country_SKU_Close_List_nocomma=AllCountryActions_Country_SKU_Close_List_nocomma
        if ',' in  AllCountryActions_Country_SKU_Close_List_nocomma:
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
                    
                    
    CountrySKU_Close_List=AllCountryActions_Country_SKU_Close_List_nocomma_list
    print(AllCountryActions_Country_SKU_Close_List)
                
                                                            
    
     
    
    n=0                                             
    for bulkdatafile in os.listdir(bulkdatafilepath): #找bulkfile对应的国家文件

        Bulkfile_Country=bulkdatafile.split('_')[0]

        if Bulkfile_Country==countryname:
            SearchTermAll_Good_Country=SearchTermAll_Good[SearchTermAll_Good["COUNTRY"]==str(countryname)]
         
            
            print("SearchTermAll_Good_Country")
            bulkfile=pd.read_excel(bulkdatafilepath+str(bulkdatafile),engine="openpyxl",sheet_name=1)
            n+=1
            break
    if n==1:
        print(bulkdatafile,bulkfile)
        
        Bulkfile_SearchTerm_Add=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_auto=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])

        Bulkfile_SearchTerm_Add_auto_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_manual=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_draft99=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])
        Bulkfile_SearchTerm_Add_manual_draft=pd.DataFrame(columns=["Record ID","Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Bidding strategy","Placement Type","Increase bids by placement","更改记录"])

   

        for Stoi in range(0,len(SearchTermAll_Good_Country)):#遍历好词ddataframe的行#for A0
 
            Stoi_keyword=SearchTermAll_Good_Country.iloc[[Stoi],[3]].values[0][0]
            Stoi_keyword=str(Stoi_keyword)
            print("Stoi_keyword=",Stoi_keyword)
            Stoi_Campaign=SearchTermAll_Good_Country.iloc[[Stoi],[1]].values[0][0]#获取第n行的关键词字段的值，遍历
            print(Stoi_Campaign)
            Stoi_AdGroup=SearchTermAll_Good_Country.iloc[[Stoi],[2]].values[0][0]
        

                                                  
            averageprice=SearchTermAll_Good_Country.iloc[[Stoi],[6]].values[0][0]/SearchTermAll_Good_Country.iloc[[Stoi],[5]].values[0][0]
            print("averageprice",averageprice)

    


            type1=bulkfile.loc[(bulkfile['Record Type']=="Campaign")&(bulkfile['Campaign']==Stoi_Campaign),"Campaign Targeting Type"]
            length_type=len(bulkfile.loc[(bulkfile['Record Type']=="Campaign")&(bulkfile['Campaign']==Stoi_Campaign),"Campaign Targeting Type"])
            print(type1,length_type)
            if Stoi_keyword.isdigit():
                print("数字")
                continue
            if Stoi_keyword.startswith('b0'):
                print("暂不处理Asin")
                continue
            if length_type==0:
               print("这个好词对应的Camaign在bulkfile没有记录，生成新的Campaign")
               #生成一个新的campaign
  
               bulkfile_Campaign_SKUMax66=Allbulk_Campaign_SKUMax.groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('max')
               bulkfile_Campaign_SKU_list66=bulkfile_Campaign_SKUMax66.loc[(bulkfile_Campaign_SKUMax66["Country"]==countryname)&(bulkfile_Campaign_SKUMax66["Campaign"]==Stoi_Campaign)&(bulkfile_Campaign_SKUMax66['Ad Group']==Stoi_AdGroup),"SKU"].to_list()
               if len(bulkfile_Campaign_SKU_list66)==0:
                   print("实在找不到Sku，循环下一个词")
                   continue
               bulkfile_Campaign_SKUMax66=bulkfile_Campaign_SKU_list66[0]
               #bulkfile_Campaign_SKUMax66=Allbulk_Campaign_SKUMax[(Allbulk_Campaign_SKUMax["Campaign"]==Stoi_Campaign) &(Allbulk_Campaign_SKUMax['Ad Group']==Stoi_AdGroup)&(Allbulk_Campaign_SKUMax['Spend'].max()),"SKU"].values[0][0])  #只要落在自动的SKU上
               print("找好词对应的SKU：bulkfile_Campaign_SKUMax2",bulkfile_Campaign_SKUMax2)

               pr66_sku="Pr4a"+"-"+str(bulkfile_Campaign_SKUMax2)
               if bulkfile_Campaign_SKUMax2 not in CountrySKU_Close_List:
                   
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Campaign","Campaign":pr66_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad Group","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad ","Campaign":pr66_sku,"Ad Group":pr66_sku,"SKU":bulkfile_Campaign_SKUMax66,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Keyword","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)

               if bulkfile_Campaign_SKUMax2  in CountrySKU_Close_List:    
                   
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Campaign","Campaign":pr66_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)","更改记录":"关闭广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad Group","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled","更改记录":"关闭广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Ad ","Campaign":pr66_sku,"Ad Group":pr66_sku,"SKU":bulkfile_Campaign_SKUMax66,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","更改记录":"关闭广告"},ignore_index = True)
                   Bulkfile_SearchTerm_Add_draft99=Bulkfile_SearchTerm_Add_draft99.append({"Record Type":"Keyword","Campaign":pr66_sku,"Ad Group":pr66_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"关闭广告"},ignore_index = True)
             
                  


               
            else:
                Campaign_type_list=bulkfile.loc[(bulkfile['Record Type']=="Campaign")&(bulkfile['Campaign']==Stoi_Campaign),"Campaign Targeting Type"].to_list()
                print(Campaign_type_list)
       
#####################################如果好词从Auto广告产生########################Auto#################################################################################################                
                if Campaign_type_list[0]=="Auto" :#Camaign是否为自动 #if:AA1
                    
                    print("这个好词来自自动广告",Stoi_Campaign)
                                                                                              
                
                #找到历史花费最多的SKU
                    print("好词对应的广告从自动广告产生")
                    bulkfile_Campaign_SKUMax1=Allbulk_Campaign_SKUMax.groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('max')                                                                          
                    #下面寻找这个广告最大花费的Ad:SKU
                    bulkfile_Campaign_SKU_list77=bulkfile_Campaign_SKUMax1.loc[(bulkfile_Campaign_SKUMax1["Country"]==countryname)&(bulkfile_Campaign_SKUMax1["Campaign"]==Stoi_Campaign)&(bulkfile_Campaign_SKUMax1['Ad Group']==Stoi_AdGroup),"SKU"].to_list()
                    print("SKUlist",bulkfile_Campaign_SKU_list77)
                    
                    if len(bulkfile_Campaign_SKU_list77)==0:
                        print("bulkfile_Campaign_SKU_list长度为0")
                        bulkfile_Campaign_SKUMax77=Allbulk_Campaign_SKUMax.groupby(["Country","Campaign","Ad Group","SKU"],as_index=False)[['Spend']].agg('max')
                        bulkfile_Campaign_SKU_list77=bulkfile_Campaign_SKUMax77.loc[(bulkfile_Campaign_SKUMax77["Country"]==countryname)&(bulkfile_Campaign_SKUMax77["Campaign"]==Stoi_Campaign)&(bulkfile_Campaign_SKUMax77['Ad Group']==Stoi_AdGroup),"SKU"].to_list()
                        if len(bulkfile_Campaign_SKU_list77)==0:
                            print("实在找不到Sku，循环下一个词")
                            continue
                    else:      
                        bulkfile_Campaign_SKUMax2=bulkfile_Campaign_SKU_list77[0]
                        
                    print("选的SKU",bulkfile_Campaign_SKUMax2)
                    #bulkfile_Campaign_SKUMax2=Allbulk_Campaign_SKUMax[(Allbulk_Campaign_SKUMax["Campaign"]==Stoi_Campaign) &(Allbulk_Campaign_SKUMax['Ad Group']==Stoi_AdGroup)&(Allbulk_Campaign_SKUMax['Spend'].max()),"SKU"].values(0,0)  #只要落在自动的SKU上                                                                 
                    print("好词对应的SKU：bulkfile_Campaign_SKUMax2",bulkfile_Campaign_SKUMax2)
                    
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
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Campaign","Campaign":pr4_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)
                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase"},ignore_index = True)        


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
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                                        bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Record Type"]=="Ad")&bulkfile["SKU"]==bulkfile_Campaign_SKUMax2,"Status"]="enabled" #打开SKU的状态
                                        
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
                                        Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":bulkfileSKU_Manual_Campaign,"Ad Group":bulkfileSKUManualCampaignAdGroup_Assign,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)
                        
                                    elif bulkfile_Campaign_SKUMax2  in CountrySKU_Close_List:
                                       
                                        Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":bulkfileSKU_Manual_Campaign,"Ad Group":bulkfileSKUManualCampaignAdGroup_Assign,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"paused","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact","更改记录":"关闭广告可上传"},ignore_index = True)




                
                
 



                    
                   
                                if Stoi_keyword in (bulkfile.loc[(bulkfile["Campaign"]==bulkfileSKU_Manual_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Keyword or Product Targeting"].drop_duplicates().to_list()):
###########如果存在这个SKU的Pr4广告有这个词的phrase
                                    if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                                        
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
                                          Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":bulkfileSKU_Manual_Campaign,"Ad Group":bulkfileSKUManualCampaignAdGroup_Assign,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase"},ignore_index = True)
                      
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
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)
                                    else:                                    
                                        if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                                    
                                            Maxbid_Adgroup=0.99

                                    
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Campaign","Campaign":pr4_sku,"Campaign Daily Budget":5,"Campaign Start Date": datetime.date.today(),"Campaign Targeting Type":"Manual","Campaign Status":"enabled","Bidding strategy":"Dynamic bidding (up and down)"},ignore_index = True)
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad Group","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":Maxbid_Adgroup,"Campaign Status":"enabled","Ad Group Status":"enabled"},ignore_index = True)
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Ad ","Campaign":pr4_sku,"Ad Group":pr4_sku,"SKU":bulkfile_Campaign_SKUMax2,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled"},ignore_index = True)
                                            Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":pr4_sku,"Ad Group":pr4_sku,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)
                                    
 
#####################################Auto###################################1end


########################################Manual####################如果来自手动广告:则判断在这个Stoi_Campaign中  ############################################################################################################################################################################2
                                                                                              
           
          
                else: #如果来自手动:则就追加到在这个Stoi_Campaign中 #if:AA1
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
                        if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                        
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Campaign Status"]="enabled" #打开这个词，防止关闭
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Ad Group Status"]="enabled"#打开这个词，防止关闭
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact"),"Status"]="enabled"#打开这个词，防止关闭



                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"Campaign Status"]="enabled"  #打开这个广告的所有Campaign Status? 是否只打开Campaign？由于有优先级的问题，可以一统之。是否打开关闭广告由产品计划决定！！
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign),"Ad Group Status"]="enabled"  #打开Ad goup
                        #bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad"),"Status"]="enabled" #打开所有SKU的状态


                

                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Campaign")]) #追加Campaign              
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad Group")])#追加Ad Group
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Record Type"]=="Ad")])#追加Ad
                            Bulkfile_SearchTerm_Add_manual=Bulkfile_SearchTerm_Add_manual.append(bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="exact")])#追加Keyword

                
                    else:#if#MN-A
 #如果没有exac####生成在产生这个词的Campaign中。                                               
                 #生成一个
                        print("这个手动广告不包含这个好词的exact")
                        if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                            Bulkfile_SearchTerm_Add_manual_draft=Bulkfile_SearchTerm_Add_manual_draft.append({"Record Type":"Keyword","Campaign":Stoi_Campaign,"Ad Group":Stoi_AdGroup,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"exact"},ignore_index = True)



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
                        if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Campaign Status"]="enabled" #打开campaign status
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Ad Group Status"]="enabled" #打开Ad group status
                            bulkfile.loc[(bulkfile["Campaign"]==Stoi_Campaign)&(bulkfile["Keyword or Product Targeting"]==Stoi_keyword)&(bulkfile["Match Type"]=="phrase"),"Status"]="enabled" #打开Keyword status
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

                        if bulkfile_Campaign_SKUMax2  not in CountrySKU_Close_List:
                            Bulkfile_SearchTerm_Add_manual_draft=Bulkfile_SearchTerm_Add_manual_draft.append({"Record Type":"Keyword","Campaign":Stoi_Campaign,"Ad Group":Stoi_AdGroup,"Max Bid":averageprice,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":Stoi_keyword,"Match Type":"phrase"},ignore_index = True)


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
    
    n=0                                             
    for bulkdatafile in os.listdir(bulkdatafilepath): #找bulkfile对应的国家文件

        Bulkfile_Country=bulkdatafile.split('_')[0]

        if Bulkfile_Country==countryname22:

            bulkfile=pd.read_excel(bulkdatafilepath+str(bulkdatafile),engine="openpyxl",sheet_name=1)
            n+=1
            break
    if n==1:
        print(bulkdatafile,bulkfile)
    bulkfile["更改记录"]

    if len(SearchBadwithCountry)>0:
        for oi78 in range(len(SearchBadwithCountry)):
            
            oi78_keyword=SearchTermAll_Bad.iloc[[Stoi],[3]].values[0][0]
            oi78_keyword=str(oi78_keyword)
            
            oi78_Campaign=SearchTermAll_Bad.iloc[[Stoi],[1]].values[0][0]#获取第n行的关键词字段的值，遍历
            print(Stoi_Campaign)
            oi78_AdGroup=SearchTermAll_Bad.iloc[[Stoi],[2]].values[0][0]
            #下面找到bulkfile中对应的广告进行关闭
             #看是手动还是自动？


            # 如果是手动：
            if len(bulkfile.loc[bulkfile["Campaign"]==oi78_Campaign])>0:
            
                campaign78type=bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Record Type"]=="Campaign"),"Campaign Targeting Type"].values[0]
                if campaign78type=="Manual":

                
                    testlist78=bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact")].to_list()
                    if len(testlist78)>0:
                        bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact"),"更改记录"]="Bad"
                        bulkfile.loc[(bulkfile["Campaign"]==oi78_Campaign)&(bulkfile["Ad Group"]==oi78_AdGroup)&(bulkfile["Keyword or Product Targeting"]==oi78_keyword)&(bulkfile["Match Type"]=="exact"),"Status"]="paused"
            #添加negtive+
                    Bulkfile_SearchTerm_Add_manual_draft=Bulkfile_SearchTerm_Add_manual_draft.append({"Record Type":"Keyword","Campaign":oi78_Campaign,"Ad Group":oi78_AdGroup,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":oi78_keyword,"Match Type":"negative exact","更改记录":"Bad"},ignore_index = True)
             
                elif campaign78type=="Auto":
                
                    Bulkfile_SearchTerm_Add_auto_draft=Bulkfile_SearchTerm_Add_auto_draft.append({"Record Type":"Keyword","Campaign":oi78_Campaign,"Ad Group":oi78_AdGroup,"Campaign Status":"enabled","Ad Group Status":"enabled","Status":"enabled","Keyword or Product Targeting":oi78_keyword,"Match Type":"negative exact","更改记录":"Bad"},ignore_index = True)
                
                
                    Bulkfile_SearchTerm_Add78=pd.concat([Bulkfile_SearchTerm_Add_auto_draft,Bulkfile_SearchTerm_Add_manual_draft],ignore_index=True) #追加到Bulkfile_SearchTerm_Add  

            else:
                print("坏词对应的Campaign没了")
    Bulkfile_SearchTerm_Add78.to_excel(r'D:\\运营\\4行动表\\Bulkfile_SearchTerm\\'+'Bad_'+str(datetime.date.today())+countryname+"Bulkfile_SearchTerm"+".xlsx",index=False)
 
    
    
        
                     
