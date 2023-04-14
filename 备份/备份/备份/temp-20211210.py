import pandas as pd
print("以下为处理bulk操作报表的程序")

#定义bulk数据汇总表所在路径
Allbulkpath='D:\\运营\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表.xlsx')
print(Allbulk["Sales"])
input(" ?")
#####Campaign汇总表


AllbulkCampaign=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign"],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg('sum')
AllbulkCampaignWEEK=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg('sum')
AllbulkSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","周数"],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg('sum')
AllbulkCampaignSKUWEEK=Allbulk[Allbulk['Record Type']=="Ad"].groupby(["Country","SKU","Campaign","周数"],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg('sum')

writer=pd.ExcelWriter(Allbulkpath+'周bulk数据Summary.xlsx')
AllbulkCampaign.to_excel(writer,"Campaign汇总")
AllbulkCampaignWEEK.to_excel(writer,"CampaignWEEK汇总")
AllbulkSKUWEEK.to_excel(writer,"SKU-WEEK")
AllbulkCampaignSKUWEEK.to_excel(writer,"SKU-Campaign-WEEK")


# AllbulkCampaign.to_excel(Allbulkpath+'周bulk数据-Campaign汇总.xlsx',index=False)

##### Campaign周汇总表
#AllbulkCampaignWEEK=Allbulk[Allbulk['Record Type']=="Campaign"].groupby(["Country","Campaign","周数"],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg('sum')
#AllbulkCampaignWEEK.to_excel(Allbulkpath+'周bulk数据-Campaign-WEEK-汇总.xlsx',index=False)



print(Allbulk)
# AllbulkCampaignKeyword=Allbulk[Allbulk['Keyword or Product Targeting'].notna()].groupby(["Country","Campaign",'Keyword or Product Targeting'],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg('sum')
AllbulkCampaignKeyword=Allbulk[Allbulk['Keyword or Product Targeting'].notna()].groupby(["Country","Campaign","Keyword or Product Targeting"],as_index=False)["Impressions","Clicks","Spend","Orders","Total Units","Sales"].agg("sum")
print(AllbulkCampaignKeyword)
maxrow=len(AllbulkCampaignKeyword)
AllbulkCampaignKeyword["zhuanhualv"]=AllbulkCampaignKeyword['Orders']/AllbulkCampaignKeyword['Clicks']
print(maxrow)
print(AllbulkCampaignKeyword)
clickstest=AllbulkCampaignKeyword.iloc[[0],[5]].values[0][0]
print(clickstest)
AllbulkCampaignKeyword.to_excel(Allbulkpath+'周bulk广告数据汇总表huizong.xlsx')
 
                                
bulkoperation=pd.read_excel(Allbulkpath+'bulkoperation模板.xlsx')
bulkoperation["是否更新"]=""
bulkoperationmaxrow=len(bulkoperation)

for i in range(0,maxrow):
    clicks=AllbulkCampaignKeyword.iloc[[i],[4]].values[0][0]
    zhuanhualv=AllbulkCampaignKeyword.iloc[[i],[9]].values[0][0]
    keyword=AllbulkCampaignKeyword.iloc[[i],[2]].values[0][0]
    #print(keyword)
    input(" ?")
    #print("clicks",clicks,"zhuanhualv",zhuanhualv)
    
    condition01=clicks>16
    condition02=clicks>9 and clicks<17
    condition11=zhuanhualv<0.0625 and zhuanhualv>0 
    condition13=zhuanhualv>0.0625 and zhuanhualv<0.125
    condition12=zhuanhualv==0
    condition61=zhuanhualv>0.5
    condition63=zhuanhualv>0.2
    print(condition63)
    

    if condition02==True and condition12==True:
         for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,10]=bulkoperation.iloc[boj,10]*0.33
                bulkoperation.iloc[boj,31]='点击在10到16次之间,无转化降价0.33'
                print("newprice1",boj)
                
    elif condition01==True and condition12==True:
        #修改操作bulk表
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,17]='paused'
                bulkoperation.iloc[boj,31]='点击超过16次无订单：暂停'
                print("newstatus1")
    elif condition02==True and condition13==True:
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,10]=bulkoperation.iloc[boj,10]*0.6
                bulkoperation.iloc[boj,31]='点击在10到16次之间,转化低降价0.6'
                print("newprice2")
    elif condition01== True and condition11==True:
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,10]=bulkoperation.iloc[boj,10]*0.33
                bulkoperation.iloc[boj,31]='点击大于16次，转化率《0.0625 降价0.33'
                print("newprice3",boj)
    elif condition02==True and condition12==True:
         for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,17]='paused'
                bulkoperation.iloc[boj,31]='此行更新'
                print("newstatus",boj)
    elif condition63==True:
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            samecompaignAdCondition=bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and (bulkoperation.iloc[[boj],[1]].values[0][0]=="Campaign" or bulkoperation.iloc[[boj],[1]].values[0][0]=="Ad" or bulkoperation.iloc[[boj],[1]].values[0][0]=="Ad Group")
            samecompaignAdgroupCondition=bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and (bulkoperation.iloc[[boj],[1]].values[0][0]=="Campaign" or bulkoperation.iloc[[boj],[1]].values[0][0]=="Ad Group")
            samecompaignCondition=bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[1]].values[0][0]=="Campaign" 
            if samekeywordCondition== True:
                print("samekeywordCondition== True")
                bulkoperation.iloc[boj,17]='enabled'
                bulkoperation.iloc[boj,16]='enabled'
                bulkoperation.iloc[boj,15]='enabled'
                bulkoperation.iloc[boj,31]='此行更新'
            elif samecompaignAdCondition== True:
                print("samecompaignAdCondition== True")
                bulkoperation.iloc[boj,17]='enabled'
                bulkoperation.iloc[boj,16]='enabled'
                bulkoperation.iloc[boj,31]='此行更新'
          
            elif samecompaignAdgroupCondition==True:
                print("samecompaignAdgroupCondition==True")
                bulkoperation.iloc[boj,16]='enabled'
                bulkoperation.iloc[boj,15]='enabled'
                bulkoperation.iloc[boj,31]='此行更新'
            elif samecompaignCondition==True:
                print("samecompaign==True")
                bulkoperation.iloc[boj,15]='enabled'
                bulkoperation.iloc[boj,31]='此行更新'
               
                print("good ad confirmed",boj)
               
            
           #else:
               # 写一行此关键词的广告
            
           
