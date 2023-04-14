import pandas as pd
#定义bulk数据汇总表所在路径
Allbulkpath='D:\\运营\\'
Allbulk=pd.read_excel(Allbulkpath+'周bulk广告数据汇总表2.xlsx')
print(Allbulk)
AllbulkCampaignKeyword=Allbulk.groupby(["Country","Campaign",'Keyword or Product Targeting'],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg('sum')

print(AllbulkCampaignKeyword)
maxrow=len(AllbulkCampaignKeyword)
AllbulkCampaignKeyword["zhuanhualv"]=AllbulkCampaignKeyword['Orders']/AllbulkCampaignKeyword['Clicks']
print(maxrow)
print(AllbulkCampaignKeyword)
clickstest=AllbulkCampaignKeyword.iloc[[0],[5]].values[0][0]
print(clickstest)
AllbulkCampaignKeyword.to_excel(Allbulkpath+'周bulk广告数据汇总表huizong2.xlsx')

bulkoperation=pd.read_excel(Allbulkpath+'bulkoperation2.xlsx')
bulkoperationmaxrow=len(bulkoperation)
for i in range(0,maxrow):
    clicks=AllbulkCampaignKeyword.iloc[[i],[4]].values[0][0]
    zhuanhualv=AllbulkCampaignKeyword.iloc[[i],[9]].values[0][0]
    keyword=AllbulkCampaignKeyword.iloc[[i],[2]].values[0][0]
    #print(keyword)

    #print("clicks",clicks,"zhuanhualv",zhuanhualv)
    
    condition01=clicks>16
    condition02=clicks>9 and clicks<17
    condition11=zhuanhualv<0.0625 and zhuanhualv>0
    condition12=zhuanhualv==0
    condition61=zhuanhualv>0.5
    condition63=zhuanhualv>0.2
    print(condition63)
    

    if condition01==True and condition12==True:
        #修改操作bulk表
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,17]='paused'
                print("newstatus")
    elif condition02==True and condition11==True:
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,10]=bulkoperation.iloc[boj,10]*0.33
                print("newprice2")
    elif condition01== True and condition11==True:
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,10]=bulkoperation.iloc[boj,10]*0.33
                print("newprice1",boj)
    elif condition02==True and condition12==True:
         for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            if samekeywordCondition==True:
                bulkoperation.iloc[boj,17]='paused'
                print("newstatus",boj)
    elif condition63==True:
        for boj in range(0,bulkoperationmaxrow):
            samekeywordCondition= bulkoperation.iloc[[boj],[28]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[0]].values[0][0] and bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[11]].values[0][0]==keyword
            samecompaignAdCondition=bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and (bulkoperation.iloc[[boj],[1]].values[0][0]=="Campaign" or bulkoperation.iloc[[boj],[1]].values[0][0]=="Ad" or bulkoperation.iloc[[boj],[1]].values[0][0]=="Ad Group")
            samecompaignAdgroupCondition=bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and (bulkoperation.iloc[[boj],[1]].values[0][0]=="Campaign" or bulkoperation.iloc[[boj],[1]].values[0][0]=="Ad Group")
            samecompaignCondition=bulkoperation.iloc[[boj],[3]].values[0][0]==AllbulkCampaignKeyword.iloc[[i],[1]].values[0][0] and bulkoperation.iloc[[boj],[1]].values[0][0]=="Campaign" 
            if samekeywordCondition== True:
                print("samekeywordCondition== True")
                bulkoperation.iloc[boj,17]='enabled'
                bulkoperation.iloc[boj,16]='enabled'
                bulkoperation.iloc[boj,15]='enabled'
            elif samecompaignAdCondition== True:
                print("samecompaignAdCondition== True")
                bulkoperation.iloc[boj,17]='enabled'
                bulkoperation.iloc[boj,16]='enabled'
          
            elif samecompaignAdgroupCondition==True:
                print("samecompaignAdgroupCondition==True")
                bulkoperation.iloc[boj,16]='enabled'
                bulkoperation.iloc[boj,15]='enabled'
            elif samecompaignCondition==True:
                print("samecompaign==True")
                bulkoperation.iloc[boj,15]='enabled'
               
                print("good ad confirmed",boj)
               
            
           #else:
               # 写一行此关键词的广告
            
           
           
        

#将bulkoperation 导出

bulkoperation.to_excel(Allbulkpath+'bulkoperationnew2.xlsx')









      
#for i in range(0,maxrow+1):
   # t=AllbulkCampaignKeyword.iloc[[i],[5]].values[0][0]/AllbulkCampaignKeyword.iloc[[i],[4]].values[0][0]
    #print(t)

#Allbulk_groupbycampaign-keyword=
#bulknow=
#wordi=pd(a,i)
#zhuanhualvi=pd(a,4)/pd(a,5)
#for i=(0,maxrow)
    #if clicksi>7 and zhuanhualvi<0.05
        #for bulknow 遍历列行和列
            #if bulknow(a,i)=wordi and compaign 相等 国家 相等
                #bulknowprice=
                #bulknowprice=1/3
    #elif zhuanhualv>0.2

        #for  bulknow 遍历列行和列
            #if
            #bulknowprice=1.3*bulknowprice
            
       
