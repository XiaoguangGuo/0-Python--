# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

samplecolumn=['(Parent) ASIN', '(Child) ASIN', 'Title', 'SKU', 'Sessions', 'Session Percentage', 'Page Views', 'Page Views Percentage', 'Buy Box Percentage', 'Units Ordered', 'Unit Session Percentage', 'Ordered Product Sales', 'Total Order Items', 'Country', '日期']
#取得墨西哥和美国的标准列名
USMXColumnlist=samplecolumn 
key =['US','CA','MX','UK','IT','DE','JP','ES','FR']
#开始导入销售数据
#思路：1.遍历销售数据2. 导入file1加上国家，if 如果是美国，墨西哥（或者如果列数为固定的）则插入2个字段 ;将file1 的dataframe追加到 目标文件dataframe ;导出目标dataframe到excel
#思路：1.遍历销售数据2. 导入file1加上国家，if 如果是美国，墨西哥（或者如果列数为固定的）则插入2个字段 ;将file1 的dataframe追加到 目标文件中间件dataframe ;导出目中间件dataframe到目标dataframe到excel

salesfilepath = 'D:\\运营\\计划数据\\NewCountries\\销售数据\\'
Newallsales=pd.read_excel('D:/SailingstarFBA计划/NEW-ALL周销售数据.xlsx')
NewColumnlist= Newallsales.columns.tolist()
#取得目标文件的列名
               
a=0
for salesfile in os.listdir(salesfilepath):
    print(salesfile)
#遍历数据文件
    sourcedata=pd.read_csv(salesfilepath+str(salesfile)).assign(Country=os.path.basename(salesfile).split('_')[0], 日期=os.path.basename(salesfile).split('_')[1])
    sourcedata['日期'] = pd.to_datetime(sourcedata['日期'])
    print(sourcedata[['Country','日期']])
    for i in range(len(key)):
    #遍历国家名     
        if key[i] in salesfile:
                       
            if len(sourcedata.columns)==15:
            #如果列数=14
                #USMXColumnlist =["(Parent) ASIN","(Child) ASIN","Title","SKU","Sessions","Session Percentage","Page Views","Page Views Percentage","Buy Box Percentage","Units Ordered","Unit Session Percentage","Ordered Product Sales","Total Order Items","Country","日期"]
                #定义标准列名，给美国和MX赋值列名，reindex列名

                print(USMXColumnlist)
                 

                sourcedata.columns=USMXColumnlist
             
                #给数据源文件列名赋标准值
                
                print(NewColumnlist)
                sourcedata=sourcedata.reindex(columns=NewColumnlist)
                #给源文件插入全部缺失的列名，行值为0
                print(len(sourcedata.columns))
                #test.to_excel('D:/运营/计划数据/NewCountries/销售数据/temp.xlsx')
            elif len(sourcedata.columns)==19:
                sourcedata['周数']=""
                print(len(sourcedata.columns),len(NewColumnlist))
                sourcedata.columns=NewColumnlist
                 ##给数据源文件列名赋目标文件列名标准值
            else:
                  print("检查销售数据文件的列名")


            Newallsales=Newallsales.append(sourcedata,ignore_index=True)
            print(Newallsales)
             #追加数据到目标文件
            print("添加销售数据"+str(key[i]))
            break
    #shutil.move(r'D:\\运营\\计划数据\\Newcountries\\销售数据\\'+ str(salesfile), 'D:/运营/HistoricalData/计划数据/Newcountries/销售数据')
    a+=1

#Newallsales["日期"] = Newallsales["日期"].dt.strftime("%Y-%m-%d")
#统一Newallsales的日期格式但是报错
maxtime=pd.to_datetime(Newallsales["日期"].max())

print("最晚时间",maxtime)
Newallsales["日期"]=pd.to_datetime(Newallsales["日期"])
#获取最晚时间
Newallsales['周数']=(maxtime-Newallsales['日期']).dt.days//7+1
#给周数赋值
Newallsales.to_excel(r'D:\SailingstarFBA计划\NEW-ALL周销售数据.xlsx',sheet_name="Sheet1",startrow=0,header=True,index=False)

print("导出目标文件到Excel，数量为"+str(a))
    
