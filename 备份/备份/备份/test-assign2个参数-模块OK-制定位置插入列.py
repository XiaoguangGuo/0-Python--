import pandas as pd
import os
salesfilepath =r'D:\运营\计划数据\NewCountries\销售数据'
test=pd.read_csv(salesfilepath+ "\BusinessReport_3-10-20-2-22_.csv").assign(Country=os.path.basename("BusinessReport_3-10-20-2-22_.csv").split('_')[0],日期=os.path.basename("BusinessReport_3-10-20-2-22_.csv").split('_')[1])
Newallsales=pd.read_excel('D:/SailingstarFBA计划/NEW-ALL周销售数据.xlsx') 
print(test)
#在第K列插入一列“Units Ordered – B2B”
#在原来的L，就是新的M列插入一列“Unit Session Percentage – B2B”





# df.iloc[：列参数]

# 取得列名，并插入" "

#取得标准列名，给美国和加拿大赋值列名，reindex列名

USMXColumnlist=["(Parent) ASIN","(Child) ASIN",	"Title","SKU","Sessions","Session Percentage","Page Views","Page Views Percentage","Buy Box Percentage",
              "Units Ordered","Unit Session Percentage","Ordered Product Sales","Total Order Items","Country","日期"]


test.columns=USMXColumnlist
if len(test.columns)==13:
    print("len(test.columns)==13")
else:
    print("len(test.columns)!=13")
NewUSMXColumnlist= Newallsales.columns.tolist()
#["(Parent) ASIN", "(Child) ASIN",	"Title","SKU","Sessions","Session Percentage","Page Views","Page Views Percentage","Buy Box Percentage",
              #"Units Ordered",Units Ordered - B2B "Unit Session Percentage","Ordered Product Sales","Total Order Items")

print(NewUSMXColumnlist)
test=test.reindex(columns=NewUSMXColumnlist,fill_value=100)

print(len(test.columns))
test.to_excel('D:/运营/计划数据/NewCountries/销售数据/temp.xlsx')
