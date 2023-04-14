import pandas as pd


shanchuduoyuhang=pd.read_excel(r'D:\\运营\\Sponsored Products Search term report.xlsx')
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

shanchuduoyuhang.to_excel(r'D:\运营\Sponsored Products Search term reporte-test123.xlsx',sheet_name="Sheet1", index=False)

