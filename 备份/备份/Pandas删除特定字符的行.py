
#qudiao 除第一行
#通过.isin()，选取包含指定字符串"boy"的行

#wertyyu=wertyyu[~wertyyu.icol(0).isin(["asin"])]  
#通过~取反，选取不包含指定字符串"boy"的行
wertyyu100=wertyyu[~wertyyu.iloc[: , 3].isin(["SKU"])]  
# 学习：删除指定列"sImagePath"=="wj"或者"sImagePath"=="/"的行数据
#df_checkimage = df_checkimage[~df_checkimage["sImagePath"].isin(["/","wj"])]

#去掉标题行，失败，实质是去掉了第一行数据
# wertyyu100.drop([0,0],inplace=True)    

wertyyu100.to_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx',index=False)
