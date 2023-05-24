#读取"D:\运营\2生成过程表\All_ad2023-05-13.xlsx"到pandas
import pandas as pd
#读取"D:\运营\2生成过程表\All_ad2023-05-13.xlsx"到pandas
All_ad= pd.read_excel(r'D:\运营\2生成过程表\All_ad2023-05-13.xlsx')

#All_ad.loc[(All_ad["Clicks_26"]>10) & (All_ad["转化率26_26"]<0.05) & (All_ad["Clicks1"]>1) & (All_ad["转化率1"]<0.15) & (All_ad["Clicks_4"]>10) & (All_ad["转化率4-4"]<0.1),"标签"]="怪兽"
Monsteradv=All_ad[((All_ad["Clicks_26"]>10) & (All_ad["转化率26_26"]<0.05))& ((All_ad["Clicks1"]>1) & (All_ad["转化率1"]<0.15))& ((All_ad["Clicks_4"]>10) & (All_ad["转化率4_4"]<0.1))]
print(Monsteradv)
#输出Monsteradv
Monsteradv.to_excel(r'D:\运营\2生成过程表\Monsteradv.xlsx',index=False)

