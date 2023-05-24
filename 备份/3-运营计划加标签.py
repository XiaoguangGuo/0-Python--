
# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil
import datetime

plan=pd.read_excel(r'D:\运营\plan.xlsx',sheet_name="Sheet1")
plan["皮质层标签"]=""
plan["行动方案"]=""
plan.loc[(plan["Fufillable"]<1),"皮质层标签"] = plan["皮质层标签"].astype(str)+"无库存"
plan.loc[(plan["Fufillable"]<5) &(plan["Fufillable"]>0),"皮质层标签"] = plan["皮质层标签"].astype(str)+"低库存"

#平均每天销售一个的产品为重点产品
plan.loc[(plan["SELLING10"]>10),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 重点产品"
plan.loc[(plan["Adjusted-Week2"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要快递发货"
plan.loc[(plan["Adjusted-Week4"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要空运发货"
plan.loc[(plan["Adjusted-Week8"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要快海运发货"
plan.loc[(plan["Adjusted-Week15"].astype(float)>0),"皮质层标签" ] = plan["皮质层标签"].astype(str)+"需要慢海运发货"
plan.loc[(plan["GGZZ1"]>0.5),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告增长"
plan.loc[(plan["GGZZ1"]<-0.5),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告降低"
plan.loc[((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0)&((plan["STOCKALL"]/(plan["1"]+plan["2"]+plan["3"]+plan["4"])*4)<4),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 4周后缺货-提价"

plan.loc[(plan["广告1"]<0.5)| (plan["广告1"].isna())|(plan["广告1"]==""),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周无广告"

#每个订单花费广告费超过2美元就是效果差
plan.loc[((plan["Country"]=="US")|(plan["Country"]=="CA"))&( plan["BILI1"]>2 ),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告效果差"
plan.loc[((plan["Country"]=="US")|(plan["Country"]=="CA"))&( plan["BILI1"]<0.3 ),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 广告花费占比小"
#4周广告费比上四周订单超过2美元就是长期效果差

plan.loc[((plan["Country"]=="US")|(plan["Country"]=="CA"))&(((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0) & ((plan["广告1"]+plan["广告2"]+plan["广告3"]+plan["广告4"])/(plan["1"]+plan["2"]+plan["3"]+plan["4"])>2)),"皮质层标签" ] = plan["皮质层标签"].astype(str)+" 长期广告效果差"

#墨西哥广告超过40比索是效果差
plan.loc[((plan["Country"]=="MX")&( plan["BILI1"]>40 )),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告效果差"

plan.loc[(plan["Country"]=="MX")&(((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0) & ((plan["广告1"]+plan["广告2"]+plan["广告3"]+plan["广告4"])/(plan["1"]+plan["2"]+plan["3"]+plan["4"])>40)),"皮质层标签" ] = plan["皮质层标签"].astype(str)+" 长期广告效果差"

plan.loc[(plan["广告1"]>0) &(plan["1"]==0),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周广告效果差"

#All_Product_Analyzefile_LABEL=All_Product_Analyzefile[(All_Product_Analyzefile["销量"]<2)&(All_Product_Analyzefile["FBA可售"]>=50)&(All_Product_Analyzefile["周数"]==1)]
plan.loc[(plan["1"]<1),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周无销量"

plan.loc[(plan["Zhouzhuan10"]>50)| (plan["SELLING10"]==0)&(plan["STOCKALL"]>0),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 长期积压"

plan.loc[(plan["2"]>0)& (plan["1"]/plan["2"]<0.5),"皮质层标签"] = plan["皮质层标签"].astype(str)+" 本周销量暴跌"
############################################################

plan.loc[plan["皮质层标签"].str.contains("提价","本周广告增长"),"行动方案"]=plan["行动方案"].astype(str)+"1.暂停广告"
plan.loc[(((plan["1"]+plan["2"]+plan["3"]+plan["4"])==0)&(plan["Fufillable"]>30))|((plan["1"]+plan["2"]+plan["3"]+plan["4"])>0)&(((plan["Fufillable"]/(plan["1"]+plan["2"]+plan["3"]+plan["4"])*4)>4)),"行动方案"] = plan["行动方案"].astype(str)+"2.确保广告是开的"







############################################################


#All_Product_Analyzefile.to_excel(r'E:\\20220615网盘备份\\运营数据\\All_Product_Analyzefile.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)
plan.to_excel(r'D:\运营\数据分析结果\计划结果\New_plan.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)



