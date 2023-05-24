#################################################################################################做Summa
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')

df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)
df = df[df['日期'].notna()]

df['Spend'] = df['Spend'].astype(float)
df["Max Bid"] = df["Max Bid"].astype(float)
df['Sales'] = df['Sales'].astype(float)
df['日期'] = pd.to_datetime(df['日期'])

def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    df['日期'] = pd.to_datetime(df['日期'])
    df['周数'] = ((last_saturday - df['日期']).dt.days // 7) + 1
    return df

All_adv = update_week_numbers(df)
All_adv=All_adv[All_adv["周数"]<53]
#获取周数=1时"Keyword or Product Targeting" ,Campaign, Ad Group,Match Type, Max Bid, Campaign Status,Ad Group Status,Status
All_adv_basic1=All_adv[All_adv["周数"]==1].loc[:,["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","Max Bid","Campaign Status","Ad Group Status","Status"]]





ALLbulkCampaignSKU=All_adv.loc[All_adv['SKU'].notnull(),['Country','Campaign','SKU']].drop_duplicates()
CamaignSKUAgg=ALLbulkCampaignSKU.groupby(["Country","Campaign"],as_index=False).agg({'SKU':[",".join]})#追加的新的汇总comaignSKU
CamaignSKUAgg.columns =[ 'Country','Campaign','MergedSKU']
print(CamaignSKUAgg)
input("Press Enter to continue...")

All_adv=pd.merge(All_adv,CamaignSKUAgg,how='left',on=["Country","Campaign"])

 
All_adv_group52 = All_adv.groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group52.loc[All_adv_group52["Clicks"]>0,"转化率52"]=All_adv_group52["Orders"]/All_adv_group52["Clicks"]
All_adv_group26 = All_adv[All_adv["周数"]<27].groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group26.loc[All_adv_group26["Clicks"]>0,"转化率26"]=All_adv_group26["Orders"]/All_adv_group26["Clicks"]
All_adv_group8 = All_adv[All_adv["周数"]<9].groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group26.loc[All_adv_group26["Clicks"]>0,"转化率8"]=All_adv_group26["Orders"]/All_adv_group26["Clicks"]
All_adv_group4 = All_adv[All_adv["周数"]<5].groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group4.loc[All_adv_group4["Clicks"]>0,"转化率4"]=All_adv_group4["Orders"]/All_adv_group4["Clicks"]

cols_to_convert1 = ["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales", "ACoS"]

All_adv[cols_to_convert1]=All_adv[cols_to_convert1].replace(',','.',regex=True)
All_adv["Acos"]=All_adv["ACoS"]*1

All_adv_selected=All_adv.loc[All_adv["Keyword or Product Targeting"].notnull(),["Country","Campaign","MergedSKU","Ad Group","Keyword or Product Targeting",	"Match Type","周数","Impressions","Clicks","Spend","Orders","Total Units","Sales","ACoS"]]
All_adv_selected.loc[All_adv_selected["Clicks"]>0,"转化率"]=All_adv_selected["Orders"]/All_adv_selected["Clicks"]
All_adv_selected_uniques=All_adv_selected[["Country","Campaign","Ad Group",	"Keyword or Product Targeting",	"Match Type"]].drop_duplicates()
All_adv_selected_uniques=All_adv_selected_uniques.merge(All_adv_basic1,how='left',on=["Country","Campaign","Ad Group",	"Keyword or Product Targeting",	"Match Type"])


# 先设定你的目标列
cols_to_convert = ["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales", "ACoS","转化率"]

for col in cols_to_convert:
    All_adv_selected[col] = pd.to_numeric(All_adv_selected[col], errors='coerce')



# 假设你的周数范围是从1到52
for week in range(1, 53):
    # 先为每一周的数据创建一个临时的数据框
    temp_df = All_adv_selected[All_adv_selected['周数'] == week]
    
    # 更新列名以反映当前的周数
    temp_df = temp_df.rename(columns={col: col + str(week) for col in cols_to_convert})
    
    # 按照指定的列来合并数据
    merge_cols = ["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"]
    All_adv_selected_uniques = pd.merge(All_adv_selected_uniques, temp_df, how='left', on=merge_cols)

# 打印最终的结果
All_adv= All_adv_selected_uniques
All_adv=pd.merge(All_adv,All_adv_group52,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])
All_adv=pd.merge(All_adv,All_adv_group26,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"],suffixes=('', '_26'))
All_adv=pd.merge(All_adv,All_adv_group8,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"],suffixes=('', '_8'))
All_adv=pd.merge(All_adv,All_adv_group4,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"],suffixes=('', '_4'))


Allbulkpath='D:\\运营\\2生成过程表\\'
#获取周六日期并转换为2023-5-13格式


last_saturday=find_last_saturday().strftime('%Y-%m-%d')
 

writer=pd.ExcelWriter(Allbulkpath+'All_ad'+last_saturday+'.xlsx')
 
All_adv.to_excel(writer,"All_adv",index=False) 
writer.close()
