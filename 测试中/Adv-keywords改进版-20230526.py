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
All_adv=All_adv.drop_duplicates()
#获取周数=1时"Keyword or Product Targeting" ,Campaign, Ad Group,Match Type, Max Bid, Campaign Status,Ad Group Status,Status
All_adv_basic1=All_adv[All_adv["周数"]==1].loc[:,["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","Max Bid","Campaign Status","Ad Group Status","Status"]]





ALLbulkCampaignSKU=All_adv.loc[All_adv['SKU'].notnull(),['Country','Campaign','SKU']].drop_duplicates()
CamaignSKUAgg=ALLbulkCampaignSKU.groupby(["Country","Campaign"],as_index=False).agg({'SKU':[",".join]})#追加的新的汇总comaignSKU
CamaignSKUAgg.columns =[ 'Country','Campaign','MergedSKU']
print(CamaignSKUAgg)



#将Bulkfiles的数据加入主要SKU列
def process_spend_summary(BulkFile_df):
    # 检查是否有 "Country" 列
    has_country = "Country" in BulkFile_df.columns

    if not has_country:
        print("没有国家列")

    # 选择 SKU 列不为空的行
    BulkFile_df = BulkFile_df[BulkFile_df['SKU'].notna()]

    # 按 "Country"（如果有），"Campaign"，"Ad Group" 和 "SKU" 对 "Spend" 进行汇总
    group_columns = ["Country", "Campaign", "Ad Group", "SKU"] if has_country else ["Campaign", "Ad Group", "SKU"]
    spend_summary = BulkFile_df.groupby(group_columns).agg({"Spend": "sum"}).reset_index()

    # 为每个 "Country"（如果有），"Campaign" 和 "Ad Group" 找到具有最大 "Spend" 的 SKU
    group_columns = ["Country", "Campaign", "Ad Group"] if has_country else ["Campaign", "Ad Group"]
    spend_summary = spend_summary.loc[spend_summary.groupby(group_columns)["Spend"].idxmax()]

    # 将结果重命名为 "主要SKU"
    spend_summary = spend_summary.rename(columns={"SKU": "主要SKU"})
    print(spend_summary)

    # 将结果合并到周BUlkFIle原始数据表，创建一个新列 "主要SKU"


    return spend_summary 


All_adv=pd.merge(All_adv,CamaignSKUAgg,how='left',on=["Country","Campaign"])
spend_summary=process_spend_summary(All_adv)

merge_columns = ["Country", "Campaign", "Ad Group","主要SKU"]
All_adv = All_adv.merge(spend_summary[merge_columns], on=["Country", "Campaign", "Ad Group"], how="left")

print(All_adv)


 
All_adv_group52 = All_adv.groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group52.loc[All_adv_group52["Clicks"]>0,"转化率Sum52"]=All_adv_group52["Orders"]/All_adv_group52["Clicks"]
All_adv_group26 = All_adv[All_adv["周数"]<27].groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group26.loc[All_adv_group26["Clicks"]>0,"转化率Sum26"]=All_adv_group26["Orders"]/All_adv_group26["Clicks"]
All_adv_group8 = All_adv[All_adv["周数"]<9].groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group26.loc[All_adv_group26["Clicks"]>0,"转化率Sum8"]=All_adv_group26["Orders"]/All_adv_group26["Clicks"]
All_adv_group4 = All_adv[All_adv["周数"]<5].groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group4.loc[All_adv_group4["Clicks"]>0,"转化率Sum4"]=All_adv_group4["Orders"]/All_adv_group4["Clicks"]
All_adv_group2 = All_adv[All_adv["周数"]<3].groupby(["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])[["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales"]].sum().reset_index()
All_adv_group2.loc[All_adv_group4["Clicks"]>0,"转化率Sum2"]=All_adv_group2["Orders"]/All_adv_group2["Clicks"]

cols_to_convert1 = ["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales", "ACoS"]

All_adv[cols_to_convert1]=All_adv[cols_to_convert1].replace(',','.',regex=True)
All_adv["Acos"]=All_adv["ACoS"]*1

All_adv_selected=All_adv.loc[All_adv["Keyword or Product Targeting"].notnull(),["Country","Campaign","MergedSKU","主要SKU","Ad Group","Keyword or Product Targeting",	"Match Type","周数","Impressions","Clicks","Spend","Orders","Total Units","Sales","ACoS"]]
All_adv_selected.loc[All_adv_selected["Clicks"]>0,"转化率"]=All_adv_selected["Orders"]/All_adv_selected["Clicks"]
All_adv_selected_uniques=All_adv_selected[["Country","Campaign","Ad Group",	"Keyword or Product Targeting",	"Match Type","MergedSKU","主要SKU"]].drop_duplicates()
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
    merge_cols = ["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","主要SKU","MergedSKU"]
    All_adv_selected_uniques = pd.merge(All_adv_selected_uniques, temp_df, how='left', on=merge_cols)

# 打印最终的结果
All_adv= All_adv_selected_uniques
All_adv=All_adv.drop_duplicates(subset=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","主要SKU","MergedSKU"],keep='first')



All_adv=pd.merge(All_adv,All_adv_group52,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"])
All_adv=pd.merge(All_adv,All_adv_group26,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"],suffixes=('', '_26'))
All_adv=pd.merge(All_adv,All_adv_group8,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"],suffixes=('', '_8'))
All_adv=pd.merge(All_adv,All_adv_group4,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"],suffixes=('', '_4'))
All_adv=pd.merge(All_adv,All_adv_group2,how='left',on=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type"],suffixes=('', '_2'))


#加好词标签 IF(AND(M2>20, Q2>0.25), "强词", IF(AND(M2>10, Q2>0.15), "好词",IF(AND(M2>20,Q2<0.05),"差词",IF(AND(M2>5,Q2>0.1),"保持词",""))))
All_adv["好词标签"]=""
import re

def classify_words(row, col1, col2):
    num = ''.join(re.findall(r'\d+',col2))  # 使用正则表达式从字符串中提取数字
    if row[col1] > 50 and row[col2] > 0.25:
        return num + "强词"
    elif row[col1] > 10 and row[col2] > 0.15:
        return num + "好词"
    elif row[col1] > 100 and row[col2] < 0.05:
        return num + "高点击差词"
    elif row[col1] > 20 and row[col2] < 0.05:
        return num + "差词"
    elif row[col1] > 5 and row[col2] > 0.1:
        return num + "保持词"
    else:
        return ""

columns = [('Clicks','转化率Sum52'), ('Clicks_26','转化率Sum26'), ('Clicks_8','转化率Sum8'), ('Clicks_4','转化率Sum4'),('Clicks_2','转化率Sum2'),('Clicks1','转化率1')]  # 这是你的列对，你需要将其替换为实际的列名对

All_adv['好词标签'] = All_adv.apply(lambda row: ' '.join([classify_words(row, col1, col2) for col1, col2 in columns]), axis=1)

All_adv["皮质层标签"]=""

def icy_words(row, col1, col2):
    num = ''.join(re.findall(r'\d+',col2))  # 使用正则表达式从字符串中提取数字
    
    
    if row[col1] > 60 and row[col2]==0:
        return num + "高点击零点词" 
    elif row[col1] > 20 and row[col2]==0:
        return num + "零点词"
    elif row[col1] > 60 and row[col2] < 0.03:

        return num + "高点击冰点词"
    elif row[col1] > 20 and row[col2] < 0.03:
        return num + "冰点词"
    
    else:
        return ""
    

All_adv["皮质层标签"]=All_adv.apply(lambda row: ' '.join([icy_words(row, col1, col2) for col1, col2 in columns]), axis=1)


All_adv = All_adv[All_adv['Country'].str.contains('GV')]

All_adv_simple=All_adv[["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","Campaign Status","Ad Group Status","Status","主要SKU","MergedSKU","好词标签","皮质层标签","Impressions1","Clicks1","Spend1","Orders1","Total Units1","Sales1","ACoS1","转化率1", "转化率2","Clicks","转化率Sum52","Clicks_26","转化率Sum26","Clicks_8","转化率Sum8","Clicks_4","转化率Sum4","Clicks_2","转化率Sum2"]]

#将日销售目标合并进来
#读取"D:\运营\2生成过程表\TESTAll_Product_Analyzefile.xlsx"
TESTAll_Product_Analyzefile=pd.read_excel('D:\\运营\\2生成过程表\\TESTAll_Product_Analyzefile.xlsx')
TESTAll_Product_Analyzefile_adMount=TESTAll_Product_Analyzefile[['站点','计算销售日目标','MNSKU']]
TESTAll_Product_Analyzefile_adMount=TESTAll_Product_Analyzefile_adMount.rename(columns={'站点':'Country'})
TESTAll_Product_Analyzefile_adMount['站点'] = TESTAll_Product_Analyzefile_adMount['Country'].replace({
    '英国': 'NEW-UK',
    '法国': 'NEW-FR',
    '加拿大': 'NEW-CA',
    '日本': 'NEW-JP',
    '德国': 'NEW-DE',
    '西班牙': 'NEW-ES',
    '美国': 'NEW-US',
    '墨西哥': 'NEW-MX',
})
TESTAll_Product_Analyzefile_adMount['MNSKU'] = TESTAll_Product_Analyzefile_adMount['MNSKU'].str.split(',')

# 使用 explode 函数将 MNSKU 列进行拆分
TESTAll_Product_Analyzefile_adMount = TESTAll_Product_Analyzefile_adMount.explode('MNSKU')

All_adv=pd.merge(All_adv,TESTAll_Product_Analyzefile_adMount,how='left',left_on=["Country","主要SKU"],right_on=["Country","MNSKU"],suffixes=('', '_y'))

All_adv_simple=All_adv[["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","Campaign Status","Ad Group Status","Status","主要SKU","MergedSKU","好词标签","皮质层标签","计算销售日目标","Impressions1","Clicks1","Spend1","Orders1","Total Units1","Sales1","ACoS1","转化率1", "转化率2","Clicks","转化率Sum52","Clicks_26","转化率Sum26","Clicks_8","转化率Sum8","Clicks_4","转化率Sum4","Clicks_2","转化率Sum2"]]


Allbulkpath='D:\\运营\\2生成过程表\\'
#获取周六日期并转换为2023-5-13格式


last_saturday=find_last_saturday().strftime('%Y-%m-%d')
 

writer=pd.ExcelWriter(Allbulkpath+'All_ad_GV-'+last_saturday+'.xlsx')
 
All_adv.to_excel(writer,"All_adv",index=False) 
All_adv_simple.to_excel(writer,"All_adv_simple",index=False)
writer.close()
