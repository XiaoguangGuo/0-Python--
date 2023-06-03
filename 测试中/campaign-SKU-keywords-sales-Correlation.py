import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import openpyxl

def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    date_column = "日期" if "日期" in df.columns else "Date"
    df[date_column] = pd.to_datetime(df[date_column])
    df['周数'] = ((last_saturday - df[date_column]).dt.days // 7) + 1
    return df

def correlations_SKU_Campaign(df, columns):#计算相关系数的函数，按coloumns进行分组，计算相关系数，返回结果
    grouped_df = df.groupby(columns)
    result_df = pd.DataFrame(columns=columns + ['相关系数'])

    for group, group_df in grouped_df:
        if 'Spend' in group_df.columns and 'Units Ordered' in group_df.columns:
            correlation = group_df['Spend'].corr(group_df['Units Ordered'])
            temp_data = {key: value for key, value in zip(columns, group)}
            temp_data.update({'相关系数': correlation})
            temp_df = pd.DataFrame(temp_data, index=[0])
            result_df = pd.concat([result_df, temp_df], ignore_index=True)

    return result_df

def update_bulkfiles_week_numbers(database_path, input_weeks):#更新Bulkfiles的周数的函数
    last_saturday = find_last_saturday()
    start_date = last_saturday - timedelta(weeks=input_weeks)
    conn = sqlite3.connect(database_path)
    df = pd.read_sql_query(f"SELECT * FROM Bulkfiles WHERE 日期 >= '{start_date}'", conn)
    df = df[df['日期'].notna()]
    df['Spend'] = df['Spend'].astype(float)
    conn.close()
    return update_week_numbers(df)


#以下是主程序

input_weeks = input("请输入考察周数，直接回车为默认52周：")#输入考察周数
input_weeks = int(input_weeks) if input_weeks else 52

Bulkfiles_DF = update_bulkfiles_week_numbers(r'D:\运营\sqlite\AmazonData.db', input_weeks)#获取数据并加周数
Bulkfiles_DF_keyword = Bulkfiles_DF.loc[Bulkfiles_DF['Keyword or Product Targeting'].notnull(), ['Country', 'Campaign', 'Ad Group','Keyword or Product Targeting','Match Type','Spend', '周数']]#筛选列

Bulkfiles_DF=Bulkfiles_DF[Bulkfiles_DF['周数']<(input_weeks-1)]#筛选周数
Bulkfiles_DF_keyword=Bulkfiles_DF_keyword[Bulkfiles_DF_keyword ['周数']<(input_weeks-1) ]

def process_spend_summary(BulkFile_df,Bulkfiles_DF_keyword):
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
    merge_columns = ["Country", "Campaign", "Ad Group", "主要SKU"] if has_country else ["Campaign", "Ad Group", "主要SKU"]
    BulkFile_df_keywords_mainSKU = Bulkfiles_DF_keyword.merge(spend_summary[merge_columns], on=group_columns, how="left")


    # 将结果合并到周BUlkFIle原始数据表，创建一个新列 "主要SKU"
 

    return BulkFile_df_keywords_mainSKU


Bulkes_DF_withMainSKU=process_spend_summary(Bulkfiles_DF,Bulkfiles_DF_keyword)

print(Bulkes_DF_withMainSKU)


#以下获取销售数据
Sales_DF = pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx', sheet_name='Sheet1')#读取周销售数据总表
Sales_DF = update_week_numbers(Sales_DF)#加周数
Sales_DF = Sales_DF.loc[Sales_DF['SKU'].notnull(), ['Country', 'SKU', 'Units Ordered', '周数']]#筛选列
Sales_DF=Sales_DF[Sales_DF['周数']<(input_weeks-1)]#筛选周数
print(Sales_DF)
#筛选数据
Bulkfiles_DF2 = Bulkes_DF_withMainSKU[Bulkes_DF_withMainSKU['Country'].isin(Sales_DF['Country'])]#筛选广告数据的国家
print(Bulkfiles_DF2.columns)


#对广告数据按Country，Campai

Bulkfiles_DF2.rename(columns={"主要SKU":"SKU"}, inplace=True)
print(Bulkfiles_DF2.columns)
 
#将销售数据合并到广告数据
Bulkfiles_DF3 = pd.merge(Bulkfiles_DF2, Sales_DF, how='outer', on=['Country', 'SKU', '周数']).fillna(0)#合并两个表

 



#计算按 Country ，Campaign和 SKU 汇总的相关性
column=['Country','Campaign','SKU','Ad Group','Keyword or Product Targeting','Match Type']
corr_SKU_Campaign_DF = correlations_SKU_Campaign(Bulkfiles_DF3,column)
corr_SKU_Campaign_DF.to_excel(r'D:\运营\2生成过程表\correlations_SKU_Campaign.xlsx', index=False)


##以下为包含Campaign的移动时间段相关性计算
def calculate_moving_periods_Campaign_correlations(df, start_week, end_week):
    df = df[(df['周数'] >= start_week) & (df['周数'] <= end_week)]
    return correlations_SKU_Campaign(df, ['Country', 'SKU', 'Campaign','Ad Group','Keyword or Product Targeting','Match Type'])

# 定义移动时间段
moving_periods = [
    (1, input_weeks-1),
    (1, 8),
    (2, 9),
    (3, 10),
    (4, 11),
    (5, 12)
]
# 计算带Campaign的数据的指定移动时间段的相关系数

moving_periods_corr_Campaign_df = pd.DataFrame()

for start_week, end_week in moving_periods:
    
    if start_week == 1 and end_week == input_weeks - 1:
        period_label = f'Weeks 1 to {input_weeks - 1}'
    else:
        period_label = f'Weeks {start_week} to {end_week}'

    period_corr_Campaign = calculate_moving_periods_Campaign_correlations(Bulkfiles_DF, start_week, end_week)

    period_corr_Campaign['Time Period'] = period_label

    period_corr_Campaign.to_excel(r'D:\运营\2生成过程表\period_corr_Campaign'+str(period_label)+".xlsx", engine="openpyxl",index=False)
    moving_periods_corr_Campaign_df = pd.concat([moving_periods_corr_Campaign_df, period_corr_Campaign], ignore_index=True)
    

combined_corr_campaign_df=moving_periods_corr_Campaign_df
combined_corr_campaign_df.to_excel(r'D:\运营\2生成过程表\combined_correlations_Campaign.xlsx', index=False)
#把combined_corr_campaign_df 按Country，SKU，Campaign进行Pivot，横轴为Time Period，纵轴为相关系数，然后输出成excel
combined_corr_campaign_df = combined_corr_campaign_df.pivot_table(index=['Country', 'SKU', 'Campaign','Ad Group','Keyword or Product Targeting','Match Type'], columns='Time Period', values='相关系数', aggfunc='first').reset_index()
combined_corr_campaign_df.to_excel(r'D:\运营\2生成过程表\combined_correlations_Campaign_pivot.xlsx', index=False)

#以下为计算不包含Campaign的相关性和移动时间段相关性
##计算按 Country 和 SKU 汇总的相关性
####按 Country 和 SKU 对 Spend 汇总
grouped_spend_df = Bulkfiles_DF2.groupby(['Country', "周数",'SKU']).agg({'Spend': 'sum'}).reset_index()


# 将汇总的 Spend 与 Sales_DF 合并
merged_sales_spend_df = pd.merge(Sales_DF, grouped_spend_df, on=['Country', "周数",'SKU'], how='outer').fillna(0)

#输出至 Excel
merged_sales_spend_df.to_excel(r'D:\运营\2生成过程表\merged_sales_spend.xlsx', index=False)


# 计算按 Country 和 SKU 汇总的相关性
corr_SKU_summary_DF = correlations_SKU_Campaign(merged_sales_spend_df, ['Country', 'SKU'])

# 将按 Country 和 SKU 汇总的相关性输出至 Excel
corr_SKU_summary_DF.to_excel(r'D:\运营\2生成过程表\correlations_SKU_summary.xlsx', index=False)

##以下为移动时间段相关性计算

def calculate_moving_periods_correlations(df, start_week, end_week):
    df = df[(df['周数'] >= start_week) & (df['周数'] <= end_week)]
    return correlations_SKU_Campaign(df, ['Country', 'SKU'])

# 定义移动时间段
moving_periods = [
    (1, input_weeks-1),
    (1, 8),
    (2, 9),
    (3, 10),
    (4, 11),
    (5, 12)
]
# 计算指定移动时间段的相关系数

moving_periods_corr_df = pd.DataFrame()


for start_week, end_week in moving_periods:
    if start_week == 1 and end_week == input_weeks - 1:
        period_label = f'Weeks 1 to {input_weeks - 1}'
    else:
        period_label = f'Weeks {start_week} to {end_week}'
    period_corr = calculate_moving_periods_correlations(merged_sales_spend_df, start_week, end_week)
    period_corr['Time Period'] = period_label
    moving_periods_corr_df = pd.concat([moving_periods_corr_df, period_corr], ignore_index=True)

combined_corr_df = moving_periods_corr_df
combined_corr_df.to_excel(r'D:\运营\2生成过程表\combined_correlations.xlsx', index=False)
#把combined_corr_campaign_df 按Country，SKU，进行Pivot，横轴为Time Period，纵轴为相关系数，然后输出成excel
combined_corr_df = combined_corr_df.pivot_table(index=['Country', 'SKU'], columns='Time Period', values='相关系数', aggfunc='first').reset_index()
combined_corr_df.to_excel(r'D:\运营\2生成过程表\combined_correlations_pivot.xlsx', index=False)




