import pandas as pd
import sqlite3
from datetime import datetime, timedelta

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

def correlations_SKU_Campaign(df, columns):
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

def update_bulkfiles_week_numbers(database_path, input_weeks):
    last_saturday = find_last_saturday()
    start_date = last_saturday - timedelta(weeks=input_weeks)
    conn = sqlite3.connect(database_path)
    df = pd.read_sql_query(f"SELECT * FROM Bulkfiles WHERE 日期 >= '{start_date}'", conn)
    df = df[df['日期'].notna()]
    df['Spend'] = df['Spend'].astype(float)
    conn.close()
    return update_week_numbers(df)

input_weeks = input("请输入考察周数，直接回车为默认52周：")
input_weeks = int(input_weeks) if input_weeks else 52
Bulkfiles_DF = update_bulkfiles_week_numbers(r'D:\运营\sqlite\AmazonData.db', input_weeks)
Bulkfiles_DF = Bulkfiles_DF.loc[Bulkfiles_DF['SKU'].notnull(), ['Country', 'SKU', 'Campaign', 'Spend', '周数']]
Bulkfiles_DF=Bulkfiles_DF[Bulkfiles_DF['周数']<(input_weeks-1)]
Sales_DF = pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx', sheet_name='Sheet1')
Sales_DF = update_week_numbers(Sales_DF)
Sales_DF = Sales_DF.loc[Sales_DF['SKU'].notnull(), ['Country', 'SKU', 'Units Ordered', '周数']]
Sales_DF=Sales_DF[Sales_DF['周数']<(input_weeks-1)]
Bulkfiles_DF = Bulkfiles_DF[Bulkfiles_DF['Country'].isin(Sales_DF['Country'])]
Bulkfiles_DF2=Bulkfiles_DF
Bulkfiles_DF = pd.merge(Bulkfiles_DF, Sales_DF, how='outer', on=['Country', 'SKU', '周数']).fillna(0)

Bulkfiles_DF.to_excel(r'D:\运营\2生成过程表\Bulkfiles_DFforcorrel.xlsx', index=False)
corr_SKU_Campaign_DF = correlations_SKU_Campaign(Bulkfiles_DF, ['Country', 'SKU', 'Campaign'])
corr_SKU_Campaign_DF.to_excel(r'D:\运营\2生成过程表\correlations_SKU_Campaign.xlsx', index=False)

# 按 Country 和 SKU 对
# 按 Country 和 SKU 对 Spend 汇总
grouped_spend_df = Bulkfiles_DF2.groupby(['Country', "周数",'SKU']).agg({'Spend': 'sum'}).reset_index()

# 将汇总的 Spend 与 Sales_DF 合并
merged_sales_spend_df = pd.merge(Sales_DF, grouped_spend_df, on=['Country', "周数",'SKU'], how='outer').fillna(0)
#输出至 Excel
merged_sales_spend_df.to_excel(r'D:\运营\2生成过程表\merged_sales_spend.xlsx', index=False)
# 计算按 Country 和 SKU 汇总的相关性
corr_SKU_summary_DF = correlations_SKU_Campaign(merged_sales_spend_df, ['Country', 'SKU'])

# 将按 Country 和 SKU 汇总的相关性输出至 Excel
corr_SKU_summary_DF.to_excel(r'D:\运营\2生成过程表\correlations_SKU_summary.xlsx', index=False)

def calculate_moving_periods_correlations(df, start_week, end_week, period_length):
    df = df[(df['周数'] >= start_week) & (df['周数'] < start_week + period_length)]
    return correlations_SKU_Campaign(df, ['Country', 'SKU'])

# 定义移动时间段的起始周数
moving_periods_start_weeks = [
    1,
    2,
    3,
    4,
    5,
    6
]

period_length = 8

# 初始化移动时间段的结果 DataFrame
moving_periods_corr_df = pd.DataFrame()

# 计算指定移动时间段的相关系数
for start_week in moving_periods_start_weeks:
    end_week = start_week + period_length - 1
    period_label = 'Recent 8 weeks' if start_week == 1 else f'Weeks {start_week} to {end_week}'
    period_corr = calculate_moving_periods_correlations(merged_sales_spend_df, start_week, end_week, period_length)
    period_corr['Time Period'] = period_label
    moving_periods_corr_df = pd.concat([moving_periods_corr_df, period_corr], ignore_index=True)

# 将移动时间段的相关性输出至 Excel
moving_periods_corr_df.to_excel(r'D:\运营\2生成过程表\moving_periods_correlations.xlsx', index=False)



