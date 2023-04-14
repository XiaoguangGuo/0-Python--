import sqlite3
import pandas as pd
from datetime import datetime, timedelta

# 连接到 SQLite 数据库
conn = sqlite3.connect('D:\\运营\\sqlite\\your_database_name.db')

# 使用 pandas 从 your_table_name 读取数据
df = pd.read_sql_query('SELECT * FROM your_table_name', conn)

# 更新周数
max_date = df['日期'].max()
df['日期'] = pd.to_datetime(df['日期'])
df['周数'] = ((df['日期'] - max_date) // timedelta(days=7)).abs() + 1

# 创建数据透视表
pivot_df = df.groupby(['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting',
                       'Match Type', 'Campaign Status', 'Ad Group Status', 'Status', '周数']).agg({
                           'Impressions': 'sum',
                           'Clicks': 'sum',
                           'Spend': 'sum',
                           'Orders': 'sum',
                           'Total Units': 'sum',
                           'Sales': 'sum'
                       }).reset_index()

# 计算周转率
pivot_df['周转率'] = pivot_df['Orders'] / pivot_df['Clicks']

# 对每一周的数据进行排序
sorted_weeks = sorted(pivot_df['周数'].unique())
sorted_columns = ['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting',
                  'Match Type', 'Campaign Status', 'Ad Group Status', 'Status', '周转率']

for week in sorted_weeks:
    week_df = pivot_df[pivot_df['周数'] == week].copy()
    week_df['Impressions'] = week_df['Impressions'].astype(str) + ' (周 ' + week_df['周数'].astype(str) + ')'
    week_df['Clicks'] = week_df['Clicks'].astype(str) + ' (周 ' + week_df['周数'].astype(str) + ')'
    week_df['Spend'] = week_df['Spend'].astype(str) + ' (周 ' + week_df['周数'].astype(str) + ')'
    week_df['Orders'] = week_df['Orders'].astype(str) + ' (周 ' + week_df['周数'].astype(str) + ')'
    week_df['Total Units'] = week_df['Total Units'].astype(str) + ' (周 ' + week_df['周数'].astype(str) + ')'
    week_df['Sales'] = week_df['Sales'].astype(str) + ' (周 ' + week_df['周数'].astype(str) + ')'
    
    sorted_columns += ['Impressions', 'Clicks', 'Spend', 'Orders', 'Total Units', 'Sales']

    pivot_df = pivot_df.merge(week_df, on=['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting',
                                           'Match Type', 'Campaign Status', 'Ad Group Status', 'Status', '周数'],
                              how='left', suffixes=('', '_y'))

    # 删除重复的合并列
    pivot_df.drop(columns=[col for col in pivot_df.columns if col.endswith('_y')], inplace=True)

# 按指定的列顺序排序
pivot_df = pivot_df[sorted_columns]

# 保存结果到 Excel 文件
pivot_df.to_excel('D:\\运营\\sqlite\\result.xlsx', index=False)
