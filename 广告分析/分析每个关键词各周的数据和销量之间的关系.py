import sqlite3
import pandas as pd

# 连接 SQLite 数据库并读取数据
conn = sqlite3.connect('bulk_files.db')
df = pd.read_sql_query("SELECT * from bulk_files", conn)

# 选取需要的列
df = df[['Country', 'Campaign Name', 'Ad Group Name', 'Match Type', 'Targeting', 'Week', 'Impressions', 'Clicks', 'Orders', 'Total Spend']]

# 将关键词和匹配方式拆分为两列
df[['Keyword', 'Match Type']] = df['Targeting'].str.split('_', expand=True)

# 将 Week 转换为整数类型
df['Week'] = df['Week'].astype(int)

# 计算转化率
df['Conversion Rate'] = df['Orders'] / df['Clicks']

# 透视表格，汇总每个关键词每周的数据
pivot_df = df.pivot_table(index=['Country', 'Campaign Name', 'Ad Group Name', 'Keyword', 'Match Type'], columns=['Week'], values=['Impressions', 'Clicks', 'Orders', 'Conversion Rate'], aggfunc='sum', fill_value=0)

# 将透视表格展平
flat_df = pd.DataFrame(pivot_df.to_records())

# 重新排列列顺序
cols = ['Country', 'Campaign Name', 'Ad Group Name', 'Keyword', 'Match Type']
for i in range(1, 53):
    cols += [('Impressions', i), ('Clicks', i), ('Orders', i), ('Conversion Rate', i)]
flat_df = flat_df[cols]

# 保存为 CSV 文件
flat_df.to_csv('keyword_stats.csv', index=False)
