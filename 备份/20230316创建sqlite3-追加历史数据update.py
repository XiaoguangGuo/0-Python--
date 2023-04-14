import sqlite3
import pandas as pd
from datetime import datetime, timedelta
# 读取 Excel 文件
pd.set_option('display.max_rows', None)
conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')




# 使用 pandas 从 your_table_name 读取数据
df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)


# 获取行数
row_count = len(df)

# 获取列名
column_names = df.columns

print(f"行数: {row_count}")
print("列名:")
for column_name in column_names:
    print(column_name)


# 获取并去重日期列
unique_dates = df['日期'].drop_duplicates().sort_values().reset_index(drop=True)
print(unique_dates)

# 从 2022 年开始检查日期是否中断
start_year = 2022
date_intervals_are_continuous = True

for i in range(len(unique_dates) - 1):
    date1 = datetime.strptime(unique_dates[i], '%Y-%m-%d %H:%M:%S')
    date2 = datetime.strptime(unique_dates[i + 1], '%Y-%m-%d %H:%M:%S')
    
    if date1.year >= start_year and (date2 - date1).days != 7:
        date_intervals_are_continuous = False
        if (date2 - date1).days !=7:
            print(f"间隔不等于7 的两个日期：{unique_dates[i]} 和 {unique_dates[i + 1]}")
        break

if date_intervals_are_continuous:
    print("从 2022 年起，日期没有中断。")
else:
    print("从 2022 年起，日期存在中断。")
    
# 检查并去除重复行
duplicated_rows = df.duplicated().sum()
if duplicated_rows > 0:
    print(f"发现 {duplicated_rows} 个重复行。正在去重...")
    df = df.drop_duplicates()

    # 将去重后的数据更新到 your_table_name
    with conn:
        c = conn.cursor()
        c.execute("DELETE FROM [Bulkfiles]")
        df.to_sql("Bulkfiles", conn, if_exists='append', index=False)
    print("已更新 your_table_name 表，去除重复行。")
else:
    print("没有发现重复行。")

# 关闭数据库连接
conn.close()

