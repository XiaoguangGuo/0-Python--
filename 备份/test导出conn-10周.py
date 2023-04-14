import sqlite3
import pandas as pd
from datetime import datetime, timedelta

# 创建 SQLite 数据库连接
conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')

# 从数据库中读取数据
df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)



# 计算重复行的数量
duplicate_rows_count = df.duplicated().sum()

if duplicate_rows_count > 0:
    print(f"发现 {duplicate_rows_count} 个重复行。")
else:
    print("没有发现重复行。")

# 关闭数据库连接
conn.close()



# 将字符串转换为 datetime 对象
target_date = datetime.strptime('2023-03-25', '%Y-%m-%d')

# 计算 10 周之前的日期
ten_weeks_before = target_date - timedelta(weeks=10)

# 将日期列转换为 datetime 类型
df['日期'] = pd.to_datetime(df['日期'])

# 筛选出日期列在范围内的数据
filtered_df = df[(df['日期'] >= ten_weeks_before) & (df['日期'] <= target_date)]

# 将筛选后的数据导出到 Excel 文件
filtered_df.to_excel(r'D:\\运营\\周bulk数据检查.xlsx', index=False)

# 关闭数据库连接
conn.close()
