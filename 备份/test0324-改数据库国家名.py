import sqlite3
import pandas as pd

# 创建一个与 SQLite 数据库文件的连接
conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')

# 使用 pandas 从 "Bulkfiles" 读取数据
df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)
 
df=df[df["日期"]<11]
df.to_excel(r'D:\\运营\周bulkfiles
 

# 关闭数据库连接
conn.close()
