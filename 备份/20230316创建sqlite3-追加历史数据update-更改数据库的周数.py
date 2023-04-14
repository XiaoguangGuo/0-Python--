import pandas as pd
import sqlite3
from datetime import datetime, timedelta

def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    df['日期'] = pd.to_datetime(df['日期'])
    df['周数'] = ((last_saturday - df['日期']).dt.days // 7) + 1
    return df

# 连接到SQLite3数据库
conn = sqlite3.connect("Amazondata.db")

# 从数据库中读取bulkfiles表
df = pd.read_sql_query("SELECT * from bulkfiles", conn)

# 更新DataFrame中的周数
updated_df = update_week_numbers(df)

# 将更新后的DataFrame数据写回到数据库中
updated_df.to_sql("bulkfiles", conn, if_exists="replace", index=False)

# 关闭数据库连接
conn.close()
