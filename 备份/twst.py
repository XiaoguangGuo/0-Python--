import sqlite3

# 连接到数据库
conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')
c = conn.cursor()

# 更新 Spend, Max Bid, Sales 列中的逗号为小数点
c.execute("UPDATE Bulkfiles SET Spend = REPLACE(Spend, ',', '.'), [Max Bid] = REPLACE([Max Bid], ',', '.'), Sales = REPLACE(Sales, ',', '.')")

# 提交更改并关闭连接
conn.commit()
conn.close()


