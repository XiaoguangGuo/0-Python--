import sqlite3

# 创建一个与 SQLite 数据库文件的连接。
conn = sqlite3.connect('D:\\运营\\sqlite\\AmazonData.db')
c = conn.cursor()

# 创建表
c.execute('''
CREATE TABLE IF NOT EXISTS Bulkfiles (
    "Record ID" INTEGER,
    "Record Type" TEXT,
    "Campaign ID" TEXT,
    Campaign TEXT,
    "Campaign Daily Budget" REAL,
    "Portfolio ID" TEXT,
    "Campaign Start Date" TEXT,
    "Campaign End Date" TEXT,
    "Campaign Targeting Type" TEXT,
    "Ad Group" TEXT,
    "Max Bid" REAL,
    "Keyword or Product Targeting" TEXT,
    "Product Targeting ID" TEXT,
    "Match Type" TEXT,
    SKU TEXT,
    "Campaign Status" TEXT,
    "Ad Group Status" TEXT,
    Status TEXT,
    Impressions INTEGER,
    Clicks INTEGER,
    Spend REAL,
    Orders INTEGER,
    "Total Units" INTEGER,
    Sales REAL,
    ACoS REAL,
    "Bidding strategy" TEXT,
    "Placement Type" TEXT,
    "Increase bids by placement" REAL,
    Country TEXT,
    日期 TEXT,
    周数 INTEGER
)
''')

# 提交更改并关闭连接
conn.commit()
conn.close()







