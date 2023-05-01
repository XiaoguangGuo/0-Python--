import pandas as pd
import sqlite3

# 连接到数据库
conn = sqlite3.connect('rules.db')

# 创建规则表
conn.execute('''CREATE TABLE rules
             (id INTEGER PRIMARY KEY,
             rule TEXT,
             algorithm TEXT,
             algorithm_code TEXT);''')

# 插入一些规则
conn.execute("INSERT INTO rules (rule, algorithm, algorithm_code) \
              VALUES ('rule1', 'algorithm1', 'def algorithm1():\\n    # code for algorithm1\\n    pass')")
conn.execute("INSERT INTO rules (rule, algorithm, algorithm_code) \
              VALUES ('rule2', 'algorithm2', 'def algorithm2():\\n    # code for algorithm2\\n    pass')")
conn.execute("INSERT INTO rules (rule, algorithm, algorithm_code) \
              VALUES ('rule3', 'algorithm3', 'def algorithm3():\\n    # code for algorithm3\\n    pass')")

# 查询规则
cursor = conn.execute("SELECT * from rules")
for row in cursor:
    print("ID = ", row[0])
    print("Rule = ", row[1])
    print("Algorithm = ", row[2])
    print("Algorithm Codes = ", row[3])
    print("\n")

# 获取算法代码
def get_algorithm(rule):
    cursor = conn.execute("SELECT algorithm_code from rules WHERE rule=?", (rule,))
    code = cursor.fetchone()
    if code is not None:
        return code[0]
    else:
        return None

# 读取ProductActions表
plan = pd.read_excel("ProductActions.xlsx")

# 添加Algorithm Codes列
plan["Algorithm Codes"] = ""

# 遍历规则表，执行算法
for row in conn.execute("SELECT rule, algorithm from rules"):
    rule = row[0]
    algorithm_name = row[1]
    algorithm_code = get_algorithm(rule)
    if algorithm_code is not None:
        exec(algorithm_code)
        plan.loc[plan["皮质层标签"].str.contains(rule), "Algorithm Codes"] += algorithm_code

# 输出ProductActions表
print(plan)

# 关闭数据库连接
conn.close()
