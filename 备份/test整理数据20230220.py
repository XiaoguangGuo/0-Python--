import os
import shutil

# 定义原始数据文件夹路径
raw_data_path = "D:\\运营\\原始数据"

# 定义历史数据文件夹路径
historical_data_path = "D:\\运营\\HistoricalData\\计划数据\\老站\\销售数据"

# 定义销售数据文件夹路径
sales_data_path = "D:\\运营\\1数据源\\计划数据\\老站\\销售数据"

# 如果历史数据文件夹不存在，则创建
if not os.path.exists(historical_data_path):
    os.makedirs(historical_data_path)

# 遍历销售数据文件夹下所有文件
for root, dirs, files in os.walk(sales_data_path):
    for file in files:
        file_path = os.path.join(root, file)
        # 移动文件到历史数据文件夹中
        shutil.move(file_path, historical_data_path)

# 遍历原始数据文件夹下所有文件
for root, dirs, files in os.walk(raw_data_path):
    for file in files:
        # 如果文件名包含"BusinessReport"且为csv文件
        if "BusinessReport" in file and file.endswith(".csv"):
            file_path = os.path.join(root, file)
            # 复制文件到销售数据文件夹中
            shutil.copy(file_path, sales_data_path)
