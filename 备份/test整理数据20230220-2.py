import os
import shutil
import re

# 源目录
source_dir = r'D:\运营\原始数据'
# 销售数据目标目录
target_dir = r'D:\运营\1数据源\计划数据\老站\销售数据'
# 历史数据目录
historical_data_dir = r'D:\运营\HistoricalData\计划数据\老站\销售数据'

# 创建HistoricalData文件夹
if not os.path.exists(historical_data_dir):
    os.makedirs(historical_data_dir)

# 移动原有的文件到HistoricalData目录中
for file_name in os.listdir(target_dir):
    file_path = os.path.join(target_dir, file_name)
    if os.path.isfile(file_path):
        shutil.move(file_path, historical_data_dir)

# 拷贝BusinessReport文件到销售数据目录中
for file_name in os.listdir(source_dir):
    if 'BusinessReport' in file_name and file_name.endswith('.csv'):
        file_path = os.path.join(source_dir, file_name)
        # 复制文件到销售数据目录中
        shutil.copy(file_path, target_dir)
        # 检查文件名是否符合要求
        if not re.match(r'(US_|CA_|MX_).*', file_name):
            print(f"文件名{file_name}不是以US_或CA_或MX_开头，跳过。")
        continue
# 检查2：检查日期格式是否为“2023-2-18”
        file_date = file_name.split("")[1]
        if not re.match(r'\d{4}-\d{1,2}-\d{1,2}', file_date):
            print(f"文件名{file_name}日期格式不是“2023-2-18”，跳过。")
        continue
