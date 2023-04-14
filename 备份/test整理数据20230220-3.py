import os
import shutil
import re
import os
import shutil
import re

# Define source and destination directories
src_dir = "D:\\运营\\1数据源\\计划数据\\老站\\销售数据"
dest_dir = "D:\\运营\\HistoricalData\\计划数据\\老站\\销售数据"
raw_data_dir = "D:\\运营\\原始数据"

# Move files from source to destination directory, overwriting same named files
for filename in os.listdir(src_dir):
    src_path = os.path.join(src_dir, filename)
    dest_path = os.path.join(dest_dir, filename)
    shutil.move(src_path, dest_path)

# Copy matching files from raw data directory to destination directory
pattern=re.compile(r'(US|MX|CA)_\d{4}-\d{1,2}-\d{1,2}_BusinessReport\.csv$')
for filename in os.listdir(raw_data_dir):

    if pattern.match(filename):
        continue

    else:
        print(f"Invalid filename: {filename}")
        

    src_path = os.path.join(raw_data_dir, filename)
    dest_path = os.path.join(dest_dir, filename)
    shutil.copy(src_path, dest_path)
print("处理完成。")


# Define source and destination directories
source_dir2 = r"D:\运营\1数据源\计划数据\老站\当日库存"
dest_dir2 = r"D:\运营\HistoricalData\计划数据\老站\当日库存"
raw_dir2 = r"D:\运营\原始数据"

# Move files from source directory to destination directory
for file2 in os.listdir(source_dir2):
    src_path2 = os.path.join(source_dir2, file2)
    dest_path2= os.path.join(dest_dir2, file2)
    shutil.move(src_path2, dest_path2)

# Copy CSV files from raw directory to destination directory
for file2 in os.listdir(raw_dir2):
    if  re.match(r"(US|CA|MX)_\d{12,15}\.csv", file2):
        src_path2 = os.path.join(raw_dir2, file2)
        dest_path2 = os.path.join(dest_dir2, file2)
        shutil.copy2(src_path2, dest_path2)
