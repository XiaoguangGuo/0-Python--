import os
import shutil


#将所有数据放到"新原始数据"下

source_folder = 'D:\\运营\\1数据源'
target_folder = 'D:\\运营\\HistoricalData'

raw_data_dir = "D:\\运营\\原始数据"
newraw_data_dir="D:\\运营\\NEW原始数据"
oldraw_data_dir = "D:\\运营\\HistoricalData\\原始数据\\"
#######################把原始数据移到历史数据\原始数据里
for oldfile in os.listdir(raw_data_dir):
    src_path = os.path.join(raw_data_dir, oldfile)
    dest_path = os.path.join(oldraw_data_dir, oldfile)
    shutil.move(src_path, dest_path)
#######################把新原始数据移到 D：原始数据里
for newfile in os.listdir(newraw_data_dir):
    src_path = os.path.join(newraw_data_dir, newfile)
    dest_path = os.path.join(raw_data_dir, newfile)
    shutil.move(src_path, dest_path)
    print("完成",newfile)

folder_mapping = {
    "周Bulk广告数据": "周Bulk广告数据",
    "bulkoperationfiles": "周Bulk广告数据",
    "周SearchTerm数据": "广告周数据",
    "计划数据\\老站\\销售数据": "计划数据\\老站\\销售数据",
    "计划数据\\老站\\当日库存": "计划数据\\老站\\当日库存",
    "计划数据\\老站\\在途库存": "计划数据\\老站\\在途库存",
    "Product_Analyze产品分析": "Product_Analyze产品分析"
}

for folder_name, target_name in folder_mapping.items():
    folder_path = os.path.join(source_folder, folder_name)
    target_path = os.path.join(target_folder, target_name)
    if os.path.exists(folder_path):
        if len(os.listdir(folder_path)) == 0:
            print(f"{folder_name} 文件夹为空")
        else:
            for file_name in os.listdir(folder_path):
                source_file_path = os.path.join(folder_path, file_name)
                target_file_path = os.path.join(target_path, file_name)
                shutil.move(source_file_path, target_file_path)
                print(f"Moved {source_file_path} to {target_file_path}")
    else:
        print(f"{folder_name} 文件夹不存在")
print("All files have been moved.")

################################################检查Bulk文件是否齐全名称是否对#####################################################################3
import os

source_folder = r'D:\运营\原始数据'

# Site mapping and missing site list
site_mapping = {
    'US': 'GV-US',
    'MX': 'GV-MX',
    'CA': 'GV-CA'
}


for file_name in os.listdir(source_folder):
    # Check if the file name contains "Bulk"
    if 'bulk' in file_name:
        # Get the site name before the first "_"
        site_name = file_name.split('_')[0]
        # Check if the site name is in the mapping dictionary
        if site_name in site_mapping:
            new_site_name = site_mapping[site_name]
            # Rename the file by replacing the site name with the new site name
            new_file_name = file_name.replace(site_name, new_site_name)
            # Move the file to the target folder
            source_file_path = os.path.join(source_folder, file_name)
            target_file_path = os.path.join(source_folder, new_file_name)
            os.rename(source_file_path, target_file_path)
            print(f"Renamed {source_file_path} to {target_file_path}")
        #else:
            #print(f"Warning: Site {site_name} is not in the site mapping list.")



def check_sites(source_folder):
    expected_sites = ["NEW-US", "NEW-CA", "NEW-MX", "NEW-JP", "NEW-DE", "NEW-FR", "NEW-ES", "NEW-UK", "HM-US", "GV-US", "GV-CA", "GV-MX"]
    actual_sites = []
    for foldername, subfolders, filenames in os.walk(source_folder):
        for filename in filenames:
            if 'bulk' in filename:
                site = filename.split('_')[0]
                actual_sites.append(site)
    actual_sites = list(set(actual_sites))
    print("实际站点名列表：", actual_sites)
    missing_sites = set(expected_sites) - set(actual_sites)
    if missing_sites:
        print("Bulk缺少以下站点名：", list(missing_sites))
    else:
        print("所有bulk站点名均已存在。")

# 运行程序


check_sites(source_folder)

#这个代码将遍历源文件夹中的文件，如果文件名包含"Bulk"，则获取站点名称。然后，它将检查站点名称是否在站点映射字典中，如果是，则将站点名称替换为映射字典中的新站点名称。
#最后，它将重命名文件并将其移动到原始数据文件夹中。如果站点名称不在映射字典中，则它将显示警告消息并将该站点名称添加到缺失的站点列表中。完成文件重命名后，它将打印缺少的站点列表。

######################################################################################检查销售文件##############33
import os

site_names = []

for file_name in os.listdir(r'D:\运营\原始数据'):
    if file_name.endswith(".csv") and "BusinessReport" in file_name:
        site_name = file_name.split("_")[0]
        if site_name in ["US", "CA", "MX"]:
            site_names.append(site_name)

# 检查缺失的站点名
missing_sites = list(set(["US", "CA", "MX"]) - set(site_names))
if missing_sites:
    print("缺失的站点名：", missing_sites)
else:
    print("3. 所有BusinessReport站点名都存在。")
################################################检查当日库存#######################################################################
import os
import re

site_names = []

for file_name in os.listdir("D:\\运营\\原始数据"):
    if file_name.endswith(".csv"):
         
        parts = file_name.split("_")
        if len(parts) > 1 and re.match("^[a-zA-Z]+$", parts[0]) and re.match("^\d{9,13}", parts[1]):
            site_name = parts[0]
            print(site_name)
            if site_name in ["US", "CA", "MX"]:
                print(site_name)
                site_names.append(site_name)

# 检查缺失的站点名
missing_sites = list(set(["US", "CA", "MX"]) - set(site_names))
if missing_sites:
    print("当日库存缺失的站点名：", missing_sites)
else:
    print("当日库存所有站点名都存在。")
###################################################################33

import os

site_names = []

for file_name in os.listdir("D:\\运营\\原始数据"):
    if file_name.endswith(".tsv"):
        parts = file_name.split("_")
        if len(parts) > 1:
            site_name = parts[0]
            if site_name in ["US", "CA", "MX"]:
                site_names.append(site_name)

# 检查缺失的站点名
missing_sites = list(set(["US", "CA", "MX"]) - set(site_names))
if missing_sites:
    print("在途库存缺失的站点名：", missing_sites)
else:
    print("在途库存所有站点名都存在。")
######################################################################
input("确认检查结果，如果没问题则回车-拷贝入D运营，如有问题则相应修改和补充")
###########################################复制##############################################333
import os
import shutil

# 定义文件路径
data_path = "D:\\运营\\原始数据"
sales_path = "D:\\运营\\1数据源\\计划数据\\老站\\销售数据"
stock_path = "D:\\运营\\1数据源\\计划数据\\老站\\当日库存"
transit_path = "D:\\运营\\1数据源\\计划数据\\老站\\在途库存"
searchterm_path = "D:\\运营\\1数据源\\周SearchTerm数据"
bulk_path = "D:\\运营\\1数据源\\周Bulk广告数据"

# 复制 BusinessReport 文件到销售数据文件夹中
for file in os.listdir(data_path):

    if not "BusinessReport" in file:
        continue

    if not re.match(r'^(US|CA|MX)_\d{4}-\d{1,2}-\d{1,2}_[\w\(\)\s-]+\.csv$',file):
        print(f"Invalid filename: {file}")
        input("修改文件名")
        continue

    #旧的不要了if "BusinessReport" in file and file.endswith(".csv"):
    shutil.copy(os.path.join(data_path, file), sales_path)





        

# 复制数字格式的 CSV 文件到当日库存文件夹中
for file in os.listdir(data_path):

    
    if file.endswith(".csv"):
        file_prefix = file.split("_")[0]
        if file_prefix.isalpha() and file.split(".")[0][-13:-4].isdigit() and 9 <= len(file.split(".")[0][-13:-4]) <= 13:
            shutil.copy(os.path.join(data_path, file), stock_path)


import os
import shutil



valid_prefix = ['US', 'CA', 'MX']

for file in os.listdir(data_path):
    if file.endswith('.csv'):
        filename = os.path.splitext(file)[0]
        filename_parts = filename.split('_')
        if len(filename_parts) == 2 and len(filename_parts[0]) == 2 and filename_parts[0] in valid_prefix \
                and filename_parts[1].isdigit() and 9 <= len(filename_parts[1]) <= 13:
            source_path = os.path.join(data_path, file)
            destination_path = os.path.join(stock_path, file)
            shutil.copy2(source_path, destination_path)
            print(f'Copied {file} to {stock_path}')




            

# 复制 TSV 文件到在途库存文件夹中
for file in os.listdir(data_path):
    if file.endswith(".tsv"):
        file_prefix = file.split("_")[0]
        if file_prefix in ["US", "CA", "MX"]:
            shutil.copy(os.path.join(data_path, file), transit_path)

# 复制 Sponsored 文件到周 SearchTerm 数据文件夹中
for file in os.listdir(data_path):
    if "Sponsored" in file:
        shutil.copy(os.path.join(data_path, file), searchterm_path)

# 复制 bulk 文件到周 Bulk 广告数据文件夹中
for file in os.listdir(data_path):
    if "bulk" in file:
        shutil.copy(os.path.join(data_path, file), bulk_path)
