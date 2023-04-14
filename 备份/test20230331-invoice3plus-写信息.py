import pandas as pd
import os

# 定义路径和文件名
template_dir = "D:\\运营\\发票模板"
template_file = "template.xlsx"
base_info_file = "发票基础信息表.xlsx"
sheet_name = "列名对应"

# 读取模板文件并找到表头所在行
template_path = os.path.join(template_dir, template_file)
template_df = pd.read_excel(template_path, header=None)
header_row = None
for i in range(len(template_df)):
    if template_df.iloc[i].isnull().all() and i > 0:
        if template_df.iloc[i-1].notnull().any():
            header_row = i-1
            break

# 提取表头信息，设置列名
if header_row is not None:
    template_df = template_df.iloc[header_row:]
    template_df.columns = template_df.iloc[0]
    template_df = template_df[1:]
    template_df = template_df.reset_index(drop=True)
    template_df["模板列名"] = template_df.iloc[0,1:].tolist()

# 读取基础信息文件中的列名对应sheet，并将空值替换为empty
base_info_path = os.path.join(template_dir, base_info_file)
base_info_df = pd.read_excel(base_info_path, sheet_name=sheet_name)
base_info_df = base_info_df.fillna("empty")

# 将基础信息文件中的列名对应信息merge到模板dataframe中
merged_df = pd.merge(template_df, base_info_df, on="模板列名")

# 提取中间列名列的list
middle_col_list = merged_df["中间列名"].tolist()

# 输出结果
print(middle_col_list)
