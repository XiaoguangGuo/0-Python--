
# -*- coding:utf-8 –*-
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil

import pandas as pd

# 从G盘根目录读取bulk-a2u3r73mnhpwps-20230316-20230317-1678902566983的xlsx文件
bulk_file = "G:/bulk-a2u3r73mnhpwps-20230316-20230317-1678902566983.xlsx"
bulk_df = pd.read_excel(bulk_file,sheet_name="Sponsored Products Campaigns")

# 从D盘/运营/2生成过程表中读取 “周bulk数据Summary”的XLSX表的 "Keyword-Campaign-WEEK" 工作表
summary_file = "D:/运营/2生成过程表/周bulk数据Summary.xlsx"
summary_df = pd.read_excel(summary_file, sheet_name="Keyword-Campaign-WEEK")

# 汇总周bulk数据Summary所有周的各列数据
summary_df = summary_df.groupby(['Country', 'Campaign', 'Keyword or Product Targeting', 'Ad Group']).sum().reset_index()
print(summary_df)
# 选择Country为GV-US的数据
summary_df_us = summary_df[summary_df['Country'] == 'GV-US']

# 按照 Country, Campaign, Keyword or Product Targeting, Ad Group 作为匹配条件merge到bulk文件的dataframe中
merged_df = pd.merge(bulk_df, summary_df_us, on=['Campaign', 'Keyword or Product Targeting', 'Ad Group'], how='left')

# 输出merge后的dataframe为xlsx文件，目录为D盘/运营/2生成过程表，名称为美国广告测试.xlsx
output_file = "D:/运营/2生成过程表/美国广告测试.xlsx"
merged_df.to_excel(output_file, index=False)
