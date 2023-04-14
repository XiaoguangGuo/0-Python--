import os
import pandas as pd
import glob

source_folder = 'D:\\运营\\1数据源\\待修改Bulk文件'
output_folder = 'D:\\运营\\1数据源\\处理后待修改Bulk文件'

# 1. 读取 source_folder 中的所有文件
file_list = glob.glob(os.path.join(source_folder, '*.xlsx'))






# 2. 读取 output_summary.xlsx 文件
summary_df = pd.read_excel('D:\\运营\\output_summary.xlsx')

# 保留 summary_df 中需要的列
summary_df = summary_df[['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting', 'Match Type', 'Impressions', 'Clicks', 'Spend', 'Orders', 'Total Units', 'Sales', '转化率', '点击率', '标签','Campaign Status_merged','主要SKU']]




for file in file_list:
    # 读取待修改的文件
    source_df = pd.read_excel(file, sheet_name=1)

    # 为 source_df 添加 Country 列
    country_name = os.path.basename(file).split('_')[0]
    source_df['Country'] = country_name




    #加状态标签
    # 获取 "Record Type" 列值为 "Campaign" 的 [Campaign] 列值对应的 "Campaign Status" 列的值
    campaign_statuses = source_df.loc[source_df['Record Type'] == 'Campaign', ['Campaign', 'Campaign Status']]

    # 去重，以确保每个 Campaign 只有一个对应的 Campaign Status
    campaign_statuses = campaign_statuses.drop_duplicates(subset='Campaign')

    # 将结果合并到 [Campaign] 列具有相同 "Campaign" 值的所有行
    source_df = source_df.merge(campaign_statuses, on='Campaign', suffixes=('', '_merged'))



    # 3. 根据条件合并两个 DataFrame
    merged_df = source_df.merge(summary_df, on=['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting', 'Match Type'], how='left')
 

    # 将合并后的文件保存到 output_folder 中
    output_file = os.path.join(output_folder, 'Revised_' + os.path.basename(file))
    merged_df.to_excel(output_file, index=False)
