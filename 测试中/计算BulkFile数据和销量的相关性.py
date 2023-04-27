result_df = pd.DataFrame(columns=['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting', 'Match Type', '主要SKU', '计算相关性的列名', '相关系数'])

# 分组计算相关性
grouped_df = df.groupby(['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting', 'Match Type', '主要SKU'])

for group, group_df in grouped_df:
    for col in cols_to_calculate_corr:
        # 计算相关系数
        correlation = group_df['Spend'].corr(group_df[col])

        # 将相关系数添加到结果 DataFrame
        temp_data = {'Country': group[0], 'Campaign': group[1], 'Ad Group': group[2], 'Keyword or Product Targeting': group[3], 'Match Type': group[4], '主要SKU': group[5], '计算相关性的列名': col, '相关系数': correlation}
        temp_df = pd.DataFrame(temp_data, index=[0])
        
        result_df = pd.concat([result_df, temp_df], ignore_index=True)