#获取最近一个周六的函数
def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday
#将df中的日期列转换为周数并添加周数列
def update_week_numbers(df):


    last_saturday = find_last_saturday()
    print(last_saturday)

    # 检查输入 DataFrame 的列名中哪一个表示日期
    date_column = "日期" if "日期" in df.columns else "Date"

    df[date_column] = pd.to_datetime(df[date_column])
    df['周数'] = ((last_saturday - df[date_column]).dt.days // 7) + 1
    return df


#将Bulkfiles的数据加入主要SKU列
def process_spend_summary(BulkFile_df):
    # 检查是否有 "Country" 列
    has_country = "Country" in BulkFile_df.columns

    if not has_country:
        print("没有国家列")

    # 选择 SKU 列不为空的行
    BulkFile_df = BulkFile_df[BulkFile_df['SKU'].notna()]

    # 按 "Country"（如果有），"Campaign"，"Ad Group" 和 "SKU" 对 "Spend" 进行汇总
    group_columns = ["Country", "Campaign", "Ad Group", "SKU"] if has_country else ["Campaign", "Ad Group", "SKU"]
    spend_summary = BulkFile_df.groupby(group_columns).agg({"Spend": "sum"}).reset_index()

    # 为每个 "Country"（如果有），"Campaign" 和 "Ad Group" 找到具有最大 "Spend" 的 SKU
    group_columns = ["Country", "Campaign", "Ad Group"] if has_country else ["Campaign", "Ad Group"]
    spend_summary = spend_summary.loc[spend_summary.groupby(group_columns)["Spend"].idxmax()]

    # 将结果重命名为 "主要SKU"
    spend_summary = spend_summary.rename(columns={"SKU": "主要SKU"})

    # 将结果合并到周BUlkFIle原始数据表，创建一个新列 "主要SKU"
    merge_columns = ["Country", "Campaign", "Ad Group", "主要SKU"] if has_country else ["Campaign", "Ad Group", "主要SKU"]
    BulkFile_df = BulkFile_df.merge(spend_summary[merge_columns], on=group_columns, how="left")

    return BulkFile_df


#输出Campaign和主要SKU的对应表（如果有国家就包含国家）
 def process_spend_summary(pivot_df):
    # 检查是否有 "Country" 列/或者"COUNTRY"列
    

 
    #如果有COUNTRY列，将其改为Country
    if "COUNTRY" in pivot_df.columns:
        pivot_df.rename(columns={"COUNTRY": "Country"}, inplace=True)
        has_country = "Country" in pivot_df.columns



    if not has_country:
        print("没有国家列")

    # 选择 SKU 列不为空的行
    pivot_df = pivot_df[pivot_df['SKU'].notna()]

    # 按 "Country"（如果有），"Campaign"，"Ad Group" 和 "SKU" 对 "Spend" 进行汇总
    group_columns = ["Country", "Campaign", "Ad Group", "SKU"] if has_country else ["Campaign", "Ad Group", "SKU"]
    spend_summary = pivot_df.groupby(group_columns).agg({"Spend": "sum"}).reset_index()

    # 为每个 "Country"（如果有），"Campaign" 和 "Ad Group" 找到具有最大 "Spend" 的 SKU
    group_columns = ["Country", "Campaign", "Ad Group"] if has_country else ["Campaign", "Ad Group"]
    spend_summary = spend_summary.loc[spend_summary.groupby(group_columns)["Spend"].idxmax()]

    # 将结果重命名为 "主要SKU"
    spend_summary = spend_summary.rename(columns={"SKU": "主要SKU"})

    # 只保留 "Country"（如果有），"Campaign" 和 "主要SKU" 列
    output_columns = ["Country", "Campaign", "Ad Group","主要SKU"] if has_country else ["Campaign", "Ad Group","主要SKU"]
    spend_summary = spend_summary[output_columns]

    return spend_summary



import sqlite3
import pandas as pd
from datetime import datetime, timedelta

def update_bulkfiles_week_numbers(database_path):
    # Connect to the database and retrieve the Bulkfiles table
    conn = sqlite3.connect(database_path)
    df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)
    
    # Filter out rows without a date
    df = df[df['日期'].notna()]

    # Convert columns to appropriate types
    df['Spend'] = df['Spend'].astype(float)
    df["Max Bid"] = df["Max Bid"].astype(float)
    df['Sales'] = df['Sales'].astype(float)
    df['日期'] = pd.to_datetime(df['日期'])

    # Function to find the date of the last Saturday
    def find_last_saturday():
        today = datetime.now()
        last_saturday = today - timedelta(days=today.weekday() + 2)
        return last_saturday

    # Function to update the week number based on the date of the last Saturday
    def update_week_numbers(df):
 
    # 检查输入 DataFrame 的列名中哪一个表示日期
        date_column = "日期" if "日期" in df.columns else "Date"

        df[date_column] = pd.to_datetime(df[date_column])
        df['周数'] = ((last_saturday - df[date_column]).dt.days // 7) + 1
        return df


    # Update the week numbers in the dataframe
    updated_df = update_week_numbers(df)

    # Write the updated dataframe back to the database
    updated_df.to_sql('Bulkfiles', conn, if_exists='replace', index=False)

    # Close the database connection
    conn.close()

    print("Bulkfiles week numbers updated successfully.")
#update_bulkfiles_week_numbers('D:/运营/sqlite/AmazonData.db')
#import sqlite3
#import pandas as pd
#from datetime import datetime, timedelta

#计算Keyword Targeting和销量的相关系数

def calculate_correlations(sales_column, df, exclude_columns=[]):
    # 从 DataFrame 中删除不需要计算相关性的列
    cols_to_calculate_corr = [col for col in df.columns if col not in exclude_columns + [sales_column]]

    # 初始化结果 DataFrame
    result_df = pd.DataFrame(columns=list(exclude_columns) + ['计算相关性的列名', '相关系数'])

    # 分组计算相关性
    grouped_df = df.groupby(list(exclude_columns))

    for group, group_df in grouped_df:
        for col in cols_to_calculate_corr:
            # 计算相关系数
            correlation = group_df[sales_column].corr(group_df[col])

            # 将相关系数添加到结果 DataFrame
            if not isinstance(group, tuple):
                group = (group,)
            temp_data = {key: value for key, value in zip(list(exclude_columns), group)}
            temp_data.update({'计算相关性的列名': col, '相关系数': correlation})
            temp_df = pd.DataFrame(temp_data, index=[0])
            
            result_df = pd.concat([result_df, temp_df], ignore_index=True)

    return result_df



#据说速度会快一些的方法计算相关系数。
def calculate_correlations(df):
    # 按 Country, Campaign, Ad Group, Keyword or Product Targeting, Match Type, 主要SKU 对数据分组
    group_columns = ['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting', 'Match Type', '主要SKU']
    grouped_df = df.groupby(group_columns)

    # 从 DataFrame 中删除不需要计算相关性的列
    exclude_columns = group_columns + ['周数']
    cols_to_calculate_corr = [col for col in df.columns if col not in exclude_columns + ('Spend', 'Unit Ordered')]

    # 初始化结果 DataFrame
    result_df = pd.DataFrame(columns=group_columns + ['计算相关性的列名', '相关系数'])

    for group, group_df in grouped_df:
        # 将需要计算相关系数的列转换为数组
        data = group_df[cols_to_calculate_corr].values.T

        # 计算相关系数矩阵
        corr_matrix = np.corrcoef(data)

        # 提取所需的系数
        for i, col in enumerate(cols_to_calculate_corr):
            for j, other_col in enumerate(cols_to_calculate_corr[i+1:]):
                correlation = corr_matrix[i, j+i+1]

                # 将相关系数添加到结果 DataFrame
                temp_data = {key: value for key, value in zip(group_columns, group)}
                temp_data.update({'计算相关性的列名': f'{col} - {other_col}', '相关系数': correlation})
                temp_df = pd.DataFrame(temp_data, index=[0])
                
                result_df = pd.concat([result_df, temp_df], ignore_index=True)

    return result_df[['Country', '主要SKU', 'Campaign', 'Keyword or Product Targeting', 'Ad Group',
                      'Match Type', '计算相关性的列名', '相关系数']]

#一种把排除列作为输入的方法，计算相关系数。
def calculate_correlations(sales_column, df, exclude_columns=[]):
    # 从 DataFrame 中删除不需要计算相关性的列
    cols_to_calculate_corr = [col for col in df.columns if col not in exclude_columns + [sales_column]]

    # 初始化结果 DataFrame
    result_df = pd.DataFrame(columns=list(exclude_columns) + ['计算相关性的列名', '相关系数'])

    # 分组计算相关性
    grouped_df = df.groupby(list(exclude_columns))

    for group, group_df in grouped_df:
        for col in cols_to_calculate_corr:
            # 计算相关系数
            correlation = group_df[sales_column].corr(group_df[col])

            # 将相关系数添加到结果 DataFrame
            if not isinstance(group, tuple):
                group = (group,)
            temp_data = {key: value for key, value in zip(list(exclude_columns), group)}
            temp_data.update({'计算相关性的列名': col, '相关系数': correlation})
            temp_df = pd.DataFrame(temp_data, index=[0])
            
            result_df = pd.concat([result_df, temp_df], ignore_index=True)

    return result_df

#以下是chatgpt给的第1种计算相关系数的函数。
def calculate_correlations(df):
    # 按 Country, Campaign, Ad Group, Keyword or Product Targeting, Match Type, 主要SKU 对数据分组
    group_columns = ['Country', 'Campaign', 'Ad Group', 'Keyword or Product Targeting', 'Match Type', '主要SKU']
    grouped_df = df.groupby(group_columns)

    # 初始化结果 DataFrame
    result_df = pd.DataFrame(columns=['Country', '主要SKU', 'Campaign', 'Keyword or Product Targeting',
                                      'Ad Group', 'Match Type', 'Unit Ordered', 'Spend', '相关系数'])

    # 遍历每个分组，计算相关性并将结果添加到结果 DataFrame 中
    for group, group_df in grouped_df:
        # 只计算存在 Spend 和 Unit Ordered 列的分组
        if 'Spend' in group_df.columns and 'Unit Ordered' in group_df.columns:
            # 计算相关系数
            correlation = group_df['Spend'].corr(group_df['Unit Ordered'])

            # 将相关系数添加到结果 DataFrame
            temp_data = {key: value for key, value in zip(group_columns, group)}
            temp_data.update({'Unit Ordered': group_df['Unit Ordered'].iloc[0],
                               'Spend': group_df['Spend'].iloc[0],
                               '相关系数': correlation})
            temp_df = pd.DataFrame(temp_data, index=[0])
            result_df = pd.concat([result_df, temp_df], ignore_index=True)

    return result_df[['Country', '主要SKU', 'Campaign', 'Keyword or Product Targeting', 'Ad Group',
                      'Match Type', 'Unit Ordered', 'Spend', '相关系数']]