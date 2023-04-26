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
