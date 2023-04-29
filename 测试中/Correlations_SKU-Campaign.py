# -*- coding: utf-8 -*-
import pandas as pd
import sqlite3
from datetime import datetime, timedelta



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


def correlations_SKU_Campaign(df):
    # 按 Country, Campaign, Ad Group, Keyword or Product Targeting, Match Type, 主要SKU 对数据分组
    group_columns = ['Country', 'Campaign', 'SKU']
    grouped_df = df.groupby(group_columns)

    # 初始化结果 DataFrame
    result_df = pd.DataFrame(columns=['Country', 'SKU', 'Campaign',  'Spend', 'Unit Ordered', '相关系数'])

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

    return result_df[['Country', '主要SKU', 'Campaign',  'Spend', 'Unit Ordered','相关系数']]



def update_bulkfiles_week_numbers(database_path):
    input_weeks = input("请输入考察周数，直接回车为默认52周：")
    if input_weeks == '':
        input_weeks = 52
    else:
        input_weeks = int(input_weeks)
    # Connect to the database and retrieve the Bulkfiles table
    conn = sqlite3.connect(database_path)
    #读取database_path中Bulkfiles表，获取距离last_saturday 的日期在input_weeks周以内的数据，将其转换为df
    
    df = pd.read_sql_query(f"SELECT * FROM Bulkfiles WHERE 日期 >= date('now', '-{input_weeks} week')", conn)


    
    # Filter out rows without a date
    df = df[df['日期'].notna()]

    # Convert columns to appropriate types
    df['Spend'] = df['Spend'].astype(float)
    df["Max Bid"] = df["Max Bid"].astype(float)
    df['Sales'] = df['Sales'].astype(float)
    df['日期'] = pd.to_datetime(df['日期'])

    # Function to find the date of the last Saturday
    
    last_saturday=find_last_saturday()

    # Function to update the week number based on the date of the last Saturday
    
    updated_df = update_week_numbers(df)

    # Write the updated dataframe back to the database
 
    conn.close()

    print("Bulkfiles week numbers updated successfully.")
    return updated_df

 
Bulkfiles_DF=update_bulkfiles_week_numbers(r'D:\运营\sqlite\AmazonData.db')
#筛选Bulkfiles_DF中SKU列不为空的行，列名为：['Country', 'SKU', 'Campaign', 'Spend']
Bulkfiles_DF=Bulkfiles_DF.loc[Bulkfiles_DF['SKU'].notnull(),['Country', 'SKU', 'Campaign', 'Spend']]


# retrieve the Sales table"D:\运营\2生成过程表\周销售数据总表.xlsx"
Sales_DF=pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx',sheet_name='Sheet1')
#计算周数，更新周数列，算法为日期值距离last_saturday的周数
Sales_DF=update_week_numbers(Sales_DF)
#筛选Sales_DF中SKU列不为空的行，列名为：['Country', 'SKU','Units Ordered']
Sales_DF=Sales_DF.loc[Sales_DF['SKU'].notnull(),['Country', 'SKU','Units Ordered']]
#将Sales_DFmerge到Bulkfiles_DF中，列名为：['Country', 'SKU', 'Campaign', 'Spend', 'Units Ordered']
Bulkfiles_DF=pd.merge(Bulkfiles_DF,Sales_DF,how='left',on=['Country', 'SKU'])
#计算相关系数
corr_SKU_Campaign_DF=correlations_SKU_Campaign(Bulkfiles_DF)
#将Correlations_DF输出至excel
corr_SKU_Campaign_DF.to_excel(r'D:\运营\2生成过程表\correlations_SKU_Campaign.xlsx',index=False)


 

 


