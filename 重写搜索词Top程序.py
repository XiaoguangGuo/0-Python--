import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import re
import os

# 变量设置
clicktemp = 9
weekscope = 105
topnumber = 20
conversationrate_set = 0.03

# 获取数据库
conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')


def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday


def update_week_numbers(df):
    last_saturday = find_last_saturday()
    df['日期'] = pd.to_datetime(df['日期'])
    df['周数'] = ((last_saturday - df['日期']).dt.days // 7) + 1
    return df


def preprocess_dataframe(df):
    df = df[df['日期'].notna()]
    df = df.drop_duplicates()
    df['Spend'] = df['Spend'].astype(float)
    df["Max Bid"] = df["Max Bid"].astype(float)
    df['Sales'] = df['Sales'].astype(float)
    df['日期'] = pd.to_datetime(df['日期'])
    return df


def calculate_conversion_rate(df):
    df['转化率'] = df['Orders'] / df['Clicks']
    df['点击率'] = df['Clicks'] / df['Impressions']
    return df


def load_data():
    last_saturday = find_last_saturday()
    weeks_ago_27 = last_saturday - timedelta(weeks=weekscope)
    df_SummaryCountry = pd.read_sql_query(f'SELECT * FROM "Bulkfiles" WHERE 日期 >= "{weeks_ago_27}"', conn)
    df_SummaryCountry = preprocess_dataframe(df_SummaryCountry)
    updated_df = update_week_numbers(df_SummaryCountry)
    updated_df = updated_df[updated_df["周数"] < weekscope]
    return updated_df


def pivot_data(df):
    df = df.drop(["Campaign Status", "Ad Group Status", "Status"], axis=1)
    pivot_df = df.groupby(["Country", "Campaign", "Ad Group", "Keyword or Product Targeting",
                           "Match Type"]).agg({
                               "Impressions": 'sum',
                               'Clicks': 'sum',
                               'Spend': 'sum',
                               'Orders': 'sum',
                               "Total Units": 'sum',
                               'Sales': 'sum'
                           }).reset_index()

    pivot_df = calculate_conversion_rate(pivot_df)
    return pivot_df


def classify_data(df):
    df['标签'] = '无'
    df.loc[(df['Clicks'] >= 10) & (df['转化率'] >= 0.2), "标签"] = '好targeting'
    df.loc[(df['Clicks'] >= 10) & ((df['转化率'] >= 0.1) & (df['转化率'] < 0.2)), "标签"] = '可用Targeting'
    df.loc[(df['Clicks'] >= 20) & ((df['转化率'] >= 0.05) & (df['转化率'] < 0.1)), "标签"] = '差Targeting-挑选'
    df.loc[(df['Clicks'] >= 20) & (df['转化率'] < 0.05), '标签'] = '差




def filter_condition(pivot_df):
    return (pivot_df['Clicks'] >= 20) & (pivot_df['转化率'] < 0.05)

# 更新符合条件的行
def update_filtered_rows(pivot_df):
    pivot_df.loc[filter_condition(pivot_df), '标签'] = '差Targeting-淘汰'
    return pivot_df

# 读取文件
filepath = 'D:\\运营\\1数据源\\周Bulk广告数据\\'
filename = '20210926-20211002.xlsx'
updated_df = pd.read_excel(filepath + filename)

# 筛选条件
clicktemp = 20
topnumber = 5
conversationrate_set = 0.03

# 应用筛选条件和更新操作
pivot_df = update_filtered_rows(updated_df)

# ... (保留原始代码)

# 合并表格时使用筛选条件
amazon_bulk_df_merge_keywords.loc[filter_condition(amazon_bulk_df_merge_keywords), 'Status'] = 'paused'
amazon_bulk_df_merge_keywords.loc[filter_condition(amazon_bulk_df_merge_keywords), '变更记录'] = '暂停所有其他非选词'

# ... (保留原始代码)

# 将结果保存到新的Excel文件时，应用筛选条件
merged_df = merged_df[filter_condition(merged_df)]