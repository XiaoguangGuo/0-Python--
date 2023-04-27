import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import numpy as np

def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    print(last_saturday)
    # 检查输入 DataFrame 的列名中哪一个表示日期
    date_column = "日期" if "日期" in df.columns else "Date"

    df[date_column] = pd.to_datetime(df[date_column])
    df['周数'] = ((last_saturday - df[date_column]).dt.days // 7) + 1
    return df

#读取"D:\运营\sqlite\AmazonData.db"中的Bulkfile表作为pivot	
conn = sqlite3.connect("D:\运营\sqlite\AmazonData.db")
#读取Bulkfiles表中距离last_saturday27周以内的数据成dataframe


last_saturday = find_last_saturday()

# 计算 27 周以前的日期
start_date = last_saturday - timedelta(weeks=27)

# 修改 SQL 查询以获取 27 周以后的数据
query = f'SELECT * FROM "Bulkfiles" WHERE "日期" >= \'{start_date}\''

# 使用修改后的查询获取数据
pivot_df = pd.read_sql_query(query, conn)

print(pivot_df)
 
conn.close()

#将df中的日期列转换为周数并添加周数列


pivot_df = update_week_numbers(pivot_df)
#输出到excel
print(pivot_df)
 


def process_spend_summary(pivot_df):

    # 查找是否有 "Country" 或 "COUNTRY" 列
    has_country = False

    # 如果有 "COUNTRY" 列，将其改为 "Country"
    if "COUNTRY" in pivot_df.columns:
        pivot_df.rename(columns={"COUNTRY": "Country"}, inplace=True)
        has_country = "Country" in pivot_df.columns
    elif "Country" in pivot_df.columns:
        has_country = True

    if not has_country:
        print("没有国家列")




    pivot_df = pivot_df[pivot_df['SKU'].notna()]
    # 按 "Country"（如果有），"Campaign"，"Ad Group" 和 "SKU" 对 "Spend" 进行汇总
    pivot_df["Spend"] = pd.to_numeric(pivot_df["Spend"], errors="coerce")
    pivot_df["Spend"].fillna(0, inplace=True)
    group_columns = ["Country", "Campaign", "Ad Group", "SKU"] if has_country else ["Campaign", "Ad Group", "SKU"]
    spend_summary = pivot_df.groupby(group_columns).agg({"Spend": "sum"}).reset_index()

    # 为每个 "Country"（如果有），"Campaign" 和 "Ad Group" 找到具有最大 "Spend" 的 SKU
    group_columns2 = ["Country", "Campaign", "Ad Group"] if has_country else ["Campaign", "Ad Group"]
    spend_summary = spend_summary.loc[spend_summary.groupby(group_columns2)["Spend"].idxmax()]

    # 将结果重命名为 "主要SKU"
    spend_summary = spend_summary.rename(columns={"SKU": "主要SKU"})

    # 只保留 "Country"（如果有），"Campaign" 和 "主要SKU" 列
    output_columns = ["Country", "Campaign", "Ad Group","主要SKU"] if has_country else ["Campaign", "Ad Group","主要SKU"]
    spend_summary = spend_summary[output_columns]

    return spend_summary


# 读取搜索报告"D:\运营\2生成过程表\Sponsored Products Search term report.xlsx"
search_report = pd.read_excel(r'D:\运营\2生成过程表\Sponsored Products Search term report.xlsx')
#获取距lastsaturday27周内的数据，lastsaturday为最近一个周六，使用本程序的函数计算
search_report = search_report[search_report['Date'] > find_last_saturday() - timedelta(weeks=27)]
#search_report的"COUNTRY"列改为"Country"
search_report.rename(columns={"COUNTRY": "Country"}, inplace=True)

#按Date距离last Saturday的日期重新计算周数并更新周数列

spend_summary=process_spend_summary(pivot_df)
#输出到"D:\运营\2生成过程表\spend_summary.xlsx"
spend_summary.to_excel(r'D:\运营\2生成过程表\spend_summary.xlsx', index=False)
 
# 将结果合并到搜索表周原始数据表，创建一个新列 "主要SKU"
merged_data = pd.merge(search_report, spend_summary, left_on=["Country", "Campaign Name", "Ad Group Name"],right_on=["Country","Campaign","Ad Group"], how="left")
merged_data = merged_data.drop(columns=["Campaign", "Ad Group"])
#输出到"D:\运营\2生成过程表\merged_data.xlsx"
 



#读取周销售数据表 在 运营 2019计划中
sales_df = pd.read_excel(r'D:\运营\2019plan\周销售数据.xlsx')
last_saturday = find_last_saturday()
#增加一列周数，值为周数
sales_df["周数"] = ((last_saturday - sales_df["日期"]).dt.days // 7) + 1
#增加一列Country，值为GV-US
sales_df["Country"] = "GV-US"
#将sales_df中的数据按国家和SKU和周数为索引，sum of sales为值，创建一个新的dataframe
sales_df = sales_df.groupby(["Country", "SKU", "周数"])["Units Ordered"].sum().reset_index()
#将sales_df中的SKU列改为主要SKU
sales_df.rename(columns={"SKU": "主要SKU"}, inplace=True)
print(sales_df)

#键盘输入国家名和主要SKU名，输出这个国家这个主要SKU的相关性
country = input('请输入国家名：')
sku = input('请输入主要SKU：')
#筛选出这个国家这个主要SKU的数据
merged_dataforpivot = merged_data[(merged_data['Country'] == country) & (merged_data['主要SKU'] == sku)]
#筛选出sales_df中这个国家这个主要SKU的数据
sales_df = sales_df[(sales_df['Country'] == country) & (sales_df['主要SKU'] == sku)]
#输出sales_df到"D:\运营\2生成过程表\sales_df.xlsx"GV
sales_df.to_excel(r'D:\运营\2生成过程表\sales_df.xlsx')
 
#用sku，campaign，ad group，周数作为索引，customersearchterm作为列，sum of spend作为值，创建一个新的dataframe
pivot_df = merged_dataforpivot.pivot_table(index=["Country", "Campaign Name", "Ad Group Name", "主要SKU", "周数"], columns="Customer Search Term", values="Spend", aggfunc="sum").reset_index()
#输出到excel

#将sales_df合并到pivot_df中，按国家，SKU，周数为索引，sum of sales为值，创建一个新的dataframe
pivot_df = pd.merge(pivot_df, sales_df, on=["Country", "主要SKU", "周数"], how="left")

pivot_df = pivot_df[pivot_df["主要SKU"].notna()]
pivot_df.to_excel(r'D:\运营\2生成过程表\pivot_df.xlsx')
input('请按回车键继续：')

def calculate_correlations(df):
    # 列出需要排除的列名
    exclude_columns = ['Country', 'Campaign Name', 'Ad Group Name', '主要SKU', '周数', 'Units Ordered']

    # 找出search term列名
    search_term_cols = list(set(df.columns) - set(exclude_columns))

    # 按照分组计算相关系数
    grouped_df = df.groupby(['Country', 'Campaign Name',  'Ad Group Name', '主要SKU'])

    # 初始化结果 DataFrame
    result_df = pd.DataFrame(columns=['Country', 'Campaign Name', 'Ad Group Name', '主要SKU', 'search term 列名', '相关系数'])

    for group, group_df in grouped_df:
        # 将需要计算相关系数的列转换为数组
        data = group_df[search_term_cols].values.T

        # 找到不为0的列
        nonzero_cols = np.array(search_term_cols)[np.sum(group_df[search_term_cols], axis=0) != 0]
        
        # 计算相关系数矩阵
        corr_matrix = np.corrcoef(data[:, np.isin(search_term_cols, nonzero_cols)])

        # 提取所需的系数
        for i, col in enumerate(nonzero_cols):
            for j, other_col in enumerate(nonzero_cols[i+1:]):
                correlation = corr_matrix[i, j+i+1]

                # 将相关系数添加到结果 DataFrame
                temp_data = {key: value for key, value in zip(['Country', 'Campaign', 'Name', 'Ad Group Name', '主要SKU'], group)}
                temp_data.update({'search term 列名': f'{col} - {other_col}', '相关系数': correlation})
                temp_df = pd.DataFrame(temp_data, index=[0])

                result_df = pd.concat([result_df, temp_df], ignore_index=True)

    return result_df


# 调用函数并输出结果
pivot_df.columns = pivot_df.columns.astype(str)
result_df = calculate_correlations(pivot_df)
print(result_df)
#输出结果到excel
result_df.to_excel(r'D:\运营\2生成过程表\关键词相关性.xlsx')

   



