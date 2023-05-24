import pandas as pd
import numpy as np
import sqlite3
from datetime import datetime
from datetime import timedelta
def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    df['日期'] = pd.to_datetime(df['日期'])
    df['周数'] = ((last_saturday - df['日期']).dt.days // 7) + 1
    return df

# 通过键盘输入选取国家
selected_country = input("Please input the Country: ")
 
#
conn = sqlite3.connect('D:\运营\sqlite\AmazonData.db')

# 从Bulkfiles表中读取数据
bulk_summary = pd.read_sql_query("SELECT * FROM Bulkfiles", conn)

# 关闭连接
conn.close()


bulk_summary=update_week_numbers(bulk_summary)

#选择100周以内的数据

bulk_summary=bulk_summary[(bulk_summary['周数']<=53)&(bulk_summary['Country'] == selected_country)]
bulk_summary['Spend'] = bulk_summary['Spend'].astype(float)
spend_summary =bulk_summary.groupby(["Country","Campaign", "SKU"]).agg({"Spend": "sum"}).reset_index()

 
# 为每个 "Campaign" 和 "Country" 找到具有最大 "Spend" 的 SKU

spend_summary = spend_summary.loc[spend_summary.groupby(["Country","Campaign"])["Spend"].idxmax()]


# 将结果重命名为 "主要SKU"
spend_summary = spend_summary.rename(columns={"SKU": "主要SKU"})
print(spend_summary)


# 将结果合并到周BUlkFIle原始数据表，创建一个新列 "主要SKU"
bulk_summary = bulk_summary.merge(spend_summary[["Country","Campaign", "主要SKU"]], on=["Country","Campaign"], how="left")
bulk_summary=bulk_summary[bulk_summary["Spend"]>0].fillna(0)
print(bulk_summary)

# 选取特定国家和主要SKU的数据
selected_data = bulk_summary


# 创建Campaign花费表


campaign_spend = pd.pivot_table(selected_data, values='Spend', index=['周数', '主要SKU'], columns='Campaign', aggfunc=np.sum)
campaign_spend = campaign_spend.reset_index()
campaign_spend=campaign_spend.fillna(0)
campaign_spend=campaign_spend.rename(columns={'主要SKU':'SKU'})

campaign_spend.to_excel(r'D:\运营\campaign_spend.xlsx')
print("campaign_spend.columns",campaign_spend)

# 创建产品销售表
# 读取产品销售数据 "D:\运营\2019plan\周销售数据.xlsx"
product_sales = pd.read_excel(r'D:\运营\2019plan\周销售数据.xlsx', sheet_name=0)
product_sales['Country']="US"
#周数列转换为int类型
product_sales['周数'] = product_sales['周数'].astype(int)
 
product_sales.rename(columns={'Units Ordered': 'Sales'}, inplace=True)
product_sales = product_sales[['Country', '周数', 'SKU', 'Sales']]

# 匹配主要SKU与销售表中的SKU，并合并两个表

merged_data = pd.merge(campaign_spend, product_sales, on=['周数','SKU'], how='left')
merged_data['Sales']=merged_data['Sales'].fillna(0)
 
merged_data.to_excel(r'D:\运营\merged_data.xlsx')

print(merged_data)

skus = merged_data['SKU'].unique()

# 创建一个空的 DataFrame 来存储结果
results = pd.DataFrame(columns=['SKU', 'Campaign', 'Correlation'])
q=0
# 对于每个 SKU
for sku in skus:
    # 从数据中筛选出该 SKU 的行
    q=q+1
    sku_data = merged_data[merged_data['SKU'] == sku].copy()
    sku_data['Sales'] = sku_data['Sales'].fillna(0).astype(float)
    print(sku_data)
    sku_data.to_excel(r'D:\\运营\\2生成过程表\\'+str(sku)+".xlsx")
    # 对于该 SKU 下的每个 Campaign
    for campaign in sku_data.columns:
         
        if campaign not in ['SKU', '周数','Country']:
            # 计算该 SKU 和 Campaign 的相关系数
            sku_data[campaign]=sku_data[campaign].fillna(0).astype(float)
            corr = sku_data[campaign].corr(sku_data['Sales'])
            
            # 将结果添加到结果 DataFrame 中
            results = pd.concat([results, pd.DataFrame({'SKU': sku, 'Campaign': campaign, 'Correlation': corr}, index=[0])], ignore_index=True)
print(results)
results.to_excel(r'D:\运营\3数据分析结果\Camaign相关性分析.xlsx')
                 
