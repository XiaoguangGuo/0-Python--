import pandas as pd
import datetime

def calculate_sales_summary(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    # 计算周数
    today = datetime.datetime.now()
    df['周数'] = (today - df['日期']).dt.days // 7+1
    
    # 按Country和Week进行分组，并计算销售额、产品销售数、订单数和广告总额的和
    grouped = df.groupby(['站点', '周数']).agg({'销量': 'sum', '销售额': 'sum', '广告花费': 'sum', '广告订单量': 'sum'})
    
    # 计算毛利A并添加到结果中
    grouped['毛利A'] = (grouped['销售额'] - grouped['广告花费']) / grouped['销量']

    pivot_table = pd.pivot_table(grouped, index=['站点'], columns=['周数'], values=['销售额', '销量', '广告花费', '毛利A'])
    return pivot_table

    
    # 返回汇总结果
    return pivot_table


file_path=r'D:\运营\\2生成过程表\\All_Product_Analyzefile.xlsx'
pivot_table_Total=calculate_sales_summary(file_path)
print(pivot_table_Total)
pivot_table_Total.to_excel(r'D:\运营\\3数据分析结果\\SailingstarTotalWeek.xlsx')
