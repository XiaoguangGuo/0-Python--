

import pandas as pd
import datetime
import requests



def fetch_exchange_rates(app_id, currencies):
    api_url = 'https://openexchangerates.org/api/latest.json?app_id={}&symbols={}'
    rates = {}

    try:
        response = requests.get(api_url.format(app_id, ','.join(currencies)), headers={'Authorization': f'Token {app_id}'})
        response.raise_for_status()
        data = response.json()
        for currency in currencies:
            rate = data['rates'][currency]
            rates[currency] = rate
        print('连接API成功！以下是货币对美元的汇率：')
        print(rates)
    except requests.exceptions.HTTPError as error:
        print(f'连接API出错：{error}')
        
    return rates



if __name__ == '__main__':
    app_id = '438a43ad7170441aa0c7a00caebf086f'
    currencies = ['USD', 'CAD', 'EUR', 'GBP', 'JPY', 'MXN', 'SEK']
    exchange_rates = fetch_exchange_rates(app_id, currencies)
    exchange_rates = pd.to_numeric(exchange_rates, errors='coerce')

exchangerate_dic={"GV-US":"USD","GV-CA":"CAD","NEW-UK":"GBP","NEW-JP":"JPY","NEW-CA":"CAD","NEW-IT":"EUR","NEW-DE":"EUR","NEW-ES":"EUR","NEW-FR":"EUR","NEW-US":"USD","HM-US":"USD","GV-MX":"MXN","NEW-MX":"MXN"}
country_dic={"英国":"NEW-UK","日本":"NEW-JP","加拿大": "NEW-CA","意大利":"NEW-IT","德国":"NEW-DE" ,"西班牙":"NEW-ES","法国": "NEW-FR","美国": "NEW-US" }

def calculate_sales_summary(file_path):

    df = pd.read_excel(file_path)
    df['站点'] = df['站点'].map(country_dic)
    # 计算周数
    today = datetime.datetime.now()
    df['周数'] = (today - df['日期']).dt.days // 7+1
    exchange_rates = pd.Series(fetch_exchange_\rates(app_id, currencies))
    df['汇率'] = df['站点'].map(exchangerate_dic).map(exchange_rates)
    df['销售额']=df['销售额'].fillna(0)
    df['广告花费']=df['广告花费'].fillna(0)
    df['销售额'].astype(float)
    df['广告花费'].astype(float)
    df['销售额']=df['销售额']/df['汇率']
    df['广告花费']=df['广告花费']/df['汇率']
    
    # 按Country和Week进行分组，并计算销售额、产品销售数、订单数和广告总额的和
    grouped = df.groupby(['站点', '周数']).agg({'销量': 'sum', '销售额': 'sum', '广告花费': 'sum', '广告订单量': 'sum'})
    
    # 计算毛利A并添加到结果中
    grouped['毛利A'] = (grouped['销售额'] - grouped['广告花费']) / grouped['销售额']
    
    # 计算毛利B并添加到结果中
    grouped['毛利B'] = (grouped['销售额'] - grouped['广告花费']) / grouped['销量']
    
    # 计算销售额占比并添加到结果中
    grouped['销售额占比'] = grouped['销售额'] / grouped['销售额'].sum()
    
    # 计算销售额占比并添加到结果中
    grouped['广告花费占比'] = grouped['广告花费'] / grouped['广告花费'].sum()
    
    # 计算销售额占比并添加到结果中
    grouped['销售额占比'] = grouped['销售额'] / grouped['销售额'].sum()
    
    # 计算销售额占比并添加到结果中
    grouped['广告花费占比'] = grouped['广告
    df['汇率'] = df['站点'].map(exchangerate_dic).map(exchange_rates)
    df['销售额']=df['销售额'].fillna(0)
    df['广告花费']=df['广告花费'].fillna(0)
    df['销售额'].astype(float)
    df['广告花费'].astype(float)
    df['销售额']=df['销售额']/df['汇率']
    df['广告花费']=df['广告花费']/df['汇率']
    
    # 按Country和Week进行分组，并计算销售额、产品销售数、订单数和广告总额的和
    grouped = df.groupby(['站点', '周数']).agg({'销量': 'sum', '销售额': 'sum', '广告花费': 'sum', '广告订单量': 'sum'})
    
    # 计算毛利A并添加到结果中
    grouped['毛利A'] = (grouped['销售额'] - grouped['广告花费']) / grouped['销量']

    pivot_table = pd.pivot_table(grouped, index=['站点'], columns=['周数'], values=['销售额', '销量', '广告花费', '毛利A'])
    
    # 设置 Style，并输出到 Excel
    style = pivot_table.style.applymap(lambda x: 'background-color: #c8f7c5', subset=pd.IndexSlice[:, ['销售额']])
    style = style.applymap(lambda x: 'background-color: #d5d5ff', subset=pd.IndexSlice[:, ['销量']])
    style = style.applymap(lambda x: 'background-color: #ffccc7', subset=pd.IndexSlice[:, ['广告花费']])
    style = style.applymap(lambda x: 'background-color: #fefebf', subset=pd.IndexSlice[:, ['毛利A']])
    
    # 在 Excel 中输出时，将 Style 一起写入到文件中
    with pd.ExcelWriter(r'D:\运营\\3数据分析结果\\SailingstarTotalWeek.xlsx') as writer:
        pivot_table.to_excel(writer, sheet_name='Sheet1', index=True, header=True)
        style.to_excel(writer, sheet_name='Sheet1', index=True, header=True, startrow=len(pivot_table)+1)
    
    return pivot_table

file_path=r'D:\运营\\2生成过程表\\All_Product_Analyzefile.xlsx'
calculate_sales_summary(file_path)
 

