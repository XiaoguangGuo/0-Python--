import pandas as pd
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


#读取D:\运营\2生成过程表\All_Product_Analyzefile.xlsx,sheet1
Product_Analyzefile_df = pd.read_excel(r'D:\运营\2生成过程表\All_Product_Analyzefile.xlsx',sheet_name='sheet1')
all_sales_df=pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx',sheet_name="Sheet1")

ASIN_SKU = Product_Analyzefile_df[["站点", "ASIN", "MSKU"]].drop_duplicates()
ASIN_SKU_sales=all_sales_df[["Country","(Child) ASIN","SKU"]].drop_duplicates()
ASIN_SKU_sales.rename(columns={'(Child) ASIN':'ASIN','Country':'站点',"SKU":"MSKU"},inplace=True)
ASIN_SKU=pd.concat([ASIN_SKU,ASIN_SKU_sales],axis=0,ignore_index=True) 
ASIN_SKU=ASIN_SKU.drop_duplicates()

 

#按Country和Asin分组后MSKU列不同行的字符串合并，用逗号拼接。形成新的MSKU列，然后drop掉MSKU列，去重。
 

# 定义一个函数，用于合并和去重MSKU
def merge_and_deduplicate(mskus):
    unique_mskus = set()
    for msku in mskus:
        # 分割逗号分隔的字符串，并将结果添加到集合中
        
        if isinstance(msku, (float, int)):
            unique_mskus.update(str(msku).split(','))
        else:
            unique_mskus.update(msku.split(','))
    return ','.join(unique_mskus)

# 使用groupby和apply方法合并和去重MSKU
ASIN_SKU_unique = ASIN_SKU.groupby(['站点', 'ASIN'])['MSKU'].apply(merge_and_deduplicate).reset_index()
ASIN_SKU_unique.columns = ['站点', 'ASIN', 'MNSKU']

 

def merge_and_expand_msku(df):
    df['MSKU'] = df['MSKU'].apply(lambda x: set(str(x).split(',')) if isinstance(x, (float, int)) else set(x.split(',')))

    merged_df = df.groupby(['站点', 'ASIN'])['MSKU'].apply(lambda x: set.union(*x)).reset_index()
    expanded_df = merged_df.explode('MSKU').reset_index(drop=True)
    return expanded_df

ASIN_SKU_multiple = merge_and_expand_msku(ASIN_SKU)
 



Product_Analyzefile_df=update_week_numbers(Product_Analyzefile_df)
Product_Analyzefile_df["周数"]+=1
#筛选Product_Analyzefile_df中”站点"，"ASIN"，"FBA可售"，可售天数预估，建议补货量三列，另存为一个dataframe，FBA可售改为Fufillable
Product_Analyzefile_Fufillable_df=Product_Analyzefile_df.loc[Product_Analyzefile_df["周数"]==1,["站点","ASIN","FBA可售","可售天数预估","建议补货量"]]
Product_Analyzefile_Fufillable_df.rename(columns={'FBA可售':'Fufillable'},inplace=True)
 





all_sales_df=update_week_numbers(all_sales_df)

all_advertise_df=pd.read_excel(r'D:\运营\2生成过程表\周bulk数据Summary.xlsx',sheet_name="SKU-WEEK")


all_advertise_df=all_advertise_df[['Country','SKU','周数','Impressions','Clicks','Spend','Orders','Total Units','Sales']]
all_advertise_df=pd.merge(all_advertise_df,ASIN_SKU_multiple,left_on=["Country","SKU"],right_on=["站点","MSKU"],how="left")

all_advertise_df=all_advertise_df.groupby(["Country","ASIN","周数"])['Impressions','Clicks','Spend','Orders','Total Units','Sales'].sum().reset_index()
print(all_advertise_df)
#把all_advertise_df merge到all_sales_df，匹配 Country，SKU，week周数，all_advertise_df concat只取[Impressions,	Clicks	,Spend	,Orders	,Total Units,	Sales]这几列\
all_advertise_df=all_advertise_df[['Country', 'ASIN', '周数', 'Impressions', 'Clicks', 'Spend', 'Orders',
       'Total Units', 'Sales']]
#输出all_advertise_df

 
all_sales_df = pd.merge(all_sales_df, all_advertise_df, how='outer', left_on=['Country','ASIN','周数'], right_on=['Country','ASIN','周数'])







#all_salses_df rename
all_sales_df.rename(columns={'(Child) ASIN':'ASIN','Country':'站点','Sessions - Total':'Sessions','Units Ordered':'销量','Ordered Product Sales':'销售额','Clicks':'广告点击量','Total Units':'广告订单量','Sales':'广告销售额','Spend':'广告花费'},inplace=True)
all_sales_df["店铺"]=all_sales_df["站点"]
all_sales_df["MSKU"]=all_sales_df["SKU"]
#把all_sales_df concat到Product_Analyzefile_df

Product_Analyzefile_df = pd.concat([Product_Analyzefile_df,all_sales_df],axis=0,ignore_index=True) 
#获取ASIN列和MSKU的对照表









Product_Analyzefile_df=Product_Analyzefile_df[Product_Analyzefile_df["周数"]<=52]
exchangerate_20221217={"GV-US":1,"GV-CA":1.3701,"NEW-UK":0.8223,"NEW-JP":136.6790,"NEW-CA":1.3701,"NEW-IT":0.9457,"NEW-DE":0.9457,"NEW-ES":0.9457,"NEW-FR":0.9457,"NEW-US":1,"HM-US":1,"GV-MX":19.774,"NEW-MX":19.774}

columns_to_convert = ['销售额', '广告花费', '广告销售额', '毛利润']


    # Convert the column to a numeric data type
for column in columns_to_convert:
    # Fill NaN values with 0
    Product_Analyzefile_df[column] = Product_Analyzefile_df[column].fillna(0)

    # Convert the column to a numeric data type
    Product_Analyzefile_df[column] = pd.to_numeric(Product_Analyzefile_df[column], errors='coerce')

country_codes = Product_Analyzefile_df['站点'].unique()
zhandian_dic_country={"GV-US":"GV-US","GV-CA":"GV-CA","英国":"NEW-UK","日本":"NEW-JP","加拿大":"NEW-CA","意大利":"NEW-IT","德国":"NEW-DE","西班牙":"NEW-ES","法国":"NEW-FR","美国":"NEW-US","HM-US":"HM-US","GV-MX":"GV-MX","墨西哥":"NEW-MX"}
for country_code in country_codes:
    country_code_new=zhandian_dic_country[country_code]
    exchange_rate = exchangerate_20221217.get(country_code_new, 1)
    
    # 筛选出该国家的数据
    country_data = Product_Analyzefile_df[Product_Analyzefile_df['站点'] == country_code]

    # 更新每个要转换的列的值
    for column in columns_to_convert:
        Product_Analyzefile_df.loc[country_data.index, columns_to_convert] = country_data[columns_to_convert] / exchange_rate


# Create the pivot table
#输出Product_Analyzefile_df到D:\运营\2生成过程表\TESTAll_Product_Analyzefilebeforepivot.xlsx
Product_Analyzefile_df.to_excel(r'D:\运营\2生成过程表\TESTAll_Product_Analyzefilebeforepivot.xlsx')



# Create the pivot table
Product_Analyzefile_df_pivot = Product_Analyzefile_df.pivot_table(
    index=['ASIN', '店铺', '站点'],
    columns='周数',
    values=['Sessions', '销量', '销售额', '广告点击量','广告花费', '广告订单量', '广告销售额','毛利润'],
    aggfunc='sum'
)




# Generate column names with the original column names and the week number
new_columns = [(col[0] + str(col[1])) for col in Product_Analyzefile_df_pivot.columns]

# Reset the index and set new column names
Product_Analyzefile_df_pivot.reset_index(inplace=True)
Product_Analyzefile_df_pivot.columns = ['ASIN', '店铺', '站点'] + new_columns







#################################加入老站的库存和在途######################################
Stock_US=pd.read_excel(r'D:\运营\2019plan\当日Amazon库存.xlsx')
Stock_US.rename(columns = {'sku':"SKU",'asin':"ASIN","afn-fulfillable-quantity":"Fufillable","afn-inbound-receiving-quantity":"Receiving","afn-reserved-quantity":"Reserved"}, inplace = True)
Stock_US["站点"]="GV-US"
Stock_US=Stock_US[["站点","ASIN","SKU","Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]]

Stock_CA=pd.read_excel(r'D:\运营\2019plan\Canada当前Amazon库存.xlsx')
Stock_CA["站点"]="GV-CA"
Stock_CA.rename(columns = {'sku':"SKU",'asin':"ASIN","afn-fulfillable-quantity":"Fufillable","afn-inbound-receiving-quantity":"Receiving","afn-reserved-quantity":"Reserved"}, inplace = True)
Stock_CA=Stock_CA[["站点","ASIN","SKU","Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]]

Stock_MX=pd.read_excel(r'D:\运营\2019plan\Mexico当日Amazon库存.xlsx')
Stock_MX.rename(columns = {'sku':"SKU",'asin':"ASIN","afn-fulfillable-quantity":"Fufillable","afn-inbound-receiving-quantity":"Receiving","afn-reserved-quantity":"Reserved"}, inplace = True)
Stock_MX["站点"]="GV-MX"
Stock_MX=Stock_MX[["站点","ASIN","SKU","Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]]                

Stock_All=pd.concat([Stock_US,Stock_CA,Stock_MX])

#Stock_All groupy by Asin and Country，对["Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]进行sum，只显示["Asin","Country"，"Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"]reindex。
Stock_All=Stock_All.groupby(["ASIN","站点"])["Fufillable","Reserved","afn-inbound-working-quantity","afn-inbound-shipped-quantity","Receiving"].sum().reset_index()
           
#下面处理在途计划
Intransit_us=pd.read_excel(r'D:\运营\2019plan\在途库存.xlsx')
Intransit_us["站点"]="GV-US"
Intransit_us.rename(columns = {'Merchant SKU':"SKU"}, inplace = True)

Intransit_ca=pd.read_excel(r'D:\运营\2019plan\Canada在途库存.xlsx')
Intransit_ca["站点"]="GV-CA"
Intransit_ca.rename(columns = {'Merchant SKU':"SKU"}, inplace = True)
        
Intransit_mx=pd.read_excel(r'D:\运营\2019plan\Mexico在途库存.xlsx')
Intransit_mx["站点"]="GV-MX"              
Intransit_mx.rename(columns = {'Merchant SKU':"SKU"}, inplace = True) 

Intransit_All=pd.concat([Intransit_us,Intransit_ca,Intransit_mx])    
#对Intransit_All groupy by Asin and Country，对["Shipped"]进行sum，只显示["Asin","Country"，"Shipped"]reindex。
Intransit_All=Intransit_All.groupby(["ASIN","站点","周数"])["Shipped"].sum().reset_index()

max_week=11
Intransit_Weeks = Intransit_All[["站点","ASIN"]].drop_duplicates()
for i in range(1,max_week):
  Intransit_All2=Intransit_All.groupby(["站点","ASIN","周数"],as_index=False)[['Shipped']].agg('sum')
  Intransit_Weeks_i=Intransit_All2.loc[Intransit_All2["周数"]==i]
  if len(Intransit_Weeks_i)>0:    
      Intransit_Weeks_i=Intransit_Weeks_i[["站点","ASIN","Shipped"]]
      Intransit_Weeks_i.rename(columns = {"Shipped":"第"+str(i)+"周入库"}, inplace = True)

  Intransit_Weeks =pd.merge(Intransit_Weeks,Intransit_Weeks_i,on=["站点","ASIN"] ,how="left")
column_names_intransit = ["站点", "ASIN"] + [f"第{i}周入库" for i in range(1, 16)]
Intransit_Weeks_Columnsall = pd.DataFrame(columns=column_names_intransit)
Intransit_Weeks=pd.concat([Intransit_Weeks_Columnsall,Intransit_Weeks],axis=0,ignore_index=True)


Product_Analyzefile_Fufillable_df=pd.concat([Product_Analyzefile_Fufillable_df,Stock_All],axis=0,ignore_index=True)
Product_Analyzefile_df_pivot = pd.merge(Product_Analyzefile_df_pivot, Product_Analyzefile_Fufillable_df, how='outer', left_on=['站点', 'ASIN'], right_on=['站点', 'ASIN'])


Product_Analyzefile_df_pivot = pd.merge(Product_Analyzefile_df_pivot, Intransit_Weeks, how='outer', left_on=['站点', 'ASIN'], right_on=['站点', 'ASIN'])

Product_Analyzefile_df_pivot["STOCKALL"] = (
    Product_Analyzefile_df_pivot["Fufillable"].fillna(0)  +
    Product_Analyzefile_df_pivot["Receiving"].fillna(0)  +
    Product_Analyzefile_df_pivot["Reserved"].fillna(0)  +
    Product_Analyzefile_df_pivot["afn-inbound-shipped-quantity"].fillna(0))

#计算Product_Analyzefile_df_pivot10周总销量
Product_Analyzefile_df_pivot["Selling10"]=Product_Analyzefile_df_pivot["销量1"].fillna(0)+Product_Analyzefile_df_pivot["销量2"].fillna(0)+Product_Analyzefile_df_pivot["销量3"].fillna(0)+Product_Analyzefile_df_pivot["销量4"].fillna(0)+Product_Analyzefile_df_pivot["销量5"].fillna(0)+Product_Analyzefile_df_pivot["销量6"].fillna(0)+Product_Analyzefile_df_pivot["销量7"].fillna(0)+Product_Analyzefile_df_pivot["销量8"].fillna(0)+Product_Analyzefile_df_pivot["销量9"].fillna(0)+Product_Analyzefile_df_pivot["销量10"].fillna(0)
#计算4周总销量
Product_Analyzefile_df_pivot["Selling4"]=Product_Analyzefile_df_pivot["销量1"].fillna(0)+Product_Analyzefile_df_pivot["销量2"].fillna(0)+Product_Analyzefile_df_pivot["销量3"].fillna(0)+Product_Analyzefile_df_pivot["销量4"].fillna(0)
Product_Analyzefile_df_pivot["ZZ1"]=Product_Analyzefile_df_pivot["销量1"].fillna(0)-Product_Analyzefile_df_pivot["销量2"].fillna(0)
Product_Analyzefile_df_pivot["ZZ2"]=(Product_Analyzefile_df_pivot["销量2"].fillna(0)+Product_Analyzefile_df_pivot["销量1"].fillna(0)-Product_Analyzefile_df_pivot["销量3"].fillna(0)-Product_Analyzefile_df_pivot["销量4"].fillna(0))/2
Product_Analyzefile_df_pivot["计算销售日目标"]=Product_Analyzefile_df_pivot["STOCKALL"]/45
Product_Analyzefile_df_pivot["计算销售日目标"]=(Product_Analyzefile_df_pivot["计算销售日目标"].apply(lambda x:int(x+0.5)))   
Product_Analyzefile_df_pivot["计算销售周目标"]=Product_Analyzefile_df_pivot["计算销售日目标"]*7
#“销售目标差”=“计算销售周目标”-“销量1”
Product_Analyzefile_df_pivot["销售目标差"]=Product_Analyzefile_df_pivot["计算销售周目标"]-Product_Analyzefile_df_pivot["销量1"].fillna(0)
#for i in (1,10) if 销量i>0 ,BILIi=广告金额i/销量i
for i in range(1, 11):
    sales_col_i = "销量" + str(i)
    ad_col = "广告花费" + str(i)
    bili_col = "BILI" + str(i)

    sales_series = Product_Analyzefile_df_pivot[sales_col_i].fillna(0)
    ad_series = Product_Analyzefile_df_pivot[ad_col].fillna(0)

    bili_series = pd.Series(index=sales_series.index)
    for idx in bili_series.index:
        if sales_series[idx] > 0:
            bili_series.loc[idx] = ad_series.loc[idx] / sales_series.loc[idx]
        else:
            bili_series.loc[idx] = 99999999999

    Product_Analyzefile_df_pivot[bili_col] = bili_series



#for i=1 to 10  ，广告转化率i=广告订单i/广告点击量i

import numpy as np

for i in range(1, 11):
    ad_clicks_col = "广告点击量" + str(i)
    ad_orders_col = "广告订单量" + str(i)
    ad_conversion_rate_col = "广告转化率" + str(i)

    ad_clicks_series = Product_Analyzefile_df_pivot[ad_clicks_col].fillna(0)
    ad_orders_series = Product_Analyzefile_df_pivot[ad_orders_col].fillna(0)

    Product_Analyzefile_df_pivot[ad_conversion_rate_col] = ad_clicks_series.apply(lambda x: ad_orders_series[x] / ad_clicks_series[x] if ad_clicks_series[x] > 0 else (np.nan if ad_orders_series[x] == 0 and ad_clicks_series[x] == 0 else 99999999999))

#计算每一行[销量1,销量2,销量3,销量4,销量5,销量6,销量7,销量8,销量9,销量10]和[广告花费1,广告花费2,广告花费3,广告花费4,广告花费5,广告花费6,广告花费7,广告花费8,广告花费9,广告花费10]的相关系数
import numpy as np

import numpy as np

import numpy as np
import pandas as pd

def row_corr(row):
    sales_cols = pd.to_numeric(row[["销量1", "销量2", "销量3", "销量4", "销量5", "销量6", "销量7", "销量8", "销量9", "销量10"]].values, errors='coerce')
    ad_spend_cols = pd.to_numeric(row[["广告花费1", "广告花费2", "广告花费3", "广告花费4", "广告花费5", "广告花费6", "广告花费7", "广告花费8", "广告花费9", "广告花费10"]].values, errors='coerce')

    sales_cols_masked = np.ma.masked_invalid(sales_cols)
    ad_spend_cols_masked = np.ma.masked_invalid(ad_spend_cols)

    corr = np.ma.corrcoef(sales_cols_masked, ad_spend_cols_masked)[0, 1]
    return corr

Product_Analyzefile_df_pivot["10周广告相关系数"] = Product_Analyzefile_df_pivot.apply(row_corr, axis=1)

def row_corr4(row):
    sales_cols = pd.to_numeric(row[["销量1", "销量2", "销量3", "销量4"]].values, errors='coerce')
    ad_spend_cols = pd.to_numeric(row[["广告花费1", "广告花费2", "广告花费3", "广告花费4"]].values, errors='coerce')

    sales_cols_masked = np.ma.masked_invalid(sales_cols)
    ad_spend_cols_masked = np.ma.masked_invalid(ad_spend_cols)

    corr = np.ma.corrcoef(sales_cols_masked, ad_spend_cols_masked)[0, 1]
    return corr

Product_Analyzefile_df_pivot["4周广告相关系数"] = Product_Analyzefile_df_pivot.apply(row_corr4, axis=1)


#如果"Selling10"=0且STACKALL>0,"ZHOUZHUAN10"=99999999999；如果"Selling10"=0且STACKALL=0,"ZHOUZHUAN10"为空值；如果"Selling10">0，"ZHOUZHUAN10"=STACKALL/Selling10*10
Product_Analyzefile_df_pivot.loc[(Product_Analyzefile_df_pivot["Selling10"].fillna(0) == 0) & (Product_Analyzefile_df_pivot["STOCKALL"] > 0), "ZHOUZHUAN10"] = 99999999999

Product_Analyzefile_df_pivot.loc[(Product_Analyzefile_df_pivot["Selling10"].fillna(0) == 0) & (Product_Analyzefile_df_pivot["STOCKALL"] == 0), "ZHOUZHUAN10"] = np.nan

Product_Analyzefile_df_pivot.loc[~((Product_Analyzefile_df_pivot["Selling10"].fillna(0) == 0) & (Product_Analyzefile_df_pivot["STOCKALL"] >= 0)), "ZHOUZHUAN10"] = Product_Analyzefile_df_pivot["STOCKALL"] / Product_Analyzefile_df_pivot["Selling10"] * 10


WeekSalesIndex_Dic = {
    "week1": 0.2,
    "week2": 0.2,
    "week3": 0.1,
    "week4": 0.1,
    "week5": 0.1,
    "week6": 0.1,
    "week7": 0.1,
    "week8": 0.1,
}

WeekSales = 0

for i in range(1, 9):
    WeekSales += WeekSalesIndex_Dic["week" + str(i)] * Product_Analyzefile_df_pivot["销量" + str(i)]

#做计划：
def update_weekly_demands(dataframe, WeekSales):
    for i in range(2, 16):
        dataframe[f"For第{i}周销售的到货需求"] = WeekSales * i - dataframe["Fufillable"] - dataframe["Receiving"]
        for j in range(1, i):
            dataframe[f"For第{i}周销售的到货需求"] -= dataframe[f"第{j}周入库"]
        dataframe[f"For第{i}周销售的到货需求"] -= dataframe["Reserved"]

    for i in range(2, 16):
        dataframe[f"Adjusted-Week{i}"] = dataframe["ZZ2"] * 0.7 * i + dataframe[f"For第{i}周销售的到货需求"]

    return dataframe

Product_Analyzefile_df_pivot = update_weekly_demands(Product_Analyzefile_df_pivot, WeekSales)







def update_weekly_demands(dataframe, WeekSales):
    for i in range(2, 16):
        dataframe[f"For第{i}周销售的到货需求"] = WeekSales * i - dataframe["Fufillable"] - dataframe["Receiving"]
        for j in range(1, i):
            dataframe[f"For第{i}周销售的到货需求"] -= dataframe[f"第{j}周入库"]
        dataframe[f"For第{i}周销售的到货需求"] -= dataframe["Reserved"]

    for i in range(2, 16):
        dataframe[f"Adjusted-Week{i}"] = dataframe["ZZ2"] * 0.7 * i + dataframe[f"For第{i}周销售的到货需求"]

    return dataframe

Product_Analyzefile_df_pivot = update_weekly_demands(Product_Analyzefile_df_pivot, WeekSales)


###########################################################3333
def calculate_consecutive_weeks(df):
    def consecutive_weeks(group):
        group = group.sort_values('周数')
        group['销量变化'] = group['销量'].diff()
        
        consecutive_weeks = 0
        for change in group['销量变化'].iloc[1:]:
            if change > 0:
                if consecutive_weeks >= 0:
                    consecutive_weeks += 1
                else:
                    break
            elif change < 0:
                if consecutive_weeks <= 0:
                    consecutive_weeks -= 1
                else:
                    break
            else:
                break
                
        return consecutive_weeks
    
    for col in df.columns:
        if "销量" in col:
            df[col] = df[col].fillna(0)    

    grouped = df.groupby(['ASIN', '店铺', '站点'])
    rising_falling_weeks = grouped.apply(consecutive_weeks)

    unique_groups = df[['ASIN', '店铺', '站点']].drop_duplicates()
    consecutive_weeks_df = unique_groups.merge(rising_falling_weeks.reset_index(), 
                                               on=['ASIN', '店铺', '站点'], 
                                               how='left')
    consecutive_weeks_df.rename(columns={0: 'consecutive_weeks_fromweek1'}, inplace=True)

    return consecutive_weeks_df


result_df=calculate_consecutive_weeks(Product_Analyzefile_df)
#merge
Product_Analyzefile_df_pivot = pd.merge(Product_Analyzefile_df_pivot, result_df, how='left', left_on=['ASIN', '店铺', '站点'], right_on=['ASIN', '店铺', '站点'])

#将ASIN_SKU_unique中的MNSKU匹配到Product_Analyzefile_df_pivot中，按COuntry对应站点，ASIN对应ASIN的方式匹配
Product_Analyzefile_df_pivot = pd.merge(Product_Analyzefile_df_pivot, ASIN_SKU_unique, how='left', left_on=['站点', 'ASIN'], right_on=['站点', 'ASIN'])


#输出到D:\运营\2生成过程表\TESTAll_Product_Analyzefile.xlsx
Product_Analyzefile_df_pivot.to_excel(r'D:\运营\2生成过程表\TESTAll_Product_Analyzefile2.xlsx')

 
 








#                    


