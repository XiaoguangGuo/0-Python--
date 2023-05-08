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
Product_Analyzefile_df=update_week_numbers(Product_Analyzefile_df)
Product_Analyzefile_df["周数"]+=1
all_sales_df=pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx',sheet_name="Sheet1")
all_advertise_df=pd.read_excel(r'D:\运营\2生成过程表\周bulk数据Summary.xlsx',sheet_name="SKU-Campaign-WEEK")
all_sales_df=update_week_numbers(all_sales_df)

all_advertise_df=all_advertise_df[['Country','SKU','周数','Impressions','Clicks','Spend','Orders','Total Units','Sales']]

#把all_advertise_df merge到all_sales_df，匹配 Country，SKU，week周数，all_advertise_df concat只取[Impressions,	Clicks	,Spend	,Orders	,Total Units,	Sales]这几列
all_sales_df = pd.merge(all_sales_df, all_advertise_df, how='left', left_on=['Country','SKU','周数'], right_on=['Country','SKU','周数'])







#all_salses_df rename
all_sales_df.rename(columns={'(Child) ASIN':'ASIN','Country':'站点','Sessions - Total':'Sessions','Units Ordered':'销量','Ordered Product Sales':'销售额','Clicks':'广告点击量','Total Units':'广告订单','Sales':'广告销售额'},inplace=True)
all_sales_df["店铺"]=all_sales_df["站点"]
all_sales_df["MSKU"]=all_sales_df["SKU"]
#把all_sales_df concat到Product_Analyzefile_df
Product_Analyzefile_df = pd.concat([Product_Analyzefile_df,all_sales_df],axis=0,ignore_index=True) 
#获取ASIN列和MSKU的对照表
ASIN-SKU=Product_Analyzefile_df[["Country","ASIN","MSKU"]].unique()
#按Country和Asin分组后MSKU列不同行的字符串合并，用逗号拼接。形成新的MSKU列，然后drop掉MSKU列，去重。
 

# 定义一个函数，用于合并和去重MSKU
def merge_and_deduplicate(mskus):
    unique_mskus = set()
    for msku in mskus:
        # 分割逗号分隔的字符串，并将结果添加到集合中
        unique_mskus.update(msku.split(','))
    return ','.join(unique_mskus)

# 使用groupby和apply方法合并和去重MSKU
ASIN-SKU_unique = ASIN-SKU.groupby(['站点', 'ASIN'])['MSKU'].apply(merge_and_deduplicate).reset_index()
ASIN-SKU_unique.columns = ['站点', 'ASIN', 'MNSKU']

 

def merge_and_expand_msku(df):
    df['MSKU'] = df['MSKU'].apply(lambda x: set(x.split(',')))
    merged_df = df.groupby(['站点', 'ASIN'])['MSKU'].apply(lambda x: set.union(*x)).reset_index()
    expanded_df = merged_df.explode('MSKU').reset_index(drop=True)
    return expanded_df

ASIN-SKU_multiple = merge_and_expand_msku(ASIN-SKU)








Product_Analyzefile_df=Product_Analyzefile_df[Product_Analyzefile_df["周数"]<=52]
#输出Product_Analyzefile_df到D:\运营\2生成过程表\TESTAll_Product_Analyzefilebeforepivot.xlsx
Product_Analyzefile_df.to_excel(r'D:\运营\2生成过程表\TESTAll_Product_Analyzefilebeforepivot.xlsx')



# Create the pivot table
Product_Analyzefile_df_pivot = Product_Analyzefile_df.pivot_table(
    index=['ASIN', '店铺', '站点'],
    columns='周数',
    values=['Sessions', '销量', '销售额', '广告点击量', '广告订单', '广告销售额','毛利润'],
    aggfunc='sum'
)

# Generate column names with the original column names and the week number
new_columns = [(col[0] + str(col[1])) for col in Product_Analyzefile_df_pivot.columns]

# Reset the index and set new column names
Product_Analyzefile_df_pivot.reset_index(inplace=True)
Product_Analyzefile_df_pivot.columns = ['ASIN', '店铺', '站点'] + new_columns



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

#将ASIN-SKU_unique中的MNSKU匹配到Product_Analyzefile_df_pivot中，按COuntry对应站点，ASIN对应ASIN的方式匹配
Product_Analyzefile_df_pivot = pd.merge(Product_Analyzefile_df_pivot, ASIN-SKU_unique, how='left', left_on=['站点', 'ASIN'], right_on=['站点', 'ASIN'])


#输出到D:\运营\2生成过程表\TESTAll_Product_Analyzefile.xlsx
Product_Analyzefile_df_pivot.to_excel(r'D:\运营\2生成过程表\TESTAll_Product_Analyzefile.xlsx')

 
 








#                    


