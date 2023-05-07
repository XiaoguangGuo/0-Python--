import pandas as pd
from datetime import datetime, timedelta
all_sales_df=pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx',sheet_name="Sheet1")
all_sales_dfold=pd.read_excel(r'D:\运营\2生成过程表\周销售数据总表3.xlsx',sheet_name="Sheet1")
mexico_df=pd.read_excel(r'D:\运营\2019plan\Mexico周销售数据history.xlsx',sheet_name="Sheet1")
mexico_df["Country"]="GV-MX"
all_sales_df = all_sales_df.reset_index(drop=True)
all_sales_dfold = all_sales_dfold.reset_index(drop=True)
mexico_df = mexico_df.reset_index(drop=True)

all_sales_df = pd.concat([all_sales_df,all_sales_dfold],axis=0,ignore_index=True)
all_sales_df = pd.concat([all_sales_df,mexico_df],axis=0,ignore_index=True)
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

all_sales_df=update_week_numbers(all_sales_df)
#输出周销售数据总表
all_sales_df.to_excel(r'D:\运营\2生成过程表\周销售数据总表.xlsx',sheet_name="Sheet1",index=False)
