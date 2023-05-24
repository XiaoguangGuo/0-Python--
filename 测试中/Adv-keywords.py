#################################################################################################做Summa
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

conn = sqlite3.connect('D:/运营/sqlite/AmazonData.db')

df = pd.read_sql_query('SELECT * FROM "Bulkfiles"', conn)
df = df[df['日期'].notna()]

df['Spend'] = df['Spend'].astype(float)
df["Max Bid"] = df["Max Bid"].astype(float)
df['Sales'] = df['Sales'].astype(float)
df['日期'] = pd.to_datetime(df['日期'])

def find_last_saturday():
    today = datetime.now()
    last_saturday = today - timedelta(days=today.weekday() + 2)
    return last_saturday

def update_week_numbers(df):
    last_saturday = find_last_saturday()
    df['日期'] = pd.to_datetime(df['日期'])
    df['周数'] = ((last_saturday - df['日期']).dt.days // 7) + 1
    return df

All_adv = update_week_numbers(df)
All_adv=All_adv[All_adv["周数"]<53]


cols_to_convert1 = ["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales", "ACoS"]

All_adv[cols_to_convert1]=All_adv[cols_to_convert1].replace(',','.',regex=True)
All_adv["Acos"]=All_adv["ACoS"]*1

All_adv=All_adv[All_adv["Country"]=="NEW-JP" ]
All_adv_selected=All_adv.loc[All_adv["Keyword or Product Targeting"].notnull(),["Country","Campaign","Ad Group","Max Bid",	"Keyword or Product Targeting",	"Match Type","Campaign Status",	"Ad Group Status","Status","周数","Impressions","Clicks","Spend","Orders","Total Units","Sales","ACoS"]]
All_adv_selected.loc[All_adv_selected["Clicks"]>0,"转化率"]=All_adv_selected["Orders"]/All_adv_selected["Clicks"]

cols_to_convert = ["Impressions", "Clicks", "Spend", "Orders", "Total Units", "Sales", "ACoS","转化率"]

#for col in cols_to_convert:
    #All_adv_selected[col] = pd.to_numeric(All_adv_selected[col], errors='coerce')

All_advJp1=All_adv[(All_adv["周数"]==1 )&(All_adv["Country"]=="NEW-JP" )]
#输出All_advJp1
All_advJp1.to_excel('D:\\运营\\2生成过程表\\All_advJp1.xlsx',index=False,encoding='utf_8_sig')

input("Press Enter to continue...")
All_adv_pivot = All_adv_selected.pivot_table(
    index=["Country","Campaign","Ad Group","Keyword or Product Targeting","Match Type","Campaign Status",	"Ad Group Status","Status"],
    columns='周数',
    values=["Impressions","Clicks",	"Spend","Orders","Total Units","Sales",	"ACoS","转化率"],
    aggfunc='first'
)


# Generate column names with the original column names and the week number
new_columns = [(col[0] + str(col[1])) for col in All_adv_pivot.columns]

# Reset the index and set new column names
All_adv_pivot.reset_index(inplace=True)
All_adv_pivot.columns = ["Country","Campaign","Ad Group",	"Keyword or Product Targeting",		"Match Type","Campaign Status",	"Ad Group Status","Status"] + new_columns


Allbulkpath='D:\\运营\\2生成过程表\\'

writer=pd.ExcelWriter(Allbulkpath+'All_ad.xlsx')
 
All_adv_pivot.to_excel(writer,"All_adv__pivottest0514",index=False) 
writer.close()
