
import pandas as pd 

COLUMN_NAMES=["Record ID"," Record Type","Campaign ID","Campaign","Campaign Daily Budget","Portfolio ID","Campaign Start Date","Campaign End Date","Campaign Targeting Type","Ad Group","Max Bid","Keyword or Product Targeting","Product Targeting ID","Match Type","SKU","Campaign Status","Ad Group Status","Status","Impressions","Clicks","Spend","Orders","Total Units","Sales","ACoS","Bidding strategy","Placement Type","Increase bids by placement"]
Bulkcreation=pd.DataFrame(columns=COLUMN_NAMES) 
print(Bulkcreation.columns)



#获取某列不重复的值

#df[' sdasd'].unique()
