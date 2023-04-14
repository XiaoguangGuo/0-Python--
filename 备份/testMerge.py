import pandas as pd 

ListingPd=pd.read_excel(r'D:/运营/Listinginfo.xlsx')

print(ListingPd)

Allbulk=pd.read_excel(r'D:/运营/周bulk广告数据汇总表.xlsx')
Allbulk_fenlei=pd.merge(Allbulk,ListingPd,how="left",left_on=["Country","SKU"],right_on=["COUNTRY","SKU"])
print(Allbulk_fenlei)
Allbulk_fenlei.to_excel(r'D:/运营/周bulk广告数据汇总表fenlei.xlsx')

# yunxing man Allbulk_SKUCAMPAIGN=Allbulk[Allbulk["SKU"].notna()].drop_duplicates("Campaign","SKU",inplace=True)

Allbulk_SKUCAMPAIGN= Allbulk.groupby(["Campaign","SKU"],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].sum()
Allbulk_SKUCAMPAIGN2= Allbulk.groupby(["Campaign","SKU"],as_index=False)['Impressions','Clicks','Spend','Orders','Total Units','Sales'].agg("sum")
# Allbulk_SKUCAMPAIGN= Allbulk_fenlei.groupby(["Campaign","SKU"],).agg('Impressions','Clicks','Spend','Orders','Total Units','Sales')
#Allbulk_Dalei= Allbulk_fenlei.groupby(["大类","小类","Country"]).agg('Impressions','Clicks','Spend','Orders','Total Units','Sales')

print(Allbulk_SKUCAMPAIGN)
print(Allbulk_SKUCAMPAIGN2)
Allbulk_SKUCAMPAIGN.to_excel(r'D:/运营/Allbulk_SKUCAMPAIGN1.xlsx')
Allbulk_SKUCAMPAIGN2.to_excel(r'D:/运营/Allbulk_SKUCAMPAIGN2.xlsx')
#print(Allbulk_Dalei)
input("?")

Allbulk_SKUCAMPAIGN_Merge=pd.merge(Allbulk_SKUCAMPAIGN,ListingPd,how="left",left_on=["Country","SKU"],right_on=["COUNTRY","SKU"])
Print(Allbulk_SKUCAMPAIGN_Merge)

