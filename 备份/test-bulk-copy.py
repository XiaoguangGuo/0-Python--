import pandas as pd

bulkone=pd.read_excel(r'D:\\运营/bulkoperationfilesNEW\\bulkone.xlsx',sheet_name=0)

sku=pd.read_excel(r'D:\\运营/bulkoperationfilesNEW\\bulkone.xlsx',sheet_name=1,names=["SKU"])
skuLIST=sku["SKU"].tolist()
print(skuLIST)

bulkonenew=bulkone
for skuname in skuLIST:
    print(skuname)
    bulk2=bulkone
   
    bulk3=bulk2.replace("N20190513-Cherry-Red",skuname)
    print(bulk3)
    bulkonenew=bulkonenew.append(bulk3)

bulkonenew.to_excel(r'D:\\运营/bulkoperationfilesNEW\\bulkone-new.xlsx')
