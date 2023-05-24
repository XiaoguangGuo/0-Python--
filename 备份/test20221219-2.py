
from openpyxl import load_workbook,Workbook
import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime 
import shutil
import numpy as np

bulkfilepath=r'D:\运营\4行动表\bulkoperationfilesNEW'

CountriesProduct=pd.read_excel(r'D:\\运营\\3数据分析结果\\国家汇总.xlsx",engine="openpyxl",sheet_name="ProductActions")
dirlist=os.listdir(bulkfilepath)
for bulkchangefile in dirlist: #找bulkfile对应的国家文件 os.listdir(bulkoperationfilespath)
    countryname20001219=bulkchangefile.split('-')[2]
    print(countryname20001219)
    input()
  
    Changefiledf=pd.read_excel(bulkdatafilepath+"\\"+str(bulkchangefile),engine="openpyxl")
    Changefiledf["STOCKALL"]=""
    Changefiledf["皮质标签"]=""
    Changefiledf["行动方案"]=""

    bulkchangefileSKU_list=Changefiledf["SKU"].drop_duplicates().to_list()
    for bulkchangefileSKU in bulkchangefileSKU_list:
        if bulkchangefileSKU in CountriesProduct.loc[CountriesProduct["Country"]==countryname20001219,"SKU"].to_list():
            stock_oi=CountriesProduct.loc[(CountriesProduct["Country"]==countryname20001219)&(CountriesProduct["SKU"].str.contains(bulkchangefileSKU)),"STOCKALL"].value[0]
            label_oi=CountriesProduct.loc[(CountriesProduct["Country"]==countryname20001219)&(CountriesProduct["SKU"].str.contains(bulkchangefileSKU)),"皮质标签"].value[0]
            action_oi=CountriesProduct.loc[(CountriesProduct["Country"]==countryname20001219)&(CountriesProduct["SKU"].str.contains(bulkchangefileSKU)),"皮质标签"].value[0]                              
            
            Changefiledf.loc[Changefiledf["SKU"]==bulkchangefileSKU,"STOCKALL"]== stock_oi                          
            Changefiledf.loc[Changefiledf["SKU"]==bulkchangefileSKU,"皮质标签"]== label_oi
            Changefiledf.loc[Changefiledf["SKU"]==bulkchangefileSKU,"行动方案"]== action_oi

    bulkchangefile.to_excel(bulkfilepath+"\\NEW-"+str(bulkchangefile),index=False)
