

import pandas as pd
import os

import datetime 
import openpyxl
import numpy as np



SearchSummaryWeeks=pd.read_excel(r'D:\\运营\\2生成过程表\\Search_Term_Summary.xlsx',sheet_name="SeachTermWeekSum_Weeks")
SearchSummaryWeeks=SearchSummaryWeeks[SearchSummaryWeeks["Country"]=="NEW-US"]

SearSummaryBiaotou=SearchSummaryWeeks[["Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term"]]


Searchtempo=pd.read_excel(r'D:/运营/1数据源/temp-SearchTerm/Sponsored Products Search term report.xlsx',engine="openpyxl")

SearchtempoBoaotou=Searchtempo[["Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term"]]

SearSummaryBiaotou=pd.concat([SearSummaryBiaotou,SearchtempoBoaotou],ignore_index=True)
SearSummaryBiaotou=SearSummaryBiaotou.drop_duplicates()
SearSummaryBiaotou=pd.merge(SearSummaryBiaotou,Searchtempo,on=["Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term"] ,how="left")
SearchSummaryWeeks=pd.merge(SearSummaryBiaotou,SearchSummaryWeeks,on=["Campaign Name", "Ad Group Name","Targeting","Match Type","Customer Search Term"] ,how="left")


writer=pd.ExcelWriter(r'D:\\运营\\2生成过程表\\Search_Term_SummaryNEW-US.xlsx')

SearchSummaryWeeks.to_excel(writer,"SearchSummaryWeeksTempo",index=False)

 
writer.close()
