import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime

def get_delta_days(maxtime, datadate):
    datatimedatetime = datetime.datetime.strptime(datadate, '%Y%m%d')
    return (maxtime - datatimedatetime).days // 7 + 1

def clean_and_transform_data(sourcedata):
    columnlist = ["Campaign Daily Budget", "Max Bid", "Spend", "Sales"]
    sourcedata[columnlist] = sourcedata[columnlist].replace(',', '.', regex=True).astype(float)
    sourcedata["ACoS"] = sourcedata["ACoS"].replace(',', '.', regex=True)
    sourcedata['日期'] = pd.to_datetime(sourcedata['日期'])
    sourcedata['周数'] = 1
    return sourcedata

def main():
    newdate = input('输入最新日期y-m-d: ')
    maxtime = datetime.datetime.strptime(newdate, '%Y-%m-%d')

    bulkhzWorkbook = load_workbook(r'D:\运营\2生成过程表\周bulk广告数据汇总表.xlsx')
    sheet = bulkhzWorkbook.active

    bulkdatafilepath = 'D:\\运营\\1数据源\\周bulk广告数据\\'

    for bulkdatafile in os.listdir(bulkdatafilepath):
        datadate = bulkdatafile.split('-')[4]
        delta = get_delta_days(maxtime, datadate)

        sourcedata = pd.read_excel(bulkdatafilepath + str(bulkdatafile), engine="openpyxl", sheet_name=1).assign(
            Country=os.path.basename(bulkdatafile).split('_')[0], 日期=os.path.basename(bulkdatafile).split('-')[4])

        sourcedata = clean_and_transform_data(sourcedata)

        for row in dataframe_to_rows(sourcedata, index=False, header=False):
            sheet.append(row)

        bulkhzWorkbook.save(r'D:\运营\2生成过程表\周bulk广告数据汇总表.xlsx')

        shutil.copy(r'D:\\运营\\1数据源\\周bulk广告数据\\' + str(bulkdatafile), r'D:\\运营\\1数据源\\bulkoperationfiles\\')
        shutil.copy(r'D:\\运营\\1数据源\\周bulk广告数据\\' + str(bulkdatafile), r'D:\\运营\\HistoricalData\\周bulk广告数据\\')

    Allbulkpath = 'D:\\运营\\2生成过程表\\'
    Allbulk = pd.read_excel(Allbulkpath + '周bulk广告数据汇总表.xlsx')
    Allbulk["周数"] = (maxtime - Allbulk["日期"]).dt.days // 7 + 1
    Allbulk.to_excel(Allbulkpath + '周bulk广告数据汇总表.xlsx', sheet_name="Sheet", index=False)

if __name__ == "__main__":
    main()
