import os
import pandas as pd
import shutil
import datetime
from datetime import date

def get_file_date(file_name):
    date_str = os.path.basename(file_name).split("_")[2]
    return date_str[0:10]

def process_files(Product_Analyzepath, All_Product_Analyzefile):
    for Product_Analyzefile in os.listdir(Product_Analyzepath):
        print(Product_Analyzefile)
        date_str = get_file_date(Product_Analyzefile)
        Product_Analyzefile_DF = pd.read_excel(Product_Analyzepath + str(Product_Analyzefile)).assign(日期=date_str)
        All_Product_Analyzefile["周数"] = 1
        All_Product_Analyzefile = All_Product_Analyzefile.append(Product_Analyzefile_DF, ignore_index=True)
        shutil.move(Product_Analyzepath + str(Product_Analyzefile), 'D:/运营/HistoricalData/Product_Analyzefile')
    return All_Product_Analyzefile

def process_weekly_data(max_week, all_product_analyzefile, all_product_analyzefile_weeks):
    for i in range(1, max_week):
        all_product_analyzefile_weeks_i = all_product_analyzefile.loc[(all_product_analyzefile["周数"] == i)]

        if i == 1:
            all_product_analyzefile_weeks_i = all_product_analyzefile_weeks_i[["ASIN", "店铺", "站点", 'MSKU', "FBA可售", "可售天数预估", "标签", "销量", "销售额", '广告点击量', '广告花费', '广告订单量', '毛利润']]
            all_product_analyzefile_weeks_i.rename(columns={'销量': '销量' + str(i), '销售额': '销售额' + str(i), '广告点击量': '广告点击量' + str(i), '广告花费': '广告' + str(i), '广告订单量': '广告订单' + str(i), '毛利润': '毛利润' + str(i)}, inplace=True)
        else:
            all_product_analyzefile_weeks_i = all_product_analyzefile_weeks_i[["ASIN", "店铺", "站点", 'MSKU', "销量", "销售额", '广告点击量', '广告花费', '广告订单量', '毛利润']]
            all_product_analyzefile_weeks_i.rename(columns={'销量': '销量' + str(i), '销售额': '销售额' + str(i), '广告点击量': '广告点击量' + str(i), '广告花费': '广告' + str(i), '广告订单量': '广告订单' + str(i), '毛利润': '毛利润' + str(i)}, inplace=True)

        # 合并
        all_product_analyzefile_weeks = pd.merge(all_product_analyzefile_weeks, all_product_analyzefile_weeks_i, on=["ASIN", "店铺", "站点", "MSKU"], how="left")

        all_product_analyzefile_weeks = all_product_analyzefile_weeks.drop_duplicates()

    return all_product_analyzefile_weeks

def main():
    newdate = input('输入最新日期y-m-d：新站汇总到周日', )
    maxtime = datetime.datetime.strptime(newdate, '%Y-%m-%d')
    print(maxtime)
    maxtimeday = datetime.datetime.strptime(newdate, '%Y-%m-%d').date()
    print(maxtimeday)

    Product_Analyzepath = r'D:\\运营\\1数据源\\Product_Analyze产品分析\\'
    print(Product_Analyzepath)
    All_Product_Analyzefile = pd.read_excel(r'D:\\运营\2生成过程表\\All_Product_Analyzefile.xlsx', sheet_name=0)
    All_Product_Analyzefile = process_files(Product_Analyzepath, All_Product_Analyzefile)
    
    # ... (保留原始代码)
    All_Product_Analyzefile_Weeks=All_Product_Analyzefile[["ASIN","店铺","站点","MSKU"]].drop_duplicates()
    All_Product_Analyzefile['日期'] = pd.to_datetime(All_Product_Analyzefile['日期'])
    print(All_Product_Analyzefile[['日期']])

    All_Product_Analyzefile['周数']=(maxtime-All_Product_Analyzefile['日期']).dt.days//7+1
    #按周数重整理表格
    All_Product_Analyzefile_Weeks = process_weekly_data(max_week, All_Product_Analyzefile, All_Product_Analyzefile_Weeks)
                                     
    All_Product_Analyzefile.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

    All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile_Weeks.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)

    All_Product_Analyzefile_Weeks=All_Product_Analyzefile_Weeks[["ASIN","店铺","站点",'MSKU',"FBA可售","可售天数预估","标签","销量1","销量2","销量3","销量4","销量5","销量6","销量7","销量8","销量9","销量10","销售额1","销售额2","销售额3","销售额4","销售额5","销售额6","销售额7","销售额8","销售额9","销售额10","广告点击量1","广告点击量2","广告点击量3","广告点击量4","广告点击量5","广告点击量6","广告点击量7","广告点击量8","广告点击量9","广告点击量10","广告1", "广告2","广告3","广告4","广告5","广告6","广告7","广告8","广告9","广告10", "广告订单1","广告订单2","广告订单3","广告订单4","广告订单5","广告订单6","广告订单7","广告订单8","广告订单9","广告订单10"]]

    All_Product_Analyzefile_Weeks.to_excel(r'D:\\运营\\2生成过程表\\All_Product_Analyzefile_Weeks排序.xlsx',sheet_name="sheet1",startrow=0,header=True,index=False)




    
if __name__ == "__main__":
    main()
