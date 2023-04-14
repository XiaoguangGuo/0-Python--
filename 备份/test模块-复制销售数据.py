import glob, os
src_dir_path_sales=r'D:\运营\计划数据\老站\销售数据'
print(os.listdir(src_dir_path_sales))
data_csv2 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))

data_sales_US=pd.read_excel(r'D:\2019plan\周销售数据.xlsx')
data_sales_CA=pd.read_excel(r'D:\2019plan\Canada周销售数据.xlsx')
data_sales_MX=pd.read_excel(r'D:\2019plan\Mexico周销售数据.xlsx')
salescolumns_US=data_sales_US.columns.tolist()
salescolumns_CA=data_sales_CA.columns.tolist()
salescolumns_MX=data_sales_MX.columns.tolist()
#文件


for file in os.listdir(src_dir_path_sales):
     
    if key[0] in file
    data_csv3 = pd.read_table(r'D:\\运营\\计划数据\\老站\\销售数据\\'+ str(file))    # 读取以分
    
    data['日期'] = ''
    # 把文件名的分列第一段写入日期

    files = glob.glob('files/*.csv')
    print(files)
# basename
    data_csv3 = pd.concat([pd.read_csv(fp).assign(New=os.path.basename(fp).split('_')[0]) for fp in files])
