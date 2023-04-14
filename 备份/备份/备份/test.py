import csv
from xlsxwriter.workbook import workbook
src_dir_path_inventory=r'D:\运营\计划数据\老站\在途库存'

filePath = r'D:\PythonDocs\AmazonPlan'
f_name = os.listdir(filePath)
#读取文件夹内所有文件名
#print(f_name)

#source_xls= []
for i in f_name:
    source_xls.append(filePath + '\\' + i)
