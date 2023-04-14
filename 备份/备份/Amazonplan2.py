import os
import xlrd
import xlsxwriter


filePath = r'D:\PythonDocs\AmazonPlan'
#需合并的文件所在的文件夹路径

f_name = os.listdir(filePath)
#读取文件夹内所有文件名
#print(f_name)

source_xls= []
for i in f_name:
    source_xls.append(filePath + '\\' + i)
#将文件路径存储在列表中
#print(source_xls)

target_xls = r"D:\PythonDocs\AmazonPlan\Amazonplan.xlsx"
#合并后文件的路径

# 读取数据
data = []
for i in source_xls:
    wb = xlrd.open_workbook(i)
    for sheet in wb.sheets():
        for rownum in range(sheet.nrows):
        #合并excel中的所有数据
            a = sheet.row_values(rownum)
            a.append(i.replace((filePath + '\\'), ''))
            #将文件名做为新的一列
            data.append(a)
            #data.append(sheet.row_values(rownum))
#print(data)
# 写入数据

workbook = xlsxwriter.Workbook(target_xls)
worksheet = workbook.add_worksheet()
font = workbook.add_format({"font_size":14})
for i in range(len(data)):
    for j in range(len(data[i])):
        worksheet.write(i, j, data[i][j], font)
# 关闭文件流
workbook.close()
