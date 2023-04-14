import os
import xlrd
import xlsxwriter
import pandas as pd
import openpyxl
import xlwings


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
workbook = xlsxwriter.Workbook(target_xls)
worksheet = workbook.add_worksheet()

worksheet.write('A1', '数据1')
worksheet.write('B1', '数据2')
worksheet.write('R1', '数据')
data = []

for i in source_xls:
 
#用pandas打开
   # wb=pd.read_excel(i,engine='openpyxl')；
   #如果不能打开,除了xlrd的版本问题，还有可能是文件夹中有隐藏文件，要删除文件夹重新建立。
    wb = xlrd.open_workbook(i)
    for sheet in wb.sheets():
        for rownum in range(1,sheet.nrows)
        # 从第二行开始合并
            print(rownum)
        #合并excel中的所有数据
        
            a = sheet.row_values(rownum)
            a.append(i.replace((filePath + '\\'), ''))
            #将文件名做为新的一列
            data.append(a)
            #data.append(sheet.row_values(rownum))
#print(data)
# 写入数据



font = workbook.add_format({"font_size":14})
for i in range(len(data)):
    for j in range(len(data[i])):
        worksheet.write(i, j, data[i][j], font)

 
workbook.close()        

#分列
wertyyu = pd.read_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx')
 
#注意名称一定要有_,能分列；如果列为空就会报错。    
df=wertyyu["数据"].str.split("_",expand=True)

wertyyu["数据"]=df[0]
wertyyu["其他"]=df[1]
print(wertyyu.head(5))
print(df.head(5))

wertyyu.to_excel(r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx',index=False)


#复制一张表到另一张表
import sys

sys.path.append(r'D:\PythonDocs\Python程序')

#2、读取待复制的表格
xlpath = r'D:\PythonDocs\AmazonPlan\Amazonplan.xlsx'

xlworkbook = xlwings.Book(xlpath)
print(xlworkbook)

#3、读取待粘贴的表格
xlpath2 = r'D:\运营\计划\销售数据\sales.xlsx'
xlworkbook2 = xlwings.Book(xlpath2)

#3-1、找到最后一行的第一个单元格
rng = xlworkbook2.sheets("Sheet1").range('A1').expand('table')

cell_index = str(rng.rows.count+1)

range1 = xlworkbook2.sheets("Sheet1").range('A'+cell_index)

#3-2、按行复制数据到目标表格。
range1.value = xlworkbook.sheets("Sheet1").range('A1').expand('table').value



