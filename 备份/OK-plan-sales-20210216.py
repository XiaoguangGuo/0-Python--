import os
import xlrd
import xlsxwriter
import pandas as pd
import openpyxl
import xlwings


filePath = r'D:\运营\计划数据\Newcountries\销售数据'
#需合并的文件所在的文件夹路径

f_name = os.listdir(filePath)
#读取文件夹内所有文件名
#print(f_name)

source_xls= []
for i in f_name:
    source_xls.append(filePath + '\\' + i)
#将文件路径存储在列表中
#print(source_xls)

target_xls = r"D:\SailingstarFBA计划\newsalesdata.xlsx"
#合并后文件的路径

# 读取数据


data = []

for i in source_xls:

#用pandas打开
   # wb=pd.read_excel(i,engine='openpyxl')；
   #如果不能打开,除了xlrd的版本问题，还有可能是文件夹中有隐藏文件，要删除文件夹重新建立。
    wb = xlrd.open_workbook(i)
    for sheet in wb.sheets():
        for rownum in range(0,sheet.nrows):
            print(rownum)
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
        
worksheet.write('R1', '国家')
 
workbook.close()        

#分列
wertyyu = pd.read_excel(r"D:\SailingstarFBA计划\newsalesdata.xlsx")
 
#注意名称一定要有_,能分列；如果列为空就会报错。    
df=wertyyu["国家"].str.split("_",expand=True)

wertyyu["国家"]=df[0]
wertyyu["日期"]=df[1]
wertyyu["其他"]=df[2] #由于要有国家日期和其他，必须三个

print(wertyyu.head(5))
print(df.head(5))


#qudiao 除第一行
#通过.isin()，选取包含指定字符串"boy"的行

#wertyyu=wertyyu[~wertyyu.icol(0).isin(["asin"])]  
#通过~取反，选取不包含指定字符串"boy"的行
wertyyu100=wertyyu[~wertyyu.iloc[: , 3].isin(["SKU"])]  
# 学习：删除指定列"sImagePath"=="wj"或者"sImagePath"=="/"的行数据
#df_checkimage = df_checkimage[~df_checkimage["sImagePath"].isin(["/","wj"])]

#去掉标题行，失败，实质是去掉了第一行数据
# wertyyu100.drop([0,0],inplace=True)    

wertyyu100.to_excel(r"D:\SailingstarFBA计划\newsalesdata.xlsx",index=False)


#复制一张表到另一张表
                      
import sys

sys.path.append(r'D:\运营\Python程序')

#2、读取待复制的表格
xlpath = r'D:\SailingstarFBA计划\newsalesdata.xlsx'

xlworkbook = xlwings.Book(xlpath)
print(xlworkbook)

#3、读取待粘贴的表格
xlpath2 = r'D:\SailingstarFBA计划\NEW-ALL周销售数据.xlsx'
xlworkbook2 = xlwings.Book(xlpath2)

#3-1、找到最后一行的第一个单元格
rng = xlworkbook2.sheets("Sheet1").range('A1').expand('table')

cell_index = str(rng.rows.count+1)

range1 = xlworkbook2.sheets("Sheet1").range('A'+cell_index)

#3-2、按行复制数据到目标表格。
#从第二行开始复制
range1.value = xlworkbook.sheets("Sheet1").range('A2').expand('table').value

xlworkbook.save()
xlworkbook.close()
xlworkbook2.save()
xlworkbook2.close()

zengjiazhoushu=pd.read_excel(r'D:\SailingstarFBA计划\NEW-ALL周销售数据.xlsx')
print(zengjiazhoushu.head())
print(zengjiazhoushu["日期"].max())



