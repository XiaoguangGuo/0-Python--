# -*- coding:utf-8 –*-

'''
程序用来将csv批量转换为excel文件。指定源路径和目标路径。
在main函数中指定源文件路径source，目标文件路径ob.
这个程序假设csv文件放在："C:\\Users\\Administrator\\Desktop\\ceshi\\csv文件"
输出excel文件到："C:\\Users\\Administrator\\Desktop\\ceshi\\xlsx文件"
'''
 
# 导入pandas
import pandas as pd
import os


# 建立单个文件的excel转换成csv函数,file 是excel文件名，to_file 是csv文件名。 sep=';'以分号分隔的csv文件;error_bad_lines=False 忽略错误行数据
def csv_to_xlsx(file, to_file):

    data_csv = pd.read_csv(file, encoding='latin1', error_bad_lines=False, sep=';')     # 读取以分号为分隔符的csv文件   sep作用为指定分隔符，默认在Windows系统系分隔符为逗号
    data_csv.to_excel(to_file, sheet_name='data')

# 读取一个目录里面的所有文件：
def read_path(path):
    dirs = os.listdir(path)
    return dirs


# 主函数
def main():
    # 源文件路径
    source = "C:\\Users\\Administrator\\Desktop\\ceshi\\csv文件"

    # 目标文件路径
    ob = "C:\\Users\\Administrator\\Desktop\\ceshi\\xlsx文件"

    # 将源文件路径里面的文件转换成列表file_list
    file_list = [source + '\\' + i for i in read_path(source)]

    a = 0       # 列表索引csv文件名称放进j_list列表中，索引0即为第一个csv文件名称
    j_list = read_path(source)       # 文件夹中所有的csv文件名称提取出来按顺序放进j_list列表中
    print("---->", read_path(source))       # read_path(source) 本身就是列表
    print("read_path(source)类型：", type(read_path(source)))
    # 建立循环对于每个文件调用excel_to_csv()
    for it in file_list:
        j = j_list[a]    # 按照索引逐条将csv文件名称赋值给变量j
        # 给目标文件新建一些名字列表
        j_mid = str(j).replace(".csv", "")   # 将csv文件名中的.csv后缀去掉
        print("====", j_mid)
        j_xlsx = ob + '\\' + j_mid + ".xlsx"
        csv_to_xlsx(it, j_xlsx)
        print("######", it)
        a = a+1


if __name__ == '__main__':
    main()
