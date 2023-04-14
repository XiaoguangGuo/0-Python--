# -*- coding:utf-8 –*-
import os
import pandas as pd
import chardet


# 读取文件夹中所有的csv和xlsx文件
dir_path = r'D:\运营\1数据源\Amazon品牌搜索词'
file_list = [file for file in os.listdir(dir_path) if file.endswith('.csv')]

for file in file_list:
    # 用_分割文件名，第四个部分的前面两个字母为国家名
    file_parts = file.split('_')
    country = file_parts[3][:2]

    # 打开文件，表格中第一行A列中[]内为“站点名”
    file_path = os.path.join(dir_path, file)
    print(file_path)

with open(file_path, 'rb') as f:
    result = chardet.detect(f.read())
    encoding = result['encoding']
    print(encoding)
    df =  pd.read_csv(file_path,encoding=encoding)
    station = df.columns[0].split('[')[1].split(']')[0]

    # 第一行F列中的 [ ]内第四个“-”后面部分为“截止日期”，第一行F列中的 [ ]内第四个“-”前面部分为“开始日期”
    end_date = df.iloc[0, 5].split('-')[3]
    start_date = df.iloc[0, 5].split('-')[2]

    # 将“站点名”，“开始日期”和“截止日期”插入到表格的第1列，第2列和第3列
    df.insert(0, '站点名', station)
    df.insert(1, '开始日期', start_date)
    df.insert(2, '截止日期', end_date)

    # 其余列向后移动
    df.columns = pd.RangeIndex(len(df.columns))
    df.columns = ['列'+str(i+1) for i in range(len(df.columns))]
    df = df.rename(columns={'列1': '站点名', '列2': '开始日期', '列3': '截止日期'})
    
    # 将文件另存为csv格式
    new_file_name = f'{country}_{file_parts[4]}_{file_parts[5]}.csv'
    df.to_csv(os.path.join(dir_path, new_file_name), index=False, encoding='utf-8-sig')
