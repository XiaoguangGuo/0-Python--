# -*- coding:utf-8 –*-
import os
import pandas as pd
src_dir_path=r'D:\2019plan\data_销售'        
key =['US','CA','MX']
t=key[0]
print(t)

for file in os.listdir(src_dir_path):
    print(file)
            
    if key[0] in file:
    # 执行语句
        print("有US")
        data_csv = pd.read_csv(r'D:\\2019plan\\data_销售\\'+ str(file),encoding='utf-8 ', error_bad_lines=False)    # 读取以分
        print(data_csv)
        data_csv.to_excel(r'D:\2019plan\sales.xlsx', startrow=1,header=True,index=false)
    elif key[1]in file:
        print("有CA")
        data_csv = pd.read_csv(r'D:\\2019plan\data_销售\\'+str(file),encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        data_csv.columns=["aq","bq","cq","dq","eq","fq","hq","iq","jq","kq","lq","mq","nq","oq","pq","qq","rq"]
                         
        data_csv.to_excel(r'D:\2019plan\sales.xlsx', startrow=0,header=True,index=False)

        print(data_csv)
      
    elif key[2]in file:
        print("有MX")
        data_csv = pd.read_csv(r'D:\\2019plan\data_销售\\'+str(file), encoding='utf-8 ', error_bad_lines=False)     # 读取以分
        print(data_csv)
    else:
        print("都没有")
                



                      
         #  Print

                  #  For i in key
                      #if key[i] in file
                  #    print(key[i])
        

                    
