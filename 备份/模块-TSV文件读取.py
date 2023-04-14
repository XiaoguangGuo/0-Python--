# -*- coding: utf-8 -*-
import pandas as pd 
import csv
import numpy as np

#data_tsv5= pd.read_csv(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv', sep='\t')
 
#print(data_tsv5)
                       
#第二种方法：
#with open(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv') as fd:
    #rd = csv.reader(fd, delimiter="\t", quotechar='"')
    #for row in rd:
        #print(row)

#sep='\t',
#import scipy as sp

#data = sp.genfromtxt(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv', delimiter="\t")

#第三种方

#import csv
with open(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv', 'rb') as f:
    print(len(f.readlines()))
    datatsv=csv.reader(f,delimiter="\t")
    datanew=pd.DataFrame(datatsv)
    print(datanew)

#def read_from_tsv(file_path: str, column_names: list) -> list:
    #csv.register_dialect('tsv_dialect', delimiter='\t', quoting=csv.QUOTE_ALL,encoding='UTF-8)
    #with open(file_path, "r") as wf:
        #reader = csv.DictReader(wf, fieldnames=column_names, dialect='tsv_dialect')
       # datas = []
        #for row in reader:
            #data = dict(row)
            #datas.append(data)
    #csv.unregister_dialect('tsv_dialect')
    #return datas
#datatsv=read_from_tsv(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv',"")

#df = pd.read_csv(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv', header=None, sep='\\t', escapechar='\\t',
                 #quoting=csv.QUOTE_NONE,  engine='python',
                 #encoding='iso8859_1')

#print(df)
       
#path1 = r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv'
#file1 = open(path1)
#datanew= pd.read_csv(path1, sep='\t',error_bad_lines=False,header=0,skiprows=8)
#rint(datanew)
#datanew.to_excel(r'D:/运营/计划数据/NewCountries/在途库存/test.xlsx')
