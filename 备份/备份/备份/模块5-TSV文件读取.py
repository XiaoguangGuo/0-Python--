# -*- coding: utf-8 -*-
import pandas as pd 
import csv
import numpy as np

datatsv = pd.read_csv(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv', sep='\t' ,header=7  )
 
print(datatsv)

#第二种方法：
#with open(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv') as fd:
    #rd = csv.reader(fd, delimiter="\t", quotechar='"')
    #for row in rd:
        #print(row)

#sep='\t',
#import scipy as sp

#data = sp.genfromtxt(r'D:/运营/计划数据/NewCountries/在途库存/UK_FBA15DF4HGS4.tsv', delimiter="\t")

#第三种方
