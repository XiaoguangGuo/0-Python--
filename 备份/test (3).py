

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil

# -*- coding:utf-8 –*-
import os
import pandas as pd
import shutil




data_tsv5= pd.read_csv( r'D:/运营/计划数据/老站/在途库存/US_FBA16KFY253T.tsv',sep='\t',header=7)    # 读取以分
print(data_tsv5)

data_tsv6= pd.read_csv( r'D:/运营/计划数据/老站/在途库存/US_FBA16KFY253T.tsv',sep='\t',header=8)    # 读取以分
print(data_tsv6)
