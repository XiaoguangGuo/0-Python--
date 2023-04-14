import pandas as pd

data_sales_US=pd.read_excel(r'D:\2019plan\周销售数据.xlsx')
data_sales_CA=pd.read_excel(r'D:\2019plan\Canada周销售数据.xlsx')

print(data_sales_US.columns.size,data_sales_CA.columns.size)
a=data_sales_US.columns.size
b=data_sales_CA.columns.size
print(a)
print(b)
c=a-b
print("c",c)
if c==0:
    print("相等")
else:
    print("不相等")

d=len(data_sales_US)
print(d)
e=len(data_sales_CA)
f=d-e
if f==0:
    print("len列相等")
else:
    print("len不等")
print(f)
