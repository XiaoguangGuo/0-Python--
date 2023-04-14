
key =['US','CA','MX','UK','IT','DE','JP']
keydic = { "US" : '333' , 'CA' : '666' , 'MX' : '999','UK':'111','IT':'222','DE':'888','JP':'1.5' }
print (key)
print (len(key))
print(keydic["US"])

 
for i in range(len(key)):
    print(key[i])
    print(str(key[i]))
    print("已导出"+str(key[i])+"-restock-report")
    print(keydic[str(key[i])])

for j in range(100,200):
    print(j)
