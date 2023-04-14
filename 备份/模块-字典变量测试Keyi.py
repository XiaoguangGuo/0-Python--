
key =['US','CA','MX','UK','IT','DE','JP']
keydic = { "US" : '333' , 'CA' : '666' , 'MX' : '999','UK':'111','IT':'222','DE':'888','JP':'1.5' }
print (key)
print (len(key))
print(keydic["US"])

 
for i in range(len(key)):
    print(key[i])
    print(str(key[i]))
    print("已导出"+str(key[i])+"-restock-report")
    print('key[i]')
    print(keydic[str(key[i])])

for j in range(100,200):
    print(j)
a="ere"
print(str(a))
print(a)

for letter in 'Python':
    if letter == 'h':
        continue #此处跳出for枚举'h'的那一次循环
        print('当前字母 :', letter)


for letter in 'Python':
    if letter == 'h':
        break  #此处跳出for枚举'h'的那一次循环
        print('当前字母 :', letter)
print('aaaaa')
print("aaaaa")
