for letter in 'Python':
    if letter == 'h':
        continue #此处跳出for枚举'h'的那一次循环
        print('当前字母 :', letter)


for letter in 'Python':
    if letter == 'h':
        break  #此处跳出for枚举'h'的那一次循环
        print('当前字母 :', letter)

for i in range(10):
    print("-----%d-----" %i)
for j in range(10):
    if j > 5 and j <= 8:
        print("我是continue特殊")
        continue
    print(j)

for i in range(10):
    print("-----%d-----" %i)
for j in range(10):
    if j > 5 and j <= 8:
        print("我是continue特殊")
        break
    print(j)
