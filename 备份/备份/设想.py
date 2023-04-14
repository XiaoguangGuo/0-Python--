##将当前库存放到对应的文件中
### 首先查找文件名中带有i且不带New的文件
### 打开这个文件
###写入US当前库存 

## 循环各个国家 ( MX,CA,New-US,New-CA,New-MX)



 # 源文件路径
    source = "C:\\Users\\Administrator\\Desktop\\ceshi\\csv文件"

    # 目标文件路径
    ob = "C:\\Users\\Administrator\\Desktop\\ceshi\\xlsx文件"

    # 将源文件路径里面的文件转换成列表file_list
    file_list = [source + '\\' + i for i in read_path(source)]


#寻找文件名带有listoffilename的
# 如果带有New 则写入
# 否则写入
