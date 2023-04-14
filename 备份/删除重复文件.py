

# -*- coding:utf-8 –*-

import os           # 加载文件治理库


import os, sys
from stat import *

path = r'D:\\SearchTermreport历史\\'

for i in range(len(os.listdir(path))):
    path2=os.listdir(path)
    print(path2)
    print(path2[i])
    filename=path2[i]
    
    mode = os.stat(path+filename).st_mode
    
    if S_ISDIR(mode)!=True:
        
        filenamecut=filename.encode("utf-8").decode("utf-8")
        filenamecut=filenamecut[0:30]
        filesize = os.stat(path+filename).st_size
        print(filename)
        

        newpathdirlist=path2.remove(i)
        print(newpathdirlist)
        
        if newpathdirlist:

            for leftfilename in newpathdirlist:
                print(leftfilename)
                newfilesize = os.stat(leftfilename).st_size
                leftfilenamecut=leftfilename.decode("utf-8")
                leftfilenamecut=leftfilenamecut[0:30]
        
   
        
                if newfilesize == filesize and leftfilenamecut==filenamecut :
                    os.remove(leftfilename) 
   
