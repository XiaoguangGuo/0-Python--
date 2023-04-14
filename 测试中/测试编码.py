import chardet

with open(r'D:\运营\1数据源\TopsearchTerms\\CA_Top_Search_Terms_Simple_Month_2023_02_28.csv', 'rb') as f:
    result = chardet.detect(f.read())
    print(result['encoding'])
