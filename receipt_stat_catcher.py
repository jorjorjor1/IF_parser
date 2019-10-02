import xlwt
import re
import requests
import os
## -*- coding: utf-8 -*-

regexp = r'00\d{6}'
regexp2 = r'/https?:\/\/(?:[-\w]+\.)?([-\w]+)\.\w+(?:\.\w+)?\/?.*/i'
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet('list 1', cell_overwrite_ok=True)
number_list = []
path_to_dir = os.path.dirname(__file__)+ "/"

file_name = '30' #вставляй сюда имя файла, будет прочитано имя.txt и создан имя receipt.xls
file_name_txt = file_name+'.txt'
file_name_xls = path_to_dir+file_name+' receipts'+'.xls'

f = open(path_to_dir+file_name_txt, encoding='utf-8').read().split('\n')



#for i in f:
    #if len(i)==57:
        #sheet1.write(col+1,0, i[-8:])
        #col+=1

#book.save('simple.xls')
for i in f:
    #if "«В обработке»  «Подтверждено»" in i:
    if 'infoflot' in i and len(i)== 57:
        number = f.index(i)
        number_list.append(number)
        print (i)
print(number_list)

global j1
j1=1
global j2
j2=2
col = 1
row = 1
#print(f[number_list[j1]:number_list[j2]])
#block = f[number_list[j1]:number_list[j2]]

while True:
    try:
        block = f[number_list[j1]:number_list[j2]]
        for i in block:
            if '«Подтверждено»  «Продано у агентства»' in i and len(i)>=45:
                target = block.index(i)
                target = int(target)-2
                sheet1.write(col, 0, block[target][6:])
                sheet1.write(col, 1, block[0][-6:])
                sheet1.write(col, 2,block[0])
                sheet1.write(col, 3,i)
                col+=1
        j1+=1
        j2+=1

    except IndexError:
         break

book.save(file_name_xls)
    # pattern = re.compile('Продано')  # add another parameter `re.I` for case insensitive
    # match = pattern.search(i)
    # if match:
    #     print('True')
    # else:
    #     print('F')
    #     print(i)



# try:
#     for i in block:
#         print(i)
#         if "«В обработке»" in i:
#             print(f[number_list[j1]:number_list[j2]])
#             j1 += 1
#             j2 += 1
#         else:
#             print (f[number_list[j1]], 'blank')
#             j1 += 1
#             j2 += 1
# except IndexError:
#     pass



