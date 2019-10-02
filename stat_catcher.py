import xlwt
import re
import requests
## -*- coding: utf-8 -*-
import re
import xlrd
from xlutils.copy import copy
from xlrd import open_workbook
import os

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet('Обработка заявок', cell_overwrite_ok=True)
number_list = []
created = []
checked = []
kc_checked = []
kc_time = []
client_time = []
time_after_calc = []
hyperlinks = []
true_or_false = []
right_time = []
right_time_after_cal = []


june = [1, 2, 8, 9, 15, 16, 22, 23, 29, 30]
july = [6, 7, 13, 14, 20, 21, 27, 28]
august = [3, 4, 10, 11, 17, 18, 24, 25, 31]
sept = [1,8,15,22,29]
should_restart = True

now_month = sept  # month to scan

our_ms = ['ДмитрийФурманов', 'АлександрБенуа', 'Солнечныйгород', 'Н.А.Некрасов', 'Лебединоеозеро', 'Луннаясоната',
          'ВасилийЧапаев', 'АлександрНевский', 'КапитанПушкарев', 'Севернаясказка', 'Двестолицы',
          'ИльяМуромец', 'КосмонавтГагарин']
kc_workers = ['Бекбатырова Кристина Булатовна«Подтверждено»  «Ожидает подтверждения»',
              'Зайцева Виктория Антоновна«Подтверждено»  «Ожидает подтверждения»',
              'Татаринова Елизавета Юрьевна«Подтверждено»  «Ожидает подтверждения»'
              'Абрукова Дарья Александровна«Подтверждено»  «Ожидает подтверждения»',
              'Пахомова Галина Сергеевна«Подтверждено»  «Ожидает подтверждения»'
              ]
lebed = ['К.А.Тимирязев', 'МихаилТанич', 'Бородино']
vdh = ['АлександрСуворов', 'МихаилФрунзе', 'КонстантинКоротков', 'СеменБуденный', 'СергейКучкин']
volga_ples = ['А.И.Герцен', 'ДмитрийПожарский']
mtf = ['СергейОбразцов', 'СергейЕсенин', 'АлександрРадищев', 'НиколайКарамзин', 'МихаилБулгаков', "И.А.Крылов"]
ortodox = []
valaamskii = ['Валаамскийэкспромт']
peterb_kruizi = ['ЛеонидКрасин']
rechflot = ['ГригорийПирогов']
knyaz = ['КнязьВладимир']
rusich = ['РоднаяРусь', 'РусьВеликая']
ros_voyaj = ['АлексейТолстой']
sputnik = ['Ф.И.Панферов', 'ФедорДостоевский', 'ХирургРазумовский', 'ВалерийЧкалов']
cezar = ['Президент']
volga_volga = ['ПавелБажов', 'ВладимирМаяковский', 'МихаилКутузов', 'АлександрФадеев']
volga_dream = ['ВолгаДрим']
kama_travel = ['КозьмаМинин', 'Н.В.Гоголь']
east_travel = ['АлександрВеликий', 'Империя']

# Белый Лебедь - с 10 до 19
# ВодоходЪ - с 8 до 20
# ВолгаПлёс - с 8 до 18, хз тут
# МосТурФлот - на сайте с 9 до 21, но агентский в 18 точно уходит, значит поставим с 9 до 18
# Ортодокс - с 8 до 20
# Паломники - с 8 до 20, не нужны подтверждения
# Петербургские круизы - с 8 до 20, не нужны подтверждения, поэтому всё ок
# РечФлот - хаха, с 10 до 19
# Черноморские круизы - подтверждения нужны только в сложных ситуациях, можно ставить с 8 до 20
# Экспресс-Тур - пермь с 9 до 18, значит по нашему с 8 (как отдел открывается) до 16
# КамаТревел -
# ВолгаВолга -
# ВолгаДрим -
# РосВояж - хз, ставь с 8 до 18
# Спутник-Гермес - с 10 до 19 по идее в Самаре, у нас значит с 9 до 18
# Цезарь - с 10 до 18, пусть будет

reg2 = re.compile(r'\w+\s+\w+\s+\w+\SПодтверждено\S\s+\SОжидает')
reg3 = re.compile(r'\w+\s+\w+\s+\w+\SВ\sобработке\S\s+\SПодтверждено\S')
reg5 = re.compile(r'\w+\s\S\w+.+Менеджер взял автоматически подтвержденную заявку на проверку')
file_name = '30'  # вставляй сюда имя файла, будет прочитано имя.txt и создан имя.xls
path_to_dir = os.path.dirname(__file__)+ "/"
file_name_txt = file_name + '.txt'
file_name_xls = path_to_dir+file_name + '.xls'
print(path_to_dir)
print(path_to_dir+file_name_txt)
f = open(path_to_dir+file_name_txt, encoding='utf-8').read().split('\n')


for i in f:
    if 'infoflot' in i and len(i) == 57:
        number = f.index(i)
        number_list.append(number)

global j1
j1 = 1
global j2
j2 = 2
col = 1
row = 1

def first_confirm_time_search(modifier):
    global col
    col-=int(modifier)
    try:
        if int(bricks[-1][1][0]) == int(checked[0][0]):
            time_spent = ((int(checked[0][2]) * 60 + int(checked[0][3])) - (
                    int(created[0][2]) * 60 + int(created[0][3])))
    except  NameError:
        pass
    if ms_name in our_ms:
        sheet1.write(col, 2, '1')
        if int(created[0][0]) not in now_month:
            recurs(1, 3, 8, 20)
        else:
            # print(created[0][0])
            recurs(0, 3, 8, 20)

    else:
        sheet1.write(col, 5, '1')
        if int(created[0][0]) not in now_month:
            if ms_name in lebed or ms_name in rechflot:
                # print('lebed or rechflot')
                recurs(1, 6, 10, 19)
            elif ms_name in vdh or ms_name in valaamskii or ms_name in ortodox or ms_name in peterb_kruizi or ms_name in knyaz or ms_name in sputnik:
                # print('s 8 do 20')
                recurs(1, 6, 8, 20)
            elif ms_name in volga_ples or ms_name in rusich or ms_name in ros_voyaj:
                # print('s 8 do 16 PLES RUSICH ROSVOJAJ')
                recurs(1, 6, 8, 16)
            elif ms_name in mtf or ms_name in volga_volga:
                # print('s 9 do 18 MTF volgavolga')
                recurs(1, 6, 9, 18)
            elif ms_name in cezar or ms_name in volga_dream:
                # print('s 10 do 18 cezar vogla_dream')
                recurs(1, 6, 10, 18)
            elif ms_name in kama_travel:
                # print('s 8 do 17 kama_travel')
                recurs(1, 6, 8, 17)
            elif ms_name in east_travel:
                # print('s 8 do 13 east_land')
                recurs(1, 6, 8, 13)

            else:
                sheet1.write(col, 6, 'Ошибка! нет в списке тх')

        else:
            # print(created[0][0])
            if ms_name in lebed or ms_name in rechflot:
                # print('lebed or rechflot')
                recurs(0, 6, 10, 19)
            elif ms_name in vdh or ms_name in valaamskii or ms_name in ortodox or ms_name in peterb_kruizi or ms_name in knyaz:
                # print('s 8 do 20')
                recurs(0, 6, 8, 20)
            elif ms_name in volga_ples or ms_name in rusich or ms_name in ros_voyaj:
                # print('s 8 do 16 PLES RUSICH ROSVOJAJ')
                recurs(0, 6, 8, 16)
            elif ms_name in mtf or ms_name in sputnik or ms_name in volga_volga:
                # print('s 9 do 18 MTF SPUTNIK volgavolga')
                recurs(0, 6, 9, 18)
            elif ms_name in cezar or ms_name in volga_dream:
                # print('s 10 do 18 cezar vogla_dream')
                recurs(0, 6, 10, 18)
            elif ms_name in kama_travel:
                # print('s 8 do 17 kama_travel')
                recurs(0, 6, 8, 17)
            elif ms_name in east_travel:
                # print('s 8 do 13 east_land')
                recurs(0, 6, 8, 13)
            else:
                sheet1.write(col, 6, 'Ошибка! нет в списке тх')

    link = block[-1], block[-1][-6:]
    # hyperlinks.append(link)
    sheet1.write(col, 10, ms_name)
    sheet1.write(col, 14, i)  # кто подтвердил
    sheet1.write(col, 12, block[5][7:])  # 12 ячейку не трогать, код получает резултат из нее ><
    sheet1.write(col, 0, block[x + 2][6:])
    sheet1.write(col, 9, block[3])
    col += 1


def recurs(weekday_modifier, col_num, start_work, stop_work):
    try:
        if int(created[0][0]) == int(checked[0][0]):
            time_spent = ((int(checked[0][2]) * 60 + int(checked[0][3])) - (
                    int(created[0][2]) * 60 + int(created[0][3])))
    except  NameError:
        pass
    if int(created[0][0]) == int(checked[0][0]):
        if int(created[0][2]) < int(8):
            time_spent = ((int(checked[0][2]) * 60 + int(checked[0][3])) - (
                    start_work * 60))
            if time_spent <= int(0):
                time_spent = int(3)
        sheet1.write(col, col_num, time_spent)
    else:
        if stop_work <= int(created[0][2]) <= 23 or 0 <= int(created[0][2]) < start_work:
            time_spent = ((int(checked[0][2]) * 60 + int(checked[0][3])) - (
                    start_work * 60))
            if time_spent <= int(0):
                time_spent = int(3)
            sheet1.write(col, col_num, time_spent)
        else:
            time_spent = ((stop_work * 60 - (int(created[0][2]) * 60 + int(created[0][3]))) * weekday_modifier + (
                    (int(checked[0][2]) * 60 + int(checked[0][3])) - (start_work * 60)))
            if time_spent <= int(0):
                sheet1.write(col, col_num, 'Нужна проверка, скорее всего, тут должна стоять 3')
            else:
                sheet1.write(col, col_num, time_spent)



while True:
    try:
        block = f[number_list[j1]:number_list[j2]]
        block.reverse()
        #qq = re.findall(reg3, str(block))
        bricks = []
        for i in block:
            jj = re.search(reg2, i)
            if jj:
                true_or_false.append(block.index(i))
                print(jj)


            if '«В обработке»  «Подтверждено»' in i:
                checked.clear()
                kc_checked.clear()
                created.clear()
                ms_name = block[-2].replace(" ", "") #Имя тх первое после ссылки
                ms_link = block[-1] #Ссылка на тх первая
                when_created = re.findall(r'\d+', block[5]) #Время создания
                created.append(re.findall(r'\d+', block[5]))
                x = (int(block.index(i)))
                who = i  #Кто подтвердил заявку
                time_confirmation = re.findall(r'\d+', block[x + 2]) #Время подтверждения
                checked.append(re.findall(r'\d+', block[x + 2]))
                bricks.append([ms_link, when_created, time_confirmation, who])
                if len(bricks)>1:   #если "подтверждений" в заявке больше, чем 1, т.е. ищем самое первое
                    first_confirm_time_search(1)
                    #continue

                else:
                    first_confirm_time_search(0)
                    link = block[-1], block[-1][-6:]
                    hyperlinks.append(link)
                    #continue
            #print(i, jj,true_or_false)
                if true_or_false!=[]:
                    print("НАЙДЕНО ХЕЛЛОУ")
                    col -= 1
                    sheet1.write(col, 8, "Заявка от КЦ")
                    KC_confirm_date = str(block[int(true_or_false[0]) + 2][6:])
                    print('conf date:', KC_confirm_date)
                    kc_checked.append(re.findall(r'\d+', KC_confirm_date))
                    sheet1.write(col, 13, kc_checked[0][0] + '-' + kc_checked[0][2] + ':' + kc_checked[0][
                        3])  # 13 ячейку не трогать, код получает резултат из нее ><
                    # print(kc_checked[0][0] + "-" + kc_checked[0][2] + ':' + kc_checked[0][3])

                    col += 1
            else:
                continue

        # break
            checked.clear()
            kc_checked.clear()
            created.clear()

        true_or_false.clear()
        j1 += 1
        j2 += 1
        true_or_false.clear()
    except IndexError:
        break
# row_number = 0
# j1 = 1
# j2 = 2
# while True:
#     try:
#         block = f[number_list[j1]:number_list[j2]]
#         for i in block:
#             jj = re.search(reg2, i)
#             print(jj)
#             if '«В обработке»  «Подтверждено»' in i:
#                 true_or_false.append('true')
#             if jj and true_or_false != []:
#                 print(i)
#                 KC_confirm_date = str(block[block.index(i) + 2])[6:]
#                 kc_checked.append(re.findall(r'\d+', KC_confirm_date))
#                 sheet1.write(row_number, 13, kc_checked[0][0] + '-' + kc_checked[0][2] + ':' + kc_checked[0][
#                     3])  # 13 ячейку не трогать, код получает резултат из нее ><
#                 # print(kc_checked[0][0] + "-" + kc_checked[0][2] + ':' + kc_checked[0][3])
#                 row_number+=1
#                 true_or_false.clear()
#         j1 += 1
#         j2 += 1
#         true_or_false.clear()
#         col += 1
#     except IndexError:
#         break


book.save(file_name_xls)

rb = xlrd.open_workbook(file_name_xls, formatting_info=True)

sheet = rb.sheet_by_index(0)
for i in range(sheet.nrows):
    x = re.findall(r'\d+', sheet.cell_value(i, 13))
    kc_time.append(x)
    y = re.findall(r'\d+', sheet.cell_value(i, 12))
    client_time.append(y)
    print('y:',y, 'x:',x)

for j in kc_time:
    if j != []:
        kc_index = kc_time.index(j)
        client_index = kc_index  # kc index = [[12][12][01]] - 12 июня в 12:01, client_index = [[12][2019][12][01]0
        print(client_index)
        try:
            time_spent2 = ((int(kc_time[kc_index][1]) * 60 + int(kc_time[kc_index][2])) - (
                (int(client_time[client_index][2]) * 60 + int(client_time[client_index][3]))))

            if int(kc_time[kc_index][0]) == int(client_time[client_index][0]):
                time_after_calc.append(time_spent2)
            else:
                time_after_calc.append('другой день')
        except IndexError:
            pass
    elif j == []:
        time_after_calc.append(0)

print(time_after_calc)
row_num = 0
for time in time_after_calc:
    right_time.append([time, sheet.cell_value(row_num, 3)])
    row_num += 1
for group in right_time:
    print(group[1])
    if group[1] != '' and group[1] != 'Какая-то фигня, нужна проверка' and group[0] != 'другой день' and group[1] != 'Нужна проверка, скорее всего, тут должна стоять 3':
        right_time_after_cal.append(int((group[1]) - group[0]))
    else:
        right_time_after_cal.append(group[1])

# print(right_time_after_cal)


# rb = open_workbook(file_name_xls, formatting_info=True)
wb = copy(rb)

s = wb.get_sheet(0)
link_num = 0
right_link_num = 0
try:
    for element in time_after_calc:
        s.write(link_num + 1, 1,
                xlwt.Formula('HYPERLINK("%s";"%s")' % (hyperlinks[link_num][0], hyperlinks[link_num][1])))
        s.write(link_num, 9, element)
        link_num += 1

except IndexError:
    pass

for elem in right_time_after_cal:
    if elem == "":
        right_link_num += 1
    else:
        s.write(right_link_num, 3, elem)
        right_link_num += 1

wb.save(file_name_xls)

print('2nd page start')
sheet2 = wb.add_sheet('Проверка заявок')
j1 = 1
j2 = 2
col = 1
row = 1


while True:
    try:
        block = f[number_list[j1]:number_list[j2]]
        block.reverse()
        for i in block:
            jj = re.search(reg5,i)
            if jj:
                true_or_false.append('true')
                checked.clear()
                kc_checked.clear()
                created.clear()
                # and len(i)>=45:
                ms_name = block[1].replace(" ", "")
                print ('link:', block[-1])
                x = (int(block.index(i)))
                time_v_obrabotke = block[x+2][7:]
                print(i, time_v_obrabotke)
                sheet2.write(col, 0, xlwt.Formula('HYPERLINK("%s";"%s")' % (block[-1], block[-1][-6:])))
                sheet2.write(col, 1, time_v_obrabotke)
                sheet2.write(col, 2, i)
                col+=1
                # print('MS', ms_name)
                # print('создана в', block[-6])
                #created.append(re.findall(r'\d+', block[-6]))
                # print(created)
                #x = (int(block.index(i)))
                # print('время подтверждения', block[x-2])
                # print('кто создал', i)
                #checked.append(re.findall(r'\d+', block[x - 2]))
                #print(checked)
        j1 += 1
        j2 += 1
    except IndexError:
        break

wb.save(file_name_xls)