from bs4 import BeautifulSoup
import re
import requests
from requests import request
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import time
import os


options = webdriver.ChromeOptions()  # единые настройки для каждой копии гугла
options.add_argument('user-data-dir=C:\project1')
# options.add_argument
driver = webdriver.Chrome("/usr/bin/chromedriver", chrome_options=options)
office_list = {'kazan': 'searchQuery=+офис=%27Казань,+м.+Площадь+Тукая,+ул.+Спартаковская,+д.+2%27',
               'msk_rech': 'searchQuery=офис=%27Москва,+м.+Речной+вокзал,+Ленинградское+шоссе,+57,+корпус+1+(СРВ)%27',
               'msk_sokol': 'searchQuery=офис=%27Москва,+м.+Сокол,+Чапаевский+пер.,+д.+6%27',
               'msk_myas': 'searchQuery=офис=%27Москва,+м.+Чистые+пруды,+ул.+Мясницкая,+д.+40,+стр.1%27',
               'nnov': 'searchQuery=офис=%27Нижний+Новгород,+м.+Горьковская,+ул.+Костина,+д.+3%27',
               'rostov': 'searchQuery=офис=%27Ростов-на-Дону,+ул.+Города+Волос,+д.+6%27',
               'sam_rech': 'searchQuery=офис=%27Самара,+ул.+Максима+Горького,+д.+82%27',
               'sam_rab': 'searchQuery=офис=%27Самара,+ул.+Рабочая,+д.+15%27',
               'sam_sovarm': 'searchQuery=офис=%27Самара,+ул.+Советской+Армии,+д.+180%27',
               'lomo': 'searchQuery=офис=%27Санкт-Петербург,+м.+Ломоносовская,+ул.+Полярников,+д.+6,+литера+А,+офис+5+%27',
               'efimova': 'searchQuery=офис=%27Санкт-Петербург,+м.+Сенная+площадь,+ул.+Ефимова,+д.+3А%27'}
office_list_keys = (list(office_list.keys()))
print(office_list_keys)
list_num = int(0)

office_num = int(0)
regexp =r'00\d{6}'
regexp2 = r'[^a-z]\d{2}\.\d{2}\.\d{2}'
timeout = 30
num = {}



x = os.path.abspath(".") + "/exercises/booking_parse/" #исправить адрес, это говно из жёпы
start_date = int(30)   #с этого числа включ будет поиск
finish_date = int(1)  #по это число включ будет поиск
now_month = int(9)    #месяц нужных заявок (если нужны заявки за инюнь, то пишем 6)
file_name = str(start_date) + str('.txt')


def search():
    def parse_html():  # первичный парсинг общей страницы с заявками
        try:
            element_present = EC.presence_of_element_located((By.CLASS_NAME, 'Actions'))
            WebDriverWait(driver, timeout).until(element_present)
            soup = BeautifulSoup(driver.page_source, 'lxml')
            tr = soup.find('table', class_='table table-striped table-bordered table-hover')
            tbody = tr.find('tbody')
            tr = tbody.find_all('tr')
            all_numbers = soup.find_all('strong', text=re.compile(regexp))
            x = 0
            for i in range(0, 20):
                reserv_num = all_numbers[x].find(text=re.compile(regexp))
                reserv_time = tr[x].find_all('div')
                reserv_time2 = reserv_time[2].find('small').get_text()
                reserv_day = reserv_time2[0:2]
                reserv_month = reserv_time2[3:5]
                y = reserv_num.find('strong')
                reserv_num_str = str(reserv_num)
                num[reserv_num_str] = reserv_day, reserv_month
                print(x)
                x += 1
            print(num)

        except TimeoutException:
            print('timeout_exception')
            driver.get(driver.current_url)
        except IndexError:
            pass
    def resv_enter():
        global num
        try:
            for key in num.copy():
                print(key, num[key])
                if int(num[key][1]) > now_month:
                    num.pop(key)
                    "Идем до сл м"# 6<=6
                elif int(num[key][1]) == now_month:
                    if int(num[key][0]) > int(start_date):
                        print('идем до нужной даты...')
                        num.pop(key)
                    elif int(num[key][0]) <= int(start_date):
                        driver.get('https://booking.infoflot.com/ACP/Requests/Manage/{}'.format(key))
                        element_present = EC.presence_of_element_located((By.CLASS_NAME, 'chosen-single'))
                        WebDriverWait(driver, timeout).until(element_present)
                        stop_office()
                        num.pop(key)
                        print(num)
                elif int(num[key][1]) < now_month:  # месяц в заказе>этот месяц
                    global office_num
                    office_num += 1
                    choose_an_office(office_num)

        except TimeoutException:
            print('timeout_exception')
            resv_enter()
        except KeyError:
            print('Ошибка идем на след стр')
            change_page()


    def stop_office():
        global office_num
        element_present = EC.presence_of_element_located((By.CLASS_NAME, 'chosen-single'))
        WebDriverWait(driver, timeout).until(element_present)
        soup = BeautifulSoup(driver.page_source, 'lxml')
        body = soup.find('body')
        content = body.find(id='Content')
        span9 = content.find(class_='span9')
        form_horizontal = span9.find(class_='form-horizontal')
        request_creation_date = form_horizontal.find(class_='request-creation-date')
        req_date = request_creation_date.get_text()[0:2]
        req_month = request_creation_date.get_text()[3:5]
        print(
        'start date', start_date, 'req_date', req_date, req_month, 'finish date', finish_date)
        #if int(req_month) < int(now_month):
            #if int(req_date) == int(30) or int(req_date) == int(31):  # ВАЖНО ВЫРУБАЙ ЕСЛИ В КОНЦЕ МЕСЯЦА
                #global office_num
                #office_num += 1
                #choose_an_office(office_num)
        if int(req_date) >= int(finish_date):
            parse_page()
        elif int(req_date) < int(finish_date):
            office_num += 1
            choose_an_office(office_num)
        else:
            pass

        #else:
            #office_num += 1
            #choose_an_office(office_num)


    def parse_page():  # парсинг страницы заявки
        try:
            element_present = EC.presence_of_element_located((By.CLASS_NAME, 'chosen-single'))
            WebDriverWait(driver, timeout).until(element_present)
            soup = BeautifulSoup(driver.page_source, 'lxml')
            body = soup.find('body')
            content = body.find(id='Content')
            span9 = content.find(class_='span9')
            form_horizontal = span9.find(class_='form-horizontal')
            hist = form_horizontal.contents[21].get_text()
            tab_content = content.find(class_='tab-content')
            controls = tab_content.find(id='section-RequestsApp-Data-Cruise')
            control_group = controls.findAll('a', attrs={'class': 'chosen-single'})
            ms_name = (control_group[1].get_text())
            my_file = open(x+file_name, 'a', encoding='utf-8')
            my_file.write(driver.current_url)
            my_file.write(ms_name)
            my_file.write(hist)
            print(driver.current_url, ' recorded')
        except AttributeError:
            print('timeout_exception')
            time.sleep(5)
            driver.get(driver.current_url)

    parse_html()
    resv_enter()


def choose_an_office(office_num):
    global list_num
    list_num = 1
    driver.get(
        'https://booking.infoflot.com/ACP/Requests/List/1?&{}'.format(office_list.get(office_list_keys[office_num])))
    change_page()


def change_page():
    global list_num
    for l in range(1000):
        try:
            driver.get(
                'https://booking.infoflot.com/ACP/Requests/List/{}?&{}'.format(list_num, office_list.get(
                    office_list_keys[office_num])))
            list_num = list_num + 1
            num.clear()
            element_present = EC.presence_of_element_located((By.CLASS_NAME, 'Actions'))
            WebDriverWait(driver, timeout).until(element_present)
        except:
            print('timeout_exception')
            driver.get(driver.current_url)
        print(driver.current_url)
        search()
    driver.quit()


choose_an_office(office_num)
