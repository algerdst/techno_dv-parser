# -*- coding: utf-8 -*-
import random
import csv
import openpyxl
from openpyxl.styles import Alignment
from bs4 import BeautifulSoup as bs
import time
import cloudscraper
import logging
import base64
import smtplib
import os
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime



# Создание экземпляра скрапера
scraper = cloudscraper.create_scraper()

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Функция для отправки email
with open('email.txt', 'r') as f: email_address = f.readline()


def send_email():
    getter = email_address
    sender = "MyAppsDjangoSend@gmail.com"
    password = "rmddghigmazqocld"
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    try:
        server.login(sender, password)
        msg = MIMEMultipart()
        part = MIMEBase('app', 'csv')
        file = open('my_book.csv', 'rb').read()
        bs = base64.encodebytes(file).decode()
        part.set_payload(bs)
        part.add_header('Content-Transfer-Encoding', 'base64')
        part.add_header('Content-Disposition', 'attachment', filename='my.book.csv')
        msg.attach(part)
        server.sendmail(sender, getter, msg.as_string())

        print("The message was sent successfully!")
        with open('logging.txt', 'a') as logging_file:
            logging_file.write(f'{str(datetime.now())[:19]}Файл был отправлен на почту')
    except Exception as _ex:
        return f"{_ex}\nCheck your login or password please!"


static_text = """
==================================================================

Наша компания предоставляет полный спектр услуг по автоматизации Вашего бизнеса:

- консультации по выбору оборудования и программного обеспечения и их демонстрация;

- поставка оборудования и программного обеспечения;

- установка, настройка, сопровождение и обновление оборудования и программного обеспечения;

- обучение пользователей и ИТ-специалистов.

==================================================================

Наша компания имеет опыт проектов автоматизации с 2002 года.

Также для Вас мы оказываем и другие услуги:

- продажу, настройку  и подключение любого кассового и торгового оборудования;

- подключение к оператору фискальных данных (ОФД);

- регистрацию кассового оборудования в налоговых органах;

- продажу, установку и доработку программных продуктов 1С;

- И МНОГОЕ ДРУГОЕ!

==================================================================

Мы всегда готовы проконсультировать и подобрать оптимальное решение для Вашего бизнеса!

==================================================================

Данное предложение НЕ ЯВЛЯЕТСЯ публичной офертой!

Характеристики, цена, информация о наличии товара МОГУТ ИЗМЕНЯТЬСЯ!

Всю интересующую Вас информацию НЕОБХОДИМО УТОЧНЯТЬ по номерам телефонов компании и на сайте компании!

==================================================================
"""


# Функция для создания рандомного юзер агента
def set_random_user_agent():
    # random_user_agent = UserAgent().random
    scraper.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.5993.2470 YaBrowser/23.11.0.2470 Yowser/2.5 Safari/537.36'})


# Создание книги Excel

fields = ['Наименование', 'Артикул', 'Цена', 'Описание', 'Изображение', 'Наличие', 'Срок поставки']

items_count = 0

# Сбор всех ссылок на категории в словарь
links_dict = {'Кассы самообслуживания': 'https://tehno-dv.ru/vse-tovary/kassy-samoobsluzhivaniya/'}
link = 'https://tehno-dv.ru/vse-tovary/kassy-samoobsluzhivaniya/'
set_random_user_agent()
response = scraper.get(link)
soup = bs(response.text, 'lxml')
links = soup.findAll('li', class_='ut2-item')
for link in links:
    try:
        key = link.find('a').text.replace('\n', '').lower()
        cat_link = link.find('a')['href']
        links_dict[key] = cat_link
    except:
        continue

# Исключение из поиска, ссылок записаных в файле excludes
with open('excludes.txt', 'r', encoding='utf-8') as file:
    for exclude in file:
        exclude = exclude.replace('\n', '')
        if exclude in links_dict:
            del links_dict[exclude]
with open('my_book.csv', 'w', newline='', encoding='utf-8-sig') as file:
    writer = csv.writer(file, delimiter=';')
    writer.writerow(fields)
    for link in links_dict:
        with open('logging.txt', 'a') as logging_file:
            category_items = 0
            link = links_dict[link]
            set_random_user_agent()
            response = scraper.get(link)
            page = 1
            soup = bs(response.text, 'lxml')
            category = soup.find('h1', class_='ty-mainbox-title').text.replace('\n', ' ')
            print(f"Поиск по категории {category}")
            items = []
            try:
                while True:
                    set_random_user_agent()
                    if page == 1:
                        url = link
                    else:
                        url = f"{link}page-{page}/"
                    response = scraper.get(url)
                    response.raise_for_status()  # Проверка ошибки HTTP
                    soup = bs(response.text, 'lxml')
                    blocks = soup.findAll('div', class_='ty-product-list')
                    for block in blocks:
                        # Попытка найти у товара блок с фото
                        try:
                            image_block = block.find('div', class_='image-reload')
                            image = image_block.find('img')['src']
                        except:
                            continue
                        name = block.findNext('div', class_='ut2-pl__item-name').text
                        try:
                            article = block.find('div', class_='ty-sku-item').text.replace('КОД:', '').strip('\n')
                        except:
                            logging_file.write(f"{str(datetime.now())[:19]} Отсутствует артикул у товара {name} из категории {category}")
                            article='Артикул отсутствует'
                        in_stock = block.find('div', class_='product-list-field').text.replace('Доступность:',
                                                                                               '').replace(
                            'шт.', '').replace('\n', '').strip()
                        price = float(''.join(block.find('span', 'ty-price').text.replace('₽', '').split()))
                        price -= price / 100
                        # Попытка привести значение in_stock к int
                        try:
                            in_stock = int(in_stock)
                            if in_stock > 0:
                                in_stock = 'В наличии'
                                delivery_time = ''
                        except:
                            if in_stock == 'Нет в наличии' or in_stock=='предзаказ':
                                in_stock='Под заказ'
                                delivery_time = "От 5 до 20 дней"
                        item_link = block.find('a')['href']
                        item_response = scraper.get(item_link)
                        img = bs(item_response.text, 'lxml').find('div', class_='ut2-pb__img').find('a')['href']
                        description = str(
                            bs(item_response.text, 'lxml').find('div', id='content_description')) + '\n' + static_text
                        if name not in items:
                            items.append(name)
                        # Запись новой строки в Excel
                        fields = [name, article, price, description, img, in_stock, delivery_time]
                        writer.writerow(fields)
                        items_count += 1
                        category_items += 1
                        print(f"Собрано всего товаров {items_count}")
                    logging.info(f"СОБРАНО {page} СТРАНИЦ С ТОВАРАМИ  В КАТЕГОРИИ {category}")
                    logging_file.write(
                        f"{str(datetime.now())[:19]} СОБРАНО {page} СТРАНИЦ С ТОВАРАМИ  В КАТЕГОРИИ {category}" + '\n' +
                        f"{str(datetime.now())[:19]} Собрано {category_items} товаров из категории {category}" + '\n' + f"{str(datetime.now())[:19]} Собрано всего товаров {items_count}" + '\n')
                    page += 1
                    # time.sleep(random.randint(1, 3))
            except Exception as ex:
                print(f" {str(datetime.now())[:19]} Собрано {category_items} товаров из категории {category}")
                with open('fail_url.txt', 'a') as fail_file:
                    fail_file.write(url + str(ex) + '\n')


send_email()
