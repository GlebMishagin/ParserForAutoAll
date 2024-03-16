import openpyxl
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import pandas as pd

def записать_в_excel(файл, данные):
    workbook = openpyxl.load_workbook(файл)
    sheet = workbook.active
    последняя_строка = sheet.max_row
    новая_строка = последняя_строка + 1
    sheet.append(данные)
    workbook.save(файл)

excel_file_path = 'таблицавх.xlsx'
sheet_name = 'Лист 1'
column_name = 'Артикул'

df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_art_list = df[column_name].tolist()

for art1 in column_art_list:
    url = f"https://www.avtoall.ru/search/?GlobalFilterForm%5Bnamearticle%5D={art1}"
    print(url)
    keywords = ['HYUNDAI', 'KIA', 'DAEWOO', 'GENERAL MOTORS', 'CHEVROLET', 'SSANGYONG']

    response = requests.get(url)
    volume = 0
    mass = 0

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        image_div = soup.find_all('div', class_='image')
        if image_div:
            for image in image_div:
                try:
                    link_tag = image.find('a', href=True)
                    if link_tag:
                        absolute_image_url = urljoin(url, link_tag['href'])
                        test = absolute_image_url
                        modified = test.replace("_", " ")
                        if any(keyword.lower() in modified.lower() for keyword in keywords):
                            print(f"Ссылка на изображение: {absolute_image_url}")
                            response1 = requests.get(absolute_image_url)
                            if response1.status_code == 200:
                                soup = BeautifulSoup(response1.text, 'html.parser')
                                data_span = soup.find('b', itemprop='width')
                                if data_span:
                                    a = float(data_span.get_text(strip=True))
                                    print('Ширина, м:', a)
                                    data_span = soup.find('b', itemprop='height')
                                    b = float(data_span.get_text(strip=True))
                                    print('Высота, м:', b)
                                    data_span = soup.find('b', itemprop='weight')
                                    mass = float(data_span.get_text(strip=True)) if data_span else 0
                                    print('Вес, м:', mass)
                                    section_data_div = soup.find('div', class_='section-data parametrs flex')
                                    if section_data_div:
                                        b_tags = section_data_div.find_all('b')
                                        desired_index = 6
                                        if 0 <= desired_index < len(b_tags):
                                            desired_b_tag = b_tags[desired_index]
                                            c = float(desired_b_tag.get_text(strip=True))
                                            volume = a * b * c
                                            # print('Длинна :', c)
                                            # записать_в_excel('таблицавых.xlsx', [art1, volume, mass])
                                            break
                except AttributeError:
                    pass



    if volume !=0 and mass !=0:
        print(volume, mass)
        записать_в_excel('таблицавых.xlsx', [art1, volume, mass])
    if volume ==0 and mass ==0:
        volume = 'NAN'
        mass = 'NAN'
        print(volume, mass)
        записать_в_excel('таблицавых.xlsx', [art1, volume, mass])
    # else:
    #     print(f"Ошибка при обращении к {url}")
