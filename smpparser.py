import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


def parse_releases_and_save(url, excel_file):
    try:
        # Отправляем HTTP-запрос к веб-сайту
        response = requests.get(url)

        # Проверяем успешность запроса
        if response.status_code == 200:
            # Используем BeautifulSoup для парсинга HTML
            soup = BeautifulSoup(response.text, 'html.parser')

            # Находим блок с релизами
            release_block = soup.find('div', class_='row download-list-widget')

            # Создаем новую книгу Excel и активный лист
            workbook = Workbook()
            sheet = workbook.active

            # Записываем заголовки в Excel
            headers = ['Release Number',
                       'Release Date',
                       'Release Download',
                       'Release Enhancements']
            sheet.append(headers)

            # Находим все элементы с информацией о релизе
            releases_list = release_block.find('ol',
                                               class_='list-row-container menu'
                                               ).find_all('li')

            # Записываем информацию о релизах в Excel
            for release_item in releases_list:
                release_number = release_item.find('span',
                                                   class_='release-number'
                                                   ).text.strip()
                release_date = release_item.find('span',
                                                 class_='release-date'
                                                 ).text.strip()
                release_download = release_item.find('span',
                                                     class_='release-download'
                                                     ).find('a')['href']
                release_enhancement = release_item.find('span',
                                                        class_='release-enhancements'
                                                        ).find('a')['href']

                sheet.append([release_number,
                              release_date, release_download,
                              release_enhancement])

            # Сохраняем книгу Excel
            workbook.save(excel_file)

            print(f"Данные успешно сохранены в файл: {excel_file}")

        else:
            print(f"Ошибка при запросе. Код ответа: {response.status_code}")

    except Exception as e:
        print(f"Произошла ошибка: {e}")


url = 'https://www.python.org/downloads/'
excel_file = 'releases_data.xlsx'
parse_releases_and_save(url, excel_file)
