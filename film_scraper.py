import requests
from bs4 import BeautifulSoup
import pandas as pd
import subprocess
import time


def is_file_in_use(file_path):
    """Checks if a file is currently opened."""
    try:
        with open(file_path, 'a'):
            pass
    except IOError:
        return True
    return False


def fetch_films(url, retries=3):
    """Fetches film data from the specified URL with retries upon failure."""
    for attempt in range(retries):
        response = requests.get(url)
        if response.status_code == 200:
            return response.text
        print(f"Error loading page: {response.status_code}. Attempt {attempt + 1} of {retries}.")
        time.sleep(1)  # Сделать паузу перед следующей попыткой
    return None


# Base URL for scraping
base_url = "https://rutube.ru/feeds/top/"
data = []
unique_titles = set()  # Предназначен для хранения уникальных названий фильмов
num_pages = 3  # Количество загружаемых страниц

for page in range(1, num_pages + 1):
    current_page_url = f"{base_url}?page={page}"
    print(f"Загружаемая страница {page} от {num_pages}...")
    html_content = fetch_films(current_page_url)

    if html_content:
        print(f"Parsing page {page}...")
        soup = BeautifulSoup(html_content, "html.parser")

        # Поиск всех элементов фильма на странице
        film_items = soup.find_all("div", class_="card-original-wrapper-module__cardOriginalWrap")

        for film in film_items:
            # Извлечение названия фильма
            film_name = film.find("a",
                                  class_="wdp-link-module__link wdp-card-description-module__title wdp-card-description-module__url wdp-card-description-module__videoTitle")
            film_name = film_name.text.strip() if film_name else "Название не найдено"

            # Длительность извлечения
            film_duration = film.find("span",
                                      class_="wdp-poster-badge-module__poster-badge wdp-card-video-options-module__duration")
            film_duration = film_duration.text.strip() if film_duration else "Не указано"

            # Извлекающий автор
            film_author = film.find("a",
                                    class_="wdp-link-module__link wdp-card-description-module__author wdp-card-description-module__url")
            film_author = film_author.text.strip() if film_author else "Не указано"

            # Добавлена дата извлечения
            film_added = film.find("div", class_="wdp-card-description-meta-info-module__metaInfoPublishDate")
            film_added = film_added.text.strip() if film_added else "Нет данных"

            # Количество извлекаемых просмотров
            film_views = film.find("div", class_="wdp-card-description-meta-info-module__metaInfoViewsCountNumber")
            film_views = film_views.text.replace('\xa0', '').strip() if film_views else "Нет данных"

            # Проверьте уникальность названия фильма
            if film_name in unique_titles:
                print(f"Фильм '{film_name}' уже добавлен. Пропускаем.")
            else:
                # Добавить данные о фильме в список
                data.append({
                    "Название": film_name,
                    "Автор": film_author,
                    "Добавлен": film_added,
                    "Количество просмотров": film_views,
                    "Продолжительность": film_duration
                })
                unique_titles.add(film_name)  # Добавьте название фильма к набору уникальных имен

        print(f"Страница {page} обработано успешно.")
        time.sleep(1)  # Пауза между страницами
    else:
        print(f"Не удалось загрузить страницу {page}.")

    # Удаление дубликатов на основе данных фильма
data = [dict(t) for t in {frozenset(d.items()) for d in data}]

# Сохраните данные в файл Excel
df = pd.DataFrame(data, columns=["Название", "Автор", "Добавлен", "Количество просмотров", "Продолжительность"])
excel_file_path = "films_data.xlsx"
df.to_excel(excel_file_path, index=False)

# Откройте файл Excel
if not is_file_in_use(excel_file_path):
    try:
        subprocess.Popen(['start', excel_file_path], shell=True)
        print(f"Файл '{excel_file_path}' был создан и открыт.")
    except Exception as e:
        print(f"Не удалось открыть файл: {e}")
else:
    print("Файл открыт в другом приложении. Не удается открыть.")

# Распечатайте собранные данные
if data:
    for movie in data:
        print(movie)
else:
    print("Не удалось собрать данные о фильме.")

print(f"Количество уникальных названий: {len(unique_titles)}")
