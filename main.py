from collections import defaultdict
from datetime import datetime
import json
from pprint import pprint
from jinja2 import Environment, FileSystemLoader
import pandas as pd


def get_year_word(age):
    """
    Возвращает правильное склонение слова 'год' для настоящего времени.

    Args:
        age (int): Возраст в годах

    Returns:
        str: 'год', 'года' или 'лет'
    """
    last_two_digits = age % 100
    if 11 <= last_two_digits <= 19:
        return 'лет'
    last_digit = age % 10
    if last_digit == 1:
        return 'год'
    elif 2 <= last_digit <= 4:
        return 'года'
    else:
        return 'лет'


def clean_data(value):
    """
    Заменяет NaN значения на пустые строки.

    Args:
        value: Любое значение, возможно NaN

    Returns:
        Оригинальное значение или пустая строка если значение NaN
    """
    return value if pd.notna(value) else ''


def find_cheapest_wine(wines_data):
    """
    Находит самое дешёвое вино среди всех.

    Args:
        wines_data (list): Список словарей с данными о винах

    Returns:
        dict or None: Данные самого дешёвого вина или None если не найдено
    """
    valid_wines = [
        wine for wine in wines_data
        if pd.notna(wine.get('Цена')) and wine.get('Цена') > 0
    ]

    if not valid_wines:
        return None

    return min(valid_wines, key=lambda wine: wine['Цена'])


def process_wine_data(wine_data, cheapest_wine):
    """
    Обрабатывает данные одного вина и определяет, является ли оно самым дешёвым.

    Args:
        wine_data (dict): Данные вина из Excel
        cheapest_wine (dict): Данные самого дешёвого вина

    Returns:
        dict: Обработанные данные вина
    """
    is_cheapest = (
        wine_data.get('Название') == cheapest_wine.get('Название') and
        wine_data.get('Цена') == cheapest_wine.get('Цена')
    )

    return {
        'Название': clean_data(wine_data.get('Название', '')),
        'Сорт': clean_data(wine_data.get('Сорт', '')),
        'Цена': wine_data.get('Цена') if pd.notna(wine_data.get('Цена')) else None,
        'Картинка': clean_data(wine_data.get('Картинка', '')),
        'Акция': is_cheapest
    }


def load_wine_data_from_excel(file_path):
    """
    Загружает данные о винах из Excel файла.

    Args:
        file_path (str): Путь к Excel файлу

    Returns:
        list: Список словарей с данными о винах
    """
    try:
        excel_data = pd.read_excel(file_path, na_values='', keep_default_na=False)
        return excel_data.to_dict(orient='records')
    except FileNotFoundError:
        raise FileNotFoundError(f"Файл {file_path} не найден!")
    except Exception as e:
        raise Exception(f"Ошибка при чтении Excel-файла: {e}")


def group_wines_by_category(wines_data, cheapest_wine):
    """
    Группирует вина по категориям.

    Args:
        wines_data (list): Список словарей с данными о винах
        cheapest_wine (dict): Данные самого дешёвого вина

    Returns:
        dict: Словарь с винами, сгруппированными по категориям
    """
    grouped_wines = defaultdict(list)

    for wine in wines_data:
        category = clean_data(wine.get('Категория', '')).strip()
        if not category:
            continue

        processed_wine = process_wine_data(wine, cheapest_wine)
        grouped_wines[category].append(processed_wine)

    return dict(grouped_wines)


def save_data_to_json(data, file_path):
    """
    Сохраняет данные в JSON файл.

    Args:
        data: Данные для сохранения
        file_path (str): Путь к файлу для сохранения
    """
    try:
        with open(file_path, 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False, indent=2)
    except Exception as e:
        raise Exception(f"Ошибка при сохранении JSON: {e}")


def generate_html_page(wines_data, age, output_path='index.html'):
    """
    Генерирует HTML страницу на основе данных о винах.

    Args:
        wines_data (dict): Сгруппированные данные о винах
        age (int): Возраст винодельни
        output_path (str): Путь для сохранения HTML файла
    """
    try:
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template('template.html')

        rendered_page = template.render(
            age=age,
            get_year_word=get_year_word,
            wines=wines_data
        )

        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(rendered_page)
    except Exception as e:
        raise Exception(f"Ошибка при генерации HTML: {e}")


def print_statistics(wines_data):
    """
    Выводит статистику по винам в консоль.

    Args:
        wines_data (dict): Сгруппированные данные о винах
    """
    print("\n=== СТАТИСТИКА ===")
    total_wines = 0
    total_actions = 0

    for category, wines_in_category in wines_data.items():
        category_actions = sum(1 for wine in wines_in_category if wine['Акция'])
        print(f"Категория '{category}': {len(wines_in_category)} вин (акций: {category_actions})")
        total_wines += len(wines_in_category)
        total_actions += category_actions

    print(f"Всего вин: {total_wines}")
    print(f"Всего категорий: {len(wines_data)}")
    print(f"Всего акционных товаров: {total_actions}")


def main():
    """
    Основная функция программы.
    """

    EXCEL_FILE_PATH = 'wine3.xlsx'
    JSON_OUTPUT_PATH = 'wine3.json'
    HTML_OUTPUT_PATH = 'index.html'
    FOUNDATION_YEAR = 1920

    try:
        current_year = datetime.now().year
        winery_age = current_year - FOUNDATION_YEAR

        print("Загрузка данных о винах...")
        wines_list = load_wine_data_from_excel(EXCEL_FILE_PATH)
        print(f"Успешно загружено {len(wines_list)} вин из {EXCEL_FILE_PATH}")

        cheapest_wine = find_cheapest_wine(wines_list)
        if cheapest_wine:
            print(f"Самое дешёвое вино: '{cheapest_wine['Название']}' за {cheapest_wine['Цена']} ₽")
        else:
            print("Не удалось найти самое дешёвое вино")
            cheapest_wine = {}

        grouped_wines = group_wines_by_category(wines_list, cheapest_wine)
        print_statistics(grouped_wines)

        print(f"\nСохранение данных в {JSON_OUTPUT_PATH}...")
        save_data_to_json(grouped_wines, JSON_OUTPUT_PATH)
        print("Данные успешно сохранены")
        print(f"Генерация HTML страницы {HTML_OUTPUT_PATH}...")
        generate_html_page(grouped_wines, winery_age, HTML_OUTPUT_PATH)
        print("HTML-страница успешно создана")
        print("\n✅ Программа завершена успешно!")

    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        return 1

    return 0


if __name__ == '__main__':
    main()
