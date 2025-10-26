from collections import defaultdict
from datetime import datetime
import json
from jinja2 import Environment, FileSystemLoader
import pandas as pd
import argparse
import os


def get_year_word(age):
    """
    Возвращает правильное склонение слова 'год' для настоящего времени.
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


def find_cheapest_wine(wines_data):
    """
    Находит самое дешёвое вино среди всех.
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
    """
    is_cheapest = (
        wine_data.get('Название') == cheapest_wine.get('Название') and
        wine_data.get('Цена') == cheapest_wine.get('Цена')
    )

    return {
        'Название': wine_data.get('Название', '') if pd.notna(wine_data.get('Название')) else '',
        'Сорт': wine_data.get('Сорт', '') if pd.notna(wine_data.get('Сорт')) else '',
        'Цена': wine_data.get('Цена') if pd.notna(wine_data.get('Цена')) else None,
        'Картинка': wine_data.get('Картинка', '') if pd.notna(wine_data.get('Картинка')) else '',
        'Акция': is_cheapest
    }


def load_wine_data_from_excel(file_path):
    """
    Загружает данные о винах из Excel файла.
    """
    excel_data = pd.read_excel(file_path, na_values='', keep_default_na=False)
    return excel_data.to_dict(orient='records')


def group_wines_by_category(wines_data, cheapest_wine):
    """
    Группирует вина по категориям.
    """
    grouped_wines = defaultdict(list)

    for wine in wines_data:
        category = wine.get('Категория', '')
        if pd.notna(category):
            category = category.strip()
        else:
            category = ''

        if not category:
            continue

        processed_wine = process_wine_data(wine, cheapest_wine)
        grouped_wines[category].append(processed_wine)

    return dict(grouped_wines)


def save_data_to_json(data, file_path):
    """
    Сохраняет данные в JSON файл.
    """
    with open(file_path, 'w', encoding='utf-8') as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


def generate_html_page(wines_data, age, output_path='index.html'):
    """
    Генерирует HTML страницу на основе данных о винах.
    """
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template.html')

    rendered_page = template.render(
        age=age,
        get_year_word=get_year_word,
        wines=wines_data
    )

    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(rendered_page)


def parse_arguments():
    """
    Парсит аргументы командной строки и переменные окружения.
    """
    parser = argparse.ArgumentParser(
        description='Генератор сайта-магазина вин из Excel файла'
    )

    parser.add_argument(
        '--excel-file',
        default=os.getenv('WINE_EXCEL_FILE', 'wine3.xlsx'),
        help='Путь к Excel файлу с данными о винах'
    )
    parser.add_argument(
        '--html-output',
        default=os.getenv('WINE_HTML_OUTPUT', 'index.html'),
        help='Путь для сохранения HTML файла'
    )
    parser.add_argument(
        '--template',
        default=os.getenv('WINE_TEMPLATE', 'template.html'),
        help='Путь к HTML шаблону'
    )
    parser.add_argument(
        '--save-json',
        action='store_true',
        default=os.getenv('WINE_SAVE_JSON', '').lower() in ('true', '1', 'yes'),
        help='Сохранять данные в JSON файл'
    )
    parser.add_argument(
        '--json-output',
        default=os.getenv('WINE_JSON_OUTPUT', 'wine_data.json'),
        help='Путь для сохранения JSON файла'
    )

    return parser.parse_args()


def main():
    """
    Основная функция программы.
    """
    args = parse_arguments()

    EXCEL_FILE_PATH = args.excel_file
    HTML_OUTPUT_PATH = args.html_output
    SAVE_JSON = args.save_json
    JSON_OUTPUT_PATH = args.json_output
    FOUNDATION_YEAR = 1920

    current_year = datetime.now().year
    winery_age = current_year - FOUNDATION_YEAR

    wines_list = load_wine_data_from_excel(EXCEL_FILE_PATH)
    cheapest_wine = find_cheapest_wine(wines_list)
    grouped_wines = group_wines_by_category(wines_list, cheapest_wine or {})

    if SAVE_JSON:
        save_data_to_json(grouped_wines, JSON_OUTPUT_PATH)

    generate_html_page(grouped_wines, winery_age, HTML_OUTPUT_PATH)


if __name__ == '__main__':
    main()
