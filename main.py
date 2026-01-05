from collections import defaultdict
from datetime import datetime
import json
from jinja2 import Environment, FileSystemLoader
import pandas as pd
import argparse
from http.server import HTTPServer, SimpleHTTPRequestHandler


# Глобальная константа - она никогда не меняется
FOUNDATION_YEAR = 1920


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
    Обрабатывает данные одного вина и определяет,
    является ли оно самым дешёвым.
    """
    is_cheapest = (
        wine_data.get('Название') == cheapest_wine.get('Название') and
        wine_data.get('Цена') == cheapest_wine.get('Цена')
    )

    title = wine_data.get('Название', '')
    grape = wine_data.get('Сорт', '')
    price = wine_data.get('Цена')
    image = wine_data.get('Картинка', '')

    return {
        'Название': title if pd.notna(title) else '',
        'Сорт': grape if pd.notna(grape) else '',
        'Цена': price if pd.notna(price) else None,
        'Картинка': image if pd.notna(image) else '',
        'Акция': is_cheapest
    }


def load_wine_data_from_excel(file_path):
    """
    Загружает данные о винах из Excel файла.
    """
    excel_data = pd.read_excel(
        file_path,
        na_values='',
        keep_default_na=False
    )
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

    return grouped_wines


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
    Парсит аргументы командной строки.
    """
    parser = argparse.ArgumentParser(
        description='Генератор сайта-магазина вин из Excel файла'
    )

    parser.add_argument(
        '--excel-file',
        default='wine3.xlsx',
        help='Путь к Excel файлу с данными о винах'
    )
    parser.add_argument(
        '--html-output',
        default='index.html',
        help='Путь для сохранения HTML файла'
    )
    parser.add_argument(
        '--save-json',
        action='store_true',
        help='Сохранять данные в JSON файл'
    )
    parser.add_argument(
        '--json-output',
        default='wine_data.json',
        help='Путь для сохранения JSON файла'
    )

    return parser.parse_args()


def main():
    """
    Основная функция программы.
    """
    args = parse_arguments()

    current_year = datetime.now().year
    winery_age = current_year - FOUNDATION_YEAR

    wines_list = load_wine_data_from_excel(args.excel_file)
    cheapest_wine = find_cheapest_wine(wines_list)
    grouped_wines = group_wines_by_category(
        wines_list,
        cheapest_wine or {}
    )

    if args.save_json:
        save_data_to_json(grouped_wines, args.json_output)

    generate_html_page(grouped_wines, winery_age, args.html_output)

    print(f"HTML файл создан: {args.html_output}")
    print("Запуск веб-сервера на http://127.0.0.1:8000")

    server = HTTPServer(('127.0.0.1', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
