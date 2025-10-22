from datetime import datetime
from jinja2 import Environment, FileSystemLoader


def get_year_word(age):
    last_two = age % 100
    if 11 <= last_two <= 19:
        return 'лет'
    last_digit = age % 10
    if last_digit == 1:
        return 'год'
    elif 2 <= last_digit <= 4:
        return 'года'
    else:
        return 'лет'


def main():
    current_year = datetime.now().year
    foundation_year = 1920
    age = current_year - foundation_year

    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template.html')

    rendered_page = template.render(age=age, get_year_word=get_year_word)

    with open('index.html', 'w', encoding='utf-8') as file:
        file.write(rendered_page)


if __name__ == '__main__':
    main()
