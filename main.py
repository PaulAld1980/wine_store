from datetime import datetime
from jinja2 import Environment, FileSystemLoader


def main():
    current_year = datetime.now().year

    foundation_year = 1920
    age = current_year - foundation_year

    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template.html')

    rendered_page = template.render(age=age)

    with open('index.html', 'w', encoding='utf-8') as file:
        file.write(rendered_page)

    print('✅ Страница index.html успешно сгенерирована!')


if __name__ == '__main__':
    main()
