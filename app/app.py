from flask import Flask, request, redirect, url_for, render_template, after_this_request, flash, send_file
from flask_login import LoginManager, login_required, current_user, UserMixin, login_user
from app.forms import PeriodForm
import requests
from app.report_generator import ReportGenerator
from datetime import datetime, time
from app.logger import setup_logger
from dotenv import load_dotenv
import os
import sys

logger = setup_logger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Окружение для локального запуска
dotenv_path = os.path.join(BASE_DIR, "..", ".env")
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)

# Окружение для прода
required_envs = ["SECRET_FOR_FORM", "TOKEN_MS", "VALID_USERNAME", "VALID_PASSWORD", "BASE_URL"]
missing = [var for var in required_envs if not os.environ.get(var)]
if missing:
    print(f"Ошибка: отсутствуют переменные окружения: {missing}")
    sys.exit(1)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_FOR_FORM")

token_ms = os.environ.get("TOKEN_MS")

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Необходимо выполнить вход'


def fill_projects():
    projects = []
    r = requests.get(
        url="https://api.moysklad.ru/api/remap/1.2/entity/project",
        headers={
            "Authorization": f"Bearer {token_ms}",
        }
    )

    if r.status_code != 200:
        return {"error": "Ошибка запроса в мой склад"}

    try:
        body = r.json()
        for project in body.get("rows"):
            curr_proj = (project.get("meta").get("href"), project.get("name"))
            projects.append(curr_proj)

        return projects
    except Exception as e:
        logger.exception(f"Ошибка заполнения проектов от МС {e}")
        return {"error": "Ошибка обработки ответа"}



class User(UserMixin):
    def __init__(self, id):
        self.id = id


@login_manager.user_loader
def load_user(user_id):
    return User(user_id)


VALID_USERNAME = os.environ.get("VALID_USERNAME")
VALID_PASSWORD = os.environ.get("VALID_PASSWORD")


@app.route('/')
def test():
    return redirect(url_for('generate_report'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')

        if username == VALID_USERNAME and password == VALID_PASSWORD:
            user = User(1)
            login_user(user)
            return redirect(url_for('generate_report'))
        else:
            flash('Неверный логин или пароль', 'error')
            return redirect(url_for('login'))

    return render_template('login.html')


@app.route('/generate_report', methods=['GET', 'POST'])
@login_required
def generate_report():
    form: PeriodForm = PeriodForm()

    if request.method == 'GET':
        today = datetime.today()
        form.date_from.data = datetime.combine(today, time(hour=0, minute=0))
        form.date_to.data = datetime.combine(today, time(hour=23, minute=59))

    projects = fill_projects()

    form.projects.choices = projects

    generated_link = None
    selected_project_name = None

    if form.validate_on_submit():
        try:
            # Получаем данные из формы
            project_href = form.projects.data
            date_from = form.date_from.data
            date_to = form.date_to.data

            # Находим название выбранного проекта
            for href, name in projects:
                if href == project_href:
                    selected_project_name = name
                    break

            # ФОРМИРУЕМ ССЫЛКУ
            base_url = os.environ.get("BASE_URL")

            rg = ReportGenerator(token=token_ms)
            rg.set_urls(project=project_href, from_date=str(date_from), to_date=str(date_to))

            file_name = rg.generate_report(project=selected_project_name)

            generated_link = f"{base_url}/download_report/{file_name}"

            logger.info(f"Сгенерирована ссылка: {generated_link}")
            logger.info(f"Выбран проект: {selected_project_name}")
            logger.info(f"Период: с {date_from} по {date_to}")
        except Exception as e:
            logger.exception("Ошибка при генерации отчёта")
            flash("Произошла ошибка при генерации отчёта", "error")

    return render_template(
        "main.html",
        form=form,
        generated_link=generated_link,
        selected_project=selected_project_name
    )


@app.route('/download_report/<filename>')
@login_required
def download_report(filename):
    """Скачивание файла с автоматическим удалением после отправки"""

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    filepath = os.path.join(BASE_DIR, "temp", filename)

    # Проверяем, существует ли файл
    if not os.path.exists(filepath):
        flash('Файл не найден или уже был скачан', 'error')
        return redirect(url_for('generate_report'))

    # Регистрируем функцию, которая выполнится ПОСЛЕ отправки ответа
    @after_this_request
    def remove_file(response):
        """Удаляем файл после того, как он был отправлен клиенту"""
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
                logger.info(f"Файл {filename} успешно удален")
        except Exception as e:
            logger.exception(f"Ошибка при удалении файла {filename}: {e}")
        return response

    # Отправляем файл клиенту
    return send_file(
        filepath,
        as_attachment=True,  # Принудительное скачивание
        download_name=filename,  # Имя для скачивания
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run()
