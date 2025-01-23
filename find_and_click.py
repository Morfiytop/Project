import telebot
import pandas as pd
import gdown
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
# TOKEN
bot = telebot.TeleBot('7887269574:AAHuMtBxFodbtiJ1utONahWNTPx0hb62jRg')
# ID
file_ids = {
    'students': '1XvtlovXxrYSK0rfwSDicLa6895loXgWz',
    'topics': '17LQ-z0qRQHjuXHh6OBucnFP0L-4yDHip',
    'schedule': '1hHgJHulaCCP5BU39V-DITHryqSGN54Yo',
    'attendance': '1v2_ud3ibpzAgmF_vJ7VLv-RnbflBcXFJ'
}
output_paths = {
    'students': 'C:/Users/oxot5/Downloads/Отчет по студентам.xlsx',
    'topics': 'C:/Users/oxot5/Downloads/Отчет по темам занятий.xlsx',
    'schedule': 'C:/Users/oxot5/Downloads/Расписание группы.xlsx',
    'attendance': 'C:/Users/oxot5/Downloads/Отчет по посещаемости студентов.xlsx'
}
# google api
def load_data(file_type):
    file_id = file_ids[file_type]
    file_url = f'https://drive.google.com/uc?export=download&id={file_id}'
    output_path = output_paths[file_type]
    gdown.download(file_url, output_path, quiet=False)
    df = pd.read_excel(output_path)
    return df
# Функ кнопок
def create_keyboard():
    markup = InlineKeyboardMarkup()
    markup.row_width = 2
    markup.add(InlineKeyboardButton("Пары", callback_data="pairs"),
               InlineKeyboardButton("Темы", callback_data="topics"),
               InlineKeyboardButton("Расписание", callback_data="schedule"),
               InlineKeyboardButton("Посещаемость", callback_data="attendance"))
    return markup
# Функ назад
def create_back_button():
    markup = InlineKeyboardMarkup()
    markup.add(InlineKeyboardButton("Вернуться назад", callback_data="back"))
    return markup

# /start
@bot.message_handler(commands=['start'])
def start(message):
    send_welcome(message)
# сообщения
@bot.message_handler(func=lambda message: True)
def handle_all_messages(message):
    send_welcome(message)
# работа кнопок
@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):
    command_handlers = {
        "pairs": show_pairs,
        "topics": show_topics,
        "schedule": show_schedule,
        "attendance": show_attendance,
        "back": send_welcome
    }
    handler = command_handlers.get(call.data)
    if handler:
        handler(call.message)
# Функ ответа
def send_welcome(message):
    bot.send_message(message.chat.id, "Выберите команду:", reply_markup=create_keyboard())
def send_response(message, response):
    bot.send_message(message.chat.id, response, reply_markup=create_back_button())
# Функ ошибки
def send_error(message, error):
    bot.send_message(message.chat.id, f"Произошла ошибка: {str(error)}", reply_markup=create_back_button())
# Функ колво пар групп
def show_pairs(message):
    try:
        df = load_data('students')
        df = df.fillna({'Группа': '', 'pairs in total per month': 0})
        group_pairs = df.groupby('Группа')['pairs in total per month'].first()
        response = "Количество пар по группам:\n"
        for group, pairs in group_pairs.items():
            response += f"{group} - {pairs}\n"
        send_response(message, response)
    except Exception as e:
        send_error(message, e)

# Функ тем занятий
def show_topics(message):
    try:
        df = load_data('topics')
        df = df.fillna({'Группа': '', 'Тема урока': ''})
        group_topics = df.groupby('Группа')['Тема урока'].first()
        response = "Темы занятий по группам:\n"
        for group, topics in group_topics.items():
            response += f"{group} - {topics}\n"
        send_response(message, response)
    except Exception as e:
        send_error(message, e)

# Функ для расписания группы
def show_schedule(message):
    try:
        df = load_data('schedule')
        df = df.fillna({'Группа': '', 'Пара': '', 'Время': '', 'Понедельник. 28.10.2024': '', 'Вторник. 29.10.2024': '', 'Среда. 30.10.2024': '', 'Четверг. 31.10.2024': '', 'Пятница. 01.11.2024': '', 'Суббота. 02.11.2024': '', 'Воскресенье. 03.11.2024': ''})
        response = "Расписание группы:\n"
        for index, row in df.iterrows():
            response += f"Группа: {row['Группа']}\n"
            response += f"Пара: {row['Пара']}\n"
            response += f"Время: {row['Время']}\n"
            response += f"Понедельник: {row['Понедельник. 28.10.2024']}\n"
            response += f"Вторник: {row['Вторник. 29.10.2024']}\n"
            response += f"Среда: {row['Среда. 30.10.2024']}\n"
            response += f"Четверг: {row['Четверг. 31.10.2024']}\n"
            response += f"Пятница: {row['Пятница. 01.11.2024']}\n"
            response += f"Суббота: {row['Суббота. 02.11.2024']}\n"
            response += f"Воскресенье: {row['Воскресенье. 03.11.2024']}\n\n"
        send_response(message, response)
    except Exception as e:
        send_error(message, e)

# функ для студентов
def show_attendance(message):
    try:
        df = load_data('attendance')
        df = df.fillna({'ФИО преподавателя': '', 'Колледж': '', 'Средняя посещаемость': 0, 'Всего пар': 0, 'Всего групп': 0})
        response = "Посещаемость студентов:\n"
        for index, row in df.iterrows():
            response += f"ФИО преподавателя: {row['ФИО преподавателя']}\n"
            response += f"Колледж: {row['Колледж']}\n"
            response += f"Средняя посещаемость: {row['Средняя посещаемость']}\n"
            response += f"Всего пар: {row['Всего пар']}\n"
            response += f"Всего групп: {row['Всего групп']}\n\n"
        send_response(message, response)
    except Exception as e:
        send_error(message, e)

# Обработка помощи
@bot.message_handler(commands=['help'])
def help_command(message):
    bot.send_message(message.chat.id,
                     'Доступные команды:\n/start - начать работу с ботом\n/pairs - вывести количество пар по группам\n/topics - вывести темы заниятий по группам\n/schedule - вывести расписание группы\n/attendance - вывести посещаемость студентов')

# Запуск
bot.polling(none_stop=True)