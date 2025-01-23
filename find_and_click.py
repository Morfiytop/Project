import telebot
import pandas as pd
import gdown
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton

# tooken
bot = telebot.TeleBot('7887269574:AAHuMtBxFodbtiJ1utONahWNTPx0hb62jRg')

# google drive loader
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

def load_data(file_type):
    file_id = file_ids[file_type]
    file_url = f'https://drive.google.com/uc?export=download&id={file_id}'
    output_path = output_paths[file_type]
    gdown.download(file_url, output_path, quiet=False)
    df = pd.read_excel(output_path)
    return df

def create_keyboard():
    markup = InlineKeyboardMarkup()
    markup.row_width = 2
    markup.add(InlineKeyboardButton("Пары", callback_data="pairs"),
               InlineKeyboardButton("Темы", callback_data="topics"),
               InlineKeyboardButton("Расписание", callback_data="schedule"),
               InlineKeyboardButton("Посещаемость", callback_data="attendance"))
    return markup

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Выберите команду:", reply_markup=create_keyboard())

@bot.message_handler(func=lambda message: True)
def handle_all_messages(message):
    bot.send_message(message.chat.id, "Выберите команду:", reply_markup=create_keyboard())

@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):
    if call.data == "pairs":
        pairs(call.message)
    elif call.data == "topics":
        topics(call.message)
    elif call.data == "schedule":
        schedule(call.message)
    elif call.data == "attendance":
        attendance(call.message)

def pairs(message):
    try:
        df = load_data('students')
        group_pairs = df.groupby('Группа')['pairs in total per month'].first()
        response = "Количество пар по группам:\n"
        for group, pairs in group_pairs.items():
            response += f"{group} - {pairs}\n"
        bot.send_message(message.chat.id, response)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

def topics(message):
    try:
        df = load_data('topics')
        group_topics = df.groupby('Группа')['Тема урока'].first()
        response = "Темы занятий по группам:\n"
        for group, topics in group_topics.items():
            response += f"{group} - {topics}\n"
        bot.send_message(message.chat.id, response)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

def schedule(message):
    try:
        df = load_data('schedule')
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
        bot.send_message(message.chat.id, response)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

def attendance(message):
    try:
        df = load_data('attendance')
        response = "Посещаемость студентов:\n"
        for index, row in df.iterrows():
            response += f"ФИО преподавателя: {row['ФИО преподавателя']}\n"
            response += f"Колледж: {row['Колледж']}\n"
            response += f"Средняя посещаемость: {row['Средняя посещаемость']}\n"
            response += f"Всего пар: {row['Всего пар']}\n"
            response += f"Всего групп: {row['Всего групп']}\n\n"
        bot.send_message(message.chat.id, response)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

@bot.message_handler(commands=['help'])
def help_command(message):
    bot.send_message(message.chat.id,
                     'Доступные команды:\n/start - начать работу с ботом\n/pairs - вывести количество пар по группам\n/topics - вывести темы занятий по группам\n/schedule - вывести расписание группы\n/attendance - вывести посещаемость студентов')

bot.polling(none_stop=True)