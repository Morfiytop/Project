import telebot
import pandas as pd
import gdown

# tooken
bot = telebot.TeleBot('7887269574:AAHuMtBxFodbtiJ1utONahWNTPx0hb62jRg')

# google drive loader
file_ids = {
    'students': '1XvtlovXxrYSK0rfwSDicLa6895loXgWz',
    'topics': '17LQ-z0qRQHjuXHh6OBucnFP0L-4yDHip',
    'schedule': '1hHgJHulaCCP5BU39V-DITHryqSGN54Yo',
    'attendance': '1v2_ud3ibpzAgmF_vJ7VLv-RnbflBcXFJ',
    'homework': '1_mfyE2wh8_j-aucuwxBrYk33oDOBj0Xr',
    'homework6': '1-UpX65wTexXEz_S6EaRhvAq15QIQJnKe'
}

output_paths = {
    'students': 'C:/Users/oxot5/Downloads/Отчет по студентам.xlsx',
    'topics': 'C:/Users/oxot5/Downloads/Отчет по темам занятий.xlsx',
    'schedule': 'C:/Users/oxot5/Downloads/Расписание группы.xlsx',
    'attendance': 'C:/Users/oxot5/Downloads/Отчет по посещаемости студентов.xlsx',
    'homework': 'C:/Users/oxot5/Downloads/Отчет по домашним заданиям.xlsx',
    'homework6': 'C:/Users/oxot5/Downloads/Отчет по дз (6 задание).xlsx'
}

def load_data(file_type):
    file_id = file_ids[file_type]
    file_url = f'https://drive.google.com/uc?export=download&id={file_id}'
    output_path = output_paths[file_type]
    gdown.download(file_url, output_path, quiet=False)
    df = pd.read_excel(output_path)
    return df

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, 'Пиши /pairs, /topics, /schedule, /attendance, /homework или /homework6')

@bot.message_handler(commands=['pairs'])
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

@bot.message_handler(commands=['topics'])
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

@bot.message_handler(commands=['schedule'])
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

@bot.message_handler(commands=['attendance'])
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

@bot.message_handler(commands=['homework'])
def homework(message):
    try:
        df = load_data('homework')
        response = "Домашние задания:\n"
        for index, row in df.iterrows():
            response += f"Форма обучения: {row['Форма обучения']}\n"
            response += f"ФИО преподавателя: {row['ФИО преподавателя']}\n"
            response += f"Месяц: {row['Месяц']}\n"
            response += f"Неделя 1:\n"
            response += f"  Кол-во пар: {row['Кол-во пар']}\n"
            response += f"  Выдано: {row['Выдано']}\n"
            response += f"  Получено: {row['Получено']}\n"
            response += f"  Проверено: {row['Проверено']}\n"
            response += f"  План: {row['План']}\n"
            response += f"Неделя 2:\n"
            response += f"  Кол-во пар: {row['Кол-во пар.1']}\n"
            response += f"  Выдано: {row['Выдано.1']}\n"
            response += f"  Получено: {row['Получено.1']}\n"
            response += f"  Проверено: {row['Проверено.1']}\n"
            response += f"  План: {row['План.1']}\n"
            response += f"Неделя 3:\n"
            response += f"  Кол-во пар: {row['Кол-во пар.2']}\n"
            response += f"  Выдано: {row['Выдано.2']}\n"
            response += f"  Получено: {row['Получено.2']}\n"
            response += f"  Проверено: {row['Проверено.2']}\n"
            response += f"  План: {row['План.2']}\n\n"
        bot.send_message(message.chat.id, response)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

@bot.message_handler(commands=['homework6'])
def homework6(message):
    try:
        df = load_data('homework6')
        response = "Домашние задания (6 задание):\n"
        for index, row in df.iterrows():
            response += f"Форма обучения: {row['Форма обучения']}\n"
            response += f"ФИО преподавателя: {row['ФИО преподавателя']}\n"
            response += f"Месяц: {row['Месяц']}\n"
            response += f"Неделя 1:\n"
            response += f"  Кол-во пар: {row['Кол-во пар']}\n"
            response += f"  Выдано: {row['Выдано']}\n"
            response += f"  Получено: {row['Получено']}\n"
            response += f"  Проверено: {row['Проверено']}\n"
            response += f"  План: {row['План']}\n"
            response += f"Неделя 2:\n"
            response += f"  Кол-во пар: {row['Кол-во пар.1']}\n"
            response += f"  Выдано: {row['Выдано.1']}\n"
            response += f"  Получено: {row['Получено.1']}\n"
            response += f"  Проверено: {row['Проверено.1']}\n"
            response += f"  План: {row['План.1']}\n"
            response += f"Неделя 3:\n"
            response += f"  Кол-во пар: {row['Кол-во пар.2']}\n"
            response += f"  Выдано: {row['Выдано.2']}\n"
            response += f"  Получено: {row['Получено.2']}\n"
            response += f"  Проверено: {row['Проверено.2']}\n"
            response += f"  План: {row['План.2']}\n\n"
        bot.send_message(message.chat.id, response)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

@bot.message_handler(commands=['help'])
def help_command(message):
    bot.send_message(message.chat.id,
                     'Доступные команды:\n/start - начать работу с ботом\n/pairs - вывести количество пар по группам\n/topics - вывести темы занятий по группам\n/schedule - вывести расписание группы\n/attendance - вывести посещаемость студентов\n/homework - вывести домашние задания\n/homework6 - вывести домашние задания (6 задание)')

bot.polling(none_stop=True)