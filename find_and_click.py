import telebot
import pandas as pd
import gdown
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton

# TOKEN
bot = telebot.TeleBot('7887269574:AAHuMtBxFodbtiJ1utONahWNTPx0hb62jRg')

# ID
file_ids = {
    'students': '1XvtlovXxrYSK0rfwSDicLa6895loXgWz',
    'schedule': '1hHgJHulaCCP5BU39V-DITHryqSGN54Yo'
}
output_paths = {
    'students': 'C:/Users/oxot5/Downloads/Отчет по студентам.xlsx',
    'schedule': 'C:/Users/oxot5/Downloads/Расписание группы.xlsx'
}

# Функ загрузки данных
def load_data(file_type):
    file_id = file_ids[file_type]
    file_url = f'https://drive.google.com/uc?export=download&id={file_id}'
    output_path = output_paths[file_type]
    gdown.download(file_url, output_path, quiet=False)
    df = pd.read_excel(output_path, header=0)  # Убедитесь, что заголовки на первой строке
    return df

# Функ отправкии сообщений
def send_chunked_message(chat_id, text, chunk_size=4096):
    for i in range(0, len(text), chunk_size):
        bot.send_message(chat_id, text[i:i + chunk_size])

# Анализ отчета по студентам
def analyze_students_report(message):
    try:
        df = load_data('students')

        if df.empty or df.shape[1] <= 17:
            bot.send_message(message.chat.id, "Данные по студентам пусты или недостаточно столбцов.")
            return

        df['СреднийБал'] = df.iloc[:, 17].apply(lambda x: 0 if x == '-' else float(x))
        def convert_to_5_scale(score):
            if score <= 3:
                return 2
            elif score <= 6:
                return 3
            elif score <= 9:
                return 4
            else:
                return 5

        df['Оценка'] = df['СреднийБал'].apply(convert_to_5_scale)
        low_average = df[df['Оценка'] < 3]

        response = "Студенты с низким средним баллом:\n"
        for index, row in low_average.iterrows():
            response += f"{row.iloc[0]} - Средний балл: {row['Оценка']}\n"

        send_chunked_message(message.chat.id, response if response else "Нет студентов с низким средним баллом.")
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

# Подсчет кколво пар
def count_classes_by_group(message):
    try:
        df = load_data('schedule')

        if df.empty:
            bot.send_message(message.chat.id, "Данные по расписанию пусты.")
            return
        days_of_week = ['Понедельник. 28.10.2024', 'Вторник. 29.10.2024',
                        'Среда. 30.10.2024', 'Четверг. 31.10.2024',
                        'Пятница. 01.11.2024', 'Суббота. 02.11.2024',
                        'Воскресенье. 03.11.2024']
        all_classes = []

        for day in days_of_week:
            if day in df.columns:
                day_classes = df[[df.columns[0], df.columns[1], day]].dropna()
                day_classes.columns = ['Группа', 'Пара', 'Занятие']
                all_classes.append(day_classes)
        all_classes_df = pd.concat(all_classes, ignore_index=True)
        class_count = all_classes_df['Занятие'].value_counts()

        response = "Количество проведенных пар по дисциплинам:\n"
        for discipline, count in class_count.items():
            response += f"{discipline} - {count} пар\n"

        send_chunked_message(message.chat.id, response if response else "Нет данных о парах.")
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

# меню
def create_menu():
    markup = InlineKeyboardMarkup()
    markup.add(InlineKeyboardButton("Анализ отчета по студентам", callback_data='analyze_students'))
    markup.add(InlineKeyboardButton("Подсчет проведенных пар", callback_data='count_classes'))
    return markup
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.send_message(message.chat.id, "Добро пожаловать! Выберите действие:", reply_markup=create_menu())
@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):
    if call.data == 'analyze_students':
        analyze_students_report(call.message)
    elif call.data == 'count_classes':
        count_classes_by_group(call.message)

# ЗАПУСК
if __name__ == "__main__":
    bot.polling(none_stop=True)