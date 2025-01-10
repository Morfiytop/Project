import telebot
import pandas as pd
import gdown

# tooken
bot = telebot.TeleBot('7887269574:AAHuMtBxFodbtiJ1utONahWNTPx0hb62jRg')
#google drive loader
file_id = '1XvtlovXxrYSK0rfwSDicLa6895loXgWz'
file_url = f'https://drive.google.com/uc?export=download&id={file_id}'
output_path = 'C:/Users/oxot5/Downloads/Отчет по студентам.xlsx'

# google drive
def load_data():
    gdown.download(file_url, output_path, quiet=False)
    df = pd.read_excel(output_path)
    return df

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, 'Пиши /pairs ')

@bot.message_handler(commands=['pairs'])
def pairs(message):
    try:
        # вывод
        df = load_data()
        group_pairs = df.groupby('Группа')['pairs in total per month'].first()
        response = "Количество пар по группам:\n"
        for group, pairs in group_pairs.items():
            response += f"{group} - {pairs}\n"
        # Отправляем ответ
        bot.send_message(message.chat.id, response)
    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)}")

@bot.message_handler(commands=['help'])
def help_command(message):
    bot.send_message(message.chat.id,
                     'Доступные команды:\n/start - начать работу с ботом\n/pairs - вывести количество пар по группам')
bot.polling(none_stop=True)