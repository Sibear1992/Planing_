import pandas as pd
import telebot
from telebot import types
import re
import os


# Инициализация бота с вашим токеном
bot = telebot.TeleBot("6885364516:AAG9saYSooic1MjM5Pk-2GX3I-VlTXJl2A4")
YOUR_USER_ID = '350566705'
excel_file_path = "plan.xlsx"  # Путь к файлу Excel с данными о плане

# Переменная для хранения DataFrame с данными о плане
df = None


# Функция для обновления данных в DataFrame из нового файла Excel
def update_plan(file_path):
    global df
    old_df = df.copy() if df is not None else None  # Создаем копию, если df не None
    new_df = pd.read_excel(file_path, sheet_name='Пресс', header=7)
    new_df = new_df.iloc[3:-12]
    new_df.rename(columns={'Unnamed: 2': 'Номер пресса'}, inplace=True)
    new_df['Номер пресса'] = new_df['Номер пресса'].ffill()
    new_df['Номер пресса'] = new_df['Номер пресса'].apply(lambda x: int(extract_numbers(x)[0]) if extract_numbers(x) else None)
    new_df['Номер пресса'] = new_df['Номер пресса'].ffill()

    if old_df is not None:
        # Фильтрация новых данных
        filtered_new_df = new_df[['Направление', 'Цвет', 'Шор', 'Осталось выпустить листы, шт']].dropna()
        filtered_old_df = old_df[['Направление', 'Цвет', 'Шор', 'Осталось выпустить листы, шт']].dropna()
        merged_df = pd.merge(filtered_old_df, filtered_new_df, on=['Направление', 'Цвет', 'Шор'],
                             suffixes=('_old', '_new'), how='outer', indicator=True)

        # Удаление столбца '_merge'
        merged_df.drop('_merge', axis=1, inplace=True)


        # Группировка по направлению, цвету и шору и суммирование остальных колонок
        merged_df = merged_df.groupby(['Направление', 'Цвет', 'Шор']).sum().reset_index()



        # Вычисление изменений в количестве блоков для каждого направления и цвета
        merged_df['Разница'] = merged_df['Осталось выпустить листы, шт_new'] - merged_df['Осталось выпустить листы, шт_old']

        # Отправка статистики об изменениях
        decreased_data = merged_df[merged_df['Разница'] < 0]
        increased_data = merged_df[merged_df['Разница'] > 0]
        send_statistics(decreased_data, increased_data)

        # Сохранение объединенного DataFrame в файл Excel
        file_path = 'merged_data.xlsx'
        merged_df.to_excel(file_path, index=False)

    df = new_df  # Обновляем DataFrame

# Функция для извлечения цифр из строки
def extract_numbers(s):
    return re.findall(r'\d+', s)


# Функция для отправки статистики об изменениях
def send_statistics(decreased_data, increased_data):
    if not decreased_data.empty:
        decreased_stats = "Убавилось:\n\n"
        for index, row in decreased_data.iterrows():
            decreased_stats += f"{row['Направление']} ({row['Цвет']}): {-row['Разница']} блоков.\n"
        bot.send_message(YOUR_USER_ID, decreased_stats, parse_mode='Markdown')

    if not increased_data.empty:
        increased_stats = "Прибавилось:\n\n"
        for index, row in increased_data.iterrows():
            increased_stats += f"{row['Направление']} ({row['Цвет']}): {row['Разница']} блоков.\n"
        bot.send_message(YOUR_USER_ID, increased_stats, parse_mode='Markdown')



# Функция для обработки команды /start
@bot.message_handler(commands=['start'])
def handle_start(message):
    # Создаем разметку с кнопками для выбора номера пресса
    markup = types.ReplyKeyboardMarkup(row_width=3)
    for press_number in sorted(df['Номер пресса'].unique()):
        markup.add(types.KeyboardButton(f"Пресс {int(press_number)}"))

    bot.send_message(message.chat.id,
                     'Привет! Я бот, который поможет вам узнать, что делать на каком прессе. Просто выберите номер пресса из списка ниже:',
                     reply_markup=markup)


# Функция для обработки входящего файла Excel с данными о плане
@bot.message_handler(content_types=['document'])
def handle_document(message):
    # Проверяем, что файл отправлен от нужного пользователя
    if message.from_user.id == YOUR_USER_ID:
        bot.send_message(message.chat.id, "Извините, вы не авторизованы для обновления плана.")
        return

    # Сохраняем принятый файл
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    with open(excel_file_path, 'wb') as new_file:
        new_file.write(downloaded_file)

    # Обновляем данные в DataFrame
    update_plan(excel_file_path)

    bot.send_message(message.chat.id, "План успешно обновлен.")


# Функция для обработки входящего сообщения
@bot.message_handler(func=lambda message: True)
def handle_message(message):
    user_input = message.text

    # Логика анализа запроса пользователя и вывод информации о блоках на прессе
    if user_input.startswith('Пресс'):
        try:
            press_number = int(user_input.split()[1])  # Получаем номер пресса из сообщения пользователя
            press_data = df[df['Номер пресса'] == press_number]  # Фильтруем данные по номеру пресса

            response = ""
            if not press_data.empty:
                response += f"На прессе {press_number} осталось выпустить:\n"
                for index, row in press_data.iterrows():
                    blocks_left = row['Осталось выпустить листы, шт']
                    if blocks_left > 0:
                        if press_number == 3:
                            response += f"Шор {int(round(row['Шор']))}: "  # Округляем значение шора до целых
                        response += f"{row['Направление']} ({row['Цвет']}): {blocks_left} блоков.\n"
            else:
                response = f"Пресс {press_number} не найден в базе данных."

        except IndexError:
            response = "Пожалуйста, укажите номер пресса после команды /blocks."

        except ValueError:
            response = "Номер пресса должен быть числом."

        bot.send_message(message.chat.id, response)
    else:
        response = "Пресс 1: Выпустить новую партию. Пресс 2: Проверить цвет. Пресс 3: Остановлен по техническим причинам."
        bot.send_message(message.chat.id, response)


# Проверка наличия сохраненного файла Excel и загрузка его при запуске бота
if os.path.exists(excel_file_path):
    update_plan(excel_file_path)


# Запуск бота
bot.polling()
