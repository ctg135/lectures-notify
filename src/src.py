import gspread
from gspread import Client, service_account
from telebot import TeleBot, types
from random_unicode_emoji import random_emoji

from time import sleep
from datetime import datetime, timedelta

import config

# Настройка Google Sheets API
def get_client() -> Client:
    return service_account(config.google_creds)

def get_table_by_url(client: Client, table_url) -> gspread.spreadsheet.Spreadsheet:
    """Получение таблицы из Google Sheets по ссылке."""
    return client.open_by_url(table_url)

def get_table_by_id(client: Client, table_id) -> gspread.spreadsheet.Spreadsheet:
    """Получение таблицы из Google Sheets по ID таблицы."""
    return client.open_by_key(table_id)

# Может пригодится
def get_worksheet_info(table) -> dict:
    """Возвращает количество листов в таблице и их названия."""
    worksheets = table.worksheets()
    worksheet_info = {
        "count": len(worksheets),
        "names": [worksheet.title for worksheet in worksheets]
    }
    return worksheet_info

def get_lections(date_cell: gspread.cell.Cell, worksheet: gspread.worksheet.Worksheet) -> list:
    '''
    Функция, которая по ячейке даты возвращает расписание на день в виде массива словарей
    
    Ключи: time, name, cabinet, lecturer
    '''
    height = 1
    while worksheet.cell(date_cell.row + height, date_cell.col).value == None:
        height = height + 1
        if height == 24: break
        sleep(3)
        

    # Проверка соответствия шаблону
    if height % 3 != 0:
        print(f'Неверный шаблон расписания {date_cell}, {worksheet}')
        return []

    # Количество ячеек для пар
    lections_count = int(height / 3)
    lections = []

    for lection_number in range(0, lections_count):
        
        # Получаем время пары -> следующая ячейка справа и через три строки
        lection_time = worksheet.cell(date_cell.row + (lection_number * 3),
                                    date_cell.col + 1)
        sleep(3)
        # Название лекции -> следующая ячейка справа от времени лекции
        lection_name = worksheet.cell(lection_time.row,
                                    lection_time.col + 1)
        sleep(3)
        # Кабинет проведения лекции -> следующая ячейка вниз от названия
        lection_cabinet = worksheet.cell(lection_name.row + 1,
                                    lection_name.col)
        sleep(3)
        # Преподаватель -> следующая ячейка вниз от кабинета
        lection_lecturer = worksheet.cell(lection_cabinet.row + 1,
                                    lection_name.col)
        sleep(3)
        
        lections.append({
            'time': lection_time.value,
            'name': lection_name.value,
            'cabinet': lection_cabinet.value,
            'lecturer': lection_lecturer.value
        })
        
    return lections
        
def get_today_lections(group: dict, table_id: str) -> list:
    '''
    Уведомляет о парах на сегодня
    '''
    # Google Settings
    client = get_client()
    table = get_table_by_id(client, table_id)
    worksheet = table.worksheet(group['worksheet'])

    # Получение нужной даты
    today = datetime.now().strftime('%d.%m.%Y')
    date_cell = worksheet.find(today)

    if not date_cell:
        print(f'На сегодня нету пар для: {group['worksheet']}')
        return []
        
    return get_lections(date_cell, worksheet)

def get_tomorrow_lections(group: dict, table_id: str) -> list:
    '''
    Уведомляет о парах на завтра
    '''
    # Google Settings
    client = get_client()
    table = get_table_by_id(client, table_id)
    worksheet = table.worksheet(group['worksheet'])
    
    # Получение нужной даты
    tomorrow = (datetime.now() + timedelta(days=1)).strftime('%d.%m.%Y')    
    date_cell = worksheet.find(tomorrow)

    if not date_cell:
        print(f'На завтра нету пар для: {group['worksheet']}')
        return []
        
    return get_lections(date_cell, worksheet)

def notify_today_lections(bot: TeleBot, chat_id: int, lections: list, table_id: str = ''):
    '''
    Уведомляет в чат сообщением о парах на сегодня
    '''
    if len(lections) == 0: return
    
    msg = random_emoji()[0] + ' Не забываем, пары сегодня:\n\n'
    for lection in lections:
        if lection['name'] == None:
            continue
        msg += f'<b>{lection['time']}</b> - '
        msg += f'<i>{lection['name']}</i> '
        msg += f'({lection['lecturer']}): '
        msg += f'{lection['cabinet']}\n\n'
    
    if table_id == '':
        bot.send_message(chat_id, msg)
    else:
        button_link = types.InlineKeyboardButton('Расписание 🌐', 
                                                 url=f'https://docs.google.com/spreadsheets/d/{table_id}')
        kbd = types.InlineKeyboardMarkup()
        kbd.add(button_link)
        bot.send_message(chat_id, msg, reply_markup=kbd)
        

def notify_tomorrow_lections(bot: TeleBot, chat_id: int, lections: list, table_id: str = ''):
    '''
    Уведомляет в чат сообщением о парах на завтра
    '''
    if len(lections) == 0: return
    
    msg = '⚠️ Завтра будут пары:\n\n'
    for lection in lections:
        if lection['name'] == None:
            continue
        msg += f'<b>{lection['time']}</b> - '
        msg += f'<i>{lection['name']}</i> '
        msg += f'({lection['lecturer']}): '
        msg += f'{lection['cabinet']}\n\n'
    
    if table_id == '':
        bot.send_message(chat_id, msg)
    else:
        button_link = types.InlineKeyboardButton('Расписание 🌐', 
                                                 url=f'https://docs.google.com/spreadsheets/d/{table_id}')
        kbd = types.InlineKeyboardMarkup()
        kbd.add(button_link)
        bot.send_message(chat_id, msg, reply_markup=kbd)

def check_table():
    '''
    Из конфигурации получает список групп и чатов,
    для каждой производит проверку расписания
    '''
    bot = TeleBot(config.bot_token, parse_mode='HTML')
    
    for group in config.groups:
        notify_today_lections(
            bot,
            group['chat_id'],
            get_today_lections(group, config.table_id),
            config.table_id
        )
        notify_tomorrow_lections(
            bot,
            group['chat_id'],
            get_tomorrow_lections(group, config.table_id),
            config.table_id
        )

if __name__ == '__main__':
    check_table()

