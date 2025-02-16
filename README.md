# Бот для информирования по расписанию

Отправляет сообщения в телеграм о расписании
Получает с Google Worksheet (расписания) данные о парах и отправляет пользователю (или чату).


# Конфигурация

Для работы необходимо создать  сервисный аккаунт для Google API. Инструкция по работе можно найти или можно посмотреть готовый пример [тут](https://habr.com/ru/articles/825404/). Полученный файлик нужно добавить рядом с исходным кодом и написать его название в `config.py`

Далее необходимо указать id таблицы (в ссылке он находится тут: https://docs.google.com/spreadsheets/d/ID_ТАБЛИЦЫ/...) и проверить, что есть права доступа на эту таблицу пользователю.

Для отправки самих уведомлений, нужен токен бота в Telegram, а также заполнить список словарей `groups`, в котором поле `chat_id` - id чата, в которое будет приходить уведомление, а `worksheet` - название листа, на котором будет просматриваться данный шаблон:

| Дата | Время пары №1 | Предмет       |
| ---- | ------------- | ------------- |
|      |               | Кабинет       |
|      |               | Преподаватель |
|      | Время пары №2 | Предмет       |
|      |               | Кабинет       |
|      |               | Преподаватель |


Для удобства определения названий листов есть функция `get_worksheet_info()`, которая возвращает список листов в таблице

Также для устранения проблемы, когда бот использует всю квоту запросов к Google API, добавлены функции `sleep(3)`

