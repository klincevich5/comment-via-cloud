import requests

# URL вашего Web App из Google Apps Script
web_app_url = 'https://script.google.com/macros/s/AKfycbyIrmsV1D6Zr1TXKu-rCR3L231slS1ghd8T70GWCTuC6jV7z6F0GFGw7HqZVZuRDfUA/exec'

# Данные для запроса
data = {
    'sheetName': 'Sheet1',  # Имя листа
    'range': 'A1',  # Ячейка, куда добавить комментарий
    'comment': 'Это реальный комментарий, добавленный через скрипт!'  # Текст комментария
}

# Отправляем запрос POST
response = requests.post(web_app_url, data=data)

# Проверяем результат
if response.status_code == 200:
    print("Успешно:", response.text)
else:
    print("Ошибка:", response.status_code, response.text)
