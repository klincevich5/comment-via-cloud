from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Параметры таблицы
SPREADSHEET_ID = "1esvCh3ijbGN5fNH6NgXcJrFkEPF3JgG0mit9Zu_dFt4"  # ID вашей таблицы
SHEET_NAME = "Sheet1"  # Имя листа
CELL_ADDRESS = "A1"  # Адрес ячейки, в которую нужно добавить комментарий

# Загрузка учетных данных
SERVICE_ACCOUNT_FILE = "tdevs-444302-2ed4eef84e26.json"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Создание объекта авторизации
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Создание клиента для работы с Google Sheets API
service = build('sheets', 'v4', credentials=credentials)

# Добавление текста в указанную ячейку
def add_comment_to_cell(spreadsheet_id, sheet_name, cell_address, comment_text):
    # Определяем диапазон ячейки
    range_ = f"{sheet_name}!{cell_address}"
    
    # Обновляем значение ячейки с комментарием
    request = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_,
        valueInputOption="RAW",
        body={"values": [[comment_text]]}  # Текст для вставки
    )
    response = request.execute()
    print(f"Добавлен комментарий в {cell_address}: {comment_text}")

# Вызов функции для добавления комментария
add_comment_to_cell(SPREADSHEET_ID, SHEET_NAME, CELL_ADDRESS, "Это тестовый комментарий!")
