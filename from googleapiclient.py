from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Настройка доступа
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'tdevs-444302-2ed4eef84e26.json'

credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

# ID таблицы Google
SPREADSHEET_ID = '1esvCh3ijbGN5fNH6NgXcJrFkEPF3JgG0mit9Zu_dFt4'

# Добавление комментария (имитация через заметку в ячейке)
def add_comment():
    range_name = 'A1'  # Ячейка
    comment_text = "Комментарий из API"
    
    request = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption='RAW',
        body={'values': [[comment_text]]}
    )
    response = request.execute()
    print(f"Комментарий добавлен: {response}")

add_comment()
