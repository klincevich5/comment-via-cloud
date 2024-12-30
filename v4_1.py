from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

# Путь к JSON-файлу
SERVICE_ACCOUNT_FILE = 'storied-chariot-446307-v6-b1898c29925d.json'

# Области доступа
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ID вашей Google Таблицы
SPREADSHEET_ID = '1esvCh3ijbGN5fNH6NgXcJrFkEPF3JgG0mit9Zu_dFt4'

def add_comment_to_cell():
    # Аутентификация с использованием Service Account
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    # Комментарий к ячейке
    sheet_id = 0  # Первый лист в таблице (если у вас другой лист, укажите его ID)
    note = "Это комментарий для ячейки B5"

    # Запрос для добавления комментария
    requests = [
        {
            "updateCells": {
                "rows": [
                    {
                        "values": [
                            {
                                "userEnteredValue": {"stringValue": "1"},  # Убедимся, что значение остается
                                "note": note
                            }
                        ]
                    }
                ],
                "fields": "note",
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 4,  # Индексация начинается с 0 (B5 — строка 5 = индекс 4)
                    "endRowIndex": 5,
                    "startColumnIndex": 1,  # Столбец B = индекс 1
                    "endColumnIndex": 2
                }
            }
        }
    ]

    # Выполнение запроса
    body = {"requests": requests}
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body).execute()

    print(f"Комментарий добавлен в ячейку B5: {response}")

if __name__ == '__main__':
    add_comment_to_cell()
