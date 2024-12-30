from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import logging
import time

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler("selenium_logs.log"), logging.StreamHandler()]
)

def log_and_wait(message, seconds=2):
    logging.info(message)
    time.sleep(seconds)

# Настройки для Selenium
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

logging.info("Запуск ChromeDriver")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    spreadsheet_id = "1esvCh3ijbGN5fNH6NgXcJrFkEPF3JgG0mit9Zu_dFt4"
    google_sheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid=0"
    logging.info(f"Сформирован URL таблицы: {google_sheets_url}")

    logging.info("Открытие Google Sheets")
    driver.get(google_sheets_url)
    log_and_wait("Ожидание загрузки страницы", 10)

    # Попытка выбрать каждую ячейку в диапазоне A1:D5
    for row in range(1, 6):  # строки 1-5
        for col in range(1, 5):  # столбцы A(1)-D(4)
            cell_address = chr(64 + col) + str(row)  # Генерация адреса (A1, B1...)
            logging.info(f"Обработка ячейки: {cell_address}")

            # Попытка №1: Использование JavaScript
            try:
                script = f"""
                const cell = document.querySelector('[aria-rowindex="{row}"][aria-colindex="{col}"]');
                if (cell) {{
                    cell.click();
                    return true;
                }} else {{
                    return false;
                }}
                """
                result = driver.execute_script(script)
                if result:
                    logging.info(f"Ячейка {cell_address} выбрана с помощью JavaScript.")
                else:
                    raise Exception("Ячейка не найдена через JavaScript.")

            except Exception as e:
                logging.warning(f"JavaScript не смог выбрать ячейку {cell_address}: {e}")

            # Попытка №2: Использование клавиш навигации
            try:
                if row == 1 and col == 1:
                    # Переход в верхний левый угол, если это первая ячейка
                    ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.HOME).key_up(Keys.CONTROL).perform()
                else:
                    # Переход через стрелки
                    ActionChains(driver).send_keys(Keys.ARROW_DOWN * (row - 1)).perform()
                    ActionChains(driver).send_keys(Keys.ARROW_RIGHT * (col - 1)).perform()
                logging.info(f"Ячейка {cell_address} выбрана с помощью клавиш навигации.")
            except Exception as e:
                logging.warning(f"Клавиши навигации не смогли выбрать ячейку {cell_address}: {e}")

            # Ввод адреса в ячейку
            try:
                ActionChains(driver).send_keys(cell_address).send_keys(Keys.RETURN).perform()
                logging.info(f"Адрес {cell_address} вставлен в ячейку.")
            except Exception as e:
                logging.error(f"Не удалось вставить адрес {cell_address}: {e}")

except Exception as e:
    logging.error(f"Общая ошибка: {e}")
finally:
    logging.info("Закрытие браузера")
    driver.quit()
