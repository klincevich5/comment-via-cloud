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

    # Переключение на лист Sheet1
    try:
        sheet_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='tab' and text()='Sheet1']"))
        )
        sheet_tab.click()
        log_and_wait("Переключение на лист Sheet1", 2)
    except Exception as e:
        logging.error(f"Не удалось переключиться на лист Sheet1: {e}")

    # Инициализация актуальной позиции
    current_position = {"row": 1, "col": 1}

    # Попытка выбрать каждую ячейку в диапазоне A1:D4
    for row in range(1, 5):  # строки 1-4
        for col in range(1, 5):  # столбцы A(1)-D(4)
            cell_address = chr(64 + col) + str(row)  # Генерация адреса (A1, B1...)
            logging.info(f"Обработка ячейки: {cell_address}")

            # Попытка выбора и ввода данных с учетом текущей позиции
            try:
                # Рассчитываем количество шагов для перемещения
                row_steps = row - current_position["row"]
                col_steps = col - current_position["col"]

                # Движение вниз или вверх
                if row_steps > 0:
                    for _ in range(row_steps):
                        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
                        log_and_wait("Перемещение вниз", 0.5)
                elif row_steps < 0:
                    for _ in range(abs(row_steps)):
                        ActionChains(driver).send_keys(Keys.ARROW_UP).perform()
                        log_and_wait("Перемещение вверх", 0.5)

                # Движение вправо или влево
                if col_steps > 0:
                    for _ in range(col_steps):
                        ActionChains(driver).send_keys(Keys.ARROW_RIGHT).perform()
                        log_and_wait("Перемещение вправо", 0.5)
                elif col_steps < 0:
                    for _ in range(abs(col_steps)):
                        ActionChains(driver).send_keys(Keys.ARROW_LEFT).perform()
                        log_and_wait("Перемещение влево", 0.5)

                # Установка фокуса на текущую ячейку
                script_focus = f"""
                const cell = document.querySelector('[aria-rowindex="{row}"][aria-colindex="{col}"]');
                if (cell) cell.focus();
                """
                driver.execute_script(script_focus)
                log_and_wait("Фокус установлен на ячейку", 0.2)

                # Ввод данных в ячейку
                comment = f"Address: {cell_address} | Zone: A1:D4"
                ActionChains(driver).send_keys(Keys.CONTROL, 'a').send_keys(Keys.BACKSPACE).perform()
                ActionChains(driver).send_keys(comment).send_keys(Keys.RETURN).perform()

                # Обновление текущей позиции
                current_position = {"row": row, "col": col}

                # Проверка содержимого ячейки через JavaScript
                script_check = f"""
                const cell = document.querySelector('[aria-rowindex="{row}"][aria-colindex="{col}"]');
                return cell ? cell.textContent : null;
                """
                cell_content = driver.execute_script(script_check)
                if cell_content and comment in cell_content:
                    logging.info(f"Успешно вставлено: '{comment}' в ячейку {cell_address}.")
                else:
                    logging.warning(f"Ошибка вставки данных в ячейку {cell_address}. Ожидалось: '{comment}', Получено: '{cell_content}'")

            except Exception as e:
                logging.error(f"Ошибка обработки ячейки {cell_address}: {e}")

except Exception as e:
    logging.error(f"Общая ошибка: {e}")
finally:
    logging.info("Закрытие браузера")
    driver.quit()
