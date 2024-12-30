from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time
import logging

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("selenium_logs.log"),
        logging.StreamHandler()
    ]
)

def log_and_wait(message, seconds=2):
    logging.info(message)
    time.sleep(seconds)

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
    log_and_wait("Ожидание загрузки страницы", 5)

    # Выбор ячейки через JavaScript
    logging.info("Попытка выбора ячейки D5 через JavaScript")
    script = """
        let cell = document.querySelector('[aria-rowindex="5"][aria-colindex="4"]');
        if (cell) {
            cell.click();
        } else {
            console.log('Ячейка D5 не найдена!');
        }
    """
    driver.execute_script(script)
    log_and_wait("Ячейка выбрана", 2)

    # Добавление комментария
    logging.info("Добавление комментария через Ctrl + Alt + M")
    ActionChains(driver).key_down(Keys.CONTROL).key_down(Keys.ALT).send_keys('m').key_up(Keys.ALT).key_up(Keys.CONTROL).perform()
    log_and_wait("Ожидание открытия поля комментария", 2)

    # Ввод текста комментария
    comment_xpath = '//div[@aria-label="Добавить комментарий"]'
    logging.info(f"Поиск поля для ввода комментария по XPath: {comment_xpath}")
    comment_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, comment_xpath))
    )
    logging.info("Поле для ввода комментария найдено")
    comment_field.click()
    comment_text = "Это тестовый комментарий для ячейки D5."
    comment_field.send_keys(comment_text)
    logging.info(f"Ввод комментария: {comment_text}")
    comment_field.send_keys(Keys.RETURN)
    log_and_wait("Ожидание сохранения комментария", 5)

    logging.info("Комментарий успешно добавлен")

except Exception as e:
    logging.error(f"Ошибка: {e}")

finally:
    logging.info("Закрытие браузера")
    driver.quit()
