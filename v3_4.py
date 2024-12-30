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

    # Явное перемещение к ячейке B5 через JavaScript
    logging.info("Перемещение курсора к ячейке B5")
    script = """
        let cell = document.querySelector('[aria-rowindex="5"][aria-colindex="2"]');
        if (cell) {
            cell.click();
        } else {
            console.log('Ячейка B5 не найдена!');
        }
    """
    driver.execute_script(script)
    log_and_wait("Курсор перемещен к ячейке B5", 2)

    # Добавление комментария
    logging.info("Добавление комментария через Ctrl + Alt + M")
    ActionChains(driver).key_down(Keys.CONTROL).key_down(Keys.ALT).send_keys('m').key_up(Keys.ALT).key_up(Keys.CONTROL).perform()
    log_and_wait("Ожидание открытия поля комментария", 2)

    # Поиск окна комментария
    full_xpath = '/html/body/div[4]/div/div[2]/div/div[6]/div[14]/div/div[1]/div/div[2]/div[1]'
    logging.info(f"Поиск окна ввода комментария по XPath: {full_xpath}")

    try:
        comment_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, full_xpath))
        )
        logging.info("Окно ввода комментария найдено")

        # Ввод текста через Selenium
        logging.info("Ввод текста в окно комментария")
        comment_field.click()
        comment_field.send_keys("Это тестовый комментарий для ячейки B5.")
        log_and_wait("Текст введен", 1)

        # Нажимаем Enter для сохранения комментария
        logging.info("Сохранение комментария (нажатие Enter)")
        comment_field.send_keys(Keys.RETURN)
        log_and_wait("Ожидание сохранения комментария", 5)

        logging.info("Комментарий успешно добавлен")
    except Exception as e:
        logging.error(f"Ошибка при вводе комментария: {e}")
        raise

except Exception as e:
    logging.error(f"Общая ошибка: {e}")

finally:
    logging.info("Закрытие браузера")
    driver.quit()
