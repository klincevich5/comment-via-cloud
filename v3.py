from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Настройка драйвера (укажите путь к chromedriver)
driver = webdriver.Chrome(executable_path='/path/to/chromedriver')

# URL Таблицы
url = 'https://docs.google.com/spreadsheets/d/1esvCh3ijbGN5fNH6NgXcJrFkEPF3JgG0mit9Zu_dFt4/edit'

# Авторизация
driver.get('https://accounts.google.com/')
time.sleep(2)
driver.find_element(By.ID, 'identifierId').send_keys('ВАШ_EMAIL@gmail.com')
driver.find_element(By.ID, 'identifierId').send_keys(Keys.ENTER)
time.sleep(2)
driver.find_element(By.NAME, 'password').send_keys('ВАШ_ПАРОЛЬ')
driver.find_element(By.NAME, 'password').send_keys(Keys.ENTER)

# Открытие Таблицы
time.sleep(5)
driver.get(url)

# Выбор ячейки и добавление комментария
time.sleep(5)
cell = driver.find_element(By.XPATH, '//div[@data-col="1" and @data-row="1"]')  # A1
cell.click()
time.sleep(1)

# Эмуляция добавления комментария
driver.find_element(By.XPATH, '//div[text()="Добавить комментарий"]').click()
time.sleep(1)
comment_box = driver.find_element(By.XPATH, '//textarea')
comment_box.send_keys('Это реальный комментарий через Selenium!')
comment_box.send_keys(Keys.ENTER)

# Закрытие браузера
time.sleep(3)
driver.quit()
