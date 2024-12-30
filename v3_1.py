from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Укажите путь к chromedriver
service = Service('/Users/anton-klintsevich/Documents/cloud comment/chromedriver')

# Создайте экземпляр WebDriver
driver = webdriver.Chrome(service=service)

# Тестовый запуск
driver.get('https://www.google.com')
print("Браузер запущен успешно!")
driver.quit()
