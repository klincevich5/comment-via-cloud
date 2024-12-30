from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Укажите путь к chromedriver.exe
service = Service(r'C:\Users\klinc\comment-via-cloud\chromedriver.exe')

# Создайте экземпляр WebDriver
driver = webdriver.Chrome(service=service)

# Тестовый запуск
driver.get('https://www.google.com')
print("Браузер запущен успешно!")
driver.quit()
