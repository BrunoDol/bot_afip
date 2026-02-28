from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time

# Ajustá la ruta a donde está tu chromedriver.exe
servicio = Service("E:/Users/bruno/Desktop/python/bot afip/chromedriver-win64/chromedriver.exe")

# inicializar el navegador
driver = webdriver.Chrome(service=servicio)

# probar abrir Google
driver.get("https://www.google.com")

# dejar abierto 5 segundos
time.sleep(5)

# cerrar
driver.quit()
