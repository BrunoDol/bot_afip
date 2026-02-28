from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import time

# leer el excel
archivo = r"C:\Users\pc\Desktop\Python\bot_afip\facturas.xlsx"
df = pd.read_excel(archivo)

# tomar la primera factura
primer_factura = df.iloc[0]
cuit_receptor = str(primer_factura[0])
importe = str(primer_factura[1])
concepto = str(primer_factura[2])

# configurar selenium
servicio = Service(r"E:\Users\bruno\Desktop\python\bot afip\chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome(service=servicio)

# ACÁ ES CLAVE: vos ya tenés que estar logueado en AFIP
# entonces podemos reutilizar la sesión
# por ahora, te pido que lo arranques a mano
# y desde acá no hacemos driver.get()

time.sleep(5)  # tiempo para que muevas la pestaña manualmente a la pantalla de la factura

# EJEMPLO: completar CUIT receptor
try:
    campo_cuit = driver.find_element(By.NAME, "txtCuitRecep")
    campo_cuit.clear()
    campo_cuit.send_keys(cuit_receptor)
    time.sleep(2)

    # completar importe
    campo_importe = driver.find_element(By.NAME, "txtImpTotal")
    campo_importe.clear()
    campo_importe.send_keys(importe)
    time.sleep(2)

    # completar concepto
    campo_concepto = driver.find_element(By.NAME, "txtDetalle")
    campo_concepto.clear()
    campo_concepto.send_keys(concepto)
    time.sleep(2)

    print("Se completaron los campos de la primera factura correctamente.")

except Exception as e:
    print(f"Error: {e}")

# dejamos el navegador abierto para que revises
