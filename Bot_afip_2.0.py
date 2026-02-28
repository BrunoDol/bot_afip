from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# leer excel con columnas: CUIT, importe, concepto, fecha
archivo = r"C:\Users\pc\Desktop\Python\bot_afip\facturas.xlsx"
df = pd.read_excel(archivo)

# selenium
servicio = Service(r"E:\Users\bruno\Desktop\python\bot afip\chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome(service=servicio)

driver.get("https://afip.gob.ar/landing/default.asp")
print("Iniciá sesión manualmente y llegá al menú de RCEL.")
time.sleep(15)

# bucle para todas las facturas
for index, row in df.iterrows():
    cuit_receptor = str(row[0])
    importe = str(row[1])
    concepto_texto = str(row[2])

    # formatear fecha con barras
    if pd.api.types.is_datetime64_any_dtype(df['fecha']):
        fecha_comprobante = row[3].strftime("%d/%m/%Y")
    else:
        fecha_comprobante = str(row[3])

    print(f"\n🟢 Procesando línea {index+1}/{len(df)} - CUIT: {cuit_receptor}, importe: {importe}, concepto: {concepto_texto}, fecha: {fecha_comprobante}")

    try:
        # click en Generar Comprobantes
        generar_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Generar Comprobantes"))
        )
        generar_btn.click()
        print("Clic en Generar Comprobantes.")
        time.sleep(1.5)
    except Exception as e:
        print(f"Error al generar comprobante: {e}")
        driver.quit()
        break

    # punto de venta
    try:
        punto_venta_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "puntoDeVenta"))
        )
        selector = Select(punto_venta_select)
        selector.select_by_index(1)
        print(f"Seleccionado punto de venta: {selector.options[1].text}")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Clic en Continuar en punto de venta.")
        time.sleep(1.5)
    except Exception as e:
        print(f"Error en punto de venta: {e}")
        driver.quit()
        break

    # datos de emisión
    try:
        concepto_select = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "idconcepto"))
        )
        concepto_selector = Select(concepto_select)
        for opt in concepto_selector.options:
            if "servicio" in opt.text.lower():
                concepto_selector.select_by_visible_text(opt.text)
                print(f"Concepto seleccionado: {opt.text}")
                break

        fecha_campo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "fc"))
        )
        fecha_campo.clear()
        fecha_campo.send_keys(fecha_comprobante)
        print(f"Fecha del comprobante: {fecha_comprobante}")

        desde_campo = driver.find_element(By.ID, "fsd")
        desde_campo.clear()
        desde_campo.send_keys(fecha_comprobante)
        print(f"Periodo desde: {fecha_comprobante}")

        hasta_campo = driver.find_element(By.ID, "fsh")
        hasta_campo.clear()
        hasta_campo.send_keys(fecha_comprobante)
        print(f"Periodo hasta: {fecha_comprobante}")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Clic en Continuar en datos de emisión.")
        time.sleep(1.5)
    except Exception as e:
        print(f"Error en datos de emisión: {e}")
        driver.quit()
        break

    # datos del receptor
    try:
        iva_select = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "idivareceptor"))
        )
        iva_selector = Select(iva_select)
        for opt in iva_selector.options:
            if "consumidor" in opt.text.lower():
                iva_selector.select_by_visible_text(opt.text)
                print(f"IVA receptor: {opt.text}")
                break

        cuit_campo = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "nrodocreceptor"))
        )
        cuit_campo.clear()
        cuit_campo.send_keys(cuit_receptor)
        print(f"CUIT receptor cargado: {cuit_receptor}")

        contado_checkbox = driver.find_element(By.ID, "formadepago1")
        if not contado_checkbox.is_selected():
            contado_checkbox.click()
        print("Forma de pago: Contado.")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Clic en Continuar en datos del receptor.")
        time.sleep(1.5)
    except Exception as e:
        print(f"Error en datos de receptor: {e}")
        driver.quit()
        break

    # datos de la operación
    try:
        # concepto como descripción
        descripcion_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "detalle_descripcion1"))
        )
        descripcion_field.clear()
        descripcion_field.send_keys(concepto_texto)
        print(f"Detalle: {concepto_texto}")

        # importe como precio unitario
        precio_field = driver.find_element(By.ID, "detalle_precio1")
        precio_field.clear()
        precio_field.send_keys(importe)
        print(f"Precio unitario: {importe}")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Clic en Continuar en datos de la operación.")
        time.sleep(1.5)

        # confirmar
        confirmar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "btngenerar"))
        )
        confirmar_btn.click()
        print("Clic en Confirmar datos.")

        # aceptar alerta
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        print("Alerta de confirmación aceptada.")

        time.sleep(1.5)

        # volver al menú principal
        menu_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[contains(@value,'Menú Principal')]"))
        )
        menu_btn.click()
        print("Volviendo al menú principal para la siguiente factura.")

        time.sleep(1.5)

    except Exception as e:
        print(f"Error en datos de operación o confirmación: {e}")
        driver.quit()
        break

print("✅ Proceso completado para todas las facturas.")
driver.quit()
