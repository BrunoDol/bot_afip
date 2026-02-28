# -*- coding: utf-8 -*-
"""
Bot AFIP RCEL - Carga de comprobantes desde Excel (Hoja1) con marcación de estado
Requisitos:
  - pandas, openpyxl, selenium
  - chromedriver.exe accesible (misma carpeta o ruta indicada)
"""

from pathlib import Path
import time
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# =========================
# 1) CONFIG & LECTURA DE EXCEL (ROBUSTA)
# =========================
# Ruta del archivo de entrada
archivo_in = Path(r"C:\Users\pc\Desktop\Python\bot_afip\facturas.xlsx").resolve()
print(f"Usando archivo: {archivo_in}")
print("Última modificación del archivo:", time.ctime(archivo_in.stat().st_mtime))

# Archivo de salida (resultados con Estado/Mensaje)
archivo_out = archivo_in.with_name("facturas_resultados.xlsx")
print(f"Archivo de resultados: {archivo_out}")

# Leer por nombre de hoja
xls = pd.ExcelFile(archivo_in, engine="openpyxl")
print("Hojas encontradas:", xls.sheet_names)

df = pd.read_excel(
    archivo_in,
    sheet_name="Hoja1",     # 👈 tu hoja
    usecols="A:C",          # ajustá si tus columnas no son A:C
    dtype=str,
    engine="openpyxl"
)

# Normalización básica
df.columns = ["cuit", "importe", "concepto"]
df = df.dropna(how="all")
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Agregar columnas de tracking si no existen
for col in ["Estado", "Mensaje", "Timestamp", "Comprobante"]:
    if col not in df.columns:
        df[col] = ""

print("\nPrimeras filas que se van a procesar:")
print(df.head(10))
print(f"Total de filas (incluye ya procesadas): {len(df)}\n")

# Guardado helper (se guarda después de cada fila para no perder progreso)
def guardar_resultados():
    # Guardamos siempre la hoja Hoja1 con todas las columnas (incluye Estado/Mensaje)
    df.to_excel(archivo_out, sheet_name="Hoja1", index=False)

# =========================
# 2) SELENIUM - INICIO
# =========================
servicio = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servicio)


driver.get("https://www.afip.gob.ar/landing/default.asp")
print("Iniciá sesión MANUALMENTE y entrá al menú de RCEL (Comprobantes en Línea).")
time.sleep(15)  # aumentá si necesitás más tiempo para loguearte

# =========================
# 3) BUCLE DE FACTURAS (con estado)
# =========================
for index, row in df.iterrows():
    # Saltar filas ya procesadas en OK
    if str(row.get("Estado", "")).strip().upper() == "OK":
        print(f"⏭️  Fila {index+1}: ya estaba OK, se salta.")
        continue

    cuit_receptor = str(row["cuit"]) if pd.notna(row["cuit"]) else ""
    importe = str(row["importe"]) if pd.notna(row["importe"]) else ""
    concepto_texto = str(row["concepto"]) if pd.notna(row["concepto"]) else ""

    # Limpieza
    cuit_receptor = cuit_receptor.replace(" ", "").replace("-", "")
    concepto_texto = concepto_texto.strip()
    importe = importe.strip()

    if not cuit_receptor or not importe:
        msg = "Faltan datos (CUIT o Importe)."
        print(f"⚠️  Fila {index+1}: {msg} Se salta.")
        df.loc[index, "Estado"] = "ERROR"
        df.loc[index, "Mensaje"] = msg
        df.loc[index, "Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        guardar_resultados()
        continue

    print(f"\n🟢 Procesando fila {index+1} | CUIT {cuit_receptor} | Importe {importe}")

    try:
        # --- Generar Comprobantes ---
        generar_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Generar Comprobantes"))
        )
        generar_btn.click()
        print("Clic en 'Generar Comprobantes'.")
        time.sleep(1.5)

        # --- Punto de venta ---
        punto_venta_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "puntoDeVenta"))
        )
        selector = Select(punto_venta_select)
        if len(selector.options) > 1:
            selector.select_by_index(1)
            time.sleep(0.5)  # 👈 ESPERA 1/4 DE SEGUNDO
            print(f"Punto de venta seleccionado: {selector.options[1].text}")
        else:
            print("⚠️ Solo hay una opción de punto de venta. Se usa la única disponible.")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Continuar (punto de venta).")
        time.sleep(1.5)

        # --- Concepto ---
        concepto_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "idConcepto"))
        )
        concepto_selector = Select(concepto_select)

        elegido = False
        for opt in concepto_selector.options:
            if "servicio" in opt.text.lower():
                concepto_selector.select_by_visible_text(opt.text)
                print(f"Concepto seleccionado: {opt.text}")
                elegido = True
                break
        if not elegido and len(concepto_selector.options) > 0:
            print(f"⚠️ No encontré 'servicio'. Se deja: {concepto_selector.options[0].text}")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Continuar (concepto).")
        time.sleep(1.5)

        # --- Datos del receptor ---
        iva_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "idIVAReceptor"))
        )
        iva_selector = Select(iva_select)

        elegido = False
        for opt in iva_selector.options:
            if "consumidor" in opt.text.lower():
                iva_selector.select_by_visible_text(opt.text)
                print(f"IVA receptor: {opt.text}")
                elegido = True
                break
        if not elegido and len(iva_selector.options) > 0:
            print(f"⚠️ No encontré 'consumidor'. Se deja: {iva_selector.options[0].text}")

        cuit_campo = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.NAME, "nroDocReceptor"))
        )
        cuit_campo.clear()
        cuit_campo.send_keys(cuit_receptor)
        print(f"CUIT receptor seteado: {cuit_receptor}")

        # Forma de pago: Contado (si existe el elemento)
        try:
            contado_checkbox = driver.find_element(By.ID, "formadepago1")
            if not contado_checkbox.is_selected():
                contado_checkbox.click()
            print("Forma de pago: Contado.")
        except Exception:
            print("⚠️ No se pudo marcar 'Contado' (puede que ya esté por defecto).")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Continuar (datos del receptor).")
        time.sleep(0.5)

        # --- Detalle de la operación ---
        descripcion = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.NAME, "detalleDescripcion"))
        )
        descripcion.clear()
        descripcion.send_keys(concepto_texto)

        precio = driver.find_element(By.NAME, "detallePrecio")
        precio.clear()
        precio.send_keys(importe)
        print(f"Detalle cargado | Concepto: '{concepto_texto}' | Importe: {importe}")

        continuar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and contains(@value,'Continuar')]"))
        )
        continuar_btn.click()
        print("Continuar (datos de operación).")
        time.sleep(0.5)

        # --- Confirmar + aceptar confirmación (alerta o modal) + volver a menú ---
        confirmar_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "btngenerar"))
        )
        confirmar_btn.click()
        print("Confirmar (generar comprobante). Esperando confirmación...")

        # 1) Intentar primero un alert clásico de JS
        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
            print("Confirmación aceptada vía alerta JS.")
        except TimeoutException:
            # 2) Si no hay alert, buscamos el modal HTML con el botón "Confirmar"
            print("No se encontró alerta JS, buscando botón 'Confirmar' del modal...")
            confirmar_modal_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//span[@class='ui-button-text' and normalize-space()='Confirmar']")
                )
            )
            confirmar_modal_btn.click()
            print("Confirmación aceptada haciendo clic en el botón del modal.")

        time.sleep(1)

        # Intento de capturar algún dato de salida (opcional)
        # Si AFIP muestra número de comprobante en pantalla antes de volver al menú,
        # podrías extraerlo con un find_element y guardarlo en df.loc[index, "Comprobante"].
        # Dejo un placeholder:
        numero_comp = ""  # completar si lográs ubicar el selector
        if numero_comp:
            df.loc[index, "Comprobante"] = numero_comp

        menu_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[contains(@value,'Menú Principal')]"))
        )
        menu_btn.click()
        print("Volviendo a Menú Principal...\n")

        # Si llegó hasta acá, marcamos OK
        df.loc[index, "Estado"] = "OK"
        df.loc[index, "Mensaje"] = "Emitido correctamente"
        df.loc[index, "Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        guardar_resultados()

        time.sleep(0.5)

    except Exception as e:
        msg = f"Error: {e}"
        print(f"❌ Fila {index+1} falló | {msg}")
        df.loc[index, "Estado"] = "ERROR"
        df.loc[index, "Mensaje"] = msg
        df.loc[index, "Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        guardar_resultados()
        # seguimos con la siguiente
        continue

print("✅ Proceso finalizado. Revisá el archivo de resultados.")
print(f"Ruta resultados: {archivo_out}")

# Opcional: cerrar el navegador al terminar
# driver.quit()
