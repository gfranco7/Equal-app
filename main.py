import pandas as pd
import time
import os
import webbrowser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==== Cargar lista del usuario desde Excel ====
df_placas_usuario = pd.read_excel("./placas_usuario.xlsx")  # <- Ruta del archivo
placas_usuario = df_placas_usuario['Placa'].astype(str).str.upper().tolist()  # Homogenizar

# ==== Configurar Selenium ====
service = Service('./chromedriver.exe')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# ==== Ir a la página de ANM ====
driver.get("https://anm.gov.co/informacion-atencion-minero-estado-aviso")

# ==== Esperar tabla objetivo ====
WebDriverWait(driver, 15).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "table.views-table.cols-6"))
)
time.sleep(2)

# ==== Extraer filas de la tabla ====
tabla = driver.find_element(By.CSS_SELECTOR, "table.views-table.cols-6")
filas = tabla.find_elements(By.TAG_NAME, "tr")[1:]  # Saltar encabezado

for fila in filas:
    if fila.is_displayed():
        columnas = fila.find_elements(By.TAG_NAME, "td")
        if len(columnas) >= 6:
            fecha = columnas[0].text.strip()
            placas = columnas[1].text.strip().upper()
            tipo = columnas[2].text.strip()
            num_doc = columnas[3].text.strip()

            try:
                enlace = columnas[4].find_element(By.TAG_NAME, "a").get_attribute("href")
            except:
                enlace = "Sin enlace"

            # ===== Comparar cada placa individualmente =====
            for placa_usuario in placas_usuario:
                if placa_usuario in placas:
                    print("🔎 Placa encontrada:", placa_usuario)
                    print("📄 PDF:", enlace)
                    print("------")

                    # Abrir el PDF
                    if enlace != "Sin enlace":
                        webbrowser.open(enlace)  # Abrir en navegador predeterminado

driver.quit()
