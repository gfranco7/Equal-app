import pandas as pd
import time
import webbrowser
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==== Funciones ====

def cargar_excel():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],
        title="Selecciona un archivo Excel"
    )
    if filepath:
        try:
            df = pd.read_excel(filepath)
            placas = df['Placa'].astype(str).str.upper().tolist()
            text_area.delete("1.0", tk.END)
            text_area.insert(tk.END, "\n".join(placas))
            messagebox.showinfo("Archivo cargado", f"Cargadas {len(placas)} placas desde Excel.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")

def buscar_placas():
    placas_usuario = text_area.get("1.0", tk.END).strip().splitlines()
    placas_usuario = [p.upper() for p in placas_usuario if p.strip()]
    if not placas_usuario:
        messagebox.showwarning("Advertencia", "Por favor ingresa o carga al menos una placa.")
        return
    threading.Thread(target=ejecutar_busqueda, args=(placas_usuario,), daemon=True).start()

def ejecutar_busqueda(placas_usuario):
    btn_buscar.config(state=tk.DISABLED)
    text_resultados.delete("1.0", tk.END)
    text_resultados.insert(tk.END, "Iniciando búsqueda...\n")

    try:
        service = Service('./chromedriver.exe')  # Ajustar ruta si es necesario
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # Opcional
        driver = webdriver.Chrome(service=service, options=options)

        driver.get("https://anm.gov.co/informacion-atencion-minero-estado-aviso")

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.views-table.cols-6"))
        )
        time.sleep(2)

        tabla = driver.find_element(By.CSS_SELECTOR, "table.views-table.cols-6")
        filas = tabla.find_elements(By.TAG_NAME, "tr")[1:]

        encontrados = 0

        for fila in filas:
            if fila.is_displayed():
                columnas = fila.find_elements(By.TAG_NAME, "td")
                if len(columnas) >= 6:
                    placas = columnas[1].text.strip().upper()
                    try:
                        enlace = columnas[4].find_element(By.TAG_NAME, "a").get_attribute("href")
                    except:
                        enlace = "Sin enlace"

                    for placa_usuario in placas_usuario:
                        if placa_usuario in placas:
                            encontrados += 1
                            texto = f"🔎 Placa encontrada: {placa_usuario}\n📄 PDF: {enlace}\n------\n"
                            text_resultados.insert(tk.END, texto)
                            text_resultados.see(tk.END)
                            if enlace != "Sin enlace":
                                webbrowser.open(enlace)

        if encontrados == 0:
            text_resultados.insert(tk.END, "No se encontraron placas coincidentes.\n")

        driver.quit()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    btn_buscar.config(state=tk.NORMAL)


# ==== Configuración de la ventana principal ====
root = tk.Tk()
root.title("🌒 Buscador de placas ANM")
root.geometry("700x620")
root.configure(bg="#1e1e1e")

# ==== Estilos personalizados ====
style_config = {
    "bg": "#1e1e1e",
    "fg": "#ffffff",
    "insertbackground": "#ffffff",  # Cursor blanco
    "selectbackground": "#2ecc71"
}

btn_style = {
    "bg": "#2ecc71",
    "fg": "#1e1e1e",
    "activebackground": "#27ae60",
    "activeforeground": "#ffffff",
    "font": ("Arial", 10, "bold"),
    "bd": 0,
    "padx": 10,
    "pady": 5
}

label_style = {
    "bg": "#1e1e1e",
    "fg": "#2ecc71",
    "font": ("Arial", 11, "bold")
}

# ==== Widgets ====
tk.Label(root, text="Ingresa las placas (una por línea):", **label_style).pack(pady=(10, 0))
text_area = scrolledtext.ScrolledText(root, width=50, height=10, **style_config)
text_area.pack(pady=5)

btn_cargar = tk.Button(root, text="📂 Cargar Excel", command=cargar_excel, **btn_style)
btn_cargar.pack(pady=5)

btn_buscar = tk.Button(root, text="🔍 Buscar placas", command=buscar_placas, **btn_style)
btn_buscar.pack(pady=10)

tk.Label(root, text="Resultados:", **label_style).pack()
text_resultados = scrolledtext.ScrolledText(root, width=80, height=15, **style_config)
text_resultados.pack(pady=5)

# ==== Ejecutar ====
root.mainloop()
