import pandas as pd
import time
import webbrowser
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

# Variable global para configuración
MAX_PAGINAS = 5

# ==== Funciones ====

def configurar_max_paginas():
    """Función para configurar el número máximo de páginas a revisar"""
    global MAX_PAGINAS
    
    dialog = tk.Toplevel(root)
    dialog.title("Configurar Páginas")
    dialog.geometry("350x200")
    dialog.configure(bg="#1e1e1e")
    dialog.transient(root)
    dialog.grab_set()
    
    # Centrar el diálogo
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
    y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"+{x}+{y}")
    
    tk.Label(dialog, text="Configurar búsqueda en múltiples páginas", 
             bg="#1e1e1e", fg="#2ecc71", font=("Arial", 12, "bold")).pack(pady=10)
    
    tk.Label(dialog, text="¿Cuántas páginas máximo deseas revisar?", 
             bg="#1e1e1e", fg="#ffffff", font=("Arial", 10)).pack(pady=5)
    
    # Variable para almacenar el valor
    var_paginas = tk.IntVar(value=MAX_PAGINAS)
    
    frame_input = tk.Frame(dialog, bg="#1e1e1e")
    frame_input.pack(pady=10)
    
    tk.Label(frame_input, text="Páginas:", bg="#1e1e1e", fg="#ffffff").pack(side=tk.LEFT)
    spinbox = tk.Spinbox(frame_input, from_=1, to=20, textvariable=var_paginas, 
                         width=5, bg="#2e2e2e", fg="#ffffff", insertbackground="#ffffff")
    spinbox.pack(side=tk.LEFT, padx=(5, 0))
    
    tk.Label(dialog, text="⚠️ Más páginas = más tiempo de búsqueda", 
             bg="#1e1e1e", fg="#ffa500", font=("Arial", 9)).pack(pady=5)
    
    def guardar_configuracion():
        global MAX_PAGINAS
        MAX_PAGINAS = var_paginas.get()
        dialog.destroy()
        messagebox.showinfo("Configuración", f"Se revisarán máximo {MAX_PAGINAS} páginas en la próxima búsqueda.")
    
    def cancelar():
        dialog.destroy()
    
    frame_botones = tk.Frame(dialog, bg="#1e1e1e")
    frame_botones.pack(pady=20)
    
    btn_guardar = tk.Button(frame_botones, text="✅ Guardar", command=guardar_configuracion,
                           bg="#2ecc71", fg="#1e1e1e", font=("Arial", 10, "bold"), 
                           padx=15, pady=5)
    btn_guardar.pack(side=tk.LEFT, padx=(0, 10))
    
    btn_cancelar = tk.Button(frame_botones, text="❌ Cancelar", command=cancelar,
                            bg="#e74c3c", fg="#ffffff", font=("Arial", 10, "bold"), 
                            padx=15, pady=5)
    btn_cancelar.pack(side=tk.LEFT)

def cargar_excel():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],
        title="Selecciona un archivo Excel"
    )
    if filepath:
        try:
            # Verificar si el archivo existe
            if not os.path.exists(filepath):
                raise FileNotFoundError(f"El archivo no existe: {filepath}")
            
            # Mostrar progreso
            progress_label.config(text="Cargando archivo Excel...")
            root.update()
            
            # Primero intentar leer las hojas disponibles
            try:
                excel_file = pd.ExcelFile(filepath)
                hojas_disponibles = excel_file.sheet_names
                text_resultados.delete("1.0", tk.END)
                text_resultados.insert(tk.END, f"Hojas disponibles en el archivo: {hojas_disponibles}\n")
                
                # Verificar si existe la hoja 'Ranking'
                if 'Ranking' not in hojas_disponibles:
                    # Preguntar al usuario qué hoja usar
                    hoja_seleccionada = None
                    if hojas_disponibles:
                        respuesta = messagebox.askyesno(
                            "Hoja no encontrada", 
                            f"No se encontró la hoja 'Ranking'.\n"
                            f"Hojas disponibles: {', '.join(hojas_disponibles)}\n\n"
                            f"¿Deseas usar la primera hoja disponible: '{hojas_disponibles[0]}'?"
                        )
                        if respuesta:
                            hoja_seleccionada = hojas_disponibles[0]
                        else:
                            progress_label.config(text="Operación cancelada")
                            return
                    else:
                        raise ValueError("No se encontraron hojas en el archivo Excel")
                else:
                    hoja_seleccionada = 'Ranking'
                
                # Cargar la hoja seleccionada
                df = pd.read_excel(filepath, sheet_name=hoja_seleccionada)
                text_resultados.insert(tk.END, f"Usando hoja: {hoja_seleccionada}\n")
                text_resultados.insert(tk.END, f"Columnas disponibles: {list(df.columns)}\n\n")
                
            except Exception as e:
                # Si falla la lectura con ExcelFile, intentar lectura directa
                text_resultados.insert(tk.END, f"Advertencia al leer hojas: {e}\n")
                text_resultados.insert(tk.END, "Intentando cargar hoja 'Ranking' directamente...\n")
                df = pd.read_excel(filepath, sheet_name='Ranking')

            # Verificar si la columna 'Placa' existe
            columnas_disponibles = list(df.columns)
            columna_placa = None
            
            # Buscar columna que contenga 'Placa' (sin importar mayúsculas/minúsculas)
            for col in columnas_disponibles:
                if 'placa' in str(col).lower():
                    columna_placa = col
                    break
            
            if columna_placa is None:
                # Mostrar columnas disponibles y preguntar al usuario
                columnas_texto = '\n'.join([f"- {col}" for col in columnas_disponibles])
                messagebox.showerror(
                    "Columna no encontrada", 
                    f"No se encontró una columna llamada 'Placa'.\n\n"
                    f"Columnas disponibles:\n{columnas_texto}\n\n"
                    f"Por favor, verifica que tu archivo tenga una columna llamada 'Placa'"
                )
                progress_label.config(text="Error: Columna 'Placa' no encontrada")
                return

            # Extraer las placas
            placas_series = df[columna_placa].dropna()  # Eliminar valores vacíos
            placas = placas_series.astype(str).str.strip().str.upper().tolist()
            
            # Filtrar placas vacías o que solo contengan espacios
            placas = [p for p in placas if p and p != 'NAN' and p.strip()]

            # Mostrar en el área de texto
            text_area.delete("1.0", tk.END)
            text_area.insert(tk.END, "\n".join(placas))
            
            # Mostrar información de carga
            text_resultados.insert(tk.END, f"✅ Cargadas {len(placas)} placas válidas desde la columna '{columna_placa}'\n")
            text_resultados.insert(tk.END, f"Archivo: {os.path.basename(filepath)}\n")
            text_resultados.insert(tk.END, "=" * 50 + "\n")
            
            progress_label.config(text=f"Archivo cargado exitosamente - {len(placas)} placas")
            messagebox.showinfo("Archivo cargado", f"Cargadas {len(placas)} placas desde la columna '{columna_placa}'.")
            
        except FileNotFoundError as e:
            messagebox.showerror("Error", f"Archivo no encontrado:\n{e}")
            progress_label.config(text="Error: Archivo no encontrado")
        except PermissionError:
            messagebox.showerror("Error", "No se puede acceder al archivo.\n\nPosibles soluciones:\n- Cierra el archivo Excel si está abierto\n- Verifica los permisos del archivo\n- Ejecuta el programa como administrador")
            progress_label.config(text="Error: Sin permisos para acceder al archivo")
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{error_msg}")
            text_resultados.insert(tk.END, f"❌ Error al cargar archivo: {error_msg}\n")
            progress_label.config(text="Error al cargar archivo")

def buscar_placas():
    placas_usuario = text_area.get("1.0", tk.END).strip().splitlines()
    placas_usuario = [p.strip().upper() for p in placas_usuario if p.strip()]
    if not placas_usuario:
        messagebox.showwarning("Advertencia", "Por favor ingresa o carga al menos una placa.")
        return
    threading.Thread(target=ejecutar_busqueda, args=(placas_usuario,), daemon=True).start()

def ejecutar_busqueda(placas_usuario):
    global MAX_PAGINAS
    
    btn_buscar.config(state=tk.DISABLED)
    btn_cargar.config(state=tk.DISABLED)
    btn_configurar.config(state=tk.DISABLED)
    progress_label.config(text="Buscando placas...")
    
    text_resultados.delete("1.0", tk.END)
    text_resultados.insert(tk.END, f"🔍 Iniciando búsqueda de {len(placas_usuario)} placas...\n")
    text_resultados.insert(tk.END, "=" * 50 + "\n")

    try:
        # Verificar si existe chromedriver
        chromedriver_paths = ['./chromedriver.exe', './chromedriver', 'chromedriver.exe', 'chromedriver']
        chromedriver_path = None
        
        for path in chromedriver_paths:
            if os.path.exists(path):
                chromedriver_path = path
                break
        
        if chromedriver_path:
            service = Service(chromedriver_path)
        else:
            # Intentar usar chromedriver del PATH del sistema
            service = Service()
            
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # Ejecutar sin ventana
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        
        text_resultados.insert(tk.END, "🌐 Abriendo navegador...\n")
        text_resultados.see(tk.END)
        
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("https://anm.gov.co/informacion-atencion-minero-estado-aviso")

        text_resultados.insert(tk.END, "⏳ Esperando que cargue la página...\n")
        text_resultados.see(tk.END)

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.views-table.cols-6"))
        )
        time.sleep(2)

        encontrados = 0
        total_procesados = 0
        pagina_actual = 1
        
        text_resultados.insert(tk.END, f"📊 Comenzando búsqueda en múltiples páginas (máximo {MAX_PAGINAS} páginas)...\n")
        text_resultados.see(tk.END)

        while pagina_actual <= MAX_PAGINAS:
            text_resultados.insert(tk.END, f"\n📄 === PÁGINA {pagina_actual} ===\n")
            text_resultados.see(tk.END)
            
            # Buscar la tabla en la página actual
            try:
                tabla = driver.find_element(By.CSS_SELECTOR, "table.views-table.cols-6")
                filas = tabla.find_elements(By.TAG_NAME, "tr")[1:]  # Omitir encabezado
                
                text_resultados.insert(tk.END, f"   Analizando {len(filas)} registros en página {pagina_actual}...\n")
                text_resultados.see(tk.END)
                
                procesados_pagina = 0
                
                for i, fila in enumerate(filas):
                    if fila.is_displayed():
                        procesados_pagina += 1
                        total_procesados += 1
                        columnas = fila.find_elements(By.TAG_NAME, "td")
                        if len(columnas) >= 6:
                            placas_tabla = columnas[1].text.strip().upper()
                            try:
                                enlace = columnas[4].find_element(By.TAG_NAME, "a").get_attribute("href")
                            except:
                                enlace = "Sin enlace"

                            # Buscar coincidencias
                            for placa_usuario in placas_usuario:
                                if placa_usuario in placas_tabla:
                                    encontrados += 1
                                    # Obtener información adicional
                                    solicitante = columnas[0].text.strip() if len(columnas) > 0 else "N/A"
                                    municipio = columnas[2].text.strip() if len(columnas) > 2 else "N/A"
                                    
                                    texto = f"   ||\n"
                                    texto += f"🎯 COINCIDENCIA #{encontrados} (Página {pagina_actual})\n"
                                    texto += f"   Placa buscada: {placa_usuario}\n"
                                    texto += f"   Placa(s) en tabla: {placas_tabla}\n"
                                    texto += f"   Solicitante: {solicitante}\n"
                                    texto += f"   Municipio: {municipio}\n"
                                    texto += f"   📄 PDF: {enlace}\n"
                                    texto += "   " + "-" * 40 + "\n\n"
                                    
                                    text_resultados.insert(tk.END, texto)
                                    text_resultados.see(tk.END)
                                    
                                    # Abrir PDF si existe
                                    if enlace != "Sin enlace":
                                        try:
                                            webbrowser.open(enlace)
                                        except Exception as e:
                                            text_resultados.insert(tk.END, f"   ⚠️ Error al abrir PDF: {e}\n")

                        # Actualizar progreso cada 5 registros
                        if procesados_pagina % 5 == 0:
                            progress_label.config(text=f"Página {pagina_actual}/{MAX_PAGINAS} - Procesando {procesados_pagina}/{len(filas)} registros")
                            root.update()

                text_resultados.insert(tk.END, f"   ✅ Página {pagina_actual} completada: {procesados_pagina} registros procesados\n")
                
                # Intentar ir a la siguiente página
                if pagina_actual < MAX_PAGINAS:
                    text_resultados.insert(tk.END, f"   🔄 Navegando a página {pagina_actual + 1}...\n")
                    text_resultados.see(tk.END)
                    
                    try:
                        # Buscar el botón "Siguiente" o enlaces de paginación
                        siguiente_encontrado = False
                        
                        # Opción 1: Buscar botón "Siguiente" o "Next"
                        try:
                            boton_siguiente = driver.find_element(By.XPATH, "//a[contains(text(), 'siguiente') or contains(text(), 'Siguiente') or contains(text(), 'Next') or contains(text(), '>')]")
                            if boton_siguiente.is_enabled():
                                driver.execute_script("arguments[0].click();", boton_siguiente)
                                siguiente_encontrado = True
                        except:
                            pass
                        
                        # Opción 2: Buscar enlace con número de página
                        if not siguiente_encontrado:
                            try:
                                enlace_pagina = driver.find_element(By.XPATH, f"//a[contains(text(), '{pagina_actual + 1}')]")
                                driver.execute_script("arguments[0].click();", enlace_pagina)
                                siguiente_encontrado = True
                            except:
                                pass
                        
                        # Opción 3: Buscar en elementos de paginación comunes
                        if not siguiente_encontrado:
                            try:
                                paginacion = driver.find_element(By.CSS_SELECTOR, ".pager, .pagination, .page-navigation")
                                enlaces = paginacion.find_elements(By.TAG_NAME, "a")
                                for enlace in enlaces:
                                    if enlace.text.strip() == str(pagina_actual + 1) or "siguiente" in enlace.text.lower() or "next" in enlace.text.lower():
                                        driver.execute_script("arguments[0].click();", enlace)
                                        siguiente_encontrado = True
                                        break
                            except:
                                pass
                        
                        if siguiente_encontrado:
                            # Esperar a que cargue la nueva página
                            time.sleep(3)
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "table.views-table.cols-6"))
                            )
                            pagina_actual += 1
                        else:
                            text_resultados.insert(tk.END, f"   ⚠️ No se encontró botón para página siguiente. Terminando búsqueda.\n")
                            break
                            
                    except Exception as e:
                        text_resultados.insert(tk.END, f"   ⚠️ Error al navegar a siguiente página: {str(e)}\n")
                        text_resultados.insert(tk.END, f"   ℹ️ Continuando con las páginas encontradas hasta ahora...\n")
                        break
                else:
                    text_resultados.insert(tk.END, f"   ℹ️ Alcanzado límite de páginas ({MAX_PAGINAS})\n")
                    break
                    
            except Exception as e:
                text_resultados.insert(tk.END, f"   ❌ Error al procesar página {pagina_actual}: {str(e)}\n")
                break

        # Resumen final
        text_resultados.insert(tk.END, "\n" + "=" * 50 + "\n")
        text_resultados.insert(tk.END, f"📊 RESUMEN FINAL DE BÚSQUEDA:\n")
        text_resultados.insert(tk.END, f"   • Placas buscadas: {len(placas_usuario)}\n")
        text_resultados.insert(tk.END, f"   • Páginas revisadas: {pagina_actual - 1 if pagina_actual > 1 else 1}\n")
        text_resultados.insert(tk.END, f"   • Total registros procesados: {total_procesados}\n")
        text_resultados.insert(tk.END, f"   • Coincidencias encontradas: {encontrados}\n")
        
        if encontrados == 0:
            text_resultados.insert(tk.END, f"❌ No se encontraron placas coincidentes en {pagina_actual - 1 if pagina_actual > 1 else 1} página(s).\n")
            progress_label.config(text=f"Búsqueda completada - Sin resultados ({pagina_actual - 1 if pagina_actual > 1 else 1} páginas)")
        else:
            text_resultados.insert(tk.END, f"✅ Búsqueda completada exitosamente en {pagina_actual - 1 if pagina_actual > 1 else 1} página(s).\n")
            progress_label.config(text=f"Búsqueda completada - {encontrados} coincidencias en {pagina_actual - 1 if pagina_actual > 1 else 1} páginas")

        driver.quit()

    except Exception as e:
        error_msg = str(e)
        text_resultados.insert(tk.END, f"❌ ERROR: {error_msg}\n")
        
        # Sugerencias de solución
        if "chromedriver" in error_msg.lower():
            text_resultados.insert(tk.END, "\n💡 SOLUCIÓN:\n")
            text_resultados.insert(tk.END, "1. Descarga ChromeDriver desde: https://chromedriver.chromium.org/\n")
            text_resultados.insert(tk.END, "2. Coloca 'chromedriver.exe' en la misma carpeta que este programa\n")
            text_resultados.insert(tk.END, "3. O instala WebDriver Manager: pip install webdriver-manager\n")
        
        messagebox.showerror("Error", f"Ocurrió un error durante la búsqueda:\n{error_msg}")
        progress_label.config(text="Error en la búsqueda")

    finally:
        btn_buscar.config(state=tk.NORMAL)
        btn_cargar.config(state=tk.NORMAL)
        btn_configurar.config(state=tk.NORMAL)

# ==== Configuración de la ventana principal ====
root = tk.Tk()
root.title("🌒 Buscador de Placas ANM - Versión con Múltiples Páginas")
root.geometry("800x700")
root.configure(bg="#1e1e1e")

# ==== Estilos personalizados ====
style_config = {
    "bg": "#1e1e1e",
    "fg": "#ffffff",
    "insertbackground": "#ffffff",
    "selectbackground": "#2ecc71",
    "font": ("Consolas", 10)
}

btn_style = {
    "bg": "#2ecc71",
    "fg": "#1e1e1e",
    "activebackground": "#27ae60",
    "activeforeground": "#ffffff",
    "font": ("Arial", 10, "bold"),
    "bd": 0,
    "padx": 15,
    "pady": 8
}

label_style = {
    "bg": "#1e1e1e",
    "fg": "#2ecc71",
    "font": ("Arial", 11, "bold")
}

progress_style = {
    "bg": "#1e1e1e",
    "fg": "#ffffff",
    "font": ("Arial", 9)
}

# ==== Frame principal ====
main_frame = tk.Frame(root, bg="#1e1e1e")
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# ==== Sección de entrada ====
input_frame = tk.Frame(main_frame, bg="#1e1e1e")
input_frame.pack(fill=tk.X, pady=(0, 10))

tk.Label(input_frame, text="📝 Ingresa las placas (una por línea):", **label_style).pack(anchor=tk.W)
text_area = scrolledtext.ScrolledText(input_frame, width=60, height=8, **style_config)
text_area.pack(fill=tk.X, pady=5)

# ==== Botones ====
btn_frame = tk.Frame(main_frame, bg="#1e1e1e")
btn_frame.pack(fill=tk.X, pady=5)

btn_cargar = tk.Button(btn_frame, text="📂 Cargar Excel", command=cargar_excel, **btn_style)
btn_cargar.pack(side=tk.LEFT, padx=(0, 10))

btn_buscar = tk.Button(btn_frame, text="🔍 Buscar Placas", command=buscar_placas, **btn_style)
btn_buscar.pack(side=tk.LEFT, padx=(0, 10))

btn_configurar = tk.Button(btn_frame, text="⚙️ Configurar Páginas", command=configurar_max_paginas, **btn_style)
btn_configurar.pack(side=tk.LEFT)

# ==== Barra de progreso ====
progress_label = tk.Label(main_frame, text="Listo para cargar archivo", **progress_style)
progress_label.pack(anchor=tk.W, pady=(10, 0))

# ==== Área de resultados ====
result_frame = tk.Frame(main_frame, bg="#1e1e1e")
result_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

tk.Label(result_frame, text="📊 Resultados:", **label_style).pack(anchor=tk.W)
text_resultados = scrolledtext.ScrolledText(result_frame, width=90, height=20, **style_config)
text_resultados.pack(fill=tk.BOTH, expand=True, pady=5)

# ==== Información inicial ====
texto_inicial = """🌒 BUSCADOR DE PLACAS ANM - VERSIÓN 1.1 MÚLTIPLES PÁGINAS

📋 INSTRUCCIONES:

1. Haz clic en "📂 Cargar Excel" para cargar un archivo Excel
2. El programa buscará la hoja llamada 'Ranking' y la columna 'Placa'
3. También puedes escribir placas manualmente en el área de texto superior
4. Usa "⚙️ Configurar Páginas" para establecer cuántas páginas revisar (1-20)
5. Haz clic en "🔍 Buscar Placas" para iniciar la búsqueda


🔧 REQUISITOS:

- ChromeDriver debe estar en la misma carpeta que este programa
- Conexión a internet para acceder al sitio web de ANM

⚠️ IMPORTANTE:

- Más páginas = más tiempo de búsqueda
- El programa se detiene si no encuentra más páginas
- Se muestra progreso en tiempo real

¡Listo para comenzar la búsqueda multipágina!
"""

text_resultados.insert(tk.END, texto_inicial)

# ==== Ejecutar ====
root.mainloop() 