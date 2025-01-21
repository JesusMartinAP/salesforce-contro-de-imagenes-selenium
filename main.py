import tkinter as tk
from tkinter import messagebox, ttk
import time
import re
import pandas as pd
import os
from datetime import datetime
import threading

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =========================================================================
# CONFIGURACIÓN
# =========================================================================

USERNAME = "roberto.solari@pe.aseyco.com"
PASSWORD = "Marathon128."

URL_LOGIN = "https://staging-na01-equinox.demandware.net/on/demandware.store/Sites-Site/es%3bsite%3dMarathonChile/ViewApplication-DisplayWelcomePage"

ID_INPUT_USER = "idToken2"       # Primer input (type='text') para usuario
ID_INPUT_PASSWORD = "idToken2"   # Segundo input (type='password')
ID_BTN_LOGIN = "loginButton_0"   # Botón Log in (mismo ID en ambos pasos)

XPATH_PRODUCTOS = '//*[@id="bm_content_column"]/table/tbody/tr/td/table/tbody/tr/td[2]/div[7]/div/ul/a[1]/li/div/div[2]'
XPATH_CAMPO_BUSQUEDA = '//*[@id="WFSimpleSearch_NameOrID"]'
XPATH_BTN_BUSCAR = '//*[@id="SimpleDiv"]/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/button'
XPATH_PRIMER_RESULTADO = '//td[@class="table_detail middle e s"]/a[@class="table_detail_link"]'

# =========================================================================
# VARIABLES GLOBALES PARA CONTROL DE HILO
# =========================================================================

STOP_FLAG = False        # Cuando es True, el hilo se detiene
RUNNING_THREAD = None    # Referencia al hilo en ejecución (si existe)
START_TIME = None        # Guardamos hora de inicio para cronómetro

PARTIAL_RESULTS = []     # Datos parciales [ [codigo, imagenes], ... ]
TOTAL_CODES = 0          # Número total de códigos a procesar

# =========================================================================
# FUNCIONES
# =========================================================================

def guardar_excel(datos):
    """
    Guarda en un archivo Excel la lista de datos (código, imágenes).
    """
    if not datos:
        print("No hay datos para guardar en Excel.")
        return

    df = pd.DataFrame(datos, columns=["Artículo", "Imágenes"])
    fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_archivo = f"Salesforce_control_imagenes_{fecha_hora}.xlsx"
    ruta_archivo = os.path.join(os.getcwd(), nombre_archivo)
    df.to_excel(ruta_archivo, index=False)
    print(f"Datos guardados en: {ruta_archivo}")

def process_codes_in_thread(codigos_productos):
    """
    Función que se ejecuta en un hilo aparte.
    Realiza login, busca cada código y extrae 'Imágenes:'.
    Va actualizando PARTIAL_RESULTS y la barra de progreso.
    Si STOP_FLAG se pone en True, se detiene y genera Excel parcial.
    """
    global STOP_FLAG, PARTIAL_RESULTS

    driver = None
    try:
        driver = webdriver.Chrome()
        driver.get(URL_LOGIN)
        time.sleep(2)

        # --- LOGIN (2 pasos) ---
        # Paso 1: Usuario
        user_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.XPATH, f"//input[@id='{ID_INPUT_USER}' and @type='text']"))
        )
        user_input.clear()
        user_input.send_keys(USERNAME)

        login_btn_1 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, ID_BTN_LOGIN))
        )
        login_btn_1.click()

        # Paso 2: Password
        pwd_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.XPATH, f"//input[@id='{ID_INPUT_PASSWORD}' and @type='password']"))
        )
        pwd_input.clear()
        pwd_input.send_keys(PASSWORD)

        login_btn_2 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, ID_BTN_LOGIN))
        )
        login_btn_2.click()

        time.sleep(5)

        # Clic en "Productos"
        productos_boton = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_PRODUCTOS))
        )
        productos_boton.click()
        time.sleep(3)

        total = len(codigos_productos)
        for i, codigo in enumerate(codigos_productos):
            # Chequeo de STOP_FLAG
            if STOP_FLAG:
                print("Stop solicitado. Saliendo del bucle.")
                break

            try:
                print(f"\n** Procesando código {i+1}/{total}: {codigo}")
                campo_busqueda = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_BUSQUEDA))
                )
                campo_busqueda.clear()
                campo_busqueda.send_keys(codigo)

                boton_buscar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_BTN_BUSCAR))
                )
                boton_buscar.click()
                time.sleep(3)

                try:
                    primer_resultado = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, XPATH_PRIMER_RESULTADO))
                    )
                    primer_resultado.click()
                except:
                    print(f"No se encontró resultado para el código: {codigo}")
                    PARTIAL_RESULTS.append([codigo, "Sin resultados"])
                    driver.back()
                    time.sleep(2)
                    # Actualizar progreso en la interfaz
                    update_progress(i+1, total)
                    continue

                time.sleep(3)

                # Extraer "Imágenes:"
                texto_pagina = driver.find_element(By.TAG_NAME, 'body').text
                match = re.search(r'Imágenes:\s*(.*?)\n', texto_pagina, re.DOTALL)
                if match:
                    imagenes_texto = match.group(1).strip()
                else:
                    imagenes_texto = "No se encontró la sección de 'Imágenes'."

                print(f"{codigo} -> {imagenes_texto}")
                PARTIAL_RESULTS.append([codigo, imagenes_texto])

                driver.back()
                time.sleep(2)
                
                # Esperar a que reaparezca el buscador
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, XPATH_CAMPO_BUSQUEDA))
                )

            except Exception as e:
                print(f"Error procesando código {codigo}: {e}")
                PARTIAL_RESULTS.append([codigo, "Error al procesar"])
                driver.back()
                time.sleep(2)
                # Esperar a que reaparezca el buscador
                try:
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, XPATH_CAMPO_BUSQUEDA))
                    )
                except:
                    pass

            # Actualizar barra de progreso y etiqueta
            update_progress(i+1, total)

        # Si se terminaron todos o se rompió por STOP_FLAG, guardar Excel
        print("Generando Excel con datos (final o parcial).")
        guardar_excel(PARTIAL_RESULTS)

    except Exception as e:
        print(f"Error general en login o proceso: {e}")
    finally:
        if driver:
            driver.quit()

    # Notificar fin
    on_process_finished()

def update_progress(actual, total):
    """
    Actualiza la barra de progreso y la etiqueta de estado.
    Se llama desde el hilo secundario, pero usamos 'event_generate'
    para actualizar la interfaz en el hilo principal.
    """
    # Guardamos estos valores en variables globales o en el widget
    # y luego disparamos un evento en el main thread.
    progress_data["value"] = actual
    progress_data["max"] = total
    # Etiqueta: "Procesado X de Y"
    progress_data["text"] = f"Procesado {actual}/{total}"

    # Disparamos un evento custom para que el main loop actualice la UI
    root.event_generate("<<ProgressEvent>>", when="tail")

def on_progress_event(event):
    """
    Manejador de evento custom '<<ProgressEvent>>'.
    Toma los valores globales 'progress_data' y actualiza los widgets.
    """
    current = progress_data["value"]
    mx = progress_data["max"]
    label_progress.config(text=progress_data["text"])
    progress_bar["maximum"] = mx
    progress_bar["value"] = current

def on_process_finished():
    """
    Se llama cuando el proceso (hilo) termina.
    Desactiva el flag de ejecución y habilita botones.
    """
    global RUNNING_THREAD, STOP_FLAG
    STOP_FLAG = False
    RUNNING_THREAD = None
    print("Proceso finalizado (hilo).")
    # Al terminar, re-habilitar el botón "Iniciar"
    btn_iniciar.config(state=tk.NORMAL)
    btn_stop.config(state=tk.DISABLED)

def start_process():
    """
    Inicia el hilo de procesamiento.
    """
    global STOP_FLAG, RUNNING_THREAD, PARTIAL_RESULTS, START_TIME, TOTAL_CODES

    codigos_str = text_codigos.get("1.0", tk.END)
    codigos_productos = codigos_str.split()
    codigos_productos = list(dict.fromkeys(codigos_productos))  # eliminar duplicados, si deseas

    if not codigos_productos:
        messagebox.showwarning("Advertencia", "No se ingresaron códigos.")
        return

    # Reiniciar variables
    STOP_FLAG = False
    PARTIAL_RESULTS.clear()
    START_TIME = time.time()
    TOTAL_CODES = len(codigos_productos)

    # Inhabilitar el botón "Iniciar" mientras corre
    btn_iniciar.config(state=tk.DISABLED)
    btn_stop.config(state=tk.NORMAL)

    # Barra de progreso a 0
    progress_bar["value"] = 0
    progress_bar["maximum"] = TOTAL_CODES
    label_progress.config(text="Procesado 0/0")

    # Crear hilo
    t = threading.Thread(target=process_codes_in_thread, args=(codigos_productos,))
    t.daemon = True
    t.start()
    RUNNING_THREAD = t

def stop_process():
    """
    Marca STOP_FLAG = True, el hilo se encargará de generar el Excel parcial y terminar.
    """
    global STOP_FLAG
    if RUNNING_THREAD:
        STOP_FLAG = True
        print("Stop solicitado por el usuario.")
    else:
        print("No hay proceso corriendo.")

def update_timer():
    """
    Actualiza cada segundo el tiempo transcurrido, si el hilo sigue en ejecución.
    """
    if RUNNING_THREAD is not None:
        elapsed = time.time() - START_TIME
        label_time.config(text=f"Tiempo transcurrido: {elapsed:.1f} s")
    else:
        # Si no hay hilo corriendo, el timer se queda con lo último o se resetea
        pass

    # Llamar de nuevo en 1s
    root.after(1000, update_timer)

# =========================================================================
# INTERFAZ TKINTER
# =========================================================================

root = tk.Tk()
root.title("Control de Imágenes - Barra Progreso, Stop, Cronómetro")

# Aplicar un estilo "clam" o similar para verse más moderno
style = ttk.Style(root)
style.theme_use("clam")  # prueba "alt", "default", "clam", "vista", etc.

# Diccionario para compartir datos de progreso (desde el hilo -> main)
progress_data = {"value": 0, "max": 0, "text": ""}

# Eventos custom para actualizar la interfaz desde el hilo
root.bind("<<ProgressEvent>>", on_progress_event)

frame_top = ttk.Frame(root, padding=10)
frame_top.pack(fill=tk.BOTH, expand=True)

lbl_instruccion = ttk.Label(frame_top, text="Pega aquí los códigos (separados por espacios):")
lbl_instruccion.pack(padx=5, pady=5, anchor="w")

text_codigos = tk.Text(frame_top, width=60, height=8)
text_codigos.pack(padx=5, pady=5, fill=tk.X)

frame_buttons = ttk.Frame(root, padding=10)
frame_buttons.pack(fill=tk.X, expand=False)

btn_iniciar = ttk.Button(frame_buttons, text="Iniciar proceso", command=start_process)
btn_iniciar.pack(side=tk.LEFT, padx=5)

btn_stop = ttk.Button(frame_buttons, text="Detener", command=stop_process)
btn_stop.pack(side=tk.LEFT, padx=5)
btn_stop.config(state=tk.DISABLED)  # desactivado al inicio

# Barra de progreso y label
frame_progress = ttk.Frame(root, padding=10)
frame_progress.pack(fill=tk.X, expand=False)

label_progress = ttk.Label(frame_progress, text="Progreso...")
label_progress.pack(padx=5, pady=5, anchor="w")

progress_bar = ttk.Progressbar(frame_progress, orient=tk.HORIZONTAL, length=400, mode="determinate")
progress_bar.pack(padx=5, pady=5, fill=tk.X)

# Label para mostrar el tiempo transcurrido
label_time = ttk.Label(root, text="Tiempo transcurrido: 0.0 s")
label_time.pack(padx=5, pady=5)

# Iniciar el actualizador de cronómetro
update_timer()

root.mainloop()
