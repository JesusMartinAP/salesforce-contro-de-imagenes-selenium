import tkinter as tk
from tkinter import messagebox
import time
import re
import pandas as pd
import os
from datetime import datetime

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

# IDs/XPATH del formulario en dos pasos:
ID_INPUT_USER = "idToken2"       # Primer input (type='text') para usuario
ID_INPUT_PASSWORD = "idToken2"   # Segundo input (type='password')
ID_BTN_LOGIN = "loginButton_0"   # Botón Log in (mismo ID en ambos pasos)

# XPATH del botón "Productos":
XPATH_PRODUCTOS = '//*[@id="bm_content_column"]/table/tbody/tr/td/table/tbody/tr/td[2]/div[7]/div/ul/a[1]/li/div/div[2]'

# XPATH para el buscador y el botón buscar
XPATH_CAMPO_BUSQUEDA = '//*[@id="WFSimpleSearch_NameOrID"]'
XPATH_BTN_BUSCAR = '//*[@id="SimpleDiv"]/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/button'

# XPATH para el primer resultado en la tabla
XPATH_PRIMER_RESULTADO = '//td[@class="table_detail middle e s"]/a[@class="table_detail_link"]'

# =========================================================================
# FUNCIONES
# =========================================================================

def guardar_excel(datos):
    """
    Guarda en un archivo Excel la lista de datos (código, imágenes).
    Requiere 'openpyxl': pip install openpyxl
    """
    df = pd.DataFrame(datos, columns=["Artículo", "Imágenes"])
    fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_archivo = f"Salesforce_control_imagenes_{fecha_hora}.xlsx"
    ruta_archivo = os.path.join(os.getcwd(), nombre_archivo)
    df.to_excel(ruta_archivo, index=False)
    print(f"Datos guardados en: {ruta_archivo}")

def procesar_codigos(codigos_productos):
    """
    - Abre la URL de login.
    - (Paso 1) Ingresar usuario + clic en botón.
    - (Paso 2) Ingresar password + clic en botón.
    - Clic en 'Productos'.
    - Búsqueda de cada código -> Extracción de 'Imágenes:'.
    - Retorna [ [codigo, texto_imagenes], ... ].
    """
    datos_productos = []
    driver = webdriver.Chrome()  # O webdriver.Firefox() si usas geckodriver
    
    try:
        # == PASO 1: LOGIN USUARIO ==
        print("Navegando a la página de login...")
        driver.get(URL_LOGIN)
        time.sleep(2)

        print("Esperando campo USUARIO (type='text')...")
        user_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located(
                (By.XPATH, f"//input[@id='{ID_INPUT_USER}' and @type='text']"))
        )
        user_input.clear()
        user_input.send_keys(USERNAME)
        print("Usuario ingresado.")

        print("Clic en botón 'Log in' (primer paso)...")
        login_btn_1 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, ID_BTN_LOGIN))
        )
        login_btn_1.click()

        # == PASO 2: LOGIN PASSWORD ==
        print("Esperando campo PASSWORD (type='password')...")
        pwd_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located(
                (By.XPATH, f"//input[@id='{ID_INPUT_PASSWORD}' and @type='password']"))
        )
        pwd_input.clear()
        pwd_input.send_keys(PASSWORD)
        print("Password ingresado.")

        print("Clic en botón 'Log in' (segundo paso)...")
        login_btn_2 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, ID_BTN_LOGIN))
        )
        login_btn_2.click()

        # Esperar que la página cargue
        time.sleep(5)
        print("URL actual después de login:", driver.current_url)

        # == CLIC EN “PRODUCTOS” ==
        print("Buscando botón 'Productos'...")
        productos_boton = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_PRODUCTOS))
        )
        productos_boton.click()
        time.sleep(3)
        print("Entramos a 'Productos'. URL actual:", driver.current_url)

        # == ITERAR CÓDIGOS ==
        for codigo in codigos_productos:
            try:
                print(f"\n>> Procesando código: {codigo}")

                # Campo de búsqueda
                campo_busqueda = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_BUSQUEDA))
                )
                campo_busqueda.clear()
                campo_busqueda.send_keys(codigo)

                # Botón buscar
                boton_buscar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_BTN_BUSCAR))
                )
                boton_buscar.click()
                time.sleep(3)
                print("Búsqueda realizada, URL actual:", driver.current_url)

                # Primer resultado
                try:
                    primer_resultado = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, XPATH_PRIMER_RESULTADO))
                    )
                    primer_resultado.click()
                except:
                    print(f"No se encontró resultado para el código: {codigo}")
                    datos_productos.append([codigo, "Sin resultados"])
                    # Volvemos a la pantalla anterior
                    driver.back()
                    time.sleep(2)
                    continue

                time.sleep(3)
                print("Dentro de detalle. URL actual:", driver.current_url)

                # Extraer "Imágenes:"
                texto_pagina = driver.find_element(By.TAG_NAME, 'body').text
                match = re.search(r'Imágenes:\s*(.*?)\n', texto_pagina, re.DOTALL)
                if match:
                    imagenes_texto = match.group(1).strip()
                else:
                    imagenes_texto = "No se encontró la sección de 'Imágenes'."

                print(f"{codigo} -> {imagenes_texto}")
                datos_productos.append([codigo, imagenes_texto])

                # Volver atrás
                driver.back()
                time.sleep(2)
                
                # == ESPERAR a que el buscador reaparezca tras retroceder ==
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, XPATH_CAMPO_BUSQUEDA))
                )
                print("Regresamos a la pantalla de búsqueda correctamente.")

            except Exception as e:
                print(f"Error procesando código {codigo}: {e}")
                datos_productos.append([codigo, "Error al procesar"])
                driver.back()
                time.sleep(2)
                # Esperar a que la página se estabilice
                try:
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, XPATH_CAMPO_BUSQUEDA))
                    )
                except:
                    pass

    except Exception as e:
        print(f"Error general en login o proceso: {e}")
    finally:
        driver.quit()

    return datos_productos

def iniciar_proceso():
    """
    Lee los códigos del Text (Tkinter), llama a procesar_codigos
    y luego guarda resultados en Excel.
    """
    codigos_str = text_codigos.get("1.0", tk.END)
    codigos_productos = codigos_str.split()

    if not codigos_productos:
        messagebox.showwarning("Advertencia", "No se ingresaron códigos.")
        return

    datos_productos = procesar_codigos(codigos_productos)
    guardar_excel(datos_productos)
    messagebox.showinfo("Proceso finalizado", "El proceso ha concluido exitosamente.")

# =========================================================================
# INTERFAZ GRÁFICA - TKINTER
# =========================================================================
ventana = tk.Tk()
ventana.title("Control de Imágenes - Evitar fallo en segundo código")

lbl_instruccion = tk.Label(ventana, text="Pega aquí los códigos (separados por espacios):")
lbl_instruccion.pack(padx=10, pady=5)

text_codigos = tk.Text(ventana, width=60, height=10)
text_codigos.pack(padx=10, pady=5)

btn_iniciar = tk.Button(ventana, text="Iniciar proceso", command=iniciar_proceso, bg="lightblue")
btn_iniciar.pack(pady=10)

ventana.mainloop()
