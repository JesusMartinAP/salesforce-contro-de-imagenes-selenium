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
# Primer paso: <input type="text" id="idToken2" placeholder="User Name">
#              <input id="loginButton_0" type="submit" value="Log in">
# Segundo paso: <input type="password" id="idToken2" placeholder="Password">
#               <input id="loginButton_0" type="submit" value="Log in">
# Si el botón cambiara de ID en el segundo paso, ajusta en el script.

ID_INPUT_USER = "idToken2"       # Primer input (type="text") para usuario
ID_INPUT_PASSWORD = "idToken2"   # Segundo input (type="password"), mismo ID pero distinto type
ID_BTN_LOGIN = "loginButton_0"   # Botón 'Log in' (mismo ID en ambos pasos)

# Una vez logueado, haz clic en “Productos”:
XPATH_PRODUCTOS = '//*[@id="bm_content_column"]/table/tbody/tr/td/table/tbody/tr/td[2]/div[7]/div/ul/a[1]/li/div/div[2]'

# Búsqueda de códigos
XPATH_CAMPO_BUSQUEDA = '//*[@id="WFSimpleSearch_NameOrID"]'
XPATH_BTN_BUSCAR = '//*[@id="SimpleDiv"]/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/button'
XPATH_PRIMER_RESULTADO = '//td[@class="table_detail middle e s"]/a[@class="table_detail_link"]'

# =========================================================================
# FUNCIONES
# =========================================================================

def guardar_excel(datos):
    """
    Guarda en un archivo Excel la lista de datos (código, imágenes).
    """
    df = pd.DataFrame(datos, columns=["Artículo", "Imágenes"])
    fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_archivo = f"Salesforce_control_imagenes_{fecha_hora}.xlsx"
    ruta_archivo = os.path.join(os.getcwd(), nombre_archivo)
    df.to_excel(ruta_archivo, index=False)  # requiere: pip install openpyxl
    print(f"Datos guardados en: {ruta_archivo}")

def procesar_codigos(codigos_productos):
    """
    - Abre la URL de login.
    - Primer paso: ingresar usuario + clic en botón.
    - Segundo paso: esperar campo password + ingresar password + clic en botón.
    - Clic en “Productos”.
    - Búsqueda de cada código: extracción de 'Imágenes:'.
    - Devuelve lista con [codigo, texto_imagenes].
    """
    datos_productos = []
    driver = webdriver.Chrome()  # o webdriver.Firefox()

    try:
        # 1) Ir a la página de login (primer paso)
        driver.get(URL_LOGIN)
        time.sleep(2)

        # --- PASO 1: INGRESAR USUARIO ---
        print("Esperando campo de USUARIO (type='text')...")
        user_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.XPATH, f"//input[@id='{ID_INPUT_USER}' and @type='text']"))
        )
        user_input.clear()
        user_input.send_keys(USERNAME)
        print("Usuario ingresado.")

        # Botón de login (primer clic)
        print("Haciendo clic en botón 'Log in' (primer paso)...")
        login_btn_1 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, ID_BTN_LOGIN))
        )
        login_btn_1.click()

        # --- PASO 2: INGRESAR PASSWORD ---
        print("Esperando campo de PASSWORD (type='password')...")
        pwd_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.XPATH, f"//input[@id='{ID_INPUT_PASSWORD}' and @type='password']"))
        )
        pwd_input.clear()
        pwd_input.send_keys(PASSWORD)
        print("Password ingresado.")

        print("Haciendo clic en botón 'Log in' (segundo paso)...")
        login_btn_2 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, ID_BTN_LOGIN))
        )
        login_btn_2.click()

        # Esperar que la página cargue tras login
        time.sleep(5)

        # --- CLIC EN “PRODUCTOS” ---
        print("Buscando botón 'Productos'...")
        productos_boton = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_PRODUCTOS))
        )
        productos_boton.click()
        time.sleep(3)
        print("Entramos a 'Productos'.")

        # --- ITERAR CÓDIGOS ---
        for codigo in codigos_productos:
            try:
                print(f"\n** Procesando código: {codigo}")

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

                # Primer resultado
                try:
                    primer_resultado = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, XPATH_PRIMER_RESULTADO))
                    )
                    primer_resultado.click()
                except:
                    print(f"No se encontró resultado para el código: {codigo}")
                    datos_productos.append([codigo, "Sin resultados"])
                    driver.back()
                    time.sleep(2)
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
                datos_productos.append([codigo, imagenes_texto])

                # Regresar atrás
                driver.back()
                time.sleep(2)

            except Exception as e:
                print(f"Error procesando código {codigo}: {e}")
                datos_productos.append([codigo, "Error al procesar"])
                driver.back()
                time.sleep(2)

    except Exception as e:
        print(f"Error general en login o proceso: {e}")
    finally:
        driver.quit()

    return datos_productos

def iniciar_proceso():
    """
    Lee los códigos del Text de Tkinter, llama a procesar_codigos,
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
ventana.title("Control de Imágenes - Login en 2 pasos (Refinado)")

lbl_instruccion = tk.Label(ventana, text="Pega aquí los códigos (separados por espacios):")
lbl_instruccion.pack(padx=10, pady=5)

text_codigos = tk.Text(ventana, width=60, height=10)
text_codigos.pack(padx=10, pady=5)

btn_iniciar = tk.Button(ventana, text="Iniciar proceso", command=iniciar_proceso, bg="lightblue")
btn_iniciar.pack(pady=10)

ventana.mainloop()
