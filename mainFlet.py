import flet as ft
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
import time
import threading
import os
import re
import logging

# =========================================================================
# CONFIGURACIÓN
# =========================================================================
URL_LOGIN = "https://staging-na01-equinox.demandware.net/on/demandware.store/Sites-Site/es%3bsite%3dMarathonChile/ViewApplication-DisplayWelcomePage"
ID_INPUT_USER = "idToken2"
ID_INPUT_PASSWORD = "idToken2"
ID_BTN_LOGIN = "loginButton_0"
XPATH_PRODUCTOS = '//*[@id="bm_content_column"]/table/tbody/tr/td/table/tbody/tr/td[2]/div[7]/div/ul/a[1]/li/div/div[2]'
XPATH_CAMPO_BUSQUEDA = '//*[@id="WFSimpleSearch_NameOrID"]'
XPATH_BTN_BUSCAR = '//*[@id="SimpleDiv"]/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/button'
XPATH_PRIMER_RESULTADO = '//td[@class="table_detail middle e s"]/a[@class="table_detail_link"]'

STOP_FLAG = False
RUNNING_THREAD = None
PARTIAL_RESULTS = []
START_TIME = None
TOTAL_CODES = 0

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def guardar_excel(datos):
    if not datos:
        logging.info("No hay datos para guardar en Excel.")
        return

    df = pd.DataFrame(datos, columns=["Artículo", "Imágenes"])
    fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_archivo = f"Salesforce_control_imagenes_{fecha_hora}.xlsx"
    ruta_archivo = os.path.join(os.getcwd(), nombre_archivo)
    df.to_excel(ruta_archivo, index=False)
    logging.info(f"Datos guardados en: {ruta_archivo}")

def wait_for_element(driver, locator, timeout=15):
    """
    Espera hasta que un elemento esté presente y visible.
    """
    try:
        return WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located(locator)
        )
    except Exception as e:
        logging.error(f"Error esperando elemento: {locator}. Detalles: {e}")
        return None

def process_codes(codigos_productos, username, password, page):
    global STOP_FLAG, PARTIAL_RESULTS

    driver = None
    try:
        driver = webdriver.Chrome()
        driver.get(URL_LOGIN)
        time.sleep(2)

        # Login
        user_input = wait_for_element(driver, (By.XPATH, f"//input[@id='{ID_INPUT_USER}' and @type='text']"))
        if user_input:
            user_input.clear()
            user_input.send_keys(username)
        else:
            raise Exception("No se encontró el campo de usuario.")

        login_btn_1 = wait_for_element(driver, (By.ID, ID_BTN_LOGIN))
        if login_btn_1:
            login_btn_1.click()
        else:
            raise Exception("No se encontró el botón de login (paso 1).")

        pwd_input = wait_for_element(driver, (By.XPATH, f"//input[@id='{ID_INPUT_PASSWORD}' and @type='password']"))
        if pwd_input:
            pwd_input.clear()
            pwd_input.send_keys(password)
        else:
            raise Exception("No se encontró el campo de contraseña.")

        login_btn_2 = wait_for_element(driver, (By.ID, ID_BTN_LOGIN))
        if login_btn_2:
            login_btn_2.click()
        else:
            raise Exception("No se encontró el botón de login (paso 2).")

        time.sleep(30)  # Incrementado el tiempo de espera para aceptar la notificación del celular

        productos_boton = wait_for_element(driver, (By.XPATH, XPATH_PRODUCTOS), timeout=20)
        if productos_boton:
            productos_boton.click()
        else:
            raise Exception("No se encontró el botón de productos.")

        time.sleep(5)

        total = len(codigos_productos)
        for i, codigo in enumerate(codigos_productos):
            if STOP_FLAG:
                logging.info("Stop solicitado. Saliendo del bucle.")
                break

            try:
                logging.info(f"Procesando código {i+1}/{total}: {codigo}")
                campo_busqueda = wait_for_element(driver, (By.XPATH, XPATH_CAMPO_BUSQUEDA))
                if campo_busqueda:
                    campo_busqueda.clear()
                    campo_busqueda.send_keys(codigo)
                else:
                    raise Exception(f"No se encontró el campo de búsqueda para el código: {codigo}.")

                boton_buscar = wait_for_element(driver, (By.XPATH, XPATH_BTN_BUSCAR))
                if boton_buscar:
                    boton_buscar.click()
                else:
                    raise Exception(f"No se encontró el botón de búsqueda para el código: {codigo}.")

                time.sleep(5)

                try:
                    primer_resultado = wait_for_element(driver, (By.XPATH, XPATH_PRIMER_RESULTADO))
                    if primer_resultado:
                        primer_resultado.click()
                    else:
                        raise Exception(f"No se encontró el primer resultado para el código: {codigo}.")
                except Exception as e:
                    logging.warning(f"No se encontró resultado para el código: {codigo}. Detalles: {e}")
                    PARTIAL_RESULTS.append([codigo, "Sin resultados"])
                    driver.back()
                    time.sleep(3)
                    continue

                texto_pagina = driver.find_element(By.TAG_NAME, 'body').text
                match = re.search(r'Imágenes:\s*(.*?)\n', texto_pagina, re.DOTALL)
                if match:
                    imagenes_texto = match.group(1).strip()
                else:
                    logging.warning(f"No se encontró la sección 'Imágenes' para el código: {codigo}")
                    imagenes_texto = "No se encontró la sección de 'Imágenes'."

                PARTIAL_RESULTS.append([codigo, imagenes_texto])
                driver.back()
                time.sleep(3)

            except Exception as e:
                logging.error(f"Error procesando código {codigo}: {e}")
                PARTIAL_RESULTS.append([codigo, "Error al procesar"])
                driver.back()
                time.sleep(3)

        guardar_excel(PARTIAL_RESULTS)

    except Exception as e:
        logging.error(f"Error general: {e}")
    finally:
        if driver:
            driver.quit()

    page.update()

def main(page: ft.Page):
    global STOP_FLAG, RUNNING_THREAD, PARTIAL_RESULTS, START_TIME, TOTAL_CODES

    page.title = "Control de Imágenes"
    page.window.width = 800
    page.window.height = 600

    username = ft.TextField(label="Usuario", width=400)
    password = ft.TextField(label="Contraseña", password=True, width=400)
    input_codes = ft.TextField(label="Códigos de productos (separados por espacios)", multiline=True, width=400, height=150)
    progress = ft.ProgressBar(width=400, height=20)
    log_output = ft.TextField(label="Log", multiline=True, width=400, height=150)

    def start_process(e):
        global STOP_FLAG, PARTIAL_RESULTS, START_TIME, TOTAL_CODES, RUNNING_THREAD

        codigos = input_codes.value.split()
        codigos = list(dict.fromkeys(codigos))

        if not username.value or not password.value:
            log_output.value += "Usuario y contraseña son requeridos.\n"
            page.update()
            return

        if not codigos:
            log_output.value += "No se ingresaron códigos.\n"
            page.update()
            return

        STOP_FLAG = False
        PARTIAL_RESULTS.clear()
        START_TIME = time.time()
        TOTAL_CODES = len(codigos)

        def process():
            process_codes(codigos, username.value, password.value, page)

        RUNNING_THREAD = threading.Thread(target=process)
        RUNNING_THREAD.start()

    def stop_process(e):
        global STOP_FLAG
        STOP_FLAG = True
        log_output.value += "Proceso detenido.\n"
        page.update()

    page.add(
        ft.Column([
            username,
            password,
            input_codes,
            ft.Row([
                ft.ElevatedButton("Iniciar", on_click=start_process),
                ft.ElevatedButton("Detener", on_click=stop_process),
            ]),
            progress,
            log_output
        ])
    )

ft.app(target=main, assets_dir="assets")
