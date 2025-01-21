import tkinter as tk
from tkinter import messagebox
import time
import re
import pandas as pd
import os
from datetime import datetime

# --- Importaciones de Selenium ---
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

def guardar_excel(datos):
    """
    Guarda en un archivo Excel la lista de datos proporcionada.
    """
    df = pd.DataFrame(datos, columns=["Artículo", "Imágenes"])
    fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_archivo = f"Salesforce_control_imagenes_{fecha_hora}.xlsx"
    ruta_archivo = os.path.join(os.getcwd(), nombre_archivo)
    df.to_excel(ruta_archivo, index=False)
    print(f"Datos guardados en: {ruta_archivo}")

def procesar_codigos(codigos_productos):
    """
    1) Entra al enlace inicial.
    2) Hace clic en 'Productos'.
    3) Pega el código en el buscador.
    4) Hace clic en 'Buscar'.
    5) Hace clic en el primer resultado.
    6) Extrae la sección 'Imágenes:' (o el texto requerido).
    7) Retorna al menú principal o continua.
    8) Devuelve la lista con resultados para cada código.
    """

    datos_productos = []
    
    # Inicia el driver de Selenium (ejemplo con Chrome)
    driver = webdriver.Chrome()
    
    # 1) Abre la URL inicial (staging):
    url_inicial = "https://staging-na01-equinox.demandware.net/on/demandware.store/Sites-Site/es%3bsite%3dMarathonChile/ViewApplication-DisplayWelcomePage?csrf_token=hiFVEc0BXyMfT7BnJh0uDL-eKXWJhcbFRP98TI6UAnBSEpyA8dzNWB0ogkjkobIeZuCsdheanqNl3m-msfzI8Nwl3lHvKui1QIwx3AbTpAej0VluhXsVx5-E1PduptJMuBLu_K-0lCW9vznIEXjkvMSekftRWKrPPvM8_IxGqfhVaRMti9w="
    driver.get(url_inicial)
    
    # Esperamos unos segundos para que cargue (ajusta según tu conexión/sitio)
    time.sleep(5)

    try:
        # 2) Hacer clic en “Productos”
        # XPATH que pasaste: //*[@id="bm_content_column"]/table/tbody/tr/td/table/tbody/tr/td[2]/div[7]/div/ul/a[1]/li/div/div[2]
        # OJO: A veces conviene un XPATH más corto o un wait explicito
        productos_boton = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="bm_content_column"]/table/tbody/tr/td/table/tbody/tr/td[2]/div[7]/div/ul/a[1]/li/div/div[2]'))
        )
        productos_boton.click()
        time.sleep(3)

        # Iteramos cada código
        for codigo in codigos_productos:
            # 3) Pegar el código en el buscador
            # XPATH: //*[@id="WFSimpleSearch_NameOrID"]
            campo_busqueda = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="WFSimpleSearch_NameOrID"]'))
            )
            campo_busqueda.clear()
            campo_busqueda.send_keys(codigo)

            # 4) Clic en el botón Buscar
            # XPATH: //*[@id="SimpleDiv"]/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/button
            boton_buscar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="SimpleDiv"]/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/button'))
            )
            boton_buscar.click()
            time.sleep(3)

            # 5) Hacer clic en el primer resultado de la tabla
            # Por ejemplo: Buscar la primera fila con <a class="table_detail_link"> 
            # O tu XPATH directamente a la celda que compartiste:
            try:
                primer_resultado = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//td[@class="table_detail middle e s"]/a[@class="table_detail_link"]'))
                )
                primer_resultado.click()
            except:
                print(f"No se encontró ningún resultado para el código: {codigo}")
                datos_productos.append([codigo, "Sin resultados"])
                # Volvemos a la página de 'Productos' para el siguiente código (opcional)
                driver.back()
                time.sleep(2)
                continue

            time.sleep(3)

            # 6) Extraer texto de la página y buscar la sección “Imágenes:”
            texto_pagina = driver.find_element(By.TAG_NAME, 'body').text
            match = re.search(r'Imágenes:\s*(.*?)\n', texto_pagina, re.DOTALL)
            if match:
                imagenes_texto = match.group(1).strip()
            else:
                imagenes_texto = "No se encontró la sección de 'Imágenes'."

            print(f"{codigo} -> {imagenes_texto}")
            datos_productos.append([codigo, imagenes_texto])

            # 7) Regresar o continuar al menú principal / Productos.
            # Dependiendo de cómo funcione tu sitio, puedes:
            #    - Hacer clic en un botón "Volver"
            #    - Usar driver.back() dos veces
            #    - Ir directamente a la URL de Productos
            driver.back()  # Regresa a la página de detalles del listado
            time.sleep(2)
            # Si es necesario dar un back adicional para volver al menú "Productos", hazlo
            # driver.back()
            # time.sleep(2)

        # Al terminar, cerramos el navegador
        driver.quit()
    
    except Exception as e:
        print(f"Ocurrió un error: {e}")
        driver.quit()
        guardar_excel(datos_productos)  # guardamos lo que llevemos
        return datos_productos

    return datos_productos

def iniciar_proceso():
    """
    Función que se ejecuta al presionar el botón.
    Lee los códigos del Text (Tkinter), ejecuta procesar_codigos y luego
    guarda resultados en Excel.
    """
    codigos_str = text_codigos.get("1.0", tk.END)
    codigos_productos = codigos_str.split()
    
    if not codigos_productos:
        messagebox.showwarning("Advertencia", "No se ingresaron códigos.")
        return

    datos_productos = procesar_codigos(codigos_productos)
    guardar_excel(datos_productos)
    messagebox.showinfo("Proceso finalizado", "El proceso ha concluido exitosamente.")

# ---------------------- INTERFAZ GRÁFICA (TKINTER) ----------------------
ventana = tk.Tk()
ventana.title("Control de Imágenes - Ejemplo Selenium")

lbl_instruccion = tk.Label(ventana, text="Pega aquí los códigos (separados por espacios):")
lbl_instruccion.pack(padx=10, pady=5)

text_codigos = tk.Text(ventana, width=60, height=10)
text_codigos.pack(padx=10, pady=5)

btn_iniciar = tk.Button(ventana, text="Iniciar proceso", command=iniciar_proceso, bg="lightblue")
btn_iniciar.pack(pady=10)

ventana.mainloop()
