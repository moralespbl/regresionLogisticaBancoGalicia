import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException, NoSuchWindowException
from datetime import datetime, timedelta
import locale
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import glob
import requests 
import shutil # Necesario para mover archivos
import re # Para expresiones regulares, útil para extraer fechas del texto

# === CONFIGURACIÓN DEL LOCALE ===
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8') # Para sistemas Unix/Linux/macOS
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES') # Para algunos sistemas Windows
    except locale.Error:
        print("Advertencia: No se pudo configurar el locale español. Esto podría causar problemas con los nombres de meses.")
# === FIN CONFIGURACIÓN DEL LOCALE ===

# === CONFIGURACIÓN GLOBAL ===
AÑO_A_DESCARGAR = 2009 # ¡CAMBIAR ESTE AÑO!  
CARPETA_BASE_AÑO = f"IAMC_Informes_{AÑO_A_DESCARGAR}" # Carpeta principal para todo el año
# La carpeta de descarga temporal de Chrome será la carpeta base del año
CARPETA_DESCARGA_CHROME = os.path.abspath(CARPETA_BASE_AÑO) 

BASE_IAMC_URL = "https://www.iamc.com.ar/informediario/"

# Diccionario para mapear número de mes a nombre del mes en español para las carpetas
NOMBRES_MESES = {
    1: "01_Enero", 2: "02_Febrero", 3: "03_Marzo", 4: "04_Abril", 
    5: "05_Mayo", 6: "06_Junio", 7: "07_Julio", 8: "08_Agosto",
    9: "09_Septiembre", 10: "10_Octubre", 11: "11_Noviembre", 12: "12_Diciembre"
}

# Listas para el informe final de descargas fallidas
informes_no_encontrados = [] # Fechas donde no se encontró el enlace o se indicó "sin informes"
errores_descarga = []        # Fechas donde hubo un error técnico en la descarga (requests, selenium)
informes_fecha_incorrecta = [] # Nuevas: Para informes que abren una fecha diferente

# Crear la carpeta principal del año si no existe
if not os.path.exists(CARPETA_BASE_AÑO):
    os.makedirs(CARPETA_BASE_AÑO)
    print(f"Carpeta base del año creada: {CARPETA_BASE_AÑO}")
else:
    print(f"Carpeta base del año ya existe: {CARPETA_BASE_AÑO}")

# Opciones del navegador (comunes para todas las sesiones)
options = Options()
options.add_argument("--start-maximized")
options.add_experimental_option("excludeSwitches", ["enable-popup-blocking", "enable-logging"])
options.add_argument("--disable-popup-blocking")
# Configurar las preferencias de descarga para la carpeta base del año
prefs = {
    "download.default_directory": CARPETA_DESCARGA_CHROME,
    "download.prompt_for_download": False, 
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": False # Dejamos en False si queremos que abra el visor interno
}
options.add_experimental_option("prefs", prefs)
print(f"Configuración de Chrome para descarga de PDF aplicada a: {CARPETA_DESCARGA_CHROME}")


# Definimos driver y original_calendar_window_handle como variables globales
driver = None
original_calendar_window_handle = None

def inicializar_driver_unica_vez():
    global driver, original_calendar_window_handle
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    print("Navegador Chrome iniciado.")
    driver.get(BASE_IAMC_URL)
    print(f"Página {BASE_IAMC_URL} cargada.")
    original_calendar_window_handle = driver.current_window_handle
    print(f"DEBUG: Handle de la ventana original del calendario guardado/actualizado: {original_calendar_window_handle}")

# La inicialización del driver solo se hace una vez al principio
inicializar_driver_unica_vez()


# Función para verificar el mes y año actual en el calendario
def mes_visible_actual():
    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-calendar"))
        )
        mes_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-month"))
        )
        year_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-year"))
        )

        titulo_mes_str = mes_element.text
        titulo_year_str = year_element.text
        formato_mes_str = titulo_mes_str.title()

        current_visible_date = datetime.strptime(f"01/{formato_mes_str} {titulo_year_str}", "%d/%B %Y")
        return current_visible_date
    except Exception as e:
        print(f"ERROR: No se pudo encontrar el mes y/o año visible del calendario: {e}.")
        return datetime(1,1,1) 


# Función para navegar al mes deseado en el calendario
def ir_a_mes(fecha_objetivo):
    global original_calendar_window_handle 
    
    # Asegurarse de estar en la ventana original del calendario
    if driver.current_window_handle != original_calendar_window_handle:
        print("ADVERTENCIA: No estamos en la ventana original del calendario. Intentando volver...")
        try:
            driver.switch_to.window(original_calendar_window_handle)
        except NoSuchWindowException:
            print("CRÍTICO: La ventana original del calendario no existe. El driver se ha perdido. Reinicializando.")
            inicializar_driver_unica_vez() # Reiniciar completamente si el handle se pierde
            ir_a_mes(fecha_objetivo) # Reintentar la navegación
            return 
        except Exception as e:
            print(f"ERROR: No se pudo volver a la ventana original del calendario: {e}. Reinicializando el driver.")
            inicializar_driver_unica_vez() # Reiniciar completamente si falla el switch
            ir_a_mes(fecha_objetivo) # Reintentar la navegación
            return
            
    # Después de asegurar el handle, verificar la URL y el calendario
    if driver.current_url != BASE_IAMC_URL:
        print(f"ADVERTENCIA: La URL actual no es BASE_IAMC_URL. Recargando: {BASE_IAMC_URL}")
        driver.get(BASE_IAMC_URL)
            
    # Esperar que el calendario esté visible
    try:
        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-calendar"))
        )
        time.sleep(0.5) 
    except TimeoutException:
        print("ADVERTENCIA: Calendario no visible después de cambiar a ventana original o recargar. Recargando BASE_IAMC_URL como último recurso.")
        driver.get(BASE_IAMC_URL)
        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-calendar"))
        )
        time.sleep(0.5)

    while True:
        mes_actual = mes_visible_actual()
        if mes_actual.year == 1 and mes_actual.month == 1: 
            print("ADVERTENCIA: mes_visible_actual falló. Reintentando leer el mes actual.")
            time.sleep(2)
            if driver.current_url != BASE_IAMC_URL:
                print("DEBUG: mes_visible_actual falló y la URL no es BASE_IAMC_URL. Recargando.")
                driver.get(BASE_IAMC_URL)
                WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-calendar")))
                time.sleep(1)
            continue 

        if mes_actual.year == fecha_objetivo.year and mes_actual.month == fecha_objetivo.month:
            print(f"DEBUG: Mes actual ({mes_actual.strftime('%B %Y')}) coincide con objetivo ({fecha_objetivo.strftime('%B %Y')}).")
            break
        elif mes_actual < fecha_objetivo:
            print(f"Navegando: {mes_actual.strftime('%B %Y')} < {fecha_objetivo.strftime('%B %Y')}. Clic en 'siguiente'.")
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "ui-datepicker-next"))
            ).click()
        else: # mes_actual > fecha_objetivo
            print(f"Navegando: {mes_actual.strftime('%B %Y')} > {fecha_objetivo.strftime('%B %Y')}. Clic en 'anterior'.")
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "ui-datepicker-prev"))
            ).click()
        time.sleep(1)

# Función para esperar que un archivo de descarga aparezca y se complete
def esperar_y_mover_descarga(expected_file_name_prefix, fecha, carpeta_destino_mes, timeout=60):
    start_time = time.time()
    downloaded_file = None
    
    while time.time() - start_time < timeout:
        possible_files = glob.glob(os.path.join(CARPETA_DESCARGA_CHROME, f"{expected_file_name_prefix}*.pdf"))
        possible_crdownloads = glob.glob(os.path.join(CARPETA_DESCARGA_CHROME, f"{expected_file_name_prefix}*.crdownload"))

        if possible_files:
            downloaded_file = possible_files[0]
            if not downloaded_file.endswith(".crdownload"): 
                print(f"DEBUG: Descarga de '{downloaded_file}' detectada por Chrome (no requests).")
                break
        elif possible_crdownloads:
            print(f"DEBUG: Descarga en progreso para '{expected_file_name_prefix}' (.crdownload detectado por Chrome). Esperando...")
        
        time.sleep(1) 
    
    if downloaded_file and os.path.exists(downloaded_file) and not downloaded_file.endswith(".crdownload"):
        target_path = os.path.join(carpeta_destino_mes, os.path.basename(downloaded_file))
        try:
            shutil.move(downloaded_file, target_path)
            print(f"ÉXITO: PDF movido de '{os.path.basename(downloaded_file)}' a '{carpeta_destino_mes}'.")
            return True
        except Exception as e:
            print(f"ERROR: No se pudo mover el archivo '{downloaded_file}' a '{carpeta_destino_mes}'. Excepción: {e}")
            errores_descarga.append(f"{fecha.strftime('%d-%m-%Y')} (Error al mover archivo descargado por Chrome: {e})")
            return False
    else:
        print(f"ADVERTENCIA: No se encontró la descarga de '{expected_file_name_prefix}' en '{CARPETA_DESCARGA_CHROME}' después de {timeout} segundos (descarga por Chrome).")
        errores_descarga.append(f"{fecha.strftime('%d-%m-%Y')} (No se completó la descarga del navegador)")
        return False


# Función para parsear una fecha del texto del informe
def parsear_fecha_de_texto_informe(texto_informe):
    # Intentar con formato DD/MM/YYYY o DD/MM/YY
    match = re.search(r'\d{2}/\d{2}/(\d{4}|\d{2})', texto_informe)
    if match:
        fecha_str = match.group(0)
        try:
            return datetime.strptime(fecha_str, '%d/%m/%Y')
        except ValueError:
            try:
                return datetime.strptime(fecha_str, '%d/%m/%y')
            except ValueError:
                return None
    return None


# Función principal para intentar descargar el informe de una fecha
def intentar_descargar(fecha, carpeta_destino_mes): 
    global original_calendar_window_handle, informes_no_encontrados, errores_descarga, informes_fecha_incorrecta
    
    print(f"\n--- Procesando fecha: {fecha.strftime('%d-%m-%Y')} ---")

    ir_a_mes(fecha) 

    expected_file_name_prefix = f"InformeRentaFija_{fecha.strftime('%Y%m%d')}" 
    target_pdf_name = f"{expected_file_name_prefix}.pdf"
    target_pdf_path = os.path.join(carpeta_destino_mes, target_pdf_name) 
    
    if os.path.exists(target_pdf_path):
        print(f"INFO: El informe para {fecha.strftime('%d-%m-%Y')} ya existe en {target_pdf_path}. Saltando descarga.")
        return 

    current_calendar_window_handle = driver.current_window_handle
    print(f"DEBUG: Handle de la ventana de calendario actual antes de clic en día: {current_calendar_window_handle}")

    new_pdf_window_handle = None 
    
    try:
        day_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//a[contains(@class, 'ui-state-default') and text()='{fecha.day}']"))
        )
        print(f"DEBUG: Encontrado elemento del día {fecha.day}.")

        day_element.click()
        print(f"DEBUG: Clic en el día {fecha.day} ejecutado.")
        time.sleep(2) # Pausa adicional después del clic en el día

        try:
            sin_informes_span = WebDriverWait(driver, 2).until(
                EC.visibility_of_element_located((By.XPATH, "//span[@class='sinInformes' and contains(text(), 'No se encontraron Informes')]"))
            )
            print(f"INFO: No se encontraron informes para la fecha {fecha.strftime('%d-%m-%Y')}. Saltando.")
            informes_no_encontrados.append(f"{fecha.strftime('%d-%m-%Y')} (No se encontraron informes en la página)")
            return 
        except TimeoutException:
            pass 

        # MODIFICACIÓN CLAVE: Buscamos cualquier enlace de informe que aparezca
        # Se elimina el filtro estricto por fecha en el XPath inicial para atrapar *cualquier* informe que el sitio muestre.
        xpath_any_informe_link = "//a[./div[@class='descripcion' and (contains(normalize-space(.), 'Diario') or contains(normalize-space(.), 'Informe Renta Fija'))]]"
        
        print(f"DEBUG: Buscando CUALQUIER enlace de informe desplegado con XPath: {xpath_any_informe_link}")

        informe_link_element = WebDriverWait(driver, 20).until( # Aumentado a 20 segundos
            EC.visibility_of_element_located((By.XPATH, xpath_any_informe_link))
        )
        
        # Ahora, una vez que encontramos un elemento, extraemos su texto
        texto_link = informe_link_element.find_element(By.CLASS_NAME, "descripcion").text
        print(f"DEBUG: Texto del enlace del informe encontrado: '{texto_link}'")

        # Intentamos parsear la fecha del texto encontrado
        fecha_en_link = parsear_fecha_de_texto_informe(texto_link)

        if not fecha_en_link or fecha_en_link.date() != fecha.date():
            print(f"ADVERTENCIA: Para {fecha.strftime('%d-%m-%Y')}: Se encontró el informe '{texto_link}', pero la fecha ({fecha_en_link.strftime('%d-%m-%Y') if fecha_en_link else 'N/A'}) no coincide con la fecha objetivo. Saltando descarga para esta fecha.")
            informes_fecha_incorrecta.append(f"{fecha.strftime('%d-%m-%Y')} (Encontrado: '{texto_link}' - Fecha real: {fecha_en_link.strftime('%d-%m-%Y') if fecha_en_link else 'N/A'})")
            return # No descargamos si la fecha no coincide.

        # Si la fecha coincide, procedemos con la URL
        href_pagina_informe_relativa = informe_link_element.get_attribute("href")

        if not href_pagina_informe_relativa.startswith("http"):
            base_domain_url = BASE_IAMC_URL.split('/informediario')[0]
            full_href_pagina_informe = f"{base_domain_url}{href_pagina_informe_relativa}"
        else:
            full_href_pagina_informe = href_pagina_informe_relativa

        print(f"INFO: URL obtenida del enlace del informe: {full_href_pagina_informe}")
        print(f"DEBUG: URL en minúsculas para verificación de tipo: {full_href_pagina_informe.lower()}")

        # === LÓGICA PRINCIPAL DE DESCARGA: Priorizar requests ===
        pdf_url_to_attempt_download = None 
        
        if full_href_pagina_informe.lower().endswith(".pdf") or \
           "/handlers/basestreamer.ashx" in full_href_pagina_informe.lower() or \
           "/handlers/basehandler.ashx" in full_href_pagina_informe.lower() or \
           ("/informediario/" in full_href_pagina_informe.lower() and \
            f"/{fecha.year}/" in full_href_pagina_informe.lower() and \
            f"/{fecha.month:02d}/" in full_href_pagina_informe.lower() and \
            f"/{fecha.day:02d}" in full_href_pagina_informe.lower() and \
            not ("/contenido/detail/" in full_href_pagina_informe.lower() or "basestreamer.ashx" in full_href_pagina_informe.lower()) \
           ):
            pdf_url_to_attempt_download = full_href_pagina_informe
            print(f"INFO: La URL inicial identificada como enlace directo a PDF o streamer. Se intentará descarga con requests.")
        
        else: 
            print(f"INFO: La URL identificada como página intermedia. Abriendo en una nueva pestaña para buscar el PDF real.")
            driver.execute_script("window.open(arguments[0], '_blank');", full_href_pagina_informe)
            print(f"DEBUG: Abierta nueva pestaña para: {full_href_pagina_informe}")
            time.sleep(1) # Dar tiempo para que la pestaña se abra

            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2)) 
            
            for window_handle in driver.window_handles:
                if window_handle != current_calendar_window_handle:
                    driver.switch_to.window(window_handle)
                    new_pdf_window_handle = window_handle 
                    break
            
            if not new_pdf_window_handle:
                print(f"ERROR: No se encontró una nueva pestaña para el PDF después de abrirla para {fecha.strftime('%d-%m-%Y')}.")
                errores_descarga.append(f"{fecha.strftime('%d-%m-%Y')} (No se pudo cambiar a la nueva pestaña abierta)")
                return 
            
            print(f"DEBUG: Cambiado a la nueva pestaña con handle: {new_pdf_window_handle}. URL actual: {driver.current_url}")
            time.sleep(1.5) 

            pdf_url_from_intermediate_page = None
            try:
                pdf_object_element = WebDriverWait(driver, 5).until( 
                    EC.presence_of_element_located((By.XPATH, "//object[@type='application/pdf']"))
                )
                pdf_url_from_intermediate_page = pdf_object_element.get_attribute("data")
                print("DEBUG: <object> encontrado por tipo 'application/pdf' en página intermedia. URL extraída: " + pdf_url_from_intermediate_page)
            except TimeoutException:
                print("DEBUG: No se encontró <object type='application/pdf'> en página intermedia. Intentando buscar enlace de descarga.")
                try:
                    pdf_download_link_element = WebDriverWait(driver, 5).until( 
                        EC.presence_of_element_located((By.XPATH, "//a[contains(@class, 'pdfDownload') and contains(@href, 'BaseStreamer.ashx')]"))
                    )
                    pdf_url_from_intermediate_page = pdf_download_link_element.get_attribute("href")
                    print("DEBUG: Enlace con class='pdfDownload' y BaseStreamer.ashx encontrado en página intermedia. URL extraída: " + pdf_url_from_intermediate_page)
                except TimeoutException:
                    print("ADVERTENCIA: No se encontró un enlace de descarga de PDF ni un <object> en la página intermedia para la fecha " + fecha.strftime('%d-%m-%Y') + ". El formato puede haber cambiado.")
                    informes_no_encontrados.append(f"{fecha.strftime('%d-%m-%Y')} (No se encontró <object> ni enlace de descarga en página intermedia)")
            
            if pdf_url_from_intermediate_page:
                pdf_url_to_attempt_download = pdf_url_from_intermediate_page 

        if pdf_url_to_attempt_download:
            try:
                print(f"DEBUG: Iniciando descarga FINAL del PDF desde {pdf_url_to_attempt_download} usando requests.")
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}
                response = requests.get(pdf_url_to_attempt_download, stream=True, headers=headers, timeout=30, verify=False) 
                response.raise_for_status() 

                with open(target_pdf_path, 'wb') as pdf_file:
                    for chunk in response.iter_content(chunk_size=8192):
                        pdf_file.write(chunk)
                print(f"ÉXITO: PDF descargado y guardado como: {target_pdf_name}")
            except requests.exceptions.RequestException as e:
                print(f"ERROR: Falló la descarga del PDF usando requests desde la URL: {pdf_url_to_attempt_download} para {fecha.strftime('%d-%m-%Y')}. Excepción: {e}")
                errores_descarga.append(f"{fecha.strftime('%d-%m-%Y')} (Error de descarga con requests desde URL final: {e})")
            except Exception as e:
                print(f"ERROR: Un error inesperado ocurrió al procesar la descarga del PDF (requests desde URL final) para {fecha.strftime('%d-%m-%Y')}. Excepción: {type(e).__name__}: {e}")
                errores_descarga.append(f"{fecha.strftime('%d-%m-%Y')} (Error inesperado en descarga con requests desde URL final: {type(e).__name__} - {e})")
        else:
            print(f"ADVERTENCIA: No se pudo obtener una URL de PDF válida para descargar para {fecha.strftime('%d-%m-%Y')}.")
            informes_no_encontrados.append(f"{fecha.strftime('%d-%m-%Y')} (No se pudo extraer URL de PDF para descarga)")

    except (NoSuchElementException, TimeoutException) as e:
        print(f"ADVERTENCIA: Para {fecha.strftime('%d-%m-%Y')}: Elemento no encontrado o tiempo de espera excedido durante la navegación inicial (calendario/enlace). Excepción: {type(e).__name__}: {e}. Probablemente no hay informe o la estructura de la página cambió.")
        informes_no_encontrados.append(f"{fecha.strftime('%d-%m-%Y')} (Elemento no encontrado/Timeout en navegación inicial: {type(e).__name__})")
    except WebDriverException as e:
        print(f"ERROR: Para {fecha.strftime('%d-%m-%Y')}: Error de WebDriver inesperado - {type(e).__name__}: {e}. Reinicializando el driver para intentar recuperarse.")
        errores_descarga.append(f"{fecha.strftime('%d-%m-%Y')} (WebDriverException: {type(e).__name__} - {e})")
        try:
            driver.quit()
        except:
            pass 
        inicializar_driver_unica_vez() 
    except Exception as e:
        print(f"ERROR: Para {fecha.strftime('%d-%m-%Y')}: Error inesperado en el flujo general - {type(e).__name__}: {e}. Reinicializando el driver para intentar recuperarse.")
        errores_descarga.append(f"{fecha.strftime('%d-%m-%Y')} (Error inesperado en flujo general: {type(e).__name__} - {e})")
        try:
            driver.quit()
        except:
            pass 
        inicializar_driver_unica_vez() 
    finally:
        all_window_handles = driver.window_handles
        if new_pdf_window_handle and new_pdf_window_handle in all_window_handles and driver.current_window_handle == new_pdf_window_handle:
            try:
                driver.close() 
                print("DEBUG: Cerrada la pestaña secundaria del informe/PDF.")
            except Exception as e:
                print(f"ADVERTENCIA: Error al cerrar la pestaña secundaria: {e}. Puede que ya se haya cerrado o el handle sea inválido.")
            finally:
                if original_calendar_window_handle in driver.window_handles:
                    try:
                        driver.switch_to.window(original_calendar_window_handle)
                        print("DEBUG: Vuelto a la ventana original del calendario.")
                    except Exception as e:
                        print(f"ADVERTENCIA: Error al volver a la ventana original después de cerrar secundaria: {e}. Puede que el driver esté en un estado inconsistente.")
                else:
                    print("CRÍTICO: La ventana original del calendario no existe después de cerrar una secundaria. Es posible que el driver esté corrupto.")
        elif driver.current_window_handle == original_calendar_window_handle:
            print("DEBUG: Ya estamos en la ventana original del calendario. No se requiere cierre/cambio de pestaña.")
        elif original_calendar_window_handle in driver.window_handles:
            driver.switch_to.window(original_calendar_window_handle)
            print("DEBUG: Vuelto a la ventana original del calendario (estábamos en una pestaña inesperada).")
        else:
            print("ADVERTENCIA: Estado de ventanas inesperado y la ventana principal no está disponible. El driver podría necesitar reinicio.")


# === Bucle principal para el año completo ===
print(f"\n--- Iniciando descarga de informes para el año {AÑO_A_DESCARGAR} ---")

# Para el año completo, se puede usar range(1, 13). 
# Para empezar en un mes específico (ej. Agosto=8), range(8, 13).
for mes_num in range(1, 13): # Iterar de Enero (1) a Diciembre (12)
    nombre_mes = NOMBRES_MESES[mes_num]
    carpeta_destino_mes = os.path.join(CARPETA_BASE_AÑO, nombre_mes)

    if not os.path.exists(carpeta_destino_mes):
        os.makedirs(carpeta_destino_mes)
        print(f"\nCarpeta de destino creada para {nombre_mes}: {carpeta_destino_mes}")
    else:
        print(f"DEBUG: Carpeta de destino ya existe para {nombre_mes}: {carpeta_destino_mes}")

    inicio_mes = datetime(AÑO_A_DESCARGAR, mes_num, 1)
    if mes_num == 12:
        fin_mes = datetime(AÑO_A_DESCARGAR + 1, 1, 1) - timedelta(days=1)
    else:
        fin_mes = datetime(AÑO_A_DESCARGAR, mes_num + 1, 1) - timedelta(days=1)

    fecha_actual = inicio_mes
    while fecha_actual <= fin_mes:
        # Comprobar si el día es fin de semana (sábado=5, domingo=6)
        if fecha_actual.weekday() < 5: 
            intentar_descargar(fecha_actual, carpeta_destino_mes) # Pasar la carpeta del mes
        else:
            print(f"{fecha_actual.strftime('%d-%m-%Y')}: Fin de semana. Saltando.")
        fecha_actual += timedelta(days=1)

print(f"\n--- Proceso de descarga completado para el año {AÑO_A_DESCARGAR} ---")

# === Generación del informe final ===
reporte_path = os.path.join(CARPETA_BASE_AÑO, f"reporte_descargas_{AÑO_A_DESCARGAR}.txt")
with open(reporte_path, 'w', encoding='utf-8') as f:
    f.write(f"--- Reporte de Descargas para el Año {AÑO_A_DESCARGAR} ---\n")
    f.write(f"Fecha de generación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

    if informes_no_encontrados:
        f.write("--- Días donde NO se encontró el informe (No hay enlace o 'Sin informes') ---\n")
        for fecha_info in informes_no_encontrados:
            f.write(f"- {fecha_info}\n")
        f.write("\n")
    else:
        f.write("--- Todos los informes se encontraron (sin 'sin informes') ---\n\n")

    if informes_fecha_incorrecta: # Nuevo informe para fechas incorrectas
        f.write("--- Días donde el informe encontrado CORRESPONDE a una FECHA DISTINTA a la buscada ---\n")
        f.write("(Esto indica un comportamiento inusual del calendario o redirección del sitio)\n")
        for fecha_info in informes_fecha_incorrecta:
            f.write(f"- {fecha_info}\n")
        f.write("\n")
    else:
        f.write("--- No se registraron informes con fecha incorrecta ---\n\n")

    if errores_descarga:
        f.write("--- Días con ERRORES TÉCNICOS durante la descarga ---\n")
        f.write("(Esto puede incluir problemas de red, errores de Selenium, o URL no válidas)\n")
        for error_info in errores_descarga:
            f.write(f"- {error_info}\n")
        f.write("\n")
    else:
        f.write("--- No se registraron errores técnicos de descarga ---\n\n")

    if not informes_no_encontrados and not errores_descarga and not informes_fecha_incorrecta:
        f.write("¡Felicitaciones! Todos los informes se procesaron correctamente sin errores o informes faltantes conocidos.\n")

print(f"\n--- Informe final de descargas guardado en: {reporte_path} ---")

if driver:
    driver.quit()
    print("Navegador Chrome cerrado.")