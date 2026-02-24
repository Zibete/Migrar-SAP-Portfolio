from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

from datetime import datetime

import keyring
import time
import os
import glob

entorno = os.getenv("MIGRAR_SAP_RETAILWEB_KEYRING", "<REDACTED>")
linkSB = os.getenv("MIGRAR_SAP_RETAILWEB_URL", "<REDACTED_URL>")

def esperar_descarga(carpeta, patron="*.xlsx", timeout=180, intervalo=1.0, chequeos_estables=2):
    """
    Espera hasta que haya un archivo que:
      - cumpla el patron,
      - sea de hoy,
      - no tenga .crdownload,
      - y mantenga tamano estable 'chequeos_estables' veces seguidas.
    Devuelve la ruta del archivo o None si expira el timeout.
    """
    fin = time.time() + timeout
    candidato = None
    size_ok_repetido = 0
    size_prev = -1

    while time.time() < fin:
        # Ignora descargas temporales.
        crdowns = set(os.path.splitext(p)[0] for p in glob.glob(os.path.join(carpeta, "*.crdownload")))

        # Selecciona el .xlsx mas nuevo de hoy sin archivo temporal asociado.
        hoy = datetime.now().date()
        xls = sorted(glob.glob(os.path.join(carpeta, patron)), key=os.path.getmtime, reverse=True)

        candidato = next(
            (p for p in xls
             if os.path.splitext(p)[0] not in crdowns
             and datetime.fromtimestamp(os.path.getmtime(p)).date() == hoy),
            None
        )

        if candidato and os.path.exists(candidato):
            size_act = os.path.getsize(candidato)
            if size_act == size_prev and size_act > 0:
                size_ok_repetido += 1
            else:
                size_ok_repetido = 0
                size_prev = size_act

            # Si el tamano se mantiene estable, se considera descarga completa.
            if size_ok_repetido >= chequeos_estables:
                return candidato

        time.sleep(intervalo)

    return None

def click_seguro(xpath, timeout=60):
    overlays = [
        (By.XPATH, "//div[contains(@class, 'waiting-panel')]"),
        (By.ID, "col")
    ]
    for locator in overlays:
        try:
            WebDriverWait(driver, 5).until(EC.invisibility_of_element_located(locator))
        except:
            pass

    elemento = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    try:
        elemento.click()
    except:
        driver.execute_script("arguments[0].click();", elemento)

try:

    start_time = time.time()

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")  # Ejecuta sin interfaz
    chrome_options.add_argument("--disable-gpu")  # Evita problemas en headless
    chrome_options.add_argument("--window-size=1920,1080")  # Simula pantalla completa
    chrome_options.add_argument("--disable-extensions")  # Desactiva extensiones
    chrome_options.add_argument("--disable-notifications")  # Bloquea notificaciones
    chrome_options.add_argument("--disable-infobars")  # Oculta mensajes de "Chrome esta controlado..."
    chrome_options.add_argument("--mute-audio")  # Silencia sonidos de la pagina

    driver = webdriver.Chrome(options=chrome_options)

    driver.get(linkSB)

    credential = keyring.get_credential(entorno, None)
    if not credential:
        raise RuntimeError("No se encontraron credenciales en keyring.")

    username_field = driver.find_element(By.NAME, "dgf_login_form_fd-username")
    username_field.send_keys(credential.username)

    password_field = driver.find_element(By.NAME, "dgf_login_form_fd-password_encrypted")
    password_field.send_keys(credential.password)

    WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.ID, "form.login.title"))
    ).click()

    idVentanaHome = driver.current_window_handle  # Asignar el valor del handle de la ventana actual

    # Abre la ventana de recepciones.

    driver.switch_to.window(idVentanaHome)

    num_ventanas_anterior = len(driver.window_handles)

    click_seguro("//a[.//span[contains(@class,'pull-left') and normalize-space()='Control de Inventarios']]")
    click_seguro("//a[.//span[contains(@class,'pull-left') and normalize-space()='Control de Recepciones Pagadas']]")

    WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) == num_ventanas_anterior + 1)

    idVentanaRecepciones = driver.window_handles[-1]

    # Abre la ventana de reportes.

    driver.switch_to.window(idVentanaHome)

    time.sleep(1)

    num_ventanas_anterior = len(driver.window_handles)

    click_seguro("//a[.//span[contains(@class,'pull-left') and normalize-space()='Sistema']]")
    click_seguro("//a[.//span[contains(@class,'pull-left') and normalize-space()='Reportes para Descargar']]")

    WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) == num_ventanas_anterior + 1)

    idVentanaReportes = driver.window_handles[-1]

    # Configura el reporte en RetailWeb.

    driver.switch_to.window(idVentanaRecepciones)

    time.sleep(1)

    # Filtra por pago pendiente = Si.
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(@class, 'btn') and contains(text(), 'S') and .//input[contains(@name, 'pendingPay')]]"))
    ).click()

    # Filtra por anulado = No.
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(@class, 'btn') and contains(text(), 'N') and .//input[contains(@name, 'reversed')]]"))
    ).click()

    click_seguro("//a[contains(@class, 'btn') and contains(text(), 'Buscar')]")

    time.sleep(1)

    # Espera a que desaparezca el panel de carga.
    WebDriverWait(driver, 60).until(
        lambda d: not d.find_elements(By.XPATH, "//div[@id='waitpane']//div[contains(@class, 'panel waiting-panel')]")
    )

    click_seguro("//button[contains(@class, 'btn btn-default btn-sm dropdown-toggle ')]")
    click_seguro("//a[contains(@id, 'excel')]")

    # Espera a que desaparezca el panel de carga.
    WebDriverWait(driver, 60).until(
        lambda d: not d.find_elements(By.XPATH, "//div[@id='waitpane']//div[contains(@class, 'panel waiting-panel')]")
    )

    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@id, 'action-volver') and contains(text(), 'Volver')]"))).click()

    # Descarga el reporte generado desde la ventana de reportes.

    driver.switch_to.window(idVentanaReportes)

    driver.refresh()
    time.sleep(1)

    while True:
        tabla = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//tbody[contains(@id, 'scroll_body')]")))
        fila = tabla.find_elements(By.TAG_NAME, "tr")[0]
        fechadisponible = fila.find_elements(By.TAG_NAME, "td")[6]
        fechaarchivo = fila.find_elements(By.TAG_NAME, "td")[5]

        if fechadisponible is not None and fechadisponible.text.strip() != "":
            break
        else:
            driver.refresh()
            time.sleep(1)

    fila.click()

    click_seguro("//a[contains(@id, 'DownloadResults-download')]")

    carpeta_descargas = os.path.expanduser("~/Downloads")
    archivo = esperar_descarga(carpeta_descargas, patron="*{fecha}*.xlsx".format(
        fecha=datetime.now().strftime("%Y%m%d")
    ))

    if not archivo:
        raise FileNotFoundError("No se completo la descarga del .xlsx dentro del tiempo esperado.")

    print(archivo, end="", flush=True)

except Exception as e:
    import traceback
    elapsed_time = time.time() - start_time
    minutes = int(elapsed_time // 60)
    seconds = int(elapsed_time % 60)

    mensaje = f"ERROR a {minutes} minutos y {seconds} segundos de iniciar.\nError: {e}\n\n"
    mensaje += traceback.format_exc()
    print(mensaje)

finally:

    if 'driver' in locals():
        driver.quit()
