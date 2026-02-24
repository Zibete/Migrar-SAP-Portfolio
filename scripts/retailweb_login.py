from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

from datetime import datetime, timedelta

import keyring
import os
import shutil
import time

email = os.getenv("MIGRAR_SAP_RETAILWEB_EMAIL", "<REDACTED>")
entorno = os.getenv("MIGRAR_SAP_RETAILWEB_KEYRING", "<REDACTED>")
linkSB = os.getenv("MIGRAR_SAP_RETAILWEB_URL", "<REDACTED_URL>")
destinoIndicadores = os.getenv(
    "MIGRAR_SAP_RETAILWEB_DESTINO",
    "<REDACTED_PATH>",
)

mensajeSubject = ""
mensajeBody = ""
ruta_mes = ""

idVentanaLogin = None
idVentanaRecepciones = None
idVentanaReportes = None

MONTHS_ES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]

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

    # Inicializa el WebDriver con las opciones definidas.
    driver = webdriver.Chrome(options=chrome_options)

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

    class RetailWebSession:

        def __init__(self):  # Constructor recibe el driver
            self.fechaarchivo = None
            self.Rep_hasta_dia = None
            self.Rep_mes = None
            self.Rep_anio = None
            self.nombreNuevo = None

        def loginRetailWeb(self):

            driver.get(linkSB)

            credential = keyring.get_credential(entorno, None)
            if not credential:
                raise RuntimeError("No se encontraron credenciales en keyring.")

            username_field = driver.find_element(By.NAME, "dgf_login_form_fd-username")
            username_field.send_keys(credential.username)

            password_field = driver.find_element(By.NAME, "dgf_login_form_fd-password_encrypted")
            password_field.send_keys(credential.password)

            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "form.login.title"))
            ).click()

            global idVentanaLogin
            idVentanaLogin = driver.current_window_handle  # Asignar el valor del handle de la ventana actual

        def abrirVentanaRecepciones(self):

            driver.switch_to.window(idVentanaLogin)

            num_ventanas_anterior = len(driver.window_handles)

            click_seguro("//span[contains(text(), 'Control de Inventarios')]")
            click_seguro("//span[contains(text(), 'Control de Recepciones Pagadas')]")

            WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) == num_ventanas_anterior + 1)

            global idVentanaRecepciones
            idVentanaRecepciones = driver.window_handles[-1]

        def abrirVentanaReportes(self):

            driver.switch_to.window(idVentanaLogin)

            num_ventanas_anterior = len(driver.window_handles)

            click_seguro("//span[contains(@class, 'pull-left') and contains(text(), 'Sistema')]")
            click_seguro("//span[contains(@class, 'pull-left') and contains(text(), 'Reportes para Descargar')]")

            WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) == num_ventanas_anterior + 1)

            global idVentanaReportes
            idVentanaReportes = driver.window_handles[-1]

        def asignarCarpeta(self):

            global mensajeBody

            # Construye la ruta del anio.
            ruta_anio = os.path.join(destinoIndicadores, str(self.Rep_anio))

            # Crea la carpeta del anio si no existe.
            if not os.path.exists(ruta_anio):
                os.makedirs(ruta_anio)
                mensaje = f"Carpeta {self.Rep_anio} creada en: {destinoIndicadores}"
                print(mensaje)
                mensajeBody += mensaje
            else:
                print(f"La carpeta {self.Rep_anio} ya existe en: {destinoIndicadores}")

            nombre_mes_espanol = MONTHS_ES[self.Rep_mes - 1].capitalize()

            global ruta_mes
            ruta_mes = os.path.join(ruta_anio, f"{self.Rep_mes:02d}-{nombre_mes_espanol}")

            # Crea la carpeta del mes si no existe.
            if not os.path.exists(ruta_mes):
                os.makedirs(ruta_mes)
                mensaje = f"Carpeta {self.Rep_mes:02d}-{nombre_mes_espanol} creada en: {ruta_anio}."
                print(mensaje)
                mensajeBody += f"{mensaje}\n"
            else:
                print(f"La carpeta {self.Rep_mes:02d}-{nombre_mes_espanol} ya existe en: {ruta_anio}")

            elementos = os.listdir(ruta_mes)

            archivos_excel = [f for f in elementos if os.path.isfile(os.path.join(ruta_mes, f)) and f.lower().endswith(('.xlsx'))]

            cantidad_elementos = len(archivos_excel)

            nombreNuevo = f"01.{self.Rep_mes:02}.{self.Rep_anio} al {self.Rep_hasta_dia:02}.{self.Rep_mes:02}.{self.Rep_anio}"

            if self.Rep_mes == datetime.now().month and any(
                f.endswith(f"{nombreNuevo}.xlsx") for f in elementos if os.path.isfile(os.path.join(ruta_mes, f))
            ):
                mensaje = f"No se realizo la descarga: El reporte: {nombreNuevo} ya existe en {ruta_mes}"
                print(mensaje)
                mensajeBody += f"{mensaje}\n"
            else:

                excel_creado_hoy = any(
                    datetime.fromtimestamp(os.path.getctime(os.path.join(ruta_mes, f))).date() == datetime.now().date()
                    for f in os.listdir(ruta_mes)
                    if os.path.isfile(os.path.join(ruta_mes, f)) and f.endswith(f"{nombreNuevo}.xlsx")
                )

                if excel_creado_hoy:
                    mensaje = f"No se realizo la descarga: El reporte {nombreNuevo} ya existe en {ruta_mes} y fue creado hoy."
                    print(mensaje)
                    mensajeBody += f"{mensaje}\n"
                else:
                    self.nombreNuevo = f"{cantidad_elementos+1:02d}. {nombreNuevo}"

        def generarReporteRetailWeb(self):

            driver.switch_to.window(idVentanaRecepciones)

            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//label[contains(@class, 'btn') and contains(text(), 'Todos')]"))).click()

            formatted_Rep = f"01/{self.Rep_mes:02d}/{self.Rep_anio} 00:00 - {self.Rep_hasta_dia:02d}/{self.Rep_mes:02d}/{self.Rep_anio} 23:59"

            campo = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.XPATH, "//input[contains(@class, 'input') and contains(@name, 'businessDate')]"))
            )
            campo.clear()
            campo.send_keys(formatted_Rep)

            click_seguro(f"//div[.//span[@class='drp-selected' and contains(text(), '{formatted_Rep}')]]//button[contains(@class, 'btn-primary')]")
            click_seguro("//a[contains(@class, 'btn') and contains(text(), 'Buscar')]")

            WebDriverWait(driver, 60).until_not(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@id='waitpane']//div[contains(@class, 'panel waiting-panel')]"))
            )

            click_seguro("//button[contains(@class, 'btn btn-default btn-sm dropdown-toggle ')]")
            click_seguro("//a[contains(@id, 'excel')]")

            WebDriverWait(driver, 60).until_not(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@id='waitpane']//div[contains(@class, 'panel waiting-panel')]"))
            )

            click_seguro("//a[contains(@id, 'action-volver') and contains(text(), 'Volver')]")

        def descargarReporteRetailWeb(self):

            driver.switch_to.window(idVentanaReportes)

            driver.refresh()
            time.sleep(1)

            while True:
                tabla = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//tbody[contains(@id, 'scroll_body')]")))
                fila = tabla.find_elements(By.TAG_NAME, "tr")[0]
                fechadisponible = fila.find_elements(By.TAG_NAME, "td")[6]
                self.fechaarchivo = fila.find_elements(By.TAG_NAME, "td")[5]

                if fechadisponible is not None and fechadisponible.text.strip() != "":
                    print("Fecha disponible encontrada:", fechadisponible.text.strip())
                    break
                else:
                    print("Fecha disponible vacia, refrescando...")
                    driver.refresh()
                    time.sleep(1)

            fila.click()
            click_seguro("//a[contains(@id, 'DownloadResults-download')]")

            time.sleep(2)

        def moverArchivo(self):
            from pathlib import Path
            import time
            import glob
            from datetime import datetime

            download_dir = Path.home() / "Downloads"
            fecha_hoy_str = datetime.now().strftime("%Y%m%d")

            MAX_WAIT_SECONDS = 300
            POLL_INTERVAL = 1

            archivo: Path | None = None
            t0 = time.time()

            while time.time() - t0 < MAX_WAIT_SECONDS:
                # Busca archivos .xls/.xlsx que incluyan la fecha de hoy.
                candidatos = []
                for patron in [f"*{fecha_hoy_str}*.xls", f"*{fecha_hoy_str}*.xlsx"]:
                    candidatos.extend([Path(p) for p in glob.glob(str(download_dir / patron))])

                hoy = datetime.now().date()
                candidatos_hoy = [p for p in candidatos if datetime.fromtimestamp(p.stat().st_mtime).date() == hoy]

                if candidatos_hoy:
                    candidato = max(candidatos_hoy, key=lambda p: p.stat().st_mtime)

                    # Verifica que no siga descargando (.crdownload).
                    base = candidato.name
                    cr_pending = (
                        list(download_dir.glob(f"{base}.crdownload")) or
                        list(download_dir.glob(f"*{fecha_hoy_str}*.crdownload"))
                    )
                    if not cr_pending:
                        archivo = candidato
                        break

                time.sleep(POLL_INTERVAL)

            if not archivo:
                raise FileNotFoundError(
                    "No se encontro un archivo Excel descargado hoy cuyo nombre contenga la fecha "
                    f"'{fecha_hoy_str}' en {download_dir}."
                )

            print(f"[INFO] Archivo encontrado: {archivo}")

            # Mueve al destino manteniendo el nombre original.
            ruta_destino = Path(ruta_mes)
            ruta_destino.mkdir(parents=True, exist_ok=True)

            movido = ruta_destino / archivo.name
            shutil.move(str(archivo), str(movido))

            # Renombra con el nombre final y preserva la extension original.
            nombre_final = ruta_destino / f"{self.nombreNuevo}{movido.suffix}"
            movido.rename(nombre_final)

            mensaje = f"Se creo el siguiente archivo: {nombre_final}."
            print(mensaje)
            global mensajeBody
            mensajeBody += f"{mensaje}\n"

    # Inicializa la sesion de RetailWeb.
    retailweb = RetailWebSession()

    # Reporte del periodo actual.
    retailweb.Rep_hasta_dia = (datetime.now() - timedelta(days=1)).day
    retailweb.Rep_mes = (datetime.now() - timedelta(days=1)).month
    retailweb.Rep_anio = (datetime.now() - timedelta(days=1)).year

    retailweb.loginRetailWeb()
    retailweb.abrirVentanaRecepciones()
    retailweb.abrirVentanaReportes()

    retailweb.asignarCarpeta()

    if retailweb.nombreNuevo is not None:
        retailweb.generarReporteRetailWeb()
        retailweb.descargarReporteRetailWeb()
        retailweb.moverArchivo()

    # Reporte complementario del mes anterior (solo primera quincena).
    if datetime.now().day < 15:

        retailweb = RetailWebSession()

        # Calcula el primer dia del mes actual para derivar el ultimo dia del mes anterior.
        first_day_of_current_month = datetime(datetime.now().year, datetime.now().month, 1)
        retailweb.Rep_hasta_dia = (first_day_of_current_month - timedelta(days=1)).day
        retailweb.Rep_mes = (first_day_of_current_month - timedelta(days=1)).month
        retailweb.Rep_anio = (first_day_of_current_month - timedelta(days=1)).year

        retailweb.asignarCarpeta()

        if retailweb.nombreNuevo is not None:
            retailweb.generarReporteRetailWeb()
            retailweb.descargarReporteRetailWeb()
            retailweb.moverArchivo()

    # Calcula el tiempo total de ejecucion.
    elapsed_time = time.time() - start_time
    minutes = int(elapsed_time // 60)
    seconds = int(elapsed_time % 60)

    mensaje = f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - TAREA PROGRAMADA: EXITO en {minutes} minutos {seconds} segundos."

    print(mensaje)

    mensajeBody += f"{mensaje}\n"
    mensajeSubject = mensaje

except Exception as e:

    elapsed_time = time.time() - start_time
    minutes = int(elapsed_time // 60)
    seconds = int(elapsed_time % 60)

    mensaje = f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - TAREA PROGRAMADA: FALLO a {minutes} minutos y {seconds} segundos de iniciar. Error: {e}"
    print(mensaje)

    mensajeBody += f"{mensaje}\n"
    mensajeSubject = mensaje

finally:

    if 'driver' in locals():
        driver.quit()

    def enviar_correo():

        credential = keyring.get_credential(entorno, None)
        password = credential.password if credential else ""

        if not password:
            print(f"No se pudo obtener la contrasena en las credenciales: {entorno}")
            return

        try:

            import win32com.client as win32

            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)

            mail.Subject = mensajeSubject
            mail.Body = mensajeBody
            mail.To = email

            mail.Send()
            print("Correo enviado con exito.")

        except Exception as e:
            print(f"Error al enviar el correo: {e}")

    enviar_correo()

