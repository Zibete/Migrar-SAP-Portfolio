import argparse
import glob
import os
import subprocess
import sys
import time

import keyring
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

KEYRING_SERVICE = os.getenv("MIGRAR_SAP_KEYRING_SERVICE", "<REDACTED>")
PORTAL_URL = os.getenv(
    "MIGRAR_SAP_SAP_PORTAL_URL",
    "<REDACTED_URL>",
)

def _parse_args(argv):
    parser = argparse.ArgumentParser(
        description="Inicia sesion en el portal y descarga el archivo .sap."
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="Abre SAP GUI con el archivo .sap descargado.",
    )
    return parser.parse_args(argv)

def _esperar_archivo_sap(timeout=30):
    carpeta = os.path.expanduser("~/Downloads")
    for _ in range(timeout * 2):
        archivos = glob.glob(os.path.join(carpeta, "*.sap"))
        if archivos:
            return archivos[0]
        time.sleep(0.5)
    raise FileNotFoundError("No se encontro el archivo .sap")

def main(argv):
    args = _parse_args(argv)

    cred = keyring.get_credential(KEYRING_SERVICE, None)
    if not cred:
        print("No se encontraron credenciales en keyring.", file=sys.stderr)
        return 1

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--mute-audio")

    driver = webdriver.Chrome(options=chrome_options)
    try:
        driver.get(PORTAL_URL)

        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.NAME, "j_username"))
        )
        driver.find_element(By.NAME, "j_username").send_keys(cred.username)
        driver.find_element(By.NAME, "j_password").send_keys(cred.password)
        driver.find_element(
            By.XPATH,
            "//input[@type='image' and contains(@src, 'login-boton-inicio.png')]",
        ).click()

        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//img[@alt='Acceso a Sistema Central']")
            )
        ).click()

        ruta_sap = _esperar_archivo_sap()

        # Solo la ruta en stdout para que VBA la lea sin ruido.
        print(ruta_sap)

        if args.open:
            subprocess.Popen(ruta_sap, shell=True)

    finally:
        driver.quit()

    return 0

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
