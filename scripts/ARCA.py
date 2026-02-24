import argparse
import os
import sys
import time

from selenium import webdriver
from selenium.webdriver.common.by import By

def _env_bool(name, default=False):
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in ("1", "true", "yes", "y", "on")

def _wait_for_user(wait_seconds):
    try:
        if sys.stdin and sys.stdin.isatty():
            input("Presiona Enter para cerrar el navegador...")
            return
    except Exception:
        pass
    time.sleep(wait_seconds)

def _parse_args(argv):
    parser = argparse.ArgumentParser(
        description="Completa los datos de verificacion en ARCA (AFIP)."
    )
    parser.add_argument("link", help="URL de ARCA (cae.aspx o caea.aspx).")
    parser.add_argument("cuit", help="CUIT emisor.")
    parser.add_argument("cae", help="CAE/CAEA.")
    parser.add_argument("fecha", help="Fecha de emision (dd/mm/aaaa).")
    parser.add_argument("tipo_doc", help="Tipo de comprobante.")
    parser.add_argument("pdv", help="Punto de venta.")
    parser.add_argument("comprobante", help="Numero de comprobante.")
    parser.add_argument("importe", help="Importe total.")
    parser.add_argument("cuit_pae", help="CUIT del proveedor.")
    parser.add_argument(
        "--wait-seconds",
        type=int,
        default=int(os.getenv("MIGRAR_SAP_ARCA_WAIT", "3600")),
        help="Segundos de espera antes de cerrar (si no hay consola).",
    )
    parser.add_argument(
        "--detach",
        action="store_true",
        default=_env_bool("MIGRAR_SAP_ARCA_DETACH", False),
        help="Deja el navegador abierto al finalizar el script.",
    )
    return parser.parse_args(argv)

def main(argv):
    args = _parse_args(argv)

    options = webdriver.ChromeOptions()
    if args.detach:
        options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    # Navega a la pagina de validacion.
    driver.get(args.link)

    campo = driver.find_element(By.ID, "p_CUIT")
    campo.send_keys(args.cuit)

    if args.link.endswith("caea.aspx"):
        id_cae = "p_CAEA"
    elif args.link.endswith("cae.aspx"):
        id_cae = "p_CAE"
    else:
        raise ValueError("URL invalida: se esperaba cae.aspx o caea.aspx.")

    campo = driver.find_element(By.ID, id_cae)
    campo.send_keys(args.cae)

    campo = driver.find_element(By.ID, "p_fch_emision")
    campo.send_keys(args.fecha)

    campo = driver.find_element(By.ID, "ctl00_cphBody_p_tipo_cbte")
    campo.send_keys(args.tipo_doc)

    campo = driver.find_element(By.ID, "p_pto_vta")
    campo.send_keys(args.pdv)

    campo = driver.find_element(By.ID, "p_nro_cbte")
    campo.send_keys(args.comprobante)

    campo = driver.find_element(By.ID, "p_importe")
    campo.send_keys(args.importe)

    campo = driver.find_element(By.ID, "ctl00_cphBody_p_tipo_doc")
    campo.send_keys("80")

    campo = driver.find_element(By.ID, "p_nro_doc")
    campo.send_keys(args.cuit_pae)

    campo = driver.find_element(By.ID, "ctl00_cphBody_txtsolucion")
    campo.click()

    _wait_for_user(args.wait_seconds)

    if not args.detach:
        driver.quit()

if __name__ == "__main__":
    main(sys.argv[1:])
