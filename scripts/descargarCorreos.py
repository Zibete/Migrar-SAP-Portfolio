import argparse
import datetime as dt
import os
import unicodedata

import fitz
import win32com.client

DEFAULT_BASE_PATH = r"<REDACTED_PATH>"
DEFAULT_ACCOUNT = "<REDACTED>"
DEFAULT_INBOX = "Bandeja de entrada"
DEFAULT_SUBFOLDER = "Facturacion"
DEFAULT_DAYS = 50
SPLIT_CODES = {"<REDACTED_ID_05>", "<REDACTED_ID_04>", "<REDACTED_ID_01>"}

def _parse_args(argv):
    parser = argparse.ArgumentParser(
        description="Descarga PDFs de Outlook y los guarda por proveedor."
    )
    parser.add_argument(
        "--base-path",
        default=DEFAULT_BASE_PATH,
        help="Ruta base de salida.",
    )
    parser.add_argument(
        "--account",
        default=DEFAULT_ACCOUNT,
        help="Nombre de la cuenta de Outlook.",
    )
    parser.add_argument(
        "--inbox",
        default=DEFAULT_INBOX,
        help="Nombre de la carpeta de entrada.",
    )
    parser.add_argument(
        "--subfolder",
        default=DEFAULT_SUBFOLDER,
        help="Subcarpeta dentro de la bandeja de entrada.",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=DEFAULT_DAYS,
        help="Dias hacia atras a considerar.",
    )
    return parser.parse_args(argv)

def _normalize_name(value):
    value = value.strip().lower()
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", value) if not unicodedata.combining(ch)
    )

def _find_account(outlook, account_name):
    for account in outlook.Folders:
        if account.Name == account_name:
            return account
    return None

def _find_folder(container, folder_name):
    target = _normalize_name(folder_name)
    for folder in container.Folders:
        if _normalize_name(folder.Name) == target:
            return folder
    return None

def _resolve_facturacion_dir(carpeta_proveedor):
    for nombre in ("Facturacion", "Facturación"):
        ruta = os.path.join(carpeta_proveedor, nombre)
        if os.path.exists(ruta):
            return ruta
    return os.path.join(carpeta_proveedor, "Facturacion")

def main(argv):
    args = _parse_args(argv)

    base_path = args.base_path
    anio_actual = str(dt.datetime.now().year)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    facturas_account = _find_account(outlook, args.account)

    if not facturas_account:
        raise Exception("No se encontro la cuenta de facturas.")

    inbox = _find_folder(facturas_account, args.inbox)
    if not inbox:
        raise Exception(f"No se encontro la carpeta de entrada: {args.inbox}")

    facturacion_folder = _find_folder(inbox, args.subfolder)
    if not facturacion_folder:
        raise Exception(f"No se encontro la subcarpeta: {args.subfolder}")

    limite = dt.datetime.now() - dt.timedelta(days=args.days)

    for carpeta in facturacion_folder.Folders:
        print(f"Procesando carpeta Outlook: {carpeta.Name}")

        partes = carpeta.Name.split(" ", 1)
        if len(partes) < 2:
            print(f"  Carpeta omitida (formato inesperado): {carpeta.Name}")
            continue

        codigo, nombre = partes[0], partes[1]
        path_anio = os.path.join(base_path, anio_actual)
        carpeta_proveedor = None

        if os.path.exists(path_anio):
            for nombre_carpeta in os.listdir(path_anio):
                if codigo in nombre_carpeta:
                    carpeta_proveedor = os.path.join(path_anio, nombre_carpeta)
                    break

        if not carpeta_proveedor:
            nombre_nuevo = f"{nombre} ({codigo})"
            carpeta_proveedor = os.path.join(path_anio, nombre_nuevo)
            os.makedirs(carpeta_proveedor, exist_ok=True)
            print(f"Carpeta creada: {carpeta_proveedor}")

        carpeta_facturacion = _resolve_facturacion_dir(carpeta_proveedor)
        os.makedirs(carpeta_facturacion, exist_ok=True)

        contador_pdf = 1

        for mensaje in carpeta.Items:
            if not hasattr(mensaje, "Unread"):
                continue
            if not mensaje.Unread:
                continue
            if mensaje.ReceivedTime < limite:
                continue

            fecha_correo = mensaje.ReceivedTime.strftime("%d.%m.%Y")
            base_nombre = f"Fecha base {fecha_correo}"

            for adjunto in mensaje.Attachments:
                if not adjunto.FileName.lower().endswith(".pdf"):
                    continue

                nombre_final = f"{base_nombre}-{contador_pdf}.pdf"
                ruta_guardado = os.path.join(carpeta_facturacion, nombre_final)
                adjunto.SaveAsFile(ruta_guardado)
                print(f"PDF guardado: {ruta_guardado}")
                contador_pdf += 1

                mensaje.Unread = False
                mensaje.Save()

                doc = fitz.open(ruta_guardado)
                total_paginas = doc.page_count

                if total_paginas > 1 and codigo in SPLIT_CODES:
                    for i in range(total_paginas):
                        nueva_ruta = ruta_guardado.replace(".pdf", f"_pag{i + 1}.pdf")
                        nuevo_doc = fitz.open()
                        nuevo_doc.insert_pdf(doc, from_page=i, to_page=i)
                        nuevo_doc.save(nueva_ruta)
                        nuevo_doc.close()
                    doc.close()
                    os.remove(ruta_guardado)
                    print(f"    PDF dividido en {total_paginas} paginas.")
                else:
                    doc.close()

if __name__ == "__main__":
    import sys

    main(sys.argv[1:])
