import argparse
import json
import os

DEFAULT_INDEX_FILE = os.path.join(os.path.dirname(__file__), "archivo_index.json")
DEFAULT_PATHS = [
    r"<REDACTED_PATH>",
    r"<REDACTED_PATH>",
]
DEFAULT_EXTS = [".pdf", ".xlsx", ".docx"]

def _parse_args(argv):
    parser = argparse.ArgumentParser(
        description="Crea un indice JSON de archivos para busquedas rapidas."
    )
    parser.add_argument(
        "--path",
        action="append",
        dest="paths",
        default=[],
        help="Ruta a indexar (se puede repetir).",
    )
    parser.add_argument(
        "--ext",
        action="append",
        dest="exts",
        default=[],
        help="Extension a incluir (ej: .pdf).",
    )
    parser.add_argument(
        "--no-filter",
        action="store_true",
        help="No filtrar por extension (indexa todo).",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_INDEX_FILE,
        help="Ruta del archivo de salida JSON.",
    )
    return parser.parse_args(argv)

def _iter_files(root):
    for base, dirs, files in os.walk(root, onerror=lambda _: None):
        for name in files:
            yield os.path.join(base, name)

def crear_indice(paths, output_path, extensiones_filtradas=None):
    archivos = []
    extensiones = (
        [ext.lower() for ext in extensiones_filtradas]
        if extensiones_filtradas
        else None
    )

    for ruta_inicial in paths:
        for full_path in _iter_files(ruta_inicial):
            if extensiones and not any(full_path.lower().endswith(ext) for ext in extensiones):
                continue
            archivos.append(full_path)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(archivos, f, indent=2)
    print(f"Indice creado con {len(archivos)} archivos.")

def main(argv):
    args = _parse_args(argv)
    paths = args.paths if args.paths else DEFAULT_PATHS
    exts = None if args.no_filter else (args.exts if args.exts else DEFAULT_EXTS)
    crear_indice(paths, args.output, extensiones_filtradas=exts)

if __name__ == "__main__":
    import sys

    main(sys.argv[1:])
