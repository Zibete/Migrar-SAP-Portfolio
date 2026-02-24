import argparse
import json
import os

DEFAULT_INDEX_FILE = os.path.join(os.path.dirname(__file__), "archivo_index.json")

def buscar_en_indice(index_path, palabra_clave, limit):
    with open(index_path, "r", encoding="utf-8") as f:
        archivos = json.load(f)

    resultados = [
        archivo
        for archivo in archivos
        if palabra_clave.lower() in os.path.basename(archivo).lower()
    ]

    for r in resultados[:limit]:
        print(r)
    print(f"\n{len(resultados)} archivos encontrados.")

def _parse_args(argv):
    parser = argparse.ArgumentParser(description="Busca en el indice JSON.")
    parser.add_argument("query", help="Palabra clave a buscar.")
    parser.add_argument(
        "--index",
        default=DEFAULT_INDEX_FILE,
        help="Ruta al archivo de indice JSON.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=50,
        help="Cantidad maxima de resultados a mostrar.",
    )
    return parser.parse_args(argv)

def main(argv):
    args = _parse_args(argv)
    buscar_en_indice(args.index, args.query, args.limit)

if __name__ == "__main__":
    import sys

    main(sys.argv[1:])
