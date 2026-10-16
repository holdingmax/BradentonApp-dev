"""
Corre los lectores sobre los PDFs de BRADENTON_TEST_DATA y guarda su salida
como la esperada por tests/test_lectores.py.

    python tests/regenerar_esperados.py            todos los casos
    python tests/regenerar_esperados.py cec        solo los que contienen "cec"
    python tests/regenerar_esperados.py --ver      no guarda: solo dice qué cambió

Guardar solo después de revisar que lo que cambió es una mejora.
"""

import os
import sys
import tempfile

# Igual que conftest.py: nunca contra los datos reales.
os.environ["BRADENTON_DATA_DIR"] = tempfile.mkdtemp(prefix="bradenton_pruebas_")
os.environ.pop("DATABASE_URL", None)
sys.path[:0] = [os.path.dirname(os.path.dirname(os.path.abspath(__file__))), os.path.dirname(os.path.abspath(__file__))]

import lectores  # noqa: E402


def main(args):
    only_show = "--ver" in args
    filters = [a.lower() for a in args if not a.startswith("--")]
    cases = lectores.load_cases()
    if not cases:
        sys.exit(f"Falta {lectores.DATA_ENV} (la carpeta con casos.json, pdfs/ y esperado/).")
    for case in cases:
        cid = lectores.case_id(case)
        if filters and not any(f in cid.lower() for f in filters):
            continue
        output = lectores.run_case(case)
        try:
            state = "igual" if lectores.load_expected(case) == output else "CAMBIÓ"
        except FileNotFoundError:
            state = "nuevo"
        if state != "igual" and not only_show:
            lectores.save_expected(case, output)
        print(f"{state:7} {cid}", flush=True)


if __name__ == "__main__":
    main(sys.argv[1:])
