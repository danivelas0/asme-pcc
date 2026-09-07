# -*- coding: utf-8 -*-
"""
completar_seccion_ii - corrige las citas de pagina de ASME BPVC Seccion II-D en
`resources/ASME_BPVC/Sec_II/`.

POR QUE EXISTE ESTE SCRIPT
--------------------------
`verificar_seccion_ii.py` contrasta cada tabla contra su PDF buscando el
`table_id` en la pagina que la tabla dice ocupar. De las 27 tablas de cada
edicion, 19 caen exactamente donde declaran y **8 apuntan una pagina de mas**:
quien siga la cita de la Tabla TE-2 aterriza en la TE-3.

Son las mismas ocho en las dos ediciones -TE-2, TE-3, TE-5, TM-2, TM-3, TM-4,
TM-5 e Y-2- y el desplazamiento es siempre el mismo, -1. No es una conjetura:
se midio tabla por tabla buscando el rotulo impreso en una ventana de mas menos
tres paginas alrededor de la declarada.

Es un defecto de CITA, no de valores: el `row_count` de las 27 tablas coincide
con las filas cargadas, y los valores ya los audita `verificar.py` fila a fila
contra la hoja.

USO
---
    python completar_seccion_ii.py --resources ..\\..\\..\\resources [--dry-run]

Es idempotente: comprueba el estado de partida y se puede volver a correr.
"""

import argparse
import datetime
import json
import sys
from pathlib import Path

SEC_II = Path("ASME_BPVC") / "Sec_II"

EDICIONES = ("bpvc_ii_d_metric_2025", "bpvc_ii_d_customary_2025")

# Tabla -> desplazamiento medido contra el PDF. Mismo en las dos ediciones.
DESPLAZADAS = {
    "table_te_2.json": -1,
    "table_te_3.json": -1,
    "table_te_5.json": -1,
    "table_tm_2.json": -1,
    "table_tm_3.json": -1,
    "table_tm_4.json": -1,
    "table_tm_5.json": -1,
    "table_y_2.json": -1,
}

# Paginas declaradas ANTES de la correccion, por edicion y tabla. Sirve de
# candado: si el archivo ya no trae eso, algo cambio y el script se para.
ESPERADO = {
    "bpvc_ii_d_metric_2025": {
        "table_te_2.json": 1211, "table_te_3.json": 1212, "table_te_5.json": 1223,
        "table_tm_2.json": 1241, "table_tm_3.json": 1242, "table_tm_4.json": 1243,
        "table_tm_5.json": 1244, "table_y_2.json": 1204,
    },
    "bpvc_ii_d_customary_2025": {
        "table_te_2.json": 1207, "table_te_3.json": 1208, "table_te_5.json": 1219,
        "table_tm_2.json": 1237, "table_tm_3.json": 1238, "table_tm_4.json": 1239,
        "table_tm_5.json": 1240, "table_y_2.json": 1200,
    },
}


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args()

    total = 0
    for ed in EDICIONES:
        base = a.resources / SEC_II / ed
        if not base.is_dir():
            print("ERROR: no existe %s" % base, file=sys.stderr)
            return 2
        print("=" * 74)
        print(ed)
        print("=" * 74)
        n = 0
        for nombre, off in sorted(DESPLAZADAS.items()):
            ruta = base / nombre
            doc = json.loads(ruta.read_text(encoding="utf-8"))
            p = doc.get("pdf_pages")
            if not (isinstance(p, list) and len(p) == 2):
                raise SystemExit("%s: `pdf_pages` inesperado: %r" % (nombre, p))
            previsto = ESPERADO[ed][nombre]
            if p[0] == previsto + off:
                continue  # ya corregida
            if p[0] != previsto:
                raise SystemExit(
                    "%s / %s: se esperaba pdf_pages[0] = %d (sin corregir) o %d "
                    "(corregida), y hay %d. Vuelva a medirlo contra el PDF."
                    % (ed, nombre, previsto, previsto + off, p[0]))
            doc["pdf_pages"] = [p[0] + off, p[1] + off]
            doc["extraction_amendments"] = {
                "fecha": datetime.date.today().isoformat(),
                "script": "completar_seccion_ii.py",
                "alcance": "Solo la cita de pagina. Ningun valor ni ninguna fila "
                           "se toca: `row_count` ya coincidia con las filas "
                           "cargadas en las 27 tablas.",
                "cambio": "pdf_pages %s -> %s, medido buscando el rotulo impreso "
                          "de la tabla en el PDF de origen" % (p, doc["pdf_pages"]),
            }
            if not a.dry_run:
                ruta.write_text(json.dumps(doc, ensure_ascii=False, indent=2) + "\n",
                                encoding="utf-8")
            print("  %-16s %s -> %s" % (nombre, p, doc["pdf_pages"]))
            n += 1
        if not n:
            print("  ninguna (ya estaban corregidas)")
        total += n
        print()

    print("%d citas corregidas%s." % (total, " (dry-run, no se escribio nada)"
                                      if a.dry_run else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
