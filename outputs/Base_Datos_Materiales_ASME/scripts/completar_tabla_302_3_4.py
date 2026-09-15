# -*- coding: utf-8 -*-
"""
completar_tabla_302_3_4 - reconstruye el CUERPO de la Tabla 302.3.4-1 del ASME
B31.3 («Longitudinal Weld Joint Quality Factor, Ej») en `resources/`.

POR QUE EXISTE ESTE SCRIPT
--------------------------
`completar_tabla_302_3_3.py` ya declaro el hueco: la extraccion de
`table_302_3_4_1.json` colapso el cuerpo de la tabla DENTRO de los encabezados
de columna (el encabezado del factor es literalmente «Factor, Ej 0.60
[Note (1)] 0.85 0.80 0.90 1.00») y `rows` solo conservaba dos fragmentos
sueltos. Sin el cuerpo, Buscar_Ej_A3 no podia ofrecer ningun factor
incrementado del para. 302.3.4(b): declaraba el hueco, transcribia el parrafo y
la Nota (1), y remitia al folio impreso.

Cerrarlo de verdad exige el folio impreso -pagina 51 del PDF del codigo- o su
OCR estructurado. Ni el PDF de los capitulos del B31.3 ni su OCR estan en el
repositorio (copyright de ASME). El folio lo aporto el ingeniero: un PDF de
tres paginas con el fragmento «Weld Joint and Casting Quality Factors» de
ASME B31.3-2024, que en su pagina 3 (folio impreso «18» del codigo) reproduce
la Tabla 302.3.4-1 completa, con reglas de celda impresas -no es un problema de
geometria de fuente proporcional como el de la Seccion II (ver secii_tablas.py):
las diez filas son legibles sin ambiguedad, cada una en su celda.

Este script NO hace OCR ni inferencia de columnas por bbox: transcribe
literalmente esas diez filas (verificadas contra el PDF citado, cuyo sha256
queda registrado para auditoria) y las escribe en `columns`/`rows` con el
mismo contrato de forma que usa `build_db_materiales.build_factores` para leer
cualquier tabla de este directorio.

USO
---
    python completar_tabla_302_3_4.py --resources ..\\..\\..\\resources \\
        --pdf "C:\\ruta\\ASME B31.3-WELD JOINT AND CASTING QUALITY FACTORS.pdf" \\
        [--dry-run]

Es idempotente: si el cuerpo ya esta reconstruido (10 filas, sin el sintoma de
cabecera colapsada), no vuelve a escribir nada y lo informa.
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import re
import sys
from pathlib import Path

TABLAS = Path("asme_b31") / "asme_b31_3" / "chapters" / "tables"
CANONICO = "table_302_3_4_1.json"

# Transcripcion literal de la Tabla 302.3.4-1, verificada contra el PDF citado
# en --pdf (pagina 3 de ese archivo, folio impreso "18" == folio 51 del codigo,
# que es el que ya declaraba `pdf_pages` en el JSON original). Diez filas: los
# No. 1 y 2 no tienen subdivision; el No. 3 se parte en (a)/(b), cada uno con
# tres niveles de examen; el No. 4 tiene dos. Los cinco primeros valores en
# orden -0.60, 0.85, 0.80, 0.90, 1.00- coinciden exactamente con los que la
# extraccion rota ya habia dejado ver en el encabezado colapsado: es la
# comprobacion cruzada de que la transcripcion nueva es consistente con lo
# poco que la vieja alcanzo a capturar antes de perder el resto.
FILAS = [
    dict(no="1", tipo_junta="Furnace butt weld, continuous weld",
         tipo_costura="Straight",
         examen="As required by listed specification",
         factor_ej=0.60, nota="(1)"),
    dict(no="2", tipo_junta="Electric resistance weld",
         tipo_costura="Straight or spiral (helical seam)",
         examen="As required by listed specification",
         factor_ej=0.85, nota="(1)"),
    dict(no="3(a)",
         tipo_junta="Electric fusion weld \u2014 (a) Single butt weld "
                    "(with or without filler metal)",
         tipo_costura="Straight or spiral (helical seam)",
         examen="As required by listed specification or this Code",
         factor_ej=0.80, nota=None),
    dict(no="3(a)",
         tipo_junta="Electric fusion weld \u2014 (a) Single butt weld "
                    "(with or without filler metal)",
         tipo_costura="Straight or spiral (helical seam)",
         examen="Additionally spot radiographed in accordance with para. 341.5.1",
         factor_ej=0.90, nota=None),
    dict(no="3(a)",
         tipo_junta="Electric fusion weld \u2014 (a) Single butt weld "
                    "(with or without filler metal)",
         tipo_costura="Straight or spiral (helical seam)",
         examen="Additionally 100% radiographed in accordance with para. "
                "344.5.1 and Table 341.3.2-1",
         factor_ej=1.00, nota=None),
    dict(no="3(b)",
         tipo_junta="Electric fusion weld \u2014 (b) Double butt weld "
                    "(with or without filler metal)",
         tipo_costura="Straight or spiral (helical seam) (except as provided "
                      "in 4 below)",
         examen="As required by listed specification or this Code",
         factor_ej=0.85, nota=None),
    dict(no="3(b)",
         tipo_junta="Electric fusion weld \u2014 (b) Double butt weld "
                    "(with or without filler metal)",
         tipo_costura="Straight or spiral (helical seam) (except as provided "
                      "in 4 below)",
         examen="Additionally spot radiographed in accordance with para. 341.5.1",
         factor_ej=0.90, nota=None),
    dict(no="3(b)",
         tipo_junta="Electric fusion weld \u2014 (b) Double butt weld "
                    "(with or without filler metal)",
         tipo_costura="Straight or spiral (helical seam) (except as provided "
                      "in 4 below)",
         examen="Additionally 100% radiographed in accordance with para. "
                "344.5.1 and Table 341.3.2-1",
         factor_ej=1.00, nota=None),
    dict(no="4",
         tipo_junta="Specific specification \u2014 API 5L, electric fusion "
                    "weld, double butt seam",
         tipo_costura="Straight (with one or two seams) or spiral (helical seam)",
         examen="As required by specification",
         factor_ej=0.95, nota=None),
    dict(no="4",
         tipo_junta="Specific specification \u2014 API 5L, electric fusion "
                    "weld, double butt seam",
         tipo_costura="Straight (with one or two seams) or spiral (helical seam)",
         examen="Additionally 100% radiographed in accordance with para. "
                "344.5.1 and Table 341.3.2-1",
         factor_ej=1.00, nota=None),
]

COLUMNAS = [
    ("no", "No."),
    ("type_of_joint", "Type of Joint"),
    ("type_of_seam", "Type of Seam"),
    ("examination", "Examination"),
    ("factor_ej", "Factor, Ej"),
    ("note", "Note"),
]

_SINTOMA = re.compile(r"Factor,\s*Ej\s*0\.\d")


def sha256(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def esta_colapsada(d) -> bool:
    cabeceras = " ".join(str((c or {}).get("header") or "")
                         for c in d.get("columns") or [])
    return bool(_SINTOMA.search(cabeceras))


def ya_reconstruida(d) -> bool:
    filas = d.get("rows") or []
    return (len(filas) == len(FILAS)
            and all(r.get("factor_ej") is not None for r in filas)
            and not esta_colapsada(d))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--pdf", required=True, type=Path,
                    help="PDF con el folio impreso de la Tabla 302.3.4-1, "
                         "aportado por el ingeniero (no versionado en el repo)")
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args(argv)

    if not a.pdf.is_file():
        print(f"ERROR: no existe {a.pdf}", file=sys.stderr)
        return 2
    ruta = a.resources / TABLAS / CANONICO
    if not ruta.is_file():
        print(f"ERROR: no existe {ruta}", file=sys.stderr)
        return 2

    d = json.loads(ruta.read_text(encoding="utf-8"))
    pdf_hash = sha256(a.pdf)

    print("=" * 78)
    print("Tabla 302.3.4-1  —  Longitudinal Weld Joint Quality Factor, Ej")
    print("=" * 78)
    print(f"  archivo               : {CANONICO}")
    print(f"  PDF fuente            : {a.pdf.name}")
    print(f"  PDF sha256            : {pdf_hash}")
    print(f"  filas transcritas     : {len(FILAS)}")

    if ya_reconstruida(d):
        print("\n  Ya esta reconstruida (10 filas, sin el sintoma de cabecera "
              "colapsada). No se toca nada.")
        return 0
    if not esta_colapsada(d):
        print("\nERROR: el archivo no muestra el sintoma conocido (cabecera con "
              "'Factor, Ej 0.6...') pero tampoco tiene las 10 filas esperadas.\n"
              "       Alguien lo toco de otra forma; revise antes de sobreescribir.",
              file=sys.stderr)
        return 2

    if a.dry_run:
        print("\n(--dry-run: no se ha escrito nada)")
        return 0

    hoy = datetime.date.today().isoformat()
    d["columns"] = [{"key": k, "header": h} for k, h in COLUMNAS]
    d["rows"] = [
        {
            "no": f["no"],
            "type_of_joint": f["tipo_junta"],
            "type_of_seam": f["tipo_costura"],
            "examination": f["examen"],
            "factor_ej": f["factor_ej"],
            "note": f["nota"],
        }
        for f in FILAS
    ]
    d["row_count"] = len(FILAS)
    d.pop("extraction_gap", None)
    d["extraction_amendments"] = {
        "fecha": hoy,
        "script": "completar_tabla_302_3_4.py",
        "alcance": "Reconstruye columns/rows: la extraccion original colapso el "
                   "cuerpo de la tabla dentro de los encabezados de columna "
                   "(ver el `extraction_gap` que este script retira, declarado "
                   "por completar_tabla_302_3_3.py). No toca notes, title, "
                   "table_id ni canonical_declaration.",
        "metodo": "Transcripcion literal de las 10 filas contra el folio "
                  "impreso, aportado por el ingeniero como PDF de 3 paginas "
                  "(pagina 3 de ese archivo = folio impreso \"18\" del codigo "
                  "= folio 51, el que ya declaraba `pdf_pages`). La tabla tiene "
                  "reglas de celda impresas: no es un problema de geometria de "
                  "fuente proporcional (a diferencia de la Seccion II, ver "
                  "secii_tablas.py); las diez filas se leen sin ambiguedad.",
        "verificacion_cruzada": "los cinco primeros valores de Ej en orden de "
                                "aparicion -0.60, 0.85, 0.80, 0.90, 1.00- "
                                "coinciden exactamente con los que la "
                                "extraccion colapsada ya mostraba en su "
                                "encabezado roto ('Factor, Ej 0.60 [Note (1)] "
                                "0.85 0.80 0.90 1.00'), antes de perder el "
                                "resto de la tabla.",
        "fuente": {
            "archivo": a.pdf.name,
            "sha256": pdf_hash,
            "descripcion": "PDF de 3 paginas, 'ASME B31.3-WELD JOINT AND "
                           "CASTING QUALITY FACTORS', aportado por el "
                           "ingeniero fuera del repositorio (copyright ASME, "
                           "no se versiona).",
            "pagina_del_pdf_aportado": 3,
            "folio_impreso_del_codigo": "18",
        },
        "nota_1": "No es permitido incrementar el factor Ej de las juntas 1 "
                  "(furnace butt weld) y 2 (ERW) por examen adicional: lo dice "
                  "la propia Nota (1) de la tabla, ya presente en `notes`.",
    }
    ruta.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nEscrito {CANONICO}: {len(FILAS)} filas, cuerpo reconstruido.")
    print("El duplicado table_table_302_3_4_1.json NO se toca: sigue como "
          "testigo del fragmento crudo (regla del proyecto).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
