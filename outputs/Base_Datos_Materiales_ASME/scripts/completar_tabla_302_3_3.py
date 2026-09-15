# -*- coding: utf-8 -*-
"""
completar_tabla_302_3_3 - cierra los huecos de la Tabla 302.3.3-1 del ASME B31.3
(«Increased Casting Quality Factors, Ec») en `resources/`.

POR QUE EXISTE ESTE SCRIPT
--------------------------
La Tabla 302.3.3-1 es la que dice hasta cuanto puede SUBIR el factor de calidad
de fundicion Ec cuando se hace examen suplementario. Sin ella, el motor de A-2
mostraria el factor basico como si fuera el definitivo, que es exactamente el
error que la Nota (4) de A-2 advierte que no se cometa. Tenia dos defectos:

1. ESTA DUPLICADA. Existen `table_302_3_3_1.json` y `table_table_302_3_3_1.json`
   con contenido identico. Un motor que cite «la Tabla 302.3.3-1 de resources/»
   sin decir cual de los dos archivos no es auditable.

   El patron es mas amplio: en `chapters/tables/` hay 32 pares con ese doble
   prefijo. En 21 son identicos; en 11 NO lo son, y ahi el archivo de prefijo
   simple trae las filas ya recompuestas y el de prefijo doble los fragmentos
   crudos del corte de linea (`table_302_3_3_2.json` tiene 8 filas legibles
   frente a las 19 partidas de su gemelo). Es decir: el de prefijo SIMPLE es la
   extraccion buena. Este script declara eso para el par de la 302.3.3-1, que es
   el unico que entra en el alcance del plan PLAN-BUSC-A2A3-001, y NO toca los
   otros 31 pares: cambiarlos es una decision que no es de este plan.

2. PERDIO LOS ENCABEZADOS DE COLUMNA. Las dos columnas llegan con
   `"header": null`. El impreso no se capturo y el PDF del codigo no esta en el
   repositorio (copyright de ASME), asi que NO se puede escribir el encabezado
   impreso. Lo que si hay en `resources/` es el texto normativo del propio
   codigo, para. 302.3.3(c), que describe las dos columnas literalmente:

     «Table 302.3.3-1 states the increased casting quality factors, Ec, that may
      be used for various combinations of supplementary examination.»

   De ahi se DERIVAN los dos rotulos, y se guardan como `header_derivado` junto
   con la frase de la que salen. El campo `header` se deja en null: sigue siendo
   verdad que el impreso no se capturo. Un rotulo derivado y declarado como tal
   es auditable; uno escrito de memoria no lo seria.

TERCERA PREGUNTA DE LA FASE 0: ¿HAY EQUIVALENTE PARA Ej?
--------------------------------------------------------
Si, y el script lo verifica leyendolo de `resources/`. El codigo lo dice en
para. 302.3.4(b): «Table 302.3.4-1 also indicates higher joint quality factors
that may be substituted for those in Table A-3 for certain kinds of welds if
additional examination is performed beyond that required by the product
specification.» No es una tabla aparte como la 302.3.3-1 del Ec: son filas de la
propia Tabla 302.3.4-1, y su Nota (1) prohibe el incremento para las juntas 1 y 2.

PERO la extraccion de `table_302_3_4_1.json` esta INSERVIBLE: el cuerpo de la
tabla se colapso dentro de los ENCABEZADOS de columna (el encabezado del factor
es literalmente «Factor, Ej 0.60 [Note (1)] 0.85 0.80 0.90 1.00») y `rows` solo
conserva dos fragmentos. Este script lo comprueba y lo declara. El motor de A-3
por tanto NO ofrece un factor incrementado: cita el parrafo y la nota, y remite
al folio impreso. Es el hueco declarado que exige la regla nº 1 del proyecto.

USO
---
    python completar_tabla_302_3_3.py --resources ..\\..\\..\\resources [--dry-run]

Es idempotente: compara ignorando lo que el propio script escribe.
"""
from __future__ import annotations

import argparse
import copy
import datetime
import hashlib
import json
import re
import sys
from pathlib import Path

TABLAS = Path("asme_b31") / "asme_b31_3" / "chapters" / "tables"
CAPITULOS = Path("asme_b31") / "asme_b31_3" / "chapters"

CANONICO = "table_302_3_3_1.json"
DUPLICADO = "table_table_302_3_3_1.json"
TABLA_EJ = "table_302_3_4_1.json"

# Claves que escribe este script, mas la que escribe declarar_tablas_canonicas.py
# sobre los 32 pares de la carpeta -este par entre ellos-. Se ignoran al comparar
# los dos archivos, para que una segunda corrida, o la del otro script, no hagan
# creer que el contenido ha cambiado.
CLAVES_PROPIAS = ("extraction_amendments", "superseded_by", "canonical_declaration",
                  "extraction_gap")
CLAVES_COL_PROPIAS = ("header_derivado", "fuente_derivacion")

# El parrafo del codigo del que se derivan los dos rotulos. Si el texto de
# resources/ deja de casar con esto, el script para: prefiere no rotular a
# rotular con algo que ya no dice el codigo.
_RE_302_3_3_C = re.compile(
    r"Table 302\.3\.3-1 states the (?P<col2>.+?), that may be used for "
    r"(?P<col1>.+?)\.")
_RE_302_3_4_B = re.compile(
    r"Table 302\.3\.4-1 also indicates (?P<que>.+?)\.", re.S)


# ---------------------------------------------------------------------------
# Lectura del texto normativo
# ---------------------------------------------------------------------------
def _parrafos(o):
    """Todo texto de parrafo del capitulo, sin importar como este anidado."""
    if isinstance(o, dict):
        for k, v in o.items():
            if k == "text" and isinstance(v, str):
                yield v
            else:
                yield from _parrafos(v)
    elif isinstance(o, list):
        for v in o:
            yield from _parrafos(v)


def texto_normativo(res: Path):
    """(frase de 302.3.3(c), frase de 302.3.4(b)) leidas del capitulo II."""
    ruta = res / CAPITULOS / "chapter_02.json"
    if not ruta.is_file():
        raise SystemExit(f"ERROR: falta {ruta}")
    cap = json.loads(ruta.read_text(encoding="utf-8"))
    c33 = c34 = None
    for t in _parrafos(cap):
        if c33 is None and _RE_302_3_3_C.search(t):
            c33 = t.strip()
        if c34 is None and _RE_302_3_4_B.search(t):
            c34 = t.strip()
    if c33 is None:
        raise SystemExit(
            "ERROR: no se encuentra en chapter_02.json la frase de para. 302.3.3(c) "
            "que describe las dos columnas de la Tabla 302.3.3-1. Sin ella no se "
            "puede derivar ningun rotulo, y no se inventa.")
    if c34 is None:
        raise SystemExit(
            "ERROR: no se encuentra en chapter_02.json el para. 302.3.4(b), que es "
            "lo que responde si el codigo publica un mecanismo de incremento para Ej.")
    return c33, c34


def rotulos_derivados(frase: str):
    m = _RE_302_3_3_C.search(frase)
    col2, col1 = m.group("col2").strip(), m.group("col1").strip()
    cap = lambda s: s[0].upper() + s[1:]
    return cap(col1), cap(col2)


# ---------------------------------------------------------------------------
# Comparacion del par duplicado
# ---------------------------------------------------------------------------
def _sin_claves_propias(doc):
    d = copy.deepcopy(doc)
    for k in CLAVES_PROPIAS:
        d.pop(k, None)
    for col in d.get("columns") or []:
        for k in CLAVES_COL_PROPIAS:
            col.pop(k, None)
    return d


def sha256(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def censo_de_pares(base: Path):
    """(pares totales, identicos, distintos) del doble prefijo en chapters/tables."""
    nombres = {p.name for p in base.glob("*.json")}
    pares = []
    for n in sorted(nombres):
        if not n.startswith("table_table_"):
            continue
        simple = "table_" + n[len("table_table_"):]
        if simple in nombres:
            pares.append((simple, n, sha256(base / simple) == sha256(base / n)))
    return pares


# ---------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args(argv)

    base = a.resources / TABLAS
    if not base.is_dir():
        print(f"ERROR: no existe {base}", file=sys.stderr)
        return 2

    ruta_can, ruta_dup = base / CANONICO, base / DUPLICADO
    for r in (ruta_can, ruta_dup):
        if not r.is_file():
            print(f"ERROR: falta {r}", file=sys.stderr)
            return 2

    can = json.loads(ruta_can.read_text(encoding="utf-8"))
    dup = json.loads(ruta_dup.read_text(encoding="utf-8"))
    if _sin_claves_propias(can) != _sin_claves_propias(dup):
        print("ERROR: los dos archivos del par YA NO son el mismo contenido.\n"
              "       Elegir uno dejaria de ser inocuo. Revise cual es el bueno\n"
              "       antes de declarar nada.", file=sys.stderr)
        return 2
    if len(can["rows"]) != 6:
        print(f"ERROR: la Tabla 302.3.3-1 deberia traer 6 filas y trae "
              f"{len(can['rows'])}", file=sys.stderr)
        return 2

    c33, c34 = texto_normativo(a.resources)
    col1, col2 = rotulos_derivados(c33)

    pares = censo_de_pares(base)
    n_ident = sum(1 for _, _, ok in pares if ok)

    # --- Tabla 302.3.4-1: se comprueba el estado real de su extraccion --------
    ruta_ej = base / TABLA_EJ
    ej_utilizable = None
    if ruta_ej.is_file():
        ej = json.loads(ruta_ej.read_text(encoding="utf-8"))
        # El sintoma es inconfundible: los factores estan DENTRO del encabezado
        # de columna en vez de en las filas.
        cabeceras = " ".join(str((c or {}).get("header") or "")
                             for c in ej.get("columns") or [])
        ej_utilizable = not re.search(r"Factor,\s*Ej\s*0\.\d", cabeceras)

    print("=" * 78)
    print("Tabla 302.3.3-1  —  Increased Casting Quality Factors, Ec")
    print("=" * 78)
    print(f"  archivo canonico declarado : {CANONICO}")
    print(f"  duplicado                  : {DUPLICADO}  (sha256 {sha256(ruta_dup)[:16]}…)")
    print(f"  contenido identico         : si (elegir uno no pierde nada)")
    print(f"  filas                      : {len(can['rows'])}")
    print(f"  factores Ec                : "
          f"{[r['column_2'] for r in can['rows']]}")
    print()
    print("  rotulos DERIVADOS de para. 302.3.3(c) (el impreso no se capturo):")
    print(f"    column_1 -> {col1}")
    print(f"    column_2 -> {col2}")
    print()
    print(f"  censo del doble prefijo en chapters/tables: {len(pares)} pares, "
          f"{n_ident} identicos, {len(pares) - n_ident} distintos.")
    print("    En los distintos, el de prefijo SIMPLE trae las filas recompuestas.")
    print("    Fuera del alcance de este plan: no se tocan.")
    print()
    print("  ¿publica el codigo un incremento equivalente para Ej?  SI")
    print(f"    para. 302.3.4(b): {c34[:200]}…")
    if ej_utilizable is False:
        print("    PERO table_302_3_4_1.json esta INSERVIBLE: el cuerpo de la tabla")
        print("    se colapso dentro de los encabezados de columna. El motor de A-3")
        print("    declara el hueco y remite al folio impreso; no ofrece numero.")
    elif ej_utilizable:
        print("    table_302_3_4_1.json parece utilizable: revise el motor de A-3.")

    if a.dry_run:
        print("\n(--dry-run: no se ha escrito nada)")
        return 0

    hoy = datetime.date.today().isoformat()
    for c, rot in zip(can["columns"], (col1, col2)):
        c["header_derivado"] = rot
        c["fuente_derivacion"] = (
            "chapters/chapter_02.json, para. 302.3.3(c): " + c33)
    can["extraction_amendments"] = {
        "fecha": hoy,
        "script": "completar_tabla_302_3_3.py",
        "alcance": "Solo metadatos. Ningun valor ni nombre de fila se altera: los "
                   "seis factores Ec y sus combinaciones de examen se conservan tal "
                   "como estaban cargados (regla 9 del proyecto).",
        "archivo_canonico": CANONICO,
        "duplicado_descartado": {
            "archivo": DUPLICADO,
            "sha256": sha256(ruta_dup),
            "motivo": "mismo contenido, prefijo 'table_' aplicado dos veces por la "
                      "extraccion. Se cita SIEMPRE el de prefijo simple.",
        },
        "encabezados": "columns[].header sigue en null porque el impreso no se "
                       "capturo y el PDF del codigo no esta en el repositorio. Los "
                       "rotulos van en header_derivado, con la frase del codigo de "
                       "la que salen en fuente_derivacion.",
        "incremento_de_ej": {
            "existe": True,
            "donde": "para. 302.3.4(b) y Table 302.3.4-1",
            "texto": c34,
            "extraccion_utilizable": bool(ej_utilizable),
            "nota": "El cuerpo de table_302_3_4_1.json se colapso dentro de los "
                    "encabezados de columna: el motor de A-3 declara el hueco y "
                    "remite al folio impreso en vez de ofrecer un factor.",
        },
        "censo_doble_prefijo": {
            "pares": len(pares), "identicos": n_ident,
            "distintos": len(pares) - n_ident,
            "criterio": "el archivo de prefijo SIMPLE es la extraccion recompuesta; "
                        "el de prefijo doble conserva los fragmentos del corte de "
                        "linea. Solo se declara el par de la 302.3.3-1.",
        },
    }
    ruta_can.write_text(json.dumps(can, ensure_ascii=False, indent=1),
                        encoding="utf-8")
    dup["superseded_by"] = {
        "archivo": CANONICO,
        "motivo": "duplicado por doble prefijo de la extraccion. Cite siempre "
                  "table_302_3_3_1.json; este archivo se conserva como testigo.",
        "fecha": hoy,
        "script": "completar_tabla_302_3_3.py",
    }
    ruta_dup.write_text(json.dumps(dup, ensure_ascii=False, indent=1),
                        encoding="utf-8")
    escritos = [ruta_can.name, ruta_dup.name]

    # El hueco de la Tabla 302.3.4-1 se declara DENTRO de su propio archivo, no
    # solo en el motor: quien abra table_302_3_4_1.json tiene que ver, sin salir
    # de el, que su cuerpo no esta y por que. Un hueco que solo se nombra en la
    # herramienta que lo consume se cita a ciegas desde cualquier otra.
    if ruta_ej.is_file() and ej_utilizable is False:
        ej = json.loads(ruta_ej.read_text(encoding="utf-8"))
        ej["extraction_gap"] = {
            "fecha": hoy,
            "script": "completar_tabla_302_3_3.py",
            "que_falta": "el CUERPO de la tabla. Las filas se colapsaron dentro "
                         "de los encabezados de columna y `rows` solo conserva "
                         "dos fragmentos sueltos.",
            "sintoma": "el encabezado de la columna del factor contiene los "
                       "propios valores: 'Factor, Ej 0.60 [Note (1)] 0.85 0.80 "
                       "0.90 1.00'. Es la comprobacion que hace este script.",
            "que_si_esta": "los encabezados con el texto crudo de las columnas "
                           "-tipo de junta, tipo de costura y examen- y la Nota "
                           "(1), que prohibe incrementar el factor para las "
                           "juntas 1 y 2.",
            "consecuencia": "no se puede ofrecer el factor Ej incrementado del "
                            "para. 302.3.4(b). Buscar_Ej_A3 declara el hueco, "
                            "transcribe el parrafo y la Nota (1) y remite al "
                            "folio impreso; no aproxima ningun valor.",
            "como_cerrarlo": "hace falta el folio impreso de esta tabla -pagina "
                             f"{'-'.join(str(p) for p in ej.get('pdf_pages') or [])} "
                             "del PDF del codigo- o su OCR estructurado, como se "
                             "hizo con el Apendice C. Ni el PDF de los capitulos "
                             "ni su OCR estan en el repositorio.",
        }
        ruta_ej.write_text(json.dumps(ej, ensure_ascii=False, indent=1),
                           encoding="utf-8")
        escritos.append(ruta_ej.name)

    print("\nEscritos " + ", ".join(escritos) + ".")
    return 0


if __name__ == "__main__":
    sys.exit(main())
