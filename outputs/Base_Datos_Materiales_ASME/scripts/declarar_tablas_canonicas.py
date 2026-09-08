# -*- coding: utf-8 -*-
"""
declarar_tablas_canonicas - resuelve el doble prefijo de
`resources/ASME B31/ASME B31.3/CHAPTERS/tables/`.

EL PROBLEMA
-----------
En esa carpeta hay 66 archivos que son en realidad 32 PARES: la extraccion
aplico el prefijo `table_` dos veces y dejo `table_X.json` junto a
`table_table_X.json`. Un motor que cite "la Tabla 302.3.3-1 de resources/" sin
decir cual de los dos archivos no es auditable, que es justo lo que la regla nro 1
del proyecto exige.

Y no son duplicados sin mas: en 11 de los 32 el contenido DIFIERE. Elegir "el
primero que salga" seria elegir a ciegas.

EL CRITERIO, QUE ES MECANICO Y NO DE JUICIO
-------------------------------------------
El archivo de prefijo SIMPLE es la extraccion recompuesta y el de prefijo doble
conserva los fragmentos crudos del corte de linea. Se comprueba par a par y solo
se declara si el par cae en una de estas cinco clases:

  IDENTICO                       los dos archivos son el mismo contenido.
                                 Descartar uno no pierde nada.            (21)
  NOTAS_MAS_COMPLETAS            las filas coinciden y las notas del gemelo son
                                 un subconjunto ESTRICTO de las del simple. (3)
  FILAS_RECOMPUESTAS             el texto de cada COLUMNA es identico caracter a
                                 caracter; el simple trae menos filas porque
                                 reintegro las que el corte de linea habia
                                 partido.                                  (6)
  CONTINUACION_NORMALIZADA       igual, y ademas el simple quita el marcador
                                 "(Cont'd)" que el codigo reimprime al pasar de
                                 folio. Es artefacto de pagina.            (1)
  ETIQUETA_DE_GRUPO_NO_REPETIDA  igual, y ademas al fusionar una fila partida la
                                 etiqueta de grupo que esa fila repetia pasa de
                                 N a N-1 repeticiones.                     (1)

Cualquier par que no encaje en ninguna ABORTA el script. No se declara nada a
medias: o el criterio lo resuelve, o se dice que no lo resuelve.

POR QUE SE COMPARA POR COLUMNA Y NO EN LECTURA PLANA
-----------------------------------------------------
La recomposicion no reordena el texto al azar: fusiona, DENTRO DE UNA COLUMNA,
los fragmentos que el corte de linea habia repartido entre filas consecutivas.
En lectura plana esos fragmentos no son contiguos -entre "Steel" y "T <= 25 mm"
va toda la fila de la norma y el nivel de aceptacion-, asi que comparar la
concatenacion plana daria una diferencia que no existe. Por columna la igualdad
es exacta, y es la misma comprobacion sin perdida que sostiene secii_tablas.py:
demuestra que descartar el gemelo no pierde ni un dato.

QUE ESCRIBE
-----------
Solo metadatos. Ningun valor ni nombre de fila se altera.
  - en el archivo canonico: `canonical_declaration`, con la clase, el motivo y
    el SHA-256 del descartado;
  - en el gemelo: `superseded_by`, para que quien lo abra sepa que no es la
    fuente que hay que citar.

Donde `completar_tabla_302_3_3.py` ya dejo su propia declaracion, este script la
LEE y comprueba que dice lo mismo; si no coincidiera, aborta. Dos mecanismos que
afirman el mismo hecho tienen que afirmarlo igual.

USO
---
    python declarar_tablas_canonicas.py --resources ..\\..\\..\\resources [--dry-run]

Es idempotente: ignora lo que el propio script escribe al comparar.
"""
from __future__ import annotations

import argparse
import copy
import datetime
import hashlib
import json
import re
import sys
from collections import Counter
from pathlib import Path

TABLAS = Path("ASME B31") / "ASME B31.3" / "CHAPTERS" / "tables"

# Claves que escribe esta herramienta (y su hermana de la 302.3.3-1). Se ignoran
# al comparar para que una segunda corrida no crea que el contenido cambio.
CLAVES_PROPIAS = ("canonical_declaration", "superseded_by", "extraction_amendments",
                  "extraction_gap")
CLAVES_COL_PROPIAS = ("header_derivado", "fuente_derivacion")

IDENTICO = "IDENTICO"
NOTAS = "NOTAS_MAS_COMPLETAS"
RECOMPUESTAS = "FILAS_RECOMPUESTAS"
CONTINUACION = "CONTINUACION_NORMALIZADA"
ETIQUETA = "ETIQUETA_DE_GRUPO_NO_REPETIDA"
CLASES = (IDENTICO, NOTAS, RECOMPUESTAS, CONTINUACION, ETIQUETA)

MOTIVOS = {
    IDENTICO: "los dos archivos traen el mismo contenido; el prefijo 'table_' "
              "se aplico dos veces y descartar el gemelo no pierde nada",
    NOTAS: "las filas coinciden y las notas del gemelo son un subconjunto "
           "estricto de las del canonico",
    RECOMPUESTAS: "el texto de cada COLUMNA es identico caracter a caracter; el "
                  "canonico trae menos filas porque reintegro, dentro de su "
                  "columna, las que el corte de linea habia partido",
    CONTINUACION: "ademas de reintegrar las filas partidas, el canonico quita el "
                  "marcador (Cont'd) que el codigo reimprime al pasar de folio; "
                  "es artefacto de pagina, no contenido",
    ETIQUETA: "ademas de lo anterior, al fusionar una fila partida la etiqueta de "
              "grupo que esa fila repetia pasa de N a N-1 repeticiones. El texto "
              "de cada columna coincide colapsando las repeticiones consecutivas: "
              "no se pierde ningun dato, solo una copia de un rotulo que el codigo "
              "reimprime en cada fila",
}

# El codigo imprime el apostrofo tipografico (U+2019) en (Cont'd); la extraccion
# conserva a veces el recto. Se aceptan los dos.
_RE_CONTD = re.compile("\\s*\\(\\s*Cont[’']?\\s*d\\s*\\)", re.I)
_RE_WS = re.compile(r"\s+")


def _limpio(doc):
    d = copy.deepcopy(doc)
    for k in CLAVES_PROPIAS:
        d.pop(k, None)
    for col in d.get("columns") or []:
        if isinstance(col, dict):
            for k in CLAVES_COL_PROPIAS:
                col.pop(k, None)
    return d


def _celdas(doc):
    """Texto de todas las celdas, en orden de lectura."""
    out = []
    for row in doc.get("rows") or []:
        if not isinstance(row, dict):
            continue
        for v in row.values():
            if v is not None:
                t = _RE_WS.sub(" ", str(v)).strip()
                if t:
                    out.append(t)
    return out


def _por_columna(doc, sin_contd=False, colapsar=False):
    """Texto de cada COLUMNA, concatenado en orden de fila y sin espacios."""
    cols, ultimo = {}, {}
    for row in doc.get("rows") or []:
        if not isinstance(row, dict):
            continue
        for k, v in row.items():
            if v is None:
                continue
            t = str(v)
            if sin_contd:
                t = _RE_CONTD.sub("", t)
            t = _RE_WS.sub("", t)
            # `colapsar` funde las repeticiones CONSECUTIVAS del mismo valor.
            # Solo hace falta para las etiquetas de grupo, que se reimprimen en
            # cada fila: al fusionar una fila partida, la etiqueta pasa de N a
            # N-1 repeticiones sin que se pierda ningun dato.
            if colapsar and ultimo.get(k) == t:
                continue
            ultimo[k] = t
            cols[k] = cols.get(k, "") + t
    return cols


def _notas(doc):
    return {_RE_WS.sub("", str(n)) for n in (doc.get("notes") or [])}


def clasificar(a, b):
    """Clase del par (canonico = a, de prefijo simple; gemelo = b), o None.

    Se prueban de la mas fuerte a la mas debil: cada clase posterior tolera un
    artefacto mas, y solo se llega a ella si la anterior no basta. Asi los pares
    que se demuestran con la prueba fuerte conservan esa prueba y no quedan
    cubiertos por una tolerancia que no necesitan.
    """
    la, lb = _limpio(a), _limpio(b)
    if la == lb:
        return IDENTICO
    if Counter(_celdas(la)) == Counter(_celdas(lb)) and _notas(lb) < _notas(la):
        return NOTAS
    if _por_columna(la) == _por_columna(lb):
        return RECOMPUESTAS
    if _por_columna(la, True) == _por_columna(lb, True):
        return CONTINUACION
    if (_por_columna(la, True, True) == _por_columna(lb, True, True)
            and len(la.get("rows") or []) < len(lb.get("rows") or [])):
        return ETIQUETA
    return None


def sha256(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def pares_de(base: Path):
    nombres = {p.name for p in base.glob("*.json")}
    out = []
    for n in sorted(nombres):
        if n.startswith("table_table_"):
            simple = "table_" + n[len("table_table_"):]
            if simple in nombres:
                out.append((simple, n))
    return out


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

    pares = pares_de(base)
    if not pares:
        print("ERROR: no se encuentra ningun par de doble prefijo", file=sys.stderr)
        return 2

    decididos, sin_criterio = [], []
    for simple, doble in pares:
        da = json.loads((base / simple).read_text(encoding="utf-8"))
        db = json.loads((base / doble).read_text(encoding="utf-8"))
        clase = clasificar(da, db)
        # Donde completar_tabla_302_3_3.py ya declaro, se comprueba la coherencia.
        previo = (da.get("extraction_amendments") or {}).get("archivo_canonico")
        if previo and previo != simple:
            print(f"ERROR: {simple} ya declara como canonico a {previo!r}, que no "
                  f"es el mismo archivo. Dos mecanismos no pueden discrepar.",
                  file=sys.stderr)
            return 2
        (decididos if clase else sin_criterio).append((simple, doble, clase, da, db))

    print("=" * 78)
    print(f"Doble prefijo en CHAPTERS/tables: {len(pares)} pares")
    print("=" * 78)
    for clase in CLASES:
        del_clase = [p for p in decididos if p[2] == clase]
        print(f"  {clase:31s} {len(del_clase):3d}")
        for simple, _, _, _, _ in del_clase:
            print(f"      {simple}")
    if sin_criterio:
        print()
        print("  SIN CRITERIO - no se declara ninguno de estos pares:")
        for simple, doble, _, _, _ in sin_criterio:
            print(f"      {simple}  vs  {doble}")
        print("\nERROR: el criterio mecanico no resuelve todos los pares. No se "
              "escribe nada: o se resuelven todos, o se dice que no se resuelven.",
              file=sys.stderr)
        return 2

    if a.dry_run:
        print("\n(--dry-run: no se ha escrito nada)")
        return 0

    hoy = datetime.date.today().isoformat()
    for simple, doble, clase, da, db in decididos:
        ruta_s, ruta_d = base / simple, base / doble
        da["canonical_declaration"] = {
            "fecha": hoy,
            "script": "declarar_tablas_canonicas.py",
            "archivo_canonico": simple,
            "clase": clase,
            "motivo": MOTIVOS[clase],
            "alcance": "Solo metadatos. Ningun valor ni nombre de fila se altera.",
            "duplicado_descartado": {"archivo": doble, "sha256": sha256(ruta_d)},
        }
        db["superseded_by"] = {
            "fecha": hoy,
            "script": "declarar_tablas_canonicas.py",
            "archivo": simple,
            "clase": clase,
            "motivo": "duplicado por doble prefijo de la extraccion. "
                      + MOTIVOS[clase]
                      + ". Cite siempre el canonico; este archivo se conserva "
                        "como testigo.",
        }
        ruta_s.write_text(json.dumps(da, ensure_ascii=False, indent=1),
                          encoding="utf-8")
        ruta_d.write_text(json.dumps(db, ensure_ascii=False, indent=1),
                          encoding="utf-8")
    print(f"\nDeclarados {len(decididos)} pares. Canonico = prefijo simple en todos.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
