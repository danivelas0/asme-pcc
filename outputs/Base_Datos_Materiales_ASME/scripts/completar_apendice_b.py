# -*- coding: utf-8 -*-
"""
completar_apendice_b - repara la capa de IDENTIFICACION y ESTRUCTURA del
Apendice B de ASME B31.3 (Tablas B-1, B-1C, B-2..B-6 y el indice de
especificaciones) en `resources/`.

POR QUE EXISTE ESTE SCRIPT
--------------------------
Los VALORES del Apendice B estaban todos bien: se cotejaron fila a fila las 7
tablas y las 28 entradas del indice contra los folios impresos 399-405 y no hay
un solo numero mal. Lo que falla es como quedaron identificadas y estructuradas
esas filas, y hay un caso que puede hacer dano de verdad:

  B-5 declara las columnas `°C | °F | Maximum °F` donde el codigo imprime
  `Minimo °C | Minimo °F | Maximo °C | Maximo °F`. La cabecera se desalineo y
  el maximo de 232 C quedo guardado en un campo llamado `c`. Quien lo lea como
  "temperatura minima" se lleva 232 C de minimo en un tubo de vidrio
  borosilicato.

Los otros cinco defectos son de identificacion: celdas fusionadas en B-1, tres
columnas de B-3 colapsadas en una sola celda, y el nombre de material a dos
lineas de la F1974 de B-6 partido entre dos filas.

FUENTE
------
`resources/.../appendix_b/fuente/APENDICE_B_ASME_B31_3_2024.pdf`, folios
impresos 399-405. Es un escaneo SIN capa de texto (pdfplumber devuelve cero
lineas en las 8 paginas) y no hay OCR estructurado, asi que las correcciones van
aqui como TABLA LITERAL, cada una citando su folio, verificadas ampliando la
pagina. Mismo estatus que `decisiones_map_grupo.json`: dato introducido por una
persona, declarado como tal y auditable.

TRES RAREZAS DEL PROPIO CODIGO, QUE NO SE TOCAN
-----------------------------------------------
Se conservan tal como estan impresas (Regla 9) y se dejan declaradas en
`observaciones_del_codigo` de cada archivo:
1. B-1 imprime "..." como temperatura minima de D2846/CPVC4120; B-1C imprime
   73 F para esa misma fila.
2. F2389: B-1 da 110 C de maximo y B-1C da 210 F (= 98,9 C), que no se
   corresponden; y la designacion de tuberia es "PR" en SI y "IPS Sch. 80" en US.
3. B-6, fila F1282 a 862 kPa: imprime 100 psi donde sus filas gemelas imprimen
   125 psi (862 kPa = 125 psi).

USO
---
    python completar_apendice_b.py --resources ..\\..\\..\\resources [--dry-run]

Es idempotente: comprueba el estado de partida y se puede volver a correr.
"""

import argparse
import datetime
import json
import re
import sys
from pathlib import Path

APXB = Path("asme_b31") / "asme_b31_3" / "appex" / "appendix_b"

FOLIOS = {
    "table_b_1.json": [400, 401], "table_b_1c.json": [402, 403],
    "table_b_2.json": [403], "table_b_3.json": [403],
    "table_b_4.json": [404], "table_b_5.json": [404],
    "table_b_6.json": [405], "spec_index_b.json": [399],
}

# ---------------------------------------------------------------------------
# 1. Celdas fusionadas de B-1
# ---------------------------------------------------------------------------
# El impreso deja la columna "ASTM Spec. No." con la elipsis del codigo en la
# fila del ABS, y la extraccion la pego a la designacion de tuberia de la
# columna siguiente. Lo mismo con la F2389. B-1C separo bien las dos columnas en
# esas mismas filas, lo que confirma que el defecto es exclusivo de B-1.
FUSION_B1 = {
    # material_designation -> (astm_spec_no correcto, pipe_designation correcta)
    "ABS": (None, "PR"),
    "PP-R": ("F2389", "PR"),
    "PP-RCT": ("F2389", "PR"),
}

# ---------------------------------------------------------------------------
# 2. Especificacion partida por el salto de linea (B-1 y B-1C)
# ---------------------------------------------------------------------------
# El impreso parte "F2788/F2788M" en dos lineas y la extraccion dejo el salto
# como un espacio. El indice de especificaciones la escribe sin espacio, asi que
# tal como esta ninguna de las dos tablas empareja con el indice.
SPEC_PARTIDA = {"F2788/ F2788M": "F2788/F2788M"}

# ---------------------------------------------------------------------------
# 3. B-3: tres columnas colapsadas en una celda
# ---------------------------------------------------------------------------
# El impreso es una rejilla de 3 columnas x 2 filas con SEIS especificaciones.
# Se leen POR COLUMNAS, que es como quedan en orden ascendente y como las
# enumera el indice:  D2517 D2996 | D2997 D3517 | D3754 AWWA C950.
B3_ESPERADO = ["D2517 D2997 D3754", "D2996 D3517 AWWA C950"]
B3_CORRECTO = ["D2517", "D2996", "D2997", "D3517", "D3754", "AWWA C950"]

# ---------------------------------------------------------------------------
# 4 y 5. Columnas de temperatura de B-4 y B-5
# ---------------------------------------------------------------------------
# Las dos tablas imprimen "Recommended Temperature Limits [Note (1)]" repartido
# en Minimo y Maximo, cada uno con su °C y su °F: CUATRO columnas. La extraccion
# guardo dos en B-4 y tres desalineadas en B-5.
B4_COLS = ["Spec. No.", "Material", "Class", "kPa", "psi",
           "Minimum °C", "Minimum °F", "Maximum °C", "Maximum °F"]
B4_GRUPOS = ["Allowable Gage Pressure", "Recommended Temperature Limits [Note (1)]",
             "Minimum", "Maximum"]
B5_COLS = ["ASTM Spec. No.", "Material", "Size Range DN", "NPS", "kPa",
           "Allowable psi", "Minimum °C", "Minimum °F", "Maximum °C", "Maximum °F"]
B5_GRUPOS = ["Size Range", "Allowable Gage Pressure",
             "Recommended Temperature Limits [Note (1)]", "Minimum", "Maximum"]

# ---------------------------------------------------------------------------
# 6. B-6: nombre de material a dos lineas partido entre filas
# ---------------------------------------------------------------------------
# El impreso da a la F1974 dos materiales: "Metal insert fittings for
# PEX-AL-PEX systems" (una fila) y "Metal insert fittings for PE-AL-PE systems"
# (dos filas). La extraccion corto el segundo por el salto de linea y ascendio
# su segunda mitad a material de la fila siguiente.
B6_MATERIAL_PARTIDO = ("Metal insert fittings for", "PE-AL-PE systems",
                       "Metal insert fittings for PE-AL-PE systems")

# ---------------------------------------------------------------------------
# 7. Notas del indice de especificaciones (folio 399)
# ---------------------------------------------------------------------------
# La Nota (2) es normativa: fija la equivalencia de designacion que usan B-2 y
# B-3. No estaba en el JSON.
INDICE_NOTAS = [
    "GENERAL NOTE: It is not practical to refer to a specific edition of each "
    "standard throughout the Code text. Instead, the approved edition references, "
    "along with the names and websites of the sponsoring organizations, are shown "
    "in Appendix E.",
    "NOTES: (1) For names of plastics identified only by abbreviation, see para. "
    "A326.4. (2) The term fiberglass RTR takes the place of the ASTM designation "
    "“fiberglass” (glass-fiber-reinforced thermosetting resin).",
]

OBSERVACIONES = {
    "table_b_1.json": [
        "B-1 imprime \"...\" (sin valor) como temperatura minima de D2846 / CPVC4120, "
        "mientras que B-1C imprime 73 F para esa misma fila. Asimetria del codigo; se "
        "conserva lo impreso en cada edicion (folios 400 y 402).",
        "F2389 (PP-R y PP-RCT): B-1 imprime 110 C de temperatura maxima y B-1C imprime "
        "210 F, que equivale a 98,9 C. Las dos ediciones no se corresponden. Ademas la "
        "designacion de tuberia es \"PR\" en B-1 e \"IPS Sch. 80\" en B-1C. Se conservan "
        "las dos tal como estan impresas (folios 400 y 402).",
    ],
    "table_b_1c.json": [
        "F2389 (PP-R y PP-RCT): B-1C imprime 210 F de maximo (= 98,9 C) frente a los "
        "110 C de B-1, y designa la tuberia \"IPS Sch. 80\" donde B-1 dice \"PR\". Se "
        "conservan las dos tal como estan impresas (folios 400 y 402).",
    ],
    "table_b_6.json": [
        "F1282 a 862 kPa: el codigo imprime 100 psi donde sus filas gemelas de F1281 y "
        "F1974 imprimen 125 psi para la misma presion (862 kPa = 125 psi). Se conserva "
        "el 100 tal como esta impreso (folio 405).",
    ],
}


def cargar(ruta):
    return json.loads(ruta.read_text(encoding="utf-8"))


def guardar(ruta, doc):
    ruta.write_text(json.dumps(doc, ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8")


def sellar(doc, archivo, cambios):
    doc["extraction_amendments"] = {
        "fecha": datetime.date.today().isoformat(),
        "script": "completar_apendice_b.py",
        "fuente": "appendix_b/fuente/APENDICE_B_ASME_B31_3_2024.pdf (escaneo, sin capa "
                  "de texto; correcciones transcritas y verificadas a ojo sobre la "
                  "pagina ampliada)",
        "folios_impresos": FOLIOS[archivo],
        "alcance": "Solo identificacion y estructura: celdas fusionadas, columnas "
                   "colapsadas, cabeceras desalineadas y notas. Los valores NO se "
                   "tocan: se cotejaron fila a fila contra el impreso y son correctos.",
        "cambios": cambios,
    }
    if archivo in OBSERVACIONES:
        doc["observaciones_del_codigo"] = OBSERVACIONES[archivo]


# ---------------------------------------------------------------------------
def arreglar_b1(doc, es_si):
    cambios = []
    for i, r in enumerate(doc["rows"]):
        # Celdas fusionadas: solo B-1 las trae; B-1C ya separa bien.
        if es_si:
            objetivo = FUSION_B1.get(r.get("material_designation"))
            if objetivo and r.get("pipe_designation") is None:
                spec, pipe = objetivo
                cambios.append("fila %d (%s): astm_spec_no %r -> %r, "
                               "pipe_designation None -> %r"
                               % (i, r["material_designation"], r["astm_spec_no"],
                                  spec, pipe))
                r["astm_spec_no"], r["pipe_designation"] = spec, pipe
        # Especificacion partida por el salto de linea: en las dos ediciones.
        if r.get("astm_spec_no") in SPEC_PARTIDA:
            nuevo = SPEC_PARTIDA[r["astm_spec_no"]]
            cambios.append("fila %d: astm_spec_no %r -> %r"
                           % (i, r["astm_spec_no"], nuevo))
            r["astm_spec_no"] = nuevo
    return cambios


def arreglar_b3(doc):
    actuales = [r.get("spec_nos_astm_except_as_noted") for r in doc["rows"]]
    if actuales == B3_CORRECTO:
        return []
    if actuales != B3_ESPERADO:
        raise SystemExit("B-3: estado de partida inesperado: %r" % actuales)
    doc["rows"] = [{"spec_nos_astm_except_as_noted": s} for s in B3_CORRECTO]
    doc["row_count"] = len(B3_CORRECTO)
    return ["las 3 columnas impresas se leen por columnas: 2 celdas con tres "
            "especificaciones cada una -> %d filas, una por especificacion (%s)"
            % (len(B3_CORRECTO), ", ".join(B3_CORRECTO))]


def arreglar_b4(doc):
    if doc["columns"] == B4_COLS:
        return []
    cambios = ["columnas de temperatura: ['°C','°F'] -> "
               "['Minimum °C','Minimum °F','Maximum °C','Maximum °F'] "
               "(el impreso reparte Recommended Temperature Limits en minimo y "
               "maximo, cada uno con sus dos unidades)"]
    for r in doc["rows"]:
        # Los cuatro valores estan impresos como elipsis en las 9 filas: no se
        # pierde ningun numero, pero el esquema pasa a representar la tabla.
        for viejo in ("c", "f"):
            if r.pop(viejo, None) is not None:
                raise SystemExit("B-4: se esperaba que %r fuese nulo" % viejo)
        r["minimum_c"] = r["minimum_f"] = r["maximum_c"] = r["maximum_f"] = None
    doc["columns"] = B4_COLS
    doc["column_groups"] = B4_GRUPOS
    doc["value_axis"] = "Recommended Temperature Limits [Note (1)]"
    return cambios


def arreglar_b5(doc):
    if doc["columns"] == B5_COLS:
        return []
    cambios = ["cabecera desalineada: ['°C','°F','Maximum °F'] -> "
               "['Minimum °C','Minimum °F','Maximum °C','Maximum °F']. El campo "
               "`c` guardaba el MAXIMO (232 C), no el minimo"]
    for i, r in enumerate(doc["rows"]):
        maxc, maxf = r.pop("c", None), r.pop("maximum_f", None)
        if r.pop("f", None) is not None:
            raise SystemExit("B-5 fila %d: se esperaba `f` nulo" % i)
        if (maxc, maxf) != (232, 450):
            raise SystemExit("B-5 fila %d: se esperaba (232, 450), hay %r"
                             % (i, (maxc, maxf)))
        r["minimum_c"] = r["minimum_f"] = None
        r["maximum_c"], r["maximum_f"] = maxc, maxf
    doc["columns"] = B5_COLS
    doc["column_groups"] = B5_GRUPOS
    doc["value_axis"] = "Recommended Temperature Limits [Note (1)]"
    return cambios


def arreglar_b6(doc):
    cambios = []
    trunco, cola, entero = B6_MATERIAL_PARTIDO
    for i, r in enumerate(doc["rows"]):
        if r.get("material") == trunco:
            r["material"] = entero
            cambios.append("fila %d: material %r -> %r (el impreso lo escribe a dos "
                           "lineas)" % (i, trunco, entero))
        elif r.get("material") == cola:
            r["material"] = None
            cambios.append("fila %d: material %r -> None (era la segunda linea del "
                           "nombre de la fila anterior, no un material propio)"
                           % (i, cola))
        # La tabla solo publica limites MAXIMOS: `f` es el maximo en °F.
        if "f" in r:
            r["maximum_f"] = r.pop("f")
    if any("maximum_f" in r for r in doc["rows"]) and "°F" in doc["columns"]:
        doc["columns"] = [c if c != "°F" else "Maximum [Note (1)] °F"
                          for c in doc["columns"]]
        cambios.append("columna '°F' renombrada a 'Maximum [Note (1)] °F' y campo "
                       "`f` -> `maximum_f`: la tabla solo publica limites maximos")
    return cambios


def arreglar_indice(doc):
    if doc.get("notes") == INDICE_NOTAS:
        return []
    doc["notes"] = INDICE_NOTAS
    return ["notas del folio 399 anadidas: GENERAL NOTE y Notas (1) y (2). La (2) "
            "es normativa: fija que 'fiberglass RTR' sustituye a la designacion "
            "ASTM 'fiberglass'"]


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args()

    base = a.resources / APXB
    if not base.is_dir():
        print("ERROR: no existe %s" % base, file=sys.stderr)
        return 2

    trabajo = [
        ("table_b_1.json", lambda d: arreglar_b1(d, es_si=True)),
        ("table_b_1c.json", lambda d: arreglar_b1(d, es_si=False)),
        ("table_b_3.json", arreglar_b3),
        ("table_b_4.json", arreglar_b4),
        ("table_b_5.json", arreglar_b5),
        ("table_b_6.json", arreglar_b6),
        ("spec_index_b.json", arreglar_indice),
    ]

    total = 0
    for archivo, fn in trabajo:
        ruta = base / archivo
        doc = cargar(ruta)
        cambios = fn(doc)
        print("=" * 74)
        print("%s  folios %s" % (archivo, FOLIOS[archivo]))
        print("=" * 74)
        if cambios:
            for c in cambios:
                print("  -", c)
            total += len(cambios)
        else:
            print("  ninguna (ya estaba corregido)")
        sellar(doc, archivo, cambios)
        if not a.dry_run:
            guardar(ruta, doc)
        print()

    print("%d correcciones%s." % (total, " (dry-run, no se escribio nada)"
                                  if a.dry_run else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
