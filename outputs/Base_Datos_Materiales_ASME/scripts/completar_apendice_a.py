# -*- coding: utf-8 -*-
"""
completar_apendice_a - repara los desplazamientos de columna del Apendice A de
ASME B31.3 (Tablas A-1, A-1C, A-3, A-4 y A-4C) en `resources/`.

POR QUE EXISTE ESTE SCRIPT
--------------------------
El Apendice A se imprime a PAGINAS ENFRENTADAS: la izquierda lleva la
identificacion del material y la derecha el numero de linea con la rejilla de
esfuerzos admisibles. La extraccion tiene que fusionar las dos por `Line No.`, y
en esa costura se colaron tres defectos, uno de ellos con perdida de dato:

1. A-1C PIERDE 197 VALORES DE ESFUERZO A 200 F. En el bloque de aleaciones de
   niquel (folios 334-347) la pagina izquierda termina con las columnas
   "Min. Temp. to 100 | 200 | 300". El extractor pego la palabra "Metal" del
   titulo de banda -"Basic Allowable Stress, S, ksi, at Metal Temperature, F"-
   al encabezado "200", invento el campo de identificacion `metal_200` y metio
   ahi el esfuerzo. Resultado: la curva de esas 197 filas empieza en 300 F y el
   punto de 200 F, que SI esta impreso, no se puede consultar.

2. A-3 FUSIONA "Class (or Type)" CON "Description" EN 38 FILAS. Cuando la
   descripcion arranca en la misma linea visual que la clase, las dos celdas
   quedaron pegadas: "… Seamless pipe", "Type S Seamless pipe", "All Welded
   pipe". La descripcion quedo vacia.

3. LA ELIPSIS DE UNA COLUMNA VACIA CONTIGUA SE PEGA AL CAMPO ANTERIOR. En el
   codigo, "…" significa "sin valor". Al fusionar quedo dentro del campo vecino:
   `spec_no = "A179 …"`, `material = "A1011 Gr. 30 …"`, `p_no = "1 …"`.

FUENTE
------
`resources/.../appendix_a/fuente/APENDICE_A_ASME_B31_3_2024.pdf`, folios
impresos 166-398 (233 paginas). Escaneo SIN capa de texto y sin OCR
estructurado: las decisiones que no son mecanicas van aqui como tabla literal
citando su folio, verificadas ampliando la pagina.

LO QUE ESTE SCRIPT NO TOCA
--------------------------
- Ningun valor tabulado se altera: el de 200 F se MUEVE de un campo de
  identificacion a su punto de la curva, que es donde lo imprime el codigo.
- Los nombres de columna que el propio codigo escribe de dos maneras
  ("Min. Tensile Strength, ksi" en unos bloques y "Minimum Tensile Strength,
  ksi" en otros, "Class/ Condition/ Temper" y "Class/ Condi- tion/ Temper")
  se CONSERVAN como claves distintas: es lo que esta impreso, y el builder ya
  las resuelve con sus alias en `g(row, ...)`. Normalizarlas seria reescribir
  el codigo, no corregir la extraccion.

USO
---
    python completar_apendice_a.py --resources ..\\..\\..\\resources [--dry-run]

Es idempotente: comprueba el estado de partida y se puede volver a correr.
"""

import argparse
import datetime
import json
import re
import sys
from pathlib import Path

APXA = Path("asme_b31") / "asme_b31_3" / "appex" / "appendix_a"

FOLIOS = {
    "table_a_1.json": [174, 283], "table_a_1c.json": [284, 363],
    "table_a_3.json": [366, 368], "table_a_4.json": [372, 387],
    "table_a_4c.json": [388, 397],
}

# ---------------------------------------------------------------------------
# 1. A-1C: el esfuerzo a 200 F que quedo como campo de identificacion
# ---------------------------------------------------------------------------
CAMPO_200 = "metal_200"
COL_200 = "200"
COL_BASURA = "Metal 200"

# ---------------------------------------------------------------------------
# 2. A-3: las tres filas cuya clase NO es la elipsis
# ---------------------------------------------------------------------------
# Verificadas ampliando el folio que se cita. El resto de las 38 empieza por la
# elipsis del codigo y se resuelve solo.
A3_LITERAL = {
    # valor fusionado -> (Class (or Type), Description, folio)
    "Type S Seamless pipe": ("Type S", "Seamless pipe", 366),
    "All Welded pipe": ("All", "Welded pipe", 368),
    # API 5L, tercera fila del bloque: la clase va en blanco (continuacion) y
    # todo el texto es la descripcion.
    "Electric welded pipe": (None, "Electric welded pipe", 366),
}

ELIPSIS = "…"
# Solo se quita la elipsis cuando es un TOKEN SUELTO al principio o al final.
# Nunca se toca una elipsis interior: ahi no significa columna vacia.
RE_ELI_FIN = re.compile(r"\s+" + ELIPSIS + r"\s*$")
RE_ELI_INI = re.compile(r"^\s*" + ELIPSIS + r"\s+")
# Nombre partido por el salto de linea del impreso.
BARRA_PARTIDA = re.compile(r"/\s+")

CAMPOS_IDENT = ("material", "nominal_composition", "product_form", "spec_no",
                "type_grade", "uns_no", "class_condition_temper",
                "class_condi_tion_temper", "class_cond_temper", "size_in",
                "size_mm", "size_range_in", "size_range_dia_mm",
                "size_or_thickness_range_in", "p_no", "notes")


def limpiar_valor(v):
    """Quita la elipsis suelta y el espacio del salto de linea. Devuelve
    (valor_nuevo, cambio_aplicado)."""
    if not isinstance(v, str):
        return v, None
    orig = v
    v = RE_ELI_FIN.sub("", v)
    v = RE_ELI_INI.sub("", v)
    v = v.strip()
    if v == ELIPSIS or v == "":
        return None, "elipsis sola -> None" if orig.strip() else None
    if BARRA_PARTIDA.search(v):
        v = BARRA_PARTIDA.sub("/", v)
    if v == orig:
        return orig, None
    # "1 …" -> 1 : si lo que queda es un entero puro, se devuelve como numero,
    # que es como esta cargado en el resto de las filas.
    if re.fullmatch(r"-?\d+", v):
        return int(v), "%r -> %d" % (orig, int(v))
    return v, "%r -> %r" % (orig, v)


# ---------------------------------------------------------------------------
def arreglar_elipsis(doc):
    cambios, n = [], 0
    for i, r in enumerate(doc["rows"]):
        for k in CAMPOS_IDENT:
            if k not in r:
                continue
            nuevo, msg = limpiar_valor(r[k])
            if msg:
                r[k] = nuevo
                n += 1
                if len(cambios) < 6:
                    cambios.append("fila %d  %s: %s" % (i, k, msg))
    if n:
        cambios.insert(0, "elipsis de columna vacia contigua y saltos de linea "
                          "limpiados en %d celdas de identificacion" % n)
    return cambios


def arreglar_titulo_absorbido(doc):
    """Un titulo de banda de seccion que acabo dentro de `spec_no`.

    Caso unico y verificado: A-1C bloque 3, linea 86, UNS K81340 lleva
    'A353 … Low and Intermediate Alloy Steel - Forgings and Fittings'. La
    elipsis marca el Type/Grade vacio y lo que sigue es el encabezado de la
    seccion SIGUIENTE -las filas 87 en adelante son forjados y accesorios-.
    A-1 tiene esa misma linea como spec 'A353', forma 'Plate'. Una
    especificacion no contiene nunca una elipsis, asi que el corte es seguro.
    """
    cambios = []
    for i, r in enumerate(doc["rows"]):
        v = r.get("spec_no")
        if isinstance(v, str) and ELIPSIS in v:
            nuevo = v.split(ELIPSIS)[0].strip()
            cambios.append("fila %d  spec_no: %r -> %r (titulo de seccion "
                           "absorbido)" % (i, v, nuevo))
            r["spec_no"] = nuevo
    if cambios:
        cambios.insert(0, "titulo de banda de seccion retirado de spec_no en "
                          "%d fila(s)" % len(cambios))
    return cambios


def arreglar_spec_grado_pegados(doc, ref):
    """Separa Spec. No. y Type/Grade cuando quedaron pegados en A-1C.

    Solo se toca una fila cuando el valor coincide EXACTAMENTE con
    `spec + ' ' + grado` de la misma (bloque, linea) en A-1: eso demuestra por
    si solo que son dos celdas pegadas. Las dos ediciones NO siempre nombran
    igual el grado -la SI imprime API 5L L245 donde la US imprime API 5L B-,
    asi que A-1 no vale como referencia general: solo como prueba del pegado.
    """
    idx = {}
    for r in ref["rows"]:
        idx.setdefault((r.get("block"), r.get("line_no")), []).append(r)

    cambios, n = [], 0
    for i, r in enumerate(doc["rows"]):
        v = r.get("spec_no")
        if not isinstance(v, str) or " " not in v.strip():
            continue
        if r.get("type_grade") not in (None, ""):
            continue
        cands = idx.get((r.get("block"), r.get("line_no")))
        if not cands or len(cands) != 1:
            continue
        sp, gr = cands[0].get("spec_no"), cands[0].get("type_grade")
        if not sp or gr in (None, ""):
            continue
        if v.strip() != ("%s %s" % (sp, gr)).strip():
            continue
        r["spec_no"], r["type_grade"] = sp, gr
        n += 1
        if len(cambios) < 6:
            cambios.append("fila %-5d %r -> spec=%r  grado=%r" % (i, v, sp, gr))
    if n:
        cambios.insert(0, "Spec. No. y Type/Grade separados en %d filas "
                          "(coincidencia exacta con A-1 en la misma linea)" % n)
    return cambios


def arreglar_metal_200(doc):
    """Devuelve el esfuerzo de 200 F a su punto de la curva."""
    afectadas = [r for r in doc["rows"] if r.get(CAMPO_200) is not None]
    if not afectadas and CAMPO_200 not in {k for r in doc["rows"] for k in r}:
        return []
    movidas = 0
    for r in doc["rows"]:
        if CAMPO_200 not in r:
            continue
        v = r.pop(CAMPO_200)
        if v is None:
            continue
        vals = r.setdefault("values", {})
        if COL_200 in vals and vals[COL_200] is not None:
            raise SystemExit("A-1C: la fila ya tiene un valor a 200 F; no se "
                             "sobreescribe (%r)" % r.get("spec_no"))
        # Se reinserta respetando el orden de temperaturas del codigo.
        nuevo = {COL_200: v}
        nuevo.update({k: x for k, x in vals.items() if k != COL_200})
        r["values"] = nuevo
        movidas += 1
    ci = doc.get("columns_identification") or []
    if COL_BASURA in ci:
        doc["columns_identification"] = [c for c in ci if c != COL_BASURA]
    if not movidas:
        return []
    return ["%d esfuerzos a 200 F devueltos de la identificacion `%s` a "
            "values['200'] (bloque de niquel, folios 334-347)" % (movidas, CAMPO_200)]


def arreglar_a3(doc):
    cambios, n = [], 0
    for i, r in enumerate(doc["rows"]):
        if r.get("description") not in (None, ""):
            continue
        c = r.get("class_or_type")
        if not isinstance(c, str):
            continue
        if c.lstrip().startswith(ELIPSIS):
            clase, desc = None, c.lstrip()[1:].strip()
        elif c in A3_LITERAL:
            clase, desc, _ = A3_LITERAL[c]
        else:
            raise SystemExit("A-3 fila %d: fusion no prevista %r. Verifiquela "
                             "contra el folio antes de tocarla." % (i, c))
        r["class_or_type"], r["description"] = clase, desc
        n += 1
        if len(cambios) < 6:
            cambios.append("fila %-4d %r -> clase=%r  descripcion=%r"
                           % (i, c, clase, desc))
    if n:
        cambios.insert(0, "Class (or Type) y Description separadas en %d filas" % n)
    return cambios


# ---------------------------------------------------------------------------
# UNS con la Clase/Condicion/Temple pegada (A-1C, bloque 6)
# ---------------------------------------------------------------------------
# El folio 336 imprime "UNS No." y "Class/Condition/Temper" en columnas
# separadas -N08031 | Annealed-. En 186 filas del bloque de niquel quedaron
# pegadas, y `class_condition_temper` se quedo vacio en 208 de las 218 filas del
# bloque. A-1 no tiene el defecto. Un UNS es una letra y cinco digitos: todo lo
# que venga detras es de otra columna, asi que el corte es determinista.
RE_UNS = re.compile(r"^([A-Z]\d{5})\b\s*(.*)$")

# Caso aparte, folio 340: el nombre de la clase se imprime a dos lineas y la
# primera mitad se pego al UNS mientras la segunda quedo en su columna.
# Se recomponen uniendolas, no sobreescribiendo.
UNS_CLASE_A_DOS_LINEAS = {"N08810 Sol. tr. or", "N08811 Sol. tr. or"}

# Y un UNS con un caracter espurio delante, de una columna que no se puede
# atribuir: A536 Gr. 65-45-12 es F33100. Un UNS nunca empieza por ">".
UNS_CON_BASURA = {">F33100": "F33100"}


def arreglar_uns_con_clase(doc):
    cambios, n, unidos = [], 0, 0
    for i, r in enumerate(doc["rows"]):
        v = r.get("uns_no")
        if not isinstance(v, str) or not v.strip():
            continue
        v = v.strip()
        if v in UNS_CON_BASURA:
            cambios.append("fila %d  uns_no %r -> %r (caracter espurio de otra "
                           "columna; un UNS no empieza por un signo)"
                           % (i, v, UNS_CON_BASURA[v]))
            r["uns_no"] = UNS_CON_BASURA[v]
            continue
        m = RE_UNS.match(v)
        if not m or not m.group(2):
            continue
        uns, resto = m.group(1), m.group(2).strip()
        actual = r.get("class_condition_temper")
        if v in UNS_CLASE_A_DOS_LINEAS and actual:
            nueva = "%s %s" % (resto, actual)
            unidos += 1
        elif actual:
            # No se sobreescribe una clase existente sin saber por que.
            raise SystemExit("A-1C fila %d: el UNS trae %r pegado pero la fila ya "
                             "declara la clase %r. Verifiquelo contra el folio."
                             % (i, resto, actual))
        else:
            nueva = resto
        r["uns_no"], r["class_condition_temper"] = uns, nueva
        n += 1
        if len(cambios) < 4:
            cambios.append("fila %-5d %r -> uns=%r  clase=%r" % (i, v, uns, nueva))
    if n:
        cambios.insert(0, "UNS y Clase/Condicion/Temple separados en %d filas "
                          "(%d de ellas con el nombre de clase a dos lineas)"
                       % (n, unidos))
    return cambios


# ---------------------------------------------------------------------------
# Las 21 paginas derechas huerfanas, recuperadas del folio impreso
# ---------------------------------------------------------------------------
# En 21 lineas del bloque 6 de A-1C la fusion de paginas enfrentadas perdio la
# pagina IZQUIERDA entera: esas filas se quedaron sin identificacion y sin los
# tres primeros puntos de la curva. Se transcriben aqui de los folios 336, 338,
# 340, 342, 344 y 346, leidos ampliados, y contrastados uno a uno con la edicion
# SI (A-1) por composicion, spec, UNS, P-No. y resistencias convertidas.
#
# CUIDADO: la numeracion de linea NO esta alineada entre las dos ediciones -la
# linea 203 de A-1C no es la 203 de A-1-, asi que cada fila se leyo en SU pagina
# US, no por la posicion en A-1.
#
# Campos por linea:
#   composicion, spec, tipo/grado, UNS, clase, tamano, P-No, notas,
#   Tmin F, Rm ksi, Re ksi, Tmax F, S a 100 F, S a 200 F, S a 300 F
NI = "31Ni–33Fe–27Cr–6.5Mo–Cu–N"
FE = "46Fe–24Ni–21Cr–6Mo–Cu–N"
LA = "57Ni–22Cr–14W–2Mo–La"
GT = ">3∕16"
LE = "≤3∕16"

RECUPERADAS_A1C = {
    # linea: (composicion, spec, grado, uns, clase, tamano, p_no, notas,
    #         tmin, rm, re, tmax, s100, s200, s300, folio)
    53:  (NI, "B619", None, "N08031", "Annealed", None, 45, None, -325, 94, 40, 800, 26.7, 26.7, 26.7, 336),
    54:  (NI, "B622", None, "N08031", "Annealed", None, 45, None, -325, 94, 40, 800, 26.7, 26.7, 26.7, 336),
    61:  (FE, "B675", None, "N08367", "Annealed", GT, 45, None, -325, 95, 45, 900, 30.0, 30.0, 29.9, 336),
    62:  (FE, "B690", None, "N08367", "Annealed", GT, 45, None, -325, 95, 45, 900, 30.0, 30.0, 29.9, 336),
    63:  (FE, "B804", None, "N08367", "Annealed", GT, 45, None, -325, 95, 45, 900, 30.0, 30.0, 29.9, 336),
    64:  (FE, "B675", None, "N08367", "Annealed", LE, 45, None, -325, 100, 45, 900, 30.0, 30.0, 30.0, 336),
    65:  (FE, "B690", None, "N08367", "Annealed", LE, 45, None, -325, 100, 45, 900, 30.0, 30.0, 30.0, 336),
    66:  (FE, "B804", None, "N08367", "Annealed", LE, 45, None, -325, 100, 45, 800, 30.0, 30.0, 29.9, 336),
    90:  (LA, "B619", None, "N06230", "Sol. ann.", None, 43, None, -325, 110, 45, 1650, 30.0, 30.0, 30.0, 338),
    91:  (LA, "B622", None, "N06230", "Sol. ann.", None, 43, None, -325, 110, 45, 1650, 30.0, 30.0, 30.0, 338),
    92:  (LA, "B626", None, "N06230", "Sol. ann.", None, 43, None, -325, 110, 45, 1650, 30.0, 30.0, 30.0, 338),
    118: (NI, "B625", None, "N08031", "Annealed", "All", 45, None, -325, 94, 40, 800, 26.7, 26.7, 26.7, 340),
    122: (LA, "B435", None, "N06230", "Sol. ann.", "All", 43, None, -325, 110, 45, 1650, 30.0, 30.0, 30.0, 340),
    125: (FE, "B688", None, "N08367", "Annealed", GT, 45, None, -325, 95, 45, 900, 30.0, 30.0, 29.9, 340),
    126: (FE, "B688", None, "N08367", "Annealed", LE, 45, None, -325, 100, 45, 900, 30.0, 30.0, 30.0, 340),
    159: (NI, "B366", None, "N08031", "Sol. ann.", "All", 45, "(17)", -325, 94, 40, 800, 26.7, 26.7, 26.7, 342),
    184: (LA, "B564", None, "N06230", "Sol. ann.", "All", 43, None, -325, 110, 45, 1650, 30.0, 30.0, 30.0, 344),
    185: (LA, "B366", None, "N06230", "Sol. ann.", "All", 43, "(17)", -325, 110, 45, 1650, 30.0, 30.0, 30.0, 344),
    203: (NI, "B649", None, "N08031", "Annealed", "All", 45, None, -325, 94, 40, 800, 26.7, 26.7, 26.7, 344),
    211: (LA, "B572", None, "N06230", "Sol. ann.", "All", 43, None, -325, 110, 45, 1650, 30.0, 30.0, 30.0, 344),
    216: ("59Ni–22Cr–14Mo–4Fe–3W", "A494", "CX2MW", "N26022", None, None, 43,
          "(9)", -325, 80, 45, 500, 26.7, 26.7, 26.7, 346),
}

CAMPOS_REC = ("nominal_composition", "spec_no", "type_grade", "uns_no",
              "class_condition_temper", "size_range_in", "p_no", "notes",
              "min_temp_f", "minimum_tensile_strength_ksi",
              "minimum_yield_strength_ksi", "max_temp_f")


def recuperar_huerfanas(doc):
    """Rellena las 21 filas con lo leido del folio impreso."""
    cambios, n, ya = [], 0, 0
    for r in doc["rows"]:
        if r.get("block") != 6:
            continue
        dato = RECUPERADAS_A1C.get(r.get("line_no"))
        if not dato:
            continue
        if r.get("spec_no"):
            ya += 1
            continue  # ya recuperada en una corrida anterior
        vals = r.get("values") or {}
        if "200" in vals or "300" in vals:
            raise SystemExit("A-1C linea %s: se esperaba la fila huerfana, pero ya "
                             "trae 200/300 F." % r.get("line_no"))
        for k, v in zip(CAMPOS_REC, dato[:12]):
            r[k] = v
        r["product_form"] = None      # el impreso deja la columna vacia:
        r["min_temp_to_100"] = dato[12]  # la forma la da el encabezado de seccion
        nuevo = {"200": dato[13], "300": dato[14]}
        nuevo.update(vals)
        r["values"] = nuevo
        n += 1
        if len(cambios) < 5:
            cambios.append("linea %-4s folio %s  %s %s %s  S(100/200/300 F) = "
                           "%s / %s / %s"
                           % (r["line_no"], dato[15], dato[1], dato[3],
                              dato[4] or "", dato[12], dato[13], dato[14]))
    if n + ya != len(RECUPERADAS_A1C):
        raise SystemExit("A-1C: se esperaban %d filas de la tabla de recuperacion y "
                         "se localizaron %d (%d rellenadas ahora, %d ya estaban)."
                         % (len(RECUPERADAS_A1C), n + ya, n, ya))
    if n:
        cambios.insert(0, "%d paginas derechas huerfanas recuperadas del folio "
                          "impreso: identificacion completa y los puntos de 100, "
                          "200 y 300 F (%d valores de esfuerzo)" % (n, n * 3))
    doc["observaciones_del_codigo"] = [
        "Las dos ediciones NO numeran igual sus lineas: la linea 203 de A-1C "
        "(B649 / N08031) no es la 203 de A-1 (B164 / N04400). Cada fila "
        "recuperada se leyo en SU pagina de la edicion US.",
        "Discrepancia entre ediciones conservada tal como esta impresa: para el "
        "46Fe-24Ni-21Cr-6Mo-Cu-N (N08367) A-1 publica 427 C de temperatura "
        "maxima -unos 800 F- mientras que A-1C publica 900 F en las lineas 61 a "
        "65, 125 y 126, y 800 F en la 66. No se armoniza (regla 9).",
    ]
    return cambios


# ---------------------------------------------------------------------------
# Comprobacion posterior: el hueco tiene que haber quedado cerrado
# ---------------------------------------------------------------------------
# En 21 lineas del bloque 6 de A-1C la fusion de paginas enfrentadas fallo del
# todo: solo sobrevivio la pagina DERECHA. Esas filas se quedaron sin ninguna
# identificacion -sin spec, sin UNS, sin composicion- y sin los tres primeros
# puntos de la curva, que van impresos en la pagina izquierda: `Min. Temp. to
# 100`, `200` y `300` F. Su curva arranca en 400 F.
#
# NO se rellenan desde A-1: los valores estan en otras unidades y en otra
# rejilla de temperaturas, y convertirlos violaria la regla 9. La identificacion
# si podria copiarse de A-1 por (bloque, linea) -son B619, B622, B625, B626,
# B675, B690 y B804-, pero seria inferencia, no extraccion, y dejaria una fila
# con identidad prestada y valores incompletos: peor que una fila declarada como
# incompleta.
#
# Consecuencia real, acotada: esos 21 materiales de niquel NO son seleccionables
# en la edicion US del buscador. Al no tener identificacion no llegan a la
# cascada, asi que no devuelven un valor equivocado: faltan. Se cierra leyendo
# las 7 paginas izquierdas correspondientes del folio impreso.
LINEAS_HUERFANAS_A1C = [53, 54, 61, 62, 63, 64, 65, 66, 90, 91, 92, 118, 122,
                        125, 126, 159, 184, 185, 203, 211, 216]


def comprobar_sin_huerfanas(doc):
    """Ninguna fila puede quedar sin identificacion ni sin su punto de 200 F."""
    sin_200 = sorted(r.get("line_no") for r in doc["rows"]
                     if r.get("block") == 6 and (r.get("values") or {})
                     and "200" not in r["values"])
    sin_id = sorted(r.get("line_no") for r in doc["rows"]
                    if r.get("block") == 6 and not r.get("spec_no"))
    if sin_200 or sin_id:
        raise SystemExit(
            "A-1C: quedan filas huerfanas en el bloque 6.\n"
            "  sin el punto de 200 F: %s\n  sin identificacion: %s\n"
            "Las conocidas son %s: si aparecen otras, leanse en su folio."
            % (sin_200, sin_id, LINEAS_HUERFANAS_A1C))
    return []


def sellar(doc, archivo, cambios):
    doc["extraction_amendments"] = {
        "fecha": datetime.date.today().isoformat(),
        "script": "completar_apendice_a.py",
        "fuente": "appendix_a/fuente/APENDICE_A_ASME_B31_3_2024.pdf (escaneo, sin "
                  "capa de texto; correcciones verificadas a ojo sobre la pagina "
                  "ampliada)",
        "folios_impresos": FOLIOS[archivo],
        "alcance": "Solo desplazamientos de columna de la fusion de paginas "
                   "enfrentadas. Ningun valor tabulado se altera; el esfuerzo a "
                   "200 F de A-1C se MUEVE al punto de la curva donde lo imprime "
                   "el codigo.",
        "cambios": [c for c in cambios if not c.startswith("fila")],
    }


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args()

    base = a.resources / APXA
    if not base.is_dir():
        print("ERROR: no existe %s" % base, file=sys.stderr)
        return 2

    # A-1 se carga aparte: es la referencia que demuestra que en A-1C hay dos
    # celdas pegadas. No se usa para copiar datos, solo para probar el pegado.
    a1 = json.loads((base / "table_a_1.json").read_text(encoding="utf-8"))

    trabajo = [
        ("table_a_1.json", [arreglar_elipsis, arreglar_titulo_absorbido]),
        ("table_a_1c.json", [arreglar_metal_200, arreglar_elipsis,
                             arreglar_titulo_absorbido,
                             lambda d: arreglar_spec_grado_pegados(d, a1),
                             arreglar_uns_con_clase,
                             recuperar_huerfanas,
                             comprobar_sin_huerfanas]),
        ("table_a_3.json", [arreglar_a3]),
        ("table_a_4.json", [arreglar_elipsis]),
        ("table_a_4c.json", [arreglar_elipsis]),
    ]

    total = 0
    for archivo, fns in trabajo:
        ruta = base / archivo
        doc = json.loads(ruta.read_text(encoding="utf-8"))
        cambios = []
        for fn in fns:
            cambios += fn(doc)
        print("=" * 74)
        print("%s  folios %s" % (archivo, FOLIOS[archivo]))
        print("=" * 74)
        if cambios:
            for c in cambios:
                print("  -", c)
            total += sum(1 for c in cambios if not c.startswith("fila"))
        else:
            print("  ninguna (ya estaba corregido)")
        sellar(doc, archivo, cambios)
        if not a.dry_run:
            ruta.write_text(json.dumps(doc, ensure_ascii=False, indent=2) + "\n",
                            encoding="utf-8")
        print()

    print("%d correcciones%s." % (total, " (dry-run, no se escribio nada)"
                                  if a.dry_run else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
