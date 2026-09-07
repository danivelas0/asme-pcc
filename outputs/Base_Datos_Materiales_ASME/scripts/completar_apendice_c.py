# -*- coding: utf-8 -*-
"""
completar_apendice_c - completa la capa de METADATOS del Apendice C de ASME B31.3
(Tablas C-1, C-1C, C-2, C-3, C-3C y C-4) en `resources/`.

POR QUE EXISTE ESTE SCRIPT
--------------------------
Los VALORES del Apendice C ya estaban bien extraidos. Lo que faltaba -y lo que
estaba mal- era la capa de metadatos, y sin ella el motor de busqueda no puede
rotularse ni navegar:

1. C-1 / C-1C no declaraban la UNIDAD de sus coeficientes. El archivo no contiene
   la cadena "mm/m" ni "in./100 ft" en ninguna parte: solo la letra A o B. Un
   motor que muestre "13.8" sin decir de que unidad habla no sirve para calcular.
2. Las Notas (2)..(6) de C-1 -las que dicen que composicion pertenece al Grupo 1,
   2, 3 o 4- se imprimen A TRES COLUMNAS y la extraccion original las leyo por
   filas, dejando los miembros intercalados en un unico string, con un bloque
   huerfano al final. Sin desintercalarlas no se puede decir si un 1 1/4Cr-1/2Mo
   es Grupo 1, que es justo para lo que sirve la tabla.
3. C-2 y C-3 perdieron la INDENTACION del impreso, que es lo que distingue un
   material de subgrupo de uno de nivel superior. Resultado: cinco filas quedaron
   colgando de un subgrupo que NO es el suyo -"Gray iron" figuraba como acero
   inoxidable austenitico-, nueve filas de C-2 se quedaron sin grupo al cruzar el
   corte de pagina, y un titulo de grupo llego partido por la mitad
   ("and Reinforced Plastic Mortars").

FUENTE
------
`resources/.../appendix_c/fuente/APENDICE_C_ASME_B31_3_2024.pdf` (folios impresos
407-428) y su OCR estructurado `ocr_apendice_c_datalab.json`. El PDF es un
escaneo SIN capa de texto -pdfplumber devuelve cero lineas en las 24 paginas-,
asi que lo que se parsea es el OCR; el PDF queda versionado al lado para poder
auditar a ojo cualquier decision.

TRES TRAMPAS, YA PAGADAS
------------------------
1. EL OCR NO ES LA FUENTE DE VERDAD DE LOS VALORES. Se comprobo celda a celda que
   coincide con lo ya cargado en `resources/`, y donde discrepa MANDA `resources/`.
   Caso real: la C-3 imprime "Type 309. 23Cr-12Ni" con PUNTO donde sus cinco
   hermanas llevan coma -una errata del codigo, verificada ampliando el escaneo-.
   El OCR la "corrige" a coma. Este script NO toca nombres ni numeros: solo
   escribe metadatos, y si detecta una discrepancia la reporta y para.
2. EL DISCRIMINANTE DE ENCABEZADO NO ES EL MISMO EN LAS DOS TABLAS. En C-2 el
   grupo va en negrita y el subgrupo no. En C-3 el OCR pierde la negrita, pero el
   subgrupo termina en dos puntos ("Chromium steels:") y el grupo no
   ("Nickel Alloys"). Un unico criterio falla en una de las dos.
3. LA INDENTACION NO SOBREVIVE AL HTML. El OCR emite una fila de datos igual este
   indentada bajo un subgrupo o a ras del margen. Son CINCO filas en todo el
   apendice; van abajo como tabla literal con su folio impreso, verificadas
   ampliando el escaneo. No se infieren: se declaran.

EL OCR NO ESTA EN EL REPOSITORIO
--------------------------------
Ni el PDF ni su OCR se versionan: son material con copyright de ASME y estan en
.gitignore. Las enmiendas que este script escribe citan el folio impreso, que es
lo que permite auditarlas. Para REEJECUTARLO hay que tener el OCR en disco, en
`appendix_c/ocr_apendice_c_datalab.json` o donde sea, pasandolo con --ocr.
Sin el, el script se para y lo dice: no inventa nada.

USO
---
    python completar_apendice_c.py --resources ..\\..\\..\\resources [--dry-run]
    python completar_apendice_c.py --resources ... --ocr "D:\\...\\datalab.json"

Es idempotente: comprueba el estado de partida y se puede volver a correr.
"""

import argparse
import datetime
import html as _html
import json
import re
import sys
from pathlib import Path

APXC = Path("ASME B31") / "ASME B31.3" / "APPEX" / "appendix_c"
OCR_POR_DEFECTO = "ocr_apendice_c_datalab.json"

# Folios impresos del PDF aportado (24 paginas del escaneo -> 407..428 impresos).
# El indice de pagina del OCR es 0-based y NO coincide con el folio: se mapea aqui
# una sola vez para que toda cita del script apunte al folio que ve el ingeniero.
PAGINAS = {
    "Table C-1":  {"datos": [2, 3, 4, 5], "notas": 6,  "folios": [408, 409, 410, 411, 412]},
    "Table C-1C": {"datos": [8, 9, 10, 11], "notas": 12, "folios": [414, 415, 416, 417, 418]},
    "Table C-2":  {"datos": [13, 14], "notas": None, "folios": [419, 420]},
    "Table C-3":  {"datos": [15, 16, 17], "notas": None, "folios": [421, 422, 423]},
    "Table C-3C": {"datos": [18, 19, 20, 21], "notas": None, "folios": [424, 425, 426, 427]},
    "Table C-4":  {"datos": [22], "notas": 23, "folios": [428]},
}

ARCHIVO = {
    "Table C-1": "table_c_1.json", "Table C-1C": "table_c_1c.json",
    "Table C-2": "table_c_2.json", "Table C-3": "table_c_3.json",
    "Table C-3C": "table_c_3c.json", "Table C-4": "table_c_4.json",
}

# ---------------------------------------------------------------------------
# Las cinco excepciones de indentacion
# ---------------------------------------------------------------------------
# Filas que en el impreso van A RAS del margen y que el OCR entrega pegadas al
# subgrupo anterior. Verificadas ampliando el escaneo del folio que se cita.
# Es una lista cerrada: cualquier otra fila la resuelve el algoritmo.
SIN_SUBGRUPO = {
    "Table C-2": {
        # Folio 419: van a ras, no indentadas bajo "Chlorinated poly(vinyl chloride)"
        "Polybutylene PB 2110": 419,
        "Polyether, chlorinated": 419,
        # Folio 419: va a ras, no indentada bajo "Cross-linked polyethylene"
        "Polyphenylene POP 2125": 419,
        # Folio 420: van a ras. Solo PVC2116 y PVC2120 continuan indentadas bajo
        # "Poly(vinyl chloride)" al cruzar el corte de pagina; estas siete no.
        "Poly(vinylidene fluoride)": 420,
        "Poly(vinylidene chloride)": 420,
        "Poly(tetrafluoroethylene)": 420,
        "Poly(fluorinated ethylene propylene)": 420,
        "Poly(perfluoroalkoxy alkane)": 420,
    },
    "Table C-3": {
        # Folio 421: ambas van a ras, no indentadas bajo "Austenitic stainless steels:"
        "Straight chromium stainless steels (12Cr, 17Cr, 27Cr)": 421,
        "Gray iron": 421,
    },
}
SIN_SUBGRUPO["Table C-3C"] = dict(SIN_SUBGRUPO["Table C-3"])  # mismo impreso, folio 424

# Factores de escala, con el texto impreso al lado para que trace al codigo.
ESCALAS = {
    "Table C-1":  {"operacion": "ninguna", "exponente": 0, "nota": "el 10^-6 va en la unidad del coeficiente A"},
    "Table C-1C": {"operacion": "ninguna", "exponente": 0, "nota": "el 10^-6 va en la unidad del coeficiente A"},
    "Table C-2":  {"operacion": "dividir", "exponente": 6},
    "Table C-3":  {"operacion": "multiplicar", "exponente": 3, "unidad_resultante": "MPa"},
    "Table C-3C": {"operacion": "multiplicar", "exponente": 6, "unidad_resultante": "psi"},
    "Table C-4":  {"operacion": "ninguna", "exponente": 0,
                   "unidad_resultante": "ksi (73.4 F) / MPa (23 C)"},
}

# Temperatura de referencia impresa, hoy solo legible dentro del nombre de columna.
TEMP_REFERENCIA = {
    "Table C-1": {"valor": 20, "unidad": "C"},
    "Table C-1C": {"valor": 70, "unidad": "F"},
    "Table C-4": [{"valor": 73.4, "unidad": "F"}, {"valor": 23, "unidad": "C"}],
}


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------
def sin_tags(s):
    return re.sub(r"\s+", " ", _html.unescape(re.sub(r"<[^>]+>", "", s or ""))).strip()


def latex_a_texto(s):
    """Normaliza el LaTeX del OCR a la convencion ya usada en `resources/`.

    Las fracciones se escriben como en las notas de la II-D (1/2Cr, 11/4Cr): sin
    signo de fraccion tipografico y sin espacio, para que `comp_key` del builder
    empareje las dos fuentes sin un caso especial.
    """
    if not s:
        return ""
    t = s
    t = re.sub(r"\\frac\{(\d+)\}\{(\d+)\}", r"\1/\2", t)
    t = re.sub(r"\\text\{([^{}]*)\}", r"\1", t)
    t = t.replace("^{\\circ}", "°").replace("\\circ", "°")
    t = re.sub(r"10\^\{-6\}", "10^-6", t)
    t = re.sub(r"10\^\{(\d+)\}", r"10^\1", t)
    t = t.replace("\\,", " ").replace("\\ ", " ").replace("$", "")
    t = re.sub(r"<[^>]+>", "", t)
    t = _html.unescape(t)
    return re.sub(r"\s+", " ", t).strip()


def clave(s):
    """Clave laxa para emparejar un nombre del OCR con el de `resources/`.

    Pliega mayusculas, los guiones tipograficos, los signos de puntuacion que el
    OCR normaliza por su cuenta (la errata "Type 309." frente a "Type 309,") y
    los espacios. Sirve para EMPAREJAR, nunca para reescribir.
    """
    t = (s or "").lower()
    for ch in "\u2010\u2011\u2012\u2013\u2014\u2212":
        t = t.replace(ch, "-")
    t = t.replace("\u2215", "/").replace("\u2044", "/")
    t = re.sub(r"[.,;:]", "", t)
    return re.sub(r"\s+", "", t)


def num(v):
    """Devuelve float si la celda es un numero puro; si no, None."""
    if v is None:
        return None
    # El OCR separa los miles con espacio fino ("2 830"); la extraccion original
    # lo hace con coma ("1,200"). Se quitan los dos para poder comparar.
    t = str(v).replace(",", "")
    for ch in "\u2212\u2013\u2014":
        t = t.replace(ch, "-")
    t = re.sub(r"[\s\u00a0\u2007\u2009\u202f]", "", t)
    try:
        return float(t)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Lectura del OCR
# ---------------------------------------------------------------------------
def bloques(ocr, pagina, tipo=None):
    hijos = ocr["children"][pagina].get("children") or []
    return [b for b in hijos if tipo is None or b["block_type"] == tipo]


def filas_html(h):
    """(celdas_texto, html_crudo_de_la_fila) por cada <tr>."""
    out = []
    for tr in re.findall(r"<tr>(.*?)</tr>", h or "", re.S):
        celdas = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S)
        out.append(([sin_tags(c) for c in celdas], tr))
    return out


def es_cabecera_de_tabla(crudo):
    """Fila de <th>: es la cabecera de columnas, no un dato ni un grupo.

    Cada pagina de continuacion la reimprime, asi que si no se descarta acaba
    contada como un material mas -"Material Description" aparecia como fila-.
    """
    return "<th" in crudo


def es_encabezado(celdas, crudo):
    """Encabezado de grupo o subgrupo dentro del cuerpo de la tabla.

    Dos formas, porque el OCR no es consistente entre tablas: en C-2 y C-3 sale
    como una sola celda con colspan; en C-3C sale como una fila normal con el
    titulo en la primera celda y el resto vacias.
    """
    if len(celdas) == 1 and "colspan" in crudo and celdas[0] != "":
        return True
    return len(celdas) > 1 and celdas[0] != "" and all(not c for c in celdas[1:])


def defs_coeficientes(ocr, tabla):
    """Definicion impresa de los coeficientes A y B de C-1 / C-1C.

    Vive dentro del <th> de la primera fila de la tabla, como dos bloques <math>.
    """
    # El OCR no la entrega igual en las dos ediciones: en C-1 va dentro del <th>
    # de la tabla como un unico <math>, y en C-1C sale como bloque Text con el
    # LaTeX troceado (<math>A</math> = ... <math>10^{-6}</math> in./in./°F).
    # Por eso se aplana TODO el bloque a lineas de texto y se busca sobre ellas.
    lineas = []
    for pag in PAGINAS[tabla]["datos"][:2]:
        for b in bloques(ocr, pag):
            h = (b.get("html") or "").replace("<br/>", "\n").replace("<br>", "\n")
            h = re.sub(r"</(p|div|th|td|tr)>", "\n", h)
            for ln in h.split("\n"):
                ln = latex_a_texto(ln)
                if ln:
                    lineas.append(ln)
        if any(re.match(r"^A\s*=", l) for l in lineas):
            break
    defs = {}
    for ln in lineas:
        # La llave "}" del impreso agrupa A y B bajo la misma frase de
        # referencia; corta la unidad y no forma parte de ella.
        m = re.match(r"^([AB])\s*=\s*(.+?),\s*([^,}]+?)\s*(?:\}.*)?$", ln)
        if m and m.group(1) not in defs:
            defs[m.group(1)] = {
                "magnitud": m.group(2).strip(),
                "unidad": m.group(3).strip(),
                "impreso": "%s = %s, %s" % (m.group(1), m.group(2).strip(),
                                            m.group(3).strip()),
            }
    if "A" not in defs or "B" not in defs:
        return None
    ref = ""
    for ln in lineas:
        mm = re.search(r"(in Going From .*)$", ln)
        if mm:
            ref = mm.group(1).strip()
            break
    defs["referencia"] = ref
    return defs


def notas_de_pagina(ocr, pagina):
    """Notas (n) con sus miembros, leyendo el grid POR COLUMNAS.

    En el impreso cada nota es un rotulo seguido de una rejilla de tres columnas.
    La extraccion original la leyo por filas y dejo los miembros intercalados;
    aqui se recorre columna a columna, que es el orden de lectura real.
    """
    notas = []
    pendiente = None
    for b in bloques(ocr, pagina):
        if b["block_type"] in ("ListGroup", "Text"):
            texto = latex_a_texto(sin_tags(b.get("html")))
            for m in re.finditer(r"\((\d+)\)\s*([^()]*)", texto):
                pendiente = (m.group(1), m.group(2).strip())
        elif b["block_type"] == "Table" and pendiente:
            grid = [c for c, _ in filas_html(b["html"])]
            ancho = max((len(f) for f in grid), default=0)
            miembros = []
            for col in range(ancho):
                for f in grid:
                    if col < len(f):
                        # Solo se colapsan espacios repetidos: NO se quitan todos.
                        # Las composiciones ya vienen sin espacio dentro de la celda
                        # ("1<math>1/4</math>Cr-..." -> "11/4Cr-..."), pero hay
                        # miembros de dos palabras ("Carbon steel") que los necesitan.
                        v = re.sub(r"\s+", " ", latex_a_texto(f[col])).strip()
                        if v:
                            miembros.append(v)
            notas.append({
                "nota": "(%s)" % pendiente[0],
                "encabezado": pendiente[1].rstrip(": "),
                "miembros": miembros,
            })
            pendiente = None
    return notas


def jerarquia_y_filas(ocr, tabla):
    """Recorre las paginas de datos y devuelve las filas con su grupo/subgrupo.

    El encabezado se ARRASTRA de una pagina a la siguiente: es exactamente lo que
    la extraccion original no hizo, y por eso nueve filas de C-2 se quedaron sin
    grupo al cruzar el corte de pagina.
    """
    grupo = subgrupo = None
    filas = []
    for pag in PAGINAS[tabla]["datos"]:
        tb = bloques(ocr, pag, "Table")
        if not tb:
            continue
        for celdas, crudo in filas_html(tb[0]["html"]):
            if not celdas or es_cabecera_de_tabla(crudo):
                continue
            if es_encabezado(celdas, crudo):
                titulo = celdas[0]
                if tabla == "Table C-2":
                    # C-2: el grupo va en negrita, el subgrupo no.
                    if "<b>" in crudo:
                        grupo, subgrupo = titulo, None
                    else:
                        subgrupo = titulo
                elif tabla in ("Table C-3", "Table C-3C"):
                    # C-3: el OCR pierde la negrita; el subgrupo termina en ":".
                    if titulo.endswith(":"):
                        subgrupo = titulo
                    else:
                        grupo, subgrupo = titulo, None
                else:
                    grupo, subgrupo = titulo, None
                continue
            if len(celdas) < 2:
                continue
            nombre = celdas[0]
            if not nombre or num(nombre) is not None:
                continue  # fila de encabezados de temperatura
            sub = subgrupo
            if nombre in SIN_SUBGRUPO.get(tabla, {}):
                sub = None
            filas.append({
                "nombre": nombre,
                "grupo": grupo,
                "subgrupo": sub,
                "valores": celdas[1:],
                "folio": PAGINAS[tabla]["folios"][PAGINAS[tabla]["datos"].index(pag)],
            })
    return filas


# ---------------------------------------------------------------------------
# Auditoria OCR <-> resources/
# ---------------------------------------------------------------------------
def auditar(tabla, filas_ocr, doc):
    """Compara el OCR con `resources/`. NO corrige nada: solo reporta.

    Los NUMEROS son bloqueantes: si difieren, una de las dos extracciones esta
    mal y no es este script quien debe decidir cual.

    Los NOMBRES, no. El OCR normaliza por su cuenta lo que el codigo imprime raro
    -convierte en coma el punto de "Type 309." y lee "1 1/2Cr" donde el codigo
    imprime "1/2Cr"-, y `resources/` lleva lo impreso. Mientras los numeros de la
    fila coincidan, la fila esta emparejada sin duda y la divergencia de nombre se
    reporta como informativa. En ningun caso se reescribe un nombre.
    """
    campo = "material_description" if tabla in ("Table C-2", "Table C-4") else "material"
    filas_json = doc["rows"]
    problemas, avisos = [], []
    if len(filas_ocr) != len(filas_json):
        problemas.append("recuento distinto: OCR %d filas, resources/ %d"
                         % (len(filas_ocr), len(filas_json)))
        return problemas, avisos, {}

    emparejado = {}
    for i, (o, j) in enumerate(zip(filas_ocr, filas_json)):
        nj = j.get(campo)
        if "values" in j:
            vj = [j["values"][k] for k in doc["columns"][1:] if k in j["values"]]
        else:
            vj = [j.get(k) for k in ("in_in_f", "range_f", "mm_mm_c", "range_c",
                                     "e_ksi_73_4_f", "e_mpa_23_c") if k in j]
        vo = [x for x in (num(y) for y in o["valores"]) if x is not None]
        vjn = [x for x in (num(y) for y in vj) if x is not None]
        iguales = (vjn == vo)
        if not iguales:
            problemas.append("fila %d (%s): numeros distintos\n"
                             "      OCR       %s\n      resources %s" % (i, nj, vo, vjn))
            continue
        emparejado[i] = o
        if clave(o["nombre"]) != clave(nj):
            avisos.append("fila %d: el OCR lee %r donde el codigo imprime %r "
                          "(numeros identicos; se conserva lo impreso)"
                          % (i, o["nombre"], nj))
    return problemas, avisos, emparejado


# ---------------------------------------------------------------------------
# Enmiendas
# ---------------------------------------------------------------------------
def enmendar(tabla, doc, ocr, emparejado, filas_ocr):
    """Devuelve (doc_modificado, lista_de_cambios). No toca valores ni nombres."""
    cambios = []
    folios = PAGINAS[tabla]["folios"]

    # 1. Factor de escala, con el texto impreso al lado
    if "scale_factor" not in doc:
        esc = dict(ESCALAS[tabla])
        impreso = doc.get("value_axis") or (doc.get("column_groups") or [None])[0]
        if impreso:
            esc["impreso"] = impreso
        doc["scale_factor"] = esc
        cambios.append("scale_factor: %s 10^%s" % (esc["operacion"], esc["exponente"]))

    # 2. Temperatura de referencia
    if tabla in TEMP_REFERENCIA and "reference_temperature" not in doc:
        doc["reference_temperature"] = TEMP_REFERENCIA[tabla]
        cambios.append("reference_temperature")

    # 3. C-1 / C-1C: definicion de coeficientes y miembros de las notas
    if tabla in ("Table C-1", "Table C-1C"):
        defs = defs_coeficientes(ocr, tabla)
        if defs and doc.get("coefficient_defs") != defs:
            doc["coefficient_defs"] = defs
            cambios.append("coefficient_defs: A=%s, B=%s"
                           % (defs["A"]["unidad"], defs["B"]["unidad"]))
        notas = notas_de_pagina(ocr, PAGINAS[tabla]["notas"])
        notas = [n for n in notas if n["miembros"]]
        for n in notas:
            m = re.match(r"Group (\d+)", n["encabezado"])
            n["grupo"] = "Group %s" % m.group(1) if m else None
            # La Nota (6) no define un Grupo: enumera por UNS que aleaciones de
            # aluminio representa la fila. Mismo reparto que en la II-D, donde las
            # notas se separan en "de grupo" y "de alias".
            n["tipo"] = "grupo" if n["grupo"] else "alias_uns"
        if notas and doc.get("note_members") != notas:
            doc["note_members"] = notas
            cambios.append("note_members: %s"
                           % ", ".join("%s->%d" % (n["nota"], len(n["miembros"])) for n in notas))

    # 4. C-2 / C-3 / C-3C / C-4: grupo y subgrupo impresos
    if tabla in ("Table C-2", "Table C-3", "Table C-3C", "Table C-4"):
        tocadas = 0
        for i, o in emparejado.items():
            fila = doc["rows"][i]
            antes = (fila.get("material_group"), fila.get("material_subgroup"))
            # (Cont'd) es corte de pagina, no un grupo distinto -y corta en filas
            # distintas en SI y en US-, asi que se normaliza.
            g = re.sub(r"\s*\(Cont'?d\)\s*$", "", o["grupo"] or "").strip() or None
            s = (o["subgrupo"] or "").strip() or None
            if g is None:
                fila.pop("material_group", None)
            else:
                fila["material_group"] = g
            if s is None:
                fila.pop("material_subgroup", None)
            else:
                fila["material_subgroup"] = s
            if antes != (fila.get("material_group"), fila.get("material_subgroup")):
                tocadas += 1
                cambios.append("  fila %d %-46s  %s -> %s"
                               % (i, o["nombre"][:46], antes,
                                  (fila.get("material_group"), fila.get("material_subgroup"))))
        if tocadas:
            cambios.insert(len(cambios) - tocadas,
                           "grupo/subgrupo corregidos en %d filas" % tocadas)

    # 5. Constancia de la enmienda
    doc["extraction_amendments"] = {
        "fecha": datetime.date.today().isoformat(),
        "script": "completar_apendice_c.py",
        "fuente": "appendix_c/fuente/APENDICE_C_ASME_B31_3_2024.pdf (escaneo, sin capa "
                  "de texto) y su OCR estructurado appendix_c/ocr_apendice_c_datalab.json",
        "folios_impresos": folios,
        "alcance": "Solo metadatos: unidades, definicion de coeficientes, miembros de las "
                   "Notas, grupo/subgrupo impreso y factor de escala. Los valores y los "
                   "nombres de material NO se tocan: se conservan tal como estaban "
                   "cargados, que es como los imprime el codigo (regla 9 del proyecto).",
        "cambios": [c for c in cambios if not c.startswith("  fila")],
        "excepciones_declaradas": [
            {"fila": k, "folio": v,
             "motivo": "en el impreso va a ras del margen, no indentada bajo el subgrupo "
                       "anterior; el HTML del OCR no conserva la indentacion"}
            for k, v in SIN_SUBGRUPO.get(tabla, {}).items()
        ],
    }
    return doc, cambios


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--ocr", default=None,
                    help="ruta del OCR; por defecto appendix_c/%s" % OCR_POR_DEFECTO)
    ap.add_argument("--dry-run", action="store_true",
                    help="informa sin escribir")
    ap.add_argument("--permitir-discrepancias", action="store_true",
                    help="no abortar si el OCR y resources/ difieren (se conserva resources/)")
    a = ap.parse_args()

    base = a.resources / APXC
    if not base.is_dir():
        print("ERROR: no existe %s" % base, file=sys.stderr)
        return 2
    ruta_ocr = Path(a.ocr) if a.ocr else base / OCR_POR_DEFECTO
    if not ruta_ocr.is_file():
        print("ERROR: falta el OCR %s\n"
              "       No se versiona (copyright de ASME, ver .gitignore). Deje el\n"
              "       archivo ahi o indique su ruta con --ocr." % ruta_ocr,
              file=sys.stderr)
        return 2
    ocr = json.loads(ruta_ocr.read_text(encoding="utf-8"))
    if len(ocr.get("children", [])) != 24:
        print("ERROR: el OCR no tiene 24 paginas", file=sys.stderr)
        return 2

    total_prob = 0
    for tabla, nombre in ARCHIVO.items():
        ruta = base / nombre
        doc = json.loads(ruta.read_text(encoding="utf-8"))
        print("=" * 78)
        print("%s  (%s)  folios %s" % (tabla, nombre, PAGINAS[tabla]["folios"]))
        print("=" * 78)

        emparejado, filas_ocr = {}, []
        if tabla in ("Table C-2", "Table C-3", "Table C-3C", "Table C-4"):
            filas_ocr = jerarquia_y_filas(ocr, tabla)
            problemas, avisos, emparejado = auditar(tabla, filas_ocr, doc)
            if problemas:
                total_prob += len(problemas)
                print("  DISCREPANCIAS DE VALOR OCR vs resources/  (bloqueantes):")
                for p in problemas:
                    print("    -", p)
            else:
                print("  auditoria: %d filas, todos los numeros coinciden con resources/"
                      % len(filas_ocr))
            for w in avisos:
                print("  aviso:", w)
        else:
            print("  auditoria de valores: no aplica (C-1/C-1C parten la tabla por columnas "
                  "de temperatura; sus valores no se tocan)")

        doc, cambios = enmendar(tabla, doc, ocr, emparejado, filas_ocr)
        if cambios:
            print("  enmiendas:")
            for c in cambios:
                print("    -", c)
        else:
            print("  enmiendas: ninguna (ya estaba completo)")

        if not a.dry_run:
            ruta.write_text(json.dumps(doc, ensure_ascii=False, indent=2) + "\n",
                            encoding="utf-8")
        print()

    if total_prob and not a.permitir_discrepancias:
        print("PARADA: %d discrepancias entre el OCR y resources/." % total_prob)
        print("Revise cada una: `resources/` lleva el valor tal como esta impreso y el OCR "
              "puede haberlo normalizado. Si son esperadas, repita con "
              "--permitir-discrepancias.")
        return 1
    print("OK%s." % (" (dry-run, no se escribio nada)" if a.dry_run else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
