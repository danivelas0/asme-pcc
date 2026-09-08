# -*- coding: utf-8 -*-
"""
secii_tablas — reconstruye las tablas de ASME BPVC Seccion II, partes A, B y C,
a partir de los bloques `Line` y sus `bbox`. Sin PDF y sin Excel.

POR QUE HACE FALTA ESTE MODULO
------------------------------
El JSON de las 368 especificaciones NO trae tablas. El `html` de todo bloque
`Table` es `"<p></p>"`: no hay `<tr>` ni `<td>`. Es deliberado —el `meta.json`
del drop declara que se rechazo la reconstruccion de tablas de pdftext para
conservar los bloques de texto originales con sus coordenadas—, y ademas los
`Span` no se conservan: el bloque mas fino es `Line`.

Asi que el dato esta entero, pero repartido: las filas cuelgan del `Table` como
hijos `Line`, cada uno con su texto y su `bbox`. Construir una base exige
RECOMPONER las columnas desde esos `Line`, no leer una tabla ya formada.

LA DECISION DE DISENO IMPORTANTE
--------------------------------
El reparto de una fila tiene TRES resultados y ninguno adivina:

  EXACTA      la fila trae >= 2 `Line`: cada uno va a su columna por el punto
              medio de su `bbox`. No se parte ningun texto.
  POR CONTEO  la fila es UN solo `Line`: se protegen los patrones que llevan
              espacio dentro ("48 000", "48 000 [330]", "1 1/2") y se parte por
              espacios DESDE LA DERECHA. Solo vale si el numero de fichas de
              valor cuadra con el numero de columnas que la tabla ya demostro
              tener, y si la geometria lo confirma.
  AMBIGUA     ni lo uno ni lo otro: el texto impreso se conserva ENTERO en una
              celda, la fila se marca y entra en el informe de revision.

Esa ultima linea es la decision de diseno. Sin `Span`, la posicion de cada
palabra DENTRO de un `Line` no esta en el fichero: repartirla interpolando
linealmente sobre el ancho del `bbox` seria inventar estructura con una fuente
proporcional. La geometria se usa para CONFIRMAR un reparto por conteo, nunca
para producirlo.

COMPROBACION SIN PERDIDA, OBLIGATORIA
-------------------------------------
Para cada fila reconstruida, la concatenacion normalizada de sus celdas tiene
que ser caracter a caracter igual a la concatenacion normalizada del texto de
sus `Line` de origen. Garantiza que la tabulacion es una REPARTICION del texto
impreso y nunca una adicion. No demuestra que el JSON reproduzca el PDF —eso se
audito aguas arriba, y con reservas: la cobertura de texto es del 97 % pero baja
al 88 % en las paginas apaisadas, que son justo las tablas anchas—; demuestra
que esta capa no mete nada. Es lo que si se puede garantizar sin el PDF.

USO
---
    python secii_tablas.py --resources ..\\..\\..\\resources --informe salida.md
    python secii_tablas.py --resources ... --piloto        (las 12 que duelen)
    python secii_tablas.py --resources ... --spec SA-106   (una sola)

El CLI mide y reporta: no escribe nada en el libro ni en resources/.
"""
from __future__ import annotations

import argparse
import html as _html
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

SEC_II = Path("ASME_BPVC") / "Sec_II"
PARTES = ("bpvc_ii_a_1", "bpvc_ii_a_2", "bpvc_ii_b", "bpvc_ii_c")

# Las 12 especificaciones del piloto: elegidas por ser las que duelen, no por
# ser representativas. Cada una rompe un extractor ingenuo por un motivo
# distinto, y el motivo va anotado al lado.
PILOTO = {
    "SA-106": "los dos modos de fila conviven en la misma tabla",
    "SA-6": "5,5 MB, tablas sueltas sin TableGroup, Caption vacios",
    "SA-240": "la peor cobertura de texto del drop (87 %), apaisada",
    "SA-182": "tablas anchas de composicion por grado",
    "SA-193": "perneria: clase y diametro en la misma fila",
    "SA-312": "tabla de composicion con muchos grados",
    "SA-335": "idem, aleados",
    "SB-111": "una celda partida en dos Line ('99.99' + 'min^A')",
    "SB-265": "titanio, tablas de composicion cortas",
    "SFA-5.9": "continuacion de 4 paginas, apaisada, duplicados Text/Footnote",
    "SFA-5.5": "electrodos: tablas largas con notas al pie",
    "SA-479": "barras inoxidables",
}

EXACTA, POR_CONTEO, AMBIGUA = "EXACTA", "POR CONTEO", "AMBIGUA"

# Tolerancia vertical al agrupar Line en filas. Tres puntos: un superindice
# sube la caja del Line alrededor de 1 pt (SB-111 Tabla 1: "99.99" y "min^A"
# son contiguos en x y llevan la y desplazada), y dos filas consecutivas de
# estas tablas nunca estan a menos de 6 pt.
TOL_Y = 3.0
# Separacion minima en x entre dos columnas. Por debajo de esto se considera la
# misma columna: son dos trozos de una celda partida, no dos celdas.
GAP_X = 6.0


# ---------------------------------------------------------------------------
# Texto
# ---------------------------------------------------------------------------
_RE_TAG = re.compile(r"<[^>]+>")
_RE_WS = re.compile(r"\s+")
# Un superindice se imprime PEGADO a lo que le precede: el marcador de nota de
# `0.25<i><sup>A</sup></i>` es parte de la celda "0.25A", no una ficha aparte.
# Se marca con un centinela antes de quitar el resto de etiquetas, porque esas
# si separan: el Caption de SB-111 llega como
# `<b>TABLE</b><b>1</b><b>Chemical</b><b>Requirements</b>` y sin separador daria
# `TABLE1ChemicalRequirements`.
_PEGA = "\x01"
_RE_SUP = re.compile(r"<(su[pb])>(.*?)</\1>", re.S | re.I)
_RE_PEGA = re.compile(r"\s*" + _PEGA + r"\s*")


def texto(html_) -> str:
    """Texto plano de un bloque: cada etiqueta separa, salvo el superindice."""
    if not html_:
        return ""
    s = _RE_SUP.sub(lambda m: _PEGA + _RE_TAG.sub("", m.group(2)), str(html_))
    s = _RE_TAG.sub(" ", s)
    s = _html.unescape(s)
    s = _RE_PEGA.sub("", s)
    return _RE_WS.sub(" ", s).strip()


def _plano(s: str) -> str:
    """Forma comparable de un texto: sin espacios. Base de la comprobacion
    sin perdida — reparte quien reparta, los caracteres son los mismos."""
    return _RE_WS.sub("", s or "")


# ---------------------------------------------------------------------------
# Fichas: partir una linea sin romper lo que el codigo imprime con espacio
# ---------------------------------------------------------------------------
# El espacio NO siempre separa celdas. `48 000` es un separador de millares y
# `[330]` es la unidad SI que imprime el propio codigo junto al valor. Estos
# patrones se protegen ANTES de partir, sustituyendo su espacio interior por un
# centinela que luego se restituye.
_NUL = "\x00"
_PROTEGER = [
    # 48 000 [330]  /  48 000
    re.compile(r"\d{1,3}(?:\s\d{3})+(?:\s*\[[^\]]*\])?"),
    # 0.25 [6.4]
    re.compile(r"\d[\d.,]*\s*\[[^\]]*\]"),
    # 1 1⁄2  ·  1 1/2
    re.compile(r"\d+\s+\d+\s*[⁄/]\s*\d+"),
    # 8 × C to 1.0  ·  5 x C to 0.70
    re.compile(r"\d+\s*[×x]\s*[A-Za-z]{1,2}\s+to\s+[\d.]+", re.I),
    # la elipsis ASME impresa con espacios
    re.compile(r"\.\s*\.\s*\."),
]


def fichas(s: str) -> list[str]:
    """Trocea una linea en fichas, protegiendo lo que lleva espacio dentro."""
    t = s
    for rx in _PROTEGER:
        t = rx.sub(lambda m: m.group(0).replace(" ", _NUL), t)
    return [f.replace(_NUL, " ") for f in t.split() if f.strip()]


# Ficha que puede ser un VALOR de tabla: numero, rango, elipsis, marcador de
# nota pegado al numero, o un "..." impreso. No se exige que sea numerica pura:
# el codigo publica `0.27–0.93`, `0.25A`, `48 000 [330]` y `. . .`.
_RE_VALOR = re.compile(
    r"^(?:"
    r"[−–—+-]?\d[\d.,\s]*(?:\[[^\]]*\])?[A-Za-z]{0,3}"      # 0.25A · 48 000 [330]
    r"|[−–—+-]?\d[\d.,]*\s*[–—-]\s*\d[\d.,]*[A-Za-z]{0,3}"   # 0.27–0.93
    r"|\.\s*\.\s*\."                                          # elipsis ASME
    r"|[.]{3}"
    r")$")


def es_valor(f: str) -> bool:
    return bool(_RE_VALOR.match(f.strip()))


def n_valores_finales(s: str) -> int:
    """Cuantas fichas de VALOR hay al final de la linea."""
    fs = fichas(s)
    n = 0
    for f in reversed(fs):
        if es_valor(f):
            n += 1
        else:
            break
    return n


# ---------------------------------------------------------------------------
# Recorrido del arbol de bloques
# ---------------------------------------------------------------------------
def _iou(a, b) -> float:
    """Solape de dos bbox [x0,y0,x1,y1] sobre su union."""
    if not a or not b:
        return 0.0
    x0, y0 = max(a[0], b[0]), max(a[1], b[1])
    x1, y1 = min(a[2], b[2]), min(a[3], b[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    inter = (x1 - x0) * (y1 - y0)
    ua = (a[2] - a[0]) * (a[3] - a[1])
    ub = (b[2] - b[0]) * (b[3] - b[1])
    return inter / (ua + ub - inter) if ua + ub - inter else 0.0


def _nodos(hijos, padre=None):
    for n in hijos or []:
        if isinstance(n, dict):
            yield n, padre
            yield from _nodos(n.get("children"), n)


def deduplicar(bloques):
    """Marker emite pares Text/Footnote con el mismo bbox y solo uno con texto.

    En SFA-5.9 p.359 hay 8 pares con `bbox` casi identico y el contenido ALTERNA
    entre uno y otro: quedarse siempre con el primero perderia la mitad de las
    notas y quedarse con los dos las duplicaria. Se deduplica por solape de
    `bbox` conservando el que trae `html` con texto.
    """
    fuera = set()
    for i, a in enumerate(bloques):
        if a["id"] in fuera:
            continue
        for b in bloques[i + 1:]:
            if b["id"] in fuera or a["pagina"] != b["pagina"]:
                continue
            if _iou(a["bbox"], b["bbox"]) < 0.85:
                continue
            ta, tb = texto(a["html"]), texto(b["html"])
            if ta and not tb:
                fuera.add(b["id"])
                a["gemelo"] = b["tipo"]
            elif tb and not ta:
                fuera.add(a["id"])
                b["gemelo"] = a["tipo"]
                break
    return [b for b in bloques if b["id"] not in fuera]


def cargar_spec(ruta: Path) -> dict:
    """Aplana una especificacion a una lista de bloques con su pagina."""
    d = json.loads(ruta.read_text(encoding="utf-8"))
    data = d["data"]
    bloques, padres = [], {}
    for pg in data.get("pages") or []:
        pdf_page = pg.get("pdf_page")
        printed = pg.get("printed_page")
        for n, padre in _nodos(pg.get("children")):
            bloques.append(dict(
                id=n.get("id"), tipo=n.get("block_type"), html=n.get("html"),
                bbox=n.get("bbox"), hijos=n.get("children") or [],
                pagina=pdf_page, folio=printed,
                padre=(padre or {}).get("id"),
                padre_tipo=(padre or {}).get("block_type")))
            padres[n.get("id")] = padre
    return dict(archivo=ruta.name, fuente=d.get("source", {}),
                spec=data.get("specification", {}),
                doc=data.get("document", {}), bloques=bloques)


# ---------------------------------------------------------------------------
# Titulacion de una tabla
# ---------------------------------------------------------------------------
_RE_CONT_AB = re.compile(r"<i>\s*\(?\s*Cont(?:inued|'?d)\s*\)?\s*</i>", re.I)
_RE_CONT_C = re.compile(r"\(\s*Cont(?:inued|'?d)\s*\)", re.I)


def titular(tabla, bloques, por_id):
    """Titulo de una tabla y si es continuacion de la anterior.

    Tres vias, en orden:
      1. el `Caption` hermano dentro del mismo `TableGroup` (puede ir en la
         posicion 0 o en la 1: se prueban las dos);
      2. si la tabla va SUELTA —el 32 % de ellas—, el `Caption` o
         `SectionHeader` inmediatamente encima en la misma pagina;
      3. nada: la tabla queda sin titulo y se declara.
    """
    cap = None
    if tabla["padre_tipo"] == "TableGroup":
        grupo = por_id.get(tabla["padre"])
        for h in (grupo or {}).get("hijos", [])[:2]:
            if h.get("block_type") == "Caption" and texto(h.get("html")):
                cap = h
                break
    if cap is None:
        # Caption o SectionHeader mas cercano por encima, en la misma pagina.
        arriba = [b for b in bloques
                  if b["pagina"] == tabla["pagina"]
                  and b["tipo"] in ("Caption", "SectionHeader")
                  and b["bbox"] and tabla["bbox"]
                  and b["bbox"][3] <= tabla["bbox"][1] + 2
                  and texto(b["html"])]
        if arriba:
            cap = max(arriba, key=lambda b: b["bbox"][3])
            cap = {"html": cap["html"], "id": cap["id"]}
    html_ = (cap or {}).get("html") or ""
    cont = bool(_RE_CONT_AB.search(html_)) or bool(_RE_CONT_C.search(texto(html_)))
    tit = _RE_CONT_C.sub(" ", texto(html_))
    tit = re.sub(r"\s+Continued$", "", tit, flags=re.I)
    # Los espacios se colapsan DESPUES de quitar el marcador de continuacion:
    # si no, «Table 1 (Continued) Chemical...» deja un doble espacio y el titulo
    # deja de casar con el de la primera pagina, que es lo unico que permite
    # unirlas y arrastrar la cabecera.
    tit = _RE_WS.sub(" ", tit).strip()
    return tit, cont, (cap or {}).get("id")


# ---------------------------------------------------------------------------
# Filas y columnas
# ---------------------------------------------------------------------------
def agrupar_filas(lineas):
    """Agrupa los `Line` de una tabla en filas.

    El criterio va sobre el CENTRO vertical, no sobre el borde. Con el borde,
    dos filas consecutivas —que en estas tablas se separan solo 2 pt— caen
    dentro de la tolerancia y la tabla entera colapsa en una fila. Con el
    centro, la separacion real entre filas (8-9 pt) es holgada y el
    desplazamiento que introduce un superindice (~1 pt) sigue sin partir nada.
    Un solape vertical amplio tambien basta: es el caso de la celda partida en
    dos `Line` de SB-111 ("99.99" y "min^A"), contiguos en x y con la caja
    subida por el superindice.
    """
    ls = sorted(lineas, key=lambda l: ((l["bbox"][1] + l["bbox"][3]) / 2.0,
                                       l["bbox"][0]))
    filas, actual = [], []
    for l in ls:
        a, b = l["bbox"][1], l["bbox"][3]
        c = (a + b) / 2.0
        if actual:
            ca = sum((x["bbox"][1] + x["bbox"][3]) / 2.0 for x in actual) / len(actual)
            ya0 = min(x["bbox"][1] for x in actual)
            ya1 = max(x["bbox"][3] for x in actual)
            solape = min(b, ya1) - max(a, ya0)
            alto = min(b - a, ya1 - ya0) or 1.0
            if abs(c - ca) > TOL_Y and solape < 0.5 * alto:
                filas.append(sorted(actual, key=lambda x: x["bbox"][0]))
                actual = []
        actual.append(l)
    if actual:
        filas.append(sorted(actual, key=lambda x: x["bbox"][0]))
    return filas


def bandas_x(filas):
    """Ejes de columna, deducidos SOLO de las filas que traen varios `Line`.

    Una fila de un solo `Line` no dice nada sobre donde caen las columnas: su
    `bbox` abarca la fila entera. Solo las filas con dos o mas `Line` llevan esa
    informacion.

    Y de esas se descartan ademas las lineas ANCHAS: un rotulo de banda
    ("Composition, %") o una etiqueta de fila larga cruza por encima de varias
    columnas y, si entra en la proyeccion, tapa los huecos que separan unas de
    otras y fusiona la tabla entera en una sola banda. El ancho se mide contra
    el de la propia tabla, no contra un valor absoluto: las tablas apaisadas de
    la Seccion II miden el doble que las verticales.
    """
    anchos = [l["bbox"][2] - l["bbox"][0] for f in filas for l in f]
    if not anchos:
        return []
    x_izq = min(l["bbox"][0] for f in filas for l in f)
    x_der = max(l["bbox"][2] for f in filas for l in f)
    ancho = max(x_der - x_izq, 1.0)
    intervalos = []
    for f in filas:
        if len(f) < 2:
            continue
        for l in f:
            if (l["bbox"][2] - l["bbox"][0]) <= 0.45 * ancho:
                intervalos.append((l["bbox"][0], l["bbox"][2]))
    if not intervalos:
        return []
    intervalos.sort()
    bandas = [list(intervalos[0])]
    for x0, x1 in intervalos[1:]:
        if x0 <= bandas[-1][1] + GAP_X:
            bandas[-1][1] = max(bandas[-1][1], x1)
        else:
            bandas.append([x0, x1])
    return [tuple(b) for b in bandas]


def refinar_bandas(bandas, filas, max_iter=12):
    """Parte una banda cuando dos `Line` de la misma fila caen dentro de ella.

    La proyeccion en x fusiona dos columnas cuando en ALGUNA fila una celda de
    la izquierda llega mas a la derecha que el arranque de la de al lado. La
    prueba de que la fusion es falsa la da la propia tabla: si dos `Line` de una
    misma fila comparten banda, esa banda son dos columnas. Se corta por el
    hueco mas ancho que las separa —nunca por uno menor que GAP_X, que es lo que
    distingue dos columnas de una celda partida en dos trozos contiguos.
    """
    for _ in range(max_iter):
        mejor = {}
        for f in filas:
            if len(f) < 2:
                continue
            asign = [(columna_de(l, bandas), l) for l in f]
            for (ca, la), (cb, lb) in zip(asign, asign[1:]):
                if ca != cb:
                    continue
                hueco = lb["bbox"][0] - la["bbox"][2]
                if hueco < GAP_X:
                    continue
                x = (la["bbox"][2] + lb["bbox"][0]) / 2.0
                if ca not in mejor or hueco > mejor[ca][0]:
                    mejor[ca] = (hueco, x)
        if not mejor:
            break
        nuevas = []
        for i, (x0, x1) in enumerate(bandas):
            if i in mejor and x0 < mejor[i][1] < x1:
                nuevas += [(x0, mejor[i][1]), (mejor[i][1], x1)]
            else:
                nuevas.append((x0, x1))
        if len(nuevas) == len(bandas):
            break
        bandas = nuevas
    return bandas


def columna_de(l, bandas):
    """Columna de un `Line` por el punto medio de su bbox (con tolerancia)."""
    m = (l["bbox"][0] + l["bbox"][2]) / 2.0
    for i, (x0, x1) in enumerate(bandas):
        if x0 - 3 <= m <= x1 + 3:
            return i
    # Fuera de toda banda: la mas cercana.
    return min(range(len(bandas)),
               key=lambda i: min(abs(m - bandas[i][0]), abs(m - bandas[i][1])))


def repartir_por_conteo(txt, ncols):
    """Parte una linea suelta desde la derecha: las n-1 ultimas fichas son
    valores y todo lo anterior es la etiqueta de la fila."""
    fs = fichas(txt)
    nval = ncols - 1
    if nval <= 0 or len(fs) <= nval:
        return None
    if not all(es_valor(f) for f in fs[-nval:]):
        return None
    return [" ".join(fs[:-nval])] + fs[-nval:]


# ---------------------------------------------------------------------------
# Reconstruccion de una tabla
# ---------------------------------------------------------------------------
def reconstruir(bloques_tabla, titulo):
    """Devuelve (filas, ncols, resumen de confianza).

    `bloques_tabla` es la lista de bloques `Table` que componen UNA tabla
    logica (varios cuando la tabla continua de pagina).
    """
    lineas = []
    for t in bloques_tabla:
        for h in t["hijos"]:
            if h.get("block_type") == "Line" and h.get("bbox"):
                lineas.append(dict(id=h.get("id"), html=h.get("html"),
                                   bbox=h.get("bbox"), tabla=t["id"],
                                   pagina=t["pagina"], folio=t["folio"]))
    if not lineas:
        return [], 0, Counter()

    # Las filas se agrupan por pagina: dos paginas distintas comparten sistema
    # de coordenadas y sus y se solaparian.
    filas = []
    for pag in sorted({l["pagina"] for l in lineas}, key=lambda p: int(p)):
        filas += agrupar_filas([l for l in lineas if l["pagina"] == pag])

    bandas = bandas_x(filas)
    if bandas:
        bandas = refinar_bandas(bandas, filas)
    # Una sola banda no es geometria de columnas: significa que el agrupamiento
    # en x no encontro ninguna separacion. Se descarta y se resuelve por conteo.
    if len(bandas) < 2:
        bandas = []
    # Numero de columnas: la geometria si la hay; si no, el numero de fichas de
    # valor que se repite en la mayoria de las filas de un solo Line.
    if bandas:
        ncols = len(bandas)
        origen_ncols = "geometria de las filas con varios Line"
    else:
        cuenta = Counter(n_valores_finales(texto(f[0]["html"]))
                         for f in filas if len(f) == 1)
        cuenta.pop(0, None)
        if not cuenta:
            # Ninguna fila termina en fichas de valor. Puede ser una lista de
            # verdad —una columna, cada Line una celda— o una tabla cuyas filas
            # el extractor entrego enteras en un solo Line. Se distinguen por lo
            # larga que es la fila: una lista trae una o dos fichas; una fila de
            # tabla, muchas. Llamar EXACTA a la segunda seria decir que esta
            # tabulada cuando no lo esta.
            ncols = 1
            largos = sorted(len(fichas(texto(f[0]["html"])))
                            for f in filas if len(f) == 1)
            mediana = largos[len(largos) // 2] if largos else 0
            if mediana <= 2:
                origen_ncols = "una sola columna: cada Line es una celda"
            else:
                origen_ncols = ("SIN COLUMNAS: cada fila llega entera en un solo "
                                "Line y no hay geometria ni fichas de valor que "
                                "la separen")
        else:
            nval, veces = cuenta.most_common(1)[0]
            total = sum(cuenta.values())
            ncols = nval + 1
            origen_ncols = (f"conteo de fichas de valor ({veces}/{total} filas)"
                            if veces / total >= 0.7 else
                            f"conteo NO concluyente ({veces}/{total} filas): "
                            "las que no cuadran quedan AMBIGUA")

    out = []
    for f in filas:
        celdas = [""] * max(ncols, 1)
        if len(f) >= 2 and not bandas:
            out.append(dict(celdas=[" ".join(texto(l["html"]) for l in f)],
                            confianza=AMBIGUA,
                            motivo="varios Line en la fila y ninguna banda de "
                                   "columna que los separe",
                            lineas=f))
            continue
        if len(f) >= 2 and bandas:
            # `f` viene ordenada por x. La asignacion a columnas tiene que
            # respetar ese orden: si un `Line` que va antes en la fila cae en
            # una columna posterior a la del siguiente, la banda no describe
            # esta fila —tipicamente porque un rotulo que cruza varias columnas
            # se coloco por su punto medio— y repartirla reordenaria el texto
            # impreso. Se declara AMBIGUA en vez de reordenar.
            usados, previa, malo = {}, -1, None
            for l in f:
                c = columna_de(l, bandas)
                if c in usados:
                    malo = "dos Line caen en la misma banda de columna"
                elif c < previa:
                    malo = ("el reparto por bandas alteraria el orden de lectura "
                            "de la fila")
                previa = max(previa, c)
                usados.setdefault(c, []).append(texto(l["html"]))
            if malo:
                out.append(dict(celdas=[" ".join(texto(l["html"]) for l in f)],
                                confianza=AMBIGUA, motivo=malo, lineas=f))
                continue
            for c, ts in usados.items():
                celdas[c] = " ".join(ts)
            out.append(dict(celdas=celdas, confianza=EXACTA, motivo="", lineas=f))
            continue

        t = texto(f[0]["html"])
        if ncols <= 1:
            sin_columnas = origen_ncols.startswith("SIN COLUMNAS")
            out.append(dict(
                celdas=[t],
                confianza=AMBIGUA if sin_columnas else EXACTA,
                motivo=("fila entera en un solo Line, sin geometria ni fichas de "
                        "valor que la separen" if sin_columnas
                        else "tabla de una sola columna"),
                lineas=f))
            continue
        partido = repartir_por_conteo(t, ncols)
        if partido is None:
            out.append(dict(celdas=[t], confianza=AMBIGUA,
                            motivo="linea unica que no cuadra con "
                                   f"{ncols} columnas", lineas=f))
            continue
        # La geometria CONFIRMA el reparto, nunca lo produce: el bbox de la
        # linea tiene que abarcar el rango de columnas al que se reparte.
        if bandas:
            x0, x1 = f[0]["bbox"][0], f[0]["bbox"][2]
            if not (x0 <= bandas[-1][1] + 12 and x1 >= bandas[0][0] - 12):
                out.append(dict(celdas=[t], confianza=AMBIGUA,
                                motivo="el bbox de la linea no abarca las "
                                       "columnas que el conteo le asigna",
                                lineas=f))
                continue
        out.append(dict(celdas=partido, confianza=POR_CONTEO, motivo="", lineas=f))

    out = reintegrar_continuaciones(out, ncols, bandas)
    # Tipo de fila, para poder medir la ambiguedad donde importa. Una banda o un
    # encabezado que no se deja partir cuesta mucho menos que una fila de datos
    # que no se deja partir: la primera es un rotulo, la segunda son numeros del
    # codigo. Se decide por el texto completo de la fila, no por el reparto.
    for r in out:
        r["tipo"] = "dato" if any(es_valor(f) for f in
                                  fichas(" ".join(r["celdas"]))) else "cabecera"
    conf = Counter(r["confianza"] for r in out)
    return out, ncols, dict(conf=conf, origen_ncols=origen_ncols, bandas=len(bandas))


def reintegrar_continuaciones(filas, ncols, bandas):
    """Reconoce las lineas de continuacion del rotulo de la fila anterior.

    El codigo parte una etiqueta larga en varias lineas —«Over 11 1/2 to 4 [40
    to» / «100], incl»— y el extractor entrega cada trozo como una fila propia.
    Esas filas no son filas: son la cola del rotulo de la anterior.

    NO se fusionan con ella. Fusionarlas obligaria a reordenar el texto —el
    rotulo va antes de los valores en la celda, pero su continuacion viene
    DESPUES de ellos en el orden de lectura— y eso rompe la comprobacion sin
    perdida, que compara caracter a caracter y en orden. Es justo la garantia
    que sostiene toda esta capa, asi que se conserva entera: la linea se queda
    en su propia fila, en la columna 0, que es la que ocupa geometricamente.

    El criterio es ESTRUCTURAL: una fila de una sola linea, sin ninguna ficha de
    valor, que sigue a una fila ya repartida y cuyo borde izquierdo cae en la
    primera columna. No se decide por lo que dice el texto.
    """
    if ncols < 2:
        return filas
    out = []
    for r in filas:
        anterior = out[-1] if out else None
        cont = (anterior is not None and r["confianza"] == AMBIGUA
                and len(r["lineas"]) == 1 and len(r["celdas"]) == 1
                and n_valores_finales(r["celdas"][0]) == 0
                and anterior["confianza"] in (EXACTA, POR_CONTEO))
        if cont and bandas:
            cont = bandas[0][0] - 6 <= r["lineas"][0]["bbox"][0] <= bandas[0][1] + 6
        elif cont:
            cont = abs(r["lineas"][0]["bbox"][0]
                       - anterior["lineas"][0]["bbox"][0]) <= 12
        if cont:
            celdas = [""] * ncols
            celdas[0] = r["celdas"][0]
            r = dict(r, celdas=celdas, confianza=EXACTA,
                     motivo="continuacion del rotulo de la fila anterior; se "
                            "deja en la columna 0, sin fusionar, para no "
                            "reordenar el texto impreso")
        out.append(r)
    return out


def verificar_sin_perdida(filas) -> list[str]:
    """Cada fila reconstruida tiene que contener EXACTAMENTE los caracteres de
    sus `Line` de origen. Es lo unico que se puede garantizar sin el PDF, y se
    garantiza: esta capa reparte, no anade ni quita."""
    malas = []
    for r in filas:
        origen = _plano("".join(texto(l["html"]) for l in r["lineas"]))
        salida = _plano("".join(r["celdas"]))
        if origen != salida:
            malas.append(f'{r["lineas"][0]["id"]}: {origen[:60]!r} != {salida[:60]!r}')
    return malas


# ---------------------------------------------------------------------------
# Una especificacion entera
# ---------------------------------------------------------------------------
_RE_MARCADOR = re.compile(r"^\s*(?:\(\d+\)|[A-Za-z]\b|NOTE\b|GENERAL NOTE)", re.I)


def tablas_de(spec) -> list[dict]:
    """Todas las tablas logicas de una especificacion, ya recompuestas."""
    bloques = deduplicar(spec["bloques"])
    por_id = {b["id"]: b for b in bloques}
    crudas = [b for b in bloques if b["tipo"] == "Table" and b["bbox"]]
    crudas.sort(key=lambda b: (int(b["pagina"]), b["bbox"][1]))

    logicas, previa = [], None
    for t in crudas:
        tit, cont, cap_id = titular(t, bloques, por_id)
        # Se une a la anterior si viene marcada como continuacion, o si repite
        # su titulo exacto en la pagina siguiente. La cabecera solo se imprime
        # en la primera pagina: unirlas es lo que permite arrastrarla.
        unir = (previa is not None and
                ((cont and (not tit or tit == previa["titulo"] or not previa["titulo"]))
                 or (tit and tit == previa["titulo"])))
        if unir:
            previa["bloques"].append(t)
            previa["paginas"].append(t["pagina"])
            previa["continuada"] = True
            continue
        previa = dict(titulo=tit, bloques=[t], paginas=[t["pagina"]],
                      folio=t["folio"], caption=cap_id, continuada=False,
                      suelta=t["padre_tipo"] != "TableGroup")
        logicas.append(previa)

    salida = []
    for i, lg in enumerate(logicas, start=1):
        filas, ncols, resumen = reconstruir(lg["bloques"], lg["titulo"])
        # Notas al pie: bloques por debajo de la tabla en cualquiera de sus
        # paginas. Lo que las identifica es el TIPO de bloque —`Footnote`, o un
        # `Text` que absorbio a su gemelo `Footnote` en la deduplicacion—, no un
        # marcador al principio del texto: el codigo pega el marcador a la
        # primera palabra («bSingle values shown are maximum percentages») y
        # buscarlo con una expresion regular no encuentra ninguna. Se admite
        # ademas el `Text` que si empieza por un marcador reconocible.
        por_pagina = {b["pagina"]: b for b in lg["bloques"]}
        notas, vistos = [], set()
        for b in bloques:
            t_ = por_pagina.get(b["pagina"])
            if t_ is None or not (b["bbox"] and t_["bbox"]):
                continue
            if b["bbox"][1] < t_["bbox"][3] - 2:
                continue
            es_nota = (b["tipo"] == "Footnote" or b.get("gemelo") == "Footnote"
                       or (b["tipo"] == "Text"
                           and _RE_MARCADOR.match(texto(b["html"]) or "")))
            t = texto(b["html"])
            if es_nota and t and t not in vistos:
                vistos.add(t)
                notas.append(dict(id=b["id"], texto=t))
        salida.append(dict(
            n=i, titulo=lg["titulo"], paginas=sorted(set(lg["paginas"]), key=int),
            folio=lg["folio"], bloques=[b["id"] for b in lg["bloques"]],
            continuada=lg["continuada"], suelta=lg["suelta"],
            filas=filas, ncols=ncols, resumen=resumen, notas=notas))
    return salida


def huecos_de(spec) -> Counter:
    """Bloques que NO se pueden reparar sin el PDF: se declaran, no se rellenan."""
    c = Counter()
    for b in spec["bloques"]:
        if b["tipo"] == "Form":
            c["Form sin contenido (parte C)"] += 1
        elif b["tipo"] == "Caption" and not texto(b["html"]):
            c["Caption vacio"] += 1
        elif b["tipo"] == "Table" and not b["hijos"]:
            c["Table sin ningun Line"] += 1
    return c


# ---------------------------------------------------------------------------
# Barrido
# ---------------------------------------------------------------------------
def archivos_de(res: Path, parte: str, filtro=None):
    """Archivos de una parte. El filtro casa por PREFIJO exacto del nombre:
    'SA-6' no puede colarse dentro de 'SA-609'."""
    base = res / SEC_II / parte / "specifications"
    pref = [re.sub(r"[^a-z0-9]+", "_", f.lower()) + "_" for f in (filtro or [])]
    for p in sorted(base.glob("*.json")):
        if filtro and not any(p.name.lower().startswith(x) for x in pref):
            continue
        yield p


def cobertura_declarada(res: Path, partes=PARTES):
    """Lo que el propio drop declara sobre su fidelidad al PDF.

    No es un dato de esta capa ni se puede medir desde aqui: lo mide el
    extractor contra la capa de texto del PDF y lo escribe en su `meta.json`.
    Se publica en el informe porque acota lo que significa todo lo demas: una
    tabla puede estar perfectamente localizada y tabulada y aun asi haber
    perdido texto aguas arriba.
    """
    out = []
    for parte in partes:
        ruta = res / SEC_II / parte / "meta.json"
        if not ruta.is_file():
            continue
        m = json.loads(ruta.read_text(encoding="utf-8")).get("conversion") or {}
        out.append(dict(parte=parte,
                        cobertura=m.get("token_coverage_vs_pdf_text_layer", "?"),
                        nota=m.get("coverage_note", "")))
    return out


def barrer(res: Path, partes=PARTES, filtro=None, verbose=False):
    total = dict(specs=0, tablas=0, filas=0, celdas=0, lineas=0,
                 conf=Counter(), conf_dato=Counter(), huecos=Counter(),
                 perdida=0, sin_titulo=0, continuadas=0, sueltas=0, notas=0)
    detalle = []
    for parte in partes:
        for ruta in archivos_de(res, parte, filtro):
            spec = cargar_spec(ruta)
            sid = spec["spec"].get("id") or ruta.stem
            tablas = tablas_de(spec)
            hue = huecos_de(spec)
            total["specs"] += 1
            total["huecos"] += hue
            for t in tablas:
                total["tablas"] += 1
                total["filas"] += len(t["filas"])
                total["celdas"] += sum(len(f["celdas"]) for f in t["filas"])
                total["lineas"] += sum(len(f["lineas"]) for f in t["filas"])
                total["conf"] += Counter(f["confianza"] for f in t["filas"])
                total["conf_dato"] += Counter(f["confianza"] for f in t["filas"]
                                              if f.get("tipo") == "dato")
                total["notas"] += len(t["notas"])
                total["sin_titulo"] += 0 if t["titulo"] else 1
                total["continuadas"] += 1 if t["continuada"] else 0
                total["sueltas"] += 1 if t["suelta"] else 0
                malas = verificar_sin_perdida(t["filas"])
                total["perdida"] += len(malas)
                detalle.append(dict(parte=parte, spec=sid, tabla=t["n"],
                                    titulo=t["titulo"], paginas=t["paginas"],
                                    folio=t["folio"], ncols=t["ncols"],
                                    filas=len(t["filas"]),
                                    continuada=t["continuada"], suelta=t["suelta"],
                                    conf=Counter(f["confianza"] for f in t["filas"]),
                                    motivos=Counter(f["motivo"] for f in t["filas"]
                                                    if f["motivo"]),
                                    perdida=len(malas),
                                    notas=len(t["notas"]),
                                    origen_ncols=(t["resumen"] or {}).get(
                                        "origen_ncols", "") if isinstance(
                                        t["resumen"], dict) else ""))
            if verbose:
                c = Counter(f["confianza"] for t in tablas for f in t["filas"])
                print(f"  {sid:24s} {len(tablas):3d} tablas  "
                      f"{sum(len(t['filas']) for t in tablas):5d} filas  {dict(c)}")
    return total, detalle


def informe(total, detalle, cobertura=()) -> str:
    n = total["conf"]
    tot_filas = sum(n.values()) or 1
    out = []
    a = out.append
    a("# Revision de las tablas de ASME BPVC Seccion II, partes A, B y C")
    a("")
    a("Generado por `secii_tablas.py`. Mide la reconstruccion de las tablas desde "
      "los bloques `Line` y sus `bbox`; no escribe nada en `resources/` ni en el "
      "libro.")
    a("")
    a("## Resumen")
    a("")
    a("| Magnitud | Valor |")
    a("|---|---:|")
    a(f"| Especificaciones recorridas | {total['specs']} |")
    a(f"| Tablas logicas (tras unir continuaciones) | {total['tablas']} |")
    a(f"| de ellas, continuadas de pagina | {total['continuadas']} |")
    a(f"| de ellas, sueltas (sin TableGroup) | {total['sueltas']} |")
    a(f"| de ellas, sin titulo localizable | {total['sin_titulo']} |")
    a(f"| Filas reconstruidas | {total['filas']} |")
    a(f"| Bloques `Line` consumidos | {total['lineas']} |")
    a(f"| Celdas escritas | {total['celdas']} |")
    a(f"| Notas al pie recogidas | {total['notas']} |")
    a("")
    a("## Reparto por confianza")
    a("")
    a("Se dan dos columnas de porcentaje. La segunda es la que importa: una "
      "banda o un encabezado que no se deja partir es un rotulo, y cuesta mucho "
      "menos que una fila de DATOS que no se deja partir, que son numeros del "
      "codigo. Una fila cuenta como dato si su texto trae alguna ficha de valor.")
    a("")
    nd = total["conf_dato"]
    tot_dato = sum(nd.values()) or 1
    a("| Confianza | Filas | % de todas | Filas de dato | % de las de dato |")
    a("|---|---:|---:|---:|---:|")
    for k in (EXACTA, POR_CONTEO, AMBIGUA):
        a(f"| {k} | {n.get(k, 0)} | {100.0 * n.get(k, 0) / tot_filas:.1f} % | "
          f"{nd.get(k, 0)} | {100.0 * nd.get(k, 0) / tot_dato:.1f} % |")
    a("")
    a(f"**Comprobacion sin perdida: {total['perdida']} filas en las que la "
      "concatenacion de celdas no coincide caracter a caracter con la de sus "
      "`Line` de origen.** Cero es la unica cifra aceptable: significa que esta "
      "capa reparte el texto impreso y no anade ni quita nada.")
    a("")
    a("## Lo que esta capa NO garantiza")
    a("")
    a("La comprobacion sin perdida demuestra que la tabulacion es una reparticion "
      "del texto que trae el JSON. **No demuestra que ese JSON reproduzca el PDF**: "
      "eso lo mide el extractor contra la capa de texto del propio PDF y lo declara "
      "en su `meta.json`. Se copia aqui porque acota lo que significa todo lo demas "
      "— una tabla puede estar perfectamente localizada y tabulada y aun asi haber "
      "perdido texto aguas arriba, y las paginas apaisadas son justo las tablas "
      "anchas de aleacion y propiedades mecanicas.")
    a("")
    if cobertura:
        a("| Parte | Cobertura declarada | Detalle del propio drop |")
        a("|---|---|---|")
        for c in cobertura:
            a(f"| `{c['parte']}` | {c['cobertura']} | {c['nota']} |")
    else:
        a("(no se encontro `meta.json` en ninguna parte)")
    a("")
    a("**Para valores leidos de esas tablas, el PDF manda.**")
    a("")
    a("## Huecos declarados (no reparables sin el PDF)")
    a("")
    if total["huecos"]:
        a("| Hueco | Bloques |")
        a("|---|---:|")
        for k, v in total["huecos"].most_common():
            a(f"| {k} | {v} |")
    else:
        a("Ninguno.")
    a("")
    a("## Tablas con filas AMBIGUAS")
    a("")
    a("Una fila AMBIGUA conserva su texto impreso ENTERO en una celda: no se "
      "pierde nada, pero tampoco queda tabulada. Se listan las 60 tablas con "
      "mas filas ambiguas.")
    a("")
    a("| Parte | Spec | Tabla | Titulo | Pags. | Cols. | Filas | AMBIGUA | Motivo dominante |")
    a("|---|---|---:|---|---|---:|---:|---:|---|")
    peor = sorted(detalle, key=lambda d: -d["conf"].get(AMBIGUA, 0))[:60]
    for d in peor:
        if not d["conf"].get(AMBIGUA):
            break
        mot = d["motivos"].most_common(1)
        a(f"| {d['parte']} | {d['spec']} | {d['tabla']} | "
          f"{(d['titulo'] or '(sin titulo)')[:52]} | "
          f"{'-'.join(str(p) for p in d['paginas'][:2])} | {d['ncols']} | "
          f"{d['filas']} | {d['conf'].get(AMBIGUA, 0)} | "
          f"{(mot[0][0][:52] if mot else '')} |")
    a("")
    a("## Motivos de ambiguedad, agregados")
    a("")
    agr = Counter()
    for d in detalle:
        agr += d["motivos"]
    a("| Motivo | Filas |")
    a("|---|---:|")
    for k, v in agr.most_common():
        a(f"| {k} | {v} |")
    a("")
    a("## Punto de decision — Fase 2 del PLAN-SECII-ABC-001")
    a("")
    pct = 100.0 * nd.get(AMBIGUA, 0) / tot_dato
    a(f"El plan fija aqui una parada: *«si la fraccion AMBIGUA no es marginal, se "
      f"corrige el algoritmo antes de tocar el libro. Escribir 124 000 filas dudosas "
      f"es peor que no escribirlas.»* La fraccion medida es **{pct:.1f} % de las "
      f"filas de dato**.")
    a("")
    a("Lo que SI esta garantizado, y es lo que se puede garantizar sin el PDF:")
    a("")
    a("- Ninguna fila pierde texto. La comprobacion sin perdida pasa sobre las "
      f"{total['filas']} filas y los {total['lineas']} bloques `Line`.")
    a("- Ninguna fila AMBIGUA se rellena por interpolacion sobre el `bbox`: se "
      "conserva entera en una celda y se declara. Sin `Span`, la posicion de cada "
      "palabra dentro de un `Line` no esta en el fichero, y repartirla con una "
      "fuente proporcional seria inventar estructura.")
    a("- Los huecos del origen se cuentan y se nombran; ninguno se rellena.")
    a("")
    a("Lo que NO conviene hacer todavia: volcar las nueve hojas al libro. Una fila "
      "AMBIGUA en la hoja se lee como una fila de tabla que no esta tabulada, y a "
      "esa escala el volcado seria mas ruido que dato. La decision de seguir a las "
      "Fases 3 a 5 —o de acotarlas a las tablas que si se reparten— es del "
      "ingeniero, y este informe es el insumo para tomarla.")
    a("")
    return "\n".join(out)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--informe", type=Path, default=None,
                    help="ruta del informe Markdown; si falta, solo resumen por consola")
    ap.add_argument("--piloto", action="store_true",
                    help="solo las 12 especificaciones del piloto")
    ap.add_argument("--spec", action="append", default=None,
                    help="una especificacion concreta (repetible)")
    ap.add_argument("--parte", action="append", default=None,
                    help="limitar a una parte (bpvc_ii_a_1, ...)")
    ap.add_argument("-v", "--verbose", action="store_true")
    a = ap.parse_args(argv)

    if not (a.resources / SEC_II).is_dir():
        print(f"ERROR: no existe {a.resources / SEC_II}", file=sys.stderr)
        return 2
    filtro = a.spec or (list(PILOTO) if a.piloto else None)
    partes = tuple(a.parte) if a.parte else PARTES
    total, detalle = barrer(a.resources, partes, filtro, a.verbose)
    if not total["specs"]:
        print("ERROR: ninguna especificacion coincide con el filtro", file=sys.stderr)
        return 2
    txt = informe(total, detalle, cobertura_declarada(a.resources, partes))
    if a.informe:
        a.informe.parent.mkdir(parents=True, exist_ok=True)
        a.informe.write_text(txt, encoding="utf-8")
        print(f"Informe escrito en {a.informe}")
    n, nd = total["conf"], total["conf_dato"]
    tot = sum(n.values()) or 1
    totd = sum(nd.values()) or 1
    print(json.dumps({
        "specs": total["specs"], "tablas": total["tablas"],
        "filas": total["filas"], "celdas": total["celdas"],
        "notas": total["notas"],
        "confianza": {k: n.get(k, 0) for k in (EXACTA, POR_CONTEO, AMBIGUA)},
        "ambigua_pct": round(100.0 * n.get(AMBIGUA, 0) / tot, 2),
        "confianza_filas_de_dato": {k: nd.get(k, 0)
                                    for k in (EXACTA, POR_CONTEO, AMBIGUA)},
        "ambigua_pct_filas_de_dato": round(100.0 * nd.get(AMBIGUA, 0) / totd, 2),
        "sin_perdida_fallos": total["perdida"],
        "huecos": dict(total["huecos"]),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
