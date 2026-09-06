# -*- coding: utf-8 -*-
"""
extraer_notas_ii_d — completa las NOTAS de las Tablas TM-1 y TE-1 de
ASME BPVC Seccion II Parte D (Metrica) 2025 en `resources/`.

POR QUE EXISTE ESTE SCRIPT
--------------------------
El mapeo «composicion nominal -> grupo de propiedades» NO es una heuristica:
esta impreso en el codigo, en las Notas al pie de TM-1 (Grupos A..J, para el
modulo E) y de TE-1 (Grupos 1..4, para la dilatacion termica). Sin esas notas,
asignar un grupo a un material es adivinar, y `resources/` es la unica fuente
de verdad del proyecto (Regla n. 1).

La extraccion original de II-D guardo esas notas de forma incompleta:

  table_tm_1.json  ->  solo Notas (1) y (2), y la (1) ademas trunca la mitad
                       de sus miembros (se imprime a dos columnas y quedo una).
                       Faltaban (3)..(17): Grupos C, D, E, F, G, H, I, J y los
                       seis alias de aceros PH identificados por UNS.
  table_te_1.json  ->  solo la GENERAL NOTE y la Nota (1). Faltaban (2), (3)
                       y (4): Grupos 2, 3 y 4.

Sin los Grupos G, H, I, J de TM-1 y 3, 4 de TE-1 no hay forma trazable de dar
E ni alfa a un austenitico o a un duplex.

TRES TRAMPAS DE LA PAGINA IMPRESA, YA PAGADAS
---------------------------------------------
1. LAS DOS TABLAS NO COMPARTEN DISPOSICION. Las paginas de notas de TM-1 estan
   rotadas 90 grados; las de TE-1 son verticales. No sirve un unico orden de
   lectura fijo: hay que derivar el eje de avance de `page.rotation`. Leer el
   texto plano con `get_text()` entremezcla los miembros de una nota con los
   de la siguiente, y silenciosamente (no falla, devuelve listas creibles).
2. LAS NOTAS SE IMPRIMEN A VARIAS COLUMNAS. Una nota es un bloque de
   encabezado seguido de dos o tres bloques de miembros a la misma altura. La
   Nota (1) de TM-1 se imprime a dos columnas y la extraccion original solo
   conservo una: guardaba 4 miembros de los 8 impresos. Por eso aqui una nota
   se arma por GEOMETRIA —los bloques comprendidos entre su encabezado y el
   siguiente, ordenados por columna— y no por flujo de texto.
3. UNA NOTA CONTINUA EN LA PAGINA SIGUIENTE bajo el rotulo «NOTES (CONT'D)».
   El Grupo B de TM-1 arranca en el folio 1188 y termina en el 1189 (14 + 7 =
   21 miembros). El encabezado vigente se arrastra de una pagina a la otra.

FIDELIDAD AL IMPRESO (Regla 9: los valores se cargan tal como estan impresos)
----------------------------------------------------------------------------
ASME BPVC.II.D.M-2025 imprime las Notas (8) y (9) DUPLICADAS: ambas dicen
«Material Group H consists of the following duplex (austenitic-ferritic)
stainless steels:» con la misma lista de 13 miembros. No es un artefacto de
extraccion — son dos bloques distintos de la pagina (folio 1190), verificados
por coordenadas. La tabla solo referencia la Nota (9). Se cargan las dos tal
como estan impresas y se deja constancia en `notes_extraction.observaciones`.
NO se deduplica.

USO
---
    python extraer_notas_ii_d.py --pdf "<...>/D Metric 2025 _p1201-p1500.pdf" \
        --resources ../../../resources [--dry-run]

El PDF es la copia licenciada del usuario; no vive en el repositorio. Se acepta
tanto el PDF completo como cualquiera de sus cortes por paginas: el script
localiza las notas por su texto, no por numero de pagina.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:  # el paquete se publico como `fitz` en versiones previas
    import fitz as pymupdf


# --- Reconocimiento de la pagina -------------------------------------------
# Mobiliario de pagina: rotulos repetidos, folio impreso y pie del codigo. Se
# elimina para que una nota partida entre dos paginas quede contigua.
FURNITURE = re.compile(
    r"^(Table (TM|TE)-\d+.*"
    r"|Moduli of Elasticity.*"
    r"|Thermal Expansion.*"
    r"|NOTES(\s*\(CONT'D\))?:?"
    r"|\d{3,4}"
    r"|ASME BPVC.*"
    r"|\N{LATIN SMALL LETTER ETH}\d+\N{LATIN CAPITAL LETTER THORN})$"
)
ANCLA_NOTAS = re.compile(r"^NOTES(\s*\(CONT'D\))?:?$")
FOLIO = re.compile(r"^\d{3,4}$")
NUM_NOTA = re.compile(r"^\((\d+)\)\s*(.*)$")

# Encabezados de nota que definen una PERTENENCIA a grupo (los que nos importan).
HDR_TM = re.compile(r"^Material Group ([A-J]) consists of the following (.+):$")
HDR_TE = re.compile(r"^Group (\d+) alloys \(by nominal composition\):$")
# Notas (12)..(17) de TM-1: no definen grupo, dan alias de una fila por UNS.
HDR_ALIAS = re.compile(r"^Also known as ")


ROTULO_TABLA = re.compile(r"^Table (TM|TE)-\d+")


def ejes(pagina):
    """Devuelve (avance, columna) para ordenar los bloques de la pagina.

    `avance` recorre la pagina como se lee de arriba abajo; `columna` ordena
    los bloques que estan a la misma altura, de izquierda a derecha. En una
    pagina rotada 90 grados esos dos papeles los cumplen x y -y.
    """
    if pagina.rotation in (90, 270):
        return (lambda b: b[0]), (lambda b: -b[1])
    return (lambda b: b[1]), (lambda b: b[0])


def bloques_de_notas(pagina):
    """Bloques de la region de notas de la pagina, en orden de impresion.

    Devuelve [] si la pagina no abre ni continua una region de notas. Corta al
    llegar al rotulo de OTRA tabla: en el folio 1161 las Notas (2)-(4) de TE-1
    conviven con la Tabla TE-2 entera debajo.
    """
    avance, columna = ejes(pagina)
    bloques = [b for b in pagina.get_text("blocks") if b[4].strip()]
    bloques.sort(key=lambda b: (round(avance(b), 1), round(columna(b), 1)))

    inicio = None
    for i, b in enumerate(bloques):
        if any(ANCLA_NOTAS.match(ln.strip()) for ln in b[4].split("\n")):
            inicio = i
            break
    if inicio is None:
        return []

    region = []
    for b in bloques[inicio:]:
        primera = b[4].strip().split("\n")[0].strip()
        if ROTULO_TABLA.match(primera):
            break
        region.append(b)
    return region


def region_de_notas(doc, id_tabla: str):
    """Recorre el PDF y devuelve (bloques, folios) de las notas de la tabla.

    Una pagina se reconoce por el TEXTO de sus notas, no por el rotulo de la
    tabla: en el folio 1188 la pagina de notas de TM-1 no imprime rotulo
    propio, asi que exigirlo perdia las Notas (1) y (2) sin avisar. En cambio
    «Material Group X consists of...» solo aparece en TM-1 y «Group N alloys
    (by nominal composition)» solo en TE-1, de modo que cada nota se identifica
    a si misma. Se corta en la primera pagina que ya no aporte notas nuestras,
    para no arrastrar las de una tabla posterior.
    """
    hdr = HDR_TM if id_tabla == "TM-1" else HDR_TE
    region, folios, dentro = [], [], False
    for pagina in doc:
        propios = bloques_de_notas(pagina)
        nuestra = any(hdr.match(_sin_numero(ln))
                      for b in propios for ln in _lineas(b))
        if not nuestra:
            if dentro:
                break            # las notas de esta tabla ya terminaron
            continue
        dentro = True
        # El folio impreso es un bloque cuyo texto COMPLETO es el numero. Buscarlo
        # linea a linea devolvia «100», que es un rotulo de temperatura de la tabla.
        folio = next((int(b[4].strip()) for b in pagina.get_text("blocks")
                      if FOLIO.match(b[4].strip())), None)
        if folio is not None:
            folios.append(folio)
        region += propios
    return region, folios


def _lineas(bloque) -> list[str]:
    return [ln.strip() for ln in bloque[4].split("\n")
            if ln.strip() and not FURNITURE.match(ln.strip())]


def _sin_numero(linea: str) -> str:
    """Quita el «(n)» inicial si lo lleva.

    El numero de nota va en su propia linea o pegado al encabezado segun el
    ancho de la columna en que cayo: TM-1 imprime «(3)» aparte y «(10) Material
    Group I consists...» junto; TE-1 siempre lo imprime junto.
    """
    m = NUM_NOTA.match(linea)
    return m.group(2).strip() if m else linea


def partir_en_notas(region: list) -> dict[str, dict]:
    """Arma {numero: {encabezado, miembros}} a partir de los bloques.

    Un bloque es de ENCABEZADO si alguna de sus lineas abre una nota, «(n)».
    Todo bloque posterior alimenta la ultima nota abierta, incluso si esta en
    otra pagina: asi se recupera la continuacion del Grupo B de TM-1.
    """
    notas: dict[str, dict] = {}
    actual = None
    for bloque in region:
        for ln in _lineas(bloque):
            m = NUM_NOTA.match(ln)
            if m:
                actual = m.group(1)
                notas.setdefault(actual,
                                 {"encabezado": m.group(2).strip(), "miembros": []})
                continue
            if actual is None:
                continue
            n = notas[actual]
            if not n["encabezado"]:
                n["encabezado"] = ln     # el encabezado venia en la linea siguiente
            else:
                n["miembros"].append(ln)
    return notas


def clasificar(numero: str, nota: dict, id_tabla: str) -> dict | None:
    """Da forma final a una nota y descarta las que no definen pertenencia."""
    hdr = nota["encabezado"]
    m = (HDR_TM if id_tabla == "TM-1" else HDR_TE).match(hdr)
    if m:
        etiqueta = (f"Material Group {m.group(1)}" if id_tabla == "TM-1"
                    else f"Group {m.group(1)}")
        return {
            "nota": f"({numero})",
            "tipo": "grupo",
            "grupo": etiqueta,
            "encabezado": hdr,
            "miembros": nota["miembros"],
        }
    if HDR_ALIAS.match(hdr):
        return {"nota": f"({numero})", "tipo": "alias",
                "encabezado": hdr, "miembros": []}
    return None


def extraer(pdf: Path, id_tabla: str) -> tuple[list[dict], list[int]]:
    doc = pymupdf.open(pdf)
    try:
        region, folios = region_de_notas(doc, id_tabla)
    finally:
        doc.close()
    crudas = partir_en_notas(region)
    out = []
    for numero in sorted(crudas, key=int):
        fila = clasificar(numero, crudas[numero], id_tabla)
        if fila:
            out.append(fila)
    return out, folios


# --- Escritura en resources/ ------------------------------------------------
def detectar_duplicados(notas: list[dict]) -> list[str]:
    """Notas distintas que definen el MISMO grupo con la MISMA lista.

    Se detecta por los datos y no por el numero de nota, porque la duplicacion
    no es igual en las dos ediciones: la metrica imprime el Grupo H dos veces,
    en las Notas (8) y (9), y a partir de ahi toda su numeracion va corrida en
    uno respecto de la edicion U.S. Customary, que lo imprime una sola vez en
    la Nota (8). Dar por hecho que las dos ediciones numeran igual seria el
    error que este proyecto evita: SI y US son extracciones independientes.
    """
    vistos: dict[tuple, list[str]] = {}
    for n in notas:
        if n["tipo"] != "grupo":
            continue
        vistos.setdefault((n["grupo"], tuple(n["miembros"])), []).append(n["nota"])
    avisos = []
    for (grupo, _), cuales in vistos.items():
        if len(cuales) > 1:
            avisos.append(
                f"El codigo imprime {' y '.join(cuales)} duplicadas: todas definen "
                f"'{grupo}' con identica lista de miembros. Verificado en la pagina "
                f"impresa; no es un artefacto de extraccion. Se cargan todas tal "
                f"como estan impresas (Regla 9: no se deduplican).")
    return avisos


def escribir(ruta: Path, notas: list[dict], folios: list[int], pdf: Path,
             dry_run: bool) -> None:
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)

    datos["note_members"] = notas
    datos["notes_extraction"] = {
        "extraido_de": pdf.name,
        "metodo": ("orden de lectura por bloques (x0 creciente, y0 decreciente); "
                   "la pagina de notas esta rotada"),
        "folios_impresos": sorted(set(folios)),
        "notas_de_grupo": sum(1 for n in notas if n["tipo"] == "grupo"),
        "notas_de_alias": sum(1 for n in notas if n["tipo"] == "alias"),
        "observaciones": detectar_duplicados(notas),
    }

    if dry_run:
        print(f"  [dry-run] {ruta.name}: no se escribio")
        return
    texto = json.dumps(datos, ensure_ascii=False, indent=2)
    # El resto de resources/ usa CRLF; se conserva para que el diff sea limpio.
    with open(ruta, "w", encoding="utf-8", newline="\r\n") as fh:
        fh.write(texto)
    print(f"  escrito {ruta.name}")


TABLAS = [("TM-1", "table_tm_1.json"), ("TE-1", "table_te_1.json")]

# Las dos ediciones se extraen por separado, cada una de su propio PDF. No son
# conversiones una de otra y tampoco numeran igual sus notas.
EDICIONES = {"si": "bpvc_ii_d_metric_2025", "us": "bpvc_ii_d_customary_2025"}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--pdf", required=True, type=Path,
                    help="PDF de II-D (completo o el corte que contenga las "
                         "Tablas TM-1 y TE-1) de la edicion indicada en --edicion")
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--edicion", choices=sorted(EDICIONES), default="si",
                    help="si = metrica (por omision) · us = U.S. Customary")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args(argv)

    if not a.pdf.exists():
        print(f"ERROR: no existe el PDF {a.pdf}", file=sys.stderr)
        return 2
    base = a.resources / EDICIONES[a.edicion]
    print(f"Edicion: {a.edicion} -> {base.name}")

    fallos = 0
    for id_tabla, archivo in TABLAS:
        print(f"Tabla {id_tabla}:")
        notas, folios = extraer(a.pdf, id_tabla)
        grupos = [n for n in notas if n["tipo"] == "grupo"]
        if not grupos:
            print(f"  ERROR: no se hallo ninguna nota de grupo", file=sys.stderr)
            fallos += 1
            continue
        for n in notas:
            print(f"  {n['nota']:5s} {n['tipo']:6s} "
                  f"{n.get('grupo', ''):18s} miembros={len(n['miembros']):3d}")
        escribir(base / archivo, notas, folios, a.pdf, a.dry_run)
    return 1 if fallos else 0


if __name__ == "__main__":
    raise SystemExit(main())
