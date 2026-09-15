# -*- coding: utf-8 -*-
"""
completar_art_210_ce - repara la ecuacion de carbono equivalente del Art. 210
(`para. 210-4.1.1.3`), que en `resources/` llega COLAPSADA y sin imagen.

POR QUE EXISTE ESTE SCRIPT
--------------------------
El bloque `equation` del Art. 210 trae hoy este texto:

    = + + + + + + CE C Mn 6 Cr Mo V 5 Ni Cu 15

Estan todos los simbolos, pero NO la estructura: las barras de fraccion de la
ecuacion son TRAZOS del PDF, no caracteres, asi que la extraccion las perdio y
con ellas la unica informacion que dice que numero divide a que. Tal cual, el
bloque no publica una ecuacion: publica una bolsa de simbolos.

A diferencia de la ec. (2) del Art. 212, aqui NO hay imagen a la que recurrir
(el bloque no trae `image`), asi que es el Caso A de la skill `motor_pcc2`:
extraccion defectuosa que se repara desde el PDF del codigo, por script, con su
procedencia. El PDF NO se versiona (copyright de ASME) y NO alimenta al motor:
alimenta a este script, que escribe el JSON; el motor lee el JSON.

COMO SE RECUPERA LA ESTRUCTURA
------------------------------
No se teclea de memoria. Se DERIVA de la geometria de la pagina, que es lo unico
que la publica: cada barra de fraccion es un trazo horizontal, y la posicion
vertical de cada simbolo respecto de esa barra dice si es numerador o
denominador. `reconstruir_ce()` hace exactamente eso y devuelve la ecuacion
armada; el resultado se contrasta ademas contra `EC_ESPERADA`, de modo que si
ASME recompone la pagina en una edicion futura el script ABORTA en vez de
escribir una ecuacion que ya no es la impresa.

Esto importa y no es celo de mas: el Art. 208-3.4 publica OTRA definicion de CE
-`C + (Mn + Si)/6 + ...`, con silicio- para el mismo codigo. Escribir "la
formula de CE que uno recuerda" habria metido la del 208 en el 210, o al reves.

USO
---
    python completar_art_210_ce.py --resources ..\\..\\..\\resources \\
        --pdf "C:\\...\\ASME PCC-2 REPAIR OF PRESSURE EQUIPMENT AND PIPING.pdf"
        [--dry-run]

Requiere PyMuPDF (`import pymupdf`), que es con lo que se extrajo el B31.3 de
este mismo repositorio (ver tools/b31_3_extractor).
"""
from __future__ import annotations

import argparse
import copy
import datetime
import hashlib
import json
import sys
from pathlib import Path

# Los dos espejos del Art. 210 en resources/.
MIRRORS = (
    Path("asme_pcc") / "pcc_2" / "p2_welded_repairs"
    / "art_210_service_welding_onto" / "art_210.json",
    Path("asme_pcc") / "pcc_2" / "part_2_welded_repairs"
    / "article_210_in_service_welding_onto_carbon"
    / "article_210_in_service_welding_onto_carbon.json",
)

# Anclas de la pagina y de la banda de la ecuacion dentro de ella. Son texto
# impreso, no numeros de pagina: un folio se desplaza entre ediciones, un
# epigrafe no.
ANCLA_PAGINA = "210-4.1.1.3 Carbon Equivalence"
ANCLA_INICIO = "in weight percent amounts"
ANCLA_FIN = "210-4.1.1.4"

# La ecuacion, tal como tiene que salir de la geometria. NO es la fuente del
# dato -la fuente es el PDF- sino el seguro contra un cambio de maquetacion.
EC_ESPERADA = "CE = C + Mn/6 + (Cr + Mo + V)/5 + (Ni + Cu)/15"

# Firma del bloque roto: los simbolos que sí llegaron, en el orden en que
# llegaron. Identifica el bloque sin depender de su indice.
FIRMA_COLAPSADA = ("CE", "Mn", "Cr", "Mo", "Ni", "Cu", "15")

CLAVE_PROPIA = "extraction_amendments"
SCRIPT = "completar_art_210_ce.py"


def sha256(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


# --------------------------------------------------------------------------
# Reconstruccion desde la geometria del PDF
# --------------------------------------------------------------------------
def reconstruir_ce(pdf: Path):
    """La ecuacion de CE del Art. 210, armada desde las posiciones del PDF.

    Devuelve `(ecuacion, n_pagina)`. Aborta si la pagina, la banda o las barras
    de fraccion no estan donde el metodo las supone: es preferible no escribir
    nada a escribir una ecuacion mal armada.
    """
    try:
        import pymupdf
    except ImportError:                                   # pragma: no cover
        raise SystemExit(
            "ERROR: hace falta PyMuPDF para reconstruir la ecuacion desde el "
            "PDF (pip install pymupdf). Es la misma dependencia con la que se "
            "extrajo el B31.3 de este repositorio.")

    doc = pymupdf.open(pdf)
    pagina = next((i for i in range(doc.page_count)
                   if ANCLA_PAGINA in doc[i].get_text()), None)
    if pagina is None:
        raise SystemExit(
            f"ERROR: no se encuentra {ANCLA_PAGINA!r} en {pdf.name}. ¿Es el PDF "
            f"de ASME PCC-2? No se escribe nada.")
    page = doc[pagina]

    ini, fin = page.search_for(ANCLA_INICIO), page.search_for(ANCLA_FIN)
    if not ini or not fin:
        raise SystemExit(
            f"ERROR: en la pagina {pagina} no se acotan las anclas de la banda "
            f"de la ecuacion ({ANCLA_INICIO!r} / {ANCLA_FIN!r}).")
    y0, y1 = ini[0].y1, fin[0].y0

    # La ecuacion va en la columna derecha; el parrafo que la anuncia, en la
    # izquierda. COL_SPLIT separa las dos columnas de la maqueta de ASME.
    col_split = page.rect.width / 2
    fichas = []
    for b in page.get_text("dict")["blocks"]:
        for linea in b.get("lines", []):
            for s in linea.get("spans", []):
                t = s["text"].strip()
                x0, sy0, x1, sy1 = s["bbox"]
                if t and x0 > col_split and y0 - 2 <= (sy0 + sy1) / 2 <= y1 + 2:
                    fichas.append({"t": t, "x": (x0 + x1) / 2, "y": (sy0 + sy1) / 2,
                                   "x0": x0, "x1": x1})

    barras = [d["rect"] for d in page.get_drawings()
              if y0 - 4 <= d["rect"].y0 <= y1 + 4
              and d["rect"].height < 3 and d["rect"].width > 3
              and d["rect"].x0 > col_split]
    if not barras:
        raise SystemExit(
            f"ERROR: no se detecta ninguna barra de fraccion en la banda de la "
            f"ecuacion (pagina {pagina}). Sin las barras no hay estructura que "
            f"recuperar, y no se inventa.")
    barras.sort(key=lambda r: r.x0)

    # Cada ficha cae en la barra cuyo rango horizontal la contiene; lo que no
    # cae en ninguna es la linea base (CE, =, y los + que separan terminos).
    usadas, fracciones = set(), []
    for r in barras:
        num, den = [], []
        for i, f in enumerate(fichas):
            if r.x0 - 1 <= f["x"] <= r.x1 + 1:
                (num if f["y"] < r.y0 else den).append(f)
                usadas.add(i)
        if not num or not den:
            raise SystemExit(
                f"ERROR: la barra en x={r.x0:.1f}..{r.x1:.1f} no tiene "
                f"numerador y denominador completos. No se arma la ecuacion.")
        num.sort(key=lambda f: f["x"])
        den.sort(key=lambda f: f["x"])
        arriba = " ".join(f["t"] for f in num)
        abajo = " ".join(f["t"] for f in den)
        # Un numerador con mas de un simbolo necesita parentesis: (Cr + Mo + V)/5
        # NO es Cr + Mo + V/5, y la diferencia es el dato entero.
        if len(num) > 1:
            arriba = f"({arriba})"
        fracciones.append({"t": f"{arriba}/{abajo}", "x": (r.x0 + r.x1) / 2})

    base = [f for i, f in enumerate(fichas) if i not in usadas]
    piezas = sorted(base + fracciones, key=lambda f: f["x"])
    ecuacion = " ".join(f["t"] for f in piezas)
    return ecuacion, pagina


# --------------------------------------------------------------------------
# Escritura en resources/
# --------------------------------------------------------------------------
def _bloque_ce(blocks):
    """Indice del bloque `equation` de CE, hallado por su firma, no por indice."""
    for i, b in enumerate(blocks):
        if not isinstance(b, dict) or b.get("type") != "equation":
            continue
        t = b.get("text") or ""
        if all(s in t for s in FIRMA_COLAPSADA):
            return i
    return None


def _bloque_ya_fijado(blocks):
    for i, b in enumerate(blocks):
        if (isinstance(b, dict) and b.get("type") == "equation"
                and (b.get("text") or "").strip() == EC_ESPERADA):
            return i
    return None


def procesar(ruta_json: Path, ecuacion: str, pdf_rel: str, pdf_sha: str,
             pagina: int, hoy: str, dry: bool) -> bool:
    if not ruta_json.is_file():
        raise SystemExit(f"ERROR: falta {ruta_json}")
    doc = json.loads(ruta_json.read_text(encoding="utf-8"))
    original = copy.deepcopy(doc)
    blocks = doc.get("blocks") or []

    i = _bloque_ce(blocks)
    if i is None:
        if _bloque_ya_fijado(blocks) is not None:
            print(f"  {ruta_json.name}: sin cambios (ya estaba reparada).")
            return False
        raise SystemExit(
            f"ERROR: en {ruta_json.name} no hay ningun bloque `equation` con la "
            f"firma de la ecuacion de CE {FIRMA_COLAPSADA}. El script NO escribe: "
            f"la extraccion ya no esta donde se espera.")

    blocks[i]["text"] = ecuacion
    doc[CLAVE_PROPIA] = _amendment(hoy, i, pdf_rel, pdf_sha, pagina, ecuacion,
                                   doc.get(CLAVE_PROPIA))

    if doc == original:
        print(f"  {ruta_json.name}: sin cambios (ya estaba reparada).")
        return False
    if dry:
        print(f"  {ruta_json.name}: repararia el bloque {i}.")
        return False
    ruta_json.write_text(json.dumps(doc, ensure_ascii=False, indent=2),
                         encoding="utf-8")
    print(f"  {ruta_json.name}: ecuacion de CE reparada en el bloque {i}.")
    return True


def _amendment(hoy, idx, pdf_rel, pdf_sha, pagina, ecuacion, previo):
    entrada = {
        "fecha": hoy,
        "script": SCRIPT,
        "que": "Se repara la ecuacion de carbono equivalente del para. 210-4.1.1.3, "
               "que la extraccion original entrego colapsada (todos los simbolos, "
               "ninguna barra de fraccion): las barras son trazos del PDF, no "
               "caracteres.",
        "bloque": idx,
        "texto_previo": "= + + + + + + CE C Mn 6 Cr Mo V 5 Ni Cu 15",
        "texto_fijado": ecuacion,
        "pdf_fuente": pdf_rel,
        "pdf_sha256": pdf_sha,
        "pdf_pagina_0based": pagina,
        "procedencia": "Estructura DERIVADA de la geometria de la pagina (posicion "
                       "de cada simbolo respecto de cada trazo horizontal), no "
                       "tecleada de memoria. El PDF no se versiona (copyright ASME) "
                       "y no alimenta ningun motor: alimenta a este script, que "
                       "escribe el JSON (Regla nº 1).",
        "nota": "El Art. 208-3.4 publica OTRA definicion de CE -con Si en el termino "
                "de Mn- para el mismo codigo. Son distintas a proposito; no se "
                "copia una en la otra (regla 9).",
        "alcance": "Solo el bloque de la ecuacion de CE y este metadato. Ningun otro "
                   "valor del JSON se altera.",
    }
    if isinstance(previo, list):
        return [p for p in previo if p.get("script") != SCRIPT] + [entrada]
    if isinstance(previo, dict) and previo.get("script") != SCRIPT:
        return [previo, entrada]
    return entrada


def main(argv=None):
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--pdf", required=True, type=Path,
                    help="PDF de ASME PCC-2 (no se versiona; solo corrige el JSON)")
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args(argv)

    if not a.resources.is_dir():
        print(f"ERROR: no existe {a.resources}", file=sys.stderr)
        return 2
    if not a.pdf.is_file():
        print(f"ERROR: no existe el PDF {a.pdf}", file=sys.stderr)
        return 2

    ecuacion, pagina = reconstruir_ce(a.pdf)
    print("=" * 78)
    print("Art. 210  para. 210-4.1.1.3  —  carbono equivalente")
    print("=" * 78)
    print(f"  PDF                   : {a.pdf.name}")
    pdf_sha = sha256(a.pdf)
    print(f"  PDF sha256            : {pdf_sha}")
    print(f"  pagina (0-based)      : {pagina}")
    print(f"  reconstruida          : {ecuacion}")
    if ecuacion != EC_ESPERADA:
        print(f"  esperada              : {EC_ESPERADA}", file=sys.stderr)
        raise SystemExit(
            "ERROR: la ecuacion reconstruida no coincide con la esperada. Puede "
            "ser otra edicion del codigo o una maquetacion distinta. NO se "
            "escribe: revise la pagina antes de actualizar EC_ESPERADA.")
    print("  coincide con EC_ESPERADA: si\n")

    hoy = datetime.date.today().isoformat()
    cambios = 0
    for rel in MIRRORS:
        if procesar(a.resources / rel, ecuacion, a.pdf.name, pdf_sha, pagina,
                    hoy, a.dry_run):
            cambios += 1
    if a.dry_run:
        print("\n(--dry-run: no se ha escrito nada)")
    else:
        print(f"\nEspejos modificados: {cambios} de {len(MIRRORS)}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
