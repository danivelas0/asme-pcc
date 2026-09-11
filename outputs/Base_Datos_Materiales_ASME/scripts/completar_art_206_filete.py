# -*- coding: utf-8 -*-
"""
completar_art_206_filete - fija en la capa de texto del JSON del Art. 206 las dos
ecuaciones del cateto del filete de extremo, hoy presentes en `resources/` SOLO
como imagen (Figs. 206-3.5-1 y 206-3.5-2).

POR QUE EXISTE ESTE SCRIPT
--------------------------
El texto de 206-3.5 (bloques 44-45) solo enuncia la regla de rama del 1,4x
(filete completo si T_s <= 1,4 T_p; extremos as-is/achaflanados si es mas
grueso). El CATETO en si —cuanto mide la soldadura— esta en las figuras:
  Fig. 206-3.5-1 (bloque 72):  w    = T_s + G       (para T_s <= 1,4 T_p)
  Fig. 206-3.5-2 (bloque 77):  w_max = 1,4 T_p + G   (para T_s > 1,4 T_p, chaflan)
donde G = gap (luz) sleeve-portador, T_p = espesor del tubo portador, T_s =
espesor nominal del sleeve Type B.

La Regla n.1 admite que un dato viva en `resources/` como imagen: el bloque
`figure` con `image:` es parte de `resources/`. Este script recupera las dos
ecuaciones LEYENDO ambos PNG del propio `resources/` (confirmado: imprimen
`T_s + G` y `1,4·T_p + G`) y las fija en `blocks[i].text`, con su procedencia en
`extraction_amendments`. Asi el motor del Art. 206 (Paso 4) puede citarlas de la
capa de texto, trazables, sin reabrir la imagen. No hay ningun vacio real para el
Art. 206 (a diferencia del 212, que si lo tenia en el App. 501).

Es idempotente y defensivo: localiza cada figura por su PNG (subcadena 3_5_1 /
3_5_2), aborta si el bloque ya no apunta a ese PNG, y no pisa un texto distinto
del esperado.

USO
---
    python completar_art_206_filete.py --resources ..\\..\\..\\resources [--dry-run]
"""
from __future__ import annotations

import argparse
import copy
import datetime
import hashlib
import json
import sys
from pathlib import Path

MIRRORS = (
    Path("ASME PCC") / "pcc_2" / "p2_welded_repairs"
    / "art_206_full_encirclement_steel" / "art_206.json",
    Path("ASME PCC") / "pcc_2" / "part_2_welded_repairs"
    / "article_206_full_encirclement_steel"
    / "article_206_full_encirclement_steel.json",
)

# (subcadena del PNG, texto que se fija). Recuperados leyendo ambas figuras.
FIGURAS = [
    ("3_5_1", "w = Ts + G (cateto de filete completo, Ts <= 1.4 Tp) — Fig. 206-3.5-1"),
    ("3_5_2", "w_max = 1.4 Tp + G (cateto maximo, Ts > 1.4 Tp; chaflan opcional) "
              "— Fig. 206-3.5-2"),
]

CLAVE_PROPIA = "extraction_amendments"


def sha256(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def _bloque_figura(blocks, subcadena):
    for i, b in enumerate(blocks):
        if isinstance(b, dict) and subcadena in str(b.get("image", "")):
            return i
    return None


def procesar(ruta_json: Path, hoy: str, dry: bool):
    if not ruta_json.is_file():
        raise SystemExit(f"ERROR: falta {ruta_json}")
    doc = json.loads(ruta_json.read_text(encoding="utf-8"))
    original = copy.deepcopy(doc)
    blocks = doc.get("blocks") or []

    entradas = []
    for subcadena, texto in FIGURAS:
        i = _bloque_figura(blocks, subcadena)
        if i is None:
            raise SystemExit(
                f"ERROR: en {ruta_json.name} ningun bloque apunta a un PNG "
                f"'{subcadena}'. La fuente de esa ecuacion ya no esta donde se "
                f"espera (¿se reorganizo resources/?).")
        b = blocks[i]
        if b.get("type") != "figure":
            raise SystemExit(
                f"ERROR: el bloque {i} de {ruta_json.name} ('{subcadena}') no es "
                f"'figure' sino '{b.get('type')}'. No se toca.")
        png = ruta_json.parent / b["image"]
        if not png.is_file():
            raise SystemExit(
                f"ERROR: el bloque {i} referencia '{b['image']}' pero el PNG no "
                f"existe en {png}. La imagen es la fuente; sin ella no se fija.")
        actual = (b.get("text") or "").strip()
        if actual and actual != texto:
            raise SystemExit(
                f"ERROR: el bloque {i} de {ruta_json.name} ya trae un text "
                f"distinto del esperado: {actual!r}. No se pisa.")
        b["text"] = texto
        entradas.append({"bloque": i, "imagen_fuente": b["image"],
                         "imagen_sha256": sha256(png), "texto_fijado": texto})

    doc[CLAVE_PROPIA] = {
        "fecha": hoy,
        "script": "completar_art_206_filete.py",
        "que": "Se fijan en la capa de texto las dos ec. de cateto del filete de "
               "extremo (w = Ts + G y w_max = 1.4 Tp + G), hasta ahora presentes "
               "solo como imagen (Figs. 206-3.5-1 y 206-3.5-2).",
        "ecuaciones": entradas,
        "procedencia": "Recuperadas leyendo los PNG del propio resources/ (Regla "
                       "n.1: un dato puede vivir en resources/ como imagen). NO se "
                       "uso ningun folio/PDF externo ni memoria del modelo.",
        "alcance": "Solo el text de los dos bloques de figura y este metadato. "
                   "Ningun otro valor del JSON se altera (regla 9 del proyecto).",
    }

    if doc == original:
        print(f"  {ruta_json.name}: sin cambios (ya estaba fijado).")
        return False
    if dry:
        print(f"  {ruta_json.name}: fijaria {len(entradas)} ec. de cateto "
              f"(bloques {[e['bloque'] for e in entradas]}).")
        return False
    ruta_json.write_text(json.dumps(doc, ensure_ascii=False, indent=2),
                         encoding="utf-8")
    print(f"  {ruta_json.name}: cateto fijado en los bloques "
          f"{[e['bloque'] for e in entradas]}.")
    return True


def main(argv=None):
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args(argv)
    if not a.resources.is_dir():
        print(f"ERROR: no existe {a.resources}", file=sys.stderr)
        return 2
    hoy = datetime.date.today().isoformat()
    print("=" * 78)
    print("Art. 206 cateto del filete  w = Ts+G / w_max = 1.4 Tp+G  (de las figuras)")
    print("=" * 78)
    n = 0
    for rel in MIRRORS:
        if procesar(a.resources / rel, hoy, a.dry_run):
            n += 1
    print(f"\n{'(--dry-run: no se ha escrito nada)' if a.dry_run else f'Espejos modificados: {n} de {len(MIRRORS)}.'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
