# -*- coding: utf-8 -*-
"""
completar_art_212_eq2 - fija en la capa de texto del JSON del Art. 212 la ec. (2)
`F_LP = P*Dm/4`, hoy presente en `resources/` SOLO como imagen.

POR QUE EXISTE ESTE SCRIPT
--------------------------
El Art. 212 publica dos fuerzas de presion: la circunferencial ec. (1)
`F_CP = P*Dm/2` (que si esta en la capa de texto del JSON) y la longitudinal
ec. (2) `F_LP = P*Dm/4`. La ec. (2) NO llego a la capa de texto: vive dentro del
diagrama de la Fig. 212-3.2 (bloque `figure` nº 33 de los dos espejos), cuyo
`text` esta vacio y cuyo `image` apunta al PNG del diagrama.

La Regla nº 1 del proyecto obliga a que todo dato que alimente un motor salga de
`resources/`, y admite explicitamente que un dato viva ahi COMO IMAGEN: el bloque
`figure`/`equation` con `image:` es parte de `resources/`. Este script recupera la
ec. (2) LEYENDO ese PNG del propio `resources/` (no un folio externo) y la fija en
`blocks[i].text`, con su procedencia en `extraction_amendments`. Asi el motor del
Art. 212 puede citar la ec. (2) de la capa de texto, trazable, sin reabrir la
imagen.

El contenido de la imagen ya se confirmo leyendola: imprime literalmente
`F_LP = P*Dm/4` rotulada `(2)`, junto al diagrama de F_CP y F_LP.

Solo se toca el bloque de la ec. (2) y se anade un `extraction_amendments`.
Ningun otro valor del JSON cambia. Es idempotente y defensivo: aborta si el
bloque ya no apunta al PNG esperado (defensa ante una reorganizacion de
`resources/`).

USO
---
    python completar_art_212_eq2.py --resources ..\\..\\..\\resources [--dry-run]
"""
from __future__ import annotations

import argparse
import copy
import datetime
import hashlib
import json
import sys
from pathlib import Path

# Los dos espejos del Art. 212 en resources/. La subcadena identifica el bloque
# `figure` de la ec. (2) por su PNG (el mismo diagrama en las dos ramas, con
# nombre de archivo distinto por como se nombro la extraccion de cada espejo).
MIRRORS = (
    Path("ASME PCC") / "pcc_2" / "p2_welded_repairs"
    / "art_212_fillet_welded_patches" / "art_212.json",
    Path("ASME PCC") / "pcc_2" / "part_2_welded_repairs"
    / "article_212_fillet_welded_patches"
    / "article_212_fillet_welded_patches.json",
)
SUBCADENA_PNG = "diagram_3_2"

# El texto que se recupera de la imagen. Plano (la capa de texto del JSON no
# lleva subindices): F_LP = P*Dm/4, rotulo (2). El middot separa P de Dm.
TEXTO_EC2 = "F_LP = P·Dm/4 (2)"

CLAVE_PROPIA = "extraction_amendments"


def sha256(ruta: Path) -> str:
    return hashlib.sha256(ruta.read_bytes()).hexdigest()


def _bloque_ec2(blocks):
    """Indice del bloque figure de la ec. (2) (el que apunta al diagrama)."""
    for i, b in enumerate(blocks):
        if isinstance(b, dict) and SUBCADENA_PNG in str(b.get("image", "")):
            return i
    return None


def procesar(ruta_json: Path, hoy: str, dry: bool):
    if not ruta_json.is_file():
        raise SystemExit(f"ERROR: falta {ruta_json}")
    doc = json.loads(ruta_json.read_text(encoding="utf-8"))
    original = copy.deepcopy(doc)
    blocks = doc.get("blocks") or []

    i = _bloque_ec2(blocks)
    if i is None:
        raise SystemExit(
            f"ERROR: en {ruta_json.name} ningun bloque apunta a un PNG "
            f"'{SUBCADENA_PNG}'. El script NO escribe: la fuente de la ec. (2) "
            f"ya no esta donde se espera (¿se reorganizo resources/?).")
    b = blocks[i]
    if b.get("type") != "figure":
        raise SystemExit(
            f"ERROR: el bloque {i} de {ruta_json.name} apunta al PNG de la ec. (2) "
            f"pero su type es '{b.get('type')}', no 'figure'. No se toca.")

    png = ruta_json.parent / b["image"]
    if not png.is_file():
        raise SystemExit(
            f"ERROR: el bloque {i} referencia '{b['image']}' pero el PNG no existe "
            f"en {png}. La imagen es la fuente; sin ella no se fija la ec. (2).")
    png_sha = sha256(png)

    # Defensa de idempotencia: el text debe estar vacio (primera corrida) o ya
    # ser exactamente el nuestro (corrida repetida). Cualquier otra cosa se
    # respeta y se avisa: no se pisa un texto que otro proceso haya puesto.
    actual = (b.get("text") or "").strip()
    if actual and actual != TEXTO_EC2:
        raise SystemExit(
            f"ERROR: el bloque {i} de {ruta_json.name} ya trae un text distinto "
            f"del esperado: {actual!r}. No se pisa.")

    b["text"] = TEXTO_EC2
    doc[CLAVE_PROPIA] = _amendment(hoy, i, b["image"], png_sha, doc.get(CLAVE_PROPIA))

    if doc == original:
        print(f"  {ruta_json.name}: sin cambios (ya estaba fijada).")
        return False
    if dry:
        print(f"  {ruta_json.name}: fijaria la ec. (2) en el bloque {i} "
              f"(PNG sha256 {png_sha[:16]}...).")
        return False
    ruta_json.write_text(json.dumps(doc, ensure_ascii=False, indent=2),
                         encoding="utf-8")
    print(f"  {ruta_json.name}: ec. (2) fijada en el bloque {i} "
          f"(PNG sha256 {png_sha[:16]}...).")
    return True


def _amendment(hoy, idx, image_rel, png_sha, previo):
    """Registro de procedencia. Acumula por-archivo si ya habia otros amendments
    (dict con clave por script) sin pisarlos."""
    entrada = {
        "fecha": hoy,
        "script": "completar_art_212_eq2.py",
        "que": "Se fija en la capa de texto la ec. (2) F_LP = P*Dm/4, hasta ahora "
               "presente solo como imagen (diagrama de la Fig. 212-3.2).",
        "bloque": idx,
        "imagen_fuente": image_rel,
        "imagen_sha256": png_sha,
        "texto_fijado": TEXTO_EC2,
        "procedencia": "Recuperada leyendo el PNG del propio resources/ (Regla nº 1: "
                       "un dato puede vivir en resources/ como imagen). NO se uso "
                       "ningun folio/PDF externo ni memoria del modelo.",
        "alcance": "Solo el bloque de la ec. (2) y este metadato. Ningun otro valor "
                   "del JSON se altera (regla 9 del proyecto).",
    }
    if isinstance(previo, dict) and previo.get("script") != "completar_art_212_eq2.py":
        # Convive con amendments de otros scripts: pasa a lista.
        return [previo, entrada] if not isinstance(previo, list) else previo + [entrada]
    return entrada


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
    print("Art. 212 ec. (2)  F_LP = P*Dm/4  —  recuperada de la imagen de resources/")
    print("=" * 78)
    cambios = 0
    for rel in MIRRORS:
        if procesar(a.resources / rel, hoy, a.dry_run):
            cambios += 1
    if a.dry_run:
        print("\n(--dry-run: no se ha escrito nada)")
    else:
        print(f"\nEspejos modificados: {cambios} de {len(MIRRORS)}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
