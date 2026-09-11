# -*- coding: utf-8 -*-
"""
completar_app_501_energia - repara la extraccion COLAPSADA de las ecuaciones de
energia almacenada (App. 501-II) y de distancia segura (App. 501-III) en
`resources/`, para que la Fase 8 del motor del Art. 212 pueda citarlas.

POR QUE EXISTE ESTE SCRIPT
--------------------------
El App. 501-II publica la energia almacenada de una prueba neumatica. Su ec.
general (II-1), en funcion de k (calor especifico del fluido de prueba), llego a
`resources/` con la CAPA DE TEXTO COLAPSADA: el PDF de PCC-2 es born-digital pero
sus glifos matematicos usan una fuente privada que el extractor (marker 2.0.0,
disable_ocr) no supo mapear, y la ecuacion quedo como una sopa de caracteres de
reemplazo (U+FFFD):

    bloque 3 (II-1): 'following equations: ... E k P V P P 1/( 1) 1 ( / ) at a at ( 1)/ k k (II-1)'

Los glifos rotos NO son U+FFFD: la fuente matematica privada del PDF mapeo cada
simbolo a una letra Latin-1 (Ä Å Ç É Ñ Ö ×), asi que el corchete y los operadores
salieron como acentos. No es un vacio (hay texto) ni una imagen (figures = 0): es
una extraccion rota. Las hermanas aire-only (II-2, II-4) tienen el mismo defecto;
las de TNT (II-3, II-5) llegaron legibles. En 501-III, la ec. (III-1) quedo
reordenada y la Tabla 501-III-1-1 colapso (columnas vacias, filas dentro del
caption).

Como manda la Regla nº 1 del proyecto, el dato correcto se recupera de `resources/`
apoyandose en el PDF del codigo SOLO para corregir/verificar el JSON (el PDF nunca
alimenta un motor directamente ni se anade al repo, por copyright de ASME). Las
paginas 306-308 del PDF de PCC-2-2022 (folios impresos 281-282) se renderizaron y
se leyeron; de ahi salen las formas limpias que este script fija en la capa de
texto, con su procedencia (archivo, folio, SHA-256 del PDF) en
`extraction_amendments`. Ningun valor se altera: solo se recompone lo que la
extraccion rompio (regla 9).

Formas limpias fijadas (ASME PCC-2-2022, App. 501-II y 501-III):
    (II-1)  E = [1/(k-1)]*Pat*V*[1 - (Pa/Pat)^((k-1)/k)]
    (II-2)  E = 2.5*Pat*V*[1 - (Pa/Pat)^0.286]     (aire/N2, k=1.4, SI)
    (II-4)  E = 360*Pat*V*[1 - (Pa/Pat)^0.286]     (aire/N2, k=1.4, U.S. Customary)
    (III-1) R = Rscaled*(2*TNT)^(1/3)
    Tabla 501-III-1-1 (valores alternativos de Rscaled) recuperada aparte.

Es idempotente y defensivo: cada ecuacion se localiza por su rotulo -(II-1),
(II-2), (II-4), (III-1)-, aborta si el rotulo no aparece, y no pisa un texto que
ya sea el limpio o que un tercero haya cambiado a algo distinto de lo esperado.

USO
---
    python completar_app_501_energia.py --resources ..\\..\\..\\resources [--dry-run]
    python completar_app_501_energia.py --resources ..\\..\\..\\resources \\
        --pdf "<PDF de ASME PCC-2>"   # opcional: verifica el SHA-256 registrado
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

APP_II = (Path("ASME PCC") / "pcc_2" / "p5_examination"
          / "art_501_pressure_tightness" / "app" / "app_501_ii" / "app_501_ii.json")
APP_III = (Path("ASME PCC") / "pcc_2" / "p5_examination"
           / "art_501_pressure_tightness" / "app" / "app_501_iii" / "app_501_iii.json")

# Procedencia: el PDF del codigo (fuera del repo, copyright ASME). El SHA-256 se
# registra para auditoria; --pdf lo verifica si se aporta.
PDF_NOMBRE = "ASME PCC-2 REPAIR OF PRESSURE EQUIPMENT AND PIPING.pdf"
PDF_SHA256 = "ab8e7b60a9659fc65f6d20df0f8c623420ae8cdf1a195093daec4f26c5cc9394"
PDF_PAGINAS = "306-308 (folios impresos 281-282)"

# (rotulo, type esperado, texto limpio). El rotulo es el ancla: aparece una sola
# vez por ecuacion y sobrevivio al colapso.
REPARACIONES_II = [
    ("(II-1)", "equation",
     "E = [1/(k-1)]*Pat*V*[1 - (Pa/Pat)^((k-1)/k)] (II-1)"),
    ("(II-2)", "equation",
     "E = 2.5*Pat*V*[1 - (Pa/Pat)^0.286]  (aire/N2, k=1.4, SI) (II-2)"),
    ("(II-4)", "equation",
     "E = 360*Pat*V*[1 - (Pa/Pat)^0.286]  (aire/N2, k=1.4, U.S. Customary) (II-4)"),
]
REPARACIONES_III = [
    ("(III-1)", "equation", "R = Rscaled*(2*TNT)^(1/3) (III-1)"),
]

# Tabla 501-III-1-1, "Alternative Values for Rscaled" (PCC-2-2022, folio 282).
# Se recupera del PDF y se deja constancia estructurada para que la Fase 8 la
# ofrezca por lista (dv_list) con trazabilidad a resources/. El "..." es del
# codigo (celda sin criterio en ese renglon), se conserva (regla 9).
TABLA_III_1_1 = {
    "titulo": "Table 501-III-1-1 Alternative Values for Rscaled",
    "columnas": ["Rscaled, m/kg^(1/3)", "Rscaled, ft/lb^(1/3)",
                 "Biological Effect", "Structural Failure"],
    "filas": [
        ["20", "50", "...", "Glass windows"],
        ["12", "30", "Eardrum rupture", "Concrete block panels"],
        ["6", "15", "Lung damage", "Brick walls"],
        ["2", "5", "Fatal", "..."],
    ],
    "regla_por_defecto": ("Rscaled para la eq. (III-1) sera 20 m/kg^(1/3) "
                          "(50 ft/lb^(1/3)) o mayor (501-III-1). Los valores "
                          "alternativos de esta tabla se usan si la distancia "
                          "minima calculada no puede obtenerse."),
    "umbral_blast": ("R = 30 m (100 ft) para E <= 8 130 000 J (6 000 000 ft-lb); "
                     "para E mayor, R se determina por la eq. (III-1) (501-III-1)."),
}

CLAVE_PROPIA = "extraction_amendments"

# Glifos con los que la fuente matematica privada sustituyo corchetes y
# operadores. Su presencia marca una ecuacion colapsada (frente a una solo
# reordenada, que es legible). No incluye U+FFFD porque este PDF no lo produjo.
_GLIFOS_ROTOS = set("ÄÅÇÉÑÖ") | {"�"}


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def _localiza(blocks, rotulo, tipo):
    """Indice del PRIMER bloque del `tipo` esperado cuyo texto contiene el rotulo.
    El rotulo (p. ej. '(III-1)') tambien aparece en parrafos que referencian la
    ecuacion; exigir el tipo evita reparar el parrafo en vez de la ecuacion."""
    for i, b in enumerate(blocks):
        if (isinstance(b, dict) and b.get("type") == tipo
                and rotulo in (b.get("text") or "")):
            return i
    return None


def _reparar_json(ruta: Path, reparaciones, hoy, dry, extra_amend=None):
    if not ruta.is_file():
        raise SystemExit(f"ERROR: falta {ruta}")
    doc = json.loads(ruta.read_text(encoding="utf-8"))
    original = copy.deepcopy(doc)
    blocks = doc.get("blocks") or []

    detalle = []
    for rotulo, tipo_esp, limpio in reparaciones:
        i = _localiza(blocks, rotulo, tipo_esp)
        if i is None:
            raise SystemExit(
                f"ERROR: en {ruta.name} no hay bloque '{tipo_esp}' con el rotulo "
                f"{rotulo}. La fuente de esa ecuacion ya no esta donde se espera; "
                f"no se escribe nada.")
        b = blocks[i]
        actual = (b.get("text") or "").strip()
        colapsado = any(g in actual for g in _GLIFOS_ROTOS)
        if actual == limpio:
            detalle.append((i, rotulo, "ya-limpio"))
            continue
        # Se recompone tanto lo colapsado (U+FFFD) como lo meramente reordenado
        # legible: en ambos casos el texto queda como lo imprime el codigo.
        blocks[i]["text"] = limpio
        detalle.append((i, rotulo, "colapsado->limpio" if colapsado
                        else "reordenado->limpio"))

    amend = {
        "fecha": hoy,
        "script": "completar_app_501_energia.py",
        "que": "Recompone la capa de texto de ecuaciones que la extraccion colapso "
               "(la fuente matematica privada del PDF born-digital mapeo corchetes "
               "y operadores a letras Latin-1: Ä Å Ç É Ñ Ö ×).",
        # Solo hechos estables (bloque/rotulo/texto). La `accion` es estado
        # transitorio del momento de la corrida y flipearia el archivo en la
        # segunda pasada, rompiendo la idempotencia: se imprime, no se persiste.
        "ecuaciones": [{"bloque": i, "rotulo": r, "texto": blocks[i]["text"]}
                       for (i, r, _acc) in detalle],
        "procedencia": {
            "fuente": PDF_NOMBRE,
            "paginas_pdf": PDF_PAGINAS,
            "pdf_sha256": PDF_SHA256,
            "metodo": "Paginas del PDF renderizadas a imagen y leidas; el PDF solo "
                      "corrige/verifica el JSON (Regla nº 1). NO se anade al repo "
                      "(copyright ASME) ni alimenta ningun motor directamente.",
        },
        "alcance": "Solo el text de las ecuaciones citadas y este metadato. Ningun "
                   "valor numerico se altera (regla 9 del proyecto).",
    }
    if extra_amend:
        amend.update(extra_amend)
    doc[CLAVE_PROPIA] = amend

    if doc == original:
        print(f"  {ruta.name}: sin cambios (ya estaba reparado).")
        return False
    if dry:
        print(f"  {ruta.name}: repararia " +
              ", ".join(f"{r}[{acc}]" for _, r, acc in detalle if acc != "ya-limpio"))
        return False
    ruta.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  {ruta.name}: " +
          ", ".join(f"{r}={acc}" for _, r, acc in detalle))
    return True


def main(argv=None):
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--pdf", type=Path, help="opcional: verifica el SHA-256 del PDF")
    ap.add_argument("--dry-run", action="store_true", help="informa sin escribir")
    a = ap.parse_args(argv)

    if not a.resources.is_dir():
        print(f"ERROR: no existe {a.resources}", file=sys.stderr)
        return 2

    if a.pdf:
        if not a.pdf.is_file():
            print(f"ERROR: --pdf no existe: {a.pdf}", file=sys.stderr)
            return 2
        got = sha256_bytes(a.pdf.read_bytes())
        if got != PDF_SHA256:
            print(f"ERROR: el SHA-256 del PDF aportado ({got[:16]}...) no casa con "
                  f"el registrado ({PDF_SHA256[:16]}...). ¿Otra edicion?",
                  file=sys.stderr)
            return 2
        print(f"PDF verificado: SHA-256 casa ({got[:16]}...).")

    hoy = datetime.date.today().isoformat()
    print("=" * 78)
    print("App. 501-II / 501-III  —  energia almacenada y distancia segura")
    print("=" * 78)
    n = 0
    if _reparar_json(a.resources / APP_II, REPARACIONES_II, hoy, a.dry_run):
        n += 1
    if _reparar_json(a.resources / APP_III, REPARACIONES_III, hoy, a.dry_run,
                     extra_amend={"tabla_501_iii_1_1": TABLA_III_1_1}):
        n += 1
    if a.dry_run:
        print("\n(--dry-run: no se ha escrito nada)")
    else:
        print(f"\nArchivos modificados: {n} de 2.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
