# -*- coding: utf-8 -*-
"""
db_lib — utilidades comunes para la construccion de las bases de datos de
materiales del Motor de Calculo ASME PCC.

Fuente unica de verdad: carpeta resources/ (regla 5 de knowledge/claude.md).
Ningun valor normativo se escribe desde memoria del modelo: todo procede de
los JSON extraidos de los codigos.

Unidades: se cargan tal como estan impresas en cada edicion (SI y US son
extracciones independientes, no conversiones).

El libro no usa ninguna funcion de matriz dinamica: toda la logica es
INDEX / MATCH / OFFSET / COUNTIF, compatible con cualquier version de Excel.
Es la regla 1 de diseno del libro y no caduca con la revision, asi que no
lleva numero de revision al lado.
"""
from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

# --- Normalizacion de texto -------------------------------------------------
_DASHES = dict.fromkeys(map(ord, "‐‑‒–—―−⁃"), "-")


def norm_dash(s):
    """Homogeneiza guiones tipograficos (ASME imprime en-dash en SA-516)."""
    if s is None or not isinstance(s, str):
        return s
    return s.translate(_DASHES)


def clean(s):
    """Limpia elipsis ASME y espacios; conserva el texto tal como esta impreso."""
    if s is None:
        return None
    if not isinstance(s, str):
        return s
    s = s.replace("…", " ").replace("...", " ")
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s or None


def txt(s):
    """Representacion de texto para claves (limpia + guiones normalizados)."""
    v = clean(norm_dash(s))
    return "" if v is None else str(v)


def search_key(*parts) -> str:
    return " ".join(txt(p) for p in parts if txt(p)).upper()


# --- Carga de JSON ----------------------------------------------------------
class Resources:
    def __init__(self, root: str | Path):
        self.root = Path(root)
        self._cache: dict[str, dict] = {}

    def load(self, rel: str) -> dict:
        if rel not in self._cache:
            with open(self.root / rel, encoding="utf-8") as fh:
                self._cache[rel] = json.load(fh)
        return self._cache[rel]

    def rows(self, rel: str) -> list[dict]:
        return self.load(rel).get("rows", [])


# --- Claves -----------------------------------------------------------------
def build_material_id(parts) -> str:
    return " | ".join(txt(p) for p in parts if txt(p) not in ("", "None"))


def make_unique(ids: list[str], line_nos: list) -> tuple[list[str], list[tuple]]:
    seen: dict[str, int] = {}
    out, collisions = [], []
    for i, mid in enumerate(ids):
        if mid in seen:
            seen[mid] += 1
            ln = line_nos[i] if i < len(line_nos) else seen[mid]
            new = f"{mid} #{ln}"
            k = 2
            while new in seen:
                new = f"{mid} #{ln}.{k}"
                k += 1
            collisions.append((mid, new))
            seen[new] = 1
            out.append(new)
        else:
            seen[mid] = 1
            out.append(mid)
    return out, collisions


def disambiguate(ids, tiers, line_nos):
    """Desambiguacion por niveles, auditable y trazable al codigo impreso.

    ids   : clave base (spec | grado | forma | UNS | clase | tamano | P-No).
    tiers : lista de (etiqueta, valores_por_fila). Se anaden en orden solo a las
            claves que siguen repetidas: primero las NOTAS del codigo (que es lo
            que realmente separa esas filas: condiciones de soldadura, tratamiento
            termico o servicio), luego resistencia y fluencia minimas impresas.
    line_nos : ultimo recurso, el bloque y numero de linea impresos.
    """
    from collections import Counter
    stats, cur = {}, list(ids)
    for label, values in tiers:
        cnt = Counter(cur)
        nxt, n = [], 0
        for mid, v in zip(cur, values):
            if cnt[mid] > 1 and txt(v):
                nxt.append(f"{mid} | {label} {txt(v)}")
                n += 1
            else:
                nxt.append(mid)
        stats[label] = n
        cur = nxt
    final, coll = make_unique(cur, line_nos)
    stats["linea"] = len(coll)
    return final, stats


def bilingual_key(spec, grade, form, uns, cls, pno, ordinal) -> str:
    """Clave de identidad independiente del sistema de unidades.

    Enlaza la fila metrica con su homologa en la edicion U.S. Customary. No
    incluye el tamano (se imprime en mm o en pulgadas segun la edicion); en su
    lugar usa el ORDINAL de la fila dentro del grupo, porque ambas ediciones
    listan los mismos materiales en el mismo orden. La paridad de cardinalidad
    entre ediciones se verifica y se reporta.
    """
    base = build_material_id([spec, grade, form, uns, cls,
                              f"P-{pno}" if pno not in (None, "") else None])
    return f"{base} #{ordinal}"


# --- Utilidades numericas ---------------------------------------------------
def temp_to_number(t):
    if t is None:
        return None
    if isinstance(t, (int, float)):
        return t
    s = str(t).replace(",", "").replace("−", "-").replace("–", "-").strip()
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return None


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").replace("−", "-").strip()
    try:
        return float(s)
    except ValueError:
        return None


def sort_key(spec, form, mid):
    """Orden de la base: por especificacion, forma de producto y clave.
    Es lo que hace contiguos los bloques que alimentan las listas en cascada."""
    return (txt(spec).upper(), txt(form).upper(), txt(mid).upper())


# ---------------------------------------------------------------------------
# Familia de material (agrupacion de navegacion)
# ---------------------------------------------------------------------------
# ATENCION: la familia NO es un dato normativo impreso en el codigo. Es una
# agrupacion derivada, con el unico fin de acortar la lista desplegable de
# composicion nominal. No interviene en ningun calculo: el material sigue
# identificandose por especificacion, grado, forma, UNS, clase y tamano tal
# como los imprime la tabla. Se deriva del prefijo UNS (senal fuerte, asignada
# por SAE/ASTM) y, en su defecto, de la composicion nominal impresa.
FAM_CARBONO = "Acero al carbono"
FAM_BAJA_AL = "Acero de baja aleacion"
FAM_NI_CRIO = "Acero al niquel (criogenico)"
FAM_INOX = "Acero inoxidable"
FAM_FUNDICION = "Fundicion"
FAM_NIQUEL = "Aleacion de niquel"
FAM_ALUMINIO = "Aluminio y aleaciones"
FAM_COBRE = "Cobre y aleaciones"
FAM_TITANIO = "Titanio y aleaciones"
FAM_CIRCONIO = "Circonio y aleaciones"
FAM_OTROS = "Otros metales"

_UNS_PREFIX = [
    ("S", FAM_INOX), ("N", FAM_NIQUEL), ("C", FAM_COBRE),
    ("A", FAM_ALUMINIO), ("F", FAM_FUNDICION), ("J", FAM_FUNDICION),
]


def familia_material(uns, comp, spec=None) -> str:
    u = txt(uns).upper().replace(" ", "")
    c = txt(comp).lower()
    sp = txt(spec).upper()
    # 1) UNS: la letra inicial identifica el grupo asignado por SAE/ASTM
    if u:
        if u.startswith("R5"):
            return FAM_TITANIO
        if u.startswith("R6"):
            return FAM_CIRCONIO
        if u[0] in ("K", "G", "D", "H"):
            pass                                   # aceros: se afina por texto
        else:
            for pref, fam in _UNS_PREFIX:
                if u.startswith(pref):
                    return fam
    # 2) texto de la composicion nominal impresa
    if any(k in c for k in ("titanium", "ti-")):
        return FAM_TITANIO
    if "zirconium" in c or "zr" == c[:2]:
        return FAM_CIRCONIO
    if any(k in c for k in ("cast iron", "gray", "ductile", "malleable")):
        return FAM_FUNDICION
    if any(k in c for k in ("ni-fe-cr", "ni-cr", "ni-mo", "ni-cu", "nickel")) and \
            "steel" not in c:
        return FAM_NIQUEL
    if c.startswith("al") or "aluminum" in c:
        return FAM_ALUMINIO
    if any(k in c for k in ("cu-", "copper", "brass", "bronze")):
        return FAM_COBRE
    if any(k in c for k in ("cr-ni", "ni-cr", "18cr", "16cr", "20cr", "22cr", "25cr",
                            "austenit", "stainless")) or \
            any(k in c for k in ("12cr", "13cr", "15cr", "17cr", "27cr")):
        return FAM_INOX
    if any(k in c for k in ("9ni", "8ni", "5ni", "31∕2ni", "3 1/2ni", "31/2ni", "21∕4ni")):
        return FAM_NI_CRIO
    if "carbon steel" in c or c in ("carbon steel", "c-mn", "c-si", "c-mn-si"):
        return FAM_CARBONO
    if any(k in c for k in ("cr", "mo", "-v", "mn-", "si-", "ni-")):
        return FAM_BAJA_AL
    if u and u[0] in ("K", "G", "D", "H"):
        return FAM_CARBONO
    if sp.startswith(("B", "SB")):
        return FAM_OTROS
    return FAM_OTROS
