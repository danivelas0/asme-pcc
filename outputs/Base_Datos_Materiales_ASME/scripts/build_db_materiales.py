# -*- coding: utf-8 -*-
"""
build_db_materiales.py — Construye las bases de datos de materiales del
Motor de Calculo ASME PCC a partir de los JSON de resources/.

PLAN-DB-MAT-001 · Rev. 2

Cambios de la Rev. 2 frente a la Rev. 1
---------------------------------------
* El libro NO usa ninguna funcion de matriz dinamica (FILTER / SORT / UNIQUE /
  XLOOKUP / VSTACK). Toda la logica es INDEX / MATCH / OFFSET / COUNTIF, que
  funciona en cualquier version de Excel. Era la causa de que los buscadores no
  devolvieran resultados.
* La busqueda por texto libre se sustituye por CASCADA DE LISTAS DESPLEGABLES:
  Spec. No. -> Forma de producto -> Material. La unica celda donde el usuario
  escribe es la temperatura de consulta.
* Las bases se ordenan por (Spec, Forma, clave) para que cada bloque de la
  cascada sea contiguo y las listas se resuelvan con OFFSET/MATCH/COUNTIF.
* Clave bilingue: enlaza cada fila metrica con su homologa U.S. Customary y
  verifica la paridad entre ediciones.

Uso:
    python build_db_materiales.py --resources <ruta> --in <xlsx> --out <xlsx>
"""
from __future__ import annotations

import argparse
import datetime
import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

sys.path.insert(0, str(Path(__file__).resolve().parent))
from db_lib import (Resources, bilingual_key, build_material_id, clean, disambiguate,
                    familia_material, make_unique, num, search_key, sort_key,
                    temp_to_number, txt)

# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
NAVY, BLUE, GREY, YELL = "1F3864", "2F5597", "F2F2F2", "FFF2CC"
TITLE_F = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor=NAVY)
BAND_FILL = PatternFill("solid", fgColor=BLUE)
SRC_F = Font(name="Calibri", size=8, italic=True, color="595959")
HDR_F = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor=BLUE)
DATA_F = Font(name="Calibri", size=9)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LBL_F = Font(name="Calibri", size=10, bold=True)
IN_F = Font(name="Calibri", size=10, bold=True, color="0000FF")
IN_FILL = PatternFill("solid", fgColor=YELL)
OUT_F = Font(name="Calibri", size=12, bold=True, color="006100")

R_TITLE, R_SRC, R_HDR, R_DATA = 1, 2, 3, 4
N_VARIANTES = 25          # filas de la tabla comparativa de cada buscador

META: list[dict] = []
ISSUES: list[str] = []


def g(row: dict, *keys, default=None):
    for k in keys:
        if k in row and row[k] is not None:
            return row[k]
    return default


_SPLIT_RE = re.compile(r"^(material|metal)_(-?\d+)$")


def fix_merged_ident(row: dict):
    """Corrige el artefacto de extraccion 'material_<T>': el nombre del material
    y el valor de la primera columna de temperatura quedaron fusionados
    (p. ej. 'N02200 222' en TM-4 y A-1C). Se separa el ultimo token numerico."""
    for k, v in row.items():
        m = _SPLIT_RE.match(k)
        if m and isinstance(v, str):
            parts = v.rsplit(" ", 1)
            if len(parts) == 2 and num(parts[1]) is not None:
                return parts[0], {m.group(2): num(parts[1])}
            return v, {}
    return None, {}


# ---------------------------------------------------------------------------
# Escritura generica
# ---------------------------------------------------------------------------
def new_sheet(wb, name, title, source):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font, ws["A1"].fill = TITLE_F, TITLE_FILL
    ws["A2"] = source
    ws["A2"].font = SRC_F
    ws.freeze_panes = "A4"
    return ws


def write_headers(ws, cols, temps=None):
    for j, h in enumerate(cols, start=1):
        c = ws.cell(R_HDR, j, h)
        c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    if temps:
        for j, t in enumerate(temps, start=len(cols) + 1):
            c = ws.cell(R_HDR, j, t)
            c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX
            c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[R_HDR].height = 46
    return len(cols)


def write_rows(ws, records, n_ident, temps=None):
    r = R_DATA
    for ident, vals in records:
        for j, v in enumerate(ident, start=1):
            ws.cell(r, j, v).font = DATA_F
        if temps:
            for j, t in enumerate(temps, start=n_ident + 1):
                v = vals.get(t)
                if v is not None:
                    ws.cell(r, j, v).font = DATA_F
        r += 1
    return r - 1


def autosize(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def record_meta(sheet, table_id, source_file, edition, unit, rows, note=""):
    META.append(dict(hoja=sheet, tabla=table_id, archivo=source_file, edicion=edition,
                     sistema=unit, filas=rows, nota=note))


def add_name(wb, name, ref):
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names.add(DefinedName(name, attr_text=ref))


def dv_list(ws, cell, formula):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(ws[cell])


def cascade_formula(val_name, key_name, key_expr):
    """Lista dependiente clasica (sin matrices dinamicas):
    bloque contiguo de <val_name> cuyas filas tienen la clave <key_expr>."""
    return (f"=OFFSET(INDEX({val_name},1),IFERROR(MATCH({key_expr},{key_name},0),1)-1,"
            f"0,MAX(1,COUNTIF({key_name},{key_expr})),1)")


# ---------------------------------------------------------------------------
# Banda compacta: los puntos realmente tabulados, empaquetados a la izquierda
# ---------------------------------------------------------------------------
# Las tablas del codigo dejan huecos INTERIORES (la A-1 no imprime 125 C para
# A106 Gr.B). Sobre la banda impresa un MATCH aproximado caeria en la celda
# vacia. Por eso, junto a la banda impresa —que se conserva como referencia
# auditable, con los encabezados de temperatura del codigo— se escribe una
# banda COMPACTA con solo los puntos que existen, en orden ascendente y sin
# huecos. Toda la consulta trabaja sobre ella con MATCH/INDEX/OFFSET, sin una
# sola formula matricial: asi funciona igual en Excel y en Google Sheets.
def append_packed(ws, records, n_ident, temps, npts_col=None):
    t0 = n_ident + len(temps) + 1
    npack = max((len(v) for _, v in records), default=1) or 1
    v0 = t0 + npack
    for j in range(npack):
        for c, lb in ((t0 + j, f"Tc{j+1}"), (v0 + j, f"Vc{j+1}")):
            cell = ws.cell(R_HDR, c, lb)
            cell.font, cell.fill = HDR_F, HDR_FILL
            ws.column_dimensions[get_column_letter(c)].hidden = True
    for i, (_, vals) in enumerate(records):
        r = R_DATA + i
        for j, t in enumerate(sorted(vals)):
            ws.cell(r, t0 + j, t).font = DATA_F
            ws.cell(r, v0 + j, vals[t]).font = DATA_F
        if npts_col:
            ws.cell(r, npts_col, len(vals)).font = DATA_F
    return t0, v0, npack


def packed_refs(info):
    """Expresiones portables de la fila compacta de un material."""
    sh = info["sheet"]
    t0 = get_column_letter(info["pack_t0"])
    v0 = get_column_letter(info["pack_v0"])
    npc = get_column_letter(info["npts_col"])
    return dict(sheet=sh,
                npts=f"{sh}!${npc}${R_DATA}:${npc}${info['last_row']}",
                t_anchor=f"{sh}!${t0}${R_DATA}",
                v_anchor=f"{sh}!${v0}${R_DATA}")


def packed_rows(refs, fila, npts_cell):
    """(Trow, Vrow) como rangos 1 x n_pts mediante OFFSET."""
    return (f'OFFSET({refs["t_anchor"]},{fila}-1,0,1,MAX(1,{npts_cell}))',
            f'OFFSET({refs["v_anchor"]},{fila}-1,0,1,MAX(1,{npts_cell}))')


def interp_value(t1c, s1c, t2c, s2c, tq, modo):
    """S(T) a partir de las celdas auxiliares. Sin extrapolar: por debajo del
    primer punto devuelve el primero; si no hay punto superior, el ultimo."""
    return (f'=IF({s1c}="","fuera de rango",'
            f'IF({tq}<={t1c},{s1c},'
            f'IF(OR({t2c}="",{s2c}=""),{s1c},'
            f'IF({modo}="Tabulado-conservador",{s2c},'
            f'{s1c}+({s2c}-{s1c})*({tq}-{t1c})/({t2c}-{t1c})))))')


# ---------------------------------------------------------------------------
# Layout unico de las bases indexadas por material
# ---------------------------------------------------------------------------
STRESS_COLS = [
    "material_id", "Tabla", "k0", "k1", "k2", "k3", "k4", "clave_bi", "Familia",
    "Composicion nominal", "Forma de producto", "Spec. No.", "Tipo/Grado",
    "UNS / Alloy", "Clase/Cond./Temple", "Tamano/Espesor", "P-No.", "Group No.",
    "Temp. min. / curva impacto", "Resist. traccion min.", "Fluencia min.",
    "Temp. max. / limite",
    "I", "III", "VIII-1", "VIII-2", "XII", "Grafico presion externa", "Notas",
    "Bloque", "Linea", "n_pts"]
C = {n: i + 1 for i, n in enumerate(STRESS_COLS)}
N_IDENT = len(STRESS_COLS)
CL = {n: get_column_letter(i) for n, i in C.items()}
# Cascada: composicion -> forma -> especificacion -> tipo/grado.
# Las claves k1..k4 son acumulativas; la base se ordena por ellas para que cada
# nivel sea un bloque CONTIGUO y las listas dependientes se resuelvan con
# OFFSET/MATCH/COUNTIF (sin matrices dinamicas).
# Cascada de 5 niveles + variante. La familia agrupa las composiciones nominales
# para que la lista no tenga cientos de entradas.
CASCADA = [("Familia", "k0"), ("Composicion nominal", "k1"),
           ("Forma de producto", "k2"), ("Spec. No.", "k3"), ("Tipo/Grado", "k4")]

SIZE_KEYS = ("size_thickness_mm", "size_thickness_in", "size_mm", "size_in",
             "size_range_in", "size_or_thickness_range_in", "size_range_dia_mm",
             "size_range_dia_in")


def norm_row(row, tag, family):
    """Normaliza una fila de cualquiera de las tablas de material."""
    merged, extra = fix_merged_ident(row)
    d = dict(
        tabla=tag,
        spec=txt(g(row, "spec_no")) or None,
        form=clean(g(row, "product_form")),
        grade=clean(g(row, "type_grade")),
        comp=clean(g(row, "nominal_composition", "material") or merged),
        uns=clean(g(row, "uns_no", "alloy_desig_uns_no")),
        cls=clean(g(row, "class_condition_temper", "class_cond_temper",
                    "class_condi_tion_temper")),
        size=clean(g(row, *SIZE_KEYS)),
        pno=g(row, "p_no", "p_no_5", "p_no_5_7"),
        grp=g(row, "group_no"),
        tmin=g(row, "min_temp_c", "min_temp_f"),
        rm=g(row, "min_tensile_strgth_mpa", "min_tensile_strength_mpa",
             "min_tensile_strength_ksi", "minimum_tensile_strength_ksi"),
        re_=g(row, "min_yield_strgth_mpa", "min_yield_strength_mpa",
              "min_yield_strength_ksi", "minimum_yield_strength_ksi"),
        notes=clean(g(row, "notes")),
        blk=row.get("block"),
        lin=row.get("line_no"),
        extra=extra,
    )
    if family == "b313":
        d["tmax"] = g(row, "max_temp_c", "max_temp_f")
        d["ap"] = (None, None, None, None, None)
        d["chart"] = None
    elif family == "iid":
        v81 = g(row, "viii_1", "viii_1_xii",
                "applicability_and_max_temperature_limits_np_not_permitted_viii_1")
        d["tmax"] = v81
        d["ap"] = (clean(g(row, "i")),
                   clean(g(row, "limits_iii", "iii",
                           "np_not_permitted_spt_supports_only_iii")),
                   clean(v81), clean(g(row, "viii_2")), clean(g(row, "xii")))
        d["chart"] = clean(g(row, "external_pressure_chart_no"))
    else:                                   # propiedades (Tabla U / Y-1)
        d["tmax"] = None
        d["ap"] = (None, None, None, None, None)
        d["chart"] = None
    return d


def build_material_sheet(res, wb, name, title, source, tables, family, base_temp=None):
    """Construye una base indexada por material con el layout unico.

    tables: lista de (etiqueta, ruta_json).
    """
    temps, raw = set(), []
    for tag, f in tables:
        d = res.load(f)
        for t in d.get("columns_values") or []:
            tn = temp_to_number(t)
            if tn is not None:
                temps.add(tn)
        for row in d["rows"]:
            raw.append((tag, row))
    if base_temp is not None:
        temps.add(base_temp)
    temps = sorted(temps)

    recs, dropped = [], 0
    ordinal: dict[str, int] = {}
    for tag, row in raw:
        d = norm_row(row, tag, family)
        vals = {}
        if base_temp is not None:
            v0 = g(row, "min_temp_to_40", "min_temp_to_100")
            if num(v0) is not None:
                vals[base_temp] = num(v0)
        for k, v in (row.get("values") or {}).items():
            t = temp_to_number(k)
            if t is not None and num(v) is not None:
                vals[t] = num(v)
        for k, v in d["extra"].items():
            t = temp_to_number(k)
            if t is not None:
                vals.setdefault(t, v)
        mid = build_material_id([d["spec"], d["grade"], d["form"], d["uns"], d["cls"],
                                 d["size"],
                                 f"P-{d['pno']}" if d["pno"] not in (None, "") else None])
        if not mid and not vals:
            dropped += 1
            continue
        if family == "iid":
            bi_parts = [d["spec"], d["grade"], d["form"], d["uns"], d["cls"],
                        f"P-{d['pno']}" if d["pno"] not in (None, "") else None]
        elif family == "b313":
            # A-1 y su companion A-1C no extraen forma, clase ni UNS de forma
            # homogenea; el enlace entre ediciones se hace con spec + grado + ordinal.
            bi_parts = [d["spec"], d["grade"]]
        else:
            bi_parts = [d["spec"], d["grade"], d["uns"]]
        base_bi = build_material_id(bi_parts)
        ordinal[base_bi] = ordinal.get(base_bi, 0) + 1
        d["bi"] = f"{base_bi} #{ordinal[base_bi]}"
        d["mid_base"] = f"{tag} | {mid}" if family != "prop" else mid
        d["vals"] = vals
        recs.append(d)

    ids, st = disambiguate([d["mid_base"] for d in recs],
                           [("Notas", [d["notes"] for d in recs]),
                            ("Rm", [d["rm"] for d in recs]),
                            ("Re", [d["re_"] for d in recs])],
                           [f"{d['blk']}.{d['lin']}" for d in recs])
    for d, mid in zip(recs, ids):
        d["mid"] = mid


    SIN = "(sin dato)"

    def keys_of(d):
        # El codigo deja campos en blanco en algunas filas. Se sustituyen por una
        # etiqueta seleccionable para que la cascada pueda llegar a esos materiales;
        # las columnas de identificacion conservan el valor tal como esta impreso.
        fa = d["fam"]
        c = txt(d["comp"]) or SIN
        f = txt(d["form"]) or SIN
        sp = txt(d["spec"]) or SIN
        gr = txt(d["grade"]) or SIN
        return (fa, f"{fa}|{c}", f"{fa}|{c}|{f}", f"{fa}|{c}|{f}|{sp}",
                f"{fa}|{c}|{f}|{sp}|{gr}")

    for d in recs:
        d["fam"] = familia_material(d["uns"], d["comp"], d["spec"])
    recs.sort(key=lambda d: (d["fam"].upper(),
                             (txt(d["comp"]) or SIN).upper(),
                             (txt(d["form"]) or SIN).upper(),
                             (txt(d["spec"]) or SIN).upper(),
                             (txt(d["grade"]) or SIN).upper(),
                             txt(d["mid"]).upper()))
    records = []
    for d in recs:
        k0, k1, k2, k3, k4 = keys_of(d)
        d["k0"], d["k1"], d["k2"], d["k3"], d["k4"] = k0, k1, k2, k3, k4
        records.append((
            [d["mid"], d["tabla"], k0, k1, k2, k3, k4, d["bi"], d["fam"],
             d["comp"], d["form"], d["spec"], d["grade"], d["uns"], d["cls"], d["size"],
             d["pno"], d["grp"], d["tmin"], d["rm"], d["re_"], d["tmax"],
             d["ap"][0], d["ap"][1], d["ap"][2], d["ap"][3], d["ap"][4],
             d["chart"], d["notes"], d["blk"], d["lin"], len(d["vals"])],
            d["vals"]))

    ws = new_sheet(wb, name, title, source)
    n = write_headers(ws, STRESS_COLS, temps)
    last = write_rows(ws, records, n, temps)
    autosize(ws, {CL["material_id"]: 46, CL["Tabla"]: 7, CL["clave_bi"]: 40,
                  CL["Familia"]: 26, CL["Composicion nominal"]: 24, CL["Forma de producto"]: 18,
                  CL["Spec. No."]: 12, CL["Tipo/Grado"]: 12, CL["UNS / Alloy"]: 14,
                  CL["Clase/Cond./Temple"]: 16, CL["Tamano/Espesor"]: 16,
                  CL["Notas"]: 16})
    for cn in ("k0", "k1", "k2", "k3", "k4", "clave_bi", "Bloque", "Linea", "n_pts"):
        ws.column_dimensions[CL[cn]].hidden = True
    t0, v0, npack = append_packed(ws, records, n, temps, C["n_pts"])
    ws.auto_filter.ref = f"A{R_HDR}:{get_column_letter(n + len(temps))}{last}"
    ISSUES.append(f"{name}: material_id desambiguados -> " +
                  ", ".join(f"{k}={v}" for k, v in st.items()) +
                  f"; filas sin identificacion ni valores descartadas = {dropped}.")
    return dict(sheet=name, n_ident=n, temps=temps, last_row=last, recs=recs,
                pack_t0=t0, pack_v0=v0, npack=npack, npts_col=C["n_pts"])


# ---------------------------------------------------------------------------
# Bases de esfuerzos y propiedades por material
# ---------------------------------------------------------------------------
APX = "ASME B31/ASME B31.3/APPEX"


def build_b313(res, wb, system):
    si = system == "SI"
    t = [("A-1", f"{APX}/appendix_a/table_a_1.json"),
         ("A-4", f"{APX}/appendix_a/table_a_4.json")] if si else \
        [("A-1C", f"{APX}/appendix_a/table_a_1c.json"),
         ("A-4C", f"{APX}/appendix_a/table_a_4c.json")]
    name = "DB_B31_3" if si else "DB_B31_3C"
    info = build_material_sheet(
        res, wb, name,
        "BASE DE MATERIALES — ASME B31.3-2024, Apendice A · " +
        ("Tablas A-1 y A-4 (SI: MPa, C)" if si else "Tablas A-1C y A-4C (US: ksi, F)"),
        "Fuente: resources/" + t[0][1] + " y " + t[1][1] + " · ASME B31.3-2024 · "
        "Valores tal como estan impresos; celda vacia = elipsis ASME (no admisible). "
        "Filas ordenadas por Spec. No. y forma de producto para alimentar las listas "
        "en cascada.",
        t, "b313", base_temp=40 if si else 100)
    record_meta(name, "A-1 + A-4" if si else "A-1C + A-4C",
                f"{t[0][1]} ; {t[1][1]}", "2024", system,
                info["last_row"] - R_DATA + 1,
                "Esfuerzos admisibles basicos en traccion + perneria.")
    return info


def build_iid(res, wb, system, group):
    si = system == "SI"
    ed = "bpvc_ii_d_metric_2025" if si else "bpvc_ii_d_customary_2025"
    t = ([("1A", f"{ed}/table_1a.json")] if group == "1A"
         else [("1B", f"{ed}/table_1b.json"), ("3", f"{ed}/table_3.json")])
    base = "DB_BPVC_IID" if group == "1A" else "DB_BPVC_IID_B"
    name = base if si else base + ("C" if group == "1A" else "C")
    which = ("Tabla 1A — materiales ferrosos" if group == "1A"
             else "Tablas 1B (no ferrosos) y 3 (perneria)")
    info = build_material_sheet(
        res, wb, name,
        f"BASE DE MATERIALES — ASME BPVC Seccion II-D 2025, {which} " +
        ("(Metrico: MPa, C)" if si else "(U.S. Customary: ksi, F)"),
        "Fuente: " + " , ".join(f"resources/{f}" for _, f in t) + " · Edicion 2025 · "
        "Valores tal como estan impresos. NP = No permitido, SPT = Solo soportes. "
        "Rejilla de temperatura homogenea: la Tabla 1A no imprime 175/225/275 C y la "
        "1B/3 si, por eso van en hojas separadas (mezclarlas dejaria huecos interiores "
        "y romperia la interpolacion).",
        t, "iid")
    record_meta(name, which, " ; ".join(f for _, f in t), "2025", system,
                info["last_row"] - R_DATA + 1, "Maximos esfuerzos admisibles S.")
    return info


def build_prop(res, wb, system, kind):
    si = system == "SI"
    ed = "bpvc_ii_d_metric_2025" if si else "bpvc_ii_d_customary_2025"
    fn = "table_u.json" if kind == "U" else "table_y_1.json"
    tag = "U" if kind == "U" else "Y-1"
    base = "DB_Su" if kind == "U" else "DB_Sy"
    name = base if si else base + "C"
    lbl = ("Resistencia a la traccion Su" if kind == "U" else "Limite de fluencia Sy")
    info = build_material_sheet(
        res, wb, name,
        f"BASE DE PROPIEDADES — {lbl} · ASME BPVC II-D 2025, Tabla {tag} " +
        ("(MPa, C)" if si else "(ksi, F)"),
        f"Fuente: resources/{ed}/{fn} · Edicion 2025 · Valores tal como estan impresos.",
        [(tag, f"{ed}/{fn}")], "prop")
    record_meta(name, tag, f"{ed}/{fn}", "2025", system,
                info["last_row"] - R_DATA + 1, lbl)
    return info


def verificar_paridad(si_info, us_info, etiqueta):
    """Comprueba que cada fila metrica tiene homologa en la edicion US
    (clave bilingue). Se reporta; no se inventa ninguna correspondencia."""
    a = {d["bi"] for d in si_info["recs"]}
    b = {d["bi"] for d in us_info["recs"]}
    solo_si, solo_us = len(a - b), len(b - a)
    ISSUES.append(f"Paridad SI/US en {etiqueta}: {len(a & b)} materiales enlazados; "
                  f"{solo_si} solo en la edicion metrica; {solo_us} solo en la US. "
                  "Los no enlazados muestran 'sin equivalente en la edicion US'.")
    return len(a & b), solo_si, solo_us


# ---------------------------------------------------------------------------
# Bases por grupo de material (E, dilatacion, Poisson/densidad)
# ---------------------------------------------------------------------------
GRP_COLS = ["clave", "Tabla", "clave_sf", "Grupo / material", "Detalle",
            "Busqueda", "n_pts"]


def build_modulo(res, wb, system):
    si = system == "SI"
    ed = "bpvc_ii_d_metric_2025" if si else "bpvc_ii_d_customary_2025"
    temps, recs = set(), []
    for i in range(1, 6):
        d = res.load(f"{ed}/table_tm_{i}.json")
        tid = d.get("table_id")
        for row in d["rows"]:
            merged, extra = fix_merged_ident(row)
            mat = clean(merged or g(row, "materials", "material", "material_grade_uns_no"))
            grp = clean(g(row, "material_group"))
            vals = {}
            for k, v in (row.get("values") or {}).items():
                t = temp_to_number(k)
                if t is not None and num(v) is not None:
                    vals[t] = num(v)
            for k, v in extra.items():
                t = temp_to_number(k)
                if t is not None:
                    vals.setdefault(t, v)
            temps |= set(vals)
            recs.append(dict(tid=tid, mat=mat, grp=grp, vals=vals))
    temps = sorted(temps)
    recs.sort(key=lambda d: (d["tid"], txt(d["mat"]).upper()))
    keys, _ = make_unique([f"{d['tid']} | {txt(d['mat'])}" for d in recs],
                          list(range(1, len(recs) + 1)))
    records = [([k, d["tid"], d["tid"], d["mat"], d["grp"],
                 search_key(d["mat"], d["grp"]), len(d["vals"])], d["vals"])
               for k, d in zip(keys, recs)]
    name = "DB_E" if si else "DB_EC"
    fac = "x10^3 MPa" if si else "x10^6 psi"
    ws = new_sheet(wb, name,
                   f"BASE DE PROPIEDADES — Modulo de elasticidad E · ASME BPVC II-D 2025, "
                   f"Tablas TM-1 a TM-5 ({fac})",
                   f"Fuente: resources/{ed}/table_tm_1..5.json · Edicion 2025 · "
                   f"E real = valor tabulado x {'1000 MPa' if si else '1e6 psi'} (factor del "
                   "titulo de la tabla). Indexada por GRUPO de material: ver MAP_Grupo.")
    n = write_headers(ws, GRP_COLS, temps)
    last = write_rows(ws, records, n, temps)
    t0, v0, npack = append_packed(ws, records, n, temps, 7)
    autosize(ws, {"A": 42, "B": 8, "C": 8, "D": 42, "E": 20, "F": 28})
    ws.column_dimensions["C"].hidden = True
    ws.column_dimensions["F"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:{get_column_letter(n + len(temps))}{last}"
    record_meta(name, "TM-1..TM-5", f"{ed}/table_tm_1..5.json", "2025", system,
                last - R_DATA + 1, f"Modulo de elasticidad; factor {fac}.")
    return dict(sheet=name, n_ident=n, temps=temps, last_row=last,
                pack_t0=t0, pack_v0=v0, npack=npack, npts_col=7)


def build_prd(res, wb, system):
    si = system == "SI"
    ed = "bpvc_ii_d_metric_2025" if si else "bpvc_ii_d_customary_2025"
    d = res.load(f"{ed}/table_prd.json")
    rows = sorted(d["rows"], key=lambda r: (txt(g(r, "material_group")).upper(),
                                            txt(g(r, "material")).upper()))
    name = "DB_PRD" if si else "DB_PRDC"
    dens = "Densidad, kg/m3" if si else "Densidad, lb/in3"
    ws = new_sheet(wb, name,
                   "BASE DE PROPIEDADES — Coeficiente de Poisson y densidad · "
                   f"ASME BPVC II-D 2025, Tabla PRD ({'metrico' if si else 'US'})",
                   f"Fuente: resources/{ed}/table_prd.json · Edicion 2025 · "
                   "Indexada por GRUPO de material: ver MAP_Grupo.")
    n = write_headers(ws, ["clave", "Grupo", "clave_sf", "Material", "Poisson", dens])
    recs = []
    keys, _ = make_unique([f"PRD | {txt(g(r, 'material'))}" for r in rows],
                          list(range(1, len(rows) + 1)))
    for k, r in zip(keys, rows):
        grp = clean(g(r, "material_group"))
        recs.append(([k, grp, txt(grp), clean(g(r, "material")),
                      g(r, "poisson_s_ratio"),
                      g(r, "density_kg_m3", "density_lb_in3")], {}))
    last = write_rows(ws, recs, n)
    autosize(ws, {"A": 38, "B": 24, "C": 20, "D": 34, "E": 10, "F": 16})
    ws.column_dimensions["C"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:F{last}"
    record_meta(name, "PRD", f"{ed}/table_prd.json", "2025", system,
                last - R_DATA + 1, "Poisson y densidad por grupo de material.")
    return dict(sheet=name, n_ident=n, last_row=last)


def build_te(res, wb, system):
    """TE-1..5 se imprimen indexadas por temperatura con bloques de grupos.
    Se escriben tal como estan impresas (hoja de consulta con autofiltro)."""
    si = system == "SI"
    ed = "bpvc_ii_d_metric_2025" if si else "bpvc_ii_d_customary_2025"
    name = "DB_TE" if si else "DB_TEC"
    unit = "10^-6 mm/mm/C" if si else "10^-6 in/in/F"
    ws = new_sheet(wb, name,
                   "BASE DE PROPIEDADES — Dilatacion termica · ASME BPVC II-D 2025, "
                   f"Tablas TE-1 a TE-5 ({unit})",
                   f"Fuente: resources/{ed}/table_te_1..5.json · Edicion 2025 · "
                   "Indexadas por temperatura; cada bloque cubre un juego de grupos. "
                   "Coef. A = instantaneo, B = medio, C = expansion acumulada.")
    r, nrows = R_HDR, 0
    for i in range(1, 6):
        d = res.load(f"{ed}/table_te_{i}.json")
        prev = None
        for row in d["rows"]:
            keys = list(row.keys())
            if keys != prev:
                r += 1
                c = ws.cell(r, 1, f"{d.get('table_id')} — {d.get('title')}")
                c.font = Font(name="Calibri", size=9, bold=True, color=NAVY)
                r += 1
                for j, k in enumerate(keys, start=1):
                    cc = ws.cell(r, j, k)
                    cc.font, cc.fill, cc.border = HDR_F, HDR_FILL, BOX
                    cc.alignment = Alignment(wrap_text=True, vertical="center")
                prev = keys
            r += 1
            for j, k in enumerate(keys, start=1):
                if row[k] is not None:
                    ws.cell(r, j, row[k]).font = DATA_F
            nrows += 1
    autosize(ws, {get_column_letter(j): 18 for j in range(1, 13)})
    record_meta(name, "TE-1..TE-5", f"{ed}/table_te_1..5.json", "2025", system, nrows,
                "Dilatacion termica por grupo (bloques tal como estan impresos).")


# ---------------------------------------------------------------------------
# Apendice C del B31.3 y no metalicos
# ---------------------------------------------------------------------------
def _dedup_frase(t):
    """La extraccion del Apendice C repite el rotulo del grupo dos veces
    ('12Cr ... steels 12Cr ... steels'). Si el texto es exactamente X + ' ' + X,
    se colapsa a X: es un artefacto de extraccion, no contenido del codigo."""
    if not isinstance(t, str):
        return t
    w = t.strip()
    n = len(w)
    for cut in range(n // 2, n // 2 + 2):
        a, b = w[:cut].strip(), w[cut:].strip()
        if a and a == b:
            return a
    return w


def build_appendix_c(res, wb, system):
    si = system == "SI"
    suf = "" if si else "C"
    out = {}
    for tag, fn, sheet, lbl in [
            ("C-1", f"{APX}/appendix_c/table_c_1{'' if si else 'c'}.json",
             f"DB_C_dilatacion{suf}", "Dilatacion termica de metales"),
            ("C-3", f"{APX}/appendix_c/table_c_3{'' if si else 'c'}.json",
             f"DB_C_modulo{suf}", "Modulo de elasticidad de metales")]:
        d = res.load(fn)
        temps, recs = set(), []
        for row in d["rows"]:
            mat = _dedup_frase(clean(g(row, "material_uns_no", "material")))
            coef = clean(g(row, "coefficient", "coeffi_cient"))
            vals = {temp_to_number(k): num(v) for k, v in (row.get("values") or {}).items()
                    if temp_to_number(k) is not None and num(v) is not None}
            temps |= set(vals)
            recs.append(dict(mat=mat, coef=coef, vals=vals))
        temps = sorted(temps)
        recs.sort(key=lambda x: (txt(x["mat"]).upper(), txt(x["coef"])))
        keys, _ = make_unique([build_material_id([r["mat"], r["coef"]]) for r in recs],
                              list(range(1, len(recs) + 1)))
        records = [([k, tag, tag, r["mat"], r["coef"], search_key(r["mat"]),
                     len(r["vals"])], r["vals"]) for k, r in zip(keys, recs)]
        ws = new_sheet(wb, sheet,
                       f"PROPIEDADES B31.3 — {lbl} · Tabla {d.get('table_id')} "
                       f"({'SI' if si else 'US'})",
                       f"Fuente: resources/{fn} · ASME B31.3-2024 · {d.get('value_axis')}")
        n = write_headers(ws, ["clave", "Tabla", "clave_sf", "Material / grupo",
                               "Coef.", "Busqueda", "n_pts"], temps)
        last = write_rows(ws, records, n, temps)
        t0, v0, npack = append_packed(ws, records, n, temps, 7)
        autosize(ws, {"A": 50, "B": 8, "C": 8, "D": 48, "E": 8, "F": 28})
        for cc in ("C", "F", "G"):
            ws.column_dimensions[cc].hidden = True
        ws.auto_filter.ref = f"A{R_HDR}:{get_column_letter(n + len(temps))}{last}"
        record_meta(sheet, d.get("table_id"), fn, "2024", system, last - R_DATA + 1, lbl)
        out[tag] = dict(sheet=sheet, n_ident=n, temps=temps, last_row=last,
                        pack_t0=t0, pack_v0=v0, npack=npack, npts_col=7)
    return out


def build_nometalicos(res, wb):
    """Formato largo (Tabla | Material | Campo | Valor): permite cascada
    Tabla -> Material y ficha campo/valor sin matrices dinamicas."""
    tables = [("B-1", f"{APX}/appendix_b/table_b_1.json"),
              ("B-1C", f"{APX}/appendix_b/table_b_1c.json"),
              ("B-2", f"{APX}/appendix_b/table_b_2.json"),
              ("B-3", f"{APX}/appendix_b/table_b_3.json"),
              ("B-4", f"{APX}/appendix_b/table_b_4.json"),
              ("B-5", f"{APX}/appendix_b/table_b_5.json"),
              ("B-6", f"{APX}/appendix_b/table_b_6.json"),
              ("C-2", f"{APX}/appendix_c/table_c_2.json"),
              ("C-4", f"{APX}/appendix_c/table_c_4.json")]
    ws = new_sheet(wb, "DB_NoMetalicos",
                   "MATERIALES NO METALICOS — ASME B31.3-2024, Apendice B (HDS y presiones "
                   "admisibles) y Apendice C (C-2 dilatacion, C-4 modulo)",
                   "Fuente: resources/" + APX + "/appendix_b/table_b_1..b_6.json y "
                   "appendix_c/table_c_2.json, table_c_4.json · Formato largo "
                   "(tabla / material / campo / valor) para permitir la consulta por lista "
                   "desplegable. Valores tal como estan impresos.")
    n = write_headers(ws, ["clave", "Tabla", "clave_sf", "Material", "Campo", "Valor",
                           "Titulo de la tabla"])
    recs = []
    for tag, fn in tables:
        d = res.load(fn)
        title = d.get("title")
        for row in d.get("rows") or []:
            keys = [k for k in row if row[k] is not None]
            if not keys:
                continue
            mat = clean(g(row, "material_designation", "material", "material_description",
                          "astm_spec_no", "spec_no", "spec_nos_astm_except_as_noted"))
            mat = mat or clean(row[keys[0]])
            for k in keys:
                campo = k.replace("_", " ").strip().capitalize()
                recs.append(([f"{tag} | {txt(mat)}", tag, f"{tag}|{txt(mat)}",
                              mat, campo, row[k], title], {}))
    recs.sort(key=lambda r: (r[0][1], txt(r[0][3]).upper()))
    last = write_rows(ws, recs, n)
    autosize(ws, {"A": 40, "B": 8, "C": 34, "D": 34, "E": 28, "F": 22, "G": 70})
    ws.column_dimensions["C"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:G{last}"
    for tag, fn in tables:
        record_meta("DB_NoMetalicos", tag, fn, "2024", "SI/US", 0, "")
    return dict(sheet="DB_NoMetalicos", last_row=last)


# ---------------------------------------------------------------------------
# Factores, mapeo de grupos y notas
# ---------------------------------------------------------------------------
def build_map_factores(res, wb):
    fa2, fa3 = f"{APX}/appendix_a/table_a_2.json", f"{APX}/appendix_a/table_a_3.json"
    d2, d3 = res.load(fa2), res.load(fa3)
    ws = new_sheet(wb, "MAP_Factores",
                   "FACTORES DE CALIDAD ASME B31.3-2024 — Ej (Tabla A-3, junta longitudinal) "
                   "y Ec (Tabla A-2, fundicion)",
                   f"Fuente: resources/{fa3} y resources/{fa2} · ASME B31.3-2024. "
                   "El motor sustituye el E_j fijo por un lookup sobre esta hoja. "
                   "Ordenada por Spec. No. para alimentar la lista en cascada.")
    n = write_headers(ws, ["clave", "Tabla", "clave_sf", "Spec. No.", "Clase / Tipo",
                           "Descripcion", "Factor", "Notas", "Grupo de material"])
    rows = []
    for row in d3["rows"]:
        spec = txt(g(row, "spec_no"))
        cot, desc = clean(g(row, "class_or_type")), clean(g(row, "description"))
        rows.append((build_material_id([spec, cot or desc]), "A-3", spec, cot, desc,
                     g(row, "ej"), clean(g(row, "notes")), clean(g(row, "material_group"))))
    for row in d2["rows"]:
        spec, desc = txt(g(row, "spec_no")), clean(g(row, "description"))
        rows.append((f"Ec | {build_material_id([spec, desc])}", "A-2", spec, None, desc,
                     g(row, "ec"), clean(g(row, "notes")), clean(g(row, "material_group"))))
    rows.sort(key=lambda r: (r[2].upper(), r[1], txt(r[3] or r[4]).upper()))
    keys, _ = make_unique([r[0] for r in rows], list(range(1, len(rows) + 1)))
    recs = [([k, r[1], f"{r[1]}|{r[2]}", r[2] or None, r[3], r[4], r[5], r[6], r[7]], {})
            for k, r in zip(keys, rows)]
    last = write_rows(ws, recs, n)
    autosize(ws, {"A": 58, "B": 8, "C": 16, "D": 12, "E": 26, "F": 46, "G": 9,
                  "H": 10, "I": 26})
    ws.column_dimensions["C"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:I{last}"
    record_meta("MAP_Factores", "A-2 + A-3", f"{fa2} ; {fa3}", "2024", "adimensional",
                last - R_DATA + 1, "Ej y Ec por especificacion / tipo de junta.")
    return dict(sheet="MAP_Factores", last_row=last)


# ---------------------------------------------------------------------------
# Pertenencia a grupo de propiedades: se lee del codigo, no se infiere
# ---------------------------------------------------------------------------
# El grupo que da E (TM-1) y dilatacion (TE-1) NO se deduce de la composicion:
# esta impreso en las Notas al pie de esas dos tablas, que enumeran uno por uno
# los materiales de cada grupo. Aqui se cargan de resources/ y se consultan por
# coincidencia LITERAL.
#
# Hasta la Rev. 3 esto era una lista de expresiones regulares sobre la
# composicion impresa. Producia asignaciones FALSAS, no solo inciertas: el
# patron `mo\b` capturaba cualquier material con molibdeno y rotulaba
# «acero de baja aleacion» a los austeniticos 16Cr-12Ni-2Mo, a los duplex
# 22Cr-5Ni-3Mo-N y a las aleaciones de niquel 62Ni-22Mo-15Cr; el patron `8ni`
# casaba dentro de «18Ni». Una conjetura sentada junto a datos normativos es
# peor que un hueco declarado, asi que la heuristica se elimino entera.

# Algunos miembros no son un material sino una REGLA de inclusion redactada:
# la Nota (5) de TM-1 lista «9Cr-Mo, including variations thereof». Eso no se
# puede resolver por coincidencia literal ni se debe resolver por cuenta
# propia: es una decision del ingeniero, y se marca como tal.
REGLA_TEXTUAL = re.compile(r",\s*including\b", re.I)


# El codigo no imprime la barra de fraccion igual en todas partes: II-D usa la
# barra normal («C-1/2Mo») y el Apendice A del B31.3 la DIVISION SLASH U+2215
# («C-1∕2Mo»). Ninguna de las dos se pliega a la otra por NFKC, de modo que sin
# esto un contraste entre tablas falla en silencio: no da error, simplemente no
# encuentra nada.
_BARRAS = dict.fromkeys(
    (0x2044,    # FRACTION SLASH
     0x2215,    # DIVISION SLASH
     0xFF0F),   # FULLWIDTH SOLIDUS
    "/")


def comp_key(s) -> str:
    """Clave de comparacion de una composicion nominal impresa.

    El codigo imprime la misma composicion con o sin espacios alrededor del
    guion segun la columna en que caiga (`18Cr-10Ni-Cb` en la tabla,
    `18Cr - 10Ni - Cb` si la linea se justifico), de modo que el espacio no
    puede formar parte de la clave. Los guiones tipograficos ya los homogeneiza
    txt(); aqui se pliegan ademas las barras de fraccion, se quitan espacios y
    se sube a mayusculas.
    """
    return re.sub(r"\s+", "", txt(s)).translate(_BARRAS).upper()


# TM-1 rotula «Material Group H [Note (9)]» y TE-1 «... (Group 3) [Note (3)]»:
# el parentesis de cierre es opcional.
_REF_NOTA = re.compile(r"((?:Material )?Group [A-J0-9]+)\)?\s*\[Note \((\d+)\)\]")


def notas_referenciadas(datos) -> dict:
    """{grupo: nota que la propia tabla cita} leido de sus rotulos impresos.

    TM-1 lo imprime en la columna de materiales de cada fila («Material Group H
    [Note (9)]») y TE-1 en los encabezados de columna («... (Group 3) [Note
    (3)]»). Hace falta porque la edicion metrica trae el Grupo H duplicado en
    dos notas y solo una es la que la tabla referencia: citar la otra remitiria
    al lector a una nota huerfana.
    """
    textos = [str(v) for r in datos.get("rows", []) for v in r.values()]
    textos += [str(c) for c in datos.get("columns", [])]
    return {m.group(1): f"({m.group(2)})"
            for t in textos for m in _REF_NOTA.finditer(t)}


def cargar_notas_de_grupo(res, ed):
    """Indice {clave_composicion: [(tabla, nota, grupo)]} desde resources/.

    Lee `note_members` de TM-1 y TE-1, que produce extraer_notas_ii_d.py a
    partir del PDF del codigo. Devuelve tambien las reglas textuales, que no
    entran al indice literal porque nunca casarian.
    """
    indice: dict[str, list[tuple]] = {}
    textuales: list[tuple] = []
    for archivo, tabla in (("table_tm_1.json", "TM-1"), ("table_te_1.json", "TE-1")):
        datos = res.load(f"{ed}/{archivo}")
        referenciadas = notas_referenciadas(datos)
        notas = datos.get("note_members")
        if not notas:
            raise SystemExit(
                f"ERROR: {archivo} no trae `note_members`. El mapeo de grupos de "
                f"TM-1/TE-1 sale de las Notas de esas tablas; sin ellas no hay "
                f"fuente normativa y no se construye nada. Ejecute antes:\n"
                f"    python extraer_notas_ii_d.py --pdf <II-D metrica> "
                f"--resources <resources>")
        for nota in notas:
            if nota.get("tipo") != "grupo":
                continue          # las notas de alias no definen pertenencia
            # La edicion metrica imprime el Grupo H en las Notas (8) y (9), pero
            # su tabla solo apunta a la (9). Citar la (8) seria remitir a una
            # nota que la tabla no referencia: se descarta si la otra existe.
            if referenciadas.get(nota["grupo"], nota["nota"]) != nota["nota"]:
                continue
            origen = (tabla, nota["nota"], nota["grupo"])
            for miembro in nota["miembros"]:
                if REGLA_TEXTUAL.search(miembro):
                    textuales.append((*origen, miembro))
                else:
                    indice.setdefault(comp_key(miembro), []).append(origen)
    return indice, textuales


def _stem_textual(miembro: str) -> str:
    """Primer segmento de una regla textual: «9Cr-Mo, including...» -> «9CR»."""
    return comp_key(miembro.split(",")[0]).split("-")[0]


def _uns_tokens(clave: str) -> list[str]:
    return re.findall(r"[A-Z]\d{5}", clave.upper())


def _composicion_prestada(comp_idx, uns_txt, ck, notas_idx):
    """Recupera la composicion de una fila que no la imprime, via su UNS.

    Solo devuelve algo si se cumplen las tres condiciones a la vez:
      1. la fila no trae composicion propia,
      2. su UNS aparece en el libro con UNA sola composicion (si hay varias,
         el codigo no es consistente para ese UNS y no se elige por cuenta
         propia),
      3. esa composicion figura en alguna Nota de TM-1 o TE-1.
    Devuelve (composicion, hojas_de_origen, origenes_de_nota).
    """
    if ck:
        return None
    for tok in _uns_tokens(uns_txt):
        candidatas = comp_idx.get(tok)
        if not candidatas or len(candidatas) != 1:
            continue
        comp_p, hojas = next(iter(candidatas.items()))
        origenes = notas_idx.get(comp_key(comp_p))
        if origenes:
            return comp_p, hojas, origenes
    return None


def _motivo_sin_composicion(comp_idx, uns_txt) -> str:
    """Explica POR QUE no se pudo recuperar la composicion. El motivo importa:
    no es lo mismo que el UNS no aparezca en ningun lado que el codigo lo
    imprima con dos composiciones distintas."""
    for tok in _uns_tokens(uns_txt):
        candidatas = comp_idx.get(tok)
        if candidatas and len(candidatas) > 1:
            return (f"La fila no imprime composicion nominal y el libro asocia a su "
                    f"UNS «{tok}» mas de una: {'; '.join(candidatas)}. No se elige "
                    f"por cuenta propia.")
        if candidatas:
            return (f"La fila no imprime composicion nominal. Su UNS «{tok}» aparece "
                    f"como «{next(iter(candidatas))}», pero esa composicion no figura "
                    f"en ninguna Nota de TM-1 ni TE-1.")
    return ("La fila no imprime composicion nominal y su UNS no aparece con "
            "composicion en ninguna tabla del libro: no hay dato con el que "
            "buscar el grupo.")


def cargar_prd_por_uns(res, ed):
    """Indice {UNS: fila_de_PRD} desde table_prd.json.

    La Tabla PRD (Poisson y densidad) se indexa por descripcion de material,
    pero 99 de sus 115 filas NOMBRAN los UNS que cubren («N08800, N08810, and
    N08811», «S32202»). Esas si permiten un mapeo literal 1:1.

    Las 16 restantes son categorias redactadas —«Carbon steels», «1/2Cr to
    11/4Cr steels», «High alloy steels (300 series)»— que exigen interpretar a
    que categoria pertenece un material dado. Eso es criterio de ingenieria y
    no se resuelve aqui: es exactamente la clase de inferencia que la Rev. 3
    saco del motor.
    """
    idx = {}
    for r in res.load(f"{ed}/table_prd.json")["rows"]:
        etiqueta = txt(g(r, "material"))
        for tok in re.findall(r"[A-Z]\d{5}", etiqueta.upper()):
            idx[tok] = etiqueta
    return idx


def indice_composicion_por_uns(wb, infos):
    """Indice {UNS: {composicion: [hojas que la imprimen]}}.

    Sirve para las filas que NO imprimen composicion nominal: su UNS suele
    aparecer con composicion en otra tabla del mismo libro. El UNS lo asigna
    SAE/ASTM y designa el mismo material, asi que la composicion es la misma;
    pero al venir de OTRA tabla —a veces de otro codigo, el Apendice A del
    B31.3— el resultado no puede presentarse como AUTO. Se marca aparte y se
    cita de donde salio, para que el ingeniero lo confirme.
    """
    idx: dict[str, dict[str, list]] = {}
    for info in infos:
        ws = wb[info["sheet"]]
        for r in range(R_DATA, info["last_row"] + 1):
            uns = txt(ws.cell(r, C["UNS / Alloy"]).value).upper()
            comp = txt(ws.cell(r, C["Composicion nominal"]).value)
            if uns and comp:
                idx.setdefault(uns, {}).setdefault(comp, [])
                if info["sheet"] not in idx[uns][comp]:
                    idx[uns][comp].append(info["sheet"])
    return idx


# Estados del mapeo. No existe ya un estado «PROPUESTA»: o el codigo lo dice y
# se cita la nota, o se declara el hueco y el calculo queda bloqueado.
E_UNS = "AUTO (UNS exacto)"
E_NOTA = "AUTO (composicion en Nota)"
E_TEXTUAL = "REVISAR (regla textual del codigo)"
E_COMP_AJENA = "REVISAR (composicion de otra tabla)"
E_VALIDADO = "VALIDADO POR INGENIERO"
E_SIN = "SIN MAPEO"


def cargar_decisiones(ruta: Path):
    """Decisiones del ingeniero sobre lo que el codigo no resuelve.

    Cierra el circuito: `Revision_MAP_Grupo.md` dice que hay que decidir, y
    este archivo trae lo decidido de vuelta al motor. Va SIEMPRE en un estado
    propio, `VALIDADO POR INGENIERO`, nunca mezclado con las filas AUTO: quien
    audite el libro tiene que poder separar de un vistazo lo que dice el codigo
    de lo que decidio una persona.

    No vive en resources/ a proposito. resources/ es el codigo publicado; esto
    es criterio de ingenieria sobre lo que el codigo no cubre, y confundirlos
    seria exactamente el error que la Rev. 3 vino a corregir.
    """
    if not ruta.exists():
        return {}, {}
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)
    por_comp, por_uns = {}, {}
    sin_firma = 0
    for d in datos.get("decisiones", []):
        if not (txt(d.get("grupo_tm")) or txt(d.get("grupo_te"))):
            continue                     # entrada de plantilla, aun sin rellenar
        # Sin firma no se aplica. Una propuesta bien razonada sigue siendo una
        # propuesta: lo que convierte un grupo en dato utilizable es que un
        # ingeniero lo asuma. Sin esto, un archivo de propuestas copiado por
        # error entraria al calculo como si estuviese validado.
        if not txt(d.get("validado_por")):
            sin_firma += 1
            continue
        # Una decision se ancla a la composicion cuando la fila la imprime, y al
        # UNS cuando no: en esas filas lo que se decide es precisamente si ese
        # UNS designa el material cuya composicion se tomo prestada.
        # `uns` admite una lista: la misma pregunta suele cubrir varios UNS del
        # mismo material (los seis aceros que ASME imprime como «18Cr-8Ni», por
        # ejemplo). Una firma por pregunta, no por fila: repetir la firma 38
        # veces para 17 decisiones distintas invita a firmar sin mirar.
        uns_val = d.get("uns")
        if isinstance(uns_val, (list, tuple)):
            for u in uns_val:
                if txt(u):
                    por_uns[txt(u).upper()] = d
        elif txt(uns_val):
            por_uns[txt(uns_val).upper()] = d
        elif comp_key(d.get("composicion")):
            por_comp[comp_key(d["composicion"])] = d
    if sin_firma:
        ISSUES.append(f"decisiones: {sin_firma} entradas con grupo asignado pero SIN "
                      f"FIRMA en `validado_por`. NO se aplicaron: una propuesta sin "
                      f"firmar no es una decision.")
    return por_comp, por_uns


def _firma(d) -> str:
    quien = txt(d.get("validado_por")) or "sin firma"
    cuando = txt(d.get("fecha")) or "sin fecha"
    return f"{quien}, {cuando}"


def build_map_grupo(res, wb, iid_infos, comp_infos, ruta_decisiones,
                    system="SI"):
    # Las dos ediciones se mapean por separado contra SUS PROPIAS Notas: no
    # numeran igual y son extracciones independientes (ver extraer_notas_ii_d).
    si = system == "SI"
    ed = "bpvc_ii_d_metric_2025" if si else "bpvc_ii_d_customary_2025"
    nombre = "MAP_Grupo" if si else "MAP_GrupoC"
    tm_index = {}
    for i in range(1, 6):
        d = res.load(f"{ed}/table_tm_{i}.json")
        tid = d.get("table_id")
        for row in d["rows"]:
            merged, _ = fix_merged_ident(row)
            mat = merged or g(row, "materials", "material", "material_grade_uns_no")
            key = txt(mat).upper()
            for tok in re.findall(r"[A-Z]\d{5}", key):
                tm_index[tok] = (tid, txt(mat))
            tm_index.setdefault(key, (tid, txt(mat)))
    prd_por_uns = cargar_prd_por_uns(res, ed)
    comp_idx = indice_composicion_por_uns(wb, comp_infos)
    dec_comp, dec_uns = cargar_decisiones(ruta_decisiones)
    conflictos = {}

    notas_idx, textuales = cargar_notas_de_grupo(res, ed)

    ws = new_sheet(wb, nombre,
                   f"MAPEO material -> grupo de propiedades (TM / TE / PRD) · "
                   f"ASME BPVC II-D 2025 ({'Metrica' if si else 'U.S. Customary'})",
                   "TM y TE se indexan por GRUPO de material, no por especificacion. La "
                   "pertenencia esta impresa en las Notas al pie de TM-1 (Grupos A..J) y "
                   "TE-1 (Grupos 1..4), y cada fila cita la nota que la sostiene. "
                   "AUTO (UNS exacto) = el UNS figura literalmente en TM-1..TM-5. "
                   "AUTO (composicion en Nota) = la composicion nominal figura literalmente "
                   "en la lista de miembros de una nota. REVISAR = el codigo redacta una "
                   "regla de inclusion que debe aplicar el ingeniero. SIN MAPEO = II-D no "
                   "publica el dato: NO USAR E NI DILATACION.")
    n = write_headers(ws, ["material_id", "Spec. No.", "UNS / Alloy", "Composicion nominal",
                           "Grupo E (TM)", "Fuente E", "Grupo dilatacion (TE)", "Fuente alfa",
                           "Fila PRD (Poisson/densidad)", "Estado", "Motivo", "Busqueda"])
    recs, pendientes = [], []
    stats = {E_UNS: 0, E_NOTA: 0, E_VALIDADO: 0, E_TEXTUAL: 0,
             E_COMP_AJENA: 0, E_SIN: 0}
    for info in iid_infos:
        src = wb[info["sheet"]]
        for r in range(R_DATA, info["last_row"] + 1):
            mid = src.cell(r, C["material_id"]).value
            spec = src.cell(r, C["Spec. No."]).value
            uns = src.cell(r, C["UNS / Alloy"]).value
            comp = src.cell(r, C["Composicion nominal"]).value

            grp_e = fuente_e = grp_te = fuente_te = motivo = None

            # 1) UNS literal en TM-1..TM-5: el mapeo mas fuerte, fila contra fila.
            key = txt(uns).upper()
            hit = next((tm_index[t] for t in re.findall(r"[A-Z]\d{5}", key)
                        if t in tm_index), None)
            if hit is None and key in tm_index:
                hit = tm_index[key]
            if hit:
                grp_e, fuente_e, estado = hit[1], hit[0], E_UNS

            # 2) Composicion nominal listada en una Nota de TM-1 o TE-1.
            ck = comp_key(comp)
            for tabla, nota, grupo in notas_idx.get(ck, []):
                if tabla == "TM-1" and grp_e is None:
                    grp_e, fuente_e = grupo, f"TM-1 Nota {nota}"
                elif tabla == "TE-1" and grp_te is None:
                    grp_te, fuente_te = grupo, f"TE-1 Nota {nota}"
            if hit is None and (grp_e or grp_te):
                estado = E_NOTA

            # 3) Regla de inclusion redactada: la decide el ingeniero, no el script.
            if hit is None and not (grp_e or grp_te):
                stem = ck.split("-")[0]
                cand = [t for t in textuales if _stem_textual(t[3]) == stem]
                prestada = _composicion_prestada(comp_idx, key, ck, notas_idx)
                if cand:
                    tabla, nota, grupo, miembro = cand[0]
                    estado = E_TEXTUAL
                    motivo = (f"{tabla} Nota {nota} lista «{miembro}». Aplicar la regla "
                              f"y confirmar si este material queda dentro de {grupo}.")
                elif prestada:
                    # La fila no imprime composicion, pero su UNS aparece con una
                    # sola composicion en otra tabla del libro, y esa composicion
                    # si figura en una Nota. Se propone con toda la trazabilidad
                    # a la vista; NO se marca AUTO porque el dato no sale de la
                    # fila propia.
                    comp_p, hojas_p, origenes = prestada
                    for tabla, nota, grupo in origenes:
                        if tabla == "TM-1" and grp_e is None:
                            grp_e, fuente_e = grupo, f"TM-1 Nota {nota} (comp. prestada)"
                        elif tabla == "TE-1" and grp_te is None:
                            grp_te, fuente_te = grupo, f"TE-1 Nota {nota} (comp. prestada)"
                    estado = E_COMP_AJENA
                    motivo = (f"La fila no imprime composicion nominal. Su UNS «{key}» "
                              f"aparece como «{comp_p}» en {', '.join(hojas_p)}, y esa "
                              f"composicion si figura en Nota. Confirmar que es el mismo "
                              f"material antes de usar E o dilatacion.")
                elif not ck:
                    # Sin composicion impresa y sin forma de recuperarla: la fila
                    # del codigo no trae el dato de entrada.
                    estado = E_SIN
                    motivo = _motivo_sin_composicion(comp_idx, key)
                else:
                    estado = E_SIN
                    motivo = ("El UNS no figura en TM-1..TM-5 y la composicion nominal "
                              "no esta listada en ninguna Nota de TM-1 ni TE-1. "
                              "II-D no publica E ni dilatacion para este material.")
            elif grp_e is None or grp_te is None:
                falta = "E (TM-1)" if grp_e is None else "dilatacion (TE-1)"
                motivo = f"Solo se resolvio uno de los dos grupos; falta {falta}."

            # 4) Decision del ingeniero. Se aplica SOLO donde el codigo no
            #    resolvio. Si contradice algo que el codigo si dice, no se
            #    aplica: se deja el valor del codigo y se reporta el choque,
            #    porque una decision no puede pisar una fuente normativa sin
            #    que nadie se entere.
            dec = dec_comp.get(ck) or next(
                (dec_uns[t] for t in _uns_tokens(key) if t in dec_uns), None)
            if dec:
                if estado in (E_UNS, E_NOTA):
                    dtm, dte = txt(dec.get("grupo_tm")), txt(dec.get("grupo_te"))
                    if (dtm and grp_e and dtm != txt(grp_e)) or \
                       (dte and grp_te and dte != txt(grp_te)):
                        # Se agrupa por composicion: el choque es uno solo, aunque
                        # lo arrastren cientos de filas. Repetirlo por fila
                        # sepultaria el resto de las limitaciones del libro.
                        conflictos[txt(comp)] = (
                            f"el codigo asigna {txt(grp_e) or '-'} / "
                            f"{txt(grp_te) or '-'} y la decision dice "
                            f"{dtm or '-'} / {dte or '-'}. Se conserva lo del codigo.",
                            conflictos.get(txt(comp), (None, 0))[1] + 1)
                else:
                    if txt(dec.get("grupo_tm")):
                        grp_e = txt(dec.get("grupo_tm"))
                        fuente_e = f"Validado por ingeniero ({_firma(dec)})"
                    if txt(dec.get("grupo_te")):
                        grp_te = txt(dec.get("grupo_te"))
                        fuente_te = f"Validado por ingeniero ({_firma(dec)})"
                    estado = E_VALIDADO
                    motivo = txt(dec.get("justificacion")) or \
                        "Decision del ingeniero; sin justificacion registrada."

            # PRD: coincidencia literal del UNS contra las filas de PRD que
            # nombran los suyos. Hasta la Rev. 3 esto buscaba la descripcion de
            # PRD como SUBCADENA de la composicion nominal, lo que casi nunca
            # acertaba y, cuando acertaba, devolvia la familia gruesa
            # («Ferrous Materials») en vez de la fila que hay que consultar.
            prd = next((prd_por_uns[t] for t in _uns_tokens(key)
                        if t in prd_por_uns), None)
            stats[estado] += 1
            if estado in (E_TEXTUAL, E_COMP_AJENA, E_SIN):
                # Las filas sin composicion propia se agrupan por UNS: es la
                # unidad en que se decide, y agruparlas por la composicion
                # vacia las juntaria todas en una decision falsa.
                if txt(comp):
                    clave = ("composicion", txt(comp))
                else:
                    tok = next(iter(_uns_tokens(key)), "")
                    clave = ("uns", tok) if tok else ("composicion", "")
                pendientes.append((clave, estado, motivo, txt(spec), txt(uns)))
            recs.append(([mid, spec, uns, comp, grp_e, fuente_e, grp_te, fuente_te,
                          prd, estado, motivo,
                          search_key(mid, spec, uns, comp)], {}))
    last = write_rows(ws, recs, n)
    autosize(ws, {"A": 46, "B": 12, "C": 14, "D": 24, "E": 26, "F": 18, "G": 20,
                  "H": 18, "I": 20, "J": 28, "K": 60, "L": 28})
    ws.column_dimensions["L"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:K{last}"
    ISSUES.append(f"{nombre}: " + " · ".join(f"{k}={v}" for k, v in stats.items()))
    n_dec = len(dec_comp) + len(dec_uns)
    if n_dec:
        ISSUES.append(f"{nombre}: {n_dec} decisiones validadas leidas de "
                      f"{ruta_decisiones.name}; aplicadas a {stats[E_VALIDADO]} filas.")
    for comp_c, (detalle, n) in conflictos.items():
        ISSUES.append(f"{nombre} CONFLICTO decision vs codigo en «{comp_c}» "
                      f"({n} filas): {detalle}")
    n_prd = sum(1 for ident, _ in recs if ident[8])
    ISSUES.append(
        f"{nombre} columna 'Fila PRD': {n_prd} de {len(recs)} filas resueltas por "
        f"UNS literal contra las 99 filas de table_prd.json que nombran los suyos. "
        f"Las 16 filas restantes de PRD son categorias redactadas ('Carbon steels', "
        f"'High alloy steels (300 series)'): encuadrar un material en ellas es "
        f"criterio de ingenieria y no se resuelve aqui.")
    record_meta(nombre, "TM-1/TE-1 Notas + TM-1..TM-5 (UNS)",
                f"{ed}/table_tm_1.json (note_members) ; table_te_1.json "
                f"(note_members) ; table_tm_*.json ; table_prd.json", "2025", "-",
                last - R_DATA + 1,
                "Cada fila cita la Nota que sostiene su grupo. SIN MAPEO = II-D no "
                "publica el dato; no usar E ni dilatacion.")
    return last, stats, pendientes


def escribir_revision_map_grupo(pendientes, stats, ruta: Path) -> int:
    """Hoja de revision: lo que el codigo NO resuelve, agrupado para firmar.

    Se agrupa por COMPOSICION NOMINAL porque es la unidad en que el codigo
    decide: las 3 454 filas de material se reducen a unas decenas de decisiones
    reales, y revisar fila a fila seria repetir el mismo juicio cientos de
    veces. Se acompana el recuento de materiales afectados para que se vea que
    pesa cada decision.
    """
    from collections import Counter, OrderedDict
    grupos: "OrderedDict[tuple, dict]" = OrderedDict()
    for clave, estado, motivo, spec, uns in pendientes:
        g = grupos.setdefault((clave, estado), {"motivo": motivo, "n": 0,
                                                "specs": Counter(), "uns": Counter()})
        g["n"] += 1
        if spec:
            g["specs"][spec] += 1
        if uns:
            g["uns"][uns] += 1

    # Filas sin composicion impresa Y sin UNS con el que anclar la decision: no
    # hay dato de entrada que juzgar. Se cuentan aparte para no inflar la lista.
    sin_comp = grupos.pop((("composicion", ""), E_SIN), None)

    orden = {E_TEXTUAL: 0, E_COMP_AJENA: 1, E_SIN: 2}
    filas = sorted(grupos.items(), key=lambda kv: (orden[kv[0][1]], -kv[1]["n"]))

    L = ["# Revision de MAP_Grupo — grupos de propiedades sin resolver por el codigo",
         "",
         f"Generado por `build_db_materiales.py` el {datetime.date.today().isoformat()}.",
         "Fuente: Notas de las Tablas TM-1 y TE-1 de ASME BPVC II-D (Metrica) 2025,",
         "extraidas a `resources/` por `extraer_notas_ii_d.py`.",
         "",
         "## Que hay que decidir aqui",
         "",
         "Las filas que **no** aparecen en este documento ya estan resueltas contra el",
         "codigo y citan la nota que las sostiene: no requieren criterio de ingenieria.",
         "Lo que sigue es lo que el codigo no resuelve por si solo.",
         "",
         "| Estado | Filas de material |",
         "|---|---:|"]
    for k, v in stats.items():
        L.append(f"| {k} | {v} |")
    L += ["",
          f"Decisiones distintas a tomar: **{len(filas)}** "
          f"(sobre {sum(v['n'] for _, v in grupos.items())} filas de material)."]
    if sin_comp:
        L += ["",
              f"Aparte, **{sin_comp['n']} filas no imprimen composicion nominal** y su "
              "UNS no figura en TM-1..TM-5.",
              "No entran en esta revision porque no hay dato de entrada que juzgar: para",
              "asignarles grupo habria que identificar el material por otra via (la",
              "especificacion y el grado en la tabla de origen)."]
    L += ["",
          "## Decisiones",
          "",
          "Las filas se cuentan sobre las DOS ediciones (metrica y U.S. Customary):",
          "una misma decision desbloquea el material en ambas, porque la pertenencia a",
          "grupo no depende del sistema de unidades.",
          "",
          "| # | Se decide sobre | Filas | Estado | Que hay que decidir | Grupo asignado | Firma / fecha |",
          "|---:|---|---:|---|---|---|---|"]
    for i, ((clave, estado), g) in enumerate(filas, start=1):
        tipo, valor = clave
        specs = ", ".join(s for s, _ in g["specs"].most_common(3))
        if len(g["specs"]) > 3:
            specs += f", +{len(g['specs']) - 3} mas"
        motivo = (g["motivo"] or "").replace("\n", " ")
        etiqueta = f"`{valor}`" + (" (UNS)" if tipo == "uns" else "")
        L.append(f"| {i} | {etiqueta} | {g['n']} | {estado} | {motivo} "
                 f"<br>Especificaciones: {specs or '—'} |  |  |")
    L += ["",
          "## Como usar este documento",
          "",
          "1. Para cada fila, decida el grupo de TM-1 (modulo E) y/o TE-1 (dilatacion)",
          "   que corresponde, o confirme que II-D no publica el dato para ese material.",
          "2. Anote el grupo en la columna correspondiente y firme.",
          "3. Mientras una fila siga sin grupo, el motor deja el calculo BLOQUEADO para",
          "   esos materiales. Es el comportamiento correcto: el codigo prohibe",
          "   extrapolar y prohibe inventar la pertenencia a un grupo.",
          ""]
    L += ["",
          "## Vuelta al motor",
          "",
          "Para que estas decisiones lleguen al calculo, copie",
          "`decisiones_map_grupo.plantilla.json` a `decisiones_map_grupo.json`,",
          "rellene `grupo_tm` / `grupo_te` con el rotulo tal como lo imprime el codigo",
          "(«Material Group E», «Group 1») y firme cada entrada. El builder las lee en",
          "la siguiente corrida y esas filas pasan al estado VALIDADO POR INGENIERO,",
          "**siempre separado de las filas AUTO**: quien audite el libro tiene que poder",
          "distinguir lo que dice el codigo de lo que decidio una persona.",
          "",
          "Una decision no puede pisar al codigo: si contradice un grupo que el codigo",
          "si asigna, no se aplica y el choque se reporta en las limitaciones del libro.",
          ""]
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text("\n".join(L), encoding="utf-8")

    # Plantilla legible por el motor, con lo pendiente ya listado. Se escribe
    # siempre; el archivo que el ingeniero rellena es OTRO, para no pisarselo.
    plantilla = {
        "_formato": ("Decisiones de ingenieria sobre la pertenencia a grupo de "
                     "propiedades de ASME BPVC II-D, para lo que el codigo no "
                     "resuelve por si solo. Copie este archivo a "
                     "decisiones_map_grupo.json y rellenelo."),
        "_como_rellenar": {
            "composicion": "tal como la imprime el codigo; se compara sin espacios",
            "uns": ("presente en vez de `composicion` cuando la fila del codigo no "
                    "imprime composicion nominal: ahi lo que se decide es si ese UNS "
                    "designa el material cuya composicion se tomo prestada"),
            "grupo_tm": "rotulo de TM-1, p. ej. 'Material Group E' (modulo E)",
            "grupo_te": "rotulo de TE-1, p. ej. 'Group 1' (dilatacion)",
            "justificacion": "por que; queda impreso en la hoja MAP_Grupo",
            "validado_por": "iniciales o nombre",
            "fecha": "AAAA-MM-DD",
        },
        "_aviso": ("Dejar grupo_tm y grupo_te vacios equivale a no decidir: la fila "
                   "sigue bloqueada. Una decision nunca sobreescribe un grupo que el "
                   "codigo si asigna."),
        "decisiones": [
            {("uns" if tipo == "uns" else "composicion"): valor,
             "grupo_tm": "", "grupo_te": "",
             "justificacion": "", "validado_por": "", "fecha": "",
             "_filas_afectadas": g["n"], "_estado_actual": estado,
             "_motivo": (g["motivo"] or "").replace("\n", " ")}
            for ((tipo, valor), estado), g in filas
        ],
    }
    ruta.with_name("decisiones_map_grupo.plantilla.json").write_text(
        json.dumps(plantilla, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(filas)


def iter_notas(d):
    """Recorre TODAS las estructuras de notas que usan las extracciones:
    B31.3 -> 'general_notes' + 'notes';  BPVC II-D -> 'sections'[].'items'[].
    Devuelve (seccion, id, texto)."""
    out = []
    for key, sec in (("general_notes", "Notas generales"), ("notes", "Notas")):
        for it in d.get(key) or []:
            if isinstance(it, dict):
                out.append((sec, it.get("id"), it.get("text")))
            else:
                out.append((sec, None, str(it)))
    for blk in d.get("sections") or []:
        sec = blk.get("section") if isinstance(blk, dict) else None
        for it in (blk.get("items") if isinstance(blk, dict) else None) or []:
            if isinstance(it, dict):
                out.append((sec, it.get("id"), it.get("text")))
            else:
                out.append((sec, None, str(it)))
    for it in d.get("rows") or []:
        if isinstance(it, dict):
            out.append(("Notas", it.get("id", it.get("note")),
                        it.get("text", it.get("description"))))
    return out


def build_notas(res, wb):
    ws = new_sheet(wb, "Notas_Codigo",
                   "NOTAS DE LAS TABLAS DE MATERIALES — ASME B31.3-2024 y ASME BPVC II-D 2025",
                   "Texto tal como esta impreso, incluidas las NOTAS GENERALES de cada tabla. "
                   "Las notas restringen soldadura, tratamiento termico, servicio o "
                   "temperatura: leerlas SIEMPRE antes de emitir un calculo.")
    n = write_headers(ws, ["Fuente", "Tabla", "Seccion", "Nota", "Texto"])
    srcs = [("B31.3", "A-1/A-1C", f"{APX}/appendix_a/notes_tables_a_1_a_1c.json"),
            ("B31.3", "A-4/A-4C", f"{APX}/appendix_a/notes_tables_a_4_a_4c.json"),
            ("II-D", "1A", "bpvc_ii_d_metric_2025/notes_table_1a.json"),
            ("II-D", "1B", "bpvc_ii_d_metric_2025/notes_table_1b.json"),
            ("II-D", "3", "bpvc_ii_d_metric_2025/notes_table_3.json"),
            ("II-D", "U", "bpvc_ii_d_metric_2025/notes_table_u.json"),
            ("II-D", "Y-1", "bpvc_ii_d_metric_2025/notes_table_y_1.json")]
    recs = []
    total = 0
    for code, tab, f in srcs:
        items = iter_notas(res.load(f))
        for sec, nid, t in items:
            recs.append(([code, tab, clean(sec), clean(nid), clean(t)], {}))
        total += len(items)
        record_meta("Notas_Codigo", tab, f, "-", "-", len(items), f"Notas {code}")
    last = write_rows(ws, recs, n)
    for r in range(R_DATA, last + 1):
        ws.cell(r, 5).alignment = Alignment(wrap_text=True, vertical="top")
    autosize(ws, {"A": 10, "B": 10, "C": 22, "D": 10, "E": 140})
    ws.auto_filter.ref = f"A{R_HDR}:E{last}"
    ISSUES.append(f"Notas_Codigo: {total} notas cargadas (generales + numeradas de las "
                  "7 tablas).")
    return last


# ---------------------------------------------------------------------------
# DB_Listas — listas unicas de la cascada (bloques contiguos, sin formulas)
# ---------------------------------------------------------------------------
def build_listas(wb, bases, simples=()):
    """Listas de la cascada de 5 niveles (familia -> composicion -> forma ->
    especificacion -> tipo/grado). Devuelve RANGOS explicitos: la validacion de
    datos de Google Sheets solo acepta un rango literal, nunca una formula."""
    ws = new_sheet(wb, "DB_Listas",
                   "LISTAS DE LA CASCADA — alimentan las listas desplegables de los "
                   "buscadores y del motor",
                   "Bloques contiguos derivados de cada base. La FAMILIA es una agrupacion "
                   "de navegacion derivada del prefijo UNS y de la composicion nominal "
                   "impresa; no es un dato normativo y no interviene en ningun calculo. "
                   "No es fuente normativa: se regenera con build_db_materiales.py.")
    from collections import Counter
    col, rangos = 1, {}
    for pref, info in bases:
        src = wb[info["sheet"]]
        rows = []
        for r in range(R_DATA, info["last_row"] + 1):
            ks = [src.cell(r, C[f"k{n}"]).value or "" for n in range(5)]
            parts = [k.split("|") for k in ks]
            rows.append((ks[0],
                         parts[1][1] if len(parts[1]) > 1 else "",
                         parts[2][2] if len(parts[2]) > 2 else "",
                         parts[3][3] if len(parts[3]) > 3 else "",
                         parts[4][4] if len(parts[4]) > 4 else "",
                         ks[0], ks[1], ks[2], ks[3], ks[4]))
        d = {}
        # (etiqueta, indice de clave, indice de valor)
        for tag, ki, vi in [("FAM", None, 0), ("C", 5, 1), ("F", 6, 2),
                            ("S", 7, 3), ("G", 8, 4)]:
            seen, out = set(), []
            for row in rows:
                k = row[ki] if ki is not None else ""
                v = row[vi]
                if (k, v) in seen:
                    continue
                seen.add((k, v))
                out.append((k, v))
            ws.cell(R_HDR, col, f"{pref} · {tag} clave").font = HDR_F
            ws.cell(R_HDR, col).fill = HDR_FILL
            ws.cell(R_HDR, col + 1, f"{pref} · {tag} valor").font = HDR_F
            ws.cell(R_HDR, col + 1).fill = HDR_FILL
            for i2, (k, v) in enumerate(out, start=R_DATA):
                ws.cell(i2, col, k).font = DATA_F
                ws.cell(i2, col + 1, v if v not in (None, "") else "(sin dato)").font = DATA_F
            n = len(out)
            L, L2 = get_column_letter(col), get_column_letter(col + 1)
            kr = f"DB_Listas!${L}${R_DATA}:${L}${R_DATA + n - 1}"
            vr = f"DB_Listas!${L2}${R_DATA}:${L2}${R_DATA + n - 1}"
            if tag == "FAM":
                d["FAM"] = vr
            else:
                d[tag + "K"], d[tag + "V"] = kr, vr
                cnt = Counter(k for k, _ in out)
                d["max" + tag] = max(cnt.values()) if cnt else 1
            ws.column_dimensions[L].hidden = True
            ws.column_dimensions[L2].width = 26
            col += 2
        sh, lastr = info["sheet"], info["last_row"]
        d["ID"] = f"{sh}!${CL['material_id']}${R_DATA}:${CL['material_id']}${lastr}"
        d["K4"] = f"{sh}!${CL['k4']}${R_DATA}:${CL['k4']}${lastr}"
        d["maxV"] = max(Counter(r[9] for r in rows).values()) if rows else 1
        rangos[pref] = d
    for nm, values in simples:
        ws.cell(R_HDR, col, nm).font = HDR_F
        ws.cell(R_HDR, col).fill = HDR_FILL
        for i2, v in enumerate(values, start=R_DATA):
            ws.cell(i2, col, v).font = DATA_F
        L = get_column_letter(col)
        rangos[nm] = f"DB_Listas!${L}${R_DATA}:${L}${R_DATA + len(values) - 1}"
        ws.column_dimensions[L].width = 26
        col += 1
    record_meta("DB_Listas", "-", "derivado", "-", "-", 0,
                "Listas unicas de la cascada de 5 niveles + variante.")
    return rangos


# ---------------------------------------------------------------------------
# Buscador de materiales — panel de seleccion + tarjeta de resultado
# ---------------------------------------------------------------------------
CASC_COL = {"C": 50, "F": 51, "S": 52, "G": 53, "V": 54}   # listas dependientes
AUX_COL = 46                                               # auxiliares de resolucion
NCOLS = 12                                                 # ancho del panel (A..L)

CARD_FILL = PatternFill("solid", fgColor="EEF3FA")
KPI_FILL = PatternFill("solid", fgColor="DCE6F5")
CARD_BORDER = Border(left=Side("thin", color="9DB7DC"), right=Side("thin", color="9DB7DC"),
                     top=Side("thin", color="9DB7DC"), bottom=Side("thin", color="9DB7DC"))
LBL2_F = Font(name="Calibri", size=9, bold=True, color="44546A")
VAL_F = Font(name="Calibri", size=11, bold=True, color="1F3864")
UNIT_F = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
KPI_TIT_F = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
KPI_VAL_F = Font(name="Calibri", size=18, bold=True, color="1F3864")


def banda(ws, r, texto, n=NCOLS):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
    c = ws.cell(r, 1, texto)
    c.font, c.fill = TITLE_F, BAND_FILL
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[r].height = 20


def _mrg(ws, r, c1, c2, value=None):
    if c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(r, c1)
    if value is not None:
        cell.value = value
    return cell


def campo(ws, r, c1, etiqueta, formula, unidad=None):
    """etiqueta (2 col) | valor (2 col) | unidad (1 col)."""
    e = _mrg(ws, r, c1, c1 + 1, etiqueta)
    e.font = LBL2_F
    e.alignment = Alignment(vertical="center", indent=1)
    v = _mrg(ws, r, c1 + 2, c1 + 3, formula)
    v.font = VAL_F
    v.alignment = Alignment(vertical="center", wrap_text=True)
    u = ws.cell(r, c1 + 4)
    if unidad:
        u.value = unidad
    u.font = UNIT_F
    for cc in range(c1, c1 + 5):
        ws.cell(r, cc).fill = CARD_FILL
        ws.cell(r, cc).border = CARD_BORDER
    ws.row_dimensions[r].height = 17


def build_buscador(wb, curvas, name, title, pref, rng, master, us, unit_si, unit_us,
                   valor_lbl, temp_si="°C", temp_us="°F", nota=""):
    ws = new_sheet(wb, name, title,
                   "Cascada de seleccion: familia -> composicion nominal -> forma de "
                   "producto -> especificacion -> tipo/grado. La UNICA celda que se "
                   "escribe es la temperatura de consulta. " + nota)
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 26
    U_VAL = f'IF($D$5="SI","{unit_si}","{unit_us}")'
    U_TMP = f'IF($D$5="SI","{temp_si}","{temp_us}")'

    banda(ws, 4, "1 · SELECCION DEL MATERIAL   —   todo por lista desplegable")
    filas = [(5, "Sistema de unidades", "SI"),
             (6, "0 · Familia de material", ""),
             (7, "1 · Composicion nominal", ""),
             (8, "2 · Forma de producto", ""),
             (9, "3 · Especificacion (Spec. No.)", ""),
             (10, "4 · Tipo / Grado", ""),
             (11, "5 · Variante (clase / tamano) — opcional", "")]
    for r, et, val in filas:
        e = _mrg(ws, r, 1, 3, et)
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        c = _mrg(ws, r, 4, 6, val)
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        ws.row_dimensions[r].height = 17
    e = _mrg(ws, 12, 1, 3, "TEMPERATURA DE CONSULTA   (unica celda de escritura)")
    e.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    e.alignment = Alignment(vertical="center", indent=1)
    c = _mrg(ws, 12, 4, 5, 25)
    c.font, c.fill = IN_F, IN_FILL
    c.border = Border(*[Side("medium", color="C00000")] * 4)
    ws.cell(12, 6).value = f"={U_TMP}"
    ws.cell(12, 6).font = UNIT_F
    e = _mrg(ws, 13, 1, 3, "Modo de lectura")
    e.font = LBL_F
    e.alignment = Alignment(vertical="center", indent=1)
    c = _mrg(ws, 13, 4, 6, "Interpolado")
    c.font, c.fill, c.border = IN_F, IN_FILL, BOX
    ay = _mrg(ws, 5, 7, NCOLS)
    ay.value = (f'=IF($D$5="SI","Leyendo {master["sheet"]} — valores en {unit_si}, '
                f'temperatura en {temp_si}","Leyendo {us["sheet"]} — valores en {unit_us}, '
                f'temperatura en {temp_us}")')
    ay.font = Font(italic=True, color=BLUE)
    ay2 = _mrg(ws, 12, 7, NCOLS)
    ay2.value = ('=IF($D$10="","Complete la cascada hasta el paso 4 para ver el resultado",'
                 '"Seleccion completa")')
    ay2.font = Font(italic=True, color="C00000")

    dv_list(ws, "D5", '"SI,US"')
    dv_list(ws, "D13", '"Interpolado,Tabulado-conservador"')
    dv_list(ws, "D6", "=" + rng["FAM"])
    hl = get_column_letter
    levels = [("C", "D7", rng["CK"], rng["CV"], "$D$6", rng.get("maxC", 60)),
              ("F", "D8", rng["FK"], rng["FV"], '$D$6&"|"&$D$7', rng.get("maxF", 30)),
              ("S", "D9", rng["SK"], rng["SV"], '$D$6&"|"&$D$7&"|"&$D$8',
               rng.get("maxS", 40)),
              ("G", "D10", rng["GK"], rng["GV"],
               '$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9', rng.get("maxG", 40)),
              ("V", "D11", rng["K4"], rng["ID"],
               '$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9&"|"&$D$10', rng.get("maxV", 40))]
    for tag, cell, kr, vr, key, mx in levels:
        col = CASC_COL[tag]
        L = hl(col)
        mx = max(1, min(int(mx), 250))
        ws.cell(R_HDR, col, f"lista {tag}").font = SRC_F
        for i2 in range(1, mx + 1):
            ws.cell(R_DATA + i2 - 1, col).value = (
                f'=IF(COUNTIF({kr},{key})<{i2},"",'
                f'INDEX({vr},MATCH({key},{kr},0)+{i2}-1))')
        ws.column_dimensions[L].hidden = True
        dv_list(ws, cell, f"=${L}${R_DATA}:${L}${R_DATA + mx - 1}")

    # --- auxiliares de resolucion (columnas ocultas) -----------------------
    A = AUX_COL
    LA = hl(A + 1)
    ws.cell(4, A, "k4").font = SRC_F
    ws.cell(4, A + 1).value = '=$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9&"|"&$D$10'
    K4 = f"${LA}$4"
    ws.cell(5, A, "inicio k4").font = SRC_F
    ws.cell(5, A + 1).value = f'=IFERROR(MATCH({K4},{rng["K4"]},0),0)'
    ws.cell(6, A, "n variantes").font = SRC_F
    ws.cell(6, A + 1).value = f'=COUNTIF({rng["K4"]},{K4})'
    START4, NVAR = f"${LA}$5", f"${LA}$6"
    ws.cell(7, A, "material_id").font = SRC_F
    ws.cell(7, A + 1).value = (f'=IF($D$11<>"",$D$11,IF({START4}=0,"",'
                               f'INDEX({rng["ID"]},{START4})))')
    MIDC = f"${LA}$7"
    ws.cell(8, A, "fila SI").font = SRC_F
    ws.cell(8, A + 1).value = f'=IFERROR(MATCH({MIDC},{rng["ID"]},0),"")'
    FSI = f"${LA}$8"
    ws.cell(9, A, "clave bilingue").font = SRC_F
    ws.cell(9, A + 1).value = f'=IF({FSI}="","",INDEX({_rng(master,"clave_bi")},{FSI}))'
    BI = f"${LA}$9"
    ws.cell(10, A, "fila US").font = SRC_F
    ws.cell(10, A + 1).value = (f'=IF({BI}="","",IFERROR(MATCH({BI},'
                                f'{_rng(us,"clave_bi")},0),""))')
    FUS = f"${LA}$10"
    ws.cell(11, A, "fila activa").font = SRC_F
    ws.cell(11, A + 1).value = f'=IF($D$5="SI",{FSI},{FUS})'
    FIL = f"${LA}$11"
    for jj in (A, A + 1):
        ws.column_dimensions[hl(jj)].hidden = True

    def ident(colname):
        return f'IF($D$5="SI",{_rng(master, colname)},{_rng(us, colname)})'

    return dict(ws=ws, pref=pref, rng=rng, master=master, us=us, A=A,
                MIDC=MIDC, FSI=FSI, FUS=FUS, FIL=FIL, NVAR=NVAR, ident=ident,
                valor_lbl=valor_lbl, U_VAL=U_VAL, U_TMP=U_TMP,
                unit_si=unit_si, unit_us=unit_us, temp_si=temp_si, temp_us=temp_us)


def _rng(info, colname):
    L = CL[colname]
    return f"{info['sheet']}!${L}${R_DATA}:${L}${info['last_row']}"


def finish_buscador(ctx, wb, curvas, cidx):
    """Tarjeta de resultado + ficha tecnica + curva. Sin tablas de valores:
    el resultado es el del filtro aplicado en la cascada."""
    ws, hl = ctx["ws"], get_column_letter
    master, us, A = ctx["master"], ctx["us"], ctx["A"]
    FIL, U_VAL, U_TMP = ctx["FIL"], ctx["U_VAL"], ctx["U_TMP"]
    rm, ru = packed_refs(master), packed_refs(us)
    npack = max(master["npack"], us["npack"])
    SEL = '$D$5="SI"'
    T_ANC = f'IF({SEL},{rm["t_anchor"]},{ru["t_anchor"]})'
    V_ANC = f'IF({SEL},{rm["v_anchor"]},{ru["v_anchor"]})'
    LA = hl(A + 1)

    ws.cell(12, A, "n_pts").font = SRC_F
    ws.cell(12, A + 1).value = (f'=IF({FIL}="",0,IFERROR(IF({SEL},'
                                f'INDEX({rm["npts"]},{FIL}),INDEX({ru["npts"]},{FIL})),0))')
    NP = f"${LA}$12"
    TR = f'OFFSET({T_ANC},{FIL}-1,0,1,MAX(1,{NP}))'
    VR = f'OFFSET({V_ANC},{FIL}-1,0,1,MAX(1,{NP}))'
    ws.cell(13, A, "p1").font = SRC_F
    ws.cell(13, A + 1).value = f'=IF({FIL}="","",IFERROR(MATCH($D$12,{TR},1),1))'
    P1 = f"${LA}$13"
    for i2, (lb, fml) in enumerate([
            ("T1", f'=IF({FIL}="","",IFERROR(INDEX({TR},{P1}),""))'),
            ("S1", f'=IF({FIL}="","",IFERROR(INDEX({VR},{P1}),""))'),
            ("T2", f'=IF({FIL}="","",IFERROR(INDEX({TR},{P1}+1),""))'),
            ("S2", f'=IF({FIL}="","",IFERROR(INDEX({VR},{P1}+1),""))')]):
        ws.cell(14 + i2, A, lb).font = SRC_F
        ws.cell(14 + i2, A + 1).value = fml
    T1C, S1C = f"${LA}$14", f"${LA}$15"
    T2C, S2C = f"${LA}$16", f"${LA}$17"
    VALOR = (f'IF({FIL}="","",' + interp_value(T1C, S1C, T2C, S2C, "$D$12", "$D$13")[1:] + ')')
    ws.cell(18, A, "S(T)").font = SRC_F
    ws.cell(18, A + 1).value = "=" + VALOR
    VALC = f"${LA}$18"
    TMAXC = f'INDEX({ctx["ident"]("Temp. max. / limite")},{FIL})'
    ws.cell(19, A, "estado").font = SRC_F
    ws.cell(19, A + 1).value = (
        f'=IF({FIL}="","SIN SELECCION",'
        f'IF({S1C}="","FUERA DE RANGO — sin valor tabulado",'
        f'IF(AND(ISNUMBER({TMAXC}),$D$12>{TMAXC}),'
        f'"FUERA DE RANGO — T supera la Temp. max.","EN RANGO")))')
    EST = f"${LA}$19"

    # ------------------------- 2 · RESULTADO ------------------------------
    banda(ws, 15, "2 · RESULTADO DE LA CONSULTA")
    mid = _mrg(ws, 16, 1, NCOLS)
    mid.value = (f'=IF({FIL}="","Complete la cascada para obtener el material",'
                 f'"Material seleccionado:   "&{ctx["MIDC"]})')
    mid.font = Font(name="Calibri", size=12, bold=True, color=NAVY)
    mid.alignment = Alignment(vertical="center", indent=1)
    mid.fill = CARD_FILL
    ws.row_dimensions[16].height = 20
    kpis = [(1, ctx["valor_lbl"].upper() + " A LA TEMPERATURA DE CONSULTA",
             "=" + VALC, f"={U_VAL}"),
            (4, "TEMPERATURA DE CONSULTA", "=$D$12", f"={U_TMP}"),
            (7, "MODO DE LECTURA", "=$D$13", '="segun MODO_S"'),
            (10, "ESTADO DEL RANGO", "=" + EST,
             f'=IF({FIL}="","","variantes de este grado: "&{ctx["NVAR"]})')]
    for c1, tit, val, uni in kpis:
        t = _mrg(ws, 17, c1, c1 + 2, tit)
        t.font, t.fill = KPI_TIT_F, BAND_FILL
        t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        v = _mrg(ws, 18, c1, c1 + 2, val)
        v.font, v.fill = KPI_VAL_F, KPI_FILL
        v.alignment = Alignment(horizontal="center", vertical="center")
        u = _mrg(ws, 19, c1, c1 + 2, uni)
        u.font, u.fill = UNIT_F, KPI_FILL
        u.alignment = Alignment(horizontal="center")
        for r2 in (17, 18, 19):
            for cc in range(c1, c1 + 3):
                ws.cell(r2, cc).border = CARD_BORDER
    ws.row_dimensions[17].height = 26
    ws.row_dimensions[18].height = 30
    ws.row_dimensions[19].height = 14
    ws.conditional_formatting.add(
        f"A18:L19",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("FUERA DE RANGO",{EST}))'],
                    font=Font(color="9C0006", bold=True)))

    # --------------------- 3 · FICHA TECNICA ------------------------------
    banda(ws, 21, "3 · FICHA TECNICA DEL MATERIAL   —   identificacion tal como la "
                  "organiza el codigo")

    def val_of(colname):
        idx = f'INDEX({ctx["ident"](colname)},{FIL})'
        return f'=IF({FIL}="","—",IF({idx}="","—",{idx}))'

    izq = [("Tabla del codigo", val_of("Tabla"), None),
           ("Familia de material", val_of("Familia"), '="agrupacion de navegacion"'),
           ("Composicion nominal", val_of("Composicion nominal"), None),
           ("Forma de producto", val_of("Forma de producto"), None),
           ("Especificacion (Spec. No.)", val_of("Spec. No."), None),
           ("Tipo / Grado", val_of("Tipo/Grado"), None),
           ("UNS / Alloy No.", val_of("UNS / Alloy"), None),
           ("Clase / Condicion / Temple", val_of("Clase/Cond./Temple"), None),
           ("Tamano / Espesor", val_of("Tamano/Espesor"),
            f'=IF($D$5="SI","mm","in")'),
           ("Notas del codigo", val_of("Notas"), '="ver Notas_Codigo"')]
    der = [("Edicion consultada",
            f'=IF({FIL}="","—",IF($D$5="SI","{master["sheet"]} — metrica",'
            f'IF({ctx["FUS"]}="","sin equivalente en la edicion US",'
            f'"{us["sheet"]} — U.S. Customary")))', None),
           ("P-No.", val_of("P-No."), '="adimensional"'),
           ("Group No.", val_of("Group No."), '="adimensional"'),
           ("Temp. min. / curva de impacto", val_of("Temp. min. / curva impacto"),
            f'=IF(ISNUMBER({FIL}),{U_TMP},"")'),
           ("Resistencia a la traccion min.", val_of("Resist. traccion min."),
            f"={U_VAL}"),
           ("Limite de fluencia min.", val_of("Fluencia min."), f"={U_VAL}"),
           ("Temp. max. admisible / limite", val_of("Temp. max. / limite"),
            f"={U_TMP}"),
           ("Aplicabilidad  I / III",
            f'=IF({FIL}="","—",IF(INDEX({ctx["ident"]("I")},{FIL})&'
            f'INDEX({ctx["ident"]("III")},{FIL})="","no aplica (tabla del B31.3)",'
            f'INDEX({ctx["ident"]("I")},{FIL})&"  /  "&'
            f'INDEX({ctx["ident"]("III")},{FIL})))', '="NP = no permitido"'),
           ("Aplicabilidad  VIII-1 / VIII-2 / XII",
            f'=IF({FIL}="","—",IF(INDEX({ctx["ident"]("VIII-1")},{FIL})&'
            f'INDEX({ctx["ident"]("VIII-2")},{FIL})&INDEX({ctx["ident"]("XII")},{FIL})="",'
            f'"no aplica (tabla del B31.3)",'
            f'INDEX({ctx["ident"]("VIII-1")},{FIL})&"  /  "&'
            f'INDEX({ctx["ident"]("VIII-2")},{FIL})&"  /  "&'
            f'INDEX({ctx["ident"]("XII")},{FIL})))', f'={U_TMP}&" max."'),
           ("Grafico de presion externa", val_of("Grafico presion externa"),
            '="Subparte 3 II-D"')]
    for k in range(max(len(izq), len(der))):
        r2 = 22 + k
        if k < len(izq):
            campo(ws, r2, 1, izq[k][0], izq[k][1], izq[k][2])
        if k < len(der):
            campo(ws, r2, 7, der[k][0], der[k][1], der[k][2])

    # ------------------ 4 · TRAZABILIDAD DEL CALCULO ----------------------
    rt = 22 + max(len(izq), len(der)) + 1
    banda(ws, rt, "4 · TRAZABILIDAD DEL CALCULO   —   puntos tabulados usados en la "
                  "interpolacion")
    campo(ws, rt + 1, 1, "T1 — temperatura tabulada inferior", f"={T1C}", f"={U_TMP}")
    campo(ws, rt + 2, 1, "Valor en T1", f"={S1C}", f"={U_VAL}")
    campo(ws, rt + 1, 7, "T2 — temperatura tabulada superior", f"={T2C}", f"={U_TMP}")
    campo(ws, rt + 2, 7, "Valor en T2", f"={S2C}", f"={U_VAL}")
    ec = _mrg(ws, rt + 3, 1, NCOLS)
    ec.value = (f'=IF({FIL}="","",IF($D$13="Tabulado-conservador",'
                f'"Modo tabulado-conservador: se adopta el valor de T2.",'
                f'IF(OR({T2C}="",{S2C}=""),'
                f'"Sin punto tabulado por encima de T: se adopta el ultimo valor tabulado '
                f'(el codigo prohibe extrapolar).",'
                f'"Interpolacion lineal:  S = S1 + (S2-S1)*(T-T1)/(T2-T1)")))')
    ec.font = SRC_F
    ec.alignment = Alignment(vertical="center", indent=1)
    av = _mrg(ws, rt + 4, 1, NCOLS)
    av.value = ('="Verifique siempre la Temp. max. del material y sus notas antes de '
                'emitir el calculo."')
    av.font = SRC_F

    # --------------------------- 5 · CURVA --------------------------------
    rg = rt + 6
    banda(ws, rg, "5 · CURVA DEL MATERIAL   —   valor tabulado frente a la temperatura")
    chart_anchor = f"A{rg + 1}"

    c0 = 1 + cidx * 6
    q = f"'{ws.title}'!"
    trq = TR.replace("$D$5", q + "$D$5").replace(NP, q + NP).replace(FIL, q + FIL)
    vrq = VR.replace("$D$5", q + "$D$5").replace(NP, q + NP).replace(FIL, q + FIL)
    curvas.cell(1, c0, f"{ws.title} — datos de la curva (hoja auxiliar)").font = SRC_F
    curvas.cell(2, c0).value = f'=CONCATENATE("Temperatura, ",{q}$F$12)'
    curvas.cell(2, c0 + 1).value = (f'=CONCATENATE("{ctx["valor_lbl"]}, ",'
                                    f'IF({q}$D$5="SI","{ctx["unit_si"]}",'
                                    f'"{ctx["unit_us"]}"))')
    for k in range(1, npack + 1):
        rr = 2 + k
        curvas.cell(rr, c0).value = f'=IFERROR(INDEX({trq},{k}),NA())'
        curvas.cell(rr, c0 + 1).value = f'=IFERROR(INDEX({vrq},{k}),NA())'
    curvas.cell(2, c0 + 3, "T consulta").font = HDR_F
    curvas.cell(2, c0 + 4, "Punto consultado").font = HDR_F
    curvas.cell(3, c0 + 3).value = f"={q}$D$12"
    curvas.cell(3, c0 + 4).value = f"={q}{VALC}"

    from openpyxl.chart import Reference, Series, ScatterChart
    from openpyxl.chart.marker import Marker
    from openpyxl.drawing.line import LineProperties
    ch = ScatterChart()
    ch.title = f"{ctx['valor_lbl']} frente a la temperatura"
    ch.style = 13
    ch.x_axis.title = (f"Temperatura  [{ctx['temp_si']} en SI  ·  "
                       f"{ctx['temp_us']} en US]")
    ch.y_axis.title = (f"{ctx['valor_lbl']}  [{ctx['unit_si']} en SI  ·  "
                       f"{ctx['unit_us']} en US]")
    ch.height, ch.width = 9.5, 26
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    xs = Reference(curvas, min_col=c0, min_row=3, max_row=2 + npack)
    ys = Reference(curvas, min_col=c0 + 1, min_row=2, max_row=2 + npack)
    s1 = Series(ys, xs, title_from_data=True)
    s1.marker = Marker(symbol="circle", size=5)
    ch.series.append(s1)
    xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
    yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
    s2 = Series(yq, xq, title_from_data=True)
    s2.marker = Marker(symbol="diamond", size=10)
    s2.graphicalProperties.line = LineProperties(noFill=True)
    ch.series.append(s2)
    ws.add_chart(ch, chart_anchor)

    for cc, w in zip("ABCDEFGHIJKL",
                     [20, 20, 16, 16, 9, 3, 20, 20, 16, 16, 9, 3]):
        ws.column_dimensions[cc].width = w


def build_buscador_grupo(wb, curvas, bloques, rangos):
    """Propiedades indexadas por GRUPO de material (no por especificacion):
    modulo E, dilatacion termica, Poisson y densidad. Mismo formato de tarjeta.

    GRUPO es el rotulo normativo impreso en TM-1 / TE-1 («Material Group C»,
    «Group 3»), no la FAMILIA de db_lib, que es una agrupacion derivada solo
    para acortar listas desplegables y no interviene en ningun calculo.
    """
    ws = new_sheet(wb, "Buscar_Propiedades",
                   "BUSCADOR DE PROPIEDADES POR GRUPO DE MATERIAL — modulo E (TM-1..5), "
                   "dilatacion y modulo del Apendice C del B31.3, Poisson y densidad (PRD)",
                   "Estas tablas del codigo se indexan por GRUPO de material, no por "
                   "especificacion: elija primero la tabla y despues el grupo. Para saber "
                   "que grupo corresponde a su material consulte MAP_Grupo. Valores en "
                   "unidades metricas (edicion SI del codigo). La unica celda que se "
                   "escribe es la temperatura de consulta.")
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30
    banda(ws, 4, "PARAMETROS DE CONSULTA")
    e = _mrg(ws, 5, 1, 3, "TEMPERATURA DE CONSULTA   (unica celda de escritura)")
    e.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    e.alignment = Alignment(vertical="center", indent=1)
    c = _mrg(ws, 5, 4, 5, 25)
    c.font, c.fill = IN_F, IN_FILL
    c.border = Border(*[Side("medium", color="C00000")] * 4)
    ws.cell(5, 6, "°C").font = UNIT_F
    e = _mrg(ws, 6, 1, 3, "Modo de lectura")
    e.font = LBL_F
    e.alignment = Alignment(vertical="center", indent=1)
    c = _mrg(ws, 6, 4, 6, "Interpolado")
    c.font, c.fill, c.border = IN_F, IN_FILL, BOX
    dv_list(ws, "D6", '"Interpolado,Tabulado-conservador"')

    r = 8
    for i2, b in enumerate(bloques):
        banda(ws, r, b["titulo"])
        r += 1
        e = _mrg(ws, r, 1, 3, "Tabla / familia del codigo")
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        c = _mrg(ws, r, 4, 6, b["default_tabla"])
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        dv_list(ws, f"D{r}", "=" + rangos[b["lst_tabla"]])
        tcell = f"$D${r}"
        r += 1
        e = _mrg(ws, r, 1, 3, "Grupo / material")
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        c = _mrg(ws, r, 4, 8, "")
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        lc = 50 + i2
        LL = get_column_letter(lc)
        mx = max(1, min(int(b.get("max_grupo", 60)), 250))
        kr, vr2 = rangos[b["nm_key"]], rangos[b["nm_val"]]
        ws.cell(R_HDR, lc, "lista grupo").font = SRC_F
        for k in range(1, mx + 1):
            ws.cell(R_DATA + k - 1, lc).value = (
                f'=IF(COUNTIF({kr},{tcell})<{k},"",'
                f'INDEX({vr2},MATCH({tcell},{kr},0)+{k}-1))')
        ws.column_dimensions[LL].hidden = True
        dv_list(ws, f"D{r}", f"=${LL}${R_DATA}:${LL}${R_DATA + mx - 1}")
        gcell = f"$D${r}"
        r += 1
        info = b["info"]
        idn = f"{info['sheet']}!$D${R_DATA}:$D${info['last_row']}"
        ws.cell(r, 60).value = f'=IFERROR(MATCH({gcell},{idn},0),"")'
        frow = f"$BH${r}"
        ws.column_dimensions["BH"].hidden = True
        if b.get("temps"):
            rf = packed_refs(info)
            ws.cell(r, 61).value = f'=IF({frow}="","",IFERROR(INDEX({rf["npts"]},{frow}),0))'
            NP = f"$BI${r}"
            ws.column_dimensions["BI"].hidden = True
            tr = f'OFFSET({rf["t_anchor"]},{frow}-1,0,1,MAX(1,{NP}))'
            vr = f'OFFSET({rf["v_anchor"]},{frow}-1,0,1,MAX(1,{NP}))'
            P1 = f'IFERROR(MATCH($D$5,{tr},1),1)'
            for k2, fml in enumerate([f'=IF({frow}="","",IFERROR(INDEX({tr},{P1}),""))',
                                      f'=IF({frow}="","",IFERROR(INDEX({vr},{P1}),""))',
                                      f'=IF({frow}="","",IFERROR(INDEX({tr},{P1}+1),""))',
                                      f'=IF({frow}="","",IFERROR(INDEX({vr},{P1}+1),""))']):
                ws.cell(r, 62 + k2).value = fml
                ws.column_dimensions[get_column_letter(62 + k2)].hidden = True
            T1C, S1C = f"$BJ${r}", f"$BK${r}"
            T2C, S2C = f"$BL${r}", f"$BM${r}"
            t = _mrg(ws, r, 1, 3, b["valor_lbl"].upper() + " A LA TEMPERATURA DE CONSULTA")
            t.font, t.fill = KPI_TIT_F, BAND_FILL
            t.alignment = Alignment(vertical="center", indent=1)
            v = _mrg(ws, r, 4, 5)
            v.value = (f'=IF({frow}="","(elija grupo)",' +
                       interp_value(T1C, S1C, T2C, S2C, "$D$5", "$D$6")[1:] + ')')
            v.font, v.fill = KPI_VAL_F, KPI_FILL
            v.alignment = Alignment(horizontal="center", vertical="center")
            u = _mrg(ws, r, 6, 8, b["unidad"])
            u.font, u.fill = UNIT_F, KPI_FILL
            ws.row_dimensions[r].height = 26
            valr = r
            r += 1
            campo(ws, r, 1, "T1 — tabulada inferior", f"={T1C}", "°C")
            campo(ws, r, 7, "Valor en T1", f"={S1C}", b["unidad"])
            r += 1
            campo(ws, r, 1, "T2 — tabulada superior", f"={T2C}", "°C")
            campo(ws, r, 7, "Valor en T2", f"={S2C}", b["unidad"])
            r += 1
            n2 = _mrg(ws, r, 1, NCOLS, b.get("nota", ""))
            n2.font = SRC_F
            r += 1
            c0 = 200 + i2 * 6
            npack = info["npack"]
            q = f"'{ws.title}'!"
            trq = tr.replace(frow, q + frow).replace(NP, q + NP)
            vrq = vr.replace(frow, q + frow).replace(NP, q + NP)
            curvas.cell(1, c0, f"{b['titulo'][:38]} — curva").font = SRC_F
            curvas.cell(2, c0, "Temperatura, °C").font = HDR_F
            curvas.cell(2, c0 + 1, f'{b["valor_lbl"]}, {b["unidad"]}').font = HDR_F
            for k in range(1, npack + 1):
                curvas.cell(2 + k, c0).value = f'=IFERROR(INDEX({trq},{k}),NA())'
                curvas.cell(2 + k, c0 + 1).value = f'=IFERROR(INDEX({vrq},{k}),NA())'
            from openpyxl.chart import Reference, Series, ScatterChart
            from openpyxl.chart.marker import Marker
            ch = ScatterChart()
            ch.title = f'{b["valor_lbl"]} frente a la temperatura'
            ch.x_axis.title = "Temperatura  [°C]"
            ch.y_axis.title = f'{b["valor_lbl"]}  [{b["unidad"]}]'
            ch.height, ch.width = 7.5, 20
            ch.x_axis.delete = False
            ch.y_axis.delete = False
            xs = Reference(curvas, min_col=c0, min_row=3, max_row=2 + npack)
            ys = Reference(curvas, min_col=c0 + 1, min_row=2, max_row=2 + npack)
            se = Series(ys, xs, title_from_data=True)
            se.marker = Marker(symbol="circle", size=5)
            ch.series.append(se)
            ws.add_chart(ch, f"A{r}")
            r += 15
        else:
            for lbl, col, uni in b["campos"]:
                campo(ws, r, 1, lbl,
                      f'=IF({frow}="","—",INDEX({info["sheet"]}!'
                      f'${get_column_letter(col)}${R_DATA}:'
                      f'${get_column_letter(col)}${info["last_row"]},{frow}))', uni)
                r += 1
            r += 1
    for cc, w in zip("ABCDEFGHIJKL", [22, 22, 18, 16, 12, 12, 22, 22, 16, 12, 12, 4]):
        ws.column_dimensions[cc].width = w
    return ws


def build_buscador_nm(wb, info, rangos, max_mat=60):
    ws = new_sheet(wb, "Buscar_NoMetalicos",
                   "BUSCADOR — Materiales no metalicos (ASME B31.3, Apendices B y C)",
                   "Elija la tabla y despues el material: se muestran todos los campos "
                   "impresos para ese material, con las unidades que emplea el codigo. "
                   "Todo por lista desplegable.")
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    banda(ws, 4, "1 · SELECCION")
    for r, et, val, ref in ((5, "Tabla del Apendice", None, rangos["NM_TABLA"]),
                            (6, "Material", "", None)):
        e = _mrg(ws, r, 1, 3, et)
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        c = _mrg(ws, r, 4, 8, val)
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
    ws["D5"] = wb["DB_NoMetalicos"].cell(R_DATA, 2).value
    dv_list(ws, "D5", "=" + rangos["NM_TABLA"])
    mx = max(1, min(int(max_mat), 250))
    ws.cell(R_HDR, 50, "lista material").font = SRC_F
    for k in range(1, mx + 1):
        ws.cell(R_DATA + k - 1, 50).value = (
            f'=IF(COUNTIF({rangos["NM_MATK"]},$D$5)<{k},"",'
            f'INDEX({rangos["NM_MATV"]},MATCH($D$5,{rangos["NM_MATK"]},0)+{k}-1))')
    ws.column_dimensions["AX"].hidden = True
    dv_list(ws, "D6", f"=$AX${R_DATA}:$AX${R_DATA + mx - 1}")
    last = info["last_row"]
    sf = f'DB_NoMetalicos!$C${R_DATA}:$C${last}'
    key = '$D$5&"|"&$D$6'
    ws["BH1"] = f'=IFERROR(MATCH({key},{sf},0),0)'
    ws["BH2"] = f'=COUNTIF({sf},{key})'
    ws.column_dimensions["BH"].hidden = True
    banda(ws, 8, "2 · FICHA DEL MATERIAL   —   campos tal como los imprime el codigo")
    t = _mrg(ws, 9, 1, NCOLS)
    t.value = (f'=IFERROR("Tabla "&$D$5&" — "&INDEX(DB_NoMetalicos!$G${R_DATA}:$G${last},'
               f'MATCH({key},{sf},0)),"Seleccione tabla y material")')
    t.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    t.fill = CARD_FILL
    t.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[9].height = 26
    for k in range(1, 17):
        r = 9 + k
        cond = f'IF(OR($BH$1=0,{k}>$BH$2),""'
        campo(ws, r, 1,
              f'={cond},INDEX(DB_NoMetalicos!$E${R_DATA}:$E${last},$BH$1+{k}-1))',
              f'={cond},INDEX(DB_NoMetalicos!$F${R_DATA}:$F${last},$BH$1+{k}-1))')
        _mrg(ws, r, 6, NCOLS)
    av = _mrg(ws, 27, 1, NCOLS)
    av.value = ('="Las unidades son las que imprime cada tabla: HDS en MPa (B-1) o ksi '
                '(B-1C); presiones en kPa y psi; temperaturas en °C y °F; dilatacion en '
                'mm/mm·°C e in/in·°F."')
    av.font = SRC_F
    for cc, w in zip("ABCDEFGHIJKL", [26, 26, 24, 24, 14, 3, 14, 14, 14, 14, 14, 4]):
        ws.column_dimensions[cc].width = w
    return ws


# ---------------------------------------------------------------------------
# Integracion en el motor Parche_PCC2_Art212
# ---------------------------------------------------------------------------
MOTOR = "Parche_PCC2_Art212"


def integrate_motor(wb, b313, iid1a, iidb, fac_info, rangos):
    """Conecta el motor a las bases. Cascada de 5 niveles + variante, con las
    listas materializadas en columnas ocultas (unica forma portable a Google
    Sheets) y consulta sobre la banda compacta: sin formulas matriciales."""
    ws = wb[MOTOR]
    rb, ri = rangos["B313"], rangos["IID1A"]
    B, I1, I2 = b313["sheet"], iid1a["sheet"], iidb["sheet"]
    IDC, TMAXC = CL["material_id"], CL["Temp. max. / limite"]

    def col3(cl_):
        return (f'CHOOSE({{i}},{B}!${cl_}${R_DATA}:${cl_}${b313["last_row"]},'
                f'{I1}!${cl_}${R_DATA}:${cl_}${iid1a["last_row"]},'
                f'{I2}!${cl_}${R_DATA}:${cl_}${iidb["last_row"]})')
    IDS, TMAX = col3(IDC), col3(TMAXC)
    NAMES = f'CHOOSE({{i}},"{B}","{I1}","{I2}")'
    pk = {1: packed_refs(b313), 2: packed_refs(iid1a), 3: packed_refs(iidb)}
    NPTS = 'CHOOSE({i},' + ",".join(pk[k]["npts"] for k in (1, 2, 3)) + ')'
    TANC = 'CHOOSE({i},' + ",".join(pk[k]["t_anchor"] for k in (1, 2, 3)) + ')'
    VANC = 'CHOOSE({i},' + ",".join(pk[k]["v_anchor"] for k in (1, 2, 3)) + ')'
    hl = get_column_letter

    def lab(r, text, note=None):
        ws.cell(r, 1, text).font = Font(name="Calibri", size=10)
        if note:
            ws.cell(r, 7, note).font = Font(name="Calibri", size=9, italic=True,
                                            color="595959")

    def inp(cell, value=""):
        c = ws[cell]
        c.value = value
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        c.protection = Protection(locked=False)

    ws.cell(105, 1, "7.  RESOLUCION DE MATERIAL — BASE DE DATOS ASME "
                    "(cascada de listas desplegables)")
    ws.cell(105, 1).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws.cell(105, 1).fill = TITLE_FILL
    for j2, h in enumerate(["Parametro", "", "Unidad", "Metal base", "Collar / parche",
                            "", "Referencia / Notas"], start=1):
        c = ws.cell(106, j2, h)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=GREY)
    lab(107, "Modo de lectura de S(T)   [MODO_S]", "Interpolado | Tabulado-conservador")
    inp("D107", "Interpolado")
    dv_list(ws, "D107", '"Interpolado,Tabulado-conservador"')
    lab(108, "Temperatura de evaluacion", "= T de la hoja de proceso (D25)")
    ws.cell(108, 4).value = "=$D$25"
    ws.cell(108, 3).value = "°C"
    niveles = [(109, "0 · Familia de material"), (110, "1 · Composicion nominal"),
               (111, "2 · Forma de producto"), (112, "3 · Especificacion (Spec. No.)"),
               (113, "4 · Tipo / Grado"),
               (114, "5 · Variante (clase / tamano) — opcional")]
    for r, t in niveles:
        lab(r, t, "Lista desplegable en cascada")
        inp(f"D{r}")
        inp(f"E{r}")

    # --- listas materializadas (columnas ocultas) --------------------------
    FAM_COL = 11
    ws.cell(R_HDR, FAM_COL, "lista familia").font = SRC_F
    for k in range(1, 41):
        ws.cell(R_DATA + k - 1, FAM_COL).value = (
            f'=IF($D$11=1,IFERROR(INDEX({rb["FAM"]},{k}),""),'
            f'IFERROR(INDEX({ri["FAM"]},{k}),""))')
    ws.column_dimensions[hl(FAM_COL)].hidden = True
    fam_ref = f'=${hl(FAM_COL)}${R_DATA}:${hl(FAM_COL)}${R_DATA + 39}'
    dv_list(ws, "D109", fam_ref)
    dv_list(ws, "E109", fam_ref)

    lv = [("C", "CK", "CV", 110, '{c}$109'),
          ("F", "FK", "FV", 111, '{c}$109&"|"&{c}$110'),
          ("S", "SK", "SV", 112, '{c}$109&"|"&{c}$110&"|"&{c}$111'),
          ("G", "GK", "GV", 113, '{c}$109&"|"&{c}$110&"|"&{c}$111&"|"&{c}$112'),
          ("V", "K4", "ID", 114,
           '{c}$109&"|"&{c}$110&"|"&{c}$111&"|"&{c}$112&"|"&{c}$113')]
    col = FAM_COL + 1
    for who, cl in (("base", "$D"), ("collar", "$E")):
        for tag, kk, vv, target, keyfmt in lv:
            key = keyfmt.format(c=cl)
            mx = max(1, min(int(max(rb.get("max" + tag, rb["maxV"]),
                                    ri.get("max" + tag, ri["maxV"]))), 250))
            L = hl(col)
            ws.cell(R_HDR, col, f"lista {tag} {who}").font = SRC_F
            for k in range(1, mx + 1):
                ws.cell(R_DATA + k - 1, col).value = (
                    f'=IF($D$11=1,'
                    f'IF(COUNTIF({rb[kk]},{key})<{k},"",'
                    f'INDEX({rb[vv]},MATCH({key},{rb[kk]},0)+{k}-1)),'
                    f'IF(COUNTIF({ri[kk]},{key})<{k},"",'
                    f'INDEX({ri[vv]},MATCH({key},{ri[kk]},0)+{k}-1)))')
            ws.column_dimensions[L].hidden = True
            dv_list(ws, f'{"D" if cl == "$D" else "E"}{target}',
                    f"=${L}${R_DATA}:${L}${R_DATA + mx - 1}")
            col += 1

    aux = {"D": "W", "E": "X"}
    for cl, ac in aux.items():
        ws[f"{ac}109"] = (f'=${cl}$109&"|"&${cl}$110&"|"&${cl}$111&"|"&'
                          f'${cl}$112&"|"&${cl}$113')
        ws[f"{ac}110"] = (f'=IF($D$11=1,IFERROR(MATCH(${ac}$109,{rb["K4"]},0),0),'
                          f'IFERROR(MATCH(${ac}$109,{ri["K4"]},0),0))')
        ws.column_dimensions[ac].hidden = True

    etiquetas = [(115, "material_id resuelto", None), (116, "Base de datos activa", None),
                 (117, "Indice de base (1/2/3)", None), (118, "Fila localizada", None),
                 (119, "n_pts (puntos tabulados) / p1", None),
                 (120, "T1 — temperatura tabulada inferior", "°C"),
                 (121, "T2 — temperatura tabulada superior", "°C"),
                 (122, "S en T1", "MPa"), (123, "S en T2", "MPa"),
                 (124, "Temp. max. admisible / limite VIII-1", "°C"),
                 (125, "Dictamen de rango", None),
                 (126, "S(T) resuelto", "MPa")]
    for r, t, u in etiquetas:
        lab(r, t)
        if u:
            ws.cell(r, 3).value = u
    for col2 in (4, 5):
        L = hl(col2)
        ac = aux[L]
        ws.cell(115, col2).value = (
            f'=IF(${L}$114<>"",${L}$114,IF(${ac}$110=0,"",'
            f'IF($D$11=1,INDEX({rb["ID"]},${ac}$110),INDEX({ri["ID"]},${ac}$110))))')
        ws.cell(117, col2).value = f'=IF($D$11=1,1,IF(LEFT({L}115,2)="1A",2,3))'
        ic = f"{L}117"
        ws.cell(116, col2).value = f'=IF({L}115="","",{NAMES.format(i=ic)})'
        ws.cell(118, col2).value = f'=IFERROR(MATCH({L}115,{IDS.format(i=ic)},0),"")'
        F = f"{L}118"
        ws.cell(119, col2).value = f'=IF({F}="","",IFERROR(INDEX({NPTS.format(i=ic)},{F}),0))'
        NP = f"{L}119"
        tr = f'OFFSET({TANC.format(i=ic)},{F}-1,0,1,MAX(1,{NP}))'
        vr = f'OFFSET({VANC.format(i=ic)},{F}-1,0,1,MAX(1,{NP}))'
        P1 = f'IFERROR(MATCH($D$108,{tr},1),1)'
        ws.cell(120, col2).value = f'=IF({F}="","",IFERROR(INDEX({tr},{P1}),""))'
        ws.cell(121, col2).value = f'=IF({F}="","",IFERROR(INDEX({tr},{P1}+1),""))'
        ws.cell(122, col2).value = f'=IF({F}="","",IFERROR(INDEX({vr},{P1}),""))'
        ws.cell(123, col2).value = f'=IF({F}="","",IFERROR(INDEX({vr},{P1}+1),""))'
        ws.cell(124, col2).value = f'=IFERROR(INDEX({TMAX.format(i=ic)},{F}),"")'
        ws.cell(125, col2).value = (
            f'=IF({L}115="","SIN MATERIAL SELECCIONADO",'
            f'IF({F}="","MATERIAL NO ENCONTRADO",'
            f'IF({L}122="","FUERA DE RANGO (sin valor tabulado a esa T)",'
            f'IF(AND(ISNUMBER({L}124),$D$108>{L}124),'
            f'"FUERA DE RANGO (T > Temp. max.)","OK"))))')
        ws.cell(126, col2).value = (
            f'=IF({L}125<>"OK",NA(),'
            f'IF($D$108<={L}120,{L}122,IF(OR({L}121="",{L}123=""),{L}122,'
            f'IF($D$107="Tabulado-conservador",{L}123,'
            f'{L}122+({L}123-{L}122)*($D$108-{L}120)/({L}121-{L}120)))))')
        ws.cell(126, col2).font = Font(bold=True)

    fl, fa = fac_info["sheet"], fac_info["last_row"]
    lab(128, "Clave de junta longitudinal (Tabla A-3) -> Ej",
        "Lista desplegable de MAP_Factores")
    inp("D128", "A106 | Seamless pipe")
    dv_list(ws, "D128", f"={fl}!$A${R_DATA}:$A${fa}")
    ws.cell(128, 3).value = "adimensional"
    ws.cell(128, 5).value = (f'=IFERROR(INDEX({fl}!$G${R_DATA}:$G${fa},'
                             f'MATCH($D$128,{fl}!$A${R_DATA}:$A${fa},0)),1)')
    lab(129, "Ec (fundicion, Tabla A-2) — informativo")
    inp("D129")
    dv_list(ws, "D129", f"={fl}!$A${R_DATA}:$A${fa}")
    ws.cell(129, 3).value = "adimensional"
    ws.cell(129, 5).value = (f'=IFERROR(INDEX({fl}!$G${R_DATA}:$G${fa},'
                             f'MATCH($D$129,{fl}!$A${R_DATA}:$A${fa},0)),"")')

    ws["D39"] = '=IF($E$125="OK",$E$126,NA())'
    ws["G39"] = "S(T) del collar — base de datos ASME (seccion 7)"
    ws["D40"] = '=IF($D$125="OK",$D$126,NA())'
    ws["G40"] = "S(T) del metal base — base de datos ASME (seccion 7)"
    ws["D44"] = "=$E$128"
    ws["G44"] = "Ej por lookup (B31.3 Tabla A-3) — seccion 7"
    ws["D44"].font = Font(name="Calibri", size=10)
    ws["D44"].fill = PatternFill(fill_type=None)
    ws["D44"].protection = Protection(locked=True)
    ws["F90"] = ('=IF(OR($D$125<>"OK",$E$125<>"OK"),"REVISAR — MATERIAL FUERA DE RANGO",'
                 'IF(AND(F84="CUMPLE",F85="CUMPLE",F86="CUMPLE",F87="CUMPLE"),"APTO","REVISAR"))')
    ws.cell(131, 1, "La cascada filtra la base ASME por familia, composicion nominal, "
                    "forma de producto, especificacion y tipo/grado. Si un material no "
                    "aparece, localicelo en el buscador correspondiente y pegue su "
                    "material_id en la celda 'Variante'. La lista corta de Datos_Ref queda "
                    "como respaldo historico y ya no alimenta el calculo.").font = SRC_F
    ws.protection.password = "0000"
    ws.protection.sheet = True


def deprecate_datos_ref(wb):
    ws = wb["Datos_Ref"]
    ws["A40"] = ("B. [OBSOLETO — ver DB_B31_3 / DB_BPVC_IID]  ESFUERZOS ADMISIBLES S (MPa) — "
                 "acero al carbono, T <= 40 C. Se conserva como respaldo historico de las "
                 "memorias ya emitidas; el motor ya NO lee de aqui.")
    ws["A40"].font = Font(name="Calibri", size=10, bold=True, color="C00000")
    for r in range(41, 48):
        for c in range(1, 4):
            ws.cell(r, c).font = Font(name="Calibri", size=10, color="808080", italic=True)


# ---------------------------------------------------------------------------
# Instrucciones y _meta
# ---------------------------------------------------------------------------
INSTRUCCIONES = [
    ("Objeto",
     "Automatiza los calculos y las especificaciones tecnicas del ciclo de integridad de "
     "equipos a presion y tuberia bajo la familia ASME PCC, para tuberia (ASME B31.3) y "
     "recipientes (ASME BPVC VIII-1). Desde la Rev. 2 el esfuerzo admisible y las "
     "propiedades proceden de bases de datos completas y dependientes de la temperatura, "
     "extraidas de los codigos en resources/."),
    ("1. Como se consulta un material",
     "Todos los buscadores y el motor funcionan igual, con una CASCADA de listas "
     "desplegables:\n"
     "  1 · Composicion nominal (Carbon steel, 18Cr-8Ni, Nickel...)\n"
     "  2 · Forma de producto (Smls. pipe, Plate, Forgings...)\n"
     "  3 · Especificacion (Spec. No.) — ya filtrada por los dos pasos anteriores\n"
     "  4 · Tipo / Grado — se habilita al elegir la especificacion\n"
     "  5 · Variante (clase / tamano) — opcional, solo si el codigo repite ese grado\n"
     "La UNICA celda donde se escribe es la TEMPERATURA DE CONSULTA (borde rojo). Todo lo "
     "demas se elige de una lista, de modo que no es posible teclear un material que no "
     "exista en el codigo."),
    ("2. Que devuelve el buscador",
     "Primero, encima de la ficha, la TABLA DE RESULTADOS: todas las variantes de la "
     "especificacion elegida (grado, clase, tamano, UNS, P-No., resistencia, fluencia, "
     "temperatura maxima) con el valor interpolado a la temperatura de consulta para cada "
     "una — asi se comparan de un vistazo. Debajo, la FICHA del material seleccionado (el "
     "material_id se carga solo) con toda su identificacion tal como la organiza el codigo, "
     "el valor a la temperatura de consulta y la CURVA del material frente a la temperatura. "
     "La curva se muestra como grafica, no como tabla de valores; el punto en forma de rombo "
     "marca la temperatura consultada."),
    ("3. Mapa de bases de datos",
     "DB_B31_3 / DB_B31_3C — ASME B31.3-2024, Apendice A: Tabla A-1 (esfuerzos admisibles "
     "basicos en traccion) + Tabla A-4 (perneria). SI y US nativos.\n"
     "DB_BPVC_IID / DB_BPVC_IIDC — ASME BPVC II-D 2025, Tabla 1A (ferrosos).\n"
     "DB_BPVC_IID_B / DB_BPVC_IID_BC — Tablas 1B (no ferrosos) y 3 (perneria). Van aparte "
     "porque la 1A no imprime 175/225/275 C y la 1B/3 si; mezclarlas dejaria huecos en la "
     "banda de valores y romperia la interpolacion.\n"
     "DB_Su / DB_Sy — Tablas U y Y-1 (resistencia a la traccion y fluencia por temperatura).\n"
     "DB_E — Tablas TM-1..5 (modulo E) · DB_TE — Tablas TE-1..5 (dilatacion) · DB_PRD — "
     "Poisson y densidad. Estas tres se indexan por GRUPO de material (el rotulo impreso "
     "en TM-1 / TE-1, no la familia de navegacion): ver MAP_Grupo.\n"
     "DB_C_dilatacion / DB_C_modulo — Apendice C del B31.3 (lado tuberia).\n"
     "DB_NoMetalicos — Apendices B y C-2/C-4 (termoplasticos, RTR, concreto, vidrio).\n"
     "MAP_Factores (Ej/Ec) · MAP_Grupo · Notas_Codigo · DB_Listas · _meta (trazabilidad)."),
    ("4. Conmutador de unidades SI / US",
     "Cada buscador tiene la celda 'Sistema de unidades'. En SI lee la tabla metrica del "
     "codigo (MPa, C); en US lee la tabla nativa en unidades inglesas (ksi, F): el B31.3 usa "
     "sus tablas companion 'C' y la II-D su edicion U.S. Customary 2025. NO hay conversiones. "
     "La seleccion se hace siempre sobre la edicion metrica y se enlaza con la US mediante "
     "una clave de identidad; si un material no tiene homologo, la ficha lo indica. Regla: no "
     "mezclar sistemas dentro de un mismo calculo. El motor opera siempre en SI."),
    ("5. El motor de calculo",
     "En la seccion 7 del modulo hay la misma cascada, por separado para el metal base y para "
     "el collar/parche. El selector 'Aplicacion / Codigo' (D10) decide de que base se lee: "
     "B31.3 en modo tuberia, II-D Tabla 1A en modo recipiente. La seccion muestra la fila "
     "localizada, T1/T2, S en T1 y T2, la temperatura maxima admisible, el dictamen de rango "
     "y el S(T) resuelto, de forma que el calculo es auditable paso a paso."),
    ("6. Temperatura: interpolado o tabulado",
     "MODO_S = 'Interpolado' (por defecto) devuelve S por interpolacion lineal entre las dos "
     "temperaturas tabuladas que enmarcan T: S = S1 + (S2-S1)(T-T1)/(T2-T1). "
     "'Tabulado-conservador' devuelve el valor de la temperatura tabulada inmediatamente "
     "superior. Por debajo de la minima tabulada se usa el valor minimo; por encima de la "
     "Temp. max. del material el motor devuelve error controlado y dictamen 'FUERA DE RANGO'. "
     "NO se extrapola: el codigo lo prohibe."),
    ("7. Factores Ej y Ec",
     "El antiguo E_j fijo = 1,0 se sustituyo por un lookup sobre MAP_Factores (Tabla A-3). "
     "Elija en la lista la junta que corresponde al tubo real (p. ej. 'A106 | Seamless pipe' "
     "= 1,0; 'A53 | Type E | Electric resistance welded pipe' = 0,85). Ec (Tabla A-2) queda "
     "como informativo para componentes fundidos. t_req usa S x Ej."),
    ("8. Compatibilidad",
     "El libro NO usa funciones de matriz dinamica (FILTER, SORT, UNIQUE, XLOOKUP). Toda la "
     "logica es INDEX / MATCH / OFFSET / COUNTIF y listas desplegables clasicas, de modo que "
     "funciona en cualquier version de Excel de escritorio. Si una lista desplegable no se "
     "abre, compruebe que el libro se abrio en Excel de escritorio y no en un visor."),
    ("9. Trazabilidad y limitaciones",
     "La hoja _meta indica, para cada base, el archivo de resources/ del que procede, la "
     "tabla, la edicion, el sistema de unidades y el numero de filas, mas las limitaciones "
     "detectadas. Las hojas DB contienen solo valores (sin formulas): cualquier dato puede "
     "contrastarse contra el PDF del codigo. La tipografia cursiva/negrita de las tablas A-1C "
     "no viaja en la extraccion JSON: consulte la nota del material en Notas_Codigo. En "
     "MAP_Grupo toda fila con grupo asignado cita su fuente: el UNS impreso en TM-1..TM-5, "
     "o la Nota de TM-1 / TE-1 que enumera esa composicion nominal. Las filas 'REVISAR' "
     "corresponden a una regla de inclusion que redacta el codigo y debe aplicar el "
     "ingeniero; las 'SIN MAPEO' son materiales para los que II-D no publica E ni "
     "dilatacion, y en ellas el calculo queda bloqueado."),
    ("10. Convenciones",
     "Calculo en SI por defecto. Celdas de seleccion: texto azul sobre fondo amarillo; la "
     "celda de ENTRADA de temperatura lleva ademas borde rojo. Hojas de calculo protegidas "
     "con contrasena preliminar 0000; las bases y los buscadores quedan sin proteger para "
     "poder filtrar y copiar. Verde = CUMPLE/APTO, rojo = NO CUMPLE/revisar."),
    ("Aviso",
     "Herramienta de ingenieria de referencia. Verificar entradas y resultados; complementar "
     "con WPS/PQR, ATS/JSA y registros del propietario. Revision y aprobacion por personal "
     "calificado antes de ejecutar."),
]


def rewrite_instrucciones(wb, version_note):
    ws = wb["Instrucciones"]
    ws.protection.sheet = False
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=8):
        for c in row:
            c.value = None
    ws["B2"] = "MOTOR DE CALCULO — ASME PCC (CODIGOS POST-CONSTRUCCION)"
    ws["B2"].font = Font(name="Calibri", size=14, bold=True, color=NAVY)
    ws["B3"] = ("Reparacion, montaje e inspeccion de equipos a presion y tuberia segun la "
                "familia ASME PCC · Codigos de construccion: ASME B31.3-2024 y ASME BPVC "
                "Seccion VIII-1 con II-D 2025")
    ws["B3"].font = Font(name="Calibri", size=10, italic=True, color="595959")
    ws["B4"] = version_note
    ws["B4"].font = Font(name="Calibri", size=10, bold=True, color="C00000")
    r = 6
    for title, body in INSTRUCCIONES:
        ws.cell(r, 2, title).font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        c = ws.cell(r, 3, body)
        c.font = Font(name="Calibri", size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = max(30, 13 * (body.count("\n") + 1 + len(body) // 110))
        r += 1
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 118
    ws.protection.password = "0000"
    ws.protection.sheet = True


def build_meta(wb, counts):
    ws = new_sheet(wb, "_meta", "TRAZABILIDAD DE LAS BASES DE DATOS — PLAN-DB-MAT-001 Rev.2",
                   "Cada base referencia el archivo de resources/ del que procede "
                   "(regla 5.4 de knowledge/claude.md). Al cambiar de edicion del codigo, "
                   "re-ejecutar build_db_materiales.py sobre los nuevos JSON.")
    n = write_headers(ws, ["Hoja", "Tabla", "Archivo fuente (resources/)", "Edicion",
                           "Sistema", "Filas", "Nota"])
    last = write_rows(ws, [([m["hoja"], m["tabla"], m["archivo"], m["edicion"],
                             m["sistema"], m["filas"], m["nota"]], {}) for m in META], n)
    r = last + 2
    ws.cell(r, 1, "VERIFICACION DE CONTEOS").font = Font(bold=True, color=NAVY)
    for k, v in counts.items():
        r += 1
        ws.cell(r, 1, k)
        ws.cell(r, 3, v)
    r += 2
    ws.cell(r, 1, "LIMITACIONES Y OBSERVACIONES").font = Font(bold=True, color=NAVY)
    for msg in ISSUES:
        r += 1
        ws.cell(r, 1, msg).alignment = Alignment(wrap_text=True)
    autosize(ws, {"A": 24, "B": 26, "C": 62, "D": 10, "E": 10, "F": 10, "G": 60})
    return last


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def uniques(ws, last_row, key_col, val_col):
    seen, out = set(), []
    for r in range(R_DATA, last_row + 1):
        k, v = ws.cell(r, key_col).value, ws.cell(r, val_col).value
        if (k, v) in seen:
            continue
        seen.add((k, v))
        out.append((k, v))
    return out


# ---------------------------------------------------------------------------
# Dashboard y capa de navegacion (Rev. 3)
# ---------------------------------------------------------------------------
# El libro se entrega como .xlsm: la unica hoja visible es el Dashboard y la
# macro conmuta la visibilidad del resto. Los estados se graban ademas en el
# archivo, de modo que quien abra con las macros bloqueadas siga sin ver
# ninguna base de datos.
DASH = "Dashboard"

# Hojas que el usuario puede llegar a abrir. Todo lo demas queda veryHidden.
# Esta lista DEBE coincidir con HojasNavegables() de vba/mod_nav.vba.
NAVEGABLES = ["Parche_PCC2_Art212", "Buscar_B31_3", "Buscar_BPVC_IID",
              "Buscar_BPVC_IID_B", "Buscar_Su", "Buscar_Sy", "Buscar_Propiedades",
              "Buscar_NoMetalicos", "Instrucciones"]

# La clave de destino de cada boton se guarda oculta en (fila del boton,
# COL_CLAVE_BASE + columna del boton). Depende de la columna, y no solo de la
# fila, porque el Dashboard pone tres tarjetas por banda: con una unica
# columna de claves las tres escribirian en la misma celda.
#
# La base es 66 (BN) porque BM es la columna mas alta que usa cualquier hoja
# navegable. No se usa N porque Parche_PCC2_Art212 ya ocupa K..P con sus
# listas de cascada.
# DEBE coincidir con COL_CLAVE_BASE de vba/mod_nav.vba.
COL_CLAVE_BASE = 66
CLAVE_VOLVER = "VOLVER"

# Celda del aviso de macros. DEBE coincidir con CELDA_AVISO de vba/mod_nav.vba.
FILA_AVISO = 4

# Celda donde cada hoja navegable lleva su enlace de retorno. Se elige por
# hoja porque los tres tipos de layout diferen: los buscadores y el motor
# fusionan la fila 1 (y la 2) y congelan en A4, dejando la fila 3 libre;
# Instrucciones no congela y empieza en B2, dejando libre la fila 1.
ANCLA_VOLVER = {n: (3, 1, 3) for n in NAVEGABLES}    # (fila, col_ini, col_fin)
ANCLA_VOLVER["Instrucciones"] = (1, 2, 3)

DASH_NCOLS = 12
DASH_ANCHO_COL = 15
BTN_FILL = PatternFill("solid", fgColor=BLUE)
BTN_F = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
CARD_TIT_F = Font(name="Calibri", size=11, bold=True, color=NAVY)
CARD_TXT_F = Font(name="Calibri", size=9, color="44546A")
AVISO_ROJO_F = Font(name="Calibri", size=11, bold=True, color="9C0006")
AVISO_ROJO_FILL = PatternFill("solid", fgColor="FFC7CE")
PIE_F = Font(name="Calibri", size=9, italic=True, color="595959")
PIE_FILL = PatternFill("solid", fgColor=GREY)
KPI_AMBAR_F = Font(name="Calibri", size=18, bold=True, color="BF8F00")


def _boton(ws, fila, c1, c2, texto, clave):
    """Celda-boton con hipervinculo inocuo y la clave de destino en COL_CLAVE.

    El hipervinculo apunta siempre a Dashboard!A1 y no navega por si mismo:
    solo existe para que Excel dispare Workbook_SheetFollowHyperlink. Un
    hipervinculo directo a la hoja destino seria invalido, porque la hoja
    esta oculta cuando se hace clic.
    """
    b = _mrg(ws, fila, c1, c2, texto)
    b.hyperlink = Hyperlink(ref=b.coordinate, location=f"'{DASH}'!A1", display=texto)
    # El estilo se aplica DESPUES del hipervinculo: al asignarlo, Excel pinta
    # la celda con el estilo "Hyperlink" (azul subrayado) y taparia el boton.
    b.font, b.fill = BTN_F, BTN_FILL
    b.alignment = Alignment(horizontal="center", vertical="center")
    _celda_clave(ws, fila, c1).value = clave
    return b


def _celda_clave(ws, fila, columna):
    """Celda oculta con la clave de destino del boton anclado en (fila, columna)."""
    return ws.cell(fila, COL_CLAVE_BASE + columna)


def _ocultar_columnas_clave(ws, cols_boton):
    for c in sorted({COL_CLAVE_BASE + c for c in cols_boton}):
        ws.column_dimensions[get_column_letter(c)].hidden = True


def _tarjeta(ws, fila, c1, ancho, titulo, lineas, clave):
    """Tarjeta de 4 filas: titulo, dos lineas de detalle y el boton ABRIR."""
    c2 = c1 + ancho - 1
    t = _mrg(ws, fila, c1, c2, titulo)
    t.font = CARD_TIT_F
    t.alignment = Alignment(vertical="center", indent=1)
    for i, linea in enumerate(lineas[:2]):
        d = _mrg(ws, fila + 1 + i, c1, c2, linea)
        d.font = CARD_TXT_F
        d.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    _boton(ws, fila + 3, c1, c2, "▸ ABRIR", clave)
    for r in range(fila, fila + 4):
        for c in range(c1, c2 + 1):
            cel = ws.cell(r, c)
            if r != fila + 3:
                cel.fill = CARD_FILL
            cel.border = CARD_BORDER
    ws.row_dimensions[fila].height = 20
    ws.row_dimensions[fila + 1].height = 14
    ws.row_dimensions[fila + 2].height = 14
    ws.row_dimensions[fila + 3].height = 20


_RE_TXT_FORMULA = re.compile(r'^="((?:[^"]|"")*)"$', re.DOTALL)
_RE_LONGTEXT = re.compile(r'^=_xlfn\._LONGTEXT\((.*)\)$', re.DOTALL)
_RE_ARG = re.compile(r'"((?:[^"]|"")*)"')


def normalizar_textos_como_formula(wb):
    """Convierte a texto plano las celdas que guardan una cadena como formula.

    El maestro trae notas escritas como `="texto largo..."`. Excel solo admite
    255 caracteres en un literal de cadena dentro de una formula, asi que al
    reguardar el archivo -cosa que hace make_vba_seed.py al convertirlo a
    .xlsm- lo parte en `_xlfn._LONGTEXT("trozo1","trozo2")`. Esa funcion no
    existe fuera de Excel 365: en Google Sheets, en Excel de escritorio antiguo
    y en LibreOffice la celda muestra #NAME?, y ademas dispara la deteccion de
    `_xlfn` del protocolo de verificacion.

    Una nota es texto, no una formula. Se guarda como texto y el problema
    desaparece en origen, sin depender de que version de Excel toque el archivo.
    """
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if not isinstance(c.value, str) or not c.value.startswith("="):
                    continue
                m = _RE_TXT_FORMULA.match(c.value)
                if m:
                    c.value = m.group(1).replace('""', '"')
                    n += 1
                    continue
                m = _RE_LONGTEXT.match(c.value)
                if m:
                    c.value = "".join(a.replace('""', '"')
                                      for a in _RE_ARG.findall(m.group(1)))
                    n += 1
    if n:
        ISSUES.append(f"Normalizadas {n} celdas que guardaban texto como formula "
                      f"(evita _xlfn._LONGTEXT y el #NAME? fuera de Excel 365).")
    return n


def build_dashboard(wb, kpis, fecha):
    """Portada unica del libro. `kpis` es una lista de (titulo, valor, unidad, ambar)."""
    ws = wb.create_sheet(DASH)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for c in range(1, DASH_NCOLS + 1):
        ws.column_dimensions[get_column_letter(c)].width = DASH_ANCHO_COL
    # Las tarjetas se anclan en las columnas 1, 5 y 9 (tres por banda).
    _ocultar_columnas_clave(ws, (1, 5, 9))

    # --- cabecera ---------------------------------------------------------
    t = _mrg(ws, 1, 1, DASH_NCOLS, "MOTOR DE CALCULO ASME PCC          Rev. 3")
    t.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    t.fill = TITLE_FILL
    t.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 32
    s = _mrg(ws, 2, 1, DASH_NCOLS,
             "Ingenieria de reparacion · ASME PCC-2 · B31.3-2024 · BPVC VIII-1 con II-D 2025 · "
             "unidades SI")
    s.font = Font(name="Calibri", size=9, italic=True, color="FFFFFF")
    s.fill = TITLE_FILL
    s.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6

    # Aviso de macros. Se graba en rojo: es el estado correcto para un archivo
    # en disco. Workbook_Open lo pasa a verde solo si las macros corren.
    av = _mrg(ws, FILA_AVISO, 1, DASH_NCOLS,
              "MACROS DESHABILITADAS - habilitelas para navegar entre los motores")
    av.font, av.fill = AVISO_ROJO_F, AVISO_ROJO_FILL
    av.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[FILA_AVISO].height = 22
    ws.row_dimensions[FILA_AVISO + 1].height = 6

    # --- 1. motores de calculo -------------------------------------------
    r = FILA_AVISO + 2
    banda(ws, r, "1 · MOTORES DE CALCULO — ASME PCC-2", DASH_NCOLS)
    r += 1
    _tarjeta(ws, r, 1, 4, "ART. 212 · PARCHE DE PLANCHA",
             ["Parche con soldadura de filete, PCC-2 Art. 212/206",
              "Tuberia B31.3 · virola BPVC VIII-1"],
             "Parche_PCC2_Art212")
    r += 5

    # --- 2. motores de busqueda ------------------------------------------
    banda(ws, r, "2 · MOTORES DE BUSQUEDA — bases normativas", DASH_NCOLS)
    r += 1
    buscadores = [
        ("B31.3 · TABLAS A-1 y A-4", ["Esfuerzo admisible S", "MPa (SI) y ksi (US)"],
         "Buscar_B31_3"),
        ("BPVC II-D · TABLA 1A", ["Esfuerzo admisible S, ferrosos", "MPa (SI) y ksi (US)"],
         "Buscar_BPVC_IID"),
        ("BPVC II-D · TABLAS 1B y 3", ["Esfuerzo admisible S, no ferrosos", "MPa (SI) y ksi (US)"],
         "Buscar_BPVC_IID_B"),
        ("BPVC II-D · TABLA U", ["Resistencia a la traccion Su", "MPa (SI) y ksi (US)"],
         "Buscar_Su"),
        ("BPVC II-D · TABLA Y-1", ["Limite de fluencia Sy", "MPa (SI) y ksi (US)"],
         "Buscar_Sy"),
        ("PROPIEDADES POR FAMILIA", ["Modulo E, dilatacion, Poisson, densidad",
                                     "TM-1..5 · TE · PRD · B31.3 Ap. C"],
         "Buscar_Propiedades"),
        ("B31.3 · APENDICES B y C", ["Materiales no metalicos", "Temperaturas admisibles"],
         "Buscar_NoMetalicos"),
        ("MANUAL DE USO", ["Convenciones, alcance y limitaciones", "Leer antes de calcular"],
         "Instrucciones"),
    ]
    for i, (titulo, lineas, clave) in enumerate(buscadores):
        col = 1 + (i % 3) * 4
        if i and i % 3 == 0:
            r += 5
        _tarjeta(ws, r, col, 4, titulo, lineas, clave)
    r += 5

    # --- 3. estado del libro ---------------------------------------------
    banda(ws, r, "3 · ESTADO DEL LIBRO", DASH_NCOLS)
    r += 1
    for i, (titulo, valor, unidad, ambar) in enumerate(kpis):
        c1 = 1 + i * 3
        k = _mrg(ws, r, c1, c1 + 2, titulo)
        k.font, k.fill = KPI_TIT_F, BAND_FILL
        k.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        v = _mrg(ws, r + 1, c1, c1 + 2, valor)
        v.font = KPI_AMBAR_F if ambar else KPI_VAL_F
        v.fill = KPI_FILL
        v.alignment = Alignment(horizontal="center", vertical="center")
        u = _mrg(ws, r + 2, c1, c1 + 2, unidad)
        u.font, u.fill = UNIT_F, KPI_FILL
        u.alignment = Alignment(horizontal="center")
        for rr in range(r, r + 3):
            for cc in range(c1, c1 + 3):
                ws.cell(rr, cc).border = CARD_BORDER
    ws.row_dimensions[r].height = 26
    ws.row_dimensions[r + 1].height = 30
    ws.row_dimensions[r + 2].height = 14
    r += 3
    n = _mrg(ws, r, 1, DASH_NCOLS,
             f"Compilado {fecha} · fuente unica: resources/ · valores cargados tal como "
             f"estan impresos en el codigo; SI y US son extracciones independientes, "
             f"nunca conversiones.")
    n.font = SRC_F
    n.alignment = Alignment(vertical="center", indent=1)
    r += 2

    # --- pie --------------------------------------------------------------
    p = _mrg(ws, r, 1, DASH_NCOLS,
             "Herramienta de ingenieria de referencia. Verificar entradas y resultados, y "
             "leer las notas del material, antes de emitir para construccion.")
    p.font, p.fill = PIE_F, PIE_FILL
    p.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[r].height = 18
    return ws


def link_volver(wb):
    """Escribe el enlace de retorno al Dashboard en las 9 hojas navegables.

    La celda se desbloquea explicitamente: Parche_PCC2_Art212 e Instrucciones
    se entregan protegidas, y aunque Excel permite seguir un hipervinculo en
    celda bloqueada, dejarla desbloqueada evita depender de ese detalle.
    """
    for nombre in NAVEGABLES:
        if nombre not in wb.sheetnames:
            continue
        ws = wb[nombre]
        fila, c1, c2 = ANCLA_VOLVER[nombre]
        b = _boton(ws, fila, c1, c2, "◂ VOLVER AL DASHBOARD", CLAVE_VOLVER)
        b.protection = Protection(locked=False)
        _celda_clave(ws, fila, c1).protection = Protection(locked=False)
        _ocultar_columnas_clave(ws, (c1,))


def aplicar_visibilidad(wb):
    """Graba en el archivo el estado de visibilidad de cada hoja.

    No depende de la macro: si el usuario bloquea las macros, no ve ninguna
    base de datos, solo el Dashboard con el aviso en rojo. La macro reaplica
    exactamente esta misma tabla al abrir.
    """
    estados = {}
    for ws in wb.worksheets:
        if ws.title == DASH:
            ws.sheet_state = "visible"
        elif ws.title in NAVEGABLES:
            ws.sheet_state = "hidden"
        else:
            ws.sheet_state = "veryHidden"
        estados[ws.title] = ws.sheet_state
    return estados


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--resources", required=True)
    ap.add_argument("--in", dest="inp", required=True)
    ap.add_argument("--out", dest="out", required=True)
    ap.add_argument("--report", default=None)
    ap.add_argument("--decisiones", default=None,
                    help="JSON de decisiones validadas del ingeniero. Por omision "
                         "decisiones_map_grupo.json junto al reporte; si no existe "
                         "se ignora y todo queda como lo deja el codigo.")
    ap.add_argument("--revision", default=None,
                    help="Hoja de revision de MAP_Grupo (Markdown). Por omision se "
                         "escribe junto al libro como <salida>_Revision_MAP_Grupo.md")
    a = ap.parse_args(argv)

    res = Resources(a.resources)
    # keep_vba conserva vbaProject.bin del maestro sembrado por make_vba_seed.py.
    # Es lo que permite que el entregable .xlsm lleve la capa de navegacion sin
    # tocar el ZIP a mano.
    wb = openpyxl.load_workbook(a.inp, keep_vba=True)
    normalizar_textos_como_formula(wb)

    b313, b313c = build_b313(res, wb, "SI"), build_b313(res, wb, "US")
    iid, iidc = build_iid(res, wb, "SI", "1A"), build_iid(res, wb, "US", "1A")
    iidb, iidbc = build_iid(res, wb, "SI", "B"), build_iid(res, wb, "US", "B")
    su, suc = build_prop(res, wb, "SI", "U"), build_prop(res, wb, "US", "U")
    sy, syc = build_prop(res, wb, "SI", "Y1"), build_prop(res, wb, "US", "Y1")
    for si, us, lbl in ((b313, b313c, "B31.3 A-1/A-4"), (iid, iidc, "II-D 1A"),
                        (iidb, iidbc, "II-D 1B/3"), (su, suc, "Tabla U"),
                        (sy, syc, "Tabla Y-1")):
        verificar_paridad(si, us, lbl)

    e_si, e_us = build_modulo(res, wb, "SI"), build_modulo(res, wb, "US")
    build_te(res, wb, "SI"); build_te(res, wb, "US")
    prd, prdc = build_prd(res, wb, "SI"), build_prd(res, wb, "US")
    cap = build_appendix_c(res, wb, "SI")
    build_appendix_c(res, wb, "US")
    nm = build_nometalicos(res, wb)
    fac = build_map_factores(res, wb)
    ruta_dec = (Path(a.decisiones) if a.decisiones else
                Path(__file__).resolve().parent.parent / "decisiones_map_grupo.json")
    todas = [b313, b313c, iid, iidc, iidb, iidbc]
    map_last, map_stats, map_pend = build_map_grupo(
        res, wb, [iid, iidb], todas, ruta_dec, "SI")
    _, mapc_stats, mapc_pend = build_map_grupo(
        res, wb, [iidc, iidbc], todas, ruta_dec, "US")
    build_notas(res, wb)

    # listas simples de los buscadores por grupo
    wse, wsprd, wsnm = wb["DB_E"], wb["DB_PRD"], wb["DB_NoMetalicos"]
    e_pairs = uniques(wse, e_si["last_row"], 3, 4)
    c1_pairs = uniques(wb[cap["C-1"]["sheet"]], cap["C-1"]["last_row"], 3, 4)
    c3_pairs = uniques(wb[cap["C-3"]["sheet"]], cap["C-3"]["last_row"], 3, 4)
    prd_pairs = uniques(wsprd, prd["last_row"], 3, 4)
    nm_pairs = uniques(wsnm, nm["last_row"], 2, 4)
    simples = [
        ("E_TABLA", sorted({k for k, _ in e_pairs})),
        ("E_K", [k for k, _ in e_pairs]), ("E_V", [v for _, v in e_pairs]),
        ("C1_TABLA", sorted({k for k, _ in c1_pairs})),
        ("C1_K", [k for k, _ in c1_pairs]), ("C1_V", [v for _, v in c1_pairs]),
        ("C3_TABLA", sorted({k for k, _ in c3_pairs})),
        ("C3_K", [k for k, _ in c3_pairs]), ("C3_V", [v for _, v in c3_pairs]),
        ("PRD_TABLA", sorted({k for k, _ in prd_pairs if k})),
        ("PRD_K", [k for k, _ in prd_pairs]), ("PRD_V", [v for _, v in prd_pairs]),
        ("NM_TABLA", sorted({k for k, _ in nm_pairs})),
        ("NM_MATK", [k for k, _ in nm_pairs]), ("NM_MATV", [v for _, v in nm_pairs]),
    ]
    rangos = build_listas(wb, [("B313", b313), ("IID1A", iid), ("IIDB", iidb),
                               ("SU", su), ("SY", sy)], simples)

    curvas = wb.create_sheet("_Curvas")
    curvas["A1"] = ("Datos auxiliares de las graficas. Hoja oculta: no editar. "
                    "Los valores proceden de las hojas DB por INDEX/MATCH.")
    curvas.sheet_state = "hidden"

    busc = [
        ("Buscar_B31_3", "BUSCADOR — ASME B31.3-2024, Apendice A (Tablas A-1 y A-4)",
         "B313", b313, b313c, "MPa", "ksi", "S admisible"),
        ("Buscar_BPVC_IID", "BUSCADOR — ASME BPVC II-D 2025, Tabla 1A (ferrosos)",
         "IID1A", iid, iidc, "MPa", "ksi", "S admisible"),
        ("Buscar_BPVC_IID_B", "BUSCADOR — ASME BPVC II-D 2025, Tablas 1B (no ferrosos) y 3",
         "IIDB", iidb, iidbc, "MPa", "ksi", "S admisible"),
        ("Buscar_Su", "BUSCADOR — Resistencia a la traccion Su (ASME BPVC II-D, Tabla U)",
         "SU", su, suc, "MPa", "ksi", "Su"),
        ("Buscar_Sy", "BUSCADOR — Limite de fluencia Sy (ASME BPVC II-D, Tabla Y-1)",
         "SY", sy, syc, "MPa", "ksi", "Sy"),
    ]
    for i, (nme, ttl, pref, m, u, us_, uu, vl) in enumerate(busc):
        ctx = build_buscador(wb, curvas, nme, ttl, pref, rangos[pref], m, u,
                             us_, uu, vl)
        finish_buscador(ctx, wb, curvas, i)

    e_tab = sorted({k for k, _ in e_pairs})
    c1_tab = sorted({k for k, _ in c1_pairs})
    c3_tab = sorted({k for k, _ in c3_pairs})
    prd_tab = sorted({k for k, _ in prd_pairs if k})
    from collections import Counter as _Cnt
    mx_e = max(_Cnt(k for k, _ in e_pairs).values())
    mx_c1 = max(_Cnt(k for k, _ in c1_pairs).values())
    mx_c3 = max(_Cnt(k for k, _ in c3_pairs).values())
    mx_prd = max(_Cnt(k for k, _ in prd_pairs).values())
    mx_nm = max(_Cnt(k for k, _ in nm_pairs).values())
    build_buscador_grupo(wb, curvas, [
        dict(titulo="MODULO DE ELASTICIDAD E — ASME BPVC II-D, Tablas TM-1 a TM-5",
             lst_tabla="E_TABLA", nm_key="E_K", nm_val="E_V", default_tabla=e_tab[0],
             info=e_si, n_ident=e_si["n_ident"], temps=e_si["temps"],
             valor_lbl="Modulo E (valor tabulado)", unidad="x10^3 MPa", max_grupo=mx_e,
             nota="El valor real de E = valor tabulado x 10^3 MPa (factor del titulo de la tabla)."),
        dict(titulo="DILATACION TERMICA — ASME B31.3, Apendice C, Tabla C-1",
             lst_tabla="C1_TABLA", nm_key="C1_K", nm_val="C1_V", default_tabla=c1_tab[0],
             info=cap["C-1"], n_ident=cap["C-1"]["n_ident"], temps=cap["C-1"]["temps"],
             valor_lbl="Coef. de dilatacion", unidad="10^-6 mm/mm·°C", max_grupo=mx_c1,
             nota="Fila A = coeficiente medio; fila B = expansion lineal acumulada."),
        dict(titulo="MODULO DE ELASTICIDAD — ASME B31.3, Apendice C, Tabla C-3",
             lst_tabla="C3_TABLA", nm_key="C3_K", nm_val="C3_V", default_tabla=c3_tab[0],
             info=cap["C-3"], n_ident=cap["C-3"]["n_ident"], temps=cap["C-3"]["temps"],
             valor_lbl="Modulo E (valor tabulado)", unidad="x10^3 MPa", max_grupo=mx_c3,
             nota="El valor real de E = valor tabulado x 10^3 MPa."),
        dict(titulo="POISSON Y DENSIDAD — ASME BPVC II-D, Tabla PRD",
             lst_tabla="PRD_TABLA", nm_key="PRD_K", nm_val="PRD_V",
             default_tabla=prd_tab[0], info=prd, temps=None,
             valor_lbl="", unidad="", max_grupo=mx_prd,
             campos=[("Coeficiente de Poisson", 5, "adimensional"),
                     ("Densidad", 6, "kg/m3")]),
    ], rangos)
    build_buscador_nm(wb, nm, rangos, mx_nm)

    integrate_motor(wb, b313, iid, iidb, fac, rangos)
    deprecate_datos_ref(wb)

    counts = {
        "A-1 + A-4 -> DB_B31_3":
            f"{len(res.rows(f'{APX}/appendix_a/table_a_1.json')) + len(res.rows(f'{APX}/appendix_a/table_a_4.json'))}"
            f" -> {b313['last_row'] - R_DATA + 1}",
        "1A -> DB_BPVC_IID": f"{len(res.rows('bpvc_ii_d_metric_2025/table_1a.json'))}"
                             f" -> {iid['last_row'] - R_DATA + 1}",
        "1B + 3 -> DB_BPVC_IID_B":
            f"{len(res.rows('bpvc_ii_d_metric_2025/table_1b.json')) + len(res.rows('bpvc_ii_d_metric_2025/table_3.json'))}"
            f" -> {iidb['last_row'] - R_DATA + 1}",
        "U -> DB_Su": f"{len(res.rows('bpvc_ii_d_metric_2025/table_u.json'))}"
                      f" -> {su['last_row'] - R_DATA + 1}",
        "Y-1 -> DB_Sy": f"{len(res.rows('bpvc_ii_d_metric_2025/table_y_1.json'))}"
                        f" -> {sy['last_row'] - R_DATA + 1}",
        "MAP_Grupo": " · ".join(f"{k}={v}" for k, v in map_stats.items()),
        "MAP_GrupoC": " · ".join(f"{k}={v}" for k, v in mapc_stats.items()),
    }
    build_meta(wb, counts)
    rewrite_instrucciones(
        wb, "Rev. 3 — Dashboard unico de navegacion (libro con macros, .xlsm): las bases de "
            "datos quedan ocultas y se abre un motor a la vez. Consulta por cascada de "
            "listas desplegables, resultados sobre la ficha, curva del material como "
            "grafica y libro sin funciones de matriz dinamica.")

    # Los conteos del Dashboard salen de las mismas variables que alimentan
    # `counts`, nunca escritos a mano. El indicador cuenta lo que el codigo NO
    # resuelve: la regla textual que debe aplicar el ingeniero y los materiales
    # para los que II-D no publica el dato. Las filas AUTO no entran: citan la
    # nota que las sostiene y son auditables 1:1.
    n_sin_resolver = (map_stats.get(E_TEXTUAL, 0) + map_stats.get(E_COMP_AJENA, 0)
                      + map_stats.get(E_SIN, 0))
    build_dashboard(wb, [
        ("MATERIALES B31.3 · A-1 y A-4", b313["last_row"] - R_DATA + 1, "registros", False),
        ("MATERIALES II-D · TABLA 1A", iid["last_row"] - R_DATA + 1, "registros", False),
        ("MATERIALES II-D · TABLA U", su["last_row"] - R_DATA + 1, "registros", False),
        ("MAP_Grupo SIN GRUPO NORMATIVO", n_sin_resolver,
         "filas · no usar E ni dilatacion", True),
    ], datetime.date.today().isoformat())

    link_volver(wb)

    order = [DASH,
             "Instrucciones", "Parche_PCC2_Art212", "Buscar_B31_3", "Buscar_BPVC_IID",
             "Buscar_BPVC_IID_B", "Buscar_Su", "Buscar_Sy", "Buscar_Propiedades",
             "Buscar_NoMetalicos", "Datos_Ref", "DB_B31_3", "DB_B31_3C", "DB_BPVC_IID",
             "DB_BPVC_IIDC", "DB_BPVC_IID_B", "DB_BPVC_IID_BC", "DB_Su", "DB_SuC",
             "DB_Sy", "DB_SyC", "DB_E", "DB_EC", "DB_TE", "DB_TEC", "DB_PRD", "DB_PRDC",
             "DB_C_dilatacion", "DB_C_dilatacionC", "DB_C_modulo", "DB_C_moduloC",
             "DB_NoMetalicos", "MAP_Factores", "MAP_Grupo", "MAP_GrupoC",
             "Notas_Codigo", "DB_Listas",
             "_meta", "_Curvas"]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + \
                 [s for s in wb._sheets if s.title not in order]
    # Ultimo paso antes de guardar: el estado de visibilidad debe reflejar el
    # libro completo, incluidas las hojas que hubiese traido el maestro.
    estados = aplicar_visibilidad(wb)
    wb.save(a.out)

    # Por omision acompana al reporte de verificacion en la carpeta del proyecto,
    # no al libro: es documentacion de respaldo, no un entregable suelto.
    ruta_rev = Path(a.revision) if a.revision else \
        Path(__file__).resolve().parent.parent / "Revision_MAP_Grupo.md"
    n_decisiones = escribir_revision_map_grupo(
        map_pend + mapc_pend, map_stats, ruta_rev)
    ISSUES.append(f"MAP_Grupo: {n_decisiones} decisiones distintas pendientes de "
                  f"validacion del ingeniero -> {ruta_rev.name}")

    report = {"salida": a.out, "hojas": wb.sheetnames, "conteos": counts,
              "visibilidad": estados, "limitaciones": ISSUES, "meta": META,
              "revision_map_grupo": str(ruta_rev)}
    if a.report:
        Path(a.report).write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                  encoding="utf-8")
    from collections import Counter as _C
    print(json.dumps({"hojas": len(wb.sheetnames),
                      "visibilidad": dict(_C(estados.values())),
                      "conteos": counts, "limitaciones": ISSUES},
                     ensure_ascii=False, indent=2))
    return report


if __name__ == "__main__":
    main()
