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


FERROUS_RULES = [
    (r"carbon steel", "Carbon steels (TM-1 / TE-1) — definir si C <=0,30 % o > 0,30 %"),
    (r"c-?\s*mn|carbon-?manganese", "Carbon steels (TM-1 / TE-1) — definir por %C"),
    (r"\b9cr\b|9cr-1mo|gr(ade)?\s*9[12]", "9Cr-1Mo (incl. Gr. 9, 91, 911, 92) — TM-1 / TE-1"),
    (r"\b5cr\b|5cr-1mo", "5Cr-1Mo y 29Cr-7Ni-2Mo-N — TE-1 / Material Group A (TM-1)"),
    (r"2\s*1/4cr|2\.25cr|1\s*1/4cr|1cr|1/2cr|c-?1/2mo|mo\b",
     "Aceros de baja aleacion — TM-1 Group A / TE-1 Group 1-2"),
    (r"18cr-8ni|16cr-12ni|18cr-10ni|austenit|type 3\d\d",
     "Inoxidables austeniticos — TM-1 Group C / TE-1 Group 3"),
    (r"\b1[23]cr\b|\b17cr\b|\b15cr\b|ferritic|martensit", "12Cr / 15Cr / 17Cr — TE-1"),
    (r"9ni|8ni", "Aceros 8Ni y 9Ni — TE-1"),
    (r"duct(ile)? (cast )?iron", "Fundicion ductil — TM-1 / TE-1"),
]


def build_map_grupo(res, wb, iid_infos):
    ed = "bpvc_ii_d_metric_2025"
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
    prd_groups = {txt(g(r, "material")).upper(): txt(g(r, "material_group"))
                  for r in res.load(f"{ed}/table_prd.json")["rows"]}

    ws = new_sheet(wb, "MAP_Grupo",
                   "MAPEO material -> grupo de propiedades (TM / TE / PRD) · ASME BPVC II-D 2025",
                   "Artefacto derivado: TM/TE/PRD se indexan por FAMILIA de material, no por "
                   "spec. AUTO = coincidencia literal de UNS con una fila de TM-2..TM-5 "
                   "(auditable 1:1). PROPUESTA = familia sugerida a partir de la composicion "
                   "nominal impresa: REQUIERE VALIDACION antes de usar E o dilatacion "
                   "(regla 5.3 de knowledge/claude.md). SIN MAPEO = no usar E ni alfa.")
    n = write_headers(ws, ["material_id", "Spec. No.", "UNS / Alloy", "Composicion nominal",
                           "Grupo TM propuesto", "Tabla TM", "Grupo PRD propuesto",
                           "Estado", "Busqueda"])
    recs, stats = [], {"AUTO (UNS exacto)": 0, "PROPUESTA (VALIDAR)": 0, "SIN MAPEO": 0}
    for info in iid_infos:
        src = wb[info["sheet"]]
        for r in range(R_DATA, info["last_row"] + 1):
            mid = src.cell(r, C["material_id"]).value
            spec = src.cell(r, C["Spec. No."]).value
            uns = src.cell(r, C["UNS / Alloy"]).value
            comp = src.cell(r, C["Composicion nominal"]).value
            tm_tab = tm_grp = None
            estado = "SIN MAPEO"
            key = txt(uns).upper()
            hit = None
            for tok in re.findall(r"[A-Z]\d{5}", key):
                if tok in tm_index:
                    hit = tm_index[tok]
                    break
            if hit is None and key in tm_index:
                hit = tm_index[key]
            if hit:
                tm_tab, tm_grp, estado = hit[0], hit[1], "AUTO (UNS exacto)"
            else:
                low = txt(comp).lower()
                for pat, grp in FERROUS_RULES:
                    if re.search(pat, low):
                        tm_grp, tm_tab, estado = grp, "TM-1 / TE-1", "PROPUESTA (VALIDAR)"
                        break
            prd = next((v for k, v in prd_groups.items()
                        if k and k in txt(comp).upper()), None)
            stats[estado] += 1
            recs.append(([mid, spec, uns, comp, tm_grp, tm_tab, prd, estado,
                          search_key(mid, spec, uns, comp)], {}))
    last = write_rows(ws, recs, n)
    autosize(ws, {"A": 46, "B": 12, "C": 14, "D": 24, "E": 52, "F": 12, "G": 22,
                  "H": 20, "I": 28})
    ws.column_dimensions["I"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:H{last}"
    ISSUES.append("MAP_Grupo: " + " · ".join(f"{k}={v}" for k, v in stats.items()))
    record_meta("MAP_Grupo", "derivado (TM-2..5, PRD)",
                f"{ed}/table_tm_*.json ; table_prd.json", "2025", "-",
                last - R_DATA + 1, "AUTO = UNS exacto; PROPUESTA requiere validacion.")
    return last, stats


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
    """Propiedades indexadas por FAMILIA de material (no por especificacion):
    modulo E, dilatacion termica, Poisson y densidad. Mismo formato de tarjeta."""
    ws = new_sheet(wb, "Buscar_Propiedades",
                   "BUSCADOR DE PROPIEDADES POR FAMILIA DE MATERIAL — modulo E (TM-1..5), "
                   "dilatacion y modulo del Apendice C del B31.3, Poisson y densidad (PRD)",
                   "Estas tablas del codigo se indexan por FAMILIA de material, no por "
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
     "Poisson y densidad. Estas tres se indexan por FAMILIA de material: ver MAP_Grupo.\n"
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
     "MAP_Grupo, solo las filas 'AUTO (UNS exacto)' son mapeo 1:1 auditable; las 'PROPUESTA "
     "(VALIDAR)' requieren que el ingeniero confirme la familia antes de usar E o dilatacion."),
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


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--resources", required=True)
    ap.add_argument("--in", dest="inp", required=True)
    ap.add_argument("--out", dest="out", required=True)
    ap.add_argument("--report", default=None)
    a = ap.parse_args(argv)

    res = Resources(a.resources)
    wb = openpyxl.load_workbook(a.inp)

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
    map_last, map_stats = build_map_grupo(res, wb, [iid, iidb])
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
    }
    build_meta(wb, counts)
    rewrite_instrucciones(
        wb, "Rev. 2 — Consulta por cascada de listas desplegables (composicion -> forma -> "
            "especificacion -> tipo/grado), resultados sobre la ficha, curva del material "
            "como grafica y libro sin funciones de matriz dinamica.")

    order = ["Instrucciones", "Parche_PCC2_Art212", "Buscar_B31_3", "Buscar_BPVC_IID",
             "Buscar_BPVC_IID_B", "Buscar_Su", "Buscar_Sy", "Buscar_Propiedades",
             "Buscar_NoMetalicos", "Datos_Ref", "DB_B31_3", "DB_B31_3C", "DB_BPVC_IID",
             "DB_BPVC_IIDC", "DB_BPVC_IID_B", "DB_BPVC_IID_BC", "DB_Su", "DB_SuC",
             "DB_Sy", "DB_SyC", "DB_E", "DB_EC", "DB_TE", "DB_TEC", "DB_PRD", "DB_PRDC",
             "DB_C_dilatacion", "DB_C_dilatacionC", "DB_C_modulo", "DB_C_moduloC",
             "DB_NoMetalicos", "MAP_Factores", "MAP_Grupo", "Notas_Codigo", "DB_Listas",
             "_meta", "_Curvas"]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + \
                 [s for s in wb._sheets if s.title not in order]
    wb.save(a.out)
    report = {"salida": a.out, "hojas": wb.sheetnames, "conteos": counts,
              "limitaciones": ISSUES, "meta": META}
    if a.report:
        Path(a.report).write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                  encoding="utf-8")
    print(json.dumps({"hojas": len(wb.sheetnames), "conteos": counts,
                      "limitaciones": ISSUES}, ensure_ascii=False, indent=2))
    return report


if __name__ == "__main__":
    main()
