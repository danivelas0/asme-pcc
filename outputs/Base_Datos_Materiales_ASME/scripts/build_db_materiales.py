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
from openpyxl.comments import Comment

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


AUTOR_NOTA = "ASME PCC — Motor de calculo"


def _nota(cell, texto, ancho=260, alto=90):
    """Adjunta un comentario de Excel (el que se ve al pasar el mouse, no texto
    de celda) explicando que calcula la celda o que hay que ingresar en ella."""
    c = Comment(texto, AUTOR_NOTA)
    c.width, c.height = ancho, alto
    cell.comment = c


def dv_list(ws, cell, formula, comentario=None):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(ws[cell])
    if comentario:
        _nota(ws[cell], comentario)


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
    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025" if si else "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
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
    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025" if si else "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
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
    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025" if si else "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
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
    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025" if si else "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
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
    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025" if si else "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
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


# TE-1..5 publican, por cada grupo/columna, TRES coeficientes (A = instantaneo,
# B = medio de 20 C a T, C = expansion acumulada) en columnas consecutivas del
# mismo bloque. `build_te` los conserva tal como estan impresos (temperatura en
# filas, grupo en columnas) para auditar; para un buscador por MATERIAL hace
# falta el pivote inverso -grupo en filas, temperatura en columnas-, que es
# exactamente lo que `build_modulo` ya hace para TM-1..5.
#
# Solo se pivota el Coeficiente B (medio): es el que se usa en calculo de
# dilatacion/flexibilidad (la misma magnitud que "alfa" en el Apendice C del
# B31.3, Tabla C-1). Los Coeficientes A y C siguen impresos tal cual en
# DB_TE/DB_TEC -no se pierden, solo no alimentan este buscador-: decision de
# alcance declarada ("Un dato, un motor"), no un hueco silencioso.
#
# El ROTULO de cada fila tiene que ser IDENTICO al que build_map_grupo escribe
# en `Grupo dilatacion (TE)` de MAP_Grupo/MAP_GrupoC, para que la seleccion en
# Buscar_Prop_IID case letra a letra con lo que ahi se lee: "Group N" cuando el
# titulo de columna lo declara (los Grupos 1..4 de las Notas), o el titulo
# integro de la columna B menos el sufijo " B" en cualquier otro caso -con los
# artefactos de extraccion que trae impreso incluidos ("(In- cluding...")-,
# que es exactamente `etiqueta` en `columnas_nombradas_te1`.
_TE_COL_B = re.compile(r"(?:^|\s)B$")
_TE_GROUP_NUM = re.compile(r"\(Group\s+(\d+)\)")


def etiqueta_columna_b_te1(titulo):
    """Rotulo de grupo de una columna Coeficiente B de TE-1..5, identico al
    que usa `build_map_grupo` para `grupo_te`. None si el titulo no es una
    columna B identificable (p. ej. una "B" suelta, truncada en la extraccion
    sin cuerpo que la acompañe: no hay con que rotular la fila sin inventar)."""
    t = txt(titulo)
    if not t or not _TE_COL_B.search(t):
        return None
    m = _TE_GROUP_NUM.search(t)
    if m:
        return f"Group {m.group(1)}"
    etq = _TE_COL_B.sub("", t).strip()
    return etq or None


def _clave_fila_te1(titulo):
    """`columns` trae el rotulo IMPRESO de cada columna; las claves de `rows`
    son ese mismo rotulo reducido a snake_case (sin el pie de nota entre
    corchetes). Hace falta este puente porque las filas no se indexan por el
    texto impreso, y las filas son ademas dispersas -una columna sin dato en
    esa temperatura no trae su clave-, asi que no vale alinear por posicion."""
    s = re.sub(r"\[[^\]]*\]", "", titulo)
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
    return s


def build_dilatacion_grupo(res, wb, system):
    si = system == "SI"
    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025" if si else "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
    temps, recs, sin_rotulo = set(), [], []
    for i in range(1, 6):
        d = res.load(f"{ed}/table_te_{i}.json")
        tid = d.get("table_id")
        cols = d.get("columns", [])
        if not cols:
            continue
        temp_key = _clave_fila_te1(cols[0])
        clasif = {}
        for col in cols[1:]:
            etq = etiqueta_columna_b_te1(col)
            if etq:
                clasif[_clave_fila_te1(col)] = etq
            elif txt(col) and _TE_COL_B.search(txt(col)):
                sin_rotulo.append(f"{tid}: columna Coeficiente B «{txt(col)}» sin "
                                  "rotulo de grupo identificable (truncada en la "
                                  "extraccion); no se incorpora a este buscador.")
        grupos: dict[str, dict] = {}
        for row in d["rows"]:
            tnum = temp_to_number(row.get(temp_key))
            if tnum is None:
                continue
            for col_key, etq in clasif.items():
                v = num(row.get(col_key))
                if v is not None:
                    grupos.setdefault(etq, {})[tnum] = v
        for etq, vals in grupos.items():
            temps |= set(vals)
            recs.append(dict(tid=tid, mat=etq, vals=vals))
    if sin_rotulo:
        ISSUES.append(f"DB_TE_G ({system}): " + " | ".join(sin_rotulo))
    temps = sorted(temps)
    recs.sort(key=lambda d: (d["tid"], txt(d["mat"]).upper()))
    keys, _ = make_unique([f"{d['tid']} | {txt(d['mat'])}" for d in recs],
                          list(range(1, len(recs) + 1)))
    records = [([k, d["tid"], d["tid"], d["mat"], "Coeficiente B (medio)",
                 search_key(d["mat"]), len(d["vals"])], d["vals"])
               for k, d in zip(keys, recs)]
    name = "DB_TE_G" if si else "DB_TE_GC"
    unidad = "10^-6 mm/mm/C" if si else "10^-6 in/in/F"
    ws = new_sheet(wb, name,
                   "BASE DE PROPIEDADES — Dilatacion termica, Coeficiente B (medio, de 20 C "
                   f"a T) · ASME BPVC II-D 2025, Tablas TE-1 a TE-5 ({unidad})",
                   f"Fuente: resources/{ed}/table_te_1..5.json · Edicion 2025 · "
                   "Pivotada por GRUPO/columna (Coeficiente B unicamente); la banda tal como "
                   "esta impresa, con los Coeficientes A y C, esta en "
                   f"{'DB_TE' if si else 'DB_TEC'}. Indexada por GRUPO: ver "
                   f"{'MAP_Grupo' if si else 'MAP_GrupoC'}.")
    n = write_headers(ws, GRP_COLS, temps)
    last = write_rows(ws, records, n, temps)
    t0, v0, npack = append_packed(ws, records, n, temps, 7)
    autosize(ws, {"A": 60, "B": 8, "C": 8, "D": 60, "E": 22, "F": 30})
    ws.column_dimensions["C"].hidden = True
    ws.column_dimensions["F"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:{get_column_letter(n + len(temps))}{last}"
    record_meta(name, "TE-1..TE-5 (Coef. B)", f"{ed}/table_te_1..5.json", "2025", system,
                last - R_DATA + 1, "Dilatacion termica media (Coeficiente B), por grupo.")
    return dict(sheet=name, n_ident=n, temps=temps, last_row=last,
                pack_t0=t0, pack_v0=v0, npack=npack, npts_col=7)


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


# ---------------------------------------------------------------------------
# Apendice C del B31.3 — base unificada de propiedades fisicas
# ---------------------------------------------------------------------------
# Las cuatro tablas del Apendice C (C-1/C-1C dilatacion de metales, C-2
# dilatacion de no metalicos, C-3/C-3C modulo de metales y C-4 modulo de no
# metalicos) viven en UNA sola base por edicion. El nivel 0 de la cascada es la
# PROPIEDAD: es la pregunta que hace el ingeniero ("quiero el modulo E"), no la
# tabla en la que el codigo la publica.
#
# Contrato de columnas: las OCHO primeras posiciones son identicas a
# STRESS_COLS. Eso es lo que permite que build_listas y las secciones §2
# (unicidad) y §4 (contiguidad) de verificar.py funcionen sobre esta base sin
# tocar una linea.
#
# El plan PLAN-BUSC-APXC-001 enumeraba 25 columnas; aqui son 27 porque la
# definicion impresa del coeficiente y el texto impreso del factor de escala se
# guardan en columna propia en vez de componerse dentro de una formula: la
# ficha del motor los cita verbatim y asi siguen siendo auditables desde la
# hoja.
APXC_COLS = [
    "material_id", "Tabla", "k0", "k1", "k2", "k3", "k4", "clave_bi",
    "Propiedad", "Grupo impreso", "Subgrupo impreso", "Material",
    "Variante / coeficiente", "Definicion impresa", "Tipo de dato",
    "Unidad impresa", "Factor de escala", "Factor impreso",
    "Valor unico", "Valor unico (texto)", "Rango de validez",
    "T primera tabulada", "T ultima tabulada", "Notas", "Observacion",
    "Linea", "n_pts"]
CC = {n: i + 1 for i, n in enumerate(APXC_COLS)}
N_IDENT_C = len(APXC_COLS)
CLC = {n: get_column_letter(i) for n, i in CC.items()}
assert APXC_COLS[:8] == STRESS_COLS[:8], \
    "el contrato de las 8 primeras columnas es lo que reutiliza build_listas"

PROP_C1 = "DILATACION TERMICA - METALES"
PROP_C2 = "DILATACION TERMICA - NO METALICOS"
PROP_C3 = "MODULO DE ELASTICIDAD - METALES"
PROP_C4 = "MODULO ELASTICIDAD CORTO PLAZO - NO METALICOS"
# El orden alfabetico de estos cuatro rotulos coincide con el orden de las
# tablas en el codigo (C-1, C-2, C-3, C-4); la base se ordena por k0 y por eso
# no hace falta un criterio aparte.
SIN_GRUPO = "(sin grupo impreso)"
SIN_SUBGRUPO = "(sin subgrupo impreso)"
VAR_UNICA = "(unico)"
VARIANTE_COEF = {"A": "A - coeficiente medio", "B": "B - expansion total"}
TIPO_CURVA, TIPO_PUNTO = "CURVA", "PUNTO"

# Rotulos de estado del motor del Apendice C. Viven aqui, y no dentro de la
# formula, porque verificar.py §6 recalcula esa misma logica en Excel y compara
# el texto: si el motor y el verificador escribiesen cada uno su literal, una
# divergencia de redaccion pasaria desapercibida.
EST_SIN_SEL = "SIN SELECCION"
EST_PUNTO = "VALOR UNICO — no depende de T; vea el rango de validez"
EST_SIN_TAB = "SIN VALOR TABULADO"
EST_BAJO = "FUERA DE RANGO — T por debajo del primer punto tabulado"
EST_ALTO = "FUERA DE RANGO — T por encima del ultimo punto tabulado"
EST_OK = "EN RANGO"
VAL_BLOQUEADO = "BLOQUEADO"

_RE_NOTA_REF = re.compile(r"\[Note \((\d+[A-Za-z]?)\)\]")


def _notas_ref(*textos) -> str:
    """Referencias [Note (n)] que el codigo imprime en un rotulo, como '(1), (2)'."""
    vistas = []
    for t in textos:
        for n in _RE_NOTA_REF.findall(str(t or "")):
            if n not in vistas:
                vistas.append(n)
    return ", ".join(f"({n})" for n in vistas)


def _factor_escala(sf):
    """Multiplicador impreso -> valor fisico, LEIDO del JSON (nunca codificado).

    El codigo publica el factor en el encabezado de cada tabla: 'Multiply
    Tabulated Values by 10^3' (C-3), 'Divide Table Values by 10^6' (C-2). Se
    confundir 10^3 con 10^6 es un error de tres ordenes de magnitud, asi que el
    exponente se lee y se guarda junto con su texto impreso.
    """
    if not sf:
        return 1, ""
    op = str(sf.get("operacion") or "ninguna")
    exp = int(sf.get("exponente") or 0)
    if op == "multiplicar":
        f = 10 ** exp
    elif op == "dividir":
        f = 10 ** (-exp)
    else:
        f = 1
    return f, clean(sf.get("impreso")) or ""


def _seg(v, por_defecto=""):
    """Segmento de una clave de cascada: el separador es '|', asi que no puede
    aparecer dentro del valor. Ningun rotulo del Apendice C lo lleva; la
    sustitucion es defensiva, para que una reextraccion no rompa la cascada."""
    s = txt(v)
    return (s.replace("|", "/") if s else por_defecto)


def _excepciones_declaradas(d):
    """{nombre de fila -> motivo declarado por completar_apendice_c.py}."""
    out = {}
    for e in (d.get("extraction_amendments") or {}).get("excepciones_declaradas") or []:
        fila = txt(e.get("fila"))
        if fila:
            out[fila.upper()] = f"{clean(e.get('motivo'))} (folio {e.get('folio')})"
    return out


def _fila_apxc(**kw):
    """Fila normalizada de la base del Apendice C, con todos los campos."""
    d = dict(tabla="", prop="", grupo=SIN_GRUPO, subgrupo=SIN_SUBGRUPO, material="",
             variante=VAR_UNICA, definicion="", tipo=TIPO_CURVA, unidad="",
             factor=1, factor_txt="", vnum=None, vtxt=None, rango=None,
             notas="", obs=[], idx=0, prop_id="", vals={})
    d.update(kw)
    return d


def _carga_apendice_c(res, si):
    """Las 4 propiedades del Apendice C, en la edicion pedida. Solo lectura."""
    ap = f"{APX}/appendix_c"
    filas = []

    # --- C-1 / C-1C: dilatacion termica de metales, curva vs. T --------------
    fn = f"{ap}/table_c_1{'' if si else 'c'}.json"
    d = res.load(fn)
    defs = d.get("coefficient_defs") or {}
    fac, fac_txt = _factor_escala(d.get("scale_factor"))
    exc = _excepciones_declaradas(d)
    for i, row in enumerate(d["rows"], start=1):
        crudo = clean(g(row, "material_uns_no"))
        mat = _dedup_frase(crudo)
        # C-1C parte el encabezado en 'Coeffi- cient': el lector acepta las dos
        # claves. Un parser que asuma una sola falla en la mitad de los casos.
        coef = txt(g(row, "coefficient", "coeffi_cient"))
        cd = defs.get(coef) or {}
        obs = []
        if mat != crudo:
            obs.append("Nombre impreso duplicado por la extraccion; colapsado "
                       "(artefacto, el codigo lo imprime una vez)")
        if txt(mat).upper() in exc:
            obs.append(exc[txt(mat).upper()])
        vals = {temp_to_number(k): num(v)
                for k, v in (row.get("values") or {}).items()
                if temp_to_number(k) is not None and num(v) is not None}
        filas.append(_fila_apxc(
            tabla=d.get("table_id", "").replace("Table ", "") or "C-1",
            prop=PROP_C1, prop_id="C1", idx=i, material=mat,
            variante=VARIANTE_COEF.get(coef, coef or VAR_UNICA),
            definicion=clean(cd.get("impreso")) or "",
            tipo=TIPO_CURVA, unidad=clean(cd.get("unidad")) or "",
            factor=fac, factor_txt=fac_txt,
            notas=_notas_ref(mat, cd.get("referencia")),
            obs=obs, vals=vals))

    # --- C-2: dilatacion termica de no metalicos, valor unico ----------------
    fn = f"{ap}/table_c_2.json"
    d = res.load(fn)
    cols = d.get("columns") or []
    fac, fac_txt = _factor_escala(d.get("scale_factor"))
    exc = _excepciones_declaradas(d)
    # El codigo publica los DOS sistemas en la misma tabla: el conmutador cambia
    # de columna, no de hoja (regla 10). Nunca hay conversion.
    kv, kr = ("mm_mm_c", "range_c") if si else ("in_in_f", "range_f")
    unidad = clean(cols[3] if si else cols[1]) if len(cols) > 4 else ""
    u_temp = "°C" if si else "°F"
    from collections import Counter as _Cnt2
    rep = _Cnt2(txt(r.get("material_description")) for r in d["rows"])
    for i, row in enumerate(d["rows"], start=1):
        mat = clean(g(row, "material_description"))
        v = row.get(kv)
        rng = clean(row.get(kr))
        obs = []
        if txt(mat).upper() in exc:
            obs.append(exc[txt(mat).upper()])
        if isinstance(v, str):
            obs.append("El codigo publica un intervalo; se conserva como texto y "
                       "el motor no opera sobre el (regla 9)")
        # El codigo repite `Poly(perfluoroalkoxy alkane)` tres veces con rangos
        # de validez distintos: el rango es lo que separa esas filas, asi que es
        # la variante. En el resto no hay variante que elegir.
        var = (clean(rng) or f"fila impresa {i}") if rep[txt(mat)] > 1 else VAR_UNICA
        filas.append(_fila_apxc(
            tabla="C-2", prop=PROP_C2, prop_id="C2", idx=i, material=mat,
            grupo=clean(g(row, "material_group")) or SIN_GRUPO,
            subgrupo=clean(g(row, "material_subgroup")) or SIN_SUBGRUPO,
            variante=var, tipo=TIPO_PUNTO, unidad=unidad,
            factor=fac, factor_txt=fac_txt,
            vnum=v if isinstance(v, (int, float)) else None,
            vtxt=clean(v) if isinstance(v, str) else None,
            rango=f"{rng} {u_temp}" if rng else None,
            notas=_notas_ref(g(row, "material_group")), obs=obs))

    # --- C-3 / C-3C: modulo de elasticidad de metales, curva vs. T ----------
    fn = f"{ap}/table_c_3{'' if si else 'c'}.json"
    d = res.load(fn)
    fac, fac_txt = _factor_escala(d.get("scale_factor"))
    exc = _excepciones_declaradas(d)
    unidad = clean((d.get("scale_factor") or {}).get("unidad_resultante")) or ""
    for i, row in enumerate(d["rows"], start=1):
        crudo = clean(g(row, "material"))
        mat = _dedup_frase(crudo)
        obs = []
        if mat != crudo:
            obs.append("Nombre impreso duplicado por la extraccion; colapsado "
                       "(artefacto, el codigo lo imprime una vez)")
        if txt(mat).upper() in exc:
            obs.append(exc[txt(mat).upper()])
        vals = {temp_to_number(k): num(v)
                for k, v in (row.get("values") or {}).items()
                if temp_to_number(k) is not None and num(v) is not None}
        filas.append(_fila_apxc(
            tabla=d.get("table_id", "").replace("Table ", "") or "C-3",
            prop=PROP_C3, prop_id="C3", idx=i, material=mat,
            grupo=clean(g(row, "material_group")) or SIN_GRUPO,
            subgrupo=clean(g(row, "material_subgroup")) or SIN_SUBGRUPO,
            tipo=TIPO_CURVA, unidad=unidad, factor=fac, factor_txt=fac_txt,
            notas=_notas_ref(d.get("value_axis"), g(row, "material_group")),
            obs=obs, vals=vals))

    # --- C-4: modulo de elasticidad de no metalicos, valor unico ------------
    fn = f"{ap}/table_c_4.json"
    d = res.load(fn)
    cols = d.get("columns") or []
    fac, fac_txt = _factor_escala(d.get("scale_factor"))
    exc = _excepciones_declaradas(d)
    kv = "e_mpa_23_c" if si else "e_ksi_73_4_f"
    # 'E, MPa (23°C)' -> 'MPa (23°C)': se quita el rotulo de la magnitud, que ya
    # esta en el nombre de la propiedad; el resto se conserva impreso.
    hdr = clean(cols[2] if si else cols[1]) if len(cols) > 2 else ""
    unidad = re.sub(r"^E,\s*", "", hdr or "")
    ref = [t for t in (d.get("reference_temperature") or [])
           if str(t.get("unidad")) == ("C" if si else "F")]
    rango_ref = (f"{ref[0]['valor']} °{ref[0]['unidad']}" if ref else None)
    for i, row in enumerate(d["rows"], start=1):
        mat = clean(g(row, "material_description"))
        v = row.get(kv)
        obs = []
        if txt(mat).upper() in exc:
            obs.append(exc[txt(mat).upper()])
        if isinstance(v, str):
            obs.append("El codigo publica un intervalo; se conserva como texto y "
                       "el motor no opera sobre el (regla 9)")
        filas.append(_fila_apxc(
            tabla="C-4", prop=PROP_C4, prop_id="C4", idx=i, material=mat,
            grupo=clean(g(row, "material_group")) or SIN_GRUPO,
            subgrupo=clean(g(row, "material_subgroup")) or SIN_SUBGRUPO,
            tipo=TIPO_PUNTO, unidad=unidad, factor=fac, factor_txt=fac_txt,
            vnum=v if isinstance(v, (int, float)) else None,
            vtxt=clean(v) if isinstance(v, str) else None,
            rango=rango_ref,
            notas=_notas_ref(g(row, "material_group")), obs=obs))
    return filas


def build_apendice_c(res, wb, system):
    """DB_B31_C / DB_B31_CC — las 4 tablas del Apendice C en una sola base."""
    si = system == "SI"
    name = "DB_B31_C" if si else "DB_B31_CC"
    filas = _carga_apendice_c(res, si)

    # Orden: k0 -> k1 -> k2 -> material -> orden impreso. El ultimo criterio es
    # el numero de fila del codigo, no la variante: garantiza que las dos
    # ediciones queden en el MISMO orden aunque la variante se lea de una
    # columna distinta en cada una (el rango de validez de C-2 va en °C o en °F).
    filas.sort(key=lambda d: (d["prop"], _seg(d["grupo"], SIN_GRUPO).upper(),
                              _seg(d["subgrupo"], SIN_SUBGRUPO).upper(),
                              _seg(d["material"]).upper(), d["prop_id"], d["idx"]))

    temps = sorted({t for d in filas for t in d["vals"]})
    ids, _ = make_unique(
        [build_material_id([d["tabla"], d["material"], d["variante"]]) for d in filas],
        [d["idx"] for d in filas])
    records = []
    for mid, d in zip(ids, filas):
        gr = _seg(d["grupo"], SIN_GRUPO)
        sg = _seg(d["subgrupo"], SIN_SUBGRUPO)
        mt = _seg(d["material"])
        vr = _seg(d["variante"], VAR_UNICA)
        k0 = d["prop"]
        k1 = f"{k0}|{gr}"
        k2 = f"{k1}|{sg}"
        k3 = f"{k2}|{mt}"
        k4 = f"{k3}|{vr}"
        d["k"] = (k0, k1, k2, k3, k4)
        d["mid"], d["bi"] = mid, f"{d['prop_id']}#{d['idx']}"
        tprim = min(d["vals"]) if d["vals"] else None
        tult = max(d["vals"]) if d["vals"] else None
        records.append((
            [mid, d["tabla"], k0, k1, k2, k3, k4, d["bi"],
             d["prop"], d["grupo"], d["subgrupo"], d["material"],
             d["variante"], d["definicion"], d["tipo"], d["unidad"],
             d["factor"], d["factor_txt"], d["vnum"], d["vtxt"], d["rango"],
             tprim, tult, d["notas"], " · ".join(d["obs"]) or None,
             d["idx"], len(d["vals"])],
            d["vals"]))

    ws = new_sheet(
        wb, name,
        "PROPIEDADES FISICAS DE TUBERIA — ASME B31.3-2024, Apendice C · Tablas "
        + ("C-1, C-2, C-3 y C-4 (edicion metrica)" if si
           else "C-1C, C-2, C-3C y C-4 (edicion U.S. Customary)"),
        "Fuente: resources/" + APX + "/appendix_c/table_c_1"
        + ("" if si else "c") + ".json, table_c_2.json, table_c_3"
        + ("" if si else "c") + ".json y table_c_4.json · ASME B31.3-2024 · "
        "Valores tal como estan impresos: el factor de escala va en su columna y "
        "lo aplica el motor a la vista, nunca la carga. C-2 y C-4 publican los dos "
        "sistemas en la misma tabla y de ellas se lee la COLUMNA del sistema "
        "activo; C-1/C-1C y C-3/C-3C tienen una tabla por edicion. Nunca hay "
        "conversion.")
    n = write_headers(ws, APXC_COLS, temps)
    last = write_rows(ws, records, n, temps)
    autosize(ws, {CLC["material_id"]: 52, CLC["Tabla"]: 8, CLC["clave_bi"]: 10,
                  CLC["Propiedad"]: 34, CLC["Grupo impreso"]: 30,
                  CLC["Subgrupo impreso"]: 28, CLC["Material"]: 46,
                  CLC["Variante / coeficiente"]: 22, CLC["Definicion impresa"]: 44,
                  CLC["Tipo de dato"]: 12, CLC["Unidad impresa"]: 18,
                  CLC["Factor impreso"]: 44, CLC["Rango de validez"]: 16,
                  CLC["Observacion"]: 50})
    for cn in ("k0", "k1", "k2", "k3", "k4", "clave_bi", "Linea", "n_pts"):
        ws.column_dimensions[CLC[cn]].hidden = True
    t0, v0, npack = append_packed(ws, records, n, temps, CC["n_pts"])
    ws.auto_filter.ref = f"A{R_HDR}:{get_column_letter(n + len(temps))}{last}"

    ap = f"{APX}/appendix_c"
    for tag, fn, tbl in (("C-1", f"{ap}/table_c_1{'' if si else 'c'}.json", PROP_C1),
                         ("C-2", f"{ap}/table_c_2.json", PROP_C2),
                         ("C-3", f"{ap}/table_c_3{'' if si else 'c'}.json", PROP_C3),
                         ("C-4", f"{ap}/table_c_4.json", PROP_C4)):
        nf = sum(1 for d in filas if d["prop"] == tbl)
        record_meta(name, res.load(fn).get("table_id"), fn, "2024", system, nf, tbl)
    return dict(sheet=name, n_ident=n, temps=temps, last_row=last, recs=filas,
                pack_t0=t0, pack_v0=v0, npack=npack, npts_col=CC["n_pts"])


# El enlace SI<->US del Apendice C es POSICIONAL (clave_bi = propiedad#fila
# impresa): los nombres divergen entre ediciones por artefactos de impresion
# ('Type 309.' con punto en C-3 y con coma en C-3C, '25Cr–20Ni' con raya en una
# y con guion en la otra), asi que enlazar por nombre fallaria en esas filas.
# La contrapartida obligatoria de un enlace posicional es esta asercion: si el
# numero de filas por propiedad no coincide, o si un nombre normalizado difiere
# fuera de la lista blanca, el build ABORTA. Sin ella seria una bomba silenciosa.
DIVERGENCIAS_NOMBRE_C = {
    # normalizado SI -> normalizado US, con el motivo impreso
    ("C3", "TYPE309.23CR-12NI"): ("TYPE309,23CR-12NI",
                                  "C-3 imprime punto y C-3C coma tras 'Type 309'"),
}


def _norm_nombre_c(s) -> str:
    """Nombre comparable entre ediciones: sin espacios, con guiones plegados."""
    return re.sub(r"\s+", "", txt(s)).upper()


def verificar_paridad_apendice_c(si_info, us_info):
    a = {d["bi"]: d for d in si_info["recs"]}
    b = {d["bi"]: d for d in us_info["recs"]}
    if set(a) != set(b):
        raise SystemExit(
            "Apendice C: las dos ediciones no publican el mismo numero de filas por "
            f"propiedad. Solo en SI: {sorted(set(a) - set(b))[:8]}; "
            f"solo en US: {sorted(set(b) - set(a))[:8]}.")
    malas = []
    for bi, d in a.items():
        ns, nu = _norm_nombre_c(d["material"]), _norm_nombre_c(b[bi]["material"])
        if ns == nu:
            continue
        permitido = DIVERGENCIAS_NOMBRE_C.get((d["prop_id"], ns))
        if permitido and permitido[0] == nu:
            continue
        malas.append(f"{bi}: SI={d['material']!r} US={b[bi]['material']!r}")
    if malas:
        raise SystemExit(
            "Apendice C: el enlace SI/US es posicional y hay nombres que no casan "
            "fuera de la lista blanca documentada:\n  " + "\n  ".join(malas))
    ISSUES.append(
        f"Apendice C: {len(a)} filas enlazadas SI<->US por posicion (clave_bi); "
        f"{len(DIVERGENCIAS_NOMBRE_C)} divergencia(s) de nombre en lista blanca, "
        "todas artefactos de impresion documentados.")


def build_nometalicos(res, wb):
    """Formato largo (Tabla | Material | Campo | Valor): permite cascada
    Tabla -> Material y ficha campo/valor sin matrices dinamicas.

    Solo Apendice B. Las tablas C-2 y C-4 —dilatacion y modulo de no
    metalicos— salieron de aqui a DB_B31_C / DB_B31_CC: el Apendice B publica
    esfuerzos de diseno hidrostatico y presion admisible, que es otra cosa.
    Un dato, un motor.
    """
    tables = [("B-1", f"{APX}/appendix_b/table_b_1.json"),
              ("B-1C", f"{APX}/appendix_b/table_b_1c.json"),
              ("B-2", f"{APX}/appendix_b/table_b_2.json"),
              ("B-3", f"{APX}/appendix_b/table_b_3.json"),
              ("B-4", f"{APX}/appendix_b/table_b_4.json"),
              ("B-5", f"{APX}/appendix_b/table_b_5.json"),
              ("B-6", f"{APX}/appendix_b/table_b_6.json")]
    ws = new_sheet(wb, "DB_NoMetalicos",
                   "MATERIALES NO METALICOS — ASME B31.3-2024, Apendice B: esfuerzos de "
                   "diseno hidrostatico (HDS) y presiones admisibles",
                   "Fuente: resources/" + APX + "/appendix_b/table_b_1..b_6.json · "
                   "Formato largo "
                   "(tabla / material / campo / valor) para permitir la consulta por lista "
                   "desplegable. Valores tal como estan impresos. La dilatacion (C-2) y el "
                   "modulo (C-4) de no metalicos viven en DB_B31_C / DB_B31_CC.")
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
# Factores de calidad Ec (Tabla A-2) y Ej (Tabla A-3) — bases y notas
# ---------------------------------------------------------------------------
# Ec y Ej son multiplicadores directos del esfuerzo admisible en la ecuacion de
# diseno por presion, t = P·D / (2(S·E + P·Y)). Equivocar el factor mueve el
# espesor requerido hasta un 40 % —de E = 1,00 en un tubo sin costura a E = 0,60
# en uno soldado a tope en horno—, asi que no es un dato accesorio.
#
# Lo importante que estas dos tablas dicen, y que el libro no decia: el factor
# publicado es un MINIMO que puede subirse con examen suplementario. Ver
# build_buscador_factor y las Notas (4) y (5) de A-2.
FACT_COLS = [
    "material_id", "Tabla", "k0", "k1", "k2", "k3", "k4", "clave_bi",
    "Grupo impreso", "Spec. No.", "Clase o tipo", "Descripcion", "Factor",
    "Notas citadas", "Admite incremento", "Detalle del incremento",
    "Texto de las notas", "Paginas PDF (fuente)", "Archivo fuente",
    "Linea", "n_pts"]
CF = {n: i + 1 for i, n in enumerate(FACT_COLS)}
N_IDENT_F = len(FACT_COLS)
CLF = {n: get_column_letter(i) for n, i in CF.items()}
assert FACT_COLS[:8] == STRESS_COLS[:8], \
    "el contrato de las 8 primeras columnas es lo que reutiliza build_listas"

NO_APLICA = "(no aplica)"
INC_SI = "SI — admite incremento por examen suplementario [Nota (4)]"
INC_SI_YA = ("SI — admite incremento; ademas el factor basico ya supone examen "
             "suplementario [Notas (4) y (5)]")
INC_NO = "NO — el factor ya supone el examen suplementario [Nota (5)]"
INC_SILENCIO = "El codigo no se pronuncia en esta fila"

# El corte de pagina parte un grupo en dos ("Copper and Copper Alloy" y
# "Copper and Copper Alloy (Cont'd)"). No son dos grupos: es el mismo rotulo
# reimpreso al pasar de folio. Se normaliza para la navegacion y se declara
# fila a fila en la columna «Detalle del incremento» de la propia hoja.
_RE_CONTD = re.compile(r"\s*\(Cont'?d\)\s*$", re.I)


def _grupo_sin_contd(g):
    return _RE_CONTD.sub("", txt(g)).strip()


def notas_numeradas(textos) -> dict[str, str]:
    """Trocea las notas al pie de A-2 / A-3, que llegan como parrafos enteros.

    La extraccion entrega las notas de estas dos tablas como uno o dos strings
    largos, no como items. Trocearlas por «cualquier (n)» seria un error: el
    propio texto cita «para. 302.3.1(a)» y «Table A-1 (Table A-1C)». Se buscan
    los marcadores EN ORDEN —(1), luego (2) a partir de ahi, etc.—, que es lo
    unico que distingue un marcador de nota de una referencia interna.
    """
    entero = " ".join(clean(t) or "" for t in (textos or []))
    marcas = []
    desde = 0
    n = 1
    while True:
        pos = entero.find(f"({n})", desde)
        if pos < 0:
            break
        marcas.append((n, pos))
        desde = pos + 1
        n += 1
    out = {}
    for i, (num_, pos) in enumerate(marcas):
        fin = marcas[i + 1][1] if i + 1 < len(marcas) else len(entero)
        cuerpo = entero[pos + len(f"({num_})"):fin].strip()
        # El asterisco marca en el codigo las notas que repiten texto del cuerpo
        # normativo; se conserva tal como se imprime.
        out[str(num_)] = cuerpo
    return out


def _citadas(notas_txt) -> list[str]:
    """Numeros de nota que cita una fila ('(3), (4)' -> ['3','4'])."""
    return re.findall(r"\((\d+)\)", txt(notas_txt))


def build_ec_incremento(res, wb):
    """DB_Ec_Incremento — Tabla 302.3.3-1, los 6 examenes y su Ec.

    Es lo que convierte el factor basico de A-2 en el factor que de verdad
    aplica. Se cita SIEMPRE table_302_3_3_1.json: su gemelo de doble prefijo
    quedo declarado como duplicado por completar_tabla_302_3_3.py.
    """
    fn = "ASME B31/ASME B31.3/CHAPTERS/tables/table_302_3_3_1.json"
    d = res.load(fn)
    cols = d.get("columns") or []
    if len(cols) != 2:
        raise SystemExit("Tabla 302.3.3-1: se esperaban 2 columnas.")
    # Los encabezados impresos no se capturaron; completar_tabla_302_3_3.py
    # dejo en su lugar los derivados del para. 302.3.3(c), con su procedencia.
    rot = [clean(c.get("header") or c.get("header_derivado")) for c in cols]
    if not all(rot):
        raise SystemExit(
            "Tabla 302.3.3-1: las columnas no traen ni header ni header_derivado. "
            "Corra completar_tabla_302_3_3.py antes de construir el libro.")
    ws = new_sheet(
        wb, "DB_Ec_Incremento",
        f"FACTOR DE CALIDAD DE FUNDICION INCREMENTADO — ASME B31.3-2024, "
        f"{d.get('table_id')}: {d.get('title')}",
        f"Fuente: resources/{fn} · ASME B31.3-2024 · Los rotulos de columna son "
        "DERIVADOS del para. 302.3.3(c) (el impreso no se capturo); su procedencia "
        "va en fuente_derivacion dentro del JSON. Valores tal como estan impresos.")
    n = write_headers(ws, ["clave", "Tabla", "clave_sf", rot[0], rot[1],
                           "Notas de la tabla"])
    notas = " ".join(clean(t) or "" for t in d.get("notes") or [])
    recs = []
    for i, row in enumerate(d["rows"], start=1):
        recs.append(([f"302.3.3-1 #{i}", d.get("table_id"), d.get("table_id"),
                      clean(row.get("column_1")), row.get("column_2"), notas], {}))
    last = write_rows(ws, recs, n)
    autosize(ws, {"A": 16, "B": 18, "C": 18, "D": 44, "E": 12, "F": 120})
    ws.column_dimensions["C"].hidden = True
    ws.column_dimensions["F"].hidden = True
    record_meta("DB_Ec_Incremento", d.get("table_id"), fn, "2024", "adimensional",
                last - R_DATA + 1,
                "Ec incrementado por examen suplementario (para. 302.3.3(c)).")
    return dict(sheet="DB_Ec_Incremento", n_ident=n, last_row=last,
                col_examen=4, col_factor=5, notas=notas,
                tabla=d.get("table_id"), archivo=fn)


def build_factores(res, wb, cual):
    """DB_A2_Ec (24 filas) o DB_A3_Ej (127 filas), con el contrato de columnas."""
    a2 = cual == "A-2"
    fn = f"{APX}/appendix_a/table_a_{'2' if a2 else '3'}.json"
    d = res.load(fn)
    name = "DB_A2_Ec" if a2 else "DB_A3_Ej"
    campo = "ec" if a2 else "ej"
    notas = notas_numeradas(d.get("notes"))
    paginas = "-".join(str(p) for p in (d.get("pdf_pages") or []))
    folios = ((d.get("extraction_amendments") or {}).get("folios_impresos"))
    cita_pag = f"PDF {paginas}" + (f" · folios impresos {folios[0]}-{folios[-1]}"
                                   if folios else "")

    filas = []
    for i, row in enumerate(d["rows"], start=1):
        crudo = clean(g(row, "material_group"))
        grupo = _grupo_sin_contd(crudo) or "(sin grupo impreso)"
        cit = _citadas(g(row, "notes"))
        detalle = []
        if crudo and _RE_CONTD.search(crudo):
            detalle.append(f"El codigo reimprime el rotulo del grupo como "
                           f"«{crudo}» al pasar de folio; es corte de pagina, no "
                           f"un grupo distinto")
        if a2:
            # Se deriva de las notas que cita la fila, nunca de una heuristica.
            if "4" in cit and "5" in cit:
                admite = INC_SI_YA
            elif "4" in cit:
                admite = INC_SI
            elif "5" in cit:
                admite = INC_NO
            else:
                admite = INC_SILENCIO
        else:
            # A-3 no reparte el incremento fila a fila: su Nota general (2)
            # remite al para. 302.3.4(b) para TODA la tabla.
            admite = "VER para. 302.3.4(b) y Table 302.3.4-1"
        texto = " · ".join(f"({k}) {notas[k]}" for k in cit if k in notas)
        faltan = [k for k in cit if k not in notas]
        if faltan:
            detalle.append("no se pudo trocear la(s) nota(s) "
                           + ", ".join(f"({k})" for k in faltan))
        filas.append(dict(
            grupo=grupo, spec=clean(g(row, "spec_no")),
            clase=clean(g(row, "class_or_type")),
            desc=clean(g(row, "description")),
            factor=g(row, campo), notas=clean(g(row, "notes")),
            admite=admite, detalle=" · ".join(detalle) or None,
            texto=texto or None, idx=i))

    filas.sort(key=lambda x: (txt(x["grupo"]).upper(), txt(x["spec"]).upper(),
                              txt(x["clase"]).upper(), txt(x["desc"]).upper(),
                              x["idx"]))
    ids, _ = make_unique(
        [build_material_id([cual, x["spec"], x["clase"], x["desc"]]) for x in filas],
        [x["idx"] for x in filas])
    records = []
    for mid, x in zip(ids, filas):
        gr, sp = _seg(x["grupo"]), _seg(x["spec"])
        # A-2 tiene 3 niveles y A-3 cuatro; los no usados se rellenan con una
        # etiqueta seleccionable en vez de dejarse vacios, para que la cascada
        # pueda recorrerse igual en los dos motores.
        if a2:
            n1, n2, n3, n4 = sp, _seg(x["desc"]), NO_APLICA, NO_APLICA
        else:
            n1, n2, n3, n4 = (sp, _seg(x["clase"], NO_APLICA),
                              _seg(x["desc"]), NO_APLICA)
        k0 = x["grupo"]
        k1 = f"{k0}|{n1}"
        k2 = f"{k1}|{n2}"
        k3 = f"{k2}|{n3}"
        k4 = f"{k3}|{n4}"
        x["mid"], x["bi"] = mid, f"{'A2' if a2 else 'A3'}#{x['idx']}"
        records.append((
            [mid, cual, k0, k1, k2, k3, k4, x["bi"], x["grupo"], x["spec"],
             x["clase"], x["desc"], x["factor"], x["notas"], x["admite"],
             x["detalle"], x["texto"], cita_pag, f"resources/{fn}", x["idx"], 0],
            {}))

    ws = new_sheet(
        wb, name,
        f"FACTORES DE CALIDAD — ASME B31.3-2024, {d.get('table_id')}: {d.get('title')}",
        f"Fuente: resources/{fn} · ASME B31.3-2024 · {cita_pag} · "
        f"{'Ec' if a2 else 'Ej'} es adimensional y el codigo publica UNA sola tabla "
        "para los dos sistemas de unidades. El factor impreso es un MINIMO: vea la "
        "columna «Admite incremento» y el bloque 3 del motor.")
    n = write_headers(ws, FACT_COLS)
    last = write_rows(ws, records, n)
    autosize(ws, {CLF["material_id"]: 52, CLF["Tabla"]: 8, CLF["clave_bi"]: 10,
                  CLF["Grupo impreso"]: 30, CLF["Spec. No."]: 16,
                  CLF["Clase o tipo"]: 20, CLF["Descripcion"]: 56,
                  CLF["Factor"]: 9, CLF["Notas citadas"]: 12,
                  CLF["Admite incremento"]: 46, CLF["Detalle del incremento"]: 60,
                  CLF["Texto de las notas"]: 120,
                  CLF["Paginas PDF (fuente)"]: 30, CLF["Archivo fuente"]: 54})
    for cn in ("k0", "k1", "k2", "k3", "k4", "clave_bi", "Texto de las notas",
               "Linea", "n_pts"):
        ws.column_dimensions[CLF[cn]].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:{get_column_letter(n)}{last}"
    record_meta(name, d.get("table_id"), fn, "2024", "adimensional",
                last - R_DATA + 1, d.get("title"))
    return dict(sheet=name, n_ident=n, last_row=last, recs=filas,
                tabla=d.get("table_id"), archivo=fn, cita_pag=cita_pag,
                notas=notas)


def texto_incremento_ej(res):
    """Lo que el codigo dice sobre subir Ej, leido de resources/ (no de memoria).

    Devuelve (parrafo 302.3.4(b), Nota (1) de la Tabla 302.3.4-1, filas).
    `filas` es el cuerpo reconstruido de la Tabla 302.3.4-1 (10 filas: No.,
    tipo de junta, tipo de costura, examen, Ej), que completar_tabla_302_3_4.py
    escribe en resources/ a partir del folio impreso. Si la extraccion volviera
    a colapsarse (el sintoma que declaraba el hueco original), esto aborta en
    vez de mostrar un motor que aparenta tener el dato y no lo tiene.
    """
    cap = res.load("ASME B31/ASME B31.3/CHAPTERS/chapter_02.json")

    def _txt(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k == "text" and isinstance(v, str):
                    yield v
                else:
                    yield from _txt(v)
        elif isinstance(o, list):
            for v in o:
                yield from _txt(v)

    parrafo = next((clean(t) for t in _txt(cap)
                    if "Table 302.3.4-1 also indicates" in t), None)
    if not parrafo:
        raise SystemExit(
            "No se encuentra el para. 302.3.4(b) en chapter_02.json. Sin el, el "
            "motor de A-3 no puede declarar si el codigo publica un incremento "
            "para Ej, y no se afirma nada sin fuente.")
    t34 = res.load("ASME B31/ASME B31.3/CHAPTERS/tables/table_302_3_4_1.json")
    nota1 = None
    for t in t34.get("notes") or []:
        m = re.search(r"NOTE:\s*\(1\)\s*(.+?\.)", clean(t) or "")
        if m:
            nota1 = m.group(1)
            break
    cabeceras = " ".join(clean((c or {}).get("header")) or ""
                         for c in t34.get("columns") or [])
    if re.search(r"Factor,\s*Ej\s*0\.\d", cabeceras):
        raise SystemExit(
            "table_302_3_4_1.json volvio a colapsar su cuerpo dentro de los "
            "encabezados de columna (el sintoma que completar_tabla_302_3_4.py "
            "corrigio). No se construye Buscar_Ej_A3 con un dato que no esta: "
            "vuelva a correr completar_tabla_302_3_4.py --pdf <folio 302.3.4-1>.")
    filas = t34.get("rows") or []
    if len(filas) != 10 or any(r.get("factor_ej") is None for r in filas):
        raise SystemExit(
            f"table_302_3_4_1.json trae {len(filas)} filas y se esperan 10 "
            "(la Tabla 302.3.4-1 tiene 2 juntas sin subdividir + 3(a)/3(b) a "
            "tres examenes + la junta 4 a dos). Revise si alguien la toco a "
            "mano; no se construye el motor con un cuerpo distinto del que "
            "completar_tabla_302_3_4.py declaro.")
    return parrafo, nota1, filas


def bloques_tabla_ej(filas34):
    """Filas de Table 302.3.4-1 en el formato (texto, estilo) que consume el
    bloque 'declarado' de build_buscador_factor. Una linea por fila, integra y
    auditable: no se agrupan las filas de 3(a)/3(b)/4 para no esconder ninguna
    combinacion de junta/costura/examen detras de otra."""
    out = []
    for f in filas34:
        marca = "  [Nota (1): NO admite incremento por examen adicional]" \
            if f.get("note") else ""
        texto = (f"No. {f['no']} — {f['type_of_joint']} · Costura: "
                 f"{f['type_of_seam']} · Examen: {f['examination']} · "
                 f"Ej = {f['factor_ej']:.2f}{marca}")
        out.append((texto, Font(name="Calibri", size=9)))
    return out


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


# TE-1 no reparte toda la dilatacion por Grupos numerados: la mayor parte de sus
# columnas se autodescriben en el propio titulo («15Cr and 17Cr Steels»). Esas
# columnas son tan normativas como las Notas, y no leerlas dejaba sin alfa a
# materiales para los que el codigo SI la publica: los 9Cr-1Mo, los 12Cr/13Cr,
# los 15Cr/17Cr, los 27Cr, los 8Ni/9Ni.
#
# Se resuelven LITERALMENTE: del titulo se extraen las designaciones de
# composicion que enumera, y se exige coincidencia exacta. Lo que el titulo
# anade como CONDICION —un grado, un tratamiento termico— no se resuelve aqui.
_COL_TE = re.compile(r"^Coefficients for (.+?)\s*(?:Steels?|Steel)?\s*[ABC]$")
# Designacion de composicion: empieza por digito o por simbolo de elemento.
_ES_COMPOSICION = re.compile(r"^\d|^[A-Z][a-z]?[-–]")


def columnas_nombradas_te1(datos):
    """{clave_composicion: (etiqueta_columna, condicion_o_None)} desde TE-1.

    Devuelve solo las columnas que enumeran composiciones en su titulo. Las que
    llevan una condicion entre parentesis («Including Grades 9, 91, 911, and
    92») o una condicion de tratamiento («Condition 1075») se devuelven con esa
    condicion aparte, para que el motor las marque como REVISAR en vez de
    aplicarlas: ahi la pertenencia no la decide la composicion.
    """
    out: dict[str, tuple] = {}
    for col in datos.get("columns", []):
        m = _COL_TE.match(txt(col))
        if not m:
            continue
        cuerpo = m.group(1)
        if "Group" in cuerpo:
            continue                      # los Grupos 1..4 ya vienen por Nota
        condicion = None
        cond = re.search(r"\(([^)]*)\)|,\s*(Condition .+)$", cuerpo)
        if cond:
            condicion = re.sub(r"\s+", " ", (cond.group(1) or cond.group(2)))
            # El PDF parte palabras con guion al saltar de linea: «(In- cluding».
            condicion = condicion.replace("- ", "").strip()
            cuerpo = cuerpo[:cond.start()].strip()
        etiqueta = re.sub(r"\s+[ABC]$", "", txt(col))
        # Rotulos que envuelven la designacion sin formar parte de ella.
        cuerpo = re.sub(r"^Precipitation Hardened\s+", "", cuerpo)
        cuerpo = re.sub(r"\s+Stainless$", "", cuerpo)
        # Las dos ediciones no rotulan igual esta columna: la US imprime «7Ni
        # Steels» y la metrica «7% Nickel Steel». Sin unificarlo, el mismo
        # material recibiria dilatacion en una hoja y no en la otra.
        cuerpo = re.sub(r"(\d+)%\s*Nickel", r"\1Ni", cuerpo)
        # «12Cr, 12Cr-1Al, 13Cr, and 13Cr-4Ni» -> cuatro designaciones
        for p in re.split(r",(?!\s*Condition)|\band\b", cuerpo):
            # `Steels?` es parte del rotulo, no de la designacion: sin quitarlo
            # «9Cr-1Mo Steels» se convertia en la clave «9CR-1MOSTEELS», que no
            # casa con nada y dejaba la columna inservible sin dar error.
            p = re.sub(r"(?:\s+Stainless)?\s*Steels?$|\s+Stainless$", "",
                       p.strip(" .")).strip()
            if not p or not _ES_COMPOSICION.match(p):
                continue
            k = comp_key(p)
            previo = out.get(k)
            if previo is None:
                out[k] = (etiqueta, condicion)
                continue
            # La extraccion trunca el rotulo de algunas columnas (la «A» de un
            # grupo de tres), y la truncada puede perder la condicion. La
            # condicion es PEGAJOSA: si alguna variante de la columna la lleva,
            # la designacion queda condicionada. Si no, la version truncada
            # —«...17Cr-4Ni-4Cu Stainless», sin «Condition 1075»— haria que se
            # asignase la dilatacion eligiendo a ciegas entre dos tratamientos
            # termicos con valores distintos.
            cond_final = previo[1] or condicion
            etiq_final = etiqueta if len(etiqueta) > len(previo[0]) else previo[0]
            out[k] = (etiq_final, cond_final)
    return out


# `columnas_nombradas_te1` colapsa el tratamiento termico a UNA sola condicion
# por composicion (la "pegajosa": ver su docstring y el test que lo exige),
# porque asume una unica condicion por material. El 17Cr-4Ni-4Cu tiene DOS
# columnas B reales -Condition 1075 y Condition 1150-, y la segunda se pierde
# ahi. Esta funcion no colapsa nada: guarda TODAS las condiciones impresas
# para poder casarlas contra el tratamiento que la propia fila de materiales
# imprime en `Clase/Cond./Temple` (columna `class_condition_temper` de
# table_1a.json / table_3.json).
_COND_TRATAMIENTO = re.compile(r",?\s*Condition\s+(\d{3,4})\s*[ABC]?$")


def condiciones_tratamiento_te1(datos):
    """{clave_composicion: {numero_tratamiento: etiqueta_columna}} desde TE-1."""
    out: dict[str, dict[str, str]] = {}
    for col in datos.get("columns", []):
        texto = txt(col)
        m = _COND_TRATAMIENTO.search(texto)
        if not m:
            continue
        cuerpo = texto[:m.start()]
        cuerpo = re.sub(r"^Coefficients for\s+", "", cuerpo)
        cuerpo = re.sub(r"^Precipitation Hardened\s+", "", cuerpo)
        cuerpo = re.sub(r"\s+Stainless(\s+Steels?)?$", "", cuerpo).strip()
        if not cuerpo or not _ES_COMPOSICION.match(cuerpo):
            continue
        etiqueta = re.sub(r"\s+[ABC]$", "", texto)
        out.setdefault(comp_key(cuerpo), {})[m.group(1)] = etiqueta
    return out


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


def _spec_num(spec_txt) -> str:
    """Numero base de una designacion de especificacion, sin prefijo de letra
    ni puntuacion: "SA-217" (II-D) y "A217" (B31.3 Apendice A) dan los dos
    "217". Las dos fuentes y las dos ediciones escriben el mismo numero con
    guion, en dash o prefijo "S" distintos; el numero es lo unico estable."""
    m = re.search(r"\d+", txt(spec_txt))
    return m.group(0) if m else ""


def _publicada_por_iid(comp_p, notas_idx, cols_te):
    """Mecanismos LITERALES por los que TM-1/TE-1 publican una composicion.

    Son dos, y los dos estan impresos en el codigo con el mismo rango: la lista
    de miembros de una Nota, y el titulo de una columna de TE-1 que se
    autodescribe («Coefficients for 8Ni and 9Ni Steels»). El motor ya usaba los
    dos para las filas que SI imprimen su composicion; esta funcion existe para
    que el camino de composicion prestada por UNS use exactamente los mismos y
    no se quede solo con las Notas —que era el hueco: el 9Ni de K81340 tiene
    columna propia en TE-1 y aun asi salia como «II-D no publica el dato»—.

    Solo se acepta la columna SIN condicion. Una columna condicionada por grado
    («Including Grades 9, 91, 911, and 92») o por tratamiento («Condition
    1075») no la resuelve la composicion: la resuelve un dato que la fila tiene
    que imprimir aparte, y con la composicion prestada ese dato no se ha
    comprobado. Devolverla aqui seria elegir a ciegas entre columnas con
    valores distintos.

    Devuelve (origenes_de_nota, etiqueta_de_columna_o_None).
    """
    k = comp_key(comp_p)
    col = cols_te.get(k)
    return notas_idx.get(k), (col[0] if col and col[1] is None else None)


# Vias por las que se puede identificar el material de una fila que no imprime
# su composicion. Las tres se apoyan en un dato que la PROPIA FILA imprime -el
# UNS, la especificacion, el grado-, nunca en un criterio elegido aparte.
VIA_UNS, VIA_SPEC, VIA_GRADO = None, "spec", "grado"


def _candidatas_por_uns(comp_idx, comp_idx_spec, comp_idx_grado,
                        uns_txt, spec_txt, grado_txt):
    """Genera (composicion, hojas, via, clave) para el UNS de la fila.

    Primero el UNS a secas, IGUAL que siempre: si el libro trae una sola
    composicion para ese UNS, se toma de ahi (comportamiento sin cambios para
    los casos ya resueltos, y con el mismo motivo de siempre).

    Solo cuando el UNS a secas es AMBIGUO -mas de una composicion global- se
    prueban dos vias mas finas, en este orden:

      (UNS, especificacion impresa en la fila). Resuelve donde la ambiguedad
      global es en realidad el codigo repartiendo la composicion por
      especificacion: perno vs. tuerca en A193/A194, fundicion vs. tubo
      fundido en A217/A426.

      (UNS, grado impreso en la fila). Resuelve donde el codigo la reparte por
      GRADO, dentro o a traves de las especificaciones: G41400 imprime
      «Cr-Mo» para el B7 de A193 y «Cr-0.2Mo» para el B7M de la misma A193, de
      modo que la especificacion no basta; y S41000/SA-479 no tiene fila en el
      Apendice A, pero su grado impreso -«410»- si la tiene alli (A240 Gr.
      410, «13Cr»), y es el mismo grado, impreso igual por las dos fuentes.

    Probar las finas SOLO como fallback -nunca primero- importa: si se
    probaran siempre, un UNS con una unica composicion global tambien tendria
    un unico candidato en cualquier spec o grado suyo, y el motivo cambiaria
    de redaccion sin necesidad para las ~200 filas que ya resolvia bien el
    camino simple.

    Cuando ninguna de las tres deja un candidato UNICO, no se genera nada: no
    hay con que elegir sin inventar.
    """
    spec_n = _spec_num(spec_txt)
    grado_k = txt(grado_txt).upper()
    for tok in _uns_tokens(uns_txt):
        candidatas = comp_idx.get(tok)
        if not candidatas:
            continue
        if len(candidatas) == 1:
            comp_p, hojas = next(iter(candidatas.items()))
            yield comp_p, hojas, VIA_UNS, tok
            continue
        for indice, clave, via in ((comp_idx_spec, spec_n, VIA_SPEC),
                                   (comp_idx_grado, grado_k, VIA_GRADO)):
            if not clave:
                continue
            finas = indice.get((tok, clave))
            if finas and len(finas) == 1:
                comp_p, hojas = next(iter(finas.items()))
                yield comp_p, hojas, via, clave
                break


def _composicion_prestada(comp_idx, comp_idx_spec, comp_idx_grado, uns_txt,
                          spec_txt, grado_txt, ck, notas_idx, cols_te):
    """Recupera la composicion de una fila que no la imprime, via su UNS.

    Condiciones para devolver algo, en cualquiera de las tres vias de
    `_candidatas_por_uns`:
      1. la fila no trae composicion propia,
      2. el candidato es UNICO para la clave que se prueba,
      3. TM-1 o TE-1 publican esa composicion por alguno de sus dos mecanismos
         literales: la lista de miembros de una Nota, o el titulo de una
         columna de TE-1 que se autodescribe (ver _publicada_por_iid).
    Devuelve (composicion, hojas_de_origen, origenes_de_nota, columna_te, via,
    clave_usada) con `via` en None cuando se resolvio por el UNS global.
    """
    if ck:
        return None
    for comp_p, hojas, via, clave in _candidatas_por_uns(
            comp_idx, comp_idx_spec, comp_idx_grado, uns_txt, spec_txt, grado_txt):
        origenes, col_te = _publicada_por_iid(comp_p, notas_idx, cols_te)
        if origenes or col_te:
            return comp_p, hojas, origenes, col_te, via, clave
    return None


# Cierre de una fila que el codigo no resuelve. NO es un estado del mapeo -la
# fila sigue igual de bloqueada en los dos casos-: responde a otra pregunta,
# «queda esto esperando a alguien?». Sin ella, las 113 filas de la revision se
# presentaban todas como decisiones pendientes de firma, cuando la inmensa
# mayoria no admite decision alguna: el codigo sencillamente no publica el dato,
# y firmar ahi seria inventarlo. Distinguirlas es lo que permite CERRAR la
# revision sin rellenar ni una casilla que el codigo no respalde.
CIERRE_LIMITE = "CERRADA (limite de la fuente)"
CIERRE_ABIERTA = "ABIERTA (admite criterio de ingenieria)"


def _motivo_sin_composicion(comp_idx, comp_idx_spec, comp_idx_grado, uns_txt,
                            spec_txt, grado_txt, cols_te=None) -> tuple[str, str]:
    """(motivo, cierre) para una fila que no imprime composicion nominal.

    El motivo importa, y el cierre mas: no es lo mismo que el UNS no aparezca
    en ningun lado (nada que decidir, y nada que buscar: el indice ya recorre
    las dos ediciones de II-D y las dos del Apendice A del B31.3), que el
    codigo lo imprima con dos composiciones REALES y distintas que ni la
    especificacion ni el grado deshacen (ahi si hay algo que elegir), o que la
    composicion SI la publique TE-1 pero condicionada por un dato que esta fila
    no imprime (tambien decidible, con el codigo delante).

    Usa exactamente las mismas vias que `_composicion_prestada`: si una de
    ellas identifica el material y aun asi no hay grupo, el hueco es del
    codigo, no del mapeo, y la fila se CIERRA en vez de quedar pidiendo una
    firma que no cambiaria nada.
    """
    cols_te = cols_te or {}
    _COMO = {VIA_UNS: "", VIA_SPEC: "la especificacion", VIA_GRADO: "el grado"}
    for comp_p, _hojas, via, clave in _candidatas_por_uns(
            comp_idx, comp_idx_spec, comp_idx_grado, uns_txt, spec_txt, grado_txt):
        tok = next(iter(_uns_tokens(uns_txt)), "")
        if via is VIA_UNS:
            quien = f"Su UNS «{tok}» aparece como «{comp_p}»"
        else:
            quien = (f"Su UNS «{tok}» trae varias composiciones en el libro, pero "
                     f"{_COMO[via]} que la propia fila imprime («{clave}») deja "
                     f"una sola: «{comp_p}»")
        col = cols_te.get(comp_key(comp_p))
        if col and col[1]:
            # El dato SI esta publicado, pero la columna lo condiciona a un
            # grado o tratamiento. Decirlo asi -y no «II-D no publica el
            # dato»- es la diferencia entre un hueco de la fuente y una
            # decision que el ingeniero puede tomar con el codigo delante.
            return (f"La fila no imprime composicion nominal. {quien}, y TE-1 SI "
                    f"publica dilatacion para esa composicion, pero en una columna "
                    f"condicionada («{col[1]}»). La condicion no la resuelve la "
                    f"composicion prestada: exige un dato que esta fila tendria que "
                    f"imprimir y no se ha comprobado. No se elige por cuenta propia.",
                    CIERRE_ABIERTA)
        return (f"La fila no imprime composicion nominal. {quien}. El material queda "
                f"identificado; lo que falta es el dato: esa composicion no figura en "
                f"ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. "
                f"II-D no publica E ni dilatacion para ella.", CIERRE_LIMITE)
    for tok in _uns_tokens(uns_txt):
        candidatas = comp_idx.get(tok)
        if candidatas and len(candidatas) > 1:
            return (f"La fila no imprime composicion nominal y el libro asocia a su "
                    f"UNS «{tok}» mas de una: {'; '.join(candidatas)}. Ni la "
                    f"especificacion ni el grado que la propia fila imprime dejan una "
                    f"sola. No se elige por cuenta propia.", CIERRE_ABIERTA)
    return ("La fila no imprime composicion nominal y su UNS no aparece con "
            "composicion en ninguna tabla del libro: no hay dato con el que "
            "buscar el grupo.", CIERRE_LIMITE)


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
    """Indice {UNS: {composicion: [hojas que la imprimen]}} y sus dos versiones
    finas, {(UNS, num. de especificacion): ...} y {(UNS, grado impreso): ...}.

    Sirve para las filas que NO imprimen composicion nominal: su UNS suele
    aparecer con composicion en otra tabla del mismo libro. El UNS lo asigna
    SAE/ASTM y designa el mismo material, asi que la composicion es la misma;
    pero al venir de OTRA tabla —a veces de otro codigo, el Apendice A del
    B31.3— el resultado no puede presentarse como AUTO sin mas. El indice fino
    por especificacion permite que `_composicion_prestada` SI lo trate como
    AUTO (E_COMP_AJENA) cuando la propia fila imprime la especificacion y esa
    combinacion (UNS, especificacion) tiene una unica composicion en el libro:
    ya no es "el UNS trae una composicion" sino "el UNS CON ESTA
    ESPECIFICACION, que la fila ya imprime, trae una unica composicion" —nada
    que el ingeniero deba confirmar a criterio.

    El indice por GRADO existe por el mismo motivo y con el mismo rango: hay
    UNS que el codigo reparte por grado y no por especificacion (G41400 imprime
    «Cr-Mo» para el B7 de A193 y «Cr-0.2Mo» para el B7M de la MISMA A193), y
    hay filas cuya especificacion no aparece en la tabla que trae la
    composicion pero cuyo grado si (S41000/SA-479 no tiene fila en el Apendice
    A; su grado «410» si, en A240 Gr. 410). El grado se compara LITERAL, sin
    normalizar: A194 Gr. «6» y A193 Gr. «B6» son grados distintos del mismo
    UNS con composiciones distintas, y reducirlos a su digito los confundiria.
    """
    idx: dict[str, dict[str, list]] = {}
    idx_spec: dict[tuple, dict[str, list]] = {}
    idx_grado: dict[tuple, dict[str, list]] = {}
    for info in infos:
        ws = wb[info["sheet"]]
        for r in range(R_DATA, info["last_row"] + 1):
            uns = txt(ws.cell(r, C["UNS / Alloy"]).value).upper()
            comp = txt(ws.cell(r, C["Composicion nominal"]).value)
            spec = _spec_num(ws.cell(r, C["Spec. No."]).value)
            grado = txt(ws.cell(r, C["Tipo/Grado"]).value).upper()
            if uns and comp:
                idx.setdefault(uns, {}).setdefault(comp, [])
                if info["sheet"] not in idx[uns][comp]:
                    idx[uns][comp].append(info["sheet"])
                for fino, clave in ((idx_spec, spec), (idx_grado, grado)):
                    if not clave:
                        continue
                    celda = fino.setdefault((uns, clave), {}).setdefault(comp, [])
                    if info["sheet"] not in celda:
                        celda.append(info["sheet"])
    return idx, idx_spec, idx_grado


# Estados del mapeo. No existe ya un estado «PROPUESTA»: o el codigo lo dice y
# se cita la nota, o se declara el hueco y el calculo queda bloqueado.
E_UNS = "AUTO (UNS exacto)"
E_NOTA = "AUTO (composicion en Nota)"
# La composicion no la imprime la fila, pero SI la imprime el codigo para ese
# mismo UNS en otra de sus tablas. El eslabon anadido —«mismo UNS, misma
# composicion nominal»— se sostiene en que el UNS es un identificador univoco
# de material asignado por SAE/ASTM, no en criterio de nadie. Por eso es AUTO y
# no requiere firma; el estado lo dice en su nombre y la fuente cita la tabla de
# origen, de modo que la cadena queda a la vista para auditarla.
#
# En la Rev. 3 esto nacio como REVISAR por precaucion. Era una clasificacion
# equivocada: pedia firma de ingenieria para lo que es una lectura del codigo,
# y 386 filas quedaban bloqueadas esperando una decision que no existia.
E_COMP_AJENA = "AUTO (composicion via UNS en otra tabla)"
E_TEXTUAL = "REVISAR (regla textual del codigo)"
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
        # `validado_por` es opcional: se registra si esta, pero no se exige. Lo
        # que separa una decision de un dato del codigo es el ESTADO de la fila
        # —VALIDADO POR INGENIERO frente a AUTO (...)— y la justificacion que la
        # acompana, no la presencia de una cadena de texto. Exigirla solo anadia
        # ceremonia: ninguna regla de ASME PCC-2 la pide.
        if not txt(d.get("validado_por")):
            sin_firma += 1
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
        ISSUES.append(f"decisiones: {sin_firma} entradas aplicadas sin nombre en "
                      f"`validado_por`. Se aplican igual; el estado de la fila las "
                      f"marca como VALIDADO POR INGENIERO.")
    return por_comp, por_uns


def _firma(d) -> str:
    quien = txt(d.get("validado_por")) or "sin nombre"
    cuando = txt(d.get("fecha")) or "sin fecha"
    return f"{quien}, {cuando}"


def build_map_grupo(res, wb, iid_infos, comp_infos, ruta_decisiones,
                    system="SI"):
    # Las dos ediciones se mapean por separado contra SUS PROPIAS Notas: no
    # numeran igual y son extracciones independientes (ver extraer_notas_ii_d).
    si = system == "SI"
    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025" if si else "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
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
    comp_idx, comp_idx_spec, comp_idx_grado = indice_composicion_por_uns(
        wb, comp_infos)
    dec_comp, dec_uns = cargar_decisiones(ruta_decisiones)
    conflictos = {}

    notas_idx, textuales = cargar_notas_de_grupo(res, ed)
    te1_datos = res.load(f"{ed}/table_te_1.json")
    cols_te = columnas_nombradas_te1(te1_datos)
    cond_te = condiciones_tratamiento_te1(te1_datos)

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

            grp_e = fuente_e = grp_te = fuente_te = motivo = estado = None
            cierre = None

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

            # 2b) Columna nombrada de TE-1. TE-1 no reparte toda la dilatacion
            #     por Grupos numerados: la mayoria de sus columnas se
            #     autodescriben («15Cr and 17Cr Steels») y son tan normativas
            #     como las Notas. Sin leerlas, materiales con alfa publicada
            #     quedaban sin ella.
            if grp_te is None and ck in cols_te:
                etiqueta, condicion = cols_te[ck]
                if condicion is None:
                    grp_te, fuente_te = etiqueta, "TE-1 columna impresa"
                    if estado not in (E_UNS, E_NOTA, E_COMP_AJENA):
                        estado = E_NOTA
                elif condicion.startswith("Condition"):
                    # Condicion de TRATAMIENTO TERMICO: TE-1 parte este material
                    # en varias columnas con valores distintos. La propia fila SI
                    # suele imprimir su tratamiento en "Clase/Cond./Temple"
                    # (class_condition_temper: H1075, H1100, H1150...); si coincide
                    # con una columna de TE-1, se resuelve citando ese dato impreso.
                    # Elegir sin ese respaldo si seria inventar.
                    opciones = cond_te.get(ck, {})
                    temple = txt(src.cell(r, C["Clase/Cond./Temple"]).value)
                    m_temple = re.search(r"(\d{3,4})", temple)
                    if m_temple and m_temple.group(1) in opciones:
                        grp_te = opciones[m_temple.group(1)]
                        fuente_te = (f"TE-1 columna impresa · tratamiento "
                                     f"«{temple}» impreso en Clase/Cond./Temple")
                        if estado not in (E_UNS, E_NOTA, E_COMP_AJENA):
                            estado = E_NOTA
                    elif m_temple:
                        # La fila imprime un tratamiento real, pero TE-1 no
                        # publica columna para el: no es ambiguedad de mapeo,
                        # es que el codigo no tiene ese dato.
                        motivo = (f"La fila imprime el tratamiento «{temple}» en "
                                  f"Clase/Cond./Temple, pero TE-1 solo publica "
                                  f"dilatacion de «{txt(comp)}» para "
                                  + " y ".join(f"«Condition {c}»" for c in sorted(opciones))
                                  + ". El codigo no publica el dato para esta "
                                  f"condicion.")
                    else:
                        motivo = (f"TE-1 publica la dilatacion de «{txt(comp)}» en "
                                  f"columnas segun el tratamiento termico "
                                  + " y ".join(f"«Condition {c}»" for c in sorted(opciones))
                                  + ", con valores distintos, y esta fila no imprime "
                                  f"el tratamiento en Clase/Cond./Temple. Determine "
                                  f"la condicion y lea la columna en DB_TE.")
            elif grp_te is None:
                # La columna condiciona la pertenencia a un GRADO, no a la
                # composicion: «9Cr-1Mo Steels (Including Grades 9, 91, 911, and
                # 92)». El grado esta impreso en la propia fila, asi que la
                # comprobacion sigue siendo literal.
                # Los numeros del grado se extraen como tokens: quitar todo lo no
                # numerico convertia «F91 Type 2» en «912», que no casa con
                # ningun grado y dejaba fuera 7 de las 17 filas 9Cr-1Mo-V.
                grados_fila = re.findall(r"\d+", txt(src.cell(r, C["Tipo/Grado"]).value))
                for k_col, (etiqueta, condicion) in cols_te.items():
                    if not condicion or not grados_fila:
                        continue
                    if condicion.startswith("Condition"):
                        continue          # tratamiento termico: se trata arriba
                    grados = re.findall(r"\d+", condicion)
                    if (set(grados_fila) & set(grados)
                            and ck.split("-")[0] == k_col.split("-")[0]):
                        grp_te = etiqueta
                        fuente_te = f"TE-1 columna impresa · {condicion}"
                        if estado not in (E_UNS, E_NOTA, E_COMP_AJENA):
                            estado = E_NOTA
                        break

            # 3a) Regla de inclusion redactada, para el MODULO E. Se comprueba
            #     siempre que falte E, aunque la dilatacion ya este resuelta: si
            #     solo se mirase cuando faltan las dos, resolver alfa por columna
            #     haria desaparecer de la revision la unica pregunta que de
            #     verdad exige criterio (si el 9Cr-1Mo-V entra en la Nota (5)).
            if grp_e is None and ck:
                stem = ck.split("-")[0]
                cand = [t for t in textuales if _stem_textual(t[3]) == stem]
                if cand:
                    tabla, nota, grupo, miembro = cand[0]
                    estado = E_TEXTUAL
                    # El codigo redacta un criterio y no lista: eso SI es una
                    # decision de ingenieria, y la unica clase que queda abierta
                    # con respaldo textual.
                    cierre = CIERRE_ABIERTA
                    motivo = (f"{tabla} Nota {nota} lista «{miembro}». Aplicar la regla "
                              f"y confirmar si este material queda dentro de {grupo} "
                              f"(afecta solo al modulo E; la dilatacion, si figura, "
                              f"ya esta resuelta aparte).")

            # 3b) Nada resolvio: recuperar la composicion por UNS, o declarar el hueco.
            if hit is None and not (grp_e or grp_te) and estado != E_TEXTUAL:
                grado = txt(src.cell(r, C["Tipo/Grado"]).value)
                prestada = _composicion_prestada(
                    comp_idx, comp_idx_spec, comp_idx_grado, key, spec, grado,
                    ck, notas_idx, cols_te)
                if prestada:
                    # La fila no imprime composicion, pero su UNS -o su (UNS,
                    # especificacion) o su (UNS, grado) cuando eso es lo que
                    # desambigua- aparece con una sola composicion en otra tabla
                    # del libro, y esa composicion si la publica II-D. Es lectura
                    # del codigo, no criterio: el UNS identifica el material de
                    # forma univoca, y la especificacion y el grado, cuando hacen
                    # falta, los imprime la propia fila. Se resuelve, citando la
                    # tabla de la que sale la composicion.
                    comp_p, hojas_p, origenes, col_te_p, via_p, clave_p = prestada
                    for tabla, nota, grupo in (origenes or ()):
                        if tabla == "TM-1" and grp_e is None:
                            grp_e, fuente_e = grupo, f"TM-1 Nota {nota} · comp. de {hojas_p[0]}"
                        elif tabla == "TE-1" and grp_te is None:
                            grp_te, fuente_te = grupo, f"TE-1 Nota {nota} · comp. de {hojas_p[0]}"
                    # Las columnas nombradas de TE-1 son tan normativas como sus
                    # Notas, y hasta aqui solo las veian las filas que imprimen
                    # su composicion. Sin esto, un 9Ni recuperado por UNS se
                    # quedaba sin la dilatacion que el codigo SI publica para el,
                    # y las filas que la Nota resolvia solo para el modulo E
                    # (13Cr, 15Cr, 17Cr, 13Cr-4Ni) perdian su alfa en silencio.
                    if grp_te is None and col_te_p:
                        grp_te = col_te_p
                        fuente_te = f"TE-1 columna impresa · comp. de {hojas_p[0]}"
                    estado = E_COMP_AJENA
                    publica = " y ".join(
                        p for p in (("figura en Nota" if origenes else ""),
                                    ("tiene columna propia en TE-1" if col_te_p else ""))
                        if p)
                    if via_p is not VIA_UNS:
                        que = ("LA ESPECIFICACION" if via_p is VIA_SPEC
                               else "EL GRADO")
                        cual = "la especificacion" if via_p is VIA_SPEC else "el grado"
                        motivo = (f"La fila no imprime composicion nominal. Para su UNS "
                                  f"«{key}» el libro imprime mas de una composicion, "
                                  f"pero PARA {que} QUE IMPRIME ESTA FILA "
                                  f"(«{clave_p}») hay una sola: "
                                  f"«{comp_p}», en {', '.join(hojas_p)}. Esa composicion "
                                  f"{publica}. El UNS junto con {cual} "
                                  f"impreso identifica el material de forma univoca, asi "
                                  f"que el grupo se toma de ahi.")
                    else:
                        motivo = (f"La fila no imprime composicion nominal. Su UNS «{key}» "
                                  f"aparece como «{comp_p}» en {', '.join(hojas_p)}, y esa "
                                  f"composicion {publica}. El UNS identifica el "
                                  f"material de forma univoca, asi que el grupo se toma "
                                  f"de ahi.")
                elif not ck:
                    # Sin composicion impresa y sin forma de recuperarla: la fila
                    # del codigo no trae el dato de entrada.
                    estado = E_SIN
                    motivo, cierre = _motivo_sin_composicion(
                        comp_idx, comp_idx_spec, comp_idx_grado, key, spec,
                        grado, cols_te)
                else:
                    estado = E_SIN
                    # La fila IMPRIME su composicion: el material esta
                    # identificado y el contraste contra las dos vias literales
                    # del codigo ya se hizo. No queda nada que decidir salvo
                    # inventar el dato, asi que la fila se cierra como limite de
                    # la fuente en vez de quedar pidiendo una firma.
                    cierre = CIERRE_LIMITE
                    motivo = ("El UNS no figura en TM-1..TM-5 y la composicion nominal "
                              "no esta listada en ninguna Nota de TM-1 ni TE-1 ni en "
                              "ninguna columna nombrada de TE-1. "
                              "II-D no publica E ni dilatacion para este material.")
            elif grp_e is None or grp_te is None:
                # `motivo or` porque un motivo especifico —el del tratamiento
                # termico del 17Cr-4Ni-4Cu, por ejemplo— explica MEJOR el hueco
                # que la frase generica, y se estaba perdiendo al pisarlo.
                falta = "E (TM-1)" if grp_e is None else "dilatacion (TE-1)"
                motivo = motivo or f"Solo se resolvio uno de los dos grupos; falta {falta}."

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
            if estado in (E_TEXTUAL, E_SIN):
                # Las filas sin composicion propia se agrupan por UNS: es la
                # unidad en que se decide, y agruparlas por la composicion
                # vacia las juntaria todas en una decision falsa.
                if txt(comp):
                    clave = ("composicion", txt(comp))
                else:
                    tok = next(iter(_uns_tokens(key)), "")
                    clave = ("uns", tok) if tok else ("composicion", "")
                pendientes.append((clave, estado, motivo, txt(spec), txt(uns),
                                   cierre or CIERRE_LIMITE))
            recs.append(([mid, spec, uns, comp, grp_e, fuente_e, grp_te, fuente_te,
                          prd, estado, motivo,
                          search_key(mid, spec, uns, comp)], {}))
    last = write_rows(ws, recs, n)
    autosize(ws, {"A": 46, "B": 12, "C": 14, "D": 24, "E": 26, "F": 18, "G": 20,
                  "H": 18, "I": 20, "J": 28, "K": 60, "L": 28})
    ws.column_dimensions["L"].hidden = True
    ws.auto_filter.ref = f"A{R_HDR}:K{last}"
    ISSUES.append(f"{nombre}: " + " · ".join(f"{k}={v}" for k, v in stats.items()))
    # El estado dice de DONDE sale el grupo, no si estan los dos. Una fila con
    # dilatacion pero sin modulo se leia como «AUTO» a secas y su hueco no lo
    # contaba nadie. TM-1 y TE-1 no listan los mismos materiales, asi que la
    # cobertura parcial es lo normal y hay que publicarla.
    sin_e = sum(1 for ident, _ in recs if not ident[4])
    sin_te = sum(1 for ident, _ in recs if not ident[6])
    ISSUES.append(
        f"{nombre} cobertura: {len(recs) - sin_e} filas con modulo E, "
        f"{len(recs) - sin_te} con dilatacion, {len(recs)} en total. "
        f"Una fila puede tener uno y no el otro: TM-1 y TE-1 no enumeran los "
        f"mismos materiales. La columna Motivo lo dice fila a fila.")
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


def _filas_revision(entradas, inicio: int) -> list[str]:
    """Filas de la tabla de revision, con su casilla de grupo y de firma."""
    out = []
    for i, ((clave, estado), g) in enumerate(entradas, start=inicio):
        tipo, valor = clave
        specs = ", ".join(s for s, _ in g["specs"].most_common(3))
        if len(g["specs"]) > 3:
            specs += f", +{len(g['specs']) - 3} mas"
        motivo = (g["motivo"] or "").replace("\n", " ")
        etiqueta = f"`{valor}`" + (" (UNS)" if tipo == "uns" else "")
        out.append(f"| {i} | {etiqueta} | {g['n']} | {estado} | {motivo} "
                   f"<br>Especificaciones: {specs or '—'} |  |  |")
    return out


def escribir_revision_map_grupo(pendientes, stats, ruta: Path) -> int:
    """Hoja de revision: lo que el codigo NO resuelve, agrupado y CERRADO.

    Se agrupa por COMPOSICION NOMINAL porque es la unidad en que el codigo
    decide: las 3 454 filas de material se reducen a unas decenas de decisiones
    reales, y revisar fila a fila seria repetir el mismo juicio cientos de
    veces. Se acompana el recuento de materiales afectados para que se vea que
    pesa cada decision.

    El documento va en DOS bloques, y esa separacion es lo importante. Hasta la
    Rev. 4b todo lo no resuelto se listaba junto, bajo el rotulo «decisiones a
    tomar», con una casilla de firma al lado. Eso describia mal la realidad:
    solo una fraccion minima admite decision. En el resto, el material esta
    identificado sin ambiguedad y es II-D quien no tabula el dato — firmar ahi
    no seria decidir, seria inventar el valor que el codigo no publica. Se
    separan por `cierre`, que el motor deriva del mismo camino por el que llego
    al hueco, no de una lectura del texto del motivo.
    """
    from collections import Counter, OrderedDict
    grupos: "OrderedDict[tuple, dict]" = OrderedDict()
    for clave, estado, motivo, spec, uns, cierre in pendientes:
        g = grupos.setdefault((clave, estado), {"motivo": motivo, "n": 0,
                                                "specs": Counter(), "uns": Counter(),
                                                "cierre": cierre})
        g["n"] += 1
        if spec:
            g["specs"][spec] += 1
        if uns:
            g["uns"][uns] += 1

    # Filas sin composicion impresa Y sin UNS con el que anclar la decision: no
    # hay dato de entrada que juzgar. Se cuentan aparte para no inflar la lista.
    sin_comp = grupos.pop((("composicion", ""), E_SIN), None)

    orden = {E_TEXTUAL: 0, E_SIN: 1}
    filas = sorted(grupos.items(), key=lambda kv: (orden[kv[0][1]], -kv[1]["n"]))
    abiertas = [f for f in filas if f[1]["cierre"] == CIERRE_ABIERTA]
    cerradas = [f for f in filas if f[1]["cierre"] != CIERRE_ABIERTA]

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
          f"Casos distintos sin resolver por el codigo: **{len(filas)}** "
          f"(sobre {sum(v['n'] for _, v in grupos.items())} filas de material).",
          "",
          f"- **{len(abiertas)} ABIERTOS** — admiten criterio de ingenieria: el codigo "
          f"dice algo sobre ese material (una regla redactada, una columna condicionada, "
          f"o dos composiciones reales entre las que elegir) y una persona puede "
          f"resolverlo con el codigo delante.",
          f"- **{len(cerradas)} CERRADOS** — limite de la fuente: el material esta "
          f"identificado sin ambiguedad y II-D sencillamente no tabula ni modulo E ni "
          f"dilatacion para el. **No hay nada que firmar aqui.** Rellenar una casilla "
          f"seria inventar un valor que el codigo no publica, que es justo lo que la "
          f"Regla n.º 1 del proyecto prohibe. Se listan para que conste que se miraron "
          f"y por que vias.",
          ]
    if sin_comp:
        L += ["",
              f"Aparte, **{sin_comp['n']} filas no imprimen composicion nominal** y su "
              "UNS no figura en TM-1..TM-5.",
              "No entran en esta revision porque no hay dato de entrada que juzgar: para",
              "asignarles grupo habria que identificar el material por otra via (la",
              "especificacion y el grado en la tabla de origen)."]

    L += ["",
          "Las filas se cuentan sobre las DOS ediciones (metrica y U.S. Customary):",
          "un mismo caso afecta al material en ambas, porque la pertenencia a grupo no",
          "depende del sistema de unidades.",
          "",
          "## 1. Casos ABIERTOS — admiten criterio de ingenieria",
          ""]
    if abiertas:
        L += ["| # | Se decide sobre | Filas | Estado | Que hay que decidir | Grupo asignado | Firma / fecha |",
              "|---:|---|---:|---|---|---|---|"]
        L += _filas_revision(abiertas, 1)
    else:
        L += ["*Ninguno.* Todo lo que el codigo permitia decidir esta decidido y "
              "registrado en `decisiones_map_grupo.json`; el resto es limite de la "
              "fuente y esta cerrado en el bloque siguiente."]
    L += ["",
          "## 2. Casos CERRADOS — limite de la fuente, no requieren firma",
          "",
          "Para cada uno se agotaron las dos vias literales que el codigo publica —la",
          "lista de miembros de una Nota de TM-1/TE-1 y el titulo de una columna",
          "nombrada de TE-1— en LAS DOS ediciones, y la identificacion del material se",
          "intento ademas por UNS y por (UNS, especificacion impresa en la propia fila)",
          "contra las cuatro tablas indexadas del libro (II-D 1A y 1B/3, y Apendice A",
          "del B31.3, cada una en sus dos ediciones). El resultado es el correcto: la",
          "fila queda BLOQUEADA para modulo E y dilatacion.",
          "",
          "| # | Material | Filas | Estado | Por que esta cerrado |",
          "|---:|---|---:|---|---|"]
    for i, ((clave, estado), g) in enumerate(cerradas, start=1):
        tipo, valor = clave
        specs = ", ".join(s for s, _ in g["specs"].most_common(3))
        if len(g["specs"]) > 3:
            specs += f", +{len(g['specs']) - 3} mas"
        motivo = (g["motivo"] or "").replace("\n", " ")
        etiqueta = f"`{valor}`" + (" (UNS)" if tipo == "uns" else "")
        L.append(f"| {i} | {etiqueta} | {g['n']} | {estado} | {motivo} "
                 f"<br>Especificaciones: {specs or '—'} |")
    L += ["",
          "## Como usar este documento",
          "",
          "1. El bloque 2 no se rellena. Esta cerrado: la unica forma de darle grupo a",
          "   esos materiales seria que ASME publicase el dato, que hoy no publica.",
          "2. En el bloque 1, decida el grupo de TM-1 (modulo E) y/o TE-1 (dilatacion)",
          "   que corresponde, anotelo y firme.",
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
        "_solo_lo_decidible": (
            "`decisiones` trae UNICAMENTE los casos ABIERTOS de "
            "Revision_MAP_Grupo.md: aquellos en que el codigo dice algo que una "
            "persona puede aplicar. Los CERRADOS por limite de la fuente van "
            "aparte, en `_cerrados_sin_decision`, y NO se rellenan: el material "
            "esta identificado y es II-D quien no publica el dato. Ofrecerlos con "
            "una casilla vacia invitaba a rellenarla, que es inventar el valor."),
        "decisiones": [
            {("uns" if tipo == "uns" else "composicion"): valor,
             "grupo_tm": "", "grupo_te": "",
             "justificacion": "", "validado_por": "", "fecha": "",
             "_filas_afectadas": g["n"], "_estado_actual": estado,
             "_motivo": (g["motivo"] or "").replace("\n", " ")}
            for ((tipo, valor), estado), g in abiertas
        ],
        "_cerrados_sin_decision": [
            {("uns" if tipo == "uns" else "composicion"): valor,
             "_filas_afectadas": g["n"], "_estado_actual": estado,
             "_por_que_esta_cerrado": (g["motivo"] or "").replace("\n", " ")}
            for ((tipo, valor), estado), g in cerradas
        ],
    }
    ruta.with_name("decisiones_map_grupo.plantilla.json").write_text(
        json.dumps(plantilla, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(abiertas), len(cerradas)


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
            ("II-D", "1A", "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/notes_table_1a.json"),
            ("II-D", "1B", "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/notes_table_1b.json"),
            ("II-D", "3", "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/notes_table_3.json"),
            ("II-D", "U", "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/notes_table_u.json"),
            ("II-D", "Y-1", "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/notes_table_y_1.json")]
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

# Celda unica de escritura (temperatura de consulta): relleno propio, distinto
# del amarillo IN_FILL de las listas desplegables, para que salte a la vista
# cual es la unica celda que se teclea en todo el buscador.
TEMP_INPUT_FILL = PatternFill("solid", fgColor="CCC0DA")
# Semaforo de cascada completa/incompleta (formato condicional sobre el
# indicador de seleccion, ver build_buscador/finish_buscador).
SEL_OK_FILL = PatternFill("solid", fgColor="C6EFCE")
SEL_OK_FONT = Font(name="Calibri", size=11, bold=True, color="006100")
SEL_BAD_FILL = PatternFill("solid", fgColor="FFEB9C")
SEL_BAD_FONT = Font(name="Calibri", size=11, bold=True, color="9C6500")


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


def campo(ws, r, c1, etiqueta, formula, unidad=None,
         com_etq=None, com_val=None, com_uni=None):
    """etiqueta (2 col) | valor (2 col) | unidad (1 col)."""
    e = _mrg(ws, r, c1, c1 + 1, etiqueta)
    e.font = LBL2_F
    e.alignment = Alignment(vertical="center", indent=1)
    if com_etq:
        _nota(e, com_etq)
    v = _mrg(ws, r, c1 + 2, c1 + 3, formula)
    v.font = VAL_F
    v.alignment = Alignment(vertical="center", wrap_text=True)
    if com_val:
        _nota(v, com_val)
    u = ws.cell(r, c1 + 4)
    if unidad:
        u.value = unidad
    u.font = UNIT_F
    if com_uni:
        _nota(u, com_uni)
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
    filas = [(5, "Sistema de unidades", "SI",
              "Entrada: elija SI (metrico, valores en {u_si}, temperatura en {t_si}) "
              "o US (U.S. Customary, {u_us} / {t_us}). Cambia que hoja de base de datos "
              "y que edicion del codigo lee todo el buscador.".format(
                  u_si=unit_si, t_si=temp_si, u_us=unit_us, t_us=temp_us)),
             (6, "0 · Familia de material", "",
              "Entrada: paso 0 de la cascada. Elija la familia de material de la lista "
              "desplegable (agrupacion de navegacion derivada del codigo, no es un dato "
              "normativo). Habilita la lista del paso 1."),
             (7, "1 · Composicion nominal", "",
              "Entrada: paso 1 de la cascada, dependiente del paso 0. Lista solo con las "
              "composiciones que existen dentro de la familia elegida arriba."),
             (8, "2 · Forma de producto", "",
              "Entrada: paso 2 de la cascada, dependiente de los pasos 0 y 1 (forma de "
              "producto impresa por el codigo: placa, tubo, forjado, etc.)."),
             (9, "3 · Especificacion (Spec. No.)", "",
              "Entrada: paso 3 de la cascada, dependiente de los pasos 0 a 2. Spec. No. "
              "tal como lo imprime la tabla del codigo."),
             (10, "4 · Tipo / Grado", "",
              "Entrada: paso 4 de la cascada, dependiente de los pasos 0 a 3. Al "
              "completar este paso el buscador ya resuelve un material_id (salvo que "
              "existan varias variantes, ver paso 5)."),
             (11, "5 · Variante (clase / tamano) — opcional", "",
              "Entrada opcional: solo hace falta si el paso 4 deja mas de una fila "
              "posible (mismo material_id repetido por clase, condicion o tamano). Si "
              "se deja en blanco, el buscador toma la primera variante encontrada.")]
    for r, et, val, com in filas:
        e = _mrg(ws, r, 1, 3, et)
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        _nota(e, com)
        c = _mrg(ws, r, 4, 6, val)
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        ws.row_dimensions[r].height = 17
    e = _mrg(ws, 12, 1, 3, "TEMPERATURA DE CONSULTA   (unica celda de escritura)")
    e.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    e.alignment = Alignment(vertical="center", indent=1)
    com_temp = ("Entrada: la UNICA celda de escritura libre de todo el buscador. "
               "Temperatura a la que se necesita el valor, en la unidad que muestra la "
               "celda de la derecha (segun el selector SI/US de arriba). Todo el "
               "resultado, la ficha tecnica y la curva se recalculan a partir de este "
               "valor.")
    _nota(e, com_temp)
    c = _mrg(ws, 12, 4, 5, 25)
    c.font, c.fill = IN_F, TEMP_INPUT_FILL
    c.border = Border(*[Side("medium", color="C00000")] * 4)
    _nota(c, com_temp)
    ws.cell(12, 6).value = f"={U_TMP}"
    ws.cell(12, 6).font = UNIT_F
    _nota(ws.cell(12, 6), "Calculo: unidad de la temperatura de consulta; cambia entre "
                          "°C y °F segun el selector 'Sistema de unidades' (D5).")
    e = _mrg(ws, 13, 1, 3, "Modo de lectura")
    e.font = LBL_F
    e.alignment = Alignment(vertical="center", indent=1)
    com_modo = ("Entrada: 'Interpolado' aplica la interpolacion lineal del codigo "
               "S = S1 + (S2-S1)*(T-T1)/(T2-T1) entre los dos puntos tabulados que "
               "rodean la temperatura de consulta. 'Tabulado-conservador' ignora la "
               "interpolacion y adopta directamente el valor tabulado en T2 (el "
               "escalon superior), mas conservador.")
    _nota(e, com_modo)
    c = _mrg(ws, 13, 4, 6, "Interpolado")
    c.font, c.fill, c.border = IN_F, IN_FILL, BOX
    ay = _mrg(ws, 5, 7, NCOLS)
    ay.value = (f'=IF($D$5="SI","Leyendo {master["sheet"]} — valores en {unit_si}, '
                f'temperatura en {temp_si}","Leyendo {us["sheet"]} — valores en {unit_us}, '
                f'temperatura en {temp_us}")')
    ay.font = Font(italic=True, color=BLUE)
    _nota(ay, "Aviso automatico: confirma que hoja base de datos y que unidades esta "
              "leyendo el buscador, segun el selector 'Sistema de unidades' (D5). No "
              "se edita.")
    ay2 = _mrg(ws, 12, 7, 9)
    ay2.value = '=IF($D$10="","SELECCION INCOMPLETA","SELECCION COMPLETA")'
    ay2.font = SEL_BAD_FONT
    ay2.fill = SEL_BAD_FILL
    ay2.border = BOX
    ay2.alignment = Alignment(horizontal="center", vertical="center")
    _nota(ay2, "Aviso automatico: SELECCION COMPLETA (fondo verde) si la cascada de "
               "seleccion (pasos 0 a 4) esta completa; SELECCION INCOMPLETA (fondo "
               "amarillo) si falta elegir algun nivel para poder mostrar un resultado. "
               "No se edita.")
    ws.conditional_formatting.add(
        "G12:I12",
        FormulaRule(formula=['$D$10<>""'], fill=SEL_OK_FILL, font=SEL_OK_FONT))
    ws.conditional_formatting.add(
        "G12:I12",
        FormulaRule(formula=['$D$10=""'], fill=SEL_BAD_FILL, font=SEL_BAD_FONT))

    dv_list(ws, "D5", '"SI,US"', filas[0][3])
    dv_list(ws, "D13", '"Interpolado,Tabulado-conservador"', com_modo)
    dv_list(ws, "D6", "=" + rng["FAM"], filas[1][3])
    hl = get_column_letter
    levels = [("C", "D7", rng["CK"], rng["CV"], "$D$6", rng.get("maxC", 60), filas[2][3]),
              ("F", "D8", rng["FK"], rng["FV"], '$D$6&"|"&$D$7', rng.get("maxF", 30),
               filas[3][3]),
              ("S", "D9", rng["SK"], rng["SV"], '$D$6&"|"&$D$7&"|"&$D$8',
               rng.get("maxS", 40), filas[4][3]),
              ("G", "D10", rng["GK"], rng["GV"],
               '$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9', rng.get("maxG", 40), filas[5][3]),
              ("V", "D11", rng["K4"], rng["ID"],
               '$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9&"|"&$D$10', rng.get("maxV", 40),
               filas[6][3])]
    for tag, cell, kr, vr, key, mx, com in levels:
        col = CASC_COL[tag]
        L = hl(col)
        mx = max(1, min(int(mx), 250))
        ws.cell(R_HDR, col, f"lista {tag}").font = SRC_F
        for i2 in range(1, mx + 1):
            ws.cell(R_DATA + i2 - 1, col).value = (
                f'=IF(COUNTIF({kr},{key})<{i2},"",'
                f'INDEX({vr},MATCH({key},{kr},0)+{i2}-1))')
        ws.column_dimensions[L].hidden = True
        dv_list(ws, cell, f"=${L}${R_DATA}:${L}${R_DATA + mx - 1}", com)

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
    _nota(mid, "Calculo: nombre del material_id resuelto por la cascada de "
              "seleccion, o el aviso de que falta completarla.")
    ws.row_dimensions[16].height = 20
    kpis = [(1, ctx["valor_lbl"].upper() + " A LA TEMPERATURA DE CONSULTA",
             "=" + VALC, f"={U_VAL}",
             f"Calculo: {ctx['valor_lbl']} interpolado (o tabulado-conservador, "
             "segun el Modo de lectura) a la temperatura de consulta, entre los dos "
             "puntos tabulados T1/T2 de la seccion 4. NA() si el material esta fuera "
             "de rango."),
            (4, "TEMPERATURA DE CONSULTA", "=$D$12", f"={U_TMP}",
             "Calculo: repite la temperatura tecleada en la celda D12, para dejarla "
             "junto al resultado."),
            (7, "MODO DE LECTURA", "=$D$13", '="segun MODO_S"',
             "Calculo: repite el modo de lectura elegido en D13 (Interpolado o "
             "Tabulado-conservador)."),
            (10, "ESTADO DEL RANGO", "=" + EST,
             f'=IF({FIL}="","","variantes de este grado: "&{ctx["NVAR"]})',
             "Calculo: EN RANGO si T esta entre el primer punto tabulado y la Temp. "
             "max. del material; FUERA DE RANGO si no hay valor tabulado o si T "
             "supera la Temp. max. (el codigo prohibe extrapolar); SIN SELECCION si "
             "falta completar la cascada.")]
    for c1, tit, val, uni, com in kpis:
        t = _mrg(ws, 17, c1, c1 + 2, tit)
        t.font, t.fill = KPI_TIT_F, BAND_FILL
        t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        v = _mrg(ws, 18, c1, c1 + 2, val)
        v.font, v.fill = KPI_VAL_F, KPI_FILL
        v.alignment = Alignment(horizontal="center", vertical="center")
        _nota(v, com)
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

    izq = [("Tabla del codigo", val_of("Tabla"), None,
            "Calculo: tabla del codigo (A-1/A-4, 1A, 1B/3, U o Y-1 segun el buscador) "
            "de la que proviene la fila resuelta por la cascada de seleccion."),
           ("Familia de material", val_of("Familia"), '="agrupacion de navegacion"',
            "Calculo: familia de material tal como quedo clasificada en la base "
            "(agrupacion de navegacion derivada del UNS y la composicion impresa; no "
            "es un dato normativo y no entra en ningun calculo)."),
           ("Composicion nominal", val_of("Composicion nominal"), None,
            "Calculo: composicion nominal impresa por el codigo para el material "
            "resuelto por la cascada."),
           ("Forma de producto", val_of("Forma de producto"), None,
            "Calculo: forma de producto (placa, tubo, forjado, etc.) impresa por el "
            "codigo para el material resuelto."),
           ("Especificacion (Spec. No.)", val_of("Spec. No."), None,
            "Calculo: numero de especificacion (Spec. No.) impreso por el codigo para "
            "el material resuelto."),
           ("Tipo / Grado", val_of("Tipo/Grado"), None,
            "Calculo: tipo o grado impreso por el codigo para el material resuelto."),
           ("UNS / Alloy No.", val_of("UNS / Alloy"), None,
            "Calculo: numero UNS o Alloy No. impreso por el codigo; identifica la "
            "composicion del material, no su grupo de propiedades (ver MAP_Grupo)."),
           ("Clase / Condicion / Temple", val_of("Clase/Cond./Temple"), None,
            "Calculo: clase, condicion de tratamiento termico o temple impresos por "
            "el codigo, cuando aplica a esta fila."),
           ("Tamano / Espesor", val_of("Tamano/Espesor"),
            f'=IF($D$5="SI","mm","in")',
            "Calculo: rango de tamano o espesor al que aplica el admisible de esta "
            "fila, en la unidad de la celda de la derecha."),
           ("Notas del codigo", val_of("Notas"), '="ver Notas_Codigo"',
            "Calculo: numero(s) de nota al pie del codigo que restringen este "
            "material (soldadura, PWHT, servicio). Texto completo en Notas_Codigo.")]
    der = [("Edicion consultada",
            f'=IF({FIL}="","—",IF($D$5="SI","{master["sheet"]} — metrica",'
            f'IF({ctx["FUS"]}="","sin equivalente en la edicion US",'
            f'"{us["sheet"]} — U.S. Customary")))', None,
            "Calculo: confirma si la fila activa viene de la edicion SI o US, y avisa "
            "si el material no tiene equivalente publicado en la otra edicion."),
           ("P-No.", val_of("P-No."), '="adimensional"',
            "Calculo: P-No. (numero de material para WPS/PQR segun ASME IX) impreso "
            "por el codigo."),
           ("Group No.", val_of("Group No."), '="adimensional"',
            "Calculo: Group No. (subgrupo del P-No. para WPS/PQR) impreso por el "
            "codigo."),
           ("Temp. min. / curva de impacto", val_of("Temp. min. / curva impacto"),
            f'=IF(ISNUMBER({FIL}),{U_TMP},"")',
            "Calculo: temperatura minima de diseno o curva de impacto aplicable, "
            "impresa por el codigo para este material."),
           ("Resistencia a la traccion min.", val_of("Resist. traccion min."),
            f"={U_VAL}",
            "Calculo: Su minima especificada por la norma del material (valor de "
            "fabricacion; no es el esfuerzo admisible a la temperatura de consulta)."),
           ("Limite de fluencia min.", val_of("Fluencia min."), f"={U_VAL}",
            "Calculo: Sy minimo especificado por la norma del material (valor de "
            "fabricacion; no es el esfuerzo admisible a la temperatura de consulta)."),
           ("Temp. max. admisible / limite", val_of("Temp. max. / limite"),
            f"={U_TMP}",
            "Calculo: temperatura maxima admisible o limite de aplicabilidad de esta "
            "fila. Por encima de este valor el ESTADO DEL RANGO marca FUERA DE RANGO "
            "y el resultado queda bloqueado — el codigo prohibe extrapolar."),
           ("Aplicabilidad  I / III",
            f'=IF({FIL}="","—",IF(INDEX({ctx["ident"]("I")},{FIL})&'
            f'INDEX({ctx["ident"]("III")},{FIL})="","no aplica (tabla del B31.3)",'
            f'INDEX({ctx["ident"]("I")},{FIL})&"  /  "&'
            f'INDEX({ctx["ident"]("III")},{FIL})))', '="NP = no permitido"',
            "Calculo: aplicabilidad del material en las Divisiones I y III de la "
            "ASME BPVC, segun las columnas impresas por el codigo (NP = no "
            "permitido). No aplica en tablas del B31.3."),
           ("Aplicabilidad  VIII-1 / VIII-2 / XII",
            f'=IF({FIL}="","—",IF(INDEX({ctx["ident"]("VIII-1")},{FIL})&'
            f'INDEX({ctx["ident"]("VIII-2")},{FIL})&INDEX({ctx["ident"]("XII")},{FIL})="",'
            f'"no aplica (tabla del B31.3)",'
            f'INDEX({ctx["ident"]("VIII-1")},{FIL})&"  /  "&'
            f'INDEX({ctx["ident"]("VIII-2")},{FIL})&"  /  "&'
            f'INDEX({ctx["ident"]("XII")},{FIL})))', f'={U_TMP}&" max."',
            "Calculo: aplicabilidad del material en VIII-1, VIII-2 y XII de la ASME "
            "BPVC, segun las columnas impresas por el codigo. No aplica en tablas "
            "del B31.3."),
           ("Grafico de presion externa", val_of("Grafico presion externa"),
            '="Subparte 3 II-D"',
            "Calculo: referencia a la grafica de presion externa (Subparte 3 de la "
            "Seccion II-D) aplicable a este material.")]
    for k in range(max(len(izq), len(der))):
        r2 = 22 + k
        if k < len(izq):
            campo(ws, r2, 1, izq[k][0], izq[k][1], izq[k][2], com_val=izq[k][3])
        if k < len(der):
            campo(ws, r2, 7, der[k][0], der[k][1], der[k][2], com_val=der[k][3])

    # ------------------ 4 · TRAZABILIDAD DEL CALCULO ----------------------
    rt = 22 + max(len(izq), len(der)) + 1
    banda(ws, rt, "4 · TRAZABILIDAD DEL CALCULO   —   puntos tabulados usados en la "
                  "interpolacion")
    campo(ws, rt + 1, 1, "T1 — temperatura tabulada inferior", f"={T1C}", f"={U_TMP}",
         com_val="Calculo: temperatura tabulada inmediatamente inferior (o igual) a "
                "la temperatura de consulta, tomada de la banda compacta del "
                "material resuelto.")
    campo(ws, rt + 2, 1, "Valor en T1", f"={S1C}", f"={U_VAL}",
         com_val="Calculo: valor tabulado del codigo en T1, para el material "
                "resuelto.")
    campo(ws, rt + 1, 7, "T2 — temperatura tabulada superior", f"={T2C}", f"={U_TMP}",
         com_val="Calculo: temperatura tabulada inmediatamente superior a la "
                "temperatura de consulta. Vacia si T1 es el ultimo punto tabulado.")
    campo(ws, rt + 2, 7, "Valor en T2", f"={S2C}", f"={U_VAL}",
         com_val="Calculo: valor tabulado del codigo en T2, para el material "
                "resuelto. Vacio si T1 es el ultimo punto tabulado.")
    ec = _mrg(ws, rt + 3, 1, NCOLS)
    ec.value = (f'=IF({FIL}="","",IF($D$13="Tabulado-conservador",'
                f'"Modo tabulado-conservador: se adopta el valor de T2.",'
                f'IF(OR({T2C}="",{S2C}=""),'
                f'"Sin punto tabulado por encima de T: se adopta el ultimo valor tabulado '
                f'(el codigo prohibe extrapolar).",'
                f'"Interpolacion lineal:  S = S1 + (S2-S1)*(T-T1)/(T2-T1)")))')
    ec.font = SRC_F
    ec.alignment = Alignment(vertical="center", indent=1)
    _nota(ec, "Calculo: explica en texto cual de los tres casos aplico para obtener "
              "el KPI de la seccion 2 — interpolacion lineal, ultimo valor tabulado "
              "(sin extrapolar), o modo tabulado-conservador.")
    av = _mrg(ws, rt + 4, 1, NCOLS)
    av.value = ('="Verifique siempre la Temp. max. del material y sus notas antes de '
                'emitir el calculo."')
    av.font = SRC_F
    _nota(av, "Aviso fijo: recordatorio de verificar la Temp. max. del material "
              "(seccion 3) y sus notas del codigo antes de usar el resultado. No se "
              "edita.")

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
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.layout import Layout, ManualLayout
    from openpyxl.drawing.line import LineProperties
    ch = ScatterChart()
    ch.title = f"{ctx['valor_lbl']} frente a la temperatura"
    ch.style = 13
    ch.scatterStyle = "line"
    ch.x_axis.title = (f"Temperatura  [{ctx['temp_si']} en SI  ·  "
                       f"{ctx['temp_us']} en US]")
    ch.y_axis.title = (f"{ctx['valor_lbl']}  [{ctx['unit_si']} en SI  ·  "
                       f"{ctx['unit_us']} en US]")
    ch.height, ch.width = 9.5, 26
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    # Margen manual del area de trazado: sin esto, el titulo de eje rotado puede
    # quedar dibujado encima de las etiquetas numericas del eje (choque visual).
    # Va en ch.layout (no en ch.plot_area.layout): el escritor de openpyxl
    # sobreescribe plot_area.layout = ch.layout al guardar (ChartBase._write).
    ch.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.13, y=0.15, w=0.80, h=0.65))
    xs = Reference(curvas, min_col=c0, min_row=3, max_row=2 + npack)
    ys = Reference(curvas, min_col=c0 + 1, min_row=2, max_row=2 + npack)
    s1 = Series(ys, xs, title_from_data=True)
    # Linea continua, sin marcadores de punto: la curva del material se lee
    # como trazo, no como nube de puntos. Mismo azul que la banda del buscador
    # (BLUE), para que se identifique con la hoja en la que vive.
    s1.marker = Marker(symbol="none")
    # Linea recta (no suavizada) entre puntos: la interpolacion del codigo es
    # lineal (seccion 4), una curva suavizada la representaria mal.
    s1.smooth = False
    s1.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=BLUE, w=19050))
    ch.series.append(s1)
    xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
    yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
    s2 = Series(yq, xq, title_from_data=True)
    s2.marker = Marker(symbol="diamond", size=10,
                       spPr=GraphicalProperties(
                           solidFill="FF0000", ln=LineProperties(solidFill="FF0000")))
    s2.graphicalProperties.line = LineProperties(noFill=True)
    ch.series.append(s2)
    ws.add_chart(ch, chart_anchor)

    for cc, w in zip("ABCDEFGHIJKL",
                     [20, 20, 16, 16, 9, 3, 20, 20, 16, 16, 9, 3]):
        ws.column_dimensions[cc].width = w


def build_buscador_grupo(wb, curvas, bloques, rangos, name="Buscar_Prop_IID"):
    """Propiedades de la II-D indexadas por GRUPO de material (no por
    especificacion): modulo E (TM-1..5), dilatacion termica (TE-1..5,
    Coeficiente B) y Poisson/densidad (PRD).

    GRUPO es el rotulo normativo impreso en TM-1 / TE-1 («Material Group C»,
    «Group 3»), no la FAMILIA de db_lib, que es una agrupacion derivada solo
    para acortar listas desplegables y no interviene en ningun calculo.

    Conmutador SI/US (Regla 10): DB_E/DB_EC, DB_TE_G/DB_TE_GC y DB_PRD/DB_PRDC
    son UNA TABLA POR EDICION -no columnas de una tabla compartida-, asi que el
    conmutador cambia de HOJA, igual que en Buscar_Prop_B31_3. La lista de
    GRUPOS que existen por tabla se lee siempre de la edicion SI: el rotulo de
    grupo («Material Group C», «Group 3») no es un valor fisico y el codigo lo
    imprime identico en las dos ediciones (pertenencia verificada grupo a
    grupo, ver MAP_Grupo/MAP_GrupoC); lo que cambia con el selector es de que
    hoja se lee el VALOR para ese grupo.

    El Apendice C del B31.3 —que antes compartia esta hoja— tiene motor propio
    (Buscar_Prop_B31_3): se indexa por material, no por grupo, y necesita
    conmutador SI/US, rama de dato puntual y bloqueo por la propia banda
    tabulada. Un dato, un motor.
    """
    ws = new_sheet(wb, name,
                   "BUSCADOR DE PROPIEDADES POR GRUPO DE MATERIAL — ASME BPVC Seccion "
                   "II-D 2025: modulo E (TM-1..5), dilatacion termica (TE-1..5, "
                   "Coeficiente B) y Poisson/densidad (PRD)",
                   "Estas tablas del codigo se indexan por GRUPO de material, no por "
                   "especificacion: elija primero la tabla y despues el grupo. Para saber "
                   "que grupo corresponde a su material consulte MAP_Grupo (columnas "
                   "'Grupo E (TM)' y 'Grupo dilatacion (TE)'). El selector 'Sistema de "
                   "unidades' cambia la edicion del codigo leida (metrica o U.S. "
                   "Customary): los dos sistemas estan IMPRESOS, el conmutador nunca "
                   "convierte. La unica celda que se escribe es la temperatura de "
                   "consulta. Las propiedades fisicas del Apendice C del B31.3 estan en "
                   "Buscar_Prop_B31_3.")
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30
    SEL = '$D$5="SI"'
    banda(ws, 4, "PARAMETROS DE CONSULTA")
    com_sist = ("Entrada: SI lee DB_E / DB_TE_G / DB_PRD (edicion metrica); US lee "
                "DB_EC / DB_TE_GC / DB_PRDC (edicion U.S. Customary). Los dos sistemas "
                "estan IMPRESOS en el codigo: el conmutador nunca convierte.")
    e = _mrg(ws, 5, 1, 3, "Sistema de unidades")
    e.font = LBL_F
    e.alignment = Alignment(vertical="center", indent=1)
    _nota(e, com_sist)
    c = _mrg(ws, 5, 4, 6, "SI")
    c.font, c.fill, c.border = IN_F, IN_FILL, BOX
    dv_list(ws, "D5", '"SI,US"', com_sist)
    e = _mrg(ws, 6, 1, 3, "TEMPERATURA DE CONSULTA   (unica celda de escritura)")
    e.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    e.alignment = Alignment(vertical="center", indent=1)
    com_temp = ("Entrada: la UNICA celda de escritura libre de toda la hoja, en la "
               "unidad que muestra la celda de la derecha (segun el selector 'Sistema "
               "de unidades' de arriba). Se aplica por igual a los bloques de abajo "
               "(modulo E, dilatacion y Poisson/densidad).")
    _nota(e, com_temp)
    c = _mrg(ws, 6, 4, 5, 25)
    c.font, c.fill = IN_F, TEMP_INPUT_FILL
    c.border = Border(*[Side("medium", color="C00000")] * 4)
    _nota(c, com_temp)
    u_tmp = ws.cell(6, 6)
    u_tmp.value = f'=IF({SEL},"°C","°F")'
    u_tmp.font = UNIT_F
    _nota(u_tmp, "Calculo: unidad de la temperatura de consulta; cambia entre °C y °F "
                 "segun el selector 'Sistema de unidades' (D5).")
    e = _mrg(ws, 7, 1, 3, "Modo de lectura")
    e.font = LBL_F
    e.alignment = Alignment(vertical="center", indent=1)
    com_modo = ("Entrada: 'Interpolado' aplica la interpolacion lineal del codigo "
               "entre los dos puntos tabulados que rodean la temperatura de "
               "consulta. 'Tabulado-conservador' adopta directamente el valor "
               "tabulado superior (T2), sin interpolar.")
    _nota(e, com_modo)
    c = _mrg(ws, 7, 4, 6, "Interpolado")
    c.font, c.fill, c.border = IN_F, IN_FILL, BOX
    dv_list(ws, "D7", '"Interpolado,Tabulado-conservador"', com_modo)

    r = 9
    for i2, b in enumerate(bloques):
        info_si, info_us = b["info_si"], b["info_us"]
        banda(ws, r, b["titulo"])
        r += 1
        e = _mrg(ws, r, 1, 3, "Tabla / familia del codigo")
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        com_tabla = (f'Entrada: elija la tabla/familia del codigo dentro de '
                    f'"{b["titulo"]}". Habilita la lista de grupos de esa tabla.')
        _nota(e, com_tabla)
        c = _mrg(ws, r, 4, 6, b["default_tabla"])
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        dv_list(ws, f"D{r}", "=" + rangos[b["lst_tabla"]], com_tabla)
        tcell = f"$D${r}"
        r += 1
        e = _mrg(ws, r, 1, 3, "Grupo / material")
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        com_grupo = ("Entrada: elija el GRUPO de material tal como lo imprime el "
                    "codigo (p. ej. «Material Group C», «Group 3»), no la familia de "
                    "navegacion. Para saber que grupo corresponde a un material "
                    "especifico, consulte MAP_Grupo. El rotulo de grupo es identico "
                    "en las dos ediciones del codigo: esta lista no cambia con el "
                    "selector 'Sistema de unidades'.")
        _nota(e, com_grupo)
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
        dv_list(ws, f"D{r}", f"=${LL}${R_DATA}:${LL}${R_DATA + mx - 1}", com_grupo)
        gcell = f"$D${r}"
        r += 1
        idn_si = f"{info_si['sheet']}!$D${R_DATA}:$D${info_si['last_row']}"
        idn_us = f"{info_us['sheet']}!$D${R_DATA}:$D${info_us['last_row']}"
        ws.cell(r, 60).value = (f'=IF({SEL},IFERROR(MATCH({gcell},{idn_si},0),""),'
                                f'IFERROR(MATCH({gcell},{idn_us},0),""))')
        frow = f"$BH${r}"
        ws.column_dimensions["BH"].hidden = True
        if b.get("temps"):
            u_val = f'IF({SEL},"{b["unidad_si"]}","{b["unidad_us"]}")'
            u_tmp_f = f'IF({SEL},"°C","°F")'
            rm, ru = packed_refs(info_si), packed_refs(info_us)
            ws.cell(r, 61).value = (f'=IF({frow}="","",IF({SEL},'
                                    f'IFERROR(INDEX({rm["npts"]},{frow}),0),'
                                    f'IFERROR(INDEX({ru["npts"]},{frow}),0)))')
            NP = f"$BI${r}"
            ws.column_dimensions["BI"].hidden = True
            T_ANC = f'IF({SEL},{rm["t_anchor"]},{ru["t_anchor"]})'
            V_ANC = f'IF({SEL},{rm["v_anchor"]},{ru["v_anchor"]})'
            tr = f'OFFSET({T_ANC},{frow}-1,0,1,MAX(1,{NP}))'
            vr = f'OFFSET({V_ANC},{frow}-1,0,1,MAX(1,{NP}))'
            P1 = f'IFERROR(MATCH($D$6,{tr},1),1)'
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
                       interp_value(T1C, S1C, T2C, S2C, "$D$6", "$D$7")[1:] + ')')
            v.font, v.fill = KPI_VAL_F, KPI_FILL
            v.alignment = Alignment(horizontal="center", vertical="center")
            _nota(v, f'Calculo: {b["valor_lbl"]} interpolado (o tabulado-conservador, '
                     'segun el Modo de lectura de D7) a la temperatura de consulta '
                     '(D6), entre los puntos tabulados T1/T2 de abajo, para el grupo '
                     'elegido arriba y la edicion elegida en D5.')
            u = _mrg(ws, r, 6, 8)
            u.value = f"={u_val}"
            u.font, u.fill = UNIT_F, KPI_FILL
            ws.row_dimensions[r].height = 26
            valr = r
            r += 1
            campo(ws, r, 1, "T1 — tabulada inferior", f"={T1C}", f"={u_tmp_f}",
                 com_val="Calculo: temperatura tabulada inmediatamente inferior (o "
                        "igual) a la temperatura de consulta, para el grupo elegido.")
            campo(ws, r, 7, "Valor en T1", f"={S1C}", f"={u_val}",
                 com_val=f'Calculo: {b["valor_lbl"]} tabulado en T1 para el grupo '
                        'elegido.')
            r += 1
            campo(ws, r, 1, "T2 — tabulada superior", f"={T2C}", f"={u_tmp_f}",
                 com_val="Calculo: temperatura tabulada inmediatamente superior a la "
                        "de consulta. Vacia si T1 es el ultimo punto tabulado.")
            campo(ws, r, 7, "Valor en T2", f"={S2C}", f"={u_val}",
                 com_val=f'Calculo: {b["valor_lbl"]} tabulado en T2 para el grupo '
                        'elegido. Vacio si T1 es el ultimo punto tabulado.')
            r += 1
            n2 = _mrg(ws, r, 1, NCOLS)
            n2.value = f'=IF({SEL},"{b.get("nota_si", "")}","{b.get("nota_us", "")}")'
            n2.font = SRC_F
            r += 1
            c0 = 200 + i2 * 6
            npack = max(info_si["npack"], info_us["npack"])
            q = f"'{ws.title}'!"
            trq = tr.replace(frow, q + frow).replace(NP, q + NP)
            vrq = vr.replace(frow, q + frow).replace(NP, q + NP)
            curvas.cell(1, c0, f"{b['titulo'][:38]} — curva").font = SRC_F
            curvas.cell(2, c0, "Temperatura").font = HDR_F
            curvas.cell(2, c0 + 1, f'{b["valor_lbl"]}').font = HDR_F
            for k in range(1, npack + 1):
                curvas.cell(2 + k, c0).value = f'=IFERROR(INDEX({trq},{k}),NA())'
                curvas.cell(2 + k, c0 + 1).value = f'=IFERROR(INDEX({vrq},{k}),NA())'
            curvas.cell(2, c0 + 3, "T consulta").font = HDR_F
            curvas.cell(2, c0 + 4, "Punto consultado").font = HDR_F
            curvas.cell(3, c0 + 3).value = f"={q}$D$6"
            curvas.cell(3, c0 + 4).value = f"={q}$D${valr}"

            from openpyxl.chart import Reference, Series, ScatterChart
            from openpyxl.chart.marker import Marker
            from openpyxl.chart.shapes import GraphicalProperties
            from openpyxl.chart.layout import Layout, ManualLayout
            from openpyxl.drawing.line import LineProperties
            ch = ScatterChart()
            ch.title = f'{b["valor_lbl"]} frente a la temperatura'
            ch.scatterStyle = "line"
            ch.x_axis.title = "Temperatura  [°C en SI  ·  °F en US]"
            ch.y_axis.title = (f'{b["valor_lbl"]}  [{b["unidad_si"]} en SI  ·  '
                               f'{b["unidad_us"]} en US]')
            ch.height, ch.width = 7.5, 20
            ch.x_axis.delete = False
            ch.y_axis.delete = False
            ch.layout = Layout(manualLayout=ManualLayout(
                xMode="edge", yMode="edge", x=0.15, y=0.16, w=0.78, h=0.62))
            xs = Reference(curvas, min_col=c0, min_row=3, max_row=2 + npack)
            ys = Reference(curvas, min_col=c0 + 1, min_row=2, max_row=2 + npack)
            se = Series(ys, xs, title_from_data=True)
            # Linea continua sin marcadores, mismo azul de la banda del buscador
            # (BLUE) — mismo estilo que finish_buscador.
            se.marker = Marker(symbol="none")
            se.smooth = False
            se.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=BLUE, w=19050))
            ch.series.append(se)
            xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
            yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
            sq = Series(yq, xq, title_from_data=True)
            sq.marker = Marker(symbol="diamond", size=10,
                               spPr=GraphicalProperties(
                                   solidFill="FF0000",
                                   ln=LineProperties(solidFill="FF0000")))
            sq.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
            ch.series.append(sq)
            ws.add_chart(ch, f"A{r}")
            r += 15
        else:
            for lbl, col, uni_si, uni_us in b["campos"]:
                Lc = get_column_letter(col)
                val_si = f"{info_si['sheet']}!${Lc}${R_DATA}:${Lc}${info_si['last_row']}"
                val_us = f"{info_us['sheet']}!${Lc}${R_DATA}:${Lc}${info_us['last_row']}"
                campo(ws, r, 1, lbl,
                      f'=IF({frow}="","—",IF({SEL},INDEX({val_si},{frow}),'
                      f'INDEX({val_us},{frow})))',
                      f'=IF({SEL},"{uni_si}","{uni_us}")',
                      com_val=f"Calculo: {lbl} impreso por la Tabla PRD para el grupo "
                              "elegido arriba y la edicion elegida en 'Sistema de "
                              "unidades' (D5). No depende de la temperatura de "
                              "consulta (la Tabla PRD no tabula frente a T).")
                r += 1
            r += 1
    for cc, w in zip("ABCDEFGHIJKL", [22, 22, 18, 16, 12, 12, 22, 22, 16, 12, 12, 4]):
        ws.column_dimensions[cc].width = w
    return ws


# ---------------------------------------------------------------------------
# Buscar_Prop_B31_3 — propiedades fisicas del Apendice C del B31.3
# ---------------------------------------------------------------------------
# No reutiliza build_buscador/finish_buscador a proposito: esa pareja la
# comparten los cinco buscadores de esfuerzos y sostiene 271 276 valores ya
# auditados. El Apendice C necesita tres cosas que ella no tiene —rama de dato
# puntual (C-2 y C-4 no dependen de T), limite de rango tomado de la propia
# banda tabulada en vez de una columna «Temp. max.», y un conmutador que en dos
# tablas cambia de COLUMNA en lugar de hoja—. Tocarla arriesgaria la regresion
# de todo lo verificado.
#
# Columnas ocultas propias: 36..39 (listas materializadas de los niveles 1 a 4)
# y 41..42 (auxiliares de resolucion). Quedan por debajo de AUX_COL (46), de
# CASC_COL (50..54), del bloque 60..65 de build_buscador_grupo y, sobre todo,
# de COL_CLAVE_BASE (66).
APXC_LST_COL = 36          # 36, 37, 38, 39 -> niveles 1, 2, 3 y 4
APXC_AUX_COL = 41          # 41 = rotulo, 42 = valor
APXC_CURVA_COL = 400       # columna de arranque en _Curvas


def _rngc(info, colname):
    """Rango de una columna de la base del Apendice C (layout APXC_COLS)."""
    L = CLC[colname]
    return f"{info['sheet']}!${L}${R_DATA}:${L}${info['last_row']}"


def formula_estado_apxc(fil, tipo, npts, tprim, tult, tq):
    """Estado del rango del Apendice C. Bloquea en LOS DOS extremos.

    El Apendice C no publica columna «Temp. max.»: el limite es el primer y el
    ultimo punto que la propia fila tabula. Sostener el ultimo valor por encima
    del rango —que es lo que hacen los cinco buscadores de esfuerzos, donde el
    tope lo pone una columna del codigo— seria extrapolar aqui.

    La expresion se genera en una sola funcion porque verificar.py §6 la vuelve
    a emitir para recalcularla en Excel: si viviese suelta dentro del motor, la
    prueba estaria comprobando una copia y no el original.
    """
    return (f'=IF({fil}="","{EST_SIN_SEL}",'
            f'IF({tipo}="{TIPO_PUNTO}","{EST_PUNTO}",'
            f'IF({npts}=0,"{EST_SIN_TAB}",'
            f'IF({tq}<{tprim},"{EST_BAJO}",'
            f'IF({tq}>{tult},"{EST_ALTO}",'
            f'"{EST_OK}")))))')


def formula_valor_apxc(fil, tipo, vu, vt, fac, est, vtab):
    """Valor de la propiedad CON el factor de escala aplicado (regla 6: la
    unidad final va visible junto al KPI; la seccion 4 muestra el impreso)."""
    return (f'=IF({fil}="","",'
            f'IF({tipo}="{TIPO_PUNTO}",IF(ISNUMBER({vu}),{vu}*{fac},{vt}),'
            f'IF(ISNUMBER(SEARCH("FUERA DE RANGO",{est})),"{VAL_BLOQUEADO}",'
            f'IF({est}="{EST_SIN_TAB}","{EST_SIN_TAB}",'
            f'IF(ISNUMBER({vtab}),{vtab}*{fac},"")))))')


def build_buscador_prop_c(wb, curvas, rng, master, us):
    name = "Buscar_Prop_B31_3"
    hl = get_column_letter
    ws = new_sheet(
        wb, name,
        "PROPIEDADES FISICAS DE MATERIALES DE TUBERIA — ASME B31.3-2024, APENDICE C",
        "Tablas C-1/C-1C y C-3/C-3C (metales) · C-2 y C-4 (no metalicos). Se elige "
        "primero QUE PROPIEDAD se necesita y despues el material. Valores tal como "
        "estan impresos: el factor de escala se aplica a la vista y se declara. "
        "La unica celda que se escribe es la temperatura de consulta.")
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30

    SEL = '$D$5="SI"'

    def ident(colname):
        return f'IF({SEL},{_rngc(master, colname)},{_rngc(us, colname)})'

    banda(ws, 4, "1 · QUE PROPIEDAD QUIERE CONSULTAR   —   todo por lista desplegable")
    com_sist = ("Entrada: SI lee la edicion metrica del Apendice C (Tablas C-1, C-3 y "
                "las columnas metricas de C-2 y C-4); US lee la edicion U.S. Customary "
                "(C-1C, C-3C y las columnas en pulgadas de C-2 y C-4). Los dos sistemas "
                "estan IMPRESOS en el codigo: el conmutador nunca convierte.")
    filas = [
        (5, "Sistema de unidades", "SI", com_sist),
        (6, "0 · Propiedad", "",
         "Entrada: paso 0 de la cascada, y el nivel que distingue este motor: se elige "
         "primero QUE propiedad se necesita (dilatacion o modulo, metal o no metal) y "
         "el motor decide en que tabla del Apendice C vive."),
        (7, "1 · Grupo del codigo", "",
         "Entrada: paso 1, dependiente del paso 0. Encabezado de grupo tal como lo "
         "imprime la tabla. Las tablas que no imprimen grupo ofrecen "
         f"«{SIN_GRUPO}»."),
        (8, "2 · Subgrupo", "",
         "Entrada: paso 2, dependiente de los pasos 0 y 1. Subgrupo impreso "
         f"(va indentado en el codigo). Si la tabla no lo imprime, «{SIN_SUBGRUPO}»."),
        (9, "3 · Material", "",
         "Entrada: paso 3, dependiente de los pasos 0 a 2. Material tal como lo nombra "
         "la tabla del codigo."),
        (10, "4 · Coeficiente / variante", "",
         "Entrada: paso 4, dependiente de los pasos 0 a 3. En C-1 elige entre el "
         "coeficiente A (medio) y el B (expansion total acumulada); en las tres filas de "
         "Poly(perfluoroalkoxy alkane) de C-2 elige el rango de validez, que es lo que "
         f"las separa. En el resto de las filas vale «{VAR_UNICA}»."),
    ]
    for r, et, val, com in filas:
        e = _mrg(ws, r, 1, 3, et)
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        _nota(e, com)
        c = _mrg(ws, r, 4, 6, val)
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        ws.row_dimensions[r].height = 17

    # Celda unica de escritura: relleno lavanda propio, distinto del amarillo de
    # las listas desplegables (regla de estilo 1 del proyecto).
    com_temp = ("Entrada: la UNICA celda de escritura libre de todo el motor. "
                "Temperatura a la que se necesita la propiedad, en la unidad de la "
                "celda de la derecha. Las propiedades de tipo PUNTO (C-2 y C-4) no "
                "dependen de la temperatura y el ESTADO lo avisa.")
    e = _mrg(ws, 11, 1, 3, "TEMPERATURA DE CONSULTA   (unica celda de escritura)")
    e.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    e.alignment = Alignment(vertical="center", indent=1)
    _nota(e, com_temp)
    c = _mrg(ws, 11, 4, 5, 350)
    c.font, c.fill = IN_F, TEMP_INPUT_FILL
    c.border = Border(*[Side("medium", color="C00000")] * 4)
    _nota(c, com_temp)
    u = ws.cell(11, 6)
    u.value = f'=IF({SEL},"°C","°F")'
    u.font = UNIT_F
    _nota(u, "Calculo: unidad de la temperatura de consulta; cambia entre °C y °F "
             "segun el selector 'Sistema de unidades' (D5).")

    com_modo = ("Entrada: 'Interpolado' aplica la interpolacion lineal entre los dos "
                "puntos tabulados que rodean la temperatura de consulta. "
                "'Tabulado-conservador' adopta el valor tabulado superior (T2). "
                "No interviene en las propiedades de tipo PUNTO.")
    e = _mrg(ws, 12, 1, 3, "Modo de lectura")
    e.font = LBL_F
    e.alignment = Alignment(vertical="center", indent=1)
    _nota(e, com_modo)
    c = _mrg(ws, 12, 4, 6, "Interpolado")
    c.font, c.fill, c.border = IN_F, IN_FILL, BOX

    dv_list(ws, "D5", '"SI,US"', com_sist)
    dv_list(ws, "D12", '"Interpolado,Tabulado-conservador"', com_modo)
    dv_list(ws, "D6", "=" + rng["FAM"], filas[1][3])
    niveles = [("D7", rng["CK"], rng["CV"], "$D$6", rng.get("maxC", 20), filas[2][3]),
               ("D8", rng["FK"], rng["FV"], '$D$6&"|"&$D$7', rng.get("maxF", 20),
                filas[3][3]),
               ("D9", rng["SK"], rng["SV"], '$D$6&"|"&$D$7&"|"&$D$8',
                rng.get("maxS", 80), filas[4][3]),
               ("D10", rng["GK"], rng["GV"], '$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9',
                rng.get("maxG", 10), filas[5][3])]
    for i, (cell, kr, vr, key, mx, com) in enumerate(niveles):
        col = APXC_LST_COL + i
        L = hl(col)
        mx = max(1, min(int(mx), 250))
        ws.cell(R_HDR, col, f"lista nivel {i + 1}").font = SRC_F
        for k in range(1, mx + 1):
            ws.cell(R_DATA + k - 1, col).value = (
                f'=IF(COUNTIF({kr},{key})<{k},"",'
                f'INDEX({vr},MATCH({key},{kr},0)+{k}-1))')
        ws.column_dimensions[L].hidden = True
        dv_list(ws, cell, f"=${L}${R_DATA}:${L}${R_DATA + mx - 1}", com)

    # --- auxiliares de resolucion (columnas ocultas) ------------------------
    A = APXC_AUX_COL
    LA = hl(A + 1)
    rm, ru = packed_refs(master), packed_refs(us)
    npack = max(master["npack"], us["npack"])

    def aux(fila, rotulo, formula):
        ws.cell(fila, A, rotulo).font = SRC_F
        ws.cell(fila, A + 1).value = formula
        return f"${LA}${fila}"

    KEY = aux(4, "clave k4", '=$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9&"|"&$D$10')
    FSI = aux(5, "fila SI", f'=IFERROR(MATCH({KEY},{rng["K4"]},0),"")')
    MIDC = aux(6, "material_id", f'=IF({FSI}="","",INDEX({rng["ID"]},{FSI}))')
    BI = aux(7, "clave bilingue",
             f'=IF({FSI}="","",INDEX({_rngc(master,"clave_bi")},{FSI}))')
    FUS = aux(8, "fila US",
              f'=IF({BI}="","",IFERROR(MATCH({BI},{_rngc(us,"clave_bi")},0),""))')
    FIL = aux(9, "fila activa", f'=IF({SEL},{FSI},{FUS})')
    NP = aux(10, "n_pts", f'=IF({FIL}="",0,IFERROR(IF({SEL},'
                          f'INDEX({rm["npts"]},{FIL}),INDEX({ru["npts"]},{FIL})),0))')
    T_ANC = f'IF({SEL},{rm["t_anchor"]},{ru["t_anchor"]})'
    V_ANC = f'IF({SEL},{rm["v_anchor"]},{ru["v_anchor"]})'
    TR = f'OFFSET({T_ANC},{FIL}-1,0,1,MAX(1,{NP}))'
    VR = f'OFFSET({V_ANC},{FIL}-1,0,1,MAX(1,{NP}))'
    # INDEX sobre una celda VACIA devuelve 0, no cadena vacia. Sin este
    # envoltorio, una fila de tipo PUNTO —que no tiene banda tabulada— daria
    # T1 = 0 y ISNUMBER(valor unico) = VERDADERO sobre un 0 inventado, y el
    # motor mostraria 0 donde el codigo publica un intervalo de texto.
    def vacio_si_vacio(expr):
        return f'IF({expr}="","",{expr})'

    def indexc(colname):
        """INDEX de una columna de la base sobre la fila activa, vacio si vacia."""
        return vacio_si_vacio(f'INDEX({ident(colname)},{FIL})')

    P1 = aux(11, "p1", f'=IF({FIL}="","",IFERROR(MATCH($D$11,{TR},1),1))')
    tabulados = [("T1", f"INDEX({TR},{P1})"), ("V1", f"INDEX({VR},{P1})"),
                 ("T2", f"INDEX({TR},{P1}+1)"), ("V2", f"INDEX({VR},{P1}+1)")]
    T1C, S1C, T2C, S2C = [
        aux(12 + k, lb, f'=IF({FIL}="","",IFERROR({vacio_si_vacio(expr)},""))')
        for k, (lb, expr) in enumerate(tabulados)]
    TIPO = aux(16, "tipo de dato",
               f'=IF({FIL}="","",INDEX({ident("Tipo de dato")},{FIL}))')
    TPRI = aux(17, "T primera", f'=IF({FIL}="","",{indexc("T primera tabulada")})')
    TULT = aux(18, "T ultima", f'=IF({FIL}="","",{indexc("T ultima tabulada")})')
    FAC = aux(19, "factor",
              f'=IF({FIL}="",1,INDEX({ident("Factor de escala")},{FIL}))')
    VU = aux(20, "valor unico", f'=IF({FIL}="","",{indexc("Valor unico")})')
    VT = aux(21, "valor unico texto",
             f'=IF({FIL}="","",{indexc("Valor unico (texto)")})')
    VTAB = aux(22, "valor tabulado",
               f'=IF({FIL}="","",' + interp_value(T1C, S1C, T2C, S2C,
                                                  "$D$11", "$D$12")[1:] + ')')
    # Bloqueo en LOS DOS extremos. El Apendice C no publica «Temp. max.»: el
    # limite es el primer y el ultimo punto tabulado de la propia fila. Sostener
    # el ultimo valor por encima del rango seria extrapolar, que es justo lo que
    # prohibe el codigo (regla 4 del proyecto).
    EST = aux(23, "estado", formula_estado_apxc(FIL, TIPO, NP, TPRI, TULT, "$D$11"))
    VALOR = aux(24, "valor aplicado",
                formula_valor_apxc(FIL, TIPO, VU, VT, FAC, EST, VTAB))
    for jj in (A, A + 1):
        ws.column_dimensions[hl(jj)].hidden = True

    # Aviso: que tabla y que edicion se esta leyendo.
    ay = _mrg(ws, 5, 7, NCOLS)
    ay.value = (f'=IF({FIL}="",IF({SEL},'
                f'"Edicion metrica — {master["sheet"]}",'
                f'"Edicion U.S. Customary — {us["sheet"]}"),'
                f'"Leyendo Table "&INDEX({ident("Tabla")},{FIL})&'
                f'" ("&$D$5&") — valores en "&INDEX({ident("Unidad impresa")},{FIL}))')
    ay.font = Font(italic=True, color=BLUE)
    ay.alignment = Alignment(vertical="center", wrap_text=True)
    _nota(ay, "Aviso automatico: dice siempre que tabla del Apendice C y que edicion "
              "esta leyendo el motor. No se edita.")

    # Semaforo de cascada. El disparador es $D$9 (nivel 3 · Material) porque el
    # nivel 4 solo tiene contenido real en C-1 y en tres filas de C-2.
    ay2 = _mrg(ws, 11, 7, 9)
    ay2.value = '=IF($D$9="","SELECCION INCOMPLETA","SELECCION COMPLETA")'
    ay2.font, ay2.fill, ay2.border = SEL_BAD_FONT, SEL_BAD_FILL, BOX
    ay2.alignment = Alignment(horizontal="center", vertical="center")
    _nota(ay2, "Aviso automatico: SELECCION COMPLETA (verde) cuando la cascada llega "
               "al material; SELECCION INCOMPLETA (amarillo) mientras falte un paso.")
    ws.conditional_formatting.add(
        "G11:I11", FormulaRule(formula=['$D$9<>""'], fill=SEL_OK_FILL, font=SEL_OK_FONT))
    ws.conditional_formatting.add(
        "G11:I11", FormulaRule(formula=['$D$9=""'], fill=SEL_BAD_FILL, font=SEL_BAD_FONT))

    # ------------------------- 2 · RESULTADO -------------------------------
    banda(ws, 14, "2 · RESULTADO DE LA CONSULTA")
    kpis = [(1, "VALOR DE LA PROPIEDAD", "=" + VALOR,
             f'=IF({FIL}="","",INDEX({ident("Unidad impresa")},{FIL}))',
             "Calculo: valor de la propiedad CON el factor de escala ya aplicado, a la "
             "temperatura de consulta. En las propiedades de tipo PUNTO es el valor "
             "unico impreso. BLOQUEADO si la temperatura cae fuera de la banda "
             "tabulada: el codigo prohibe extrapolar."),
            (4, "TEMPERATURA DE CONSULTA", "=$D$11", f'=IF({SEL},"°C","°F")',
             "Calculo: repite la temperatura tecleada en D11, junto al resultado."),
            (7, "MODO DE LECTURA", "=$D$12", '="segun MODO_S"',
             "Calculo: repite el modo de lectura elegido en D12. No interviene en las "
             "propiedades de tipo PUNTO."),
            (10, "ESTADO DEL RANGO", "=" + EST,
             f'=IF({FIL}="","",INDEX({ident("Tabla")},{FIL})&'
             f'IF(INDEX({ident("Notas")},{FIL})="",""," · Nota "&'
             f'INDEX({ident("Notas")},{FIL})))',
             "Calculo: EN RANGO si T cae entre el primer y el ultimo punto tabulado de "
             "la fila; FUERA DE RANGO en cualquiera de los dos extremos (el Apendice C "
             "no publica Temp. max.: el limite es la propia banda tabulada); VALOR "
             "UNICO en C-2 y C-4, que no dependen de la temperatura.")]
    for c1, tit, val, uni, com in kpis:
        t = _mrg(ws, 15, c1, c1 + 2, tit)
        t.font, t.fill = KPI_TIT_F, BAND_FILL
        t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        v = _mrg(ws, 16, c1, c1 + 2, val)
        v.font, v.fill = KPI_VAL_F, KPI_FILL
        v.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        _nota(v, com)
        u2 = _mrg(ws, 17, c1, c1 + 2, uni)
        u2.font, u2.fill = UNIT_F, KPI_FILL
        u2.alignment = Alignment(horizontal="center", wrap_text=True)
        for r2 in (15, 16, 17):
            for cc in range(c1, c1 + 3):
                ws.cell(r2, cc).border = CARD_BORDER
    ws.row_dimensions[15].height = 26
    ws.row_dimensions[16].height = 30
    ws.row_dimensions[17].height = 15
    ws.conditional_formatting.add(
        "A16:L17",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("FUERA DE RANGO",{EST}))'],
                    font=Font(color="9C0006", bold=True)))

    # ---------------------- 3 · FICHA DE LA PROPIEDAD ----------------------
    banda(ws, 19, "3 · FICHA DE LA PROPIEDAD   —   cada campo con su unidad impresa")

    def val_of(colname):
        idx = f'INDEX({ident(colname)},{FIL})'
        return f'=IF({FIL}="","—",IF({idx}="","—",{idx}))'

    izq = [("Tabla del codigo", val_of("Tabla"), None,
            "Calculo: tabla del Apendice C (C-1/C-1C, C-2, C-3/C-3C o C-4) de la que "
            "sale la fila resuelta por la cascada."),
           ("Propiedad consultada", val_of("Propiedad"), None,
            "Calculo: propiedad elegida en el paso 0 de la cascada."),
           ("Grupo impreso", val_of("Grupo impreso"), None,
            "Calculo: encabezado de grupo tal como lo imprime la tabla."),
           ("Subgrupo impreso", val_of("Subgrupo impreso"), None,
            "Calculo: subgrupo impreso (va indentado en el codigo)."),
           ("Material", val_of("Material"), None,
            "Calculo: material tal como lo nombra la tabla del codigo."),
           ("Variante / coeficiente", val_of("Variante / coeficiente"), None,
            "Calculo: coeficiente A o B en C-1, rango de validez en las filas que el "
            "codigo repite, «(unico)» en el resto."),
           ("Definicion impresa", val_of("Definicion impresa"), None,
            "Calculo: definicion del coeficiente tal como la imprime el codigo, con su "
            "unidad (folios 408 y 414 del B31.3-2024)."),
           ("Tipo de dato", val_of("Tipo de dato"), None,
            "Calculo: CURVA si el codigo tabula la propiedad frente a la temperatura; "
            "PUNTO si publica un valor unico (C-2 y C-4)."),
           ("Rango de validez", val_of("Rango de validez"), None,
            "Calculo: rango de temperatura de validez impreso en C-2, o temperatura de "
            "referencia impresa en C-4 (23 °C / 73,4 °F). Vacio en las tablas de curva.")]
    der = [("Edicion consultada",
            f'=IF({FIL}="","—",IF({SEL},"{master["sheet"]} — metrica",'
            f'IF({FUS}="","sin equivalente en la edicion US",'
            f'"{us["sheet"]} — U.S. Customary")))', None,
            "Calculo: confirma de que edicion sale la fila activa y avisa si el "
            "material no tiene homologo publicado en la otra."),
           ("Unidad impresa", val_of("Unidad impresa"), None,
            "Calculo: unidad en la que queda el valor DESPUES de aplicar el factor de "
            "escala, tal como la imprime el codigo."),
           ("Factor de escala aplicado", val_of("Factor de escala"),
            '="multiplicador"',
            "Calculo: multiplicador que lleva del valor tabulado al valor fisico. Se "
            "lee del encabezado de la tabla, no esta codificado en el motor."),
           ("Factor tal como se imprime", val_of("Factor impreso"), None,
            "Calculo: el texto del encabezado del que sale el factor "
            "(«Multiply Tabulated Values by 10^3», «Divide Table Values by 10^6»)."),
           ("Primer punto tabulado", val_of("T primera tabulada"),
            f'=IF({SEL},"°C","°F")',
            "Calculo: temperatura del primer punto que el codigo tabula para esta fila. "
            "Por debajo de ella el resultado queda BLOQUEADO."),
           ("Ultimo punto tabulado", val_of("T ultima tabulada"),
            f'=IF({SEL},"°C","°F")',
            "Calculo: temperatura del ultimo punto tabulado. Por encima de ella el "
            "resultado queda BLOQUEADO: el codigo prohibe extrapolar."),
           ("Notas del codigo", val_of("Notas"), '="ver Instrucciones"',
            "Calculo: notas al pie que cita esta fila o su tabla."),
           ("Observacion de extraccion", val_of("Observacion"), None,
            "Calculo: artefacto de extraccion conservado o reparado en esta fila, con "
            "el folio impreso que lo respalda. Vacio si no hay ninguno.")]
    for k in range(max(len(izq), len(der))):
        r2 = 20 + k
        if k < len(izq):
            campo(ws, r2, 1, izq[k][0], izq[k][1], izq[k][2], com_val=izq[k][3])
        if k < len(der):
            campo(ws, r2, 7, der[k][0], der[k][1], der[k][2], com_val=der[k][3])

    # ------------------ 4 · TRAZABILIDAD DEL CALCULO -----------------------
    rt = 20 + max(len(izq), len(der)) + 1
    banda(ws, rt, "4 · TRAZABILIDAD   —   valores tal como estan impresos, antes del "
                  "factor")
    campo(ws, rt + 1, 1, "T1 — tabulada inferior", f"={T1C}", f'=IF({SEL},"°C","°F")',
          com_val="Calculo: temperatura tabulada inmediatamente inferior (o igual) a la "
                  "de consulta, tomada de la banda compacta de la fila.")
    campo(ws, rt + 2, 1, "Valor impreso en T1", f"={S1C}",
          f'=IF({FIL}="","",INDEX({ident("Unidad impresa")},{FIL})&" / factor")',
          com_val="Calculo: valor TABULADO en T1, sin aplicar el factor de escala. El "
                  "KPI de la seccion 2 muestra este valor multiplicado por el factor.")
    campo(ws, rt + 1, 7, "T2 — tabulada superior", f"={T2C}", f'=IF({SEL},"°C","°F")',
          com_val="Calculo: temperatura tabulada inmediatamente superior. Vacia si T1 "
                  "es el ultimo punto tabulado.")
    campo(ws, rt + 2, 7, "Valor impreso en T2", f"={S2C}",
          f'=IF({FIL}="","",INDEX({ident("Unidad impresa")},{FIL})&" / factor")',
          com_val="Calculo: valor TABULADO en T2, sin aplicar el factor de escala.")
    ec = _mrg(ws, rt + 3, 1, NCOLS)
    ec.value = (
        f'=IF({FIL}="","",'
        f'IF({TIPO}="{TIPO_PUNTO}",'
        f'"Valor unico impreso: no hay interpolacion. Se muestra el valor de la columna '
        f'del sistema activo, multiplicado por el factor de escala.",'
        f'IF(ISNUMBER(SEARCH("FUERA DE RANGO",{EST})),'
        f'"Resultado bloqueado: la temperatura cae fuera de la banda tabulada y el '
        f'codigo prohibe extrapolar.",'
        f'IF($D$12="Tabulado-conservador",'
        f'"Modo tabulado-conservador: se adopta el valor de T2, por el factor de escala.",'
        f'"Interpolacion lineal:  V = V1 + (V2-V1)*(T-T1)/(T2-T1), por el factor de '
        f'escala."))))')
    ec.font = SRC_F
    ec.alignment = Alignment(vertical="center", indent=1)
    _nota(ec, "Calculo: dice en texto cual de los cuatro casos aplico para obtener el "
              "KPI de la seccion 2.")
    fu = _mrg(ws, rt + 4, 1, NCOLS)
    fu.value = (f'=IF({FIL}="",'
                f'"Fuente: resources/{APX}/appendix_c/",'
                f'"Fuente: resources/{APX}/appendix_c/ · Table "&'
                f'INDEX({ident("Tabla")},{FIL})&" · fila impresa "&'
                f'INDEX({ident("Linea")},{FIL}))')
    fu.font = SRC_F
    fu.alignment = Alignment(vertical="center", indent=1)
    _nota(fu, "Trazabilidad: archivo de resources/ y numero de fila impresa de la que "
              "sale el valor mostrado.")

    # --------------------------- 5 · CURVA ---------------------------------
    rg = rt + 6
    banda(ws, rg, "5 · CURVA DE LA PROPIEDAD   —   valor tabulado frente a la temperatura")
    av = _mrg(ws, rg + 1, 1, NCOLS)
    av.value = (f'=IF({TIPO}="{TIPO_PUNTO}",'
                f'"Esta propiedad no depende de la temperatura: el codigo publica un '
                f'valor unico. La grafica queda vacia a proposito.","")')
    av.font = Font(name="Calibri", size=9, italic=True, color="9C6500")
    av.alignment = Alignment(vertical="center", indent=1)
    _nota(av, "Aviso automatico: explica por que la grafica esta vacia cuando la "
              "propiedad es de tipo PUNTO (C-2 y C-4).")

    c0 = APXC_CURVA_COL
    q = f"'{ws.title}'!"

    def _q(expr):
        return (expr.replace("$D$5", q + "$D$5")
                    .replace(NP, q + NP).replace(FIL, q + FIL))

    curvas.cell(1, c0, f"{ws.title} — datos de la curva (hoja auxiliar)").font = SRC_F
    curvas.cell(2, c0).value = f'=CONCATENATE("Temperatura, ",{q}$F$11)'
    curvas.cell(2, c0 + 1).value = (
        f'=CONCATENATE("Valor tabulado, ",IF({q}{FIL}="","",'
        f'INDEX(IF({q}$D$5="SI",{_rngc(master,"Unidad impresa")},'
        f'{_rngc(us,"Unidad impresa")}),{q}{FIL})))')
    for k in range(1, npack + 1):
        curvas.cell(2 + k, c0).value = f'=IFERROR(INDEX({_q(TR)},{k}),NA())'
        curvas.cell(2 + k, c0 + 1).value = f'=IFERROR(INDEX({_q(VR)},{k}),NA())'
    curvas.cell(2, c0 + 3, "T consulta").font = HDR_F
    curvas.cell(2, c0 + 4, "Punto consultado").font = HDR_F
    curvas.cell(3, c0 + 3).value = f"={q}$D$11"
    curvas.cell(3, c0 + 4).value = f"={q}{VTAB}"

    from openpyxl.chart import Reference, Series, ScatterChart
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.layout import Layout, ManualLayout
    from openpyxl.drawing.line import LineProperties
    ch = ScatterChart()
    ch.title = "Propiedad tabulada frente a la temperatura"
    ch.style = 13
    ch.scatterStyle = "line"
    ch.x_axis.title = "Temperatura  [°C en SI  ·  °F en US]"
    ch.y_axis.title = "Valor TABULADO (antes del factor de escala)"
    ch.height, ch.width = 9.5, 26
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.13, y=0.15, w=0.80, h=0.65))
    xs = Reference(curvas, min_col=c0, min_row=3, max_row=2 + npack)
    ys = Reference(curvas, min_col=c0 + 1, min_row=2, max_row=2 + npack)
    s1 = Series(ys, xs, title_from_data=True)
    # Linea continua sin marcadores, mismo azul de la banda (regla de estilo 3).
    s1.marker = Marker(symbol="none")
    s1.smooth = False
    s1.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=BLUE, w=19050))
    ch.series.append(s1)
    xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
    yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
    s2 = Series(yq, xq, title_from_data=True)
    s2.marker = Marker(symbol="diamond", size=10,
                       spPr=GraphicalProperties(
                           solidFill="FF0000", ln=LineProperties(solidFill="FF0000")))
    s2.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.series.append(s2)
    ws.add_chart(ch, f"A{rg + 2}")

    for cc, w in zip("ABCDEFGHIJKL",
                     [22, 20, 16, 16, 10, 3, 22, 20, 16, 16, 10, 3]):
        ws.column_dimensions[cc].width = w
    return ws


# ---------------------------------------------------------------------------
# Buscar_Ec_A2 y Buscar_Ej_A3 — factores de calidad del B31.3
# ---------------------------------------------------------------------------
# Son la mitad de maquina que los otros motores, y el plan no finge lo
# contrario: Ec y Ej son ESCALARES —el codigo publica un numero por fila, no una
# funcion de T—, asi que aqui no hay interpolacion, ni banda compacta, ni curva.
# Lo que si hay, y es la razon de peso para construirlos, es el bloque 3: el
# factor publicado es un MINIMO que puede subirse con examen suplementario, y
# hasta ahora el libro no lo decia en ninguna parte.
FACT_LST_COL = 36          # 36..39 -> listas dependientes de los niveles 1 a 4
FACT_AUX_COL = 41          # 41 = rotulo, 42 = valor

# El conmutador SI/US de la regla 10 es DEGENERADO en estas dos tablas y hay que
# decirlo: Ec y Ej son adimensionales y el codigo publica UNA sola tabla para los
# dos sistemas (no existen A-2C ni A-3C). Poner un desplegable inventaria una
# distincion que el codigo no hace; poner nada dejaria al usuario sin saber en
# que sistema esta. Se resuelve con una celda fija y rotulada.
TXT_ADIMENSIONAL = ("FACTOR ADIMENSIONAL — identico en SI y en US "
                    "(el codigo publica una sola tabla)")


def _rngf(info, colname):
    """Rango de una columna de una base de factores (layout FACT_COLS)."""
    L = CLF[colname]
    return f"{info['sheet']}!${L}${R_DATA}:${L}${info['last_row']}"


def formula_factor_aplicable(fil, admite, basico, incfac):
    """Factor que de verdad aplica tras el examen suplementario.

    Sale de dos frases del codigo, no de un criterio propio:
      · Nota (4) de A-2 — «The HIGHER factor from Table 302.3.3-1 may be
        substituted for this factor»  -> MAX(basico, factor del examen).
      · para. 302.3.3(c) — «In no case shall the quality factor exceed 1.00»
        -> MIN(1, ...).
    Y solo se aplica donde la fila lo admite: la Nota (5) dice lo contrario —que
    el factor YA supone el examen—, y confundir las dos mueve el espesor.

    Se genera en una funcion porque verificar.py §6c la vuelve a emitir para
    recalcularla en Excel: asi la prueba ejerce el original, no una copia.
    """
    return (f'=IF({fil}="","",'
            f'IF(LEFT({admite},2)<>"SI",{basico},'
            f'IF({incfac}="",{basico},MIN(1,MAX({basico},{incfac})))))')


def build_buscador_factor(wb, name, titulo, subtitulo, rng, info, niveles,
                          simbolo, incremento):
    """Motor de un factor de calidad. `niveles` son los pasos 1..n de la cascada
    (el nivel 0 es siempre el grupo impreso del codigo)."""
    hl = get_column_letter
    ws = new_sheet(wb, name, titulo, subtitulo)
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30

    banda(ws, 4, f"1 · SELECCION   —   todo por lista desplegable")
    fija = _mrg(ws, 5, 1, NCOLS, TXT_ADIMENSIONAL)
    fija.font = Font(name="Calibri", size=10, bold=True, color=BLUE)
    fija.fill = CARD_FILL
    fija.border = CARD_BORDER
    fija.alignment = Alignment(vertical="center", indent=1)
    _nota(fija, "Celda fija, no editable: no es un selector. Ec y Ej son "
                "adimensionales y el B31.3 publica una sola tabla para los dos "
                "sistemas de unidades; no existen A-2C ni A-3C. Es la excepcion "
                "declarada a la regla del conmutador SI/US: se cumple su espiritu "
                "—el usuario sabe siempre en que sistema lee y nunca ve un valor "
                "convertido— sin fabricar un interruptor que no gobierna nada.")
    ws.row_dimensions[5].height = 18

    todos = [("0 · Grupo del codigo",
              "Entrada: paso 0 de la cascada. Grupo de material tal como lo "
              "encabeza la tabla del codigo.")] + list(niveles)
    for i, (et, com) in enumerate(todos):
        r = 6 + i
        e = _mrg(ws, r, 1, 3, et)
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        _nota(e, com)
        c = _mrg(ws, r, 4, 6, "")
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        ws.row_dimensions[r].height = 17
    r_ult = 6 + len(todos) - 1
    celdas = [f"$D${6 + i}" for i in range(len(todos))]

    dv_list(ws, "D6", "=" + rng["FAM"], todos[0][1])
    claves = ["$D$6"]
    for i in range(1, len(todos)):
        claves.append("&\"|\"&".join(celdas[:i + 1]))
    fuentes = [("C", rng["CK"], rng["CV"], rng.get("maxC", 20)),
               ("F", rng["FK"], rng["FV"], rng.get("maxF", 20)),
               ("S", rng["SK"], rng["SV"], rng.get("maxS", 20)),
               ("G", rng["GK"], rng["GV"], rng.get("maxG", 20))]
    for i in range(1, len(todos)):
        _, kr, vr, mx = fuentes[i - 1]
        key = claves[i - 1]
        col = FACT_LST_COL + i - 1
        L = hl(col)
        mx = max(1, min(int(mx), 250))
        ws.cell(R_HDR, col, f"lista nivel {i}").font = SRC_F
        for k in range(1, mx + 1):
            ws.cell(R_DATA + k - 1, col).value = (
                f'=IF(COUNTIF({kr},{key})<{k},"",'
                f'INDEX({vr},MATCH({key},{kr},0)+{k}-1))')
        ws.column_dimensions[L].hidden = True
        dv_list(ws, celdas[i], f"=${L}${R_DATA}:${L}${R_DATA + mx - 1}", todos[i][1])

    # Semaforo junto al ultimo nivel de la cascada.
    ay = _mrg(ws, r_ult, 7, 9)
    ay.value = f'=IF({celdas[-1]}="","SELECCION INCOMPLETA","SELECCION COMPLETA")'
    ay.font, ay.fill, ay.border = SEL_BAD_FONT, SEL_BAD_FILL, BOX
    ay.alignment = Alignment(horizontal="center", vertical="center")
    _nota(ay, "Aviso automatico: SELECCION COMPLETA (verde) cuando la cascada esta "
              "resuelta; SELECCION INCOMPLETA (amarillo) mientras falte un paso.")
    rango_sem = f"G{r_ult}:I{r_ult}"
    ws.conditional_formatting.add(rango_sem, FormulaRule(
        formula=[f'{celdas[-1]}<>""'], fill=SEL_OK_FILL, font=SEL_OK_FONT))
    ws.conditional_formatting.add(rango_sem, FormulaRule(
        formula=[f'{celdas[-1]}=""'], fill=SEL_BAD_FILL, font=SEL_BAD_FONT))

    # --- auxiliares --------------------------------------------------------
    A = FACT_AUX_COL
    LA = hl(A + 1)

    def aux(fila, rotulo, formula):
        ws.cell(fila, A, rotulo).font = SRC_F
        ws.cell(fila, A + 1).value = formula
        return f"${LA}${fila}"

    # La clave completa siempre tiene cinco segmentos: los niveles que la tabla
    # no usa van con la etiqueta (no aplica), igual que en la base.
    relleno = "".join(f'&"|{NO_APLICA}"' for _ in range(5 - len(todos)))
    KEY = aux(4, "clave k4", "=" + "&\"|\"&".join(celdas) + relleno)
    FIL = aux(5, "fila", f'=IFERROR(MATCH({KEY},{rng["K4"]},0),"")')
    MIDC = aux(6, "material_id", f'=IF({FIL}="","",INDEX({rng["ID"]},{FIL}))')

    def col(nombre):
        return f'INDEX({_rngf(info, nombre)},{FIL})'

    def val_of(nombre):
        return f'=IF({FIL}="","—",IF({col(nombre)}="","—",{col(nombre)}))'

    BASICO = aux(7, "factor basico", f'=IF({FIL}="","",{col("Factor")})')
    ADMITE = aux(8, "admite incremento",
                 f'=IF({FIL}="","",{col("Admite incremento")})')
    for jj in (A, A + 1):
        ws.column_dimensions[hl(jj)].hidden = True

    # ------------------------- 2 · FACTOR BASICO ---------------------------
    r = r_ult + 2
    banda(ws, r, f"2 · FACTOR BASICO {simbolo}   —   tal como lo imprime el codigo")
    kpis = [(1, f"{simbolo} BASICO", "=" + BASICO, '="adimensional"',
             f"Calculo: factor {simbolo} impreso por la tabla del codigo para la fila "
             "resuelta por la cascada. Es un MINIMO: vea el bloque 3."),
            (4, "TABLA DEL CODIGO", val_of("Tabla"), val_of("Paginas PDF (fuente)"),
             "Calculo: tabla del Apendice A de la que sale la fila, con las paginas "
             "del PDF del codigo que declara la extraccion."),
            (7, "NOTAS CITADAS", val_of("Notas citadas"), '="ver bloque 4"',
             "Calculo: numeros de nota al pie que cita esta fila. Su texto integro "
             "esta transcrito en el bloque 4."),
            (10, "ADMITE INCREMENTO", val_of("Admite incremento"),
             '="ver bloque 3"',
             "Calculo: si el codigo permite subir el factor con examen "
             "suplementario. Se deriva de las notas que cita la propia fila, no de "
             "una suposicion.")]
    for c1, tit, valf, uni, com in kpis:
        t = _mrg(ws, r + 1, c1, c1 + 2, tit)
        t.font, t.fill = KPI_TIT_F, BAND_FILL
        t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        v = _mrg(ws, r + 2, c1, c1 + 2, valf)
        v.font, v.fill = (KPI_VAL_F if c1 == 1 else VAL_F), KPI_FILL
        v.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        _nota(v, com)
        u = _mrg(ws, r + 3, c1, c1 + 2, uni)
        u.font, u.fill = UNIT_F, KPI_FILL
        u.alignment = Alignment(horizontal="center", wrap_text=True)
        for rr in (r + 1, r + 2, r + 3):
            for cc in range(c1, c1 + 3):
                ws.cell(rr, cc).border = CARD_BORDER
    ws.row_dimensions[r + 1].height = 26
    ws.row_dimensions[r + 2].height = 32
    ws.row_dimensions[r + 3].height = 15
    r += 5

    # --------- 3 · FACTOR INCREMENTADO POR EXAMEN SUPLEMENTARIO ------------
    banda(ws, r, "3 · FACTOR INCREMENTADO POR EXAMEN SUPLEMENTARIO")
    r += 1
    if incremento["tipo"] == "tabla":
        inc = incremento["info"]
        exa = (f'{inc["sheet"]}!${hl(inc["col_examen"])}${R_DATA}:'
               f'${hl(inc["col_examen"])}${inc["last_row"]}')
        fac = (f'{inc["sheet"]}!${hl(inc["col_factor"])}${R_DATA}:'
               f'${hl(inc["col_factor"])}${inc["last_row"]}')
        com_ex = ("Entrada: combinacion de examenes suplementarios realmente "
                  f'realizada, tal como la enumera la {inc["tabla"]}. Solo tiene '
                  "efecto si la fila seleccionada admite incremento (Nota (4)).")
        e = _mrg(ws, r, 1, 3, "Examen realizado")
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        _nota(e, com_ex)
        c = _mrg(ws, r, 4, NCOLS, "")
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        dv_list(ws, f"D{r}", "=" + exa, com_ex)
        celda_ex = f"$D${r}"
        r += 1
        INCFAC = aux(9, "factor del examen",
                     f'=IFERROR(INDEX({fac},MATCH({celda_ex},{exa},0)),"")')
        APLICA = aux(10, "factor aplicable",
                     formula_factor_aplicable(FIL, ADMITE, BASICO, INCFAC))
        campo(ws, r, 1, f"{simbolo} aplicable con ese examen", "=" + APLICA,
              '="adimensional"',
              com_val=f"Calculo: el mayor entre el {simbolo} basico y el factor que "
                      f'habilita el examen elegido, acotado a 1,00. Si la fila NO '
                      f"admite incremento, se conserva el basico.")
        campo(ws, r, 7, "Factor que habilita ese examen", "=" + INCFAC,
              f'="{inc["tabla"]}"',
              com_val=f'Calculo: factor que la {inc["tabla"]} asocia a la '
                      "combinacion de examenes elegida arriba.")
        r += 1
        av = _mrg(ws, r, 1, NCOLS)
        av.value = (f'=IF({FIL}="","",IF(LEFT({ADMITE},2)="SI",'
                    f'"Esta fila admite incremento: "&{ADMITE},'
                    f'"Esta fila NO admite incremento: "&{ADMITE}))')
        av.font = Font(name="Calibri", size=10, bold=True, color="9C6500")
        av.fill = PatternFill("solid", fgColor="FFEB9C")
        av.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
        _nota(av, "Aviso automatico: repite, en palabras, si el codigo permite subir "
                  "el factor de esta fila. Se deriva de las notas que ella cita.")
        ws.row_dimensions[r].height = 18
        r += 1
        lim = _mrg(ws, r, 1, NCOLS, incremento["limite"])
        lim.font = SRC_F
        lim.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
        ws.row_dimensions[r].height = 26
        r += 2
    else:
        for texto, estilo in incremento["bloques"]:
            b = _mrg(ws, r, 1, NCOLS, texto)
            b.font = estilo
            b.alignment = Alignment(vertical="top", indent=1, wrap_text=True)
            ws.row_dimensions[r].height = max(16, 13 * (1 + len(texto) // 130))
            r += 1
        r += 1

    # ---------------------- 4 · NOTAS DEL CODIGO ---------------------------
    banda(ws, r, "4 · NOTAS DEL CODIGO   —   transcritas, las que cita la fila")
    r += 1
    nt = _mrg(ws, r, 1, NCOLS)
    nt.value = (f'=IF({FIL}="","Complete la seleccion para ver las notas.",'
                f'IF({col("Texto de las notas")}="",'
                f'"Esta fila no cita ninguna nota numerada.",'
                f'{col("Texto de las notas")}))')
    nt.font = Font(name="Calibri", size=9)
    nt.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 74
    _nota(nt, "Calculo: texto integro de las notas al pie que cita la fila "
              "seleccionada, transcrito del codigo.")
    r += 2

    # ------------------------ 5 · TRAZABILIDAD -----------------------------
    banda(ws, r, "5 · TRAZABILIDAD")
    r += 1
    campo(ws, r, 1, "Material resuelto", f"=IF({FIL}=\"\",\"—\",{MIDC})", None,
          com_val="Calculo: clave del registro resuelto por la cascada.")
    campo(ws, r, 7, "Grupo impreso", val_of("Grupo impreso"), None,
          com_val="Calculo: grupo tal como lo encabeza la tabla del codigo.")
    r += 1
    campo(ws, r, 1, "Archivo fuente", val_of("Archivo fuente"), None,
          com_val="Trazabilidad: archivo de resources/ del que sale la fila.")
    campo(ws, r, 7, "Paginas del PDF del codigo", val_of("Paginas PDF (fuente)"),
          None, com_val="Trazabilidad: paginas del PDF del codigo que declara la "
                        "extraccion para esta tabla.")
    r += 1
    obs = _mrg(ws, r, 1, NCOLS)
    obs.value = (f'=IF({FIL}="","",IF({col("Detalle del incremento")}="","",'
                 f'"Observacion de extraccion: "&{col("Detalle del incremento")}))')
    obs.font = SRC_F
    obs.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    _nota(obs, "Calculo: artefacto de extraccion conservado o normalizado en esta "
               "fila. Vacio si no hay ninguno.")
    r += 1
    rec = _mrg(ws, r, 1, NCOLS,
               f"Recordatorio: {simbolo} entra en el diseno por presion como "
               "t = P·D / (2·(S·E + P·Y)). En ningun otro sitio.")
    rec.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
    rec.fill = PatternFill("solid", fgColor=GREY)
    rec.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[r].height = 18

    for cc, w in zip("ABCDEFGHIJKL",
                     [22, 20, 18, 18, 12, 6, 22, 20, 16, 14, 12, 4]):
        ws.column_dimensions[cc].width = w
    return ws


def build_buscador_nm(wb, info, rangos, max_mat=60):
    ws = new_sheet(wb, "Buscar_NoMetalicos",
                   "BUSCADOR — ASME B31.3, APENDICE B: esfuerzos de diseno hidrostatico "
                   "y presion admisible de tuberias no metalicas",
                   "Elija la tabla y despues el material: se muestran todos los campos "
                   "impresos para ese material, con las unidades que emplea el codigo. "
                   "Todo por lista desplegable. Las propiedades fisicas de los no "
                   "metalicos (dilatacion C-2 y modulo C-4) estan en Buscar_Prop_B31_3.")
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    banda(ws, 4, "1 · SELECCION")
    coms = {5: "Entrada: elija la tabla del Apendice B o C del B31.3 (B-1, B-1C, "
              "C-1, C-3, etc.). Habilita la lista de materiales de esa tabla.",
           6: "Entrada: elija el material dentro de la tabla elegida arriba. "
              "Muestra en la seccion 2 todos los campos que imprime esa fila del "
              "codigo."}
    for r, et, val, ref in ((5, "Tabla del Apendice", None, rangos["NM_TABLA"]),
                            (6, "Material", "", None)):
        e = _mrg(ws, r, 1, 3, et)
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        _nota(e, coms[r])
        c = _mrg(ws, r, 4, 8, val)
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
    ws["D5"] = wb["DB_NoMetalicos"].cell(R_DATA, 2).value
    dv_list(ws, "D5", "=" + rangos["NM_TABLA"], coms[5])
    mx = max(1, min(int(max_mat), 250))
    ws.cell(R_HDR, 50, "lista material").font = SRC_F
    for k in range(1, mx + 1):
        ws.cell(R_DATA + k - 1, 50).value = (
            f'=IF(COUNTIF({rangos["NM_MATK"]},$D$5)<{k},"",'
            f'INDEX({rangos["NM_MATV"]},MATCH($D$5,{rangos["NM_MATK"]},0)+{k}-1))')
    ws.column_dimensions["AX"].hidden = True
    dv_list(ws, "D6", f"=$AX${R_DATA}:$AX${R_DATA + mx - 1}", coms[6])
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
    _nota(t, "Calculo: confirma la tabla y el material seleccionados, o pide "
             "completar la seleccion.")
    ws.row_dimensions[9].height = 26
    for k in range(1, 17):
        r = 9 + k
        cond = f'IF(OR($BH$1=0,{k}>$BH$2),""'
        campo(ws, r, 1,
              f'={cond},INDEX(DB_NoMetalicos!$E${R_DATA}:$E${last},$BH$1+{k}-1))',
              f'={cond},INDEX(DB_NoMetalicos!$F${R_DATA}:$F${last},$BH$1+{k}-1))',
              com_etq="Calculo: nombre del campo tal como lo imprime la tabla del "
                     "Apendice B/C para el material elegido (fila k de la ficha; "
                     "vacio si el material tiene menos campos que k).",
              com_val="Calculo: valor de ese campo, en la unidad que imprime la "
                     "propia tabla del codigo (no se convierte de unidades).")
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

    def lab(r, text, note=None, com=None):
        cl = ws.cell(r, 1, text)
        cl.font = Font(name="Calibri", size=10)
        if note:
            ws.cell(r, 7, note).font = Font(name="Calibri", size=9, italic=True,
                                            color="595959")
        if com:
            _nota(cl, com)

    def inp(cell, value="", com=None):
        c = ws[cell]
        c.value = value
        c.font, c.fill, c.border = IN_F, IN_FILL, BOX
        c.protection = Protection(locked=False)
        if com:
            _nota(c, com)

    ws.cell(105, 1, "7.  RESOLUCION DE MATERIAL — BASE DE DATOS ASME "
                    "(cascada de listas desplegables)")
    ws.cell(105, 1).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws.cell(105, 1).fill = TITLE_FILL
    for j2, h in enumerate(["Parametro", "", "Unidad", "Metal base", "Collar / parche",
                            "", "Referencia / Notas"], start=1):
        c = ws.cell(106, j2, h)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=GREY)
    com_modo212 = ("Entrada: 'Interpolado' aplica la interpolacion lineal del "
                  "codigo entre T1 y T2 (fila 120/121). 'Tabulado-conservador' "
                  "adopta directamente el valor tabulado en T2, sin interpolar. "
                  "Rige el S(T) resuelto de metal base y collar (fila 126).")
    lab(107, "Modo de lectura de S(T)   [MODO_S]", "Interpolado | Tabulado-conservador",
       com_modo212)
    inp("D107", "Interpolado", com_modo212)
    dv_list(ws, "D107", '"Interpolado,Tabulado-conservador"', com_modo212)
    lab(108, "Temperatura de evaluacion", "= T de la hoja de proceso (D25)",
       "Calculo: repite automaticamente la Temperatura de operacion de la seccion 1 "
       "(D25). No se edita aqui — para cambiar la temperatura de evaluacion, edite "
       "D25.")
    ws.cell(108, 4).value = "=$D$25"
    ws.cell(108, 3).value = "°C"
    niveles = [(109, "0 · Familia de material",
                "Entrada: paso 0 de la cascada, para el metal base (D) y el collar/"
                "parche (E). Elija la familia de material de la lista desplegable. "
                "Habilita la lista del paso 1 de esa misma columna."),
               (110, "1 · Composicion nominal",
                "Entrada: paso 1 de la cascada, dependiente del paso 0 de la misma "
                "columna (metal base en D, collar/parche en E)."),
               (111, "2 · Forma de producto",
                "Entrada: paso 2 de la cascada, dependiente de los pasos 0-1 de la "
                "misma columna."),
               (112, "3 · Especificacion (Spec. No.)",
                "Entrada: paso 3 de la cascada, dependiente de los pasos 0-2 de la "
                "misma columna."),
               (113, "4 · Tipo / Grado",
                "Entrada: paso 4 de la cascada, dependiente de los pasos 0-3. Al "
                "completarlo ya se resuelve un material_id (salvo variantes, ver "
                "paso 5)."),
               (114, "5 · Variante (clase / tamano) — opcional",
                "Entrada opcional: solo hace falta si el paso 4 deja mas de una fila "
                "posible. En blanco, se toma la primera variante encontrada.")]
    for r, t, com in niveles:
        lab(r, t, "Lista desplegable en cascada", com)
        inp(f"D{r}", com=com)
        inp(f"E{r}", com=com)

    # --- listas materializadas (columnas ocultas) --------------------------
    FAM_COL = 11
    ws.cell(R_HDR, FAM_COL, "lista familia").font = SRC_F
    for k in range(1, 41):
        ws.cell(R_DATA + k - 1, FAM_COL).value = (
            f'=IF($D$11=1,IFERROR(INDEX({rb["FAM"]},{k}),""),'
            f'IFERROR(INDEX({ri["FAM"]},{k}),""))')
    ws.column_dimensions[hl(FAM_COL)].hidden = True
    fam_ref = f'=${hl(FAM_COL)}${R_DATA}:${hl(FAM_COL)}${R_DATA + 39}'
    dv_list(ws, "D109", fam_ref, niveles[0][2])
    dv_list(ws, "E109", fam_ref, niveles[0][2])

    lv = [("C", "CK", "CV", 110, '{c}$109'),
          ("F", "FK", "FV", 111, '{c}$109&"|"&{c}$110'),
          ("S", "SK", "SV", 112, '{c}$109&"|"&{c}$110&"|"&{c}$111'),
          ("G", "GK", "GV", 113, '{c}$109&"|"&{c}$110&"|"&{c}$111&"|"&{c}$112'),
          ("V", "K4", "ID", 114,
           '{c}$109&"|"&{c}$110&"|"&{c}$111&"|"&{c}$112&"|"&{c}$113')]
    niv_com = {110: niveles[1][2], 111: niveles[2][2], 112: niveles[3][2],
              113: niveles[4][2], 114: niveles[5][2]}
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
                    f"=${L}${R_DATA}:${L}${R_DATA + mx - 1}", niv_com[target])
            col += 1

    aux = {"D": "W", "E": "X"}
    for cl, ac in aux.items():
        ws[f"{ac}109"] = (f'=${cl}$109&"|"&${cl}$110&"|"&${cl}$111&"|"&'
                          f'${cl}$112&"|"&${cl}$113')
        ws[f"{ac}110"] = (f'=IF($D$11=1,IFERROR(MATCH(${ac}$109,{rb["K4"]},0),0),'
                          f'IFERROR(MATCH(${ac}$109,{ri["K4"]},0),0))')
        ws.column_dimensions[ac].hidden = True

    com_etiq = {
        115: "Calculo: material_id resuelto por la cascada de esta columna "
             "(o el pegado a mano en la celda 'Variante' si la cascada no alcanza "
             "a identificar el material).",
        116: "Calculo: hoja base de datos (B31.3, II-D 1A o II-D 1B/3) donde se "
             "encontro el material_id resuelto.",
        117: "Calculo: indice de la base de datos activa (1=B31.3, 2=II-D 1A, "
             "3=II-D 1B/3), derivado del selector de aplicacion (D11, seccion 1) "
             "o del prefijo del material_id.",
        118: "Calculo: numero de fila de la base de datos activa donde se localizo "
             "el material_id. Vacio si no se encuentra.",
        119: "Calculo: cantidad de puntos tabulados (n_pts) que tiene esa fila en la "
             "banda compacta del codigo; determina el rango de la busqueda de T1/T2.",
        120: "Calculo: temperatura tabulada inmediatamente inferior (o igual) a la "
             "temperatura de evaluacion (D108), tomada de la banda compacta.",
        121: "Calculo: temperatura tabulada inmediatamente superior a la de "
             "evaluacion. Vacia si T1 es el ultimo punto tabulado.",
        122: "Calculo: esfuerzo admisible S tabulado en T1, para el material "
             "resuelto.",
        123: "Calculo: esfuerzo admisible S tabulado en T2. Vacio si T1 es el "
             "ultimo punto tabulado.",
        124: "Calculo: temperatura maxima admisible o limite de aplicabilidad del "
             "material. Por encima de este valor, el Dictamen de rango marca FUERA "
             "DE RANGO (el codigo prohibe extrapolar).",
        125: "Calculo: SIN MATERIAL SELECCIONADO si falta completar la cascada; "
             "MATERIAL NO ENCONTRADO si el material_id no aparece en la base; FUERA "
             "DE RANGO si no hay valor tabulado a esa T o si T supera la Temp. max.; "
             "OK en cualquier otro caso.",
        126: "Calculo: S(T) resuelto para esta columna — interpolacion lineal o "
             "valor tabulado-conservador segun el Modo de lectura (D107), o NA() si "
             "el Dictamen de rango no es OK. Alimenta D39/D40 de la seccion 3.",
    }
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
        lab(r, t, com=com_etiq[r])
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
        for rr in (115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126):
            _nota(ws.cell(rr, col2), com_etiq[rr])

    fl, fa = fac_info["sheet"], fac_info["last_row"]
    com_ej = ("Entrada: elija la clave de junta longitudinal impresa en la Tabla "
             "A-3 del B31.3 (tipo de junta/costura). Determina el factor de "
             "eficiencia de junta longitudinal Ej (columna G de MAP_Factores) que "
             "usa la ecuacion de espesor de pared requerido (seccion 3, D65/E65/F65) "
             "via D44.")
    lab(128, "Clave de junta longitudinal (Tabla A-3) -> Ej", "Lista desplegable de "
       "MAP_Factores", com_ej)
    inp("D128", "A106 | Seamless pipe", com_ej)
    dv_list(ws, "D128", f"={fl}!$A${R_DATA}:$A${fa}", com_ej)
    ws.cell(128, 3).value = "adimensional"
    ws.cell(128, 5).value = (f'=IFERROR(INDEX({fl}!$G${R_DATA}:$G${fa},'
                             f'MATCH($D$128,{fl}!$A${R_DATA}:$A${fa},0)),1)')
    _nota(ws.cell(128, 5), "Calculo: Ej por lookup de la clave elegida en D128 contra "
                          "la Tabla A-3 de MAP_Factores; 1 si no se encuentra.")
    com_ec = ("Entrada informativa: clave de fundicion (Tabla A-2 del B31.3) para "
             "consultar su factor de eficiencia Ec. No alimenta ningun calculo de "
             "esta hoja; se deja como referencia si el metal base o el collar es de "
             "fundicion.")
    lab(129, "Ec (fundicion, Tabla A-2) — informativo", com=com_ec)
    inp("D129", com=com_ec)
    dv_list(ws, "D129", f"={fl}!$A${R_DATA}:$A${fa}", com_ec)
    ws.cell(129, 3).value = "adimensional"
    ws.cell(129, 5).value = (f'=IFERROR(INDEX({fl}!$G${R_DATA}:$G${fa},'
                             f'MATCH($D$129,{fl}!$A${R_DATA}:$A${fa},0)),"")')
    _nota(ws.cell(129, 5), "Calculo: Ec por lookup de la clave elegida en D129 contra "
                          "la Tabla A-2 de MAP_Factores; vacio si no se encuentra.")

    ws["D39"] = '=IF($E$125="OK",$E$126,NA())'
    ws["G39"] = "S(T) del collar — base de datos ASME (seccion 7)"
    _nota(ws["D39"], "Calculo: trae el S(T) del collar/parche resuelto en la seccion "
                    "7 (E126) si su Dictamen de rango es OK; NA() si no. Alimenta "
                    "'Esf. admisible del collar' de la seccion 1 (D39 original del "
                    "maestro queda sustituido por este valor).")
    ws["D40"] = '=IF($D$125="OK",$D$126,NA())'
    ws["G40"] = "S(T) del metal base — base de datos ASME (seccion 7)"
    _nota(ws["D40"], "Calculo: trae el S(T) del metal base resuelto en la seccion 7 "
                    "(D126) si su Dictamen de rango es OK; NA() si no.")
    ws["D44"] = "=$E$128"
    ws["G44"] = "Ej por lookup (B31.3 Tabla A-3) — seccion 7"
    ws["D44"].font = Font(name="Calibri", size=10)
    ws["D44"].fill = PatternFill(fill_type=None)
    ws["D44"].protection = Protection(locked=True)
    _nota(ws["D44"], "Calculo: repite el Ej resuelto en la seccion 7 (E128) a partir "
                    "de la clave de junta longitudinal elegida en D128. No se edita "
                    "aqui.")
    ws["F90"] = ('=IF(OR($D$125<>"OK",$E$125<>"OK"),"REVISAR — MATERIAL FUERA DE RANGO",'
                 'IF(AND(F84="CUMPLE",F85="CUMPLE",F86="CUMPLE",F87="CUMPLE"),"APTO","REVISAR"))')
    _nota(ws["F90"], "Calculo: DICTAMEN GLOBAL DEL DISEÑO. REVISAR — MATERIAL FUERA "
                    "DE RANGO si el metal base o el collar quedaron fuera de rango "
                    "en la seccion 7; APTO solo si ademas las 4 verificaciones de la "
                    "seccion 5 (filete, excentricidad, conformado en frio, espesor "
                    "de pared) dan CUMPLE; REVISAR en cualquier otro caso.")
    nota131 = ws.cell(131, 1, "La cascada filtra la base ASME por familia, composicion "
                    "nominal, forma de producto, especificacion y tipo/grado. Si un "
                    "material no aparece, localicelo en el buscador correspondiente y "
                    "pegue su material_id en la celda 'Variante'. La lista corta de "
                    "Datos_Ref queda como respaldo historico y ya no alimenta el "
                    "calculo.")
    nota131.font = SRC_F
    _nota(nota131, "Aviso fijo: como usar la cascada de la seccion 7 y que hacer si "
                  "un material no aparece en ella. No se edita.")

    comentar_art212_base(ws)

    ws.protection.password = "0000"
    ws.protection.sheet = True


def comentar_art212_base(ws):
    """Comentarios para las secciones 1-6 de Parche_PCC2_Art212 (geometria,
    presion de diseno, verificaciones): vienen ya escritas en el maestro
    sembrado, esta funcion NO las escribe ni las cambia, solo documenta lo que
    ya esta ahi con un comentario de Excel. Los textos describen fielmente la
    formula o el dato tal como esta en el maestro (Regla nº 1: no se inventa
    nada; ver el dump de referencia en Revision_MAP_Grupo si hace falta
    auditar)."""
    simples = {
        5: "Entrada: identificador del documento de este calculo (numero de MC).",
        6: "Entrada: descripcion del componente y el servicio reparado.",
        10: "Entrada: selector que conmuta el modo geometrico (tuberia, virola "
            "cilindrica o cabezal/esfera) y por tanto el codigo de construccion, "
            "la fuente del esfuerzo admisible y el factor kf de toda la hoja.",
        11: "Calculo: MODO = 1 si D10 es 'Tuberia (B31.3)', 3 si es 'Cabezal/"
            "esfera (VIII-1)', 2 en cualquier otro caso (virola cilindrica).",
        12: "Calculo: ASME B31.3 si MODO=1 (tuberia); ASME BPVC VIII-1 en los "
            "demas casos (virola o cabezal/esfera).",
        13: "Calculo: cita la tabla del codigo activo de la que sale el esfuerzo "
            "admisible S — Tabla A-1 del B31.3 si MODO=1, Tabla 1A de ASME II-D "
            "en los demas casos.",
        14: "Calculo: factor de geometria de la ecuacion de membrana, kf = 0,25 "
            "para esfera (MODO=3) o 0,5 para cilindro (tuberia o virola).",
        18: "Entrada: diametro nominal (NPS) en pulgadas, de la lista de B36.10M "
            "en Datos_Ref. Alimenta el lookup de OD y espesor de pared.",
        19: "Entrada: cedula (Schedule) segun B36.10M, en Datos_Ref, para el NPS "
            "elegido arriba.",
        20: "Calculo: diametro exterior (OD) por lookup del NPS (D18) en la "
            "tabla B36.10M de Datos_Ref.",
        21: "Calculo: espesor de pared por lookup del NPS (D18) y la cedula "
            "(D19) en la tabla B36.10M de Datos_Ref.",
        22: "Entrada: material del componente reparado (metal base), de la "
            "lista corta de Datos_Ref. Para el S(T) por temperatura use la "
            "cascada de la seccion 7.",
        23: "Entrada: material del parche o collar de refuerzo, de la lista "
            "corta de Datos_Ref. Para el S(T) por temperatura use la cascada de "
            "la seccion 7.",
        24: "Entrada informativa: fluido de servicio. No alimenta ningun "
            "calculo de esta hoja.",
        25: "Entrada: temperatura de operacion, en °C. Alimenta la Temperatura "
            "de evaluacion de la seccion 7 (D108) y por tanto todo el S(T) "
            "resuelto.",
        26: "Entrada: presion de operacion, en kg/cm². Es el caso 'Operacion' "
            "evaluado en la seccion 3.",
        27: "Entrada: presion de diseno tipica, en kg/cm². Es el caso 'Diseno "
            "tipico' de la seccion 3; se compara contra la presion maxima "
            "admisible del parche en la verificacion de la seccion 5 (fila 89).",
        28: "Entrada: presion envolvente (cota superior / rating), en kg/cm². "
            "Es el caso mas exigente, evaluado en la seccion 3.",
        29: "Entrada: espesor adoptado del parche o collar, en mm (debe ser >= "
            "espesor de pared). Alimenta la fuerza de membrana, el esfuerzo de "
            "soldadura y el peso estimado.",
        30: "Entrada: altura o dimension del parche, en mm, tomada del plano. "
            "Alimenta el peso estimado (seccion 4).",
        31: "Entrada: luz radial entre el parche y el componente, en mm (<=2 "
            "mm tipico). Alimenta el radio de conformado (Rf) y el desarrollo "
            "de la media carcasa.",
        32: "Entrada: cateto adoptado del filete perimetral, en mm. Se compara "
            "contra el filete minimo requerido en la verificacion de la "
            "seccion 5 (fila 84).",
        33: "Entrada: solape minimo del parche sobre metal sano, en mm, "
            "exigido por el Art. 212. Solo informativo en esta hoja.",
        34: "Entrada: diametro del defecto, caracterizado por UT/PT, en mm. "
            "Solo informativo en esta hoja.",
        35: "Entrada: distancia del defecto a la discontinuidad mas cercana "
            "(cordon o boquilla), en mm. Se compara contra L_min en la "
            "verificacion de la seccion 5 (fila 88) para decidir entre "
            "refuerzo 360° y parche local.",
        41: "Calculo: menor entre el esfuerzo admisible del collar (D39) y el "
            "del metal base (D40) — rige el diseno por criterio conservador.",
        42: "Entrada: eficiencia de junta de filete adoptada (Art. 212 ec. 4); "
            "verificar contra el WPS/PQR calificado.",
        43: "Entrada: factor Y de la Tabla 304.1.1 del B31.3, segun material y "
            "temperatura; se ingresa a mano, no se interpola automaticamente.",
        45: "Entrada: densidad del acero adoptada para el calculo de peso del "
            "parche/collar (seccion 4).",
        46: "Entrada: factor multiplicador de la presion de diseno para la "
            "prueba hidrostatica (tipico 1,5x, B31.3 345.4.2). Alimenta la "
            "especificacion de prueba de hermeticidad (fila 99).",
        47: "Constante: factor de conversion de kg/cm² a MPa (1 kg/cm² = "
            "0,0980665 MPa). No cambiar salvo error de unidades.",
        48: "Calculo: 1,5 veces el esfuerzo admisible gobernante (D41) — limite "
            "de la ecuacion 5 del Art. 212 para el esfuerzo de soldadura total.",
        52: "Calculo: diametro a media pared, Dm = OD − t.",
        53: "Calculo: radio medio, Rm = Dm/2.",
        54: "Calculo: radio interior, Ri = OD/2 − t.",
        55: "Calculo: excentricidad de la carga, e = (T_parche + t_pared)/2.",
        56: "Calculo: radio de conformado de la fibra media, Rf = OD/2 + luz + "
            "T_parche/2; se usa en la deformacion por conformado (ec. 7).",
        57: "Calculo: coeficiente C_sw tal que S_w = P·C_sw; relaciona la "
            "presion evaluada con el esfuerzo de soldadura (ec. 1 y 5 "
            "combinadas).",
        73: "Calculo: distancia minima a una discontinuidad, L_min = "
            "2·RAIZ(Rm·t) (ec. 3). Por debajo de esta distancia el defecto "
            "exige refuerzo de 360° en vez de parche local (verificacion de la "
            "seccion 5, fila 88).",
        74: "Calculo: presion maxima admisible del parche, en MPa: 1,5·Sa / "
            "C_sw.",
        75: "Calculo: igual que D74, convertido a kg/cm² con el factor D47.",
        76: "Calculo: relacion entre la presion maxima admisible del parche y "
            "la presion de operacion (D26); entre mas alto, mayor margen.",
        77: "Calculo: deformacion por conformado en frio, 50·T_parche/Rf "
            "(ec. 7), en %; debe ser <=5% (verificacion de la seccion 5, fila "
            "86) para no requerir tratamiento termico de conformado.",
        78: "Calculo: desarrollo de la media carcasa del collar, "
            "(π/2)·(OD + 2·luz + T_parche).",
        79: "Calculo: longitud real de corte de la plancha, igual al "
            "desarrollo menos 3 mm de luz de raiz.",
        80: "Calculo: peso estimado del parche/collar (dos mitades), a partir "
            "de la longitud de corte, la altura, el espesor y la densidad del "
            "acero.",
    }
    for r, com in simples.items():
        c = ws.cell(r, 4)
        if c.value is not None:
            _nota(c, com)

    trip = {
        61: "Calculo: repite, para este caso (Operacion / Diseno tipico / "
            "Envolvente), la presion correspondiente de la seccion 1, en "
            "kg/cm².",
        62: "Calculo: conversion a MPa de la presion de este caso, "
            "multiplicando por el factor de conversion D47.",
        63: "Calculo: fuerza de membrana de este caso (ec. 1 del Art. 212): "
            "kf · P(MPa) · Dm.",
        64: "Calculo: cateto de filete minimo requerido para este caso "
            "(ec. 4): F_m / (E · Sa).",
        65: "Calculo: espesor de pared requerido para este caso, por B31.3 "
            "(modo tuberia) o por VIII-1 UG-27 (modo esfera/cilindro), segun "
            "la presion evaluada de este caso.",
        66: "Calculo: componente de membrana del esfuerzo de soldadura para "
            "este caso (ec. 5): F_m / T_parche.",
        67: "Calculo: componente de flexion del esfuerzo de soldadura para "
            "este caso (ec. 5): 6·F_m·e / T_parche².",
        68: "Calculo: esfuerzo de soldadura total de este caso (membrana + "
            "flexion); se compara contra el limite 1,5·Sa en la fila 69.",
        69: "Calculo: CUMPLE si el esfuerzo de soldadura total de este caso "
            "(fila 68) no supera 1,5 veces el esfuerzo admisible gobernante "
            "(D48).",
    }
    for r, com in trip.items():
        for col in (4, 5, 6):
            c = ws.cell(r, col)
            if c.value is not None:
                _nota(c, com)

    verif = {
        84: ("Calculo: cateto de filete minimo requerido, el mayor de los tres "
             "casos de presion (fila 64).",
             "Entrada: cateto adoptado (repite D32).",
             "Calculo: CUMPLE si el cateto adoptado es mayor o igual al "
             "requerido."),
        85: ("Calculo: limite de esfuerzo por excentricidad, 1,5·Sa (repite "
             "D48).",
             "Calculo: esfuerzo de soldadura total adoptado para el diseno "
             "(repite E68, caso Diseno tipico).",
             "Calculo: CUMPLE si el esfuerzo adoptado no supera el limite."),
        86: ("Entrada: limite normativo de deformacion por conformado en "
             "frio, 5 % (ec. 7).",
             "Calculo: deformacion por conformado calculada (repite D77).",
             "Calculo: CUMPLE si la deformacion calculada no supera el 5 %."),
        87: ("Calculo: espesor de pared requerido, caso Envolvente (repite "
             "F65, el mas exigente).",
             "Entrada: espesor de pared real adoptado (repite D21).",
             "Calculo: CUMPLE si el espesor real es mayor o igual al "
             "requerido."),
        88: ("Calculo: distancia minima a la discontinuidad, L_min (repite "
             "D73).",
             "Entrada: distancia real del defecto a la discontinuidad (repite "
             "D35).",
             "Calculo: 'Refuerzo 360°' si la distancia real es menor que "
             "L_min; 'Parche local' en otro caso."),
        89: ("Calculo: presion maxima admisible del parche, en kg/cm² "
             "(repite D75).",
             "Entrada: presion de diseno adoptada (repite D27).",
             "Calculo: 'OK — parche' si la presion de diseno no supera la "
             "maxima admisible del parche; 'Migrar (Art.206)' si la supera."),
    }
    for r, (cd, ce, cf) in verif.items():
        for col, com in ((4, cd), (5, ce), (6, cf)):
            c = ws.cell(r, col)
            if c.value is not None:
                _nota(c, com)

    _nota(ws.cell(90, 1), "Resultado global del diseno: ver el dictamen calculado "
                         "en la celda F90 (a la derecha).")

    textos = {
        93: "Calculo: redacta automaticamente la especificacion de metodo de "
            "reparacion, citando la aplicacion (D10), el codigo (D12), el "
            "espesor del parche/collar (D29) y el peso estimado (D80).",
        94: "Entrada/texto estandar: especificacion de juntas de cierre "
            "(preparacion, proceso y consumibles de soldadura). Editable si "
            "el WPS del proyecto exige otra cosa.",
        95: "Calculo: redacta la especificacion de filetes perimetrales "
            "citando el cateto adoptado (D32) y el solape minimo (D33) de la "
            "seccion 1.",
        96: "Entrada/texto estandar: especificacion de soldadura en servicio "
            "(Art. 210) — electrodo, precalentamiento y END diferido. "
            "Editable segun el procedimiento calificado del proyecto.",
        97: "Entrada/texto estandar: especificacion de ensayos no "
            "destructivos. Editable segun el criterio de aceptacion del "
            "codigo de construccion activo (D12).",
        98: "Entrada/texto de ejemplo: especificacion de recubrimiento "
            "(esquema de un propietario de referencia). Ajustar a la "
            "especificacion real del propietario del equipo.",
        99: "Calculo: redacta la especificacion de prueba de hermeticidad, "
            "citando el factor de prueba hidrostatica (D46) y la presion de "
            "diseno (D27).",
    }
    for r, com in textos.items():
        c = ws.cell(r, 2)
        if c.value is not None:
            _nota(c, com)

    if ws.cell(101, 1).value is not None:
        _nota(ws.cell(101, 1), "Aviso fijo: alcance y limitaciones de la "
                              "herramienta. No se edita.")


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
     "DB_E — Tablas TM-1..5 (modulo E) · DB_TE — Tablas TE-1..5, tal como estan impresas "
     "(Coeficientes A, B y C por temperatura) · DB_TE_G — las mismas TE-1..5, pivotadas por "
     "GRUPO (solo Coeficiente B), la que alimenta Buscar_Prop_IID · DB_PRD — Poisson y "
     "densidad. DB_E, DB_TE_G y DB_PRD se indexan por GRUPO de material (el rotulo impreso "
     "en TM-1 / TE-1, no la familia de navegacion): ver MAP_Grupo.\n"
     "DB_B31_C / DB_B31_CC — Apendice C del B31.3 ENTERO en una sola base por edicion: "
     "C-1/C-1C (dilatacion de metales), C-2 (dilatacion de no metalicos), C-3/C-3C "
     "(modulo de metales) y C-4 (modulo de no metalicos). 201 filas por edicion.\n"
     "DB_NoMetalicos — Apendice B (HDS y presion admisible de tuberia no metalica).\n"
     "MAP_Factores (Ej/Ec) · MAP_Grupo · Notas_Codigo · DB_Listas · _meta (trazabilidad)."),
    ("3b. Propiedades fisicas: que motor usar",
     "Buscar_Prop_B31_3 — Apendice C del B31.3, para TUBERIA. Se elige primero QUE "
     "propiedad se necesita (nivel 0 de la cascada) y el motor decide en que tabla vive:\n"
     "  · DILATACION TERMICA - METALES → Tabla C-1 (SI) / C-1C (US). Dos coeficientes por "
     "material: A = coeficiente medio, en 10^-6 mm/mm/°C (10^-6 in./in./°F en US); "
     "B = expansion lineal acumulada desde 20 °C (70 °F), en mm/m (in./100 ft).\n"
     "  · MODULO DE ELASTICIDAD - METALES → Tabla C-3 (SI) / C-3C (US). El valor impreso se "
     "multiplica por 10^3 para dar MPa, o por 10^6 para dar psi. El motor aplica el factor "
     "a la vista y la seccion 4 muestra el valor impreso sin el.\n"
     "  · DILATACION TERMICA - NO METALICOS → Tabla C-2, y MODULO ELASTICIDAD CORTO PLAZO - "
     "NO METALICOS → Tabla C-4. Estas dos NO dependen de la temperatura: publican un valor "
     "unico, la grafica queda vacia a proposito y el ESTADO dice VALOR UNICO. C-2 se divide "
     "por 10^6 para dar mm/mm/°C.\n"
     "Por que el conmutador SI/US se comporta distinto: C-1/C-1C y C-3/C-3C son una tabla "
     "por edicion, asi que cambia de HOJA; C-2 y C-4 imprimen los dos sistemas en la MISMA "
     "tabla, asi que cambia de COLUMNA. En ningun caso hay conversion.\n"
     "El Apendice C no publica columna 'Temp. max.': el limite es el primer y el ultimo "
     "punto que la propia fila tabula. Fuera de ellos el resultado queda BLOQUEADO — "
     "sostener el ultimo valor seria extrapolar, que es lo que prohibe el codigo.\n"
     "Buscar_Prop_IID — modulo E (TM-1..5), dilatacion termica (TE-1..5, Coeficiente B, "
     "medio) y Poisson/densidad (PRD) de la II-D, para RECIPIENTES. Se indexa por GRUPO de "
     "material: consulte MAP_Grupo, columnas 'Grupo E (TM)' y 'Grupo dilatacion (TE)'. "
     "Lleva conmutador SI/US: cambia de HOJA (DB_E/DB_TE_G/DB_PRD en SI, "
     "DB_EC/DB_TE_GC/DB_PRDC en US), nunca convierte. El rotulo de GRUPO no cambia entre "
     "ediciones: solo el valor leido para ese grupo.\n"
     "Buscar_NoMetalicos — solo Apendice B: esfuerzo de diseno hidrostatico y presion "
     "admisible. Ya no ofrece propiedades fisicas."),
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
              "Buscar_BPVC_IID_B", "Buscar_Su", "Buscar_Sy", "Buscar_Prop_IID",
              "Buscar_Prop_B31_3", "Buscar_NoMetalicos", "Buscar_Ec_A2",
              "Buscar_Ej_A3", "Instrucciones"]

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
    t = _mrg(ws, 1, 1, DASH_NCOLS, "MOTOR DE CALCULO ASME PCC          Rev. 4")
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
        ("BPVC II-D · TM y PRD", ["Modulo E, Poisson y densidad",
                                  "por grupo de material · ver MAP_Grupo"],
         "Buscar_Prop_IID"),
        ("B31.3 · APENDICE C", ["Propiedades fisicas: dilatacion y modulo",
                                "Metales y no metalicos · SI y US"],
         "Buscar_Prop_B31_3"),
        ("B31.3 · APENDICE B", ["Esfuerzo de diseno hidrostatico",
                                "Tuberia no metalica · presion admisible"],
         "Buscar_NoMetalicos"),
        ("B31.3 · TABLA A-2", ["Factor de calidad de fundicion Ec",
                               "Basico y con examen suplementario"],
         "Buscar_Ec_A2"),
        ("B31.3 · TABLA A-3", ["Factor de calidad de junta longitudinal Ej",
                               "Por tipo de junta soldada"],
         "Buscar_Ej_A3"),
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
    te_g_si = build_dilatacion_grupo(res, wb, "SI")
    te_g_us = build_dilatacion_grupo(res, wb, "US")
    prd, prdc = build_prd(res, wb, "SI"), build_prd(res, wb, "US")
    apxc = build_apendice_c(res, wb, "SI")
    apxcc = build_apendice_c(res, wb, "US")
    verificar_paridad_apendice_c(apxc, apxcc)
    nm = build_nometalicos(res, wb)
    fac = build_map_factores(res, wb)
    ec_inc = build_ec_incremento(res, wb)
    a2 = build_factores(res, wb, "A-2")
    a3 = build_factores(res, wb, "A-3")
    ruta_dec = (Path(a.decisiones) if a.decisiones else
                Path(__file__).resolve().parent.parent / "decisiones_map_grupo.json")
    todas = [b313, b313c, iid, iidc, iidb, iidbc]
    map_last, map_stats, map_pend = build_map_grupo(
        res, wb, [iid, iidb], todas, ruta_dec, "SI")
    _, mapc_stats, mapc_pend = build_map_grupo(
        res, wb, [iidc, iidbc], todas, ruta_dec, "US")
    build_notas(res, wb)

    # listas simples de los buscadores por grupo
    wse, wste, wsprd, wsnm = wb["DB_E"], wb["DB_TE_G"], wb["DB_PRD"], wb["DB_NoMetalicos"]
    e_pairs = uniques(wse, e_si["last_row"], 3, 4)
    te_pairs = uniques(wste, te_g_si["last_row"], 3, 4)
    prd_pairs = uniques(wsprd, prd["last_row"], 3, 4)
    nm_pairs = uniques(wsnm, nm["last_row"], 2, 4)
    simples = [
        ("E_TABLA", sorted({k for k, _ in e_pairs})),
        ("E_K", [k for k, _ in e_pairs]), ("E_V", [v for _, v in e_pairs]),
        ("TE_TABLA", sorted({k for k, _ in te_pairs})),
        ("TE_K", [k for k, _ in te_pairs]), ("TE_V", [v for _, v in te_pairs]),
        ("PRD_TABLA", sorted({k for k, _ in prd_pairs if k})),
        ("PRD_K", [k for k, _ in prd_pairs]), ("PRD_V", [v for _, v in prd_pairs]),
        ("NM_TABLA", sorted({k for k, _ in nm_pairs})),
        ("NM_MATK", [k for k, _ in nm_pairs]), ("NM_MATV", [v for _, v in nm_pairs]),
    ]
    # El Apendice C entra en build_listas como una base mas: su contrato de las
    # 8 primeras columnas es identico al de STRESS_COLS, asi que la funcion no
    # necesita ni una linea de cambio.
    rangos = build_listas(wb, [("B313", b313), ("IID1A", iid), ("IIDB", iidb),
                               ("SU", su), ("SY", sy), ("APXC", apxc),
                               ("A2EC", a2), ("A3EJ", a3)], simples)

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
    te_tab = sorted({k for k, _ in te_pairs})
    prd_tab = sorted({k for k, _ in prd_pairs if k})
    from collections import Counter as _Cnt
    mx_e = max(_Cnt(k for k, _ in e_pairs).values())
    mx_te = max(_Cnt(k for k, _ in te_pairs).values())
    mx_prd = max(_Cnt(k for k, _ in prd_pairs).values())
    mx_nm = max(_Cnt(k for k, _ in nm_pairs).values())
    build_buscador_grupo(wb, curvas, [
        dict(titulo="MODULO DE ELASTICIDAD E — ASME BPVC II-D, Tablas TM-1 a TM-5",
             lst_tabla="E_TABLA", nm_key="E_K", nm_val="E_V", default_tabla=e_tab[0],
             info_si=e_si, info_us=e_us, temps=e_si["temps"],
             valor_lbl="Modulo E (valor tabulado)",
             unidad_si="x10^3 MPa", unidad_us="x10^6 psi", max_grupo=mx_e,
             nota_si="El valor real de E = valor tabulado x 10^3 MPa (factor del "
                     "titulo de la tabla).",
             nota_us="El valor real de E = valor tabulado x 10^6 psi (factor del "
                     "titulo de la tabla)."),
        dict(titulo="DILATACION TERMICA — ASME BPVC II-D, Tablas TE-1 a TE-5 "
                    "(Coeficiente B, medio)",
             lst_tabla="TE_TABLA", nm_key="TE_K", nm_val="TE_V", default_tabla=te_tab[0],
             info_si=te_g_si, info_us=te_g_us, temps=te_g_si["temps"],
             valor_lbl="Dilatacion (Coeficiente B, medio)",
             unidad_si="x10^-6 mm/mm/C", unidad_us="x10^-6 in/in/F", max_grupo=mx_te,
             nota_si="Coeficiente B (medio, de 20 C a T): el que se usa en calculo de "
                     "dilatacion/flexibilidad. Los Coeficientes A (instantaneo) y C "
                     "(expansion acumulada) siguen impresos tal cual, por temperatura, "
                     "en DB_TE.",
             nota_us="Coeficiente B (medio, de 70 F a T, el origen que imprime la "
                     "edicion US): el que se usa en calculo de dilatacion/"
                     "flexibilidad. Los Coeficientes A (instantaneo) y C (expansion "
                     "acumulada) siguen impresos tal cual, por temperatura, en "
                     "DB_TEC."),
        dict(titulo="POISSON Y DENSIDAD — ASME BPVC II-D, Tabla PRD",
             lst_tabla="PRD_TABLA", nm_key="PRD_K", nm_val="PRD_V",
             default_tabla=prd_tab[0], info_si=prd, info_us=prdc, temps=None,
             valor_lbl="", max_grupo=mx_prd,
             campos=[("Coeficiente de Poisson", 5, "adimensional", "adimensional"),
                     ("Densidad", 6, "kg/m3", "lb/in3")]),
    ], rangos)
    build_buscador_prop_c(wb, curvas, rangos["APXC"], apxc, apxcc)

    # --- los dos motores de factores de calidad ---------------------------
    par34, nota34, filas34 = texto_incremento_ej(res)
    build_buscador_factor(
        wb, "Buscar_Ec_A2",
        f"FACTOR DE CALIDAD DE FUNDICION Ec — ASME B31.3-2024, {a2['tabla']}",
        f"Valores tal como estan impresos · {a2['cita_pag']} · El factor publicado "
        "es un MINIMO: la Nota (4) permite subirlo con examen suplementario segun "
        "el para. 302.3.3(c) y la Tabla 302.3.3-1, y la Nota (5) avisa de lo "
        "contrario —que el factor YA supone ese examen—. Confundirlas "
        "es un error con consecuencia directa en el espesor requerido.",
        rangos["A2EC"], a2,
        [("1 · Especificacion",
          "Entrada: paso 1, dependiente del grupo. Spec. No. tal como lo imprime "
          "la Tabla A-2 (todas son ASTM, segun su Nota (1))."),
         ("2 · Descripcion",
          "Entrada: paso 2, dependiente de los pasos 0 y 1. Desambigua las "
          "especificaciones que el codigo repite en dos grupos (A352 figura en "
          "Carbon Steel y en Low and Intermediate Alloy Steel).")],
        "Ec",
        dict(tipo="tabla", info=ec_inc,
             limite=("Limite del propio codigo, para. 302.3.3(c): «Quality factors "
                     "higher than those shown in Table 302.3.3-1 do not result from "
                     "combining tests (2)(a) and (2)(b), or (3)(a) and (3)(b). In no "
                     "case shall the quality factor exceed 1.00.» De ahi salen el "
                     "MAX y el tope de 1,00 de la celda de arriba.")))
    build_buscador_factor(
        wb, "Buscar_Ej_A3",
        f"FACTOR DE CALIDAD DE JUNTA LONGITUDINAL Ej — ASME B31.3-2024, {a3['tabla']}",
        f"Valores tal como estan impresos · {a3['cita_pag']} · El factor lo decide "
        "el TIPO DE JUNTA, no el material: para el mismo A312 el codigo publica "
        "1,00 sin costura, 1,00 con EFW radiografiada al 100 %, 0,85 a doble tope y "
        "0,80 a tope simple.",
        rangos["A3EJ"], a3,
        [("1 · Especificacion",
          "Entrada: paso 1, dependiente del grupo. Spec. No. tal como lo imprime la "
          "Tabla A-3 (todas ASTM salvo API, segun su Nota (1))."),
         ("2 · Clase o tipo",
          "Entrada: paso 2, dependiente de los pasos 0 y 1. Clase o tipo impresos "
          f"(Type S/E/F, All, DW, SW, 12/22/32...). «{NO_APLICA}» cuando el codigo "
          "no imprime clase para esa especificacion."),
         ("3 · Descripcion de la junta",
          "Entrada: paso 3, y el que de verdad mueve el factor: sin costura, "
          "resistencia electrica, fusion electrica radiografiada al 100 %, doble "
          "tope, tope simple...")],
        "Ej",
        dict(tipo="declarado", bloques=[
            ("SI: el codigo publica un mecanismo para subir Ej, y no es una tabla "
             "aparte como la 302.3.3-1 del Ec — son las propias filas de la "
             "Table 302.3.4-1, transcrita integra abajo.",
             Font(name="Calibri", size=10, bold=True, color=NAVY)),
            (f"para. 302.3.4(b), transcrito: {par34}",
             Font(name="Calibri", size=9)),
            (f"Table 302.3.4-1, Nota (1): {nota34}" if nota34 else
             "No se pudo leer la Nota (1) de la Tabla 302.3.4-1 en resources/.",
             Font(name="Calibri", size=9, bold=True, color="9C6500")),
            ("El factor Ej lo decide el TIPO DE JUNTA/COSTURA/EXAMEN de la fila "
             "abajo, no la especificacion de material seleccionada arriba: "
             "identifique cual de las diez filas describe su junta y lea su Ej. "
             "El codigo no imprime una correspondencia fila a fila entre la "
             "Tabla A-3 y esta tabla, asi que el motor no la infiere.",
             Font(name="Calibri", size=9, bold=True, color="9C0006")),
        ] + bloques_tabla_ej(filas34)))
    build_buscador_nm(wb, nm, rangos, mx_nm)

    integrate_motor(wb, b313, iid, iidb, fac, rangos)
    deprecate_datos_ref(wb)

    counts = {
        "A-1 + A-4 -> DB_B31_3":
            f"{len(res.rows(f'{APX}/appendix_a/table_a_1.json')) + len(res.rows(f'{APX}/appendix_a/table_a_4.json'))}"
            f" -> {b313['last_row'] - R_DATA + 1}",
        "1A -> DB_BPVC_IID": f"{len(res.rows('ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_1a.json'))}"
                             f" -> {iid['last_row'] - R_DATA + 1}",
        "1B + 3 -> DB_BPVC_IID_B":
            f"{len(res.rows('ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_1b.json')) + len(res.rows('ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_3.json'))}"
            f" -> {iidb['last_row'] - R_DATA + 1}",
        "U -> DB_Su": f"{len(res.rows('ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_u.json'))}"
                      f" -> {su['last_row'] - R_DATA + 1}",
        "Y-1 -> DB_Sy": f"{len(res.rows('ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_y_1.json'))}"
                        f" -> {sy['last_row'] - R_DATA + 1}",
        "MAP_Grupo": " · ".join(f"{k}={v}" for k, v in map_stats.items()),
        "MAP_GrupoC": " · ".join(f"{k}={v}" for k, v in mapc_stats.items()),
    }
    build_meta(wb, counts)
    rewrite_instrucciones(
        wb, "Rev. 4 — El Apendice C del B31.3 pasa a motor propio (Buscar_Prop_B31_3) con "
            "conmutador SI/US y las cuatro tablas C-1..C-4 en una sola base; el Apendice B "
            "se queda solo en Buscar_NoMetalicos y la II-D en Buscar_Prop_IID. Se mantiene "
            "el Dashboard unico de navegacion de la Rev. 3 (libro con macros, .xlsm), la "
            "consulta por cascada de listas desplegables y el libro sin funciones de "
            "matriz dinamica.")

    # Los conteos del Dashboard salen de las mismas variables que alimentan
    # `counts`, nunca escritos a mano. El indicador cuenta lo que el codigo NO
    # resuelve: la regla textual que debe aplicar el ingeniero y los materiales
    # para los que II-D no publica el dato. Las filas AUTO no entran: citan la
    # nota que las sostiene y son auditables 1:1.
    n_sin_resolver = map_stats.get(E_TEXTUAL, 0) + map_stats.get(E_SIN, 0)
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
             "Buscar_BPVC_IID_B", "Buscar_Su", "Buscar_Sy", "Buscar_Prop_IID",
             "Buscar_Prop_B31_3", "Buscar_Ec_A2", "Buscar_Ej_A3",
             "Buscar_NoMetalicos", "Datos_Ref", "DB_B31_3", "DB_B31_3C", "DB_BPVC_IID",
             "DB_BPVC_IIDC", "DB_BPVC_IID_B", "DB_BPVC_IID_BC", "DB_Su", "DB_SuC",
             "DB_Sy", "DB_SyC", "DB_E", "DB_EC", "DB_TE", "DB_TEC", "DB_TE_G", "DB_TE_GC",
             "DB_PRD", "DB_PRDC",
             "DB_B31_C", "DB_B31_CC",
             "DB_NoMetalicos", "MAP_Factores", "DB_A2_Ec", "DB_A3_Ej",
             "DB_Ec_Incremento", "MAP_Grupo", "MAP_GrupoC",
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
    n_abiertas, n_cerradas = escribir_revision_map_grupo(
        map_pend + mapc_pend, map_stats, ruta_rev)
    ISSUES.append(f"MAP_Grupo: {n_abiertas} casos abiertos (admiten criterio de "
                  f"ingenieria) y {n_cerradas} cerrados por limite de la fuente "
                  f"(II-D no publica el dato; no se firman) -> {ruta_rev.name}")

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
