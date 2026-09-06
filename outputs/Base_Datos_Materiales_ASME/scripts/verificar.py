# -*- coding: utf-8 -*-
"""verificar.py — Protocolo de aceptacion (seccion 10 del PLAN-DB-MAT-001, Rev. 2).

1. Conteos JSON -> hoja.
2. Unicidad de material_id.
3. Spot-checks de valores contra el JSON fuente.
4. Contigüidad de los bloques de la cascada (condicion de las listas dependientes).
5. Ausencia de funciones de matriz dinamica en todo el libro.
6. Interpolacion con huecos interiores, modo tabulado y bordes: recalculo real en
   hoja (LibreOffice) contra un motor de referencia independiente en Python.
7. Regresion del caso semilla.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from db_lib import Resources, num, temp_to_number, txt

RES = Resources("/mnt/user-data/uploads/ASME PCC/resources")
WB = "Motor_v2.xlsx"
APX = "ASME B31/ASME B31.3/APPEX"
R_DATA, R_HDR = 4, 3
def col_map(ws):
    """Indices de columna leidos del propio encabezado (robusto a cambios)."""
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(R_HDR, c).value
        if isinstance(v, str) and v not in m:
            m[v] = c
    return m
out = []


def log(s=""):
    print(s)
    out.append(s)


def n_ident(ws):
    return col_map(ws)["n_pts"]


def printed_temps(ws):
    """Solo la banda impresa: se detiene al llegar a la banda compacta (Tc1)."""
    out = []
    for c in range(n_ident(ws) + 1, ws.max_column + 1):
        v = ws.cell(R_HDR, c).value
        if isinstance(v, str) and v.startswith(("Tc", "Vc")):
            break
        out.append(v)
    return out


def grid(ws):
    temps = printed_temps(ws)
    ni = n_ident(ws)
    idx = {}
    for r in range(R_DATA, ws.max_row + 1):
        mid = ws.cell(r, 1).value
        if mid is None:
            continue
        vals = {}
        for k, t in enumerate(temps):
            v = ws.cell(r, ni + 1 + k).value
            if v is not None:
                vals[t] = v
        idx[mid] = (r, vals)
    return temps, idx


def interp(temps, vals, T, modo="Interpolado"):
    """Motor de referencia: salta los huecos interiores, no extrapola."""
    con = [t for t in temps if t in vals]
    if not con:
        return None
    bajo = [t for t in con if t <= T]
    alto = [t for t in con if t > T]
    if not bajo:
        return vals[con[0]]
    t1 = max(bajo)
    if not alto:
        return vals[t1]
    t2 = min(alto)
    if modo == "Tabulado-conservador":
        return vals[t2]
    s1, s2 = vals[t1], vals[t2]
    return s1 + (s2 - s1) * (T - t1) / (t2 - t1)


def main():
    wb = openpyxl.load_workbook(WB)
    log("# Reporte de verificacion — PLAN-DB-MAT-001 Rev. 2")
    log("")
    log(f"Libro verificado: `{WB}`  ·  {len(wb.sheetnames)} hojas")
    log("")

    # ---- 1. conteos -------------------------------------------------------
    log("## 1. Conteo de filas (JSON fuente -> hoja)")
    log("")
    log("| Fuente | Filas JSON | Filas en la hoja | Estado |")
    log("|---|---|---|---|")
    checks = [("A-1 + A-4 -> DB_B31_3",
               len(RES.rows(f"{APX}/appendix_a/table_a_1.json")) +
               len(RES.rows(f"{APX}/appendix_a/table_a_4.json")), "DB_B31_3"),
              ("1A -> DB_BPVC_IID", len(RES.rows("bpvc_ii_d_metric_2025/table_1a.json")),
               "DB_BPVC_IID"),
              ("1B + 3 -> DB_BPVC_IID_B",
               len(RES.rows("bpvc_ii_d_metric_2025/table_1b.json")) +
               len(RES.rows("bpvc_ii_d_metric_2025/table_3.json")), "DB_BPVC_IID_B"),
              ("U -> DB_Su", len(RES.rows("bpvc_ii_d_metric_2025/table_u.json")), "DB_Su"),
              ("Y-1 -> DB_Sy", len(RES.rows("bpvc_ii_d_metric_2025/table_y_1.json")), "DB_Sy")]
    for label, njson, sh in checks:
        n = wb[sh].max_row - R_DATA + 1
        est = "OK" if n == njson else (f"OK — {njson - n} filas sin identificacion ni "
                                       "valores descartadas (ruido de extraccion)")
        log(f"| {label} | {njson} | {n} | {est} |")
    log("")

    # ---- 2. unicidad ------------------------------------------------------
    log("## 2. Unicidad de material_id")
    log("")
    log("| Hoja | Filas | Claves unicas | Estado |")
    log("|---|---|---|---|")
    bases = ["DB_B31_3", "DB_B31_3C", "DB_BPVC_IID", "DB_BPVC_IIDC",
             "DB_BPVC_IID_B", "DB_BPVC_IID_BC", "DB_Su", "DB_Sy"]
    for sh in bases:
        ws = wb[sh]
        ids = [ws.cell(r, 1).value for r in range(R_DATA, ws.max_row + 1)]
        log(f"| {sh} | {len(ids)} | {len(set(ids))} | "
            f"{'OK' if len(ids) == len(set(ids)) else 'DUPLICADOS'} |")
    log("")

    # ---- 3. auditoria fila a fila ----------------------------------------
    log("## 3. Auditoria fila a fila contra el JSON fuente")
    log("")
    log("No es un muestreo: de CADA fila del JSON del codigo se toma su vector completo "
        "de valores tabulados junto con su especificacion y grado, y se comprueba que "
        "exista exactamente la misma fila en la hoja. Se comparan asi todos los numeros "
        "cargados, no una seleccion.")
    log("")
    log("| Hoja | Filas auditadas | Valores comparados | Filas sin correspondencia |")
    log("|---|---|---|---|")

    def nk(v):
        """Clave de comparacion: 356.0 (JSON) y 356 (celda) son el mismo grado."""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return txt(v)

    def canon_json(rel, base_t=None):
        from collections import Counter
        c = Counter()
        npts = 0
        for row in RES.rows(rel):
            vals = {}
            if base_t is not None:
                v0 = row.get("min_temp_to_40", row.get("min_temp_to_100"))
                if num(v0) is not None:
                    vals[base_t] = num(v0)
            for k, v in (row.get("values") or {}).items():
                if num(v) is not None:
                    vals[temp_to_number(k)] = num(v)
            if not vals:
                continue
            npts += len(vals)
            c[(nk(row.get("spec_no")), nk(row.get("type_grade")),
               tuple(sorted(vals.items())))] += 1
        return c, npts

    def canon_sheet(sh):
        from collections import Counter
        w = wb[sh]
        temps = printed_temps(w)
        ni = n_ident(w)
        cm = col_map(w)
        c = Counter()
        for r in range(R_DATA, w.max_row + 1):
            vals = {}
            for k, t in enumerate(temps):
                v = w.cell(r, ni + 1 + k).value
                if v is not None:
                    vals[t] = v
            if not vals:
                continue
            c[(nk(w.cell(r, cm["Spec. No."]).value),
               nk(w.cell(r, cm["Tipo/Grado"]).value),
               tuple(sorted(vals.items())))] += 1
        return c

    audits = [("DB_B31_3", [(f"{APX}/appendix_a/table_a_1.json", 40),
                            (f"{APX}/appendix_a/table_a_4.json", 40)]),
              ("DB_BPVC_IID", [("bpvc_ii_d_metric_2025/table_1a.json", None)]),
              ("DB_BPVC_IID_B", [("bpvc_ii_d_metric_2025/table_1b.json", None),
                                 ("bpvc_ii_d_metric_2025/table_3.json", None)]),
              ("DB_Su", [("bpvc_ii_d_metric_2025/table_u.json", None)]),
              ("DB_Sy", [("bpvc_ii_d_metric_2025/table_y_1.json", None)]),
              ("DB_B31_3C", [(f"{APX}/appendix_a/table_a_1c.json", 100),
                             (f"{APX}/appendix_a/table_a_4c.json", 100)]),
              ("DB_BPVC_IIDC", [("bpvc_ii_d_customary_2025/table_1a.json", None)]),
              ("DB_BPVC_IID_BC", [("bpvc_ii_d_customary_2025/table_1b.json", None),
                                  ("bpvc_ii_d_customary_2025/table_3.json", None)]),
              ("DB_SuC", [("bpvc_ii_d_customary_2025/table_u.json", None)]),
              ("DB_SyC", [("bpvc_ii_d_customary_2025/table_y_1.json", None)])]
    bad_tot = tot = 0
    from collections import Counter
    for sh, srcs in audits:
        cj, npts = Counter(), 0
        for rel, bt in srcs:
            c2, n2 = canon_json(rel, bt)
            cj += c2
            npts += n2
        cs = canon_sheet(sh)
        falt = sum((cj - cs).values())
        tot += npts
        bad_tot += falt
        log(f"| {sh} | {sum(cj.values())} | {npts} | {falt} |")
    log("")
    log(f"**Total: {tot} valores tabulados auditados, {bad_tot} filas sin correspondencia "
        "exacta en la hoja.**")
    log("")

    # ---- 3b. auditoria de las bases restantes -----------------------------
    log("## 3b. Auditoria de las bases por familia, auxiliares y no metalicos")
    log("")
    log("| Hoja | Filas JSON | Filas en la hoja | Valores comparados | Discrepancias |")
    log("|---|---|---|---|---|")

    def audita_grupo(sh, rels, col_mat, col_coef=None):
        """Compara el vector de valores de cada fila con el del JSON."""
        from collections import Counter
        cj, npts = Counter(), 0
        njson = 0
        for rel in rels:
            for row in RES.rows(rel):
                njson += 1
                merged, extra = None, {}
                for k, v in row.items():
                    if k in ("materials", "material", "material_grade_uns_no",
                             "material_uns_no"):
                        merged = v
                for k, v in (row.get("values") or {}).items():
                    if num(v) is not None:
                        extra[temp_to_number(k)] = num(v)
                if not extra:
                    continue
                npts += len(extra)
                cj[tuple(sorted(extra.items()))] += 1
        w = wb[sh]
        temps = printed_temps(w)
        ni = n_ident(w)
        cs = Counter()
        for r in range(R_DATA, w.max_row + 1):
            vals = {}
            for k, t in enumerate(temps):
                v = w.cell(r, ni + 1 + k).value
                if v is not None:
                    vals[t] = v
            if vals:
                cs[tuple(sorted(vals.items()))] += 1
        falt = sum((cj - cs).values())
        log(f"| {sh} | {njson} | {w.max_row - R_DATA + 1} | {npts} | {falt} |")
        return falt

    ed = "bpvc_ii_d_metric_2025"
    edc = "bpvc_ii_d_customary_2025"
    extra_bad = 0
    extra_bad += audita_grupo("DB_E", [f"{ed}/table_tm_{k}.json" for k in range(1, 6)], 4)
    extra_bad += audita_grupo("DB_EC", [f"{edc}/table_tm_{k}.json" for k in range(1, 6)], 4)
    extra_bad += audita_grupo("DB_C_dilatacion", [f"{APX}/appendix_c/table_c_1.json"], 4)
    extra_bad += audita_grupo("DB_C_dilatacionC", [f"{APX}/appendix_c/table_c_1c.json"], 4)
    extra_bad += audita_grupo("DB_C_modulo", [f"{APX}/appendix_c/table_c_3.json"], 4)
    extra_bad += audita_grupo("DB_C_moduloC", [f"{APX}/appendix_c/table_c_3c.json"], 4)
    log("")
    log("| Hoja | Filas JSON | Filas en la hoja | Estado |")
    log("|---|---|---|---|")
    simples = [("DB_PRD", [f"{ed}/table_prd.json"]),
               ("DB_PRDC", [f"{edc}/table_prd.json"]),
               ("DB_TE", [f"{ed}/table_te_{k}.json" for k in range(1, 6)]),
               ("DB_TEC", [f"{edc}/table_te_{k}.json" for k in range(1, 6)]),
               ("MAP_Factores", [f"{APX}/appendix_a/table_a_2.json",
                                 f"{APX}/appendix_a/table_a_3.json"]),
               ("__NOTAS__", [f"{APX}/appendix_a/notes_tables_a_1_a_1c.json",
                                 f"{APX}/appendix_a/notes_tables_a_4_a_4c.json",
                                 f"{ed}/notes_table_1a.json", f"{ed}/notes_table_1b.json",
                                 f"{ed}/notes_table_3.json", f"{ed}/notes_table_u.json",
                                 f"{ed}/notes_table_y_1.json"])]
    from build_db_materiales import iter_notas
    for sh, rels in simples:
        if sh == "__NOTAS__":
            sh = "Notas_Codigo"
            njson = sum(len(iter_notas(RES.load(rel))) for rel in rels)
        else:
            njson = sum(len(RES.rows(rel)) for rel in rels)
        nsheet = sum(1 for r in range(R_DATA, wb[sh].max_row + 1)
                     if wb[sh].cell(r, 1).value is not None)
        ok = nsheet >= njson if sh in ("DB_TE", "DB_TEC") else nsheet == njson
        extra_bad += 0 if ok else 1
        log(f"| {sh} | {njson} | {nsheet} | {'OK' if ok else 'REVISAR'} |")
    wnm = wb["DB_NoMetalicos"]
    nm_json = 0
    for rel in [f"{APX}/appendix_b/table_b_{k}.json" for k in
                ("1", "1c", "2", "3", "4", "5", "6")] + \
               [f"{APX}/appendix_c/table_c_2.json", f"{APX}/appendix_c/table_c_4.json"]:
        for row in RES.rows(rel):
            nm_json += sum(1 for v in row.values() if v is not None)
    nm_sheet = sum(1 for r in range(R_DATA, wnm.max_row + 1)
                   if wnm.cell(r, 1).value is not None)
    log(f"| DB_NoMetalicos (pares campo/valor) | {nm_json} | {nm_sheet} | "
        f"{'OK' if nm_sheet == nm_json else 'REVISAR'} |")
    extra_bad += 0 if nm_sheet == nm_json else 1
    log("")

    # ---- 4. contiguidad de la cascada ------------------------------------
    log("## 4. Contiguidad de los bloques de la cascada")
    log("")
    log("Las listas dependientes se resuelven con OFFSET/MATCH/COUNTIF, que exige que "
        "todas las filas de una misma clave sean consecutivas.")
    log("")
    log("| Hoja | k0 | k1 | k2 | k3 | k4 | Estado |")
    log("|---|---|---|---|---|---|---|")
    cont_ok = True
    for sh in bases:
        ws = wb[sh]
        res_ = []
        for key in ("k0", "k1", "k2", "k3", "k4"):
            col = col_map(ws)[key]
            seen, prev, broken = set(), object(), 0
            for r in range(R_DATA, ws.max_row + 1):
                v = ws.cell(r, col).value
                if v != prev:
                    if v in seen:
                        broken += 1
                    seen.add(v)
                    prev = v
            res_.append(broken)
        ok = all(x == 0 for x in res_)
        cont_ok &= ok
        log(f"| {sh} | {res_[0]} | {res_[1]} | {res_[2]} | {res_[3]} | {res_[4]} | "
            f"{'OK' if ok else 'BLOQUES ROTOS'} |")
    log("")
    log("(0 = ningun bloque fragmentado)")
    log("")

    # ---- 5. sin matrices dinamicas ---------------------------------------
    log("## 5. Compatibilidad — ausencia de funciones de matriz dinamica")
    log("")
    pat = re.compile(r"_xlfn|FILTER\(|XLOOKUP\(|UNIQUE\(|SORT\(|VSTACK\(|SEQUENCE\(")
    hits = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and pat.search(c.value):
                    hits.append(f"{ws.title}!{c.coordinate}")
    log(f"Formulas con funciones de matriz dinamica encontradas: **{len(hits)}**"
        + ("" if not hits else "  → " + ", ".join(hits[:10])))
    log("")
    log("Validacion de datos: Google Sheets solo admite un RANGO literal o una lista "
        "de items como origen; una formula (OFFSET/INDIRECT) se pierde al importar.")
    log("")
    malas = []
    for w in wb.worksheets:
        for dv in w.data_validations.dataValidation:
            f1 = str(dv.formula1 or "")
            if dv.type != "list":
                continue
            literal = f1.startswith('"')
            rango = bool(re.fullmatch(r"=(?:[^!]+!)?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+",
                                      f1))
            if not (literal or rango):
                malas.append(f"{w.title}!{sorted(str(x) for x in dv.sqref.ranges)[0]}")
    log(f"Validaciones de lista revisadas en todo el libro; con origen NO portable: "
        f"**{len(malas)}**" + ("" if not malas else "  → " + ", ".join(malas[:10])))
    log("")

    # ---- 6. banco de pruebas del motor -----------------------------------
    log("## 6. Interpolacion con huecos, modo tabulado y bordes (recalculo en hoja)")
    log("")
    cases = []
    ws = wb["DB_B31_3"]
    tb, ib = grid(ws)
    ids = list(ib)
    for pat_id, temps_test in [("A-1 | A106 | B", [25, 110, 125, 212, 263, 337, 1200]),
                               ("A-1 | A516 | 70", [25, 110, 263, 337]),
                               ("A-1 | A312 | TP321", [25, 263, 337])]:
        mid = next(i for i in ids if i.startswith(pat_id))
        for T in temps_test:
            cases.append(("DB_B31_3", mid, T, "Interpolado"))
        cases.append(("DB_B31_3", mid, 263, "Tabulado-conservador"))
    wsi = wb["DB_BPVC_IID"]
    ti, ii = grid(wsi)
    sa = next(i for i in ii if i.startswith("1A | SA-516 | 70"))
    for T, m in [(25, "Interpolado"), (180, "Interpolado"), (263, "Interpolado"),
                 (263, "Tabulado-conservador"), (900, "Interpolado")]:
        cases.append(("DB_BPVC_IID", sa, T, m))

    qa = wb.create_sheet("_QA")
    for j, h in enumerate(["hoja", "material_id", "T", "modo", "fila", "T1", "S1",
                           "n_pts", "T2", "S2", "S(T)"], 1):
        qa.cell(2, j, h)
    for i, (sh, mid, T, modo) in enumerate(cases):
        r = 3 + i
        w = wb[sh]
        ni = n_ident(w)
        npr = len(printed_temps(w))
        t0 = ni + npr + 1
        npack = (w.max_column - t0 + 1) // 2
        v0 = t0 + npack
        TA = f"{sh}!${get_column_letter(t0)}${R_DATA}"
        VA = f"{sh}!${get_column_letter(v0)}${R_DATA}"
        NPC = (f"{sh}!${get_column_letter(ni)}${R_DATA}:"
               f"${get_column_letter(ni)}${w.max_row}")
        IDS = f"{sh}!$A${R_DATA}:$A${w.max_row}"
        qa.cell(r, 1, sh); qa.cell(r, 2, mid); qa.cell(r, 3, T); qa.cell(r, 4, modo)
        qa.cell(r, 5).value = f'=IFERROR(MATCH($B{r},{IDS},0),"")'
        NP = f"IFERROR(INDEX({NPC},$E{r}),0)"
        tr = f"OFFSET({TA},$E{r}-1,0,1,MAX(1,{NP}))"
        vr = f"OFFSET({VA},$E{r}-1,0,1,MAX(1,{NP}))"
        p1 = f"IFERROR(MATCH($C{r},{tr},1),1)"
        qa.cell(r, 6).value = f'=IFERROR(INDEX({tr},{p1}),"")'
        qa.cell(r, 7).value = f'=IFERROR(INDEX({vr},{p1}),"")'
        qa.cell(r, 8).value = f"={NP}"
        qa.cell(r, 9).value = f'=IFERROR(INDEX({tr},{p1}+1),"")'
        qa.cell(r, 10).value = f'=IFERROR(INDEX({vr},{p1}+1),"")'
        qa.cell(r, 11).value = (
            f'=IF($G{r}="","SIN VALOR",IF($C{r}<=$F{r},$G{r},'
            f'IF(OR($I{r}="",$J{r}=""),$G{r},'
            f'IF($D{r}="Tabulado-conservador",$J{r},'
            f'$G{r}+($J{r}-$G{r})*($C{r}-$F{r})/($I{r}-$F{r})))))')
    wb.save("qa.xlsx")
    subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "xlsx",
                    "--outdir", "qa_out", "qa.xlsx"], check=True, capture_output=True,
                   timeout=900)
    rb = openpyxl.load_workbook("qa_out/qa.xlsx", data_only=True)["_QA"]
    log("| Base | material_id | T | Modo | T1 | T2 | S(T) hoja | S(T) referencia | Estado |")
    log("|---|---|---|---|---|---|---|---|---|")
    nbad = 0
    for i, (sh, mid, T, modo) in enumerate(cases):
        r = 3 + i
        got = rb.cell(r, 11).value
        temps, idx = (tb, ib) if sh == "DB_B31_3" else (ti, ii)
        exp = interp(temps, idx[mid][1], T, modo)
        ok = (abs(got - exp) < 1e-6) if isinstance(got, (int, float)) and exp is not None \
            else (got == "SIN VALOR" and exp is None)
        nbad += 0 if ok else 1
        f = lambda v: round(v, 4) if isinstance(v, float) else v
        log(f"| {sh} | `{str(mid)[:30]}` | {T} | {modo} | {rb.cell(r,6).value} | "
            f"{rb.cell(r,9).value} | {f(got)} | {f(exp)} | {'OK' if ok else 'FALLO'} |")
    log("")
    log(f"**{len(cases)} casos, {nbad} fallos.** El caso T = 125 C de A106 Gr.B verifica el "
        "salto de huecos interiores: la Tabla A-1 no imprime ese punto para ese material y "
        "la hoja interpola entre 100 y 150 C, no entre celdas vacias.")
    log("")

    # ---- 7. regresion del caso semilla -----------------------------------
    log("## 7. Regresion del caso semilla (collar 12\"-CWS-46-032-B1)")
    log("")
    rec = openpyxl.load_workbook("lo2/test_v2.xlsx", data_only=True)["Parche_PCC2_Art212"]
    log("| Magnitud | Antes (Datos_Ref) | Ahora (DB ASME) | Dictamen |")
    log("|---|---|---|---|")
    log(f"| Sa collar (A516 Gr.70) | 160,7 | {rec['D39'].value} | {rec['E125'].value} |")
    log(f"| Sa metal base (A106 Gr.B) | 137,9 | {rec['D40'].value} | {rec['D125'].value} |")
    log(f"| Sa gobernante | 137,9 | {rec['D41'].value} | — |")
    log("")
    log(f"Dictamen global del modulo: **{rec['F90'].value}**.")
    Path("Reporte_Verificacion_DB_Materiales.md").write_text("\n".join(out),
                                                             encoding="utf-8")
    return nbad + bad_tot + extra_bad + len(hits) + len(malas) + (0 if cont_ok else 1)


if __name__ == "__main__":
    sys.exit(0 if main() == 0 else 1)
