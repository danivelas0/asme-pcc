# -*- coding: utf-8 -*-
"""verificar.py — Protocolo de aceptacion (seccion 10 del PLAN-DB-MAT-001, Rev. 4).

1. Conteos JSON -> hoja.
2. Unicidad de material_id.
3. Auditoria fila a fila de los valores contra el JSON fuente.
   3b. Bases por grupo, Apendice C y factores de calidad.
4. Contigüidad de los bloques de la cascada (condicion de las listas dependientes).
5. Ausencia de funciones de matriz dinamica y de validaciones no portables.
6. Interpolacion con huecos interiores, modo tabulado y bordes: recalculo real en
   hoja (Excel) contra un motor de referencia independiente en Python.
   6b. Apendice C: bloqueo en los dos extremos y rama de dato puntual.
   6c. Factores de calidad Ec y Ej, y el incremento por examen suplementario.
7. Regresion del caso semilla, leida del propio libro recalculado.
8. Capa de navegacion: visibilidad grabada y proyecto VBA intacto.
9. Mapeo de grupos de propiedades contra las Notas de TM-1 / TE-1.

Las §6b y §6c recalculan la MISMA expresion que lleva el motor: las funciones que
la generan viven en build_db_materiales.py y las emiten los dos. La prueba ejerce
el original, no una copia.

Devuelve 0 solo si todo pasa.

    python verificar.py --resources ..\\..\\..\\resources \\
                        --wb ..\\..\\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_db_materiales as B   # tabla de navegacion: DASH, NAVEGABLES, COL_CLAVE_BASE
from db_lib import Resources, num, temp_to_number, txt

# Se fijan en main() a partir de la linea de comandos. Antes vivian aqui como
# constantes apuntando a /mnt/user-data/uploads/... y a Motor_v2.xlsx, rutas de
# la maquina donde se escribio el script: el protocolo llevaba desde entonces
# sin poder ejecutarse.
RES: Resources = None       # type: ignore[assignment]
WB: str = ""
OUTDIR: Path = None         # type: ignore[assignment]
REPORTE: Path = None        # type: ignore[assignment]
APX = "ASME B31/ASME B31.3/APPEX"
# Archivo CANONICO de la Tabla 302.3.3-1. Su gemelo table_table_302_3_3_1.json
# quedo declarado duplicado por completar_tabla_302_3_3.py; citarlo seria citar
# una fuente ambigua.
TABLA_302331 = "ASME B31/ASME B31.3/CHAPTERS/tables/table_302_3_3_1.json"
R_DATA, R_HDR = 4, 3

XL_XLSX = 51                # xlOpenXMLWorkbook


def recalcular_con_excel(entrada: Path, salida: Path) -> None:
    """Recalcula el libro con Excel y lo guarda con los valores en cache.

    openpyxl no evalua formulas, asi que la unica forma de auditar lo que la
    hoja calcula de verdad -y no lo que creemos que calcula- es pasarla por un
    motor real. Antes se usaba `soffice --convert-to`, que no esta instalado en
    esta maquina y abortaba el protocolo con un FileNotFoundError indistinguible
    de un fallo de verificacion.
    """
    import win32com.client

    if salida.exists():
        salida.unlink()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    try:
        pid = _pid_de(excel)
    except Exception:  # noqa: BLE001
        pid = None
    wb = None
    try:
        wb = excel.Workbooks.Open(str(entrada.resolve()))
        excel.CalculateFullRebuild()
        wb.SaveAs(str(salida.resolve()), FileFormat=XL_XLSX)
    finally:
        for accion in (lambda: wb.Close(SaveChanges=False) if wb is not None else None,
                       excel.Quit):
            try:
                accion()
            except Exception as e:  # noqa: BLE001
                print(f"  aviso al cerrar Excel: {e}", file=sys.stderr)
        _matar(pid)


def _pid_de(excel):
    import win32process
    return win32process.GetWindowThreadProcessId(excel.Hwnd)[1]


def _matar(pid) -> None:
    """Excel via COM puede sobrevivir a Quit y bloquear el archivo."""
    if not pid:
        return
    try:
        import win32api
        import win32con
        import win32event
        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE,
                                 False, pid)
        if win32event.WaitForSingleObject(h, 5000) != win32event.WAIT_OBJECT_0:
            win32api.TerminateProcess(h, 0)
        win32api.CloseHandle(h)
    except Exception:  # noqa: BLE001
        pass
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


def auditar():
    wb = openpyxl.load_workbook(WB)
    # Copia intacta para la seccion 8: `wb` recibe la hoja _QA en la seccion 6
    # y deja de reflejar el entregable.
    wb0 = openpyxl.load_workbook(WB)
    log("# Reporte de verificacion — PLAN-DB-MAT-001 Rev. 3")
    log("")
    log(f"Libro verificado: `{Path(WB).name}`  ·  {len(wb.sheetnames)} hojas")
    log("")

    # ---- 1. conteos -------------------------------------------------------
    log("## 1. Conteo de filas (JSON fuente -> hoja)")
    log("")
    log("| Fuente | Filas JSON | Filas en la hoja | Estado |")
    log("|---|---|---|---|")
    checks = [("A-1 + A-4 -> DB_B31_3",
               len(RES.rows(f"{APX}/appendix_a/table_a_1.json")) +
               len(RES.rows(f"{APX}/appendix_a/table_a_4.json")), "DB_B31_3"),
              ("1A -> DB_BPVC_IID", len(RES.rows("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_1a.json")),
               "DB_BPVC_IID"),
              ("1B + 3 -> DB_BPVC_IID_B",
               len(RES.rows("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_1b.json")) +
               len(RES.rows("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_3.json")), "DB_BPVC_IID_B"),
              ("U -> DB_Su", len(RES.rows("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_u.json")), "DB_Su"),
              ("Y-1 -> DB_Sy", len(RES.rows("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_y_1.json")), "DB_Sy"),
              # Apendice C entero: las cuatro tablas en una sola base por edicion.
              ("C-1 + C-2 + C-3 + C-4 -> DB_B31_C",
               sum(len(RES.rows(f"{APX}/appendix_c/{f}.json"))
                   for f in ("table_c_1", "table_c_2", "table_c_3", "table_c_4")),
               "DB_B31_C"),
              ("C-1C + C-2 + C-3C + C-4 -> DB_B31_CC",
               sum(len(RES.rows(f"{APX}/appendix_c/{f}.json"))
                   for f in ("table_c_1c", "table_c_2", "table_c_3c", "table_c_4")),
               "DB_B31_CC"),
              # Factores de calidad: Ec (A-2), Ej (A-3) y el Ec incrementado.
              ("A-2 -> DB_A2_Ec", len(RES.rows(f"{APX}/appendix_a/table_a_2.json")),
               "DB_A2_Ec"),
              ("A-3 -> DB_A3_Ej", len(RES.rows(f"{APX}/appendix_a/table_a_3.json")),
               "DB_A3_Ej"),
              ("302.3.3-1 -> DB_Ec_Incremento",
               len(RES.rows(TABLA_302331)), "DB_Ec_Incremento")]
    # Las filas descartadas son ruido de extraccion conocido y documentado en
    # ISSUES por el builder, pero no pueden crecer sin que nadie se entere: se
    # tolera hasta MAX_DESCARTE por base y por encima de ahi es un fallo.
    # Antes esta seccion imprimia "OK" pasara lo que pasara y no sumaba al
    # contador, que es como el protocolo se creia verde.
    MAX_DESCARTE = 20
    cnt_bad = 0
    for label, njson, sh in checks:
        n = wb[sh].max_row - R_DATA + 1
        d = njson - n
        if d == 0:
            est = "OK"
        elif 0 < d <= MAX_DESCARTE:
            est = f"OK — {d} filas sin identificacion ni valores (ruido de extraccion)"
        else:
            est = f"FALLO — descarte de {d} filas, por encima del umbral {MAX_DESCARTE}"
            cnt_bad += 1
        log(f"| {label} | {njson} | {n} | {est} |")
    log("")

    # ---- 2. unicidad ------------------------------------------------------
    log("## 2. Unicidad de material_id")
    log("")
    log("| Hoja | Filas | Claves unicas | Estado |")
    log("|---|---|---|---|")
    # DB_B31_C y DB_B31_CC entran aqui sin ningun cambio en el codigo de las
    # secciones 2 y 4: es exactamente lo que compra el contrato de las 8
    # primeras columnas (material_id, Tabla, k0..k4, clave_bi) con STRESS_COLS.
    bases = ["DB_B31_3", "DB_B31_3C", "DB_BPVC_IID", "DB_BPVC_IIDC",
             "DB_BPVC_IID_B", "DB_BPVC_IID_BC", "DB_Su", "DB_Sy",
             "DB_B31_C", "DB_B31_CC", "DB_A2_Ec", "DB_A3_Ej"]
    # material_id es la clave con la que el motor localiza cada material: un
    # duplicado significa que el motor puede tomar el admisible equivocado.
    # Es un fallo, no una nota informativa como estaba escrito.
    uniq_bad = 0
    for sh in bases:
        ws = wb[sh]
        ids = [ws.cell(r, 1).value for r in range(R_DATA, ws.max_row + 1)]
        ok = len(ids) == len(set(ids))
        uniq_bad += 0 if ok else 1
        log(f"| {sh} | {len(ids)} | {len(set(ids))} | "
            f"{'OK' if ok else 'FALLO — DUPLICADOS'} |")
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
              ("DB_BPVC_IID", [("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_1a.json", None)]),
              ("DB_BPVC_IID_B", [("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_1b.json", None),
                                 ("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_3.json", None)]),
              ("DB_Su", [("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_u.json", None)]),
              ("DB_Sy", [("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/table_y_1.json", None)]),
              ("DB_B31_3C", [(f"{APX}/appendix_a/table_a_1c.json", 100),
                             (f"{APX}/appendix_a/table_a_4c.json", 100)]),
              ("DB_BPVC_IIDC", [("ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025/table_1a.json", None)]),
              ("DB_BPVC_IID_BC", [("ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025/table_1b.json", None),
                                  ("ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025/table_3.json", None)]),
              ("DB_SuC", [("ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025/table_u.json", None)]),
              ("DB_SyC", [("ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025/table_y_1.json", None)])]
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

    def audita_apendice_c(sh, si):
        """Auditoria fila a fila de las 4 tablas del Apendice C contra la hoja.

        No es un conteo: de cada fila del JSON se localiza su gemela en la hoja
        por (Tabla, Linea impresa) y se comparan el NOMBRE del material y el
        vector completo de valores —o el valor unico y su rango de validez en
        las tablas de dato puntual—. Un desplazamiento de una fila, que es el
        modo tipico de romper un enlace posicional, salta aqui.
        """
        w = wb[sh]
        cm = col_map(w)
        temps = printed_temps(w)
        ni = n_ident(w)
        idx = {}
        for r in range(R_DATA, w.max_row + 1):
            k = (txt(w.cell(r, cm["Tabla"]).value), w.cell(r, cm["Linea"]).value)
            idx[k] = r

        def _nm(s):
            return re.sub(r"\s+", "", txt(s)).upper()

        ap = f"{APX}/appendix_c"
        tablas = [
            ("C-1" if si else "C-1C", f"{ap}/table_c_1{'' if si else 'c'}.json",
             "material_uns_no", None, None),
            ("C-2", f"{ap}/table_c_2.json", "material_description",
             "mm_mm_c" if si else "in_in_f", "range_c" if si else "range_f"),
            ("C-3" if si else "C-3C", f"{ap}/table_c_3{'' if si else 'c'}.json",
             "material", None, None),
            ("C-4", f"{ap}/table_c_4.json", "material_description",
             "e_mpa_23_c" if si else "e_ksi_73_4_f", None),
        ]
        njson = nval = mal = 0
        for tag, rel, kmat, kval, krng in tablas:
            for i, row in enumerate(RES.rows(rel), start=1):
                njson += 1
                r = idx.get((tag, i))
                if r is None:
                    mal += 1
                    continue
                # El nombre se compara plegando espacios y guiones: la propia
                # extraccion duplica el rotulo en 21 de las 26 filas de C-1 y el
                # builder lo colapsa; eso es un artefacto declarado, no un valor.
                esperado = _nm(row.get(kmat))
                visto = _nm(w.cell(r, cm["Material"]).value)
                if esperado != visto and not esperado.startswith(visto):
                    mal += 1
                if kval is None:                       # tabla de curva
                    for k, v in (row.get("values") or {}).items():
                        if num(v) is None:
                            continue
                        t = temp_to_number(k)
                        j = temps.index(t) if t in temps else None
                        nval += 1
                        if j is None or w.cell(r, ni + 1 + j).value != num(v):
                            mal += 1
                else:                                   # tabla de dato puntual
                    v = row.get(kval)
                    nval += 1
                    if isinstance(v, str):
                        if txt(w.cell(r, cm["Valor unico (texto)"]).value) != txt(v):
                            mal += 1
                    elif v is not None:
                        if w.cell(r, cm["Valor unico"]).value != v:
                            mal += 1
                    if krng:
                        rg = txt(w.cell(r, cm["Rango de validez"]).value)
                        imp = txt(row.get(krng))
                        nval += 1
                        if imp and not rg.startswith(imp):
                            mal += 1
                        if not imp and rg:
                            mal += 1
        log(f"| {sh} | {njson} | {w.max_row - R_DATA + 1} | {nval} | {mal} |")
        return mal

    def audita_grupo_te(sh, rels):
        """DB_TE_G / DB_TE_GC: compara, por GRUPO/columna, el vector del
        Coeficiente B contra TE-1..5. Usa la MISMA clasificacion de columnas
        que el builder (`B.etiqueta_columna_b_te1` / `B._clave_fila_te1`): si
        alguien cambia esa logica, esta auditoria la ejerce cambiada, en vez
        de poder divergir con una copia propia."""
        from collections import Counter
        cj, npts = Counter(), 0
        for rel in rels:
            d = RES.load(rel)
            cols = d.get("columns", [])
            if not cols:
                continue
            temp_key = B._clave_fila_te1(cols[0])
            clasif = {}
            for col in cols[1:]:
                etq = B.etiqueta_columna_b_te1(col)
                if etq:
                    clasif[B._clave_fila_te1(col)] = etq
            grupos: dict = {}
            for row in d["rows"]:
                tnum = temp_to_number(row.get(temp_key))
                if tnum is None:
                    continue
                for col_key, etq in clasif.items():
                    v = num(row.get(col_key))
                    if v is not None:
                        grupos.setdefault(etq, {})[tnum] = v
            for vals in grupos.values():
                npts += len(vals)
                cj[tuple(sorted(vals.items()))] += 1
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
        log(f"| {sh} | {sum(cj.values())} | {w.max_row - R_DATA + 1} | {npts} | {falt} |")
        return falt

    ed = "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025"
    edc = "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"
    extra_bad = 0
    extra_bad += audita_grupo("DB_E", [f"{ed}/table_tm_{k}.json" for k in range(1, 6)], 4)
    extra_bad += audita_grupo("DB_EC", [f"{edc}/table_tm_{k}.json" for k in range(1, 6)], 4)
    extra_bad += audita_grupo_te("DB_TE_G", [f"{ed}/table_te_{k}.json" for k in range(1, 6)])
    extra_bad += audita_grupo_te("DB_TE_GC", [f"{edc}/table_te_{k}.json" for k in range(1, 6)])
    extra_bad += audita_apendice_c("DB_B31_C", True)
    extra_bad += audita_apendice_c("DB_B31_CC", False)

    def audita_factores(sh, rel, campo, con_clase):
        """Fila a fila: factor, notas citadas y descripcion contra el JSON.

        Ec y Ej multiplican directamente el esfuerzo admisible: un factor
        desplazado una fila cambia el espesor requerido sin que nada falle.
        """
        w = wb[sh]
        cm = col_map(w)
        idx = {w.cell(r, cm["Linea"]).value: r
               for r in range(R_DATA, w.max_row + 1)}
        njson = nval = mal = 0
        for i, row in enumerate(RES.rows(rel), start=1):
            njson += 1
            r = idx.get(i)
            if r is None:
                mal += 1
                continue
            comparaciones = [(row.get(campo), w.cell(r, cm["Factor"]).value),
                             (txt(row.get("notes")),
                              txt(w.cell(r, cm["Notas citadas"]).value)),
                             (txt(row.get("description")),
                              txt(w.cell(r, cm["Descripcion"]).value)),
                             (txt(row.get("spec_no")),
                              txt(w.cell(r, cm["Spec. No."]).value))]
            if con_clase:
                comparaciones.append((txt(row.get("class_or_type")),
                                      txt(w.cell(r, cm["Clase o tipo"]).value)))
            for esperado, visto in comparaciones:
                nval += 1
                if esperado != visto:
                    mal += 1
        log(f"| {sh} | {njson} | {w.max_row - R_DATA + 1} | {nval} | {mal} |")
        return mal

    extra_bad += audita_factores("DB_A2_Ec", f"{APX}/appendix_a/table_a_2.json",
                                 "ec", False)
    extra_bad += audita_factores("DB_A3_Ej", f"{APX}/appendix_a/table_a_3.json",
                                 "ej", True)

    # Tabla 302.3.3-1: los seis examenes y su Ec, contra el archivo CANONICO.
    # Las columnas se localizan por posicion porque sus rotulos no son impresos
    # sino derivados del para. 302.3.3(c) (el impreso no se capturo).
    wec = wb["DB_Ec_Incremento"]
    filas_ec = RES.rows(TABLA_302331)
    mal_ec = 0
    for i, row in enumerate(filas_ec, start=1):
        r = R_DATA + i - 1
        if txt(wec.cell(r, 4).value) != txt(row.get("column_1")):
            mal_ec += 1
        if wec.cell(r, 5).value != row.get("column_2"):
            mal_ec += 1
    log(f"| DB_Ec_Incremento | {len(filas_ec)} | {wec.max_row - R_DATA + 1} | "
        f"{2 * len(filas_ec)} | {mal_ec} |")
    extra_bad += mal_ec
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
    # Solo Apendice B: C-2 y C-4 salieron de esta hoja a DB_B31_C / DB_B31_CC.
    for rel in [f"{APX}/appendix_b/table_b_{k}.json" for k in
                ("1", "1c", "2", "3", "4", "5", "6")]:
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
    # La rejilla de referencia se resuelve por hoja, no con un ternario de dos
    # ramas: cada base nueva que entre en `cases` necesita la suya.
    grids = {"DB_B31_3": (tb, ib), "DB_BPVC_IID": (ti, ii)}

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

    # ---- 6b. Apendice C: bloqueo en los dos extremos y rama de dato puntual --
    # Se recalcula la MISMA expresion que lleva el motor —la emiten las funciones
    # formula_estado_apxc / formula_valor_apxc de build_db_materiales— resolviendo
    # la fila por material_id en vez de por la cascada. Asi la prueba ejerce el
    # original y no una copia de la logica.
    def _id_de(sh, pref):
        w = wb[sh]
        for r in range(R_DATA, w.max_row + 1):
            v = w.cell(r, 1).value
            if isinstance(v, str) and v.startswith(pref):
                return v
        raise SystemExit(f"{sh}: no existe ninguna fila que empiece por {pref!r}")

    CS = "Carbon steels with carbon content 0.30% or less"
    G1 = "Group 1 carbon and low alloy steels"
    casos_c = [
        # (hoja, material_id, T, modo)  — los cuatro primeros son el mismo
        # material a 4 temperaturas: punto exacto, interpolacion y los dos bordes.
        ("DB_B31_C", _id_de("DB_B31_C", f"C-3 | {CS}"), 25, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", f"C-3 | {CS}"), 375, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", f"C-3 | {CS}"), 700, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", f"C-3 | {CS}"), -300, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", f"C-1 | {G1}"), 400, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", f"C-1 | {G1}"), 412, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", "C-4 | Acetal"), 25, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", "C-4 | Acetal"), 300, "Interpolado"),
        ("DB_B31_C", _id_de("DB_B31_C", "C-2 | Glass-epoxy, filament-wound"), 25,
         "Interpolado"),
        # El mismo material en la edicion US: se lee su tabla nativa (C-3C), no
        # una conversion de la metrica.
        ("DB_B31_CC", _id_de("DB_B31_CC", f"C-3C | {CS}"), 662, "Interpolado"),
    ]
    qc = wb.create_sheet("_QA_C")
    for j, h in enumerate(["hoja", "material_id", "T", "modo", "fila", "T1", "V1",
                           "n_pts", "T2", "V2", "tipo", "T prim", "T ult", "factor",
                           "valor unico", "valor texto", "V(T)", "estado",
                           "valor aplicado"], 1):
        qc.cell(2, j, h)
    for i, (sh, mid, T, modo) in enumerate(casos_c):
        r = 3 + i
        w = wb[sh]
        cm = col_map(w)
        ni = n_ident(w)
        npr = len(printed_temps(w))
        t0 = ni + npr + 1
        npack = (w.max_column - t0 + 1) // 2
        v0 = t0 + npack

        def col(nombre, _w=w, _cm=cm, _sh=sh):
            L = get_column_letter(_cm[nombre])
            return f"{_sh}!${L}${R_DATA}:${L}${_w.max_row}"

        TA = f"{sh}!${get_column_letter(t0)}${R_DATA}"
        VA = f"{sh}!${get_column_letter(v0)}${R_DATA}"
        IDS = f"{sh}!$A${R_DATA}:$A${w.max_row}"
        qc.cell(r, 1, sh); qc.cell(r, 2, mid); qc.cell(r, 3, T); qc.cell(r, 4, modo)
        qc.cell(r, 5).value = f'=IFERROR(MATCH($B{r},{IDS},0),"")'
        FIL = f"$E{r}"
        NP = f"IFERROR(INDEX({col('n_pts')},{FIL}),0)"
        tr = f"OFFSET({TA},{FIL}-1,0,1,MAX(1,{NP}))"
        vr = f"OFFSET({VA},{FIL}-1,0,1,MAX(1,{NP}))"
        p1 = f"IFERROR(MATCH($C{r},{tr},1),1)"
        # Mismo envoltorio que el motor: INDEX sobre celda vacia devuelve 0, y
        # sin el las filas de tipo PUNTO entrarian con ceros inventados.
        def vac(expr):
            return f'IF({expr}="","",{expr})'

        for j, expr in ((6, f"INDEX({tr},{p1})"), (7, f"INDEX({vr},{p1})"),
                        (9, f"INDEX({tr},{p1}+1)"), (10, f"INDEX({vr},{p1}+1)")):
            qc.cell(r, j).value = f'=IFERROR({vac(expr)},"")'
        qc.cell(r, 8).value = f"={NP}"
        for j, nombre in ((11, "Tipo de dato"), (12, "T primera tabulada"),
                          (13, "T ultima tabulada"), (14, "Factor de escala"),
                          (15, "Valor unico"), (16, "Valor unico (texto)")):
            qc.cell(r, j).value = f'=IFERROR({vac(f"INDEX({col(nombre)},{FIL})")},"")'
        qc.cell(r, 17).value = (
            f'=IF($G{r}="","",IF($C{r}<=$F{r},$G{r},'
            f'IF(OR($I{r}="",$J{r}=""),$G{r},'
            f'IF($D{r}="Tabulado-conservador",$J{r},'
            f'$G{r}+($J{r}-$G{r})*($C{r}-$F{r})/($I{r}-$F{r})))))')
        qc.cell(r, 18).value = B.formula_estado_apxc(
            FIL, f"$K{r}", f"$H{r}", f"$L{r}", f"$M{r}", f"$C{r}")
        qc.cell(r, 19).value = B.formula_valor_apxc(
            FIL, f"$K{r}", f"$O{r}", f"$P{r}", f"$N{r}", f"$R{r}", f"$Q{r}")

    # ---- 6c. los dos motores de factores de calidad ------------------------
    # Aqui no hay interpolacion que recalcular: Ec y Ej son escalares. Lo que se
    # comprueba es que el motor muestra EXACTAMENTE el factor de la fila —sin
    # redondeo— y que el bloque de incremento aplica el de la Tabla 302.3.3-1
    # solo donde la fila lo admite. Se conduce escribiendo la cascada, igual que
    # el caso semilla del motor de calculo.
    def _fila_de(sh, **campos):
        """Fila de una base de factores que casa con todos los campos dados."""
        w = wb[sh]
        cm = col_map(w)
        for r in range(R_DATA, w.max_row + 1):
            if all(txt(w.cell(r, cm[k]).value) == txt(v) for k, v in campos.items()):
                return r, cm, w
        raise SystemExit(f"{sh}: ninguna fila casa con {campos}")

    EXA_MAX = "(1) and (3)(a) or (3)(b)"        # la que habilita Ec = 1,00
    # Cada motor solo puede conducirse una vez por recalculo: la cascada vive en
    # celdas concretas de su hoja. Asi que el barrido de casos se hace sobre una
    # hoja _QA_F que reemite la MISMA expresion del motor resolviendo la fila por
    # material_id, y ademas se conduce un caso vivo por motor a traves de la
    # cascada real, que es lo que comprueba el cableado de las listas.
    wi = wb["DB_Ec_Incremento"]
    inc_tab = {txt(wi.cell(rr, 4).value): wi.cell(rr, 5).value
               for rr in range(R_DATA, wi.max_row + 1)}
    casos_f = [
        ("DB_A2_Ec", {"Spec. No.": "A395",
                      "Descripcion": "Ductile and ferritic ductile iron castings"},
         EXA_MAX),
        ("DB_A2_Ec", {"Spec. No.": "A451", "Grupo impreso": "Stainless Steel"}, "(1)"),
        ("DB_A2_Ec", {"Spec. No.": "A451", "Grupo impreso": "Stainless Steel"},
         EXA_MAX),
        ("DB_A2_Ec", {"Spec. No.": "A426"}, EXA_MAX),
        ("DB_A2_Ec", {"Spec. No.": "A47"}, EXA_MAX),
        ("DB_A3_Ej", {"Spec. No.": "API 5L",
                      "Descripcion": "Continuous welded (furnace butt welded) pipe"},
         None),
        ("DB_A3_Ej", {"Spec. No.": "API 5L", "Descripcion": "Seamless pipe"}, None),
        ("DB_A3_Ej", {"Spec. No.": "A312",
                      "Descripcion": "Electric fusion welded pipe, single butt seam"},
         None),
        ("DB_A3_Ej", {"Spec. No.": "A312",
                      "Descripcion": "Electric fusion welded pipe, double butt seam"},
         None),
    ]
    qf = wb.create_sheet("_QA_F")
    for j, h in enumerate(["hoja", "material_id", "examen", "fila", "factor basico",
                           "admite", "factor del examen", "factor aplicable"], 1):
        qf.cell(2, j, h)
    esperado_f = []
    EXA = (f"DB_Ec_Incremento!$D${R_DATA}:"
           f"$D${wb['DB_Ec_Incremento'].max_row}")
    FAC = (f"DB_Ec_Incremento!$E${R_DATA}:"
           f"$E${wb['DB_Ec_Incremento'].max_row}")
    for i, (base, filtros, examen) in enumerate(casos_f):
        r0, cm, w = _fila_de(base, **filtros)
        mid = w.cell(r0, 1).value
        r = 3 + i

        def col(nombre, _w=w, _cm=cm, _b=base):
            L = get_column_letter(_cm[nombre])
            return f"{_b}!${L}${R_DATA}:${L}${_w.max_row}"

        qf.cell(r, 1, base); qf.cell(r, 2, mid); qf.cell(r, 3, examen or "")
        qf.cell(r, 4).value = (f'=IFERROR(MATCH($B{r},{base}!$A${R_DATA}:'
                               f'$A${w.max_row},0),"")')
        FIL = f"$D{r}"
        qf.cell(r, 5).value = f'=IF({FIL}="","",INDEX({col("Factor")},{FIL}))'
        qf.cell(r, 6).value = (f'=IF({FIL}="","",'
                               f'INDEX({col("Admite incremento")},{FIL}))')
        qf.cell(r, 7).value = (f'=IF($C{r}="","",'
                               f'IFERROR(INDEX({FAC},MATCH($C{r},{EXA},0)),""))')
        qf.cell(r, 8).value = B.formula_factor_aplicable(
            FIL, f"$F{r}", f"$E{r}", f"$G{r}")
        basico = w.cell(r0, cm["Factor"]).value
        admite = str(w.cell(r0, cm["Admite incremento"]).value or "")
        if examen is None:
            aplicable = basico
        else:
            aplicable = (min(1, max(basico, inc_tab[txt(examen)]))
                         if admite.startswith("SI") else basico)
        esperado_f.append((base, filtros, examen, basico, admite, aplicable))

    # Un caso vivo por motor, conducido por la cascada real. La cascada arranca
    # en D6 y ocupa un nivel por fila; el resto del layout se deriva de ahi.
    vivos = [("Buscar_Ec_A2", "DB_A2_Ec",
              {"Spec. No.": "A395",
               "Descripcion": "Ductile and ferritic ductile iron castings"},
              EXA_MAX, ("Grupo impreso", "Spec. No.", "Descripcion")),
             ("Buscar_Ej_A3", "DB_A3_Ej",
              {"Spec. No.": "API 5L",
               "Descripcion": "Continuous welded (furnace butt welded) pipe"},
              None, ("Grupo impreso", "Spec. No.", "Clase o tipo", "Descripcion"))]
    esperado_vivo = []
    for motor, base, filtros, examen, niveles in vivos:
        r0, cm, w = _fila_de(base, **filtros)
        ws_m = wb[motor]
        for k, nombre in enumerate(niveles):
            v = w.cell(r0, cm[nombre]).value
            ws_m.cell(6 + k, 4).value = v if v not in (None, "") else B.NO_APLICA
        r_ult = 5 + len(niveles)
        fila_val = r_ult + 4          # banda 2 en r_ult+2; valores dos filas mas
        fila_apl = r_ult + 9          # banda 3 en r_ult+7; examen +1; campos +2
        if examen is not None:
            ws_m.cell(fila_apl - 1, 4).value = examen
        basico = w.cell(r0, cm["Factor"]).value
        admite = str(w.cell(r0, cm["Admite incremento"]).value or "")
        aplicable = (None if examen is None else
                     (min(1, max(basico, inc_tab[txt(examen)]))
                      if admite.startswith("SI") else basico))
        esperado_vivo.append((motor, filtros, fila_val, fila_apl,
                              basico, admite, aplicable))

    # Caso semilla de la seccion 7. Se carga aqui, antes del unico recalculo,
    # para no abrir Excel dos veces.
    #
    # El motor se ENTREGA sin material seleccionado: D109..D114 y E109..E114
    # estan vacias y D115 resuelve a "". Los campos D22/D23 son descripciones
    # de texto libre, no la seleccion de la base. Asi que el caso de regresion
    # hay que conducirlo: se escribe el material_id en la celda "Variante"
    # (D114/E114), que el motor respeta por encima de la cascada
    # (`=IF($D$114<>"",$D$114,...)`). Eso fija la resolucion sin depender de
    # los cinco niveles de listas desplegables.
    motor = wb["Parche_PCC2_Art212"]
    id_base = next(i for i in ib if i.startswith("A-1 | A106 | B"))
    id_collar = next(i for i in ib if i.startswith("A-1 | A516 | 70"))
    motor["D114"] = id_base
    motor["E114"] = id_collar
    T_semilla = motor["D25"].value

    qa_in = OUTDIR / "qa.xlsx"
    qa_out = OUTDIR / "qa_recalculado.xlsx"
    wb.save(qa_in)
    recalcular_con_excel(qa_in, qa_out)
    # Un unico recalculo sirve a la seccion 6 y a la 7: el mismo libro lleva la
    # hoja _QA y el motor con el caso semilla ya cargado.
    recalc = openpyxl.load_workbook(qa_out, data_only=True)
    rb = recalc["_QA"]
    log("| Base | material_id | T | Modo | T1 | T2 | S(T) hoja | S(T) referencia | Estado |")
    log("|---|---|---|---|---|---|---|---|---|")
    nbad = 0
    for i, (sh, mid, T, modo) in enumerate(cases):
        r = 3 + i
        got = rb.cell(r, 11).value
        temps, idx = grids[sh]
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

    # ---- 6b. Apendice C: los dos extremos y el dato puntual ---------------
    log("### 6b. Apendice C — bloqueo en los dos extremos y rama de dato puntual")
    log("")
    log("El Apendice C no publica columna «Temp. max.»: el limite es el primer y el "
        "ultimo punto que tabula la propia fila. Estos casos comprueban, recalculando "
        "en Excel la misma expresion que lleva el motor, que por encima y por debajo "
        "de esa banda el resultado queda BLOQUEADO en vez de sostener el valor del "
        "extremo, y que C-2 y C-4 se resuelven por su valor unico impreso.")
    log("")
    rc = recalc["_QA_C"]
    log("| Hoja | material_id | T | Tipo | Estado hoja | Estado referencia | "
        "Valor hoja | Valor referencia | Estado |")
    log("|---|---|---|---|---|---|---|---|---|")
    nbad_c = 0
    for i, (sh, mid, T, modo) in enumerate(casos_c):
        r = 3 + i
        w = wb[sh]
        cm = col_map(w)
        temps_p = printed_temps(w)
        ni = n_ident(w)
        fila = next(rr for rr in range(R_DATA, w.max_row + 1)
                    if w.cell(rr, 1).value == mid)
        tipo = w.cell(fila, cm["Tipo de dato"]).value
        tprim = w.cell(fila, cm["T primera tabulada"]).value
        tult = w.cell(fila, cm["T ultima tabulada"]).value
        fac = w.cell(fila, cm["Factor de escala"]).value
        vu = w.cell(fila, cm["Valor unico"]).value
        vt = w.cell(fila, cm["Valor unico (texto)"]).value
        npts = w.cell(fila, cm["n_pts"]).value or 0
        vals = {}
        for k, t in enumerate(temps_p):
            v = w.cell(fila, ni + 1 + k).value
            if v is not None:
                vals[t] = v
        if tipo == "PUNTO":
            est_ref = B.EST_PUNTO
        elif npts == 0:
            est_ref = B.EST_SIN_TAB
        elif T < tprim:
            est_ref = B.EST_BAJO
        elif T > tult:
            est_ref = B.EST_ALTO
        else:
            est_ref = B.EST_OK
        if tipo == "PUNTO":
            val_ref = vu * fac if isinstance(vu, (int, float)) else (vt or "")
        elif est_ref.startswith("FUERA DE RANGO"):
            val_ref = B.VAL_BLOQUEADO
        elif est_ref == B.EST_SIN_TAB:
            val_ref = B.EST_SIN_TAB
        else:
            v = interp(temps_p, vals, T, modo)
            val_ref = v * fac if v is not None else ""
        got_e, got_v = rc.cell(r, 18).value, rc.cell(r, 19).value
        if isinstance(got_v, (int, float)) and isinstance(val_ref, (int, float)):
            ok_v = abs(got_v - val_ref) <= abs(val_ref) * 1e-9 + 1e-9
        else:
            ok_v = got_v == val_ref
        ok = (got_e == est_ref) and ok_v
        nbad_c += 0 if ok else 1
        corto = lambda s: str(s)[:34] if s is not None else "—"
        log(f"| {sh} | `{str(mid)[:34]}` | {T} | {tipo} | {corto(got_e)} | "
            f"{corto(est_ref)} | {corto(got_v)} | {corto(val_ref)} | "
            f"{'OK' if ok else 'FALLO'} |")
    log("")
    log(f"**{len(casos_c)} casos del Apendice C, {nbad_c} fallos.**")
    log("")
    nbad += nbad_c

    # ---- 6c. factores de calidad Ec y Ej ----------------------------------
    log("### 6c. Factores de calidad — el motor muestra el factor de la fila, sin "
        "redondeo")
    log("")
    log("Ec y Ej son escalares: no hay interpolacion que recalcular. Lo que se "
        "comprueba, conduciendo la cascada y recalculando en Excel, es que el KPI "
        "es EXACTAMENTE el valor impreso y que el incremento de la Tabla 302.3.3-1 "
        "solo se aplica donde la fila lo admite.")
    log("")
    log("| Base | Fila | Examen | Factor hoja | Factor codigo | Admite | "
        "Aplicable hoja | Aplicable referencia | Estado |")
    log("|---|---|---|---|---|---|---|---|---|")
    rf = recalc["_QA_F"]
    nbad_f = 0
    for i, (base, filtros, examen, basico, admite, aplicable) in enumerate(esperado_f):
        r = 3 + i
        got_b, got_a, got_ap = (rf.cell(r, 5).value, rf.cell(r, 6).value,
                                rf.cell(r, 8).value)
        ok = (got_b == basico and txt(got_a) == txt(admite)
              and isinstance(got_ap, (int, float))
              and abs(got_ap - aplicable) <= 1e-12)
        nbad_f += 0 if ok else 1
        etiq = " · ".join(f"{k}={v}" for k, v in filtros.items())
        log(f"| {base} | {etiq[:46]} | {examen or '—'} | {got_b} | {basico} | "
            f"{str(got_a)[:20]} | {got_ap} | {aplicable} | "
            f"{'OK' if ok else 'FALLO'} |")
    log("")
    log("Y el mismo calculo conducido por la cascada real de cada motor, que es lo "
        "que comprueba el cableado de las listas desplegables:")
    log("")
    log("| Motor | Fila | Factor hoja | Factor codigo | Admite hoja | "
        "Aplicable hoja | Aplicable referencia | Estado |")
    log("|---|---|---|---|---|---|---|---|")
    for motor, filtros, fila_val, fila_apl, basico, admite, aplicable in esperado_vivo:
        rm = recalc[motor]
        got_b = rm.cell(fila_val, 1).value
        got_a = rm.cell(fila_val, 10).value
        got_ap = rm.cell(fila_apl, 3).value if aplicable is not None else None
        ok = got_b == basico and txt(got_a) == txt(admite)
        if aplicable is not None:
            ok = ok and isinstance(got_ap, (int, float)) and \
                abs(got_ap - aplicable) <= 1e-12
        nbad_f += 0 if ok else 1
        etiq = " · ".join(f"{k}={v}" for k, v in filtros.items())
        log(f"| {motor} | {etiq[:46]} | {got_b} | {basico} | {str(got_a)[:20]} | "
            f"{got_ap} | {aplicable} | {'OK' if ok else 'FALLO'} |")
    log("")
    log(f"**{len(esperado_f) + len(esperado_vivo)} casos de factores, "
        f"{nbad_f} fallos.**")
    log("")
    nbad += nbad_f

    # ---- 7. regresion del caso semilla -----------------------------------
    # Antes se leia de lo2/test_v2.xlsx, un archivo que no esta en el repo, y
    # la seccion no comparaba nada: imprimia 160,7 y 137,9 como literales
    # dentro del f-string, asi que no podia fallar. Ahora lee el libro que
    # acaba de recalcular Excel y compara de verdad.
    #
    # Los valores de referencia son los IMPRESOS en B31.3-2024 Tabla A-1 a la
    # temperatura de evaluacion, y se recalculan aqui con el motor de
    # referencia en Python -no se copian a mano-. Los 160,7 / 137,9 de la nota
    # de version son los valores antiguos de Datos_Ref, la lista corta
    # obsoleta; la base ASME imprime 161 y 138.
    log("## 7. Regresion del caso semilla (collar 12\"-CWS-46-032-B1)")
    log("")
    rec = recalc["Parche_PCC2_Art212"]
    TOL = 1e-6
    ref_base = interp(tb, ib[id_base][1], T_semilla, "Interpolado")
    ref_collar = interp(tb, ib[id_collar][1], T_semilla, "Interpolado")
    esperado = [("Sa collar (A516 Gr.70)", "D39", ref_collar, "E125"),
                ("Sa metal base (A106 Gr.B)", "D40", ref_base, "D125"),
                ("Sa gobernante", "D41", min(ref_base, ref_collar), None)]
    semilla_bad = 0
    log(f"Temperatura de evaluacion: **{T_semilla} °C** · metal base `{id_base[:40]}` · "
        f"collar `{id_collar[:40]}`")
    log("")
    log("| Magnitud | Referencia Python (MPa) | Hoja recalculada (MPa) | Dictamen | Estado |")
    log("|---|---|---|---|---|")
    for etiqueta, celda, ref, celda_dict in esperado:
        got = rec[celda].value
        ok = isinstance(got, (int, float)) and abs(got - ref) <= TOL
        semilla_bad += 0 if ok else 1
        dictamen = rec[celda_dict].value if celda_dict else "—"
        log(f"| {etiqueta} | {ref} | {got} | {dictamen} | {'OK' if ok else 'FALLO'} |")
    dict_global = rec["F90"].value
    ok_global = dict_global == "APTO"
    semilla_bad += 0 if ok_global else 1
    log("")
    log(f"Dictamen global del modulo: **{dict_global}** "
        f"({'OK' if ok_global else 'FALLO — se esperaba APTO'}).")
    log("")

    # ---- 8. capa de navegacion -------------------------------------------
    log("## 8. Capa de navegacion (Dashboard y proyecto VBA)")
    log("")
    nav_bad = 0
    ruta = Path(WB)

    estados = {s.title: s.sheet_state for s in wb0.worksheets}
    visibles = [n for n, e in estados.items() if e == "visible"]
    ocultas = {n for n, e in estados.items() if e == "hidden"}
    very = {n for n, e in estados.items() if e == "veryHidden"}
    esperadas_very = set(estados) - {B.DASH} - set(B.NAVEGABLES)

    filas = [
        ("Unica hoja visible es el Dashboard", visibles == [B.DASH], ", ".join(visibles)),
        (f"Las {len(B.NAVEGABLES)} hojas navegables estan hidden",
         ocultas == set(B.NAVEGABLES), f"{len(ocultas)} hojas"),
        ("El resto esta veryHidden", very == esperadas_very, f"{len(very)} hojas"),
        ("Ninguna base de datos alcanzable desde la UI",
         not [n for n in estados if estados[n] != "veryHidden"
              and re.match(r"^(DB_|MAP_|Notas_Codigo|Datos_Ref|_)", n)], ""),
        ("El paquete conserva xl/vbaProject.bin",
         "xl/vbaProject.bin" in zipfile.ZipFile(ruta).namelist(), ruta.suffix),
    ]
    # Las claves de destino de los botones deben apuntar a hojas reales.
    dash = wb0[B.DASH]
    claves = [dash.cell(c.row, B.COL_CLAVE_BASE + c.column).value
              for row in dash.iter_rows() for c in row if c.hyperlink is not None]
    filas.append((f"Los botones cubren las {len(B.NAVEGABLES)} hojas navegables",
                  sorted(k for k in claves if k) == sorted(B.NAVEGABLES),
                  f"{len(claves)} botones"))
    filas.append(("Cada hoja navegable tiene enlace de retorno",
                  all(any(wb0[n].cell(c.row, B.COL_CLAVE_BASE + c.column).value
                          == B.CLAVE_VOLVER
                          for row in wb0[n].iter_rows() for c in row
                          if c.hyperlink is not None)
                      for n in B.NAVEGABLES), ""))

    log("| Comprobacion | Detalle | Estado |")
    log("|---|---|---|")
    for etiqueta, ok, detalle in filas:
        nav_bad += 0 if ok else 1
        log(f"| {etiqueta} | {detalle} | {'OK' if ok else 'FALLO'} |")
    log("")
    log("La visibilidad esta grabada en el archivo, no la impone la macro: con las "
        "macros bloqueadas el usuario sigue sin ver ninguna base de datos.")
    log("")

    # ---- 9. mapeo de grupos de propiedades --------------------------------
    # Cada fila de MAP_Grupo que declara un grupo tiene que citar la Nota del
    # codigo que lo sostiene, y esa Nota tiene que listar literalmente esa
    # composicion. Es la comprobacion que convierte el mapeo en auditable: sin
    # ella volveriamos a tener grupos asignados sin respaldo, que es justo el
    # defecto que la Rev. 3 elimino.
    log("## 9. Mapeo de grupos de propiedades "
        "(MAP_Grupo y MAP_GrupoC -> Notas de TM-1 / TE-1)")
    log("")
    log("Cada hoja se audita contra las Notas de SU edicion: no numeran igual, "
        "asi que cruzarlas ocultaria una cita mal puesta.")
    log("")
    map_bad = 0

    def _ck(s):
        return re.sub(r"\s+", "", str(s or "")).translate(
            dict.fromkeys(map(ord, "‐‑‒–—―−⁃"), "-")
        ).upper()

    # Indice de respaldo, leido de resources/ (no del libro). Cada hoja se
    # audita contra las Notas de SU edicion: no numeran igual, asi que cruzarlas
    # daria falsos fallos y, peor, ocultaria una cita mal puesta.
    def respaldo_de(ed):
        out = {}
        for archivo, tabla in (("table_tm_1.json", "TM-1"),
                               ("table_te_1.json", "TE-1")):
            for nota in RES.load(f"{ed}/{archivo}").get("note_members", []):
                if nota.get("tipo") != "grupo":
                    continue
                for m in nota["miembros"]:
                    out.setdefault(_ck(m), set()).add(
                        (tabla, nota["nota"], nota["grupo"]))
        return out

    estados = Counter()
    sin_cita = huerfanas = validado_sin_marca = prestada_sin_origen = 0
    for hoja, ed in (("MAP_Grupo", "ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025"),
                     ("MAP_GrupoC", "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025")):
        ws_map = wb0[hoja]
        cm = col_map(ws_map)
        respaldo = respaldo_de(ed)
        for r in range(R_DATA, ws_map.max_row + 1):
            est = str(ws_map.cell(r, cm["Estado"]).value or "")
            if not est:
                continue
            estados[est] += 1
            comp = _ck(ws_map.cell(r, cm["Composicion nominal"]).value)
            motivo = str(ws_map.cell(r, cm["Motivo"]).value or "")
            for col_g, col_f, tabla in (("Grupo E (TM)", "Fuente E", "TM-1"),
                                        ("Grupo dilatacion (TE)", "Fuente alfa", "TE-1")):
                grupo = ws_map.cell(r, cm[col_g]).value
                fuente = str(ws_map.cell(r, cm[col_f]).value or "")
                if not grupo:
                    continue
                if not fuente:
                    sin_cita += 1           # grupo sin fuente: prohibido
                    continue
                # (la marca de decision se comprueba por FILA, mas abajo)
                # Una composicion tomada de otra tabla tiene que declarar su origen.
                if "otra tabla" in est and "comp. de " not in fuente:
                    prestada_sin_origen += 1
                if "Nota" not in fuente or "comp. de " in fuente or "Validado" in fuente:
                    continue                # UNS impreso, comp. prestada o decision
                nota = fuente.split("Nota")[-1].strip()
                if (tabla, nota, grupo) not in respaldo.get(comp, set()):
                    huerfanas += 1
            # Dos redacciones posibles, segun por donde se resolvio la
            # composicion prestada (ver _composicion_prestada en el builder):
            # "aparece como" para el UNS a secas, "hay una sola" cuando hizo
            # falta desambiguar por (UNS, especificacion impresa en la fila).
            # Las dos citan de donde sale el dato; solo cambia la via.
            if "otra tabla" in est and "aparece como" not in motivo \
                    and "hay una sola" not in motivo:
                prestada_sin_origen += 1
            # La marca de decision se exige a la FILA, no a cada columna: una
            # fila validada puede tener un grupo decidido por una persona y el
            # otro leido del codigo. Es el caso del 9Cr-1Mo-V, cuyo modulo E se
            # decide por la Nota (5) mientras su dilatacion sale de la columna
            # impresa de TE-1. Exigirlo columna a columna marcaba eso como fallo.
            if est.startswith("VALIDADO"):
                fuentes = " ".join(str(ws_map.cell(r, cm[c]).value or "")
                                   for c in ("Fuente E", "Fuente alfa"))
                if "Validado por ingeniero" not in fuentes:
                    validado_sin_marca += 1

    log("| Comprobacion | Detalle | Estado |")
    log("|---|---|---|")
    # Las dos ediciones alimentan bandas de propiedades (DB_E / DB_EC, DB_TE /
    # DB_TEC). Si una se queda sin notas, media tabla pierde la unica via
    # trazable para saber a que grupo pertenece un material, y el hueco no se
    # nota al construir. Se audita aqui para que no pueda reabrirse en silencio.
    completas, faltan = [], []
    for ed in ("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025", "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"):
        for archivo in ("table_tm_1.json", "table_te_1.json"):
            grupos = [n for n in RES.load(f"{ed}/{archivo}").get("note_members", [])
                      if n.get("tipo") == "grupo"]
            (completas if grupos else faltan).append(f"{ed}/{archivo}")

    filas9 = [
        ("Todo grupo asignado cita su fuente",
         f"{sin_cita} filas con grupo y sin fuente", sin_cita == 0),
        ("La Nota citada lista esa composicion",
         f"{huerfanas} citas que el JSON del codigo no respalda", huerfanas == 0),
        ("No sobrevive ningun estado de conjetura",
         "sin filas 'PROPUESTA'",
         not any("PROPUESTA" in str(e).upper() for e in estados)),
        ("Las dos ediciones traen sus Notas de grupo",
         f"{len(completas)}/4 archivos con note_members"
         + (f" · faltan: {', '.join(faltan)}" if faltan else ""),
         not faltan),
        ("Lo decidido por una persona se declara como tal",
         f"{validado_sin_marca} filas VALIDADO sin marca en la fuente",
         validado_sin_marca == 0),
        ("Toda composicion tomada por UNS cita su tabla de origen",
         f"{prestada_sin_origen} filas sin citar la tabla de origen",
         prestada_sin_origen == 0),
    ]
    for etiqueta, detalle, ok in filas9:
        map_bad += 0 if ok else 1
        log(f"| {etiqueta} | {detalle} | {'OK' if ok else 'FALLO'} |")
    log("")
    log("| Estado del mapeo | Filas |")
    log("|---|---:|")
    for e, n in estados.most_common():
        log(f"| {e} | {n} |")
    log("")
    log("Las filas SIN MAPEO no son un defecto de la extraccion: son materiales para "
        "los que II-D no publica modulo ni dilatacion. En ellas el calculo queda "
        "bloqueado, que es lo que exige el codigo.")
    log("")

    # ---- cierre -----------------------------------------------------------
    total = (nbad + bad_tot + extra_bad + len(hits) + len(malas)
             + (0 if cont_ok else 1) + cnt_bad + uniq_bad + semilla_bad + nav_bad
             + map_bad)
    log("## Resultado")
    log("")
    log(f"| Seccion | Fallos |")
    log("|---|---|")
    for etiqueta, v in [("1. Conteos", cnt_bad), ("2. Unicidad", uniq_bad),
                        ("3. Auditoria fila a fila", bad_tot + extra_bad),
                        ("4. Contiguidad de la cascada", 0 if cont_ok else 1),
                        ("5. Portabilidad de formulas", len(hits) + len(malas)),
                        ("6. Interpolacion recalculada", nbad),
                        ("7. Caso semilla", semilla_bad),
                        ("8. Capa de navegacion", nav_bad),
                        ("9. Mapeo de grupos", map_bad)]:
        log(f"| {etiqueta} | {v} |")
    log("")
    log(f"**Total de fallos: {total}.**")
    REPORTE.write_text("\n".join(out), encoding="utf-8")
    print(f"\nReporte escrito en {REPORTE}")
    return total


def main(argv=None):
    global RES, WB, OUTDIR, REPORTE
    aqui = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--wb", required=True, type=Path,
                    help="libro a verificar (.xlsm entregable)")
    ap.add_argument("--outdir", type=Path, default=None,
                    help="directorio de trabajo para las copias recalculadas "
                         "(por defecto, uno temporal que se descarta)")
    ap.add_argument("--report", type=Path,
                    default=aqui.parent / "Reporte_Verificacion_DB_Materiales.md")
    a = ap.parse_args(argv)

    if not a.wb.exists():
        print(f"ERROR: no existe el libro {a.wb}", file=sys.stderr)
        return 2
    if not a.resources.is_dir():
        print(f"ERROR: no existe el directorio de recursos {a.resources}", file=sys.stderr)
        return 2

    RES = Resources(str(a.resources))
    WB = str(a.wb)
    REPORTE = a.report

    if a.outdir:
        a.outdir.mkdir(parents=True, exist_ok=True)
        OUTDIR = a.outdir
        return auditar()
    with tempfile.TemporaryDirectory(prefix="verificar_pcc_") as tmp:
        OUTDIR = Path(tmp)
        return auditar()


if __name__ == "__main__":
    sys.exit(0 if main() == 0 else 1)
