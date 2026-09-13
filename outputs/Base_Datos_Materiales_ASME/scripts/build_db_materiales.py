# -*- coding: utf-8 -*-
"""
build_db_materiales.py — Construye las bases de datos de materiales del
Motor de Calculo ASME PCC a partir de los JSON de resources/.

PLAN-DB-MAT-001

Que escribe
-----------
El libro entero, de una sola pasada y sin edicion manual: las bases `DB_*` y
`MAP_*` leidas de resources/, los doce motores (Art. 212, los diez buscadores
e `Instrucciones`), las nueve hojas de volcado de la Seccion II, el `Dashboard`
y las quince hojas `NAV_*` del arbol de navegacion, con los estados de
visibilidad grabados en el archivo.

Aqui NO va el historial por revision: ese vive en
`outputs/Base_Datos_Materiales_ASME/LEEME_Nota_de_Version.md`, que es su sitio.
Un changelog duplicado al principio de un modulo de 7 000 lineas se queda
viejo sin que nadie lo note — este llego a describir la Rev. 2 con el libro ya
por la 4.

Reglas de diseno que gobiernan todo lo que se emite (ver CLAUDE.md):
* Cero funciones de matriz dinamica. Solo INDEX / MATCH / OFFSET / COUNTIF.
* Validacion de datos por rango literal, nunca por formula.
* Cascada contigua de listas desplegables; la unica celda que se teclea es la
  temperatura de consulta.
* Clave bilingue: cada fila metrica enlaza con su homologa U.S. Customary y la
  paridad entre ediciones se verifica al construir.

Uso:
    python build_db_materiales.py --resources <ruta> --in <xlsm> --out <xlsm>
"""
from __future__ import annotations

import argparse
import datetime
import json
import math
import re
import sys
from collections import Counter
from copy import copy
from pathlib import Path
from typing import NamedTuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.comments import Comment
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

sys.path.insert(0, str(Path(__file__).resolve().parent))
from db_lib import (Resources, bilingual_key, build_material_id, clean, disambiguate,
                    familia_material, make_unique, num, search_key, sort_key,
                    temp_to_number, txt)
# Reconstruccion de las tablas de la Seccion II partes A, B y C. Es una
# libreria pura -sin Excel y sin PDF- y ademas CLI; aqui se usa como libreria.
import secii_tablas as secii
import b36_dimensiones

# ---------------------------------------------------------------------------
# Sistema visual — Swiss Industrial Print
# ---------------------------------------------------------------------------
# UN solo sustrato para las 70 hojas: papel de documentacion sin blanquear,
# tinta carbon y UN acento (rojo de aviacion). Sin degradados, sin sombras y
# sin pasteles: los dos grises que existen son trama de medio tono y solo
# sirven para el metadato, el estado inactivo y la reticula interior.
#
# El sustrato es CLARO por una razon de uso, no de gusto: el libro se imprime
# y se firma, y la variante oscura de terminal sale en negro sobre el papel.
#
# Dos familias, una por funcion, y las dos vienen instaladas con Windows: una
# fuente que Excel no encuentra la sustituye en silencio y deshace la retícula,
# asi que aqui no entra ninguna fuente de descarga por bien que encaje en el
# estilo (JetBrains Mono, Archivo Black).
#   MACRO  estructura: titulo de hoja, banda de seccion y cifra de KPI.
#   MONO   todo el dato: cabecera, celda, etiqueta, metadato y unidad.
PAPEL = "F4F4F0"        # sustrato
PAPEL_2 = "EAE8E3"      # sustrato de compartimento (tarjeta, KPI, campo)
TINTA = "050505"        # tinta carbon: texto y bloque estructural
TINTA_2 = "111111"      # trazo de curva
ROJO = "E61919"         # UNICO acento: aviso, bloqueo, dato vital
GRIS = "8A8A85"         # trama 55 %: metadato y tarjeta marcador
GRIS_2 = "C9C7C1"       # trama 25 %: reticula interior

# Semaforo funcional: la excepcion declarada al acento unico. En un libro de
# calculo el estado de la consulta es informacion de seguridad y hay que verla
# sin leerla, asi que se conservan los tres estados. Van como BLOQUE macizo con
# tinta encima —nunca como pastel de relleno suave— y el rojo mantiene su unico
# significado en todo el libro: bloqueado.
VERDE = "4AF626"        # relleno: seleccion completa / en rango
AMBAR = "E6A019"        # relleno: seleccion incompleta / dato con reserva
AMBAR_TXT = "8A5D00"    # el mismo ambar como TEXTO sobre papel, ya legible

# Leyenda de edicion de los DOS MOTORES DE CALCULO — excepcion declarada y
# acotada por nombre de hoja (Parche_PCC2_Art212 y Collar_PCC2_Art206), del
# mismo tipo que el semaforo funcional de arriba. En un buscador la unica celda
# que se teclea se distingue por ser el UNICO rectangulo cerrado de la zona
# (CAJA_TECLEO), y con dos o tres campos eso basta. Un motor de calculo tiene
# cuarenta y tantas celdas mezcladas —editables, de formula y de rotulo— y ahi
# el borde ya no separa nada: el ingeniero necesita ver de un golpe donde puede
# escribir y donde no. Por eso la distincion pasa al RELLENO, con tres estados
# y una leyenda impresa en la propia hoja:
#   AMARILLO  celda editable (valor libre o lista desplegable)
#   GRIS_2    celda bloqueada: formula
#   PAPEL     rotulo, unidad, descripcion o especificacion
# El tono es el minimo que se distingue del papel sin dejar de ser papel y sin
# competir con el rojo del acento; TINTA encima mantiene contraste de sobra.
AMARILLO = "FAEFC0"     # relleno: celda editable (SOLO en los dos motores)

MACRO = "Arial Black"
MONO = "Consolas"

TITLE_F = Font(name=MACRO, size=12, color=PAPEL)
TITLE_FILL = PatternFill("solid", fgColor=TINTA)
BAND_FILL = PatternFill("solid", fgColor=TINTA)
PAPEL_FILL = PatternFill("solid", fgColor=PAPEL)
SRC_F = Font(name=MONO, size=8, color=GRIS)
HDR_F = Font(name=MONO, size=9, bold=True, color=PAPEL)
HDR_FILL = PatternFill("solid", fgColor=TINTA)
DATA_F = Font(name=MONO, size=9, color=TINTA)
THIN = Side(style="thin", color=GRIS_2)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
# Regla de zona y franja de aviso. La banda de seccion y la cabecera de tabla
# cierran por abajo con la franja roja: es el limite duro entre dos zonas de
# informacion, dibujado y no insinuado.
REGLA = Side(style="medium", color=TINTA)
FRANJA = Side(style="thick", color=ROJO)
CAJA_REGLA = Border(left=REGLA, right=REGLA, top=REGLA, bottom=REGLA)
# Cabecera de tabla: conserva la reticula fina y cierra con la franja roja. Se
# declara como borde completo en vez de retocar el borde ya puesto: openpyxl
# devuelve los lados envueltos en un proxy de estilo y recomponerlos celda a
# celda es fragil.
BOX_FRANJA = Border(left=THIN, right=THIN, top=THIN, bottom=FRANJA)
LBL_F = Font(name=MONO, size=10, bold=True, color=TINTA)
# Campo de formulario. El sustrato del panel es PAPEL_2 y el campo va en PAPEL
# limpio dentro de una caja de tinta: se lee como el hueco que hay que
# rellenar, igual que en un formulario impreso. La UNICA celda que se teclea
# lleva la misma caja en ROJO -ver TEMP_INPUT_FILL / CAJA_TECLEO-, asi que
# nunca se confunde con una lista desplegable.
IN_F = Font(name=MONO, size=10, bold=True, color=TINTA)
IN_FILL = PatternFill("solid", fgColor=PAPEL)
# Leyenda de edicion de los dos motores de calculo (ver el token AMARILLO). Los
# tres rellenos se declaran JUNTOS para que la leyenda que se imprime en la hoja
# y lo que se pinta en las celdas no puedan divergir: si alguien cambia uno,
# cambia el que la leyenda muestra.
MOTOR_IN_FILL = PatternFill("solid", fgColor=AMARILLO)    # editable
MOTOR_CALC_FILL = PatternFill("solid", fgColor=GRIS_2)    # formula, bloqueada
MOTOR_LBL_FILL = PAPEL_FILL                               # rotulo / unidad / nota
# La lista desplegable lleva SOLO la linea inferior, como el hueco de un
# formulario impreso. Encajonarla por los cuatro lados se probo y se descarto al
# mirarlo exportado: siete campos seguidos con caja completa se ven como una
# escalera de barrotes, y ademas el borde de un campo compite con el de su
# vecino en la arista que comparten, que es donde el recuadro rojo de la unica
# celda que se teclea se perdia. Con la linea, la caja roja es el unico
# rectangulo cerrado de la zona y no puede confundirse.
CAJA_CAMPO = Border(bottom=REGLA)
OUT_F = Font(name=MACRO, size=12, color=TINTA)

# Microtipografia: mayusculas y separadores ASCII. Excel no tiene tracking, asi
# que el caracter de la retícula lo llevan la caja mono, la mayuscula y el
# separador; el guion largo del texto original es tipografia de prosa y en una
# banda de seccion se cambia por «//».
def rotulo(texto: str) -> str:
    """Rotulo de banda o de boton: [ MAYUSCULAS // CON SEPARADOR ASCII ]."""
    t = " ".join(str(texto).split()).replace("—", "//").replace(" - ", " // ")
    t = " ".join(t.split())
    return f"[ {t.upper()} ]"


def franja(ws, fila, c1, c2, arriba=False):
    """Franja roja al pie (o al tope) de una banda o de una cabecera de tabla.

    HAY QUE RECORRER EL RANGO, tambien si esta fusionado, y esta comprobado en
    Excel real: el RELLENO de un rango fusionado sale del formato de la celda
    ancla y cubre todo el ancho, pero el BORDE no —Excel dibuja solo el de la
    celda ancla, y la franja acaba siendo un muñon de una columna—.
    Cuidado al comprobarlo: openpyxl MIENTE en los dos sentidos. Al leer un
    libro reconstruye el borde de la ancla sobre todo el rango (parece que
    estuviera), y al escribir no propaga nada si el estilo se puso despues de
    fusionar. La unica evidencia valida es exportar la hoja e ir a mirarla.
    """
    lado = Border(top=FRANJA) if arriba else Border(bottom=FRANJA)
    for c in range(c1, c2 + 1):
        ws.cell(fila, c).border = lado


# Ancho papelado por hoja: el sustrato se aplica al nivel de COLUMNA y no celda
# a celda. Excel guarda un estilo por columna (<col style="n"/>), asi que las
# 54 198 filas del volcado de la Seccion II heredan papel y fuente mono sin un
# solo estilo de celda: el mismo resultado visual sin multiplicar el tamano del
# .xlsm por 2,6 millones de celdas con formato. Toda celda con estilo propio
# —cabecera, banda, tarjeta, campo— lo pisa, que es justo lo que se quiere.
COLS_SUSTRATO_MIN = 20


def sustrato(ws):
    """Papel + mono en TODA la hoja, y sin las lineas de reticula de Excel.

    Las gridlines se apagan porque en este sistema la reticula la DIBUJAN los
    bordes: dejar tambien las de Excel superpone dos mallas distintas. Sobre
    celda rellena Excel ya no las pinta, asi que apagarlas solo uniforma lo que
    queda fuera del ancho papelado.
    """
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TINTA
    ancho = min(100, max(COLS_SUSTRATO_MIN, ws.max_column + 8))
    for c in range(1, ancho + 1):
        dim = ws.column_dimensions[get_column_letter(c)]
        dim.fill = PAPEL_FILL
        dim.font = DATA_F


def aplicar_sustrato(wb):
    """Sustrato en las 70 hojas, incluidas las que trae el maestro."""
    for ws in wb.worksheets:
        sustrato(ws)
    return len(wb.worksheets)


# Grafica. Se estila a mano y NO con ch.style: los presets de Office traen su
# propia paleta de series de colores y su propio tipo, que es exactamente lo que
# este sistema no admite. La curva va en tinta y el punto consultado en el rojo
# del acento; la malla queda como trama de 25 %, que es la reticula de plano.
def _txt_chart(size=800, bold=False, color=TINTA):
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (CharacterProperties, Font as DFont, Paragraph,
                                       ParagraphProperties, RichTextProperties)
    cp = CharacterProperties(latin=DFont(typeface=MONO), sz=size, b=bold,
                             solidFill=color)
    return (cp, RichText(bodyPr=RichTextProperties(),
                         p=[Paragraph(pPr=ParagraphProperties(defRPr=cp),
                                      endParaRPr=cp)]))


def estilizar_chart(ch):
    """Aplica el sistema visual a una grafica ya construida (series incluidas).

    Llamar DESPUES de fijar titulo y titulos de eje: el estilo del rotulo se
    escribe sobre el objeto Title que crea openpyxl al asignarle la cadena.
    """
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.drawing.text import ParagraphProperties

    ch.graphical_properties = GraphicalProperties(
        solidFill=PAPEL, ln=LineProperties(solidFill=TINTA, w=12700))
    ch.plot_area.graphicalProperties = GraphicalProperties(
        solidFill=PAPEL, ln=LineProperties(solidFill=TINTA, w=9525))
    cp_eje, tx_eje = _txt_chart(800)
    cp_tit, _ = _txt_chart(900, bold=True)
    for ax in (ch.x_axis, ch.y_axis):
        ax.spPr = GraphicalProperties(ln=LineProperties(solidFill=TINTA, w=9525))
        ax.txPr = tx_eje
        ax.majorGridlines = ChartLines(spPr=GraphicalProperties(
            ln=LineProperties(solidFill=GRIS_2, w=3175)))
        if ax.title is not None:
            ax.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=cp_tit)
    if ch.title is not None:
        cp_h, _ = _txt_chart(1000, bold=True)
        ch.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=cp_h)
    if ch.legend is not None:
        ch.legend.txPr = tx_eje
    return ch


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
def cabecera_hoja(ws, ncols):
    """Bloque de cabecera de una hoja de datos, a lo ancho de la tabla.

    Tres franjas apiladas: el titulo en tinta maciza, la tira de procedencia en
    medio tono y la cabecera de columnas, que cierra con la franja roja
    (BOX_FRANJA). El ancho lo da la tabla —no una constante— porque estas hojas
    van de 7 a 48 columnas y un bloque mas corto que la tabla se lee como un
    titulo suelto, no como el encabezado del volcado.
    """
    for c in range(1, ncols + 1):
        cel = ws.cell(R_TITLE, c)
        cel.fill, cel.font = TITLE_FILL, TITLE_F
        src = ws.cell(R_SRC, c)
        src.fill, src.font = PatternFill("solid", fgColor=PAPEL_2), SRC_F
    ws.row_dimensions[R_TITLE].height = 22


def new_sheet(wb, name, title, source):
    ws = wb.create_sheet(name)
    ws["A1"] = rotulo(title)
    ws["A1"].font, ws["A1"].fill = TITLE_F, TITLE_FILL
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws["A2"] = source
    ws["A2"].font = SRC_F
    ws.freeze_panes = "A4"
    return ws


def write_headers(ws, cols, temps=None):
    for j, h in enumerate(cols, start=1):
        c = ws.cell(R_HDR, j, h)
        c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    if temps:
        for j, t in enumerate(temps, start=len(cols) + 1):
            c = ws.cell(R_HDR, j, t)
            c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
            c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[R_HDR].height = 46
    cabecera_hoja(ws, len(cols) + (len(temps) if temps else 0))
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
    # errorStyle="stop": Excel RECHAZA lo que no este en la lista, en vez de
    # aceptarlo en silencio. Es la mitad que le faltaba a la regla 14 — apuntar el
    # desplegable a la base no sirve de nada si el usuario puede teclear al lado.
    # Aviso honesto y deliberado: el bloqueo actua al TECLEAR; no al pegar ni al
    # escribir por macro. Reduce el error de dedo, no lo vuelve imposible.
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showErrorMessage=True, errorStyle="stop",
                        errorTitle="Valor fuera de lista",
                        error="Elija uno de los valores de la lista desplegable.")
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
                c = ws.cell(r, 1, f"{d.get('table_id')} // {d.get('title')}")
                c.font = Font(name=MONO, size=9, bold=True, color=TINTA)
                r += 1
                for j, k in enumerate(keys, start=1):
                    cc = ws.cell(r, j, k)
                    cc.font, cc.fill, cc.border = HDR_F, HDR_FILL, BOX_FRANJA
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


# ---------------------------------------------------------------------------
# DB_B31_B1 / DB_B31_B1C — Tabla B-1 del Apendice B: HDS de termoplasticos
# ---------------------------------------------------------------------------
# El Apendice B publica el esfuerzo de diseno hidrostatico en UNA TABLA POR
# EDICION —B-1 metrica y B-1C U.S. Customary—, asi que el conmutador de la
# regla 10 cambia de HOJA y no convierte nunca. Es ademas el unico dato del
# Apendice B tabulado FRENTE A LA TEMPERATURA: las demas tablas publican
# listados de especificacion (B-2, B-3) o una presion admisible puntual
# (B-4, B-5, B-6), que es otra magnitud. Un dato, un motor: B-1 sale de
# Buscar_NoMetalicos igual que salio el Apendice C.
APXB1_COLS = [
    "material_id", "Tabla", "k0", "k1", "k2", "k3", "k4", "clave_bi",
    "Designacion de material", "Spec. No. (ASTM)", "Designacion de tuberia",
    "Cell Class", "Variante",
    "Temp. min. recomendada", "Temp. max. recomendada",
    "Unidad de temperatura", "Unidad de HDS", "Encabezado impreso del HDS",
    "T primera tabulada", "T ultima tabulada",
    "Notas del HDS", "Notas de los limites", "Observacion", "Linea", "n_pts"]
CB1 = {n: i + 1 for i, n in enumerate(APXB1_COLS)}
CLB1 = {n: get_column_letter(i) for n, i in CB1.items()}
assert APXB1_COLS[:8] == STRESS_COLS[:8], \
    "el contrato de las 8 primeras columnas es lo que reutiliza build_listas"

SIN_SPEC_B1 = "(sin Spec. No. impreso)"
SIN_TUB_B1 = "(sin designacion de tuberia)"
SIN_CELL_B1 = "(sin Cell Class impresa)"

# Reparto de las 10 columnas de B-1/B-1C por POSICION. El JSON alinea `columns`
# 1:1 con las claves de cada fila —y eso se comprueba antes de leer nada—, pero
# `column_groups` llega como una lista PLANA de fragmentos de cabecera
# ("Recommended", "Temperature Limits, °C", "[Notes (1), (2)]", ...) que no se
# puede alinear por posicion con las columnas. Por eso el reparto se declara
# aqui: describe la maqueta impresa y no aporta ni un valor.
B1_COL_SPEC, B1_COL_TUB, B1_COL_MAT, B1_COL_CELL = 0, 1, 2, 3
B1_COL_TMIN, B1_COL_TMAX = 4, 5
B1_COLS_HDS = (6, 7, 8, 9)

# Rotulos de estado del motor de la Tabla B-1. Viven aqui, y no sueltos dentro
# de la formula, porque verificar.py §11 vuelve a emitir la misma expresion para
# recalcularla en Excel: si el motor y el verificador escribiesen cada uno su
# literal, una divergencia de redaccion pasaria desapercibida.
EST_B1_SIN_SEL = "SIN SELECCION"
EST_B1_SIN_TAB = "SIN HDS TABULADO — la fila solo publica limites de temperatura"
EST_B1_BAJO = "FUERA DE RANGO — T por debajo del limite minimo recomendado"
EST_B1_ALTO_LIM = "FUERA DE RANGO — T por encima del limite maximo recomendado"
EST_B1_ALTO_TAB = "FUERA DE RANGO — T por encima del ultimo punto tabulado"
EST_B1_NOTA3 = "EN RANGO — Nota (3): se sostiene el HDS de la primera T tabulada"
EST_B1_OK = "EN RANGO"
VAL_B1_BLOQUEADO = "BLOQUEADO"

_RE_NOTA_B = re.compile(r"\[Notes?\s*([^\]]*)\]")
_RE_NUM_NOTA = re.compile(r"\((\d+[A-Za-z]?)\)")


def _notas_citadas(*textos):
    """Numeros de nota al pie citados en los encabezados IMPRESOS.

    El codigo escribe '23°C [Note (3)]' en la columna de HDS y
    '[Notes (1), (2)]' en el grupo de los limites de temperatura. Se conservan
    los numeros tal como los cita el encabezado; no se deduce ninguno.
    """
    out = []
    for t in textos:
        for blk in _RE_NOTA_B.findall(txt(t)):
            for n in _RE_NUM_NOTA.findall(blk):
                if n not in out:
                    out.append(n)
    return ", ".join(out)


def _temp_de_encabezado(h):
    """Temperatura impresa en el encabezado de una columna de HDS.

    Llega como '23°C [Note (3)]' en la edicion metrica y como '100°- F' en la
    U.S. Customary (el guion es un artefacto de la extraccion). Se toma el
    primer numero, que es lo unico que el codigo imprime como temperatura de
    esa columna, y despues se coteja contra la clave de la fila.
    """
    m = re.search(r"-?\d+(?:\.\d+)?", txt(h))
    return num(m.group(0)) if m else None


def _carga_b1(res, si):
    """Tabla B-1 (metrica) o B-1C (U.S. Customary), tal como esta impresa."""
    fn = f"{APX}/appendix_b/table_b_1{'' if si else 'c'}.json"
    d = res.load(fn)
    cols = [clean(c) for c in (d.get("columns") or [])]
    rows = d.get("rows") or []
    if len(cols) != 10 or not rows:
        raise SystemExit(
            f"Tabla B-1: {fn} deberia publicar 10 columnas y al menos una fila; "
            f"trae {len(cols)} columnas y {len(rows)} filas. El motor no se "
            "construye con una extraccion que no reconoce.")
    # `columns` es de donde salen el rotulo impreso de cada campo y la
    # temperatura de cada columna de HDS. Si dejase de estar alineada 1:1 con
    # las claves de la fila, el motor leeria una columna por otra en silencio.
    for i, row in enumerate(rows, start=1):
        if len(row) != len(cols):
            raise SystemExit(
                f"Tabla B-1: la fila impresa {i} de {fn} trae {len(row)} campos "
                f"y el encabezado declara {len(cols)} columnas.")
    temps = []
    for j in B1_COLS_HDS:
        t = _temp_de_encabezado(cols[j])
        if t is None:
            raise SystemExit(
                f"Tabla B-1: no se lee la temperatura del encabezado {cols[j]!r} "
                f"de {fn}.")
        temps.append(t)
    # Cotejo cruzado: la clave de la fila lleva la misma temperatura que el
    # encabezado ('23_c' <-> '23°C'). Es la comprobacion que delata un
    # desplazamiento de columna, que es el error caro en esta tabla.
    claves = list(rows[0])
    for j, t in zip(B1_COLS_HDS, temps):
        if _temp_de_encabezado(claves[j]) != t:
            raise SystemExit(
                f"Tabla B-1: la columna {j} de {fn} declara {t} en el encabezado "
                f"{cols[j]!r} y {claves[j]!r} en la clave de la fila.")

    # Una observacion del codigo se adjudica a una fila SOLO si el texto de la
    # observacion nombra literalmente su Spec. No. impreso. Es una comprobacion
    # de contencion sobre una designacion impresa, no una inferencia.
    obs_tabla = [clean(o) for o in (d.get("observaciones_del_codigo") or [])]
    filas = []
    for i, row in enumerate(rows, start=1):
        ks = list(row)
        vals = {}
        for j, t in zip(B1_COLS_HDS, temps):
            v = num(row[ks[j]])
            if v is not None:
                vals[t] = v
        spec = clean(row[ks[B1_COL_SPEC]])
        obs = [o for o in obs_tabla if spec and txt(spec) in txt(o)]
        filas.append(dict(idx=i, spec=spec,
                          tuberia=clean(row[ks[B1_COL_TUB]]),
                          material=clean(row[ks[B1_COL_MAT]]),
                          cell=clean(row[ks[B1_COL_CELL]]),
                          tmin=num(row[ks[B1_COL_TMIN]]),
                          tmax=num(row[ks[B1_COL_TMAX]]),
                          vals=vals, obs=obs))
    return d, cols, temps, filas


def build_b1(res, wb, system):
    """DB_B31_B1 / DB_B31_B1C — la Tabla B-1 del Apendice B, por edicion."""
    si = system == "SI"
    name = "DB_B31_B1" if si else "DB_B31_B1C"
    d, cols, temps, filas = _carga_b1(res, si)
    fn = f"{APX}/appendix_b/table_b_1{'' if si else 'c'}.json"
    tid = (clean(d.get("table_id")) or "").replace("Table ", "") or \
        ("B-1" if si else "B-1C")
    u_t, u_v = ("°C", "MPa") if si else ("°F", "ksi")
    hdr_hds = clean(d.get("value_axis")) or ""
    n_hds = _notas_citadas(cols[B1_COLS_HDS[0]], hdr_hds)
    n_lim = _notas_citadas(*(d.get("column_groups") or []))

    # Orden de la base: material -> spec -> designacion de tuberia -> cell class,
    # y la fila impresa como ultimo criterio. Cada nivel queda en un bloque
    # CONTIGUO, que es lo que exige la cascada sin matrices dinamicas (regla 5).
    filas.sort(key=lambda f: (txt(f["material"]).upper(),
                              _seg(f["spec"], SIN_SPEC_B1).upper(),
                              _seg(f["tuberia"], SIN_TUB_B1).upper(),
                              _seg(f["cell"], SIN_CELL_B1).upper(),
                              f["idx"]))
    # La variante solo existe si (material, spec, tuberia, cell class) se repite.
    # Hoy no se repite en ninguna de las dos ediciones, pero una reextraccion que
    # colapsase dos filas dejaria una inalcanzable sin este nivel: es justo el
    # defecto que traia Buscar_NoMetalicos, donde 17 de las 36 filas de B-1 no
    # se podian seleccionar porque la clave era solo la designacion de material.
    rep = Counter((txt(f["material"]), txt(f["spec"]), txt(f["tuberia"]),
                   txt(f["cell"])) for f in filas)
    records = []
    for f in filas:
        mat = _seg(f["material"])
        sp = _seg(f["spec"], SIN_SPEC_B1)
        tb = _seg(f["tuberia"], SIN_TUB_B1)
        cl = _seg(f["cell"], SIN_CELL_B1)
        clave = (txt(f["material"]), txt(f["spec"]), txt(f["tuberia"]),
                 txt(f["cell"]))
        var = f"fila impresa {f['idx']}" if rep[clave] > 1 else VAR_UNICA
        k0 = mat
        k1 = f"{k0}|{sp}"
        k2 = f"{k1}|{tb}"
        k3 = f"{k2}|{cl}"
        k4 = f"{k3}|{var}"
        f["k"] = (k0, k1, k2, k3, k4)
        f["mid"] = build_material_id([tid, mat, sp, tb, cl, var])
        # Enlace SI<->US POSICIONAL, como en el Apendice C: los nombres pueden
        # divergir entre ediciones por lo que el propio codigo imprime distinto
        # (la designacion de tuberia de F2389 es 'PR' en B-1 e 'IPS Sch. 80' en
        # B-1C). La contrapartida obligatoria es verificar_paridad_b1().
        f["bi"] = f"B1#{f['idx']}"
        records.append((
            [f["mid"], tid, k0, k1, k2, k3, k4, f["bi"],
             f["material"], f["spec"], f["tuberia"], f["cell"], var,
             f["tmin"], f["tmax"], u_t, u_v, hdr_hds,
             min(f["vals"]) if f["vals"] else None,
             max(f["vals"]) if f["vals"] else None,
             n_hds, n_lim, " · ".join(f["obs"]) or None, f["idx"], len(f["vals"])],
            f["vals"]))

    ws = new_sheet(
        wb, name,
        "ESFUERZO DE DISENO HIDROSTATICO (HDS) DE TUBERIA TERMOPLASTICA — "
        "ASME B31.3-2024, Apendice B, Tabla " + tid +
        (" (edicion metrica: MPa, °C)" if si
         else " (edicion U.S. Customary: ksi, °F)"),
        f"Fuente: resources/{fn} · ASME B31.3-2024 · Valores tal como estan "
        "impresos: SI y US son extracciones independientes de la tabla que "
        "publica cada edicion, nunca una conversion (regla 9). Las columnas "
        "'Temp. min./max. recomendada' son los Recommended Temperature Limits "
        "de la propia tabla, NO el rango de HDS tabulado. " +
        (" · ".join(clean(o) for o in
                    (d.get("observaciones_del_codigo") or [])) or ""))
    n = write_headers(ws, APXB1_COLS, temps)
    last = write_rows(ws, records, n, temps)
    autosize(ws, {CLB1["material_id"]: 56, CLB1["Tabla"]: 8, CLB1["clave_bi"]: 10,
                  CLB1["Designacion de material"]: 22,
                  CLB1["Spec. No. (ASTM)"]: 16,
                  CLB1["Designacion de tuberia"]: 24, CLB1["Cell Class"]: 12,
                  CLB1["Variante"]: 16, CLB1["Encabezado impreso del HDS"]: 34,
                  CLB1["Observacion"]: 60})
    for cn in ("k0", "k1", "k2", "k3", "k4", "clave_bi", "Linea", "n_pts"):
        ws.column_dimensions[CLB1[cn]].hidden = True
    t0, v0, npack = append_packed(ws, records, n, temps, CB1["n_pts"])
    ws.auto_filter.ref = f"A{R_HDR}:{get_column_letter(n + len(temps))}{last}"
    record_meta(name, tid, fn, "2024", system, len(records),
                "Esfuerzo de diseno hidrostatico y limites de temperatura "
                "recomendados de tuberia termoplastica.")
    return dict(sheet=name, n_ident=n, temps=temps, last_row=last, recs=filas,
                pack_t0=t0, pack_v0=v0, npack=npack, npts_col=CB1["n_pts"],
                tabla=tid, u_t=u_t, u_v=u_v)


def verificar_paridad_b1(si_info, us_info):
    """Contrapartida del enlace posicional de B-1 <-> B-1C.

    Se compara lo que identifica al MATERIAL —designacion y Spec. No.—, no todo
    lo impreso: el propio codigo publica distinta designacion de tuberia y
    distinto limite maximo para F2389 en cada edicion, y eso esta declarado en
    `observaciones_del_codigo` y se conserva (regla 9). Si lo que divergiera
    fuese la identidad, el enlace estaria uniendo dos materiales distintos y el
    build ABORTA.
    """
    a = {f["bi"]: f for f in si_info["recs"]}
    b = {f["bi"]: f for f in us_info["recs"]}
    if set(a) != set(b):
        raise SystemExit(
            "Tabla B-1: las dos ediciones no publican el mismo numero de filas. "
            f"Solo en SI: {sorted(set(a) - set(b))[:8]}; "
            f"solo en US: {sorted(set(b) - set(a))[:8]}.")
    malas = []
    for bi, f in a.items():
        ident_si = (txt(f["material"]).upper(), txt(f["spec"]).upper())
        ident_us = (txt(b[bi]["material"]).upper(), txt(b[bi]["spec"]).upper())
        if ident_si != ident_us:
            malas.append(f"{bi}: SI={ident_si} US={ident_us}")
    if malas:
        raise SystemExit(
            "Tabla B-1: el enlace SI/US es posicional y hay filas cuya "
            "identidad no casa entre ediciones:\n  " + "\n  ".join(malas))
    # Lo que SI puede divergir se cuenta y se declara, no se silencia: son
    # asimetrias del propio codigo, ya recogidas en `observaciones_del_codigo`.
    dif_tub = [bi for bi, f in a.items()
               if txt(f["tuberia"]) != txt(b[bi]["tuberia"])]
    ISSUES.append(
        f"Tabla B-1: {len(a)} filas enlazadas SI<->US por posicion (clave_bi); "
        "identidad (designacion de material + Spec. No.) identica en las dos "
        f"ediciones. Designacion de tuberia distinta entre ediciones en "
        f"{len(dif_tub)} fila(s) ({', '.join(dif_tub) or 'ninguna'}): asimetria "
        "impresa por el codigo, se conserva tal cual (regla 9).")


# El RESTO del Apendice B —B-2 y B-3 (listados de especificacion de RTR y RPM)
# y B-4, B-5 y B-6 (presiones admisibles de concreto, vidrio borosilicato y
# PEX-AL-PEX)— NO se carga en este libro. Existio como DB_NoMetalicos +
# Buscar_NoMetalicos y se retiro en la Rev. 4d por decision de alcance: son
# tablas que este trabajo no usa. La extraccion sigue intacta en
# resources/.../appendix_b/table_b_2..b_6.json, que es la fuente de verdad; lo
# que se quito es la carga al libro. En el arbol aparece como tarjeta marcador
# «NO CARGADO EN ESTE LIBRO», igual que el B31.1 o la Seccion VIII: asi el nivel
# sigue explicando la taxonomia y se ve de un vistazo que falta.


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
        out.append((texto, Font(name=MONO, size=9, color=TINTA)))
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


# ---------------------------------------------------------------------------
# Seccion II, partes A, B y C (Fases 3 y 4 del plan)
# ---------------------------------------------------------------------------
# Nueve hojas de DATOS: sin buscador y sin una sola formula. El volcado integro
# no pierde nada -una fila del libro por cada fila impresa, incluidas las que no
# se dejaron repartir, que van enteras en una celda y marcadas AMBIGUA- y las
# dos hojas normalizadas solo recogen las tablas cuyos encabezados se resuelven
# ENTEROS contra el vocabulario del codigo. Ver secii_tablas.py.
#
# Estas hojas llevan una fila mas que el resto de las bases: la 3 queda libre
# para el enlace VOLVER, porque son navegables desde el Dashboard y el resto de
# hojas navegables ancla ahi su boton.
R_HDR_SECII, R_DATA_SECII = 4, 5

HOJAS_SECII_DB = {"bpvc_ii_a_1": "DB_SecII_A1", "bpvc_ii_a_2": "DB_SecII_A2",
                  "bpvc_ii_b": "DB_SecII_B", "bpvc_ii_c": "DB_SecII_C"}
NAV_SECII = ["CAT_SecII", "IDX_SecII_Tablas", "DB_SecII_A1", "DB_SecII_A2",
             "DB_SecII_B", "DB_SecII_C", "DB_SecII_Notas", "DB_SecII_Quimica",
             "DB_SecII_Traccion"]

# Excel no admite mas de 32 767 caracteres en una celda. Ninguna fila de estas
# tablas se acerca, pero el limite se comprueba y se declara en vez de
# truncar en silencio: truncar perderia texto impreso del codigo.
MAX_CELDA = 32767


def _txt_celda(ws, r, c, v):
    """Escribe una celda de TEXTO, aunque el codigo la imprima empezando por
    «=» o «+». openpyxl convertiria eso en formula, y una formula en estas
    hojas rompe la regla 1 del libro y ademas mostraria #NAME?."""
    if v in (None, ""):
        return None
    s = secii.xml_seguro(str(v))
    cel = ws.cell(r, c)
    cel.value = s
    if s[:1] in "=+-@":
        cel.data_type = "s"
    cel.font = DATA_F
    return cel


def new_sheet_secii(wb, name, title, source):
    ws = wb.create_sheet(name)
    ws["A1"] = rotulo(title)
    ws["A1"].font, ws["A1"].fill = TITLE_F, TITLE_FILL
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws["A2"] = source
    ws["A2"].font = SRC_F
    ws.freeze_panes = f"A{R_DATA_SECII}"
    return ws


def _hdr_secii(ws, cols):
    for j, h in enumerate(cols, start=1):
        c = ws.cell(R_HDR_SECII, j, h)
        c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="center")
    ws.row_dimensions[R_HDR_SECII].height = 40
    cabecera_hoja(ws, len(cols))
    return len(cols)


def _pdf_1based(p):
    """`pdf_pages` de las partes A, B y C es 0-based (al reves que la II-D):
    la hoja publica la pagina citable, +1. Ver CLAUDE.md."""
    try:
        return int(p) + 1
    except (TypeError, ValueError):
        return p


def _cargar_secii(res):
    """Recorre las cuatro partes una sola vez y devuelve lo que las nueve hojas
    necesitan. Se hace en una pasada porque son 227 MB de JSON: releerlos por
    hoja multiplicaria por nueve el tiempo de build sin ganar nada."""
    raiz = Path(res.root)
    partes = []
    for parte in secii.PARTES:
        idx_path = raiz / secii.SEC_II / parte / "index.json"
        with open(idx_path, encoding="utf-8") as fh:
            indice = json.load(fh)
        specs = []
        for ruta in secii.archivos_de(raiz, parte):
            spec = secii.cargar_spec(ruta)
            sid = spec["spec"].get("id") or ruta.stem
            specs.append(dict(id=sid, archivo=ruta.name,
                              tablas=secii.tablas_de(spec),
                              huecos=secii.huecos_de(spec)))
        partes.append(dict(parte=parte, indice=indice, specs=specs))
    return partes


def build_cat_secii(wb, datos):
    ws = new_sheet_secii(
        wb, "CAT_SecII",
        "CATALOGO — ASME BPVC Seccion II, partes A (2 vol.), B y C · Edicion 2025",
        "Una fila por entrada del indice de cada parte (especificaciones y "
        "apendices). 'Pagina PDF' se publica 1-based y citable: el `pdf_pages` "
        "de las partes A, B y C es 0-based, al reves que el de la II-D. El "
        "folio impreso es el que lleva la pagina del codigo.")
    _hdr_secii(ws, ["Parte", "Carpeta", "Especificacion", "Titulo",
                    "Designacion equivalente (ASTM/AWS)", "Pagina PDF ini.",
                    "Pagina PDF fin", "Folio impreso ini.", "Folio impreso fin",
                    "Paginas", "Figuras", "Tablas logicas detectadas",
                    "Archivo fuente (resources/)"])
    r = R_DATA_SECII
    for p in datos:
        por_id = {s["id"]: s for s in p["specs"]}
        for e in p["indice"]["entries"]:
            pdfp = e.get("pdf_pages") or [None, None]
            imp = e.get("printed_pages") or [None, None]
            spec_id = txt(e.get("specification"))
            # El indice nombra la entrada por su designacion; el JSON de la
            # especificacion trae su propio id. Se casan por el fichero, que es
            # lo unico que los dos declaran igual.
            fichero = txt(e.get("file")).split("/")[-1]
            hallado = next((s for s in p["specs"] if s["archivo"] == fichero),
                           por_id.get(spec_id))
            vals = [p["parte"], txt(e.get("folder")), spec_id or "(apendice)",
                    txt(e.get("title")),
                    txt(e.get("astm_designation") or e.get("equivalent_designation")),
                    _pdf_1based(pdfp[0]), _pdf_1based(pdfp[1]),
                    imp[0], imp[1], e.get("page_count"), e.get("figure_count"),
                    len(hallado["tablas"]) if hallado else 0,
                    f"{p['parte']}/specifications/{fichero}"]
            for j, v in enumerate(vals, start=1):
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    ws.cell(r, j, v).font = DATA_F
                else:
                    _txt_celda(ws, r, j, v)
            r += 1
    last = r - 1
    autosize(ws, {"A": 14, "B": 24, "C": 20, "D": 62, "E": 22, "F": 12, "G": 12,
                  "H": 12, "I": 12, "J": 9, "K": 9, "L": 12, "M": 52})
    ws.auto_filter.ref = f"A{R_HDR_SECII}:M{last}"
    record_meta("CAT_SecII", "Indice de las partes A/B/C",
                "ASME_BPVC/Sec_II/{bpvc_ii_a_1,a_2,b,c}/index.json", "2025", "-",
                last - R_DATA_SECII + 1,
                "Catalogo de las 379 entradas. Pagina PDF publicada 1-based.")
    return last


def build_idx_secii(wb, datos):
    ws = new_sheet_secii(
        wb, "IDX_SecII_Tablas",
        "INDICE DE TABLAS — ASME BPVC Seccion II, partes A, B y C",
        "Una fila por TABLA LOGICA (las continuaciones de pagina ya unidas). "
        "El reparto por confianza dice cuanto de la tabla quedo tabulado: "
        "AMBIGUA no pierde texto —la fila va entera en una celda— pero no queda "
        "repartida en columnas. 'Normalizada' dice si la tabla alimenta ademas "
        "DB_SecII_Quimica o DB_SecII_Traccion, y si no, por que no.")
    _hdr_secii(ws, ["Parte", "Especificacion", "Tabla", "Titulo",
                    "Paginas PDF", "Folio impreso", "Bloques de origen",
                    "Filas", "Columnas", "Origen de las columnas",
                    "Continuada", "Suelta", "EXACTA", "POR CONTEO",
                    "AMBIGUA", "Motivo dominante", "Normalizada",
                    "Por que no se normaliza"])
    r = R_DATA_SECII
    for p in datos:
        for s in p["specs"]:
            for t in s["tablas"]:
                conf = Counter(f["confianza"] for f in t["filas"])
                motivos = Counter(f["motivo"] for f in t["filas"] if f["motivo"])
                q, mq = secii.normalizar_quimica(t)
                tr, mt = secii.normalizar_traccion(t)
                norm = " + ".join(x for x in (("Quimica" if q else ""),
                                              ("Traccion" if tr else "")) if x)
                porque = "" if norm else f"quimica: {mq} · traccion: {mt}"
                vals = [p["parte"], s["id"], t["n"], t["titulo"],
                        ", ".join(str(_pdf_1based(x)) for x in t["paginas"]),
                        t["folio"], ", ".join(t["bloques"]), len(t["filas"]),
                        t["ncols"],
                        (t["resumen"] or {}).get("origen_ncols", "")
                        if isinstance(t["resumen"], dict) else "",
                        "si" if t["continuada"] else "no",
                        "si" if t["suelta"] else "no",
                        conf[secii.EXACTA], conf[secii.POR_CONTEO],
                        conf[secii.AMBIGUA],
                        motivos.most_common(1)[0][0] if motivos else "",
                        norm or "no", porque]
                for j, v in enumerate(vals, start=1):
                    if isinstance(v, int) and not isinstance(v, bool):
                        ws.cell(r, j, v).font = DATA_F
                    else:
                        _txt_celda(ws, r, j, v)
                r += 1
    last = r - 1
    autosize(ws, {"A": 14, "B": 20, "C": 7, "D": 54, "E": 14, "F": 12, "G": 30,
                  "H": 8, "I": 9, "J": 34, "K": 11, "L": 8, "M": 9, "N": 11,
                  "O": 9, "P": 44, "Q": 14, "R": 60})
    ws.auto_filter.ref = f"A{R_HDR_SECII}:R{last}"
    record_meta("IDX_SecII_Tablas", "Tablas logicas de las partes A/B/C",
                "ASME_BPVC/Sec_II/*/specifications/*.json (bloques Table -> Line)",
                "2025", "-", last - R_DATA_SECII + 1,
                "Reparto por confianza y motivo de la fila. AMBIGUA conserva el "
                "texto entero, no lo pierde.")
    return last


def build_db_secii(wb, datos, ncols_max):
    """El volcado integro: una fila del libro por cada fila impresa."""
    largas = 0
    lasts = {}
    for p in datos:
        nombre = HOJAS_SECII_DB[p["parte"]]
        ws = new_sheet_secii(
            wb, nombre,
            f"VOLCADO INTEGRO — ASME BPVC Seccion II · {p['parte']}",
            "Una fila por fila impresa, en formato ragged: C01..Cnn son las "
            "columnas que esa tabla demostro tener. Confianza EXACTA = cada "
            "Line a su columna por su bbox; POR CONTEO = fila de un solo Line "
            "partida desde la derecha con el conteo confirmado por la "
            "geometria; AMBIGUA = no se pudo repartir sin adivinar y el texto "
            "impreso se conserva ENTERO en C01. No se interpola nunca.")
        cols = ["Especificacion", "Tabla", "Titulo de la tabla", "Fila",
                "Tipo", "Confianza", "Motivo", "Pagina PDF", "Folio impreso",
                "Bloque"] + [f"C{i:02d}" for i in range(1, ncols_max + 1)]
        n_ident = 10
        _hdr_secii(ws, cols)
        r = R_DATA_SECII
        for s in p["specs"]:
            for t in s["tablas"]:
                for i, f in enumerate(t["filas"], start=1):
                    l0 = f["lineas"][0]
                    ident = [s["id"], t["n"], t["titulo"], i, f["tipo"],
                             f["confianza"], f["motivo"],
                             _pdf_1based(l0["pagina"]), l0["folio"], l0["id"]]
                    for j, v in enumerate(ident, start=1):
                        if isinstance(v, int) and not isinstance(v, bool):
                            ws.cell(r, j, v).font = DATA_F
                        else:
                            _txt_celda(ws, r, j, v)
                    for j, celda in enumerate(f["celdas"], start=n_ident + 1):
                        if len(celda) > MAX_CELDA:
                            largas += 1
                            continue
                        _txt_celda(ws, r, j, celda)
                    r += 1
        lasts[nombre] = r - 1
        autosize(ws, {"A": 20, "B": 7, "C": 40, "D": 7, "E": 10, "F": 11,
                      "G": 40, "H": 11, "I": 12, "J": 22})
        for i in range(1, ncols_max + 1):
            ws.column_dimensions[get_column_letter(n_ident + i)].width = 20
        ws.auto_filter.ref = (f"A{R_HDR_SECII}:"
                              f"{get_column_letter(n_ident + ncols_max)}{lasts[nombre]}")
        record_meta(nombre, f"Tablas de {p['parte']}",
                    f"ASME_BPVC/Sec_II/{p['parte']}/specifications/*.json",
                    "2025", "-", lasts[nombre] - R_DATA_SECII + 1,
                    "Volcado integro. Ninguna celda se rellena por "
                    "interpolacion; AMBIGUA conserva la fila entera.")
    if largas:
        ISSUES.append(f"DB_SecII_*: {largas} celdas superaban el limite de "
                      f"{MAX_CELDA} caracteres de Excel y se dejaron vacias. "
                      f"El texto sigue integro en resources/; se declara aqui "
                      f"en vez de truncarlo en silencio.")
    return lasts


def build_notas_secii(wb, datos):
    ws = new_sheet_secii(
        wb, "DB_SecII_Notas",
        "NOTAS AL PIE — tablas de ASME BPVC Seccion II, partes A, B y C",
        "Notas al pie y bloques NOTE de cada tabla. El marcador va PEGADO a la "
        "primera palabra porque el codigo lo imprime como superindice: se "
        "extrae para poder citarlo, pero el texto se conserva entero.")
    _hdr_secii(ws, ["Parte", "Especificacion", "Tabla", "Titulo de la tabla",
                    "Marcador", "Texto de la nota", "Bloque"])
    r = R_DATA_SECII
    for p in datos:
        for s in p["specs"]:
            for t in s["tablas"]:
                for nt in t["notas"]:
                    vals = [p["parte"], s["id"], t["n"], t["titulo"],
                            secii.marcador_de_nota(nt["texto"]),
                            nt["texto"][:MAX_CELDA], nt["id"]]
                    for j, v in enumerate(vals, start=1):
                        if isinstance(v, int) and not isinstance(v, bool):
                            ws.cell(r, j, v).font = DATA_F
                        else:
                            _txt_celda(ws, r, j, v)
                    r += 1
    last = r - 1
    autosize(ws, {"A": 14, "B": 20, "C": 7, "D": 40, "E": 11, "F": 110, "G": 22})
    ws.auto_filter.ref = f"A{R_HDR_SECII}:G{last}"
    record_meta("DB_SecII_Notas", "Notas al pie de las tablas de A/B/C",
                "ASME_BPVC/Sec_II/*/specifications/*.json (Footnote / Text)",
                "2025", "-", last - R_DATA_SECII + 1, "")
    return last


def build_normalizadas_secii(wb, datos):
    """DB_SecII_Quimica y DB_SecII_Traccion.

    Solo entra la tabla cuyos encabezados se resuelven ENTEROS contra el
    vocabulario del codigo. El resto se queda en el volcado integro -donde no
    se pierde nada- y `IDX_SecII_Tablas` dice por que, tabla a tabla. Es poca
    tabla: los encabezados de la Seccion II llegan en su mayoria sin partir,
    porque son filas de un solo `Line` sin fichas de valor con que partirlas.
    Forzar el encaje daria una hoja que PARECE completa y no lo es.
    """
    resultados = {}
    for hoja, fn, cols_val, titulo, fuente in (
        ("DB_SecII_Quimica", secii.normalizar_quimica, secii.COLS_QUIMICA,
         "COMPOSICION QUIMICA NORMALIZADA — Seccion II, partes A, B y C",
         "Una fila por (especificacion, tabla, grado). Valor TAL COMO ESTA "
         "IMPRESO (regla 9): la celda dice `0.27-0.93` o `0.25 max`, no un "
         "minimo y un maximo numericos —parsear un rango es interpretacion, no "
         "transcripcion—. Todo elemento que el codigo imprima y que esta hoja "
         "no tabule va ENTERO a 'Otros elementos'."),
        ("DB_SecII_Traccion", secii.normalizar_traccion, secii.COLS_TRACCION,
         "REQUISITOS DE TRACCION NORMALIZADOS — Seccion II, partes A, B y C",
         "Una fila por (especificacion, tabla, grado). La DOBLE UNIDAD es del "
         "codigo, no nuestra: `48 000 [330]` se conserva entero tal como lo "
         "imprime la tabla, nunca convertido.")):
        ws = new_sheet_secii(wb, hoja, titulo, fuente)
        extra = ["Otros elementos"] if cols_val is secii.COLS_QUIMICA else []
        _hdr_secii(ws, ["Parte", "Especificacion", "Tabla",
                        "Titulo de la tabla", "Grado / designacion",
                        "Orientacion"] + list(cols_val) + extra +
                   ["Pagina PDF", "Folio impreso", "Bloque", "Confianza"])
        r = R_DATA_SECII
        n_tablas = 0
        for p in datos:
            for s in p["specs"]:
                for t in s["tablas"]:
                    filas, _motivo = fn(t)
                    if not filas:
                        continue
                    n_tablas += 1
                    for f in filas:
                        tz = f["_traza"]
                        vals = ([p["parte"], s["id"], t["n"], t["titulo"],
                                 f["_grado"], f["_orientacion"]]
                                + [f.get(c, "") for c in cols_val]
                                + ([f.get("_otros", "")] if extra else [])
                                + [_pdf_1based(tz["pagina"]), tz["folio"],
                                   tz["bloque"], tz["confianza"]])
                        for j, v in enumerate(vals, start=1):
                            if isinstance(v, int) and not isinstance(v, bool):
                                ws.cell(r, j, v).font = DATA_F
                            else:
                                _txt_celda(ws, r, j, v)
                        r += 1
        last = r - 1
        ancho = {get_column_letter(i): w for i, w in
                 enumerate([14, 20, 7, 40, 26, 13], start=1)}
        autosize(ws, ancho)
        ws.auto_filter.ref = (f"A{R_HDR_SECII}:"
                              f"{get_column_letter(6 + len(cols_val) + len(extra) + 4)}"
                              f"{max(last, R_DATA_SECII)}")
        resultados[hoja] = (last, n_tablas)
        record_meta(hoja, "Normalizada desde las tablas de A/B/C",
                    "ASME_BPVC/Sec_II/*/specifications/*.json", "2025", "-",
                    max(last - R_DATA_SECII + 1, 0),
                    f"{n_tablas} tablas normalizadas. El resto se queda en el "
                    f"volcado integro y IDX_SecII_Tablas dice por que.")
    return resultados


def build_secii(res, wb):
    """Las nueve hojas de la Seccion II. Devuelve el resumen para el Dashboard."""
    datos = _cargar_secii(res)
    ncols_max = max((t["ncols"] for p in datos for s in p["specs"]
                     for t in s["tablas"]), default=1)
    n_tablas = sum(len(s["tablas"]) for p in datos for s in p["specs"])
    n_filas = sum(len(t["filas"]) for p in datos for s in p["specs"]
                  for t in s["tablas"])
    conf = Counter(f["confianza"] for p in datos for s in p["specs"]
                   for t in s["tablas"] for f in t["filas"])
    # La comprobacion sin perdida se corre AQUI tambien, no solo en el CLI: lo
    # que se graba en el libro tiene que ser lo mismo que se midio. Un fallo
    # aborta el build; escribir una fila que perdio texto seria peor que no
    # escribirla.
    malas = []
    for p in datos:
        for s in p["specs"]:
            for t in s["tablas"]:
                malas += secii.verificar_sin_perdida(t["filas"])
    if malas:
        raise SystemExit(
            f"ERROR: {len(malas)} filas de la Seccion II no pasan la "
            f"comprobacion sin perdida. El libro no se construye con filas que "
            f"no sean una reparticion exacta del texto impreso.\n  "
            + "\n  ".join(malas[:5]))

    build_cat_secii(wb, datos)
    build_idx_secii(wb, datos)
    lasts = build_db_secii(wb, datos, ncols_max)
    build_notas_secii(wb, datos)
    norm = build_normalizadas_secii(wb, datos)

    n_notas = sum(len(t["notas"]) for p in datos for s in p["specs"]
                  for t in s["tablas"])
    huecos = Counter()
    for p in datos:
        for s in p["specs"]:
            huecos += s["huecos"]
    ISSUES.append(
        f"Seccion II A/B/C: {n_tablas} tablas logicas y {n_filas} filas "
        f"volcadas ({conf[secii.EXACTA]} EXACTA · {conf[secii.POR_CONTEO]} POR "
        f"CONTEO · {conf[secii.AMBIGUA]} AMBIGUA), {n_notas} notas al pie. "
        f"Comprobacion sin perdida: 0 fallos sobre las {n_filas} filas.")
    ISSUES.append(
        "Seccion II A/B/C AMBIGUA: la fila conserva su texto impreso ENTERO en "
        "una celda y se marca; no se reparte por interpolacion sobre el bbox "
        "porque los `Span` no estan en el JSON y la posicion de cada palabra "
        "dentro de un `Line` no consta. No se pierde texto; no queda tabulada.")
    ISSUES.append(
        f"Seccion II A/B/C normalizadas: "
        f"{norm['DB_SecII_Quimica'][1]} tablas en DB_SecII_Quimica y "
        f"{norm['DB_SecII_Traccion'][1]} en DB_SecII_Traccion, de {n_tablas}. "
        f"Solo entra la tabla cuyos encabezados se resuelven ENTEROS contra el "
        f"vocabulario del codigo; el resto sigue integro en el volcado y "
        f"IDX_SecII_Tablas dice por que, tabla a tabla.")
    if huecos:
        ISSUES.append("Seccion II A/B/C huecos declarados (no reparables sin el "
                      "PDF): " + " · ".join(f"{k}={v}" for k, v in
                                            sorted(huecos.items())))
    return dict(tablas=n_tablas, filas=n_filas, notas=n_notas, conf=conf,
                ncols_max=ncols_max, lasts=lasts, norm=norm)


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

CARD_FILL = PatternFill("solid", fgColor=PAPEL_2)
KPI_FILL = PatternFill("solid", fgColor=PAPEL_2)
CARD_BORDER = BOX
LBL2_F = Font(name=MONO, size=9, bold=True, color=TINTA)
VAL_F = Font(name=MONO, size=11, bold=True, color=TINTA)
UNIT_F = Font(name=MONO, size=9, color=GRIS)
KPI_TIT_F = Font(name=MONO, size=9, bold=True, color=PAPEL)
# Cifra de KPI: la macrotipografia del sistema. El contraste de escala con el
# rotulo mono de 9 pt es lo que da la jerarquia; Arial Black no lleva bold
# —Excel lo sintetizaria y engorda el trazo— y baja de 18 a 16 pt porque es una
# tipografia mucho mas ancha que la que habia y el KPI de estado publica una
# frase, no solo una cifra.
KPI_VAL_F = Font(name=MACRO, size=16, color=TINTA)

# Celda unica de escritura (temperatura de consulta). El campo va en PAPEL
# limpio como cualquier otro, y lo que lo distingue de las listas desplegables
# es la CAJA ROJA: es el unico rojo de la zona de seleccion y el unico borde
# distinto del de tinta, asi que salta a la vista sin depender de un relleno de
# color que en este sistema no existe.
TEMP_INPUT_FILL = PatternFill("solid", fgColor=PAPEL)
CAJA_TECLEO = Border(*[Side("medium", color=ROJO)] * 4)
# Semaforo de cascada completa/incompleta (formato condicional sobre el
# indicador de seleccion, ver build_buscador/finish_buscador). Bloque macizo con
# tinta encima: el color lo pone el relleno, nunca el texto.
SEL_OK_FILL = PatternFill("solid", fgColor=VERDE)
SEL_OK_FONT = Font(name=MONO, size=10, bold=True, color=TINTA)
SEL_BAD_FILL = PatternFill("solid", fgColor=AMBAR)
SEL_BAD_FONT = Font(name=MONO, size=10, bold=True, color=TINTA)

# Semaforo de ACEPTACION de los dos motores de calculo (Fase 6). Mismo mecanismo
# que el de arriba —formato condicional, bloque macizo, el color lo pone el
# relleno— pero sobre la columna de Resultado de las verificaciones y sobre el
# dictamen global. Es la tercera excepcion declarada al acento unico, y la que
# mas derecho tiene a serlo: en un motor de calculo el criterio de aceptacion es
# informacion de seguridad, y la diferencia entre "cumple" y "no cumple" tiene
# que leerse sin leerse.
#
# El ROJO aparece aqui como RELLENO por primera vez fuera del aviso de macros, y
# con el mismo significado que tiene en todo el libro: bloqueado. Va en formato
# condicional (dxf), no como estilo de celda, asi que el guardia
# test_el_aviso_de_macros_es_el_unico_relleno_rojo sigue valiendo tal cual.
CUMPLE_OK_FILL = PatternFill("solid", fgColor=VERDE)
CUMPLE_OK_FONT = Font(name=MONO, size=10, bold=True, color=TINTA)
CUMPLE_BAD_FILL = PatternFill("solid", fgColor=ROJO)
CUMPLE_BAD_FONT = Font(name=MONO, size=10, bold=True, color=PAPEL)
# El dictamen global va en macrotipografia: es la frase que se lee primero.
DICTAMEN_OK_FONT = Font(name=MACRO, size=16, color=TINTA)
DICTAMEN_BAD_FONT = Font(name=MACRO, size=16, color=PAPEL)
DICTAMEN_ESPERA_FILL = PatternFill("solid", fgColor=AMBAR)
DICTAMEN_ESPERA_FONT = Font(name=MACRO, size=16, color=TINTA)


def semaforo_resultado(ws, rango, celda, favorables):
    """Verde si la celda de Resultado dice algo favorable; rojo si no.

    `favorables` son los textos que cuentan como aceptacion. No siempre es
    "CUMPLE": dos de las verificaciones del Art. 212 no resuelven en
    cumple/no cumple sino en una RUTA de reparacion —"Refuerzo 360" frente a
    "Parche local", "OK - parche" frente a "Migrar (Art.206)"—, y ahi lo verde
    es la rama que deja seguir con el parche.

    La regla del rojo es el complemento (`<>` de todas las favorables) y no una
    lista de textos desfavorables: asi una celda vacia, un #N/A o un texto que
    nadie previo salen en ROJO, que es el lado seguro. Enumerar lo malo dejaria
    lo imprevisto en blanco, indistinguible de "aun no calculado".
    """
    ok = "OR(" + ",".join(f'{celda}="{t}"' for t in favorables) + ")"
    ws.conditional_formatting.add(
        rango, FormulaRule(formula=[ok], fill=CUMPLE_OK_FILL, font=CUMPLE_OK_FONT))
    ws.conditional_formatting.add(
        rango, FormulaRule(formula=[f'AND({celda}<>"",NOT({ok}))'],
                           fill=CUMPLE_BAD_FILL, font=CUMPLE_BAD_FONT))


def banda(ws, r, texto, n=NCOLS):
    """Banda de seccion: bloque de tinta a todo el ancho, rotulo en
    macrotipografia enmarcado en ASCII y franja roja de cierre por abajo."""
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
    c = ws.cell(r, 1, rotulo(texto))
    c.font, c.fill = TITLE_F, BAND_FILL
    c.alignment = Alignment(vertical="center", indent=1)
    # El bloque de tinta lo cubre la celda ancla; la franja roja hay que
    # recorrerla celda a celda aunque el rango este fusionado (ver franja()).
    franja(ws, r, 1, n)
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


def formula_funcion_tabla(tabla_expr, tags):
    """Rotulo en español de PARA QUE sirve la tabla de la fila activa.

    Un buscador que reune dos tablas del codigo —A-1 (esfuerzos basicos en
    traccion) y A-4 (perneria)— tiene que decir cual de las dos resolvio la
    cascada Y que publica cada una: leer un admisible de perneria creyendo que
    es el de un tubo cambia el calculo. El texto es una DESCRIPCION del titulo
    impreso de la tabla, no un valor normativo, y por eso vive en el builder.

    `tags` son SOLO las tablas que esa base contiene de verdad. Emitir las
    nueve en los cinco buscadores dejaria ocho ramas muertas por motor y una
    formula tres veces mas larga sin ganar nada. Si una tabla nueva llegase sin
    descripcion, aborta el build en vez de imprimir un rotulo generico.

    Se genera en una funcion —y no suelta dentro del motor— porque
    test_dashboard.py comprueba el literal contra el mismo diccionario.
    """
    faltan = [t for t in tags if t not in FUNCION_TABLA]
    if faltan:
        raise SystemExit(
            f"No hay descripcion en FUNCION_TABLA para {faltan}. Anadala: la "
            "ficha del buscador tiene que decir que publica cada tabla.")
    expr = f'"{FUNCION_TABLA_DESCONOCIDA}"'
    for tag in sorted(tags, reverse=True):
        expr = f'IF({tabla_expr}="{tag}","{FUNCION_TABLA[tag]}",{expr})'
    return expr


# Que publica cada tabla del codigo, en español. Solo describe el titulo
# impreso; no aporta ni un valor. Las ediciones US llevan la MISMA descripcion
# que su gemela metrica: lo que cambia entre A-1 y A-1C es la unidad, no la
# funcion de la tabla.
FUNCION_TABLA_DESCONOCIDA = "tabla no descrita en este buscador"
FUNCION_TABLA = {
    "A-1": "Esfuerzos basicos admisibles en traccion para METALES (tuberia, "
           "placa, forja, fundicion)",
    "A-1C": "Esfuerzos basicos admisibles en traccion para METALES (tuberia, "
            "placa, forja, fundicion)",
    "A-4": "Esfuerzos de diseno para materiales de PERNERIA (pernos, esparragos "
           "y tuercas de union bridada)",
    "A-4C": "Esfuerzos de diseno para materiales de PERNERIA (pernos, esparragos "
            "y tuercas de union bridada)",
    "1A": "Esfuerzos admisibles de materiales FERROSOS",
    "1B": "Esfuerzos admisibles de materiales NO FERROSOS",
    "3": "Esfuerzos admisibles de PERNERIA",
    "U": "Resistencia a la traccion minima especificada Su frente a la temperatura",
    "Y-1": "Limite de fluencia minimo especificado Sy frente a la temperatura",
}


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
    # El subtitulo describe QUE publica cada tabla del buscador y en los que
    # reunen dos tablas se alarga; la altura se ajusta al texto en vez de
    # truncarlo (el ancho util del panel es de unos 190 caracteres).
    ws.row_dimensions[2].height = max(
        26, 13 * math.ceil(len(ws.cell(2, 1).value) / 190))
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
        c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
        ws.row_dimensions[r].height = 17
    e = _mrg(ws, 12, 1, 3, "TEMPERATURA DE CONSULTA  >>>  SE TECLEA")
    e.font = Font(name=MONO, size=10, bold=True, color=ROJO)
    e.alignment = Alignment(vertical="center", indent=1)
    com_temp = ("Entrada: la UNICA celda de escritura libre de todo el buscador. "
               "Temperatura a la que se necesita el valor, en la unidad que muestra la "
               "celda de la derecha (segun el selector SI/US de arriba). Todo el "
               "resultado, la ficha tecnica y la curva se recalculan a partir de este "
               "valor.")
    _nota(e, com_temp)
    c = _mrg(ws, 12, 4, 5, 25)
    c.font, c.fill = IN_F, TEMP_INPUT_FILL
    c.border = CAJA_TECLEO
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
    c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
    ay = _mrg(ws, 5, 7, NCOLS)
    ay.value = (f'=IF($D$5="SI","Leyendo {master["sheet"]} — valores en {unit_si}, '
                f'temperatura en {temp_si}","Leyendo {us["sheet"]} — valores en {unit_us}, '
                f'temperatura en {temp_us}")')
    ay.font = Font(name=MONO, size=9, color=GRIS)
    _nota(ay, "Aviso automatico: confirma que hoja base de datos y que unidades esta "
              "leyendo el buscador, segun el selector 'Sistema de unidades' (D5). No "
              "se edita.")
    ay2 = _mrg(ws, 12, 7, 9)
    ay2.value = '=IF($D$10="","SELECCION INCOMPLETA","SELECCION COMPLETA")'
    ay2.font = SEL_BAD_FONT
    ay2.fill = SEL_BAD_FILL
    ay2.border = BOX
    ay2.alignment = Alignment(horizontal="center", vertical="center")
    _nota(ay2, "Aviso automatico: SELECCION COMPLETA (bloque verde) si la cascada de "
               "seleccion (pasos 0 a 4) esta completa; SELECCION INCOMPLETA (bloque "
               "ambar) si falta elegir algun nivel para poder mostrar un resultado. "
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

    # La tabla de la fila activa, una sola vez. La ficha la usa dos veces —para
    # el rotulo y para la funcion de la tabla— y repetir el INDEX dentro de un
    # anidamiento de IF multiplicaba la formula por el numero de tablas.
    # Fila 20: las 4 a 11 son de esta funcion y las 12 a 19 las escribe
    # finish_buscador (n_pts, p1, T1/S1/T2/S2, S(T) y estado).
    ws.cell(20, A, "tabla activa").font = SRC_F
    ws.cell(20, A + 1).value = f'=IF({FIL}="","",INDEX({ident("Tabla")},{FIL}))'
    TAB = f"${LA}$20"

    # Los rotulos de tabla que esta base contiene de verdad, leidos de la propia
    # base: asi la ficha no ofrece ramas de tablas que este buscador no tiene.
    tags = sorted({d["tabla"] for d in master["recs"]} |
                  {d["tabla"] for d in us["recs"]})

    return dict(ws=ws, pref=pref, rng=rng, master=master, us=us, A=A,
                MIDC=MIDC, FSI=FSI, FUS=FUS, FIL=FIL, NVAR=NVAR, ident=ident,
                TAB=TAB, tags=tags,
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
    mid.font = Font(name=MACRO, size=11, color=TINTA)
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
        # wrap_text porque el KPI de ESTADO publica una frase, no una cifra, y
        # la macrotipografia es mucho mas ancha que la que habia: sin envolver,
        # «FUERA DE RANGO — T supera la Temp. max.» se cortaba.
        v.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        _nota(v, com)
        u = _mrg(ws, 19, c1, c1 + 2, uni)
        u.font, u.fill = UNIT_F, KPI_FILL
        u.alignment = Alignment(horizontal="center")
        for r2 in (17, 18, 19):
            for cc in range(c1, c1 + 3):
                ws.cell(r2, cc).border = CARD_BORDER
    ws.row_dimensions[17].height = 26
    ws.row_dimensions[18].height = 38
    ws.row_dimensions[19].height = 14
    ws.conditional_formatting.add(
        f"A18:L19",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("FUERA DE RANGO",{EST}))'],
                    font=Font(color=ROJO, bold=True)))

    # --------------------- 3 · FICHA TECNICA ------------------------------
    banda(ws, 21, "3 · FICHA TECNICA DEL MATERIAL   —   identificacion tal como la "
                  "organiza el codigo")

    def val_of(colname):
        idx = f'INDEX({ctx["ident"](colname)},{FIL})'
        return f'=IF({FIL}="","—",IF({idx}="","—",{idx}))'

    izq = [("Tabla del codigo", val_of("Tabla"), None,
            "Calculo: tabla del codigo (A-1/A-4, 1A, 1B/3, U o Y-1 segun el buscador) "
            "de la que proviene la fila resuelta por la cascada de seleccion."),
           # Un buscador que reune dos tablas tiene que decir cual resolvio la
           # cascada Y que publica: el admisible de un perno (A-4) y el de un
           # tubo (A-1) se leen igual y no son lo mismo.
           ("Funcion de la tabla",
            f'=IF({FIL}="","—",'
            + formula_funcion_tabla(ctx["TAB"], ctx["tags"]) + ')', None,
            "Calculo: que publica esa tabla del codigo, en español. Es una "
            "descripcion de su titulo impreso, no un dato normativo: sirve para no "
            "confundir el esfuerzo de un perno (A-4/Tabla 3) con el de un "
            "componente a presion (A-1/Tabla 1A)."),
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
            # `campo` fija 17 px, que basta para un valor corto. La funcion de la
            # tabla es una frase y se parte en tres lineas dentro de dos columnas.
            if izq[k][0] == "Funcion de la tabla":
                ws.row_dimensions[r2].height = 44
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
    ch.title = f"{ctx['valor_lbl']} frente a la temperatura".upper()
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
    # como trazo, no como nube de puntos. En TINTA, como el resto del sistema:
    # el rojo esta reservado al dato consultado y a los avisos.
    s1.marker = Marker(symbol="none")
    # Linea recta (no suavizada) entre puntos: la interpolacion del codigo es
    # lineal (seccion 4), una curva suavizada la representaria mal.
    s1.smooth = False
    s1.graphicalProperties = GraphicalProperties(
        ln=LineProperties(solidFill=TINTA_2, w=19050))
    ch.series.append(s1)
    xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
    yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
    s2 = Series(yq, xq, title_from_data=True)
    s2.marker = Marker(symbol="diamond", size=10,
                       spPr=GraphicalProperties(
                           solidFill=ROJO, ln=LineProperties(solidFill=ROJO)))
    s2.graphicalProperties.line = LineProperties(noFill=True)
    ch.series.append(s2)
    estilizar_chart(ch)
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
    c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
    dv_list(ws, "D5", '"SI,US"', com_sist)
    e = _mrg(ws, 6, 1, 3, "TEMPERATURA DE CONSULTA  >>>  SE TECLEA")
    e.font = Font(name=MONO, size=10, bold=True, color=ROJO)
    e.alignment = Alignment(vertical="center", indent=1)
    com_temp = ("Entrada: la UNICA celda de escritura libre de toda la hoja, en la "
               "unidad que muestra la celda de la derecha (segun el selector 'Sistema "
               "de unidades' de arriba). Se aplica por igual a los bloques de abajo "
               "(modulo E, dilatacion y Poisson/densidad).")
    _nota(e, com_temp)
    c = _mrg(ws, 6, 4, 5, 25)
    c.font, c.fill = IN_F, TEMP_INPUT_FILL
    c.border = CAJA_TECLEO
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
    c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
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
        c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
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
        c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
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
            v.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
            _nota(v, f'Calculo: {b["valor_lbl"]} interpolado (o tabulado-conservador, '
                     'segun el Modo de lectura de D7) a la temperatura de consulta '
                     '(D6), entre los puntos tabulados T1/T2 de abajo, para el grupo '
                     'elegido arriba y la edicion elegida en D5.')
            u = _mrg(ws, r, 6, 8)
            u.value = f"={u_val}"
            u.font, u.fill = UNIT_F, KPI_FILL
            ws.row_dimensions[r].height = 30
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
            ch.title = f'{b["valor_lbl"]} frente a la temperatura'.upper()
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
            # Linea continua sin marcadores, en tinta — mismo estilo que
            # finish_buscador (ver estilizar_chart).
            se.marker = Marker(symbol="none")
            se.smooth = False
            se.graphicalProperties = GraphicalProperties(
                ln=LineProperties(solidFill=TINTA_2, w=19050))
            ch.series.append(se)
            xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
            yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
            sq = Series(yq, xq, title_from_data=True)
            sq.marker = Marker(symbol="diamond", size=10,
                               spPr=GraphicalProperties(
                                   solidFill=ROJO,
                                   ln=LineProperties(solidFill=ROJO)))
            sq.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
            ch.series.append(sq)
            estilizar_chart(ch)
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
        c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
        ws.row_dimensions[r].height = 17

    # Celda unica de escritura: mismo campo de papel que las listas, pero con la
    # caja en ROJO (regla de estilo 1 del proyecto).
    com_temp = ("Entrada: la UNICA celda de escritura libre de todo el motor. "
                "Temperatura a la que se necesita la propiedad, en la unidad de la "
                "celda de la derecha. Las propiedades de tipo PUNTO (C-2 y C-4) no "
                "dependen de la temperatura y el ESTADO lo avisa.")
    e = _mrg(ws, 11, 1, 3, "TEMPERATURA DE CONSULTA  >>>  SE TECLEA")
    e.font = Font(name=MONO, size=10, bold=True, color=ROJO)
    e.alignment = Alignment(vertical="center", indent=1)
    _nota(e, com_temp)
    c = _mrg(ws, 11, 4, 5, 350)
    c.font, c.fill = IN_F, TEMP_INPUT_FILL
    c.border = CAJA_TECLEO
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
    c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO

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
    ay.font = Font(name=MONO, size=9, color=GRIS)
    ay.alignment = Alignment(vertical="center", wrap_text=True)
    _nota(ay, "Aviso automatico: dice siempre que tabla del Apendice C y que edicion "
              "esta leyendo el motor. No se edita.")

    # Semaforo de cascada. El disparador es $D$9 (nivel 3 · Material) porque el
    # nivel 4 solo tiene contenido real en C-1 y en tres filas de C-2.
    ay2 = _mrg(ws, 11, 7, 9)
    ay2.value = '=IF($D$9="","SELECCION INCOMPLETA","SELECCION COMPLETA")'
    ay2.font, ay2.fill, ay2.border = SEL_BAD_FONT, SEL_BAD_FILL, BOX
    ay2.alignment = Alignment(horizontal="center", vertical="center")
    _nota(ay2, "Aviso automatico: SELECCION COMPLETA (bloque verde) cuando la cascada "
               "llega al material; SELECCION INCOMPLETA (bloque ambar) mientras falte "
               "un paso.")
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
                    font=Font(color=ROJO, bold=True)))

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
    av.font = Font(name=MONO, size=9, bold=True, color=AMBAR_TXT)
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
    ch.title = "PROPIEDAD TABULADA FRENTE A LA TEMPERATURA"
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
    # Linea continua sin marcadores, en tinta (regla de estilo 3).
    s1.marker = Marker(symbol="none")
    s1.smooth = False
    s1.graphicalProperties = GraphicalProperties(
        ln=LineProperties(solidFill=TINTA_2, w=19050))
    ch.series.append(s1)
    xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
    yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
    s2 = Series(yq, xq, title_from_data=True)
    s2.marker = Marker(symbol="diamond", size=10,
                       spPr=GraphicalProperties(
                           solidFill=ROJO, ln=LineProperties(solidFill=ROJO)))
    s2.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.series.append(s2)
    estilizar_chart(ch)
    ws.add_chart(ch, f"A{rg + 2}")

    for cc, w in zip("ABCDEFGHIJKL",
                     [22, 20, 16, 16, 10, 3, 22, 20, 16, 16, 10, 3]):
        ws.column_dimensions[cc].width = w
    return ws


# ---------------------------------------------------------------------------
# Buscar_B31_B1 — esfuerzo de diseno hidrostatico del Apendice B (Tabla B-1)
# ---------------------------------------------------------------------------
# Motor propio, por la misma razon por la que el Apendice C tiene el suyo: el
# HDS es lo unico del Apendice B tabulado frente a la temperatura, y una ficha
# campo/valor no puede resolver una consulta a una T cualquiera.
#
# Lo que separa este motor de los cinco de esfuerzos —y por que no reutiliza
# build_buscador/finish_buscador— son las tres reglas de rango, que en el
# Capitulo VII son OTRAS y estan citadas una a una:
#
#   · para. A302.3.1(b) — «The stresses and allowable pressures are grouped by
#     materials and listed for stated temperatures. Straight-line interpolation
#     between temperatures is permissible.» -> se interpola, igual que en los
#     metales, y se puede elegir el modo tabulado-conservador.
#   · Nota (3) de la propia Tabla B-1/B-1C, anclada a la columna de 23 °C
#     (73 °F) — «Use these hydrostatic design stress (HDS) values at all lower
#     temperatures.» -> por debajo de la primera T tabulada NO se extrapola: se
#     SOSTIENE ese valor, y el estado lo dice citando la nota. Concuerda con
#     para. A323.2.2(b).
#   · para. A323.2.1(a) -> no se usa un material por encima de la maxima
#     temperatura para la que hay valor o rating, y las Notas (1) y (2) fijan
#     los limites recomendados. -> se bloquea por ARRIBA en dos sitios: el
#     limite maximo recomendado que imprime la fila y el ultimo punto tabulado.
#     El estado distingue los dos casos, porque no son el mismo aviso.
#
# El limite POR ABAJO es el minimo recomendado impreso, no el primer punto
# tabulado: entre uno y otro la Nota (3) sigue dando un valor valido.
#
# Columnas ocultas propias: 36..39 (listas de los niveles 1 a 4) y 41..42
# (auxiliares). Son las mismas que usan Buscar_Prop_B31_3 y los dos motores de
# factores, pero cada uno en SU hoja; quedan por debajo de COL_CLAVE_BASE (66).
APXB1_LST_COL = 36
APXB1_AUX_COL = 41
APXB1_CURVA_COL = 500      # columna de arranque en _Curvas (libres desde 500)


def _rngb1(info, colname):
    """Rango de una columna de la base de la Tabla B-1 (layout APXB1_COLS)."""
    L = CLB1[colname]
    return f"{info['sheet']}!${L}${R_DATA}:${L}${info['last_row']}"


def formula_estado_b1(fil, npts, tmin, tmax, tprim, tult, tq):
    """Estado del rango de la Tabla B-1. Cuatro bloqueos y una excepcion.

    El orden de las comprobaciones importa y no es arbitrario: primero los
    limites de temperatura RECOMENDADOS que imprime la fila (Notas (1) y (2)),
    que son un limite de servicio del material, y solo despues el extremo de la
    banda tabulada, que es un limite del dato. Un material fuera de su limite
    recomendado no se salva porque exista un HDS tabulado ahi.

    La Nota (3) es la unica excepcion al «no extrapolar» de la regla 4 del
    proyecto, y la escribe el codigo: por debajo de la primera temperatura
    tabulada el HDS se SOSTIENE, no se extrapola hacia abajo.

    Se genera en una funcion porque verificar.py §11 la vuelve a emitir para
    recalcularla en Excel: asi la prueba ejerce el original y no una copia.
    """
    return (f'=IF({fil}="","{EST_B1_SIN_SEL}",'
            f'IF({npts}=0,"{EST_B1_SIN_TAB}",'
            f'IF(AND({tmin}<>"",{tq}<{tmin}),"{EST_B1_BAJO}",'
            f'IF(AND({tmax}<>"",{tq}>{tmax}),"{EST_B1_ALTO_LIM}",'
            f'IF({tq}>{tult},"{EST_B1_ALTO_TAB}",'
            f'IF({tq}<{tprim},"{EST_B1_NOTA3}",'
            f'"{EST_B1_OK}"))))))')


def formula_valor_b1(fil, est, vtab):
    """HDS resuelto. Bloqueado en cuanto el estado diga FUERA DE RANGO.

    No hay factor de escala: la Tabla B-1 publica el HDS directamente en MPa
    (ksi en B-1C). El valor sostenido de la Nota (3) ya lo devuelve
    interp_value(), que por debajo del primer punto entrega ese primer valor.
    """
    return (f'=IF({fil}="","",'
            f'IF(ISNUMBER(SEARCH("FUERA DE RANGO",{est})),"{VAL_B1_BLOQUEADO}",'
            f'IF({est}="{EST_B1_SIN_TAB}","{EST_B1_SIN_TAB}",{vtab})))')


def build_buscador_b1(wb, curvas, rng, master, us):
    name = "Buscar_B31_B1"
    hl = get_column_letter
    ws = new_sheet(
        wb, name,
        "BUSCADOR — ESFUERZO DE DISENO HIDROSTATICO (HDS) DE TUBERIA "
        "TERMOPLASTICA · ASME B31.3-2024, Apendice B, Tablas B-1 y B-1C",
        "La Tabla B-1 publica DOS cosas por fila y no son lo mismo: el esfuerzo "
        "de diseno hidrostatico HDS —el que entra como S en la eq. (26a) del "
        "para. A304.1.2, t = PD/(2S+P)— tabulado a cuatro temperaturas, y los "
        "LIMITES DE TEMPERATURA RECOMENDADOS de las Notas (1) y (2), que son un "
        "limite de servicio del material. El motor los muestra por separado y "
        "bloquea con los dos. Conmutador SI/US: cambia de HOJA (B-1 <-> B-1C), "
        "nunca convierte. La UNICA celda que se escribe es la temperatura.")
    ws.freeze_panes = "A4"
    _mrg(ws, 1, 1, NCOLS)
    _mrg(ws, 2, 1, NCOLS)
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 56

    SEL = '$D$5="SI"'

    def ident(colname):
        return f'IF({SEL},{_rngb1(master, colname)},{_rngb1(us, colname)})'

    # ------------------------- 1 · SELECCION -------------------------------
    banda(ws, 4, "1 · SELECCION DEL TERMOPLASTICO   —   todo por lista desplegable")
    com_sist = ("Entrada: SI lee la Tabla B-1 (HDS en MPa, temperaturas en °C); US "
                "lee la Tabla B-1C (ksi, °F). El codigo publica UNA TABLA POR "
                "EDICION y el conmutador cambia de hoja: nunca hay conversion "
                "(regla 9 del proyecto).")
    filas = [
        (5, "Sistema de unidades", "SI", com_sist),
        (6, "0 · Designacion de material", "",
         "Entrada: paso 0 de la cascada. Designacion del termoplastico tal como la "
         "imprime la tabla (ABS, CPVC4120, PE4710, PVC1120, PP-R...). Habilita el "
         "paso 1."),
        (7, "1 · Spec. No. (ASTM)", "",
         "Entrada: paso 1, dependiente del paso 0. La misma designacion de material "
         "aparece bajo VARIAS especificaciones con HDS distinto —PE2708 esta en "
         "D2737, D3035 y F714—, asi que este paso no es decorativo: es el que "
         f"separa esas filas. La fila del ABS no imprime Spec. No. y ofrece "
         f"«{SIN_SPEC_B1}»."),
        (8, "2 · Designacion de tuberia", "",
         "Entrada: paso 2, dependiente de los pasos 0 y 1. Pipe Designation impresa "
         "(SDR11, Sch. 40, 80, DR-PR...). Es lo que separa CPVC4120-05 de F441 y de "
         "F442."),
        (9, "3 · Cell Class", "",
         "Entrada: paso 3, dependiente de los pasos 0 a 2. Cell Class impresa por la "
         f"tabla. Las filas que no la imprimen ofrecen «{SIN_CELL_B1}»."),
        (10, "4 · Variante", "",
         "Entrada: paso 4, dependiente de los pasos 0 a 3. Solo tiene contenido real "
         "si el codigo repitiese una fila con identificacion identica; hoy vale "
         f"«{VAR_UNICA}» en las 36 filas de las dos ediciones."),
    ]
    for r, et, val, com in filas:
        e = _mrg(ws, r, 1, 3, et)
        e.font = LBL_F
        e.alignment = Alignment(vertical="center", indent=1)
        _nota(e, com)
        c = _mrg(ws, r, 4, 6, val)
        c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
        ws.row_dimensions[r].height = 17

    # Celda unica de escritura: relleno lavanda propio (regla de estilo 1).
    com_temp = ("Entrada: la UNICA celda de escritura libre de todo el motor. "
                "Temperatura de diseno del componente (para. A301.3.2), en la unidad "
                "de la celda de la derecha. El para. A302.2.4(a) no admite margen "
                "por variaciones de presion o temperatura en tuberia no metalica: "
                "aqui va la condicion mas severa.")
    e = _mrg(ws, 11, 1, 3, "TEMPERATURA DE CONSULTA  >>>  SE TECLEA")
    e.font = Font(name=MONO, size=10, bold=True, color=ROJO)
    e.alignment = Alignment(vertical="center", indent=1)
    _nota(e, com_temp)
    c = _mrg(ws, 11, 4, 5, 23)
    c.font, c.fill = IN_F, TEMP_INPUT_FILL
    c.border = CAJA_TECLEO
    _nota(c, com_temp)
    u = ws.cell(11, 6)
    u.value = f'=IF({SEL},"°C","°F")'
    u.font = UNIT_F
    _nota(u, "Calculo: unidad de la temperatura de consulta; cambia entre °C y °F "
             "segun el selector 'Sistema de unidades' (D5).")

    com_modo = ("Entrada: 'Interpolado' aplica la interpolacion lineal recta que "
                "autoriza el para. A302.3.1(b) entre los dos puntos tabulados que "
                "rodean la temperatura de consulta. 'Tabulado-conservador' ignora la "
                "interpolacion y adopta el valor tabulado en T2 (el escalon "
                "superior), que da un HDS menor y por tanto un espesor mayor.")
    e = _mrg(ws, 12, 1, 3, "Modo de lectura")
    e.font = LBL_F
    e.alignment = Alignment(vertical="center", indent=1)
    _nota(e, com_modo)
    c = _mrg(ws, 12, 4, 6, "Interpolado")
    c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO

    dv_list(ws, "D5", '"SI,US"', com_sist)
    dv_list(ws, "D12", '"Interpolado,Tabulado-conservador"', com_modo)
    dv_list(ws, "D6", "=" + rng["FAM"], filas[1][3])
    niveles = [("D7", rng["CK"], rng["CV"], "$D$6", rng.get("maxC", 10), filas[2][3]),
               ("D8", rng["FK"], rng["FV"], '$D$6&"|"&$D$7', rng.get("maxF", 10),
                filas[3][3]),
               ("D9", rng["SK"], rng["SV"], '$D$6&"|"&$D$7&"|"&$D$8',
                rng.get("maxS", 10), filas[4][3]),
               ("D10", rng["GK"], rng["GV"], '$D$6&"|"&$D$7&"|"&$D$8&"|"&$D$9',
                rng.get("maxG", 10), filas[5][3])]
    for i, (cell, kr, vr, key, mx, com) in enumerate(niveles):
        col = APXB1_LST_COL + i
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
    A = APXB1_AUX_COL
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
             f'=IF({FSI}="","",INDEX({_rngb1(master,"clave_bi")},{FSI}))')
    FUS = aux(8, "fila US",
              f'=IF({BI}="","",IFERROR(MATCH({BI},{_rngb1(us,"clave_bi")},0),""))')
    FIL = aux(9, "fila activa", f'=IF({SEL},{FSI},{FUS})')
    NP = aux(10, "n_pts", f'=IF({FIL}="",0,IFERROR(IF({SEL},'
                          f'INDEX({rm["npts"]},{FIL}),INDEX({ru["npts"]},{FIL})),0))')
    T_ANC = f'IF({SEL},{rm["t_anchor"]},{ru["t_anchor"]})'
    V_ANC = f'IF({SEL},{rm["v_anchor"]},{ru["v_anchor"]})'
    TR = f'OFFSET({T_ANC},{FIL}-1,0,1,MAX(1,{NP}))'
    VR = f'OFFSET({V_ANC},{FIL}-1,0,1,MAX(1,{NP}))'

    # INDEX sobre una celda VACIA devuelve 0, no cadena vacia. Sin este
    # envoltorio, una fila que no imprime limite maximo —la mayoria de las de
    # PVC y PE— daria tmax = 0 y el motor bloquearia TODA temperatura por
    # encima de cero como si el codigo hubiese publicado ese limite.
    def vacio_si_vacio(expr):
        return f'IF({expr}="","",{expr})'

    def indexb(colname):
        return vacio_si_vacio(f'INDEX({ident(colname)},{FIL})')

    P1 = aux(11, "p1", f'=IF({FIL}="","",IFERROR(MATCH($D$11,{TR},1),1))')
    tabulados = [("T1", f"INDEX({TR},{P1})"), ("V1", f"INDEX({VR},{P1})"),
                 ("T2", f"INDEX({TR},{P1}+1)"), ("V2", f"INDEX({VR},{P1}+1)")]
    T1C, S1C, T2C, S2C = [
        aux(12 + k, lb, f'=IF({FIL}="","",IFERROR({vacio_si_vacio(expr)},""))')
        for k, (lb, expr) in enumerate(tabulados)]
    TMIN = aux(16, "T min. recomendada",
               f'=IF({FIL}="","",{indexb("Temp. min. recomendada")})')
    TMAX = aux(17, "T max. recomendada",
               f'=IF({FIL}="","",{indexb("Temp. max. recomendada")})')
    TPRI = aux(18, "T primera", f'=IF({FIL}="","",{indexb("T primera tabulada")})')
    TULT = aux(19, "T ultima", f'=IF({FIL}="","",{indexb("T ultima tabulada")})')
    # interp_value ya devuelve el PRIMER valor tabulado cuando T <= T1, que es
    # exactamente lo que manda la Nota (3) por debajo de la primera columna.
    VTAB = aux(20, "HDS tabulado",
               f'=IF({FIL}="","",' + interp_value(T1C, S1C, T2C, S2C,
                                                  "$D$11", "$D$12")[1:] + ')')
    EST = aux(21, "estado",
              formula_estado_b1(FIL, NP, TMIN, TMAX, TPRI, TULT, "$D$11"))
    VALOR = aux(22, "HDS resuelto", formula_valor_b1(FIL, EST, VTAB))
    for jj in (A, A + 1):
        ws.column_dimensions[hl(jj)].hidden = True

    ay = _mrg(ws, 5, 7, NCOLS)
    ay.value = (f'=IF({FIL}="",IF({SEL},'
                f'"Edicion metrica — Tabla B-1 ({master["sheet"]}) · HDS en MPa, T en °C",'
                f'"Edicion U.S. Customary — Tabla B-1C ({us["sheet"]}) · HDS en ksi, '
                f'T en °F"),'
                f'"Leyendo Table "&INDEX({ident("Tabla")},{FIL})&'
                f'" ("&$D$5&") — HDS en "&INDEX({ident("Unidad de HDS")},{FIL})&'
                f'", T en "&INDEX({ident("Unidad de temperatura")},{FIL}))')
    ay.font = Font(name=MONO, size=9, color=GRIS)
    ay.alignment = Alignment(vertical="center", wrap_text=True)
    _nota(ay, "Aviso automatico: dice siempre que tabla del Apendice B y que edicion "
              "esta leyendo el motor. No se edita.")

    # Semaforo. El disparador es $D$7 (nivel 1 · Spec. No.) y no el nivel 0:
    # la designacion de material sola NO identifica una fila —PE2708 esta en tres
    # especificaciones con HDS distinto— y decir COMPLETA ahi seria mentir.
    ay2 = _mrg(ws, 11, 7, 9)
    ay2.value = '=IF($D$7="","SELECCION INCOMPLETA","SELECCION COMPLETA")'
    ay2.font, ay2.fill, ay2.border = SEL_BAD_FONT, SEL_BAD_FILL, BOX
    ay2.alignment = Alignment(horizontal="center", vertical="center")
    _nota(ay2, "Aviso automatico: SELECCION COMPLETA (bloque verde) cuando la cascada "
               "llega a la especificacion; SELECCION INCOMPLETA (bloque ambar) mientras "
               "falte.")
    ws.conditional_formatting.add(
        "G11:I11", FormulaRule(formula=['$D$7<>""'], fill=SEL_OK_FILL, font=SEL_OK_FONT))
    ws.conditional_formatting.add(
        "G11:I11", FormulaRule(formula=['$D$7=""'], fill=SEL_BAD_FILL, font=SEL_BAD_FONT))

    # ------------------------- 2 · RESULTADO -------------------------------
    banda(ws, 14, "2 · RESULTADO DE LA CONSULTA")
    kpis = [(1, "HDS — ESFUERZO DE DISENO HIDROSTATICO", "=" + VALOR,
             f'=IF({FIL}="","",INDEX({ident("Unidad de HDS")},{FIL}))',
             "Calculo: HDS a la temperatura de consulta. Es el valor que entra como "
             "S en la eq. (26a) del para. A304.1.2: t = PD/(2S+P). BLOQUEADO si la "
             "temperatura cae fuera del limite recomendado o de la banda tabulada."),
            (4, "TEMPERATURA DE CONSULTA", "=$D$11", f'=IF({SEL},"°C","°F")',
             "Calculo: repite la temperatura tecleada en D11, junto al resultado."),
            (7, "MODO DE LECTURA", "=$D$12", '="segun A302.3.1(b)"',
             "Calculo: repite el modo elegido en D12. La interpolacion lineal recta "
             "la autoriza el para. A302.3.1(b)."),
            (10, "ESTADO DEL RANGO", "=" + EST,
             f'=IF({FIL}="","",IF(INDEX({ident("Notas del HDS")},{FIL})="","",'
             f'"HDS: Nota "&INDEX({ident("Notas del HDS")},{FIL}))&'
             f'IF(INDEX({ident("Notas de los limites")},{FIL})="",""," · limites: '
             f'Notas "&INDEX({ident("Notas de los limites")},{FIL})))',
             "Calculo: EN RANGO cuando T cae dentro de los limites recomendados y de "
             "la banda tabulada. FUERA DE RANGO por debajo del minimo recomendado, "
             "por encima del maximo recomendado o por encima del ultimo punto "
             "tabulado —tres avisos distintos porque son tres causas distintas—. "
             "EN RANGO · Nota (3) cuando T queda por debajo de la primera "
             "temperatura tabulada: ahi el codigo manda sostener ese valor, no "
             "extrapolar.")]
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
    ws.row_dimensions[15].height = 28
    ws.row_dimensions[16].height = 30
    ws.row_dimensions[17].height = 15
    ws.conditional_formatting.add(
        "A16:L17",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("FUERA DE RANGO",{EST}))'],
                    font=Font(color=ROJO, bold=True)))

    # ---------------------- 3 · FICHA DE LA FILA ---------------------------
    banda(ws, 19, "3 · FICHA DE LA FILA DEL CODIGO   —   cada campo con su unidad "
                  "impresa")

    def val_of(colname):
        idx = f'INDEX({ident(colname)},{FIL})'
        return f'=IF({FIL}="","—",IF({idx}="","—",{idx}))'

    izq = [("Tabla del codigo", val_of("Tabla"), None,
            "Calculo: B-1 en la edicion metrica y B-1C en la U.S. Customary. Son dos "
            "tablas distintas del codigo, no una convertida."),
           ("Funcion de la tabla",
            '="Esfuerzos de diseno hidrostatico (HDS) y limites de temperatura '
            'recomendados de TUBERIA TERMOPLASTICA"', None,
            "Calculo: que publica esta tabla del codigo, en español. Es una "
            "descripcion de su titulo impreso, no un dato normativo."),
           ("Designacion de material", val_of("Designacion de material"), None,
            "Calculo: designacion del termoplastico tal como la imprime la tabla."),
           ("Spec. No. (ASTM)", val_of("Spec. No. (ASTM)"), None,
            "Calculo: especificacion ASTM impresa. La misma designacion de material "
            "aparece bajo varias specs con HDS distinto."),
           ("Designacion de tuberia", val_of("Designacion de tuberia"), None,
            "Calculo: Pipe Designation impresa (SDR, Sch., DR-PR...). Puede diferir "
            "entre las dos ediciones: es una asimetria del codigo, conservada."),
           ("Cell Class", val_of("Cell Class"), '="ASTM"',
            "Calculo: Cell Class impresa por la tabla, cuando la fila la trae."),
           ("Variante", val_of("Variante"), None,
            "Calculo: nivel 4 de la cascada. «(unico)» salvo que el codigo repitiese "
            "una fila con identificacion identica.")]
    der = [("Edicion consultada",
            f'=IF({FIL}="","—",IF({SEL},"{master["sheet"]} — metrica",'
            f'IF({FUS}="","sin equivalente en la edicion US",'
            f'"{us["sheet"]} — U.S. Customary")))', None,
            "Calculo: confirma de que edicion sale la fila activa y avisa si no "
            "tuviese homologa en la otra."),
           ("Temp. min. recomendada", val_of("Temp. min. recomendada"),
            f'=IF({SEL},"°C","°F")',
            "Calculo: limite MINIMO de temperatura recomendado, Notas (1) y (2) de la "
            "tabla. Por debajo el resultado queda BLOQUEADO. No es la primera "
            "temperatura tabulada: entre las dos sigue habiendo HDS por la Nota (3)."),
           ("Temp. max. recomendada", val_of("Temp. max. recomendada"),
            f'=IF({SEL},"°C","°F")',
            "Calculo: limite MAXIMO recomendado, Notas (1) y (2). Por encima el "
            "resultado queda BLOQUEADO (para. A323.2.1). Muchas filas no lo imprimen: "
            "ahi manda el ultimo punto tabulado."),
           ("Primera T tabulada", val_of("T primera tabulada"),
            f'=IF({SEL},"°C","°F")',
            "Calculo: primera temperatura con HDS impreso (23 °C / 73 °F). Por debajo "
            "de ella la Nota (3) manda SOSTENER este valor."),
           ("Ultima T tabulada", val_of("T ultima tabulada"),
            f'=IF({SEL},"°C","°F")',
            "Calculo: ultima temperatura con HDS impreso para esta fila. Por encima el "
            "resultado queda BLOQUEADO: el codigo no publica valor y no se extrapola."),
           ("Encabezado impreso del HDS", val_of("Encabezado impreso del HDS"), None,
            "Calculo: el encabezado literal de la banda de HDS, con su unidad, tal "
            "como lo imprime el codigo."),
           ("Observacion de extraccion", val_of("Observacion"), None,
            "Calculo: asimetria entre ediciones o artefacto declarado que afecta a "
            "esta fila, con el folio impreso. Vacio si no hay ninguno.")]
    for k in range(max(len(izq), len(der))):
        r2 = 20 + k
        if k < len(izq):
            campo(ws, r2, 1, izq[k][0], izq[k][1], izq[k][2], com_val=izq[k][3])
            if izq[k][0] == "Funcion de la tabla":
                ws.row_dimensions[r2].height = 44
        if k < len(der):
            campo(ws, r2, 7, der[k][0], der[k][1], der[k][2], com_val=der[k][3])

    # ------------------ 4 · TRAZABILIDAD DEL CALCULO -----------------------
    rt = 20 + max(len(izq), len(der)) + 1
    banda(ws, rt, "4 · TRAZABILIDAD   —   puntos tabulados usados y regla aplicada")
    campo(ws, rt + 1, 1, "T1 — tabulada inferior", f"={T1C}", f'=IF({SEL},"°C","°F")',
          com_val="Calculo: temperatura tabulada inmediatamente inferior (o igual) a "
                  "la de consulta, tomada de la banda compacta de la fila.")
    campo(ws, rt + 2, 1, "HDS en T1", f"={S1C}",
          f'=IF({FIL}="","",INDEX({ident("Unidad de HDS")},{FIL}))',
          com_val="Calculo: HDS tabulado en T1, tal como lo imprime el codigo.")
    campo(ws, rt + 1, 7, "T2 — tabulada superior", f"={T2C}", f'=IF({SEL},"°C","°F")',
          com_val="Calculo: temperatura tabulada inmediatamente superior. Vacia si T1 "
                  "es el ultimo punto tabulado de la fila.")
    campo(ws, rt + 2, 7, "HDS en T2", f"={S2C}",
          f'=IF({FIL}="","",INDEX({ident("Unidad de HDS")},{FIL}))',
          com_val="Calculo: HDS tabulado en T2. Vacio si T1 es el ultimo punto.")
    ec = _mrg(ws, rt + 3, 1, NCOLS)
    ec.value = (
        f'=IF({FIL}="","",'
        f'IF({EST}="{EST_B1_SIN_TAB}",'
        f'"Esta fila no publica HDS a ninguna temperatura: el codigo solo le imprime '
        f'los limites de temperatura recomendados.",'
        f'IF(ISNUMBER(SEARCH("FUERA DE RANGO",{EST})),'
        f'"Resultado bloqueado: la temperatura cae fuera del limite recomendado o de '
        f'la banda tabulada. para. A323.2.1(a) y Notas (1) y (2) de la tabla.",'
        f'IF({EST}="{EST_B1_NOTA3}",'
        f'"Nota (3) de la Tabla B-1: se sostiene el HDS de la primera temperatura '
        f'tabulada para toda temperatura inferior. No se extrapola.",'
        f'IF($D$12="Tabulado-conservador",'
        f'"Modo tabulado-conservador: se adopta el HDS de T2, menor que el '
        f'interpolado, y por tanto conservador en el espesor.",'
        f'"Interpolacion lineal recta autorizada por el para. A302.3.1(b):  '
        f'S = S1 + (S2-S1)*(T-T1)/(T2-T1)")))))')
    ec.font = SRC_F
    ec.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[rt + 3].height = 26
    _nota(ec, "Calculo: dice en texto cual de los cinco casos aplico para obtener el "
              "KPI de la seccion 2, y con que parrafo del codigo.")
    fu = _mrg(ws, rt + 4, 1, NCOLS)
    fu.value = (f'=IF({FIL}="",'
                f'"Fuente: resources/{APX}/appendix_b/table_b_1.json y table_b_1c.json",'
                f'"Fuente: resources/{APX}/appendix_b/ · Table "&'
                f'INDEX({ident("Tabla")},{FIL})&" · fila impresa "&'
                f'INDEX({ident("Linea")},{FIL}))')
    fu.font = SRC_F
    fu.alignment = Alignment(vertical="center", indent=1)
    _nota(fu, "Trazabilidad: archivo de resources/ y numero de fila impresa de la que "
              "sale el valor mostrado.")
    av = _mrg(ws, rt + 5, 1, NCOLS)
    av.value = ('="El HDS es el esfuerzo de diseno de la eq. (26a): t = PD/(2S+P), '
                'para. A304.1.2(a). El para. A302.3.1(a) advierte que el uso del HDS '
                'para calculos distintos del diseno a presion NO esta verificado."')
    av.font = Font(name=MONO, size=9, bold=True, color=AMBAR_TXT)
    av.alignment = Alignment(vertical="center", indent=1)
    _nota(av, "Aviso del codigo, transcrito: acota para que sirve el valor que "
              "devuelve este motor.")

    # --------------------------- 5 · CURVA ---------------------------------
    rg = rt + 7
    banda(ws, rg, "5 · CURVA DEL HDS   —   valor tabulado frente a la temperatura")
    c0 = APXB1_CURVA_COL
    q = f"'{ws.title}'!"

    def _q(expr):
        return (expr.replace("$D$5", q + "$D$5")
                    .replace(NP, q + NP).replace(FIL, q + FIL))

    curvas.cell(1, c0, f"{ws.title} — datos de la curva (hoja auxiliar)").font = SRC_F
    curvas.cell(2, c0).value = f'=CONCATENATE("Temperatura, ",{q}$F$11)'
    curvas.cell(2, c0 + 1).value = (
        f'=CONCATENATE("HDS, ",IF({q}{FIL}="","",'
        f'INDEX(IF({q}$D$5="SI",{_rngb1(master,"Unidad de HDS")},'
        f'{_rngb1(us,"Unidad de HDS")}),{q}{FIL})))')
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
    ch.title = "HDS TABULADO FRENTE A LA TEMPERATURA"
    ch.scatterStyle = "line"
    ch.x_axis.title = "Temperatura  [°C en SI  ·  °F en US]"
    ch.y_axis.title = "HDS tabulado  [MPa en SI  ·  ksi en US]"
    ch.height, ch.width = 9.5, 26
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.13, y=0.15, w=0.80, h=0.65))
    xs = Reference(curvas, min_col=c0, min_row=3, max_row=2 + npack)
    ys = Reference(curvas, min_col=c0 + 1, min_row=2, max_row=2 + npack)
    s1 = Series(ys, xs, title_from_data=True)
    # Linea continua sin marcadores, en tinta (regla de estilo 3).
    s1.marker = Marker(symbol="none")
    s1.smooth = False
    s1.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=TINTA_2,
                                                                   w=19050))
    ch.series.append(s1)
    xq = Reference(curvas, min_col=c0 + 3, min_row=3, max_row=3)
    yq = Reference(curvas, min_col=c0 + 4, min_row=2, max_row=3)
    s2 = Series(yq, xq, title_from_data=True)
    s2.marker = Marker(symbol="diamond", size=10,
                       spPr=GraphicalProperties(
                           solidFill=ROJO, ln=LineProperties(solidFill=ROJO)))
    s2.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.series.append(s2)
    estilizar_chart(ch)
    ws.add_chart(ch, f"A{rg + 1}")

    for cc, w in zip("ABCDEFGHIJKL",
                     [24, 20, 16, 16, 10, 3, 24, 20, 16, 16, 10, 3]):
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
    fija.font = Font(name=MONO, size=10, bold=True, color=TINTA)
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
        c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
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
    _nota(ay, "Aviso automatico: SELECCION COMPLETA (bloque verde) cuando la cascada "
              "esta resuelta; SELECCION INCOMPLETA (bloque ambar) mientras falte un paso.")
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
        c.font, c.fill, c.border = IN_F, IN_FILL, CAJA_CAMPO
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
        av.font = Font(name=MONO, size=10, bold=True, color=TINTA)
        av.fill = PatternFill("solid", fgColor=AMBAR)
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
    nt.font = DATA_F
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
    rec.font = Font(name=MONO, size=10, bold=True, color=TINTA)
    rec.fill = PatternFill("solid", fgColor=PAPEL_2)
    rec.border = Border(top=REGLA, bottom=REGLA)
    rec.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[r].height = 18

    for cc, w in zip("ABCDEFGHIJKL",
                     [22, 20, 18, 18, 12, 6, 22, 20, 16, 14, 12, 4]):
        ws.column_dimensions[cc].width = w
    return ws


# ---------------------------------------------------------------------------
# Integracion en el motor Parche_PCC2_Art212
# ---------------------------------------------------------------------------
MOTOR = "Parche_PCC2_Art212"
# Primera fila del ANEXO DE PASOS DEL FLUJO 212 (bloques nuevos anadidos sin
# desplazar el oracle Rev0, que solo cubre 1-129; la Seccion 7 termina en 131).
# Lo que vive de esta fila en adelante es nuevo por diseno y lo fijan las anclas
# por-paso de TestBuildParcheContraOracle, no el oracle: los tests de paridad
# (validaciones y fusionados) lo excluyen de la comparacion contra el oracle.
FILA_ANEXO_FLUJO_212 = 132

# Ancho de la tabla de los dos motores: columnas A..G (ver autosize() de cada
# uno). La leyenda y su pase se mueven dentro de esa banda y no mas alla, que es
# donde viven las columnas ocultas de listas materializadas y las de clave de
# navegacion (COL_CLAVE_BASE): pintarlas seria colorear metadato invisible.
MOTOR_NCOLS = 7

# La leyenda impresa. Cada muestra lleva EL MISMO objeto de relleno que
# `aplicar_leyenda_motor` pone en las celdas, no una copia del color: asi la
# muestra no puede acabar describiendo un tono que la hoja ya no usa —el fallo
# que TEXTOS_HEREDADOS tuvo que reparar a mano cuando el maestro Rev0 seguia
# explicando sus «celdas azules sobre fondo amarillo»—.
#
# Va en D3/E3/F3 con la nota en G3 (la columna de «Referencia / Notas» de estos
# motores) y SIN FUSIONAR ninguna celda: A3:C3 ya lo ocupa el boton de volver, y
# un merge nuevo por debajo de FILA_ANEXO_FLUJO_212 romperia la comparacion de
# fusionados contra el *oracle* del 212 (que es justo la red que hay que
# conservar intacta hasta la Fase 3).
LEYENDA_MOTOR = (
    ("D3", "SE TECLEA", MOTOR_IN_FILL),
    ("E3", "FORMULA", MOTOR_CALC_FILL),
    ("F3", "ROTULO", MOTOR_LBL_FILL),
)
LEYENDA_MOTOR_NOTA = (
    "Leyenda de color de celda: amarillo = lo rellena usted (tecleado o lista "
    "desplegable) · gris = lo calcula el libro, esta bloqueada · papel = rotulo, "
    "unidad o referencia.")


def build_leyenda_motor(ws):
    """Escribe la leyenda de color en la fila 3 de un motor de calculo."""
    for celda, texto, relleno in LEYENDA_MOTOR:
        c = ws[celda]
        c.value = texto
        c.font = Font(name=MONO, size=9, bold=True, color=TINTA)
        c.fill = relleno
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        _nota(c, LEYENDA_MOTOR_NOTA)
    g = ws["G3"]
    g.value = "Leyenda de color de celda"
    g.font = SRC_F
    _nota(g, LEYENDA_MOTOR_NOTA)
    return ws


# ---------------------------------------------------------------------------
# Fase 8 — boton de reinicio de entradas
# ---------------------------------------------------------------------------
# El VBA no lleva ni una direccion de celda: cada motor publica en una columna
# oculta el MANIFIESTO de sus celdas de entrada y `LimpiarEntradas` lo lee de
# ahi. Es la misma razon por la que `HojasNavegables()` vive en un solo sitio —
# dos copias de la misma lista divergen el dia que alguien anade un campo—, y
# aqui el precio de divergir es peor que una navegacion rota: un boton que dice
# «reiniciar» y deja un campo con el valor del caso anterior.
#
# El manifiesto se DERIVA del estado real de la hoja, con el mismo criterio con
# el que `aplicar_leyenda_motor` decide pintarla de amarillo (desbloqueada, sin
# hipervinculo, dentro de A..G). Las tres cosas —lo que la leyenda promete que
# se teclea, lo que Excel deja teclear y lo que el boton limpia— quedan asi por
# construccion en el mismo conjunto: un campo nuevo entra en los tres a la vez
# o en ninguno.
# Columna 100 (CV): por encima de COL_CLAVE_BASE (66) y de la clave mas alta que
# puede grabar un boton (66 + su columna), con margen para que anadir un boton
# manana no aterrice encima del manifiesto.
COL_MANIFIESTO_RESET = 100       # DEBE coincidir con COL_MANIFIESTO del VBA.
SENTINEL_RESET = "RESET_MANIFIESTO"
PREFIJO_RESET = "RESET:"
# Texto ASCII a proposito: las dos fuentes del sistema son las que trae Windows
# (Consolas / Arial Black) y un glifo que la fuente no tenga sale como recuadro
# vacio en el entregable. Mismo formato que TXT_MANUAL ("[ ? ] MANUAL DE USO").
TXT_RESET = "[ RESET ] REINICIAR ENTRADAS"
# Fila 3 —la de acciones, la del boton VOLVER— pero a la DERECHA de la tabla
# (H..J). No entra en A..G a proposito: esa banda la compara celda a celda el
# *oracle* del 212 y la pinta `aplicar_leyenda_motor`.
FILA_RESET, COL_RESET_1, COL_RESET_2 = 3, 8, 10


def _es_entrada_motor(c):
    """True si `c` es una celda que rellena el ingeniero en un motor.

    Un boton de navegacion se entrega tambien DESBLOQUEADO (para no depender de
    que Excel deje seguir un hipervinculo en celda bloqueada), asi que pasaria
    por entrada; lo que de verdad lo distingue es el hipervinculo.
    """
    # La cola de un rango fusionado no es una celda propia: openpyxl le arrastra
    # la proteccion del ancla (por eso D5:F5 daria tres entradas donde hay un
    # solo campo) y escribir en ella por macro no es legal. Se cuenta el ancla.
    if isinstance(c, MergedCell):
        return False
    if c.hyperlink is not None:
        return False
    return c.protection is not None and c.protection.locked is False


def _fin_tabla_motor(ws):
    """Ultima fila con contenido en A..G (no `ws.max_row`).

    Las listas de cascada materializadas viven en columnas ocultas a la derecha
    y bajan cientos de filas mas que la tabla del motor.
    """
    return max((c.row for fila in ws.iter_rows(max_col=MOTOR_NCOLS) for c in fila
                if c.value is not None), default=1)


def celdas_de_entrada(ws, hasta_fila=None):
    """Direcciones de las celdas de entrada de un motor, en orden de lectura."""
    fin = hasta_fila or _fin_tabla_motor(ws)
    return [c.coordinate
            for fila in ws.iter_rows(min_row=1, max_row=fin, max_col=MOTOR_NCOLS)
            for c in fila if _es_entrada_motor(c)]


def build_reinicio_motor(ws):
    """Manifiesto de entradas + boton de reinicio de un motor de calculo.

    Se corre como ULTIMO paso, con las filas ya en su sitio definitivo (despues
    de `remapear_filas`): el manifiesto guarda direcciones, y escritas antes del
    remapeo apuntarian a las filas de antes de la mudanza.
    """
    entradas = celdas_de_entrada(ws)
    if not entradas:
        raise SystemExit(
            f"{ws.title}: el manifiesto de reinicio sale vacio. O la hoja no "
            f"tiene ni un campo editable, o inp() dejo de desbloquear las "
            f"celdas: en los dos casos el boton de reinicio mentiria.")

    col = COL_MANIFIESTO_RESET
    # El centinela es lo que el VBA comprueba antes de borrar nada: sin el, un
    # nombre de hoja equivocado en la clave del boton haria ClearContents sobre
    # una hoja que no publica manifiesto.
    cs = ws.cell(1, col, SENTINEL_RESET)
    cs.font = DATA_F
    for k, direccion in enumerate(entradas, start=2):
        ws.cell(k, col, direccion).font = DATA_F
    ws.column_dimensions[get_column_letter(col)].hidden = True

    b = _boton(ws, FILA_RESET, COL_RESET_1, COL_RESET_2, TXT_RESET,
               PREFIJO_RESET + ws.title)
    b.protection = Protection(locked=False)
    _nota(b, "Boton: vacia TODAS las celdas de entrada de esta hoja (las "
             "amarillas de la leyenda), con confirmacion previa. No toca "
             "formato, validacion ni formula: solo el contenido. Deja el motor "
             "en blanco para un caso nuevo.")
    _celda_clave(ws, FILA_RESET, COL_RESET_1).protection = Protection(locked=False)
    _ocultar_columnas_clave(ws, (COL_RESET_1,))
    return entradas


def build_botones_documentos(ws, espec=None, instr=None):
    """Atajos del motor a sus documentos (Fases 9 y 10).

    Van apilados en H1:J1 y H2:J2, encima del boton de reinicio (H3:J3): las tres
    filas estan congeladas, asi que la esquina superior derecha queda siempre a la
    vista. El retorno del arbol (A3:C3) sube al nodo del articulo; estos botones
    son el atajo lateral, para no tener que subir un nivel solo para cambiar de
    pestana del mismo articulo.
    """
    for fila, texto, destino in ((FILA_BTN_ESPEC, TXT_BTN_ESPEC, espec),
                                 (FILA_BTN_INSTR, TXT_BTN_INSTR, instr)):
        if destino is None:
            continue
        b = _boton(ws, fila, COL_BTN_DOC_1, COL_BTN_DOC_2, texto, destino)
        b.protection = Protection(locked=False)
        _celda_clave(ws, fila, COL_BTN_DOC_1).protection = Protection(locked=False)
    _ocultar_columnas_clave(ws, (COL_BTN_DOC_1,))
    return ws


# ---------------------------------------------------------------------------
# Fase 3 — la Seccion de Material sube justo detras de la Seccion 1
# ---------------------------------------------------------------------------
# El ingeniero pide leer la hoja en el orden en que se rellena: primero los
# datos de entrada, enseguida el material (que es otra entrada, la mas larga),
# y solo despues lo que el libro calcula. Hasta ahora el bloque de material
# vivia al final, detras de las especificaciones tecnicas, porque se anadio
# cuando las secciones 1-6 ya estaban ancladas al *oracle* Rev0 y moverlas
# costaba reescribir todas sus direcciones.
#
# NO se reescriben las ~1 400 lineas de direcciones literales del builder. Se
# construye la hoja donde siempre y despues se aplica UN mapa de filas. Dos
# razones, y la segunda es la que decide:
#   1. El diff queda en un solo sitio auditable —el mapa— en vez de repartido
#      por setecientas llamadas a calc().
#   2. Reescribir el FUENTE con una expresion regular es inseguro aqui: el
#      texto del builder esta lleno de cosas con forma de referencia que no lo
#      son ("A106 Gr.B", "A516 Gr.70", "D2737", "F714", "B31.3"). Sobre la hoja
#      ya construida solo hay que tocar cadenas que empiezan por "=", y dentro
#      de ellas solo lo que cae FUERA de las comillas.
#
# El ANEXO DE PASOS DEL FLUJO no se mueve, y es deliberado: deja intactas las
# direcciones que recalcula verificar.py §6e (D144, D150, D167, D183..D191) y
# las anclas por-paso de test_dashboard. El hueco que deja el bloque de
# material al subir absorbe el desplazamiento de todo lo que hay en medio.


def _mapa_filas_212():
    """fila_vieja -> fila_nueva del Art. 212. Ver el comentario de arriba.

    Fase 7: la banda "Aplicacion y codigo de construccion" gana una fila -el
    selector de sistema de unidades- y por eso todo lo que va DEBAJO de ella
    baja uno mas que en la Fase 3. El selector tiene que quedar por ENCIMA de
    los datos de entrada: gobierna las unidades de los campos que se teclean
    (kg/cm², mm, °C) y un selector debajo de los campos que rotula seria un
    selector que se descubre tarde.

    El anexo sigue sin moverse: entre el aviso (fila 130) y el PASO 1 (134)
    quedan tres filas en blanco de holgura, que es de donde sale el hueco.
    """
    m = {}
    for r in range(1, 16):      # titulo, identificacion y banda de aplicacion,
        m[r] = r                # que ahora llega hasta la fila 15 (el selector)
    for r in range(16, 36):     # Seccion 1: baja una fila
        m[r] = r + 1            # 16..35 -> 17..36
    for r in range(105, 132):   # RESOLUCION DE MATERIAL (27 filas) sube
        m[r] = r - 67           # 105..131 -> 38..64
    for r in range(37, 102):    # parametros, geometria, cargas, resultados,
        m[r] = r + 29           # verificaciones, especificaciones y aviso
    # La fila 36 (separadora) y las 102-104 (las tres en blanco que separaban el
    # aviso del bloque de material) se absorben: el hueco sigue existiendo, en
    # la 37 y en las 131-133.
    for r in range(132, 401):   # ANEXO DE PASOS DEL FLUJO: NO se mueve
        m[r] = r
    return m


def _mapa_filas_206():
    """fila_vieja -> fila_nueva del Art. 206. Mismo criterio que el 212."""
    m = {}
    for r in range(1, 35):      # ... Seccion 1 (termina en 33) + separador
        m[r] = r
    for r in range(75, 99):     # RESOLUCION DE MATERIAL (24 filas) sube
        m[r] = r - 40           # 75..98 -> 35..58
    for r in range(35, 75):     # parametros, geometria, espesor requerido,
        m[r] = r + 25           # verificaciones y dictamen global
    # 99-100 en blanco: se absorben.
    for r in range(101, 401):   # ANEXO: NO se mueve
        m[r] = r
    return m


MAPA_FILAS_212 = _mapa_filas_212()
MAPA_FILAS_206 = _mapa_filas_206()

# Celda del selector de sistema de unidades de cada motor (Fase 7). Vive en
# la banda de aplicacion y codigo, por encima de los datos de entrada.
UNIDAD_212, UNIDAD_206 = "$D$15", "$D$14"
# Y la CONDICION, aparte de la celda. Confundirlas cuesta caro y en silencio:
# `IF($D$15, a, b)` es sintaxis valida de Excel —Excel intenta leer "SI" como
# booleano— y devuelve #VALUE!, que se propaga hacia abajo sin decir de donde
# vino. Por eso hay dos nombres y no uno.
ES_SI_212, ES_SI_206 = f'{UNIDAD_212}="SI"', f'{UNIDAD_206}="SI"'


# Una referencia de celda de ESTAS hojas: columna A..G y fila de hasta tres
# digitos. Se exige que no venga precedida de letra, digito, guion bajo ni "!"
# —eso descarta el nombre de otra hoja (DB_B31_3!$A$4) y los sufijos de un
# identificador— y que no le siga un digito ni un parentesis de apertura (eso
# descarta un numero mas largo y un nombre de funcion).
_REF_MOTOR = re.compile(r'(?<![A-Za-z0-9_!$])(\$?)([A-G])(\$?)(\d{1,3})(?![0-9(])')

# Un operando de OTRA hoja, entero: "DB_B31_3!$A$4" o "DB_B36_19!$A$4:$A$50".
# Hay que reconocerlo COMPLETO y no solo por el "!": el segundo extremo de un
# rango no lo lleva delante —en "DB_B36_19!$A$4:$A$50" lo que precede a "$A$50"
# es un ":"— y sin esto _REF_MOTOR lo tomaba por una celda de la hoja del motor
# y le movia la fila. Comprobado: convertia ese rango en "$A$4:$A$79" y
# "MAP_Factores!$A$4:$A$123" en "$A$56". Es corrupcion silenciosa: la formula
# sigue siendo valida y devuelve otro numero.
_OPERANDO_EXTERNO = re.compile(
    r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d+"
    r"(?::\$?[A-Z]{1,3}\$?\d+)?")


def remapear_referencias(formula, mapa):
    """Reescribe las FILAS de las referencias A..G de una formula.

    Solo actua sobre lo que cae FUERA de un literal entre comillas dobles: ahi
    dentro vive el texto de los dictamenes ("CUMPLE", "Migrar (Art.206)", la
    prosa de las especificaciones tecnicas), que no son referencias por mucho
    que alguna se le parezca. Lo que no es una formula se devuelve tal cual.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    def sub(trozo):
        # Los operandos de otra hoja se apartan ENTEROS antes de tocar nada y
        # se reponen despues: asi ni el segundo extremo de uno de sus rangos
        # puede confundirse con una celda de esta hoja.
        externos = []

        def guardar(m):
            externos.append(m.group(0))
            return f"\x00{len(externos) - 1}\x00"

        limpio = _OPERANDO_EXTERNO.sub(guardar, trozo)
        limpio = _REF_MOTOR.sub(
            lambda m: f"{m[1]}{m[2]}{m[3]}{mapa.get(int(m[4]), int(m[4]))}", limpio)
        return re.sub(r"\x00(\d+)\x00",
                      lambda m: externos[int(m[1])], limpio)

    partes, i, dentro = [], 0, False
    for j, ch in enumerate(formula):
        if ch == '"':
            partes.append(formula[i:j] if dentro else sub(formula[i:j]))
            partes.append('"')
            dentro = not dentro
            i = j + 1
    partes.append(formula[i:] if dentro else sub(formula[i:]))
    return "".join(partes)


def refs_propias(formula):
    """Las referencias de ESTA hoja que trae una formula, en orden.

    Comparte con `remapear_referencias` las dos exclusiones que importan -lo
    que va entre comillas y los operandos de otra hoja- porque las dos preguntas
    son la misma: que trozos de la formula son celdas de este motor. Tenerlo en
    una sola funcion evita que una auditoria mire un conjunto de referencias y
    el remapeo mueva otro.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return []
    encontradas, i, dentro = [], 0, False
    for j, ch in enumerate(formula):
        if ch == '"':
            if not dentro:
                encontradas += _refs_de_trozo(formula[i:j])
            dentro = not dentro
            i = j + 1
    if not dentro:
        encontradas += _refs_de_trozo(formula[i:])
    return encontradas


def _refs_de_trozo(trozo):
    limpio = _OPERANDO_EXTERNO.sub(" ", trozo)
    return [(m[2], int(m[4])) for m in _REF_MOTOR.finditer(limpio)]


def remapear_filas(ws, mapa, ncols=None):
    """Mueve las filas de A..ncols de `ws` segun `mapa` y repunta las formulas.

    Se corre DESPUES de construir la hoja entera (comentarios y subindices
    incluidos) y antes de la leyenda de color. Mueve valor, estilo, comentario,
    hipervinculo, rangos fusionados, validaciones de datos, formato condicional
    y alto de fila.

    Lo que NO se toca, a proposito:
      - las columnas a la derecha de `ncols`: ahi viven las listas de cascada
        materializadas y las celdas de clave de navegacion. Sus filas no son
        las de la tabla, y las formulas que las leen usan columnas > G, asi que
        `remapear_referencias` tampoco las reescribe.
      - el `formula1` de una validacion (apunta a esas mismas listas ocultas);
        solo se mueve su `sqref`, que si son celdas de la tabla.
    """
    ncols = ncols or MOTOR_NCOLS
    # 1. Fotografia de lo que hay, antes de tocar nada. Se guarda el _style
    # entero (StyleArray: fuente, relleno, borde, alineacion, formato de numero
    # y proteccion de una vez) en vez de atributo por atributo.
    foto = []
    for fila in ws.iter_rows(max_col=ncols):
        for c in fila:
            if isinstance(c, MergedCell):
                continue
            if c.value is None and not c.has_style and c.comment is None:
                continue
            foto.append((c.row, c.column, c.value, copy(c._style),
                         c.comment, c.hyperlink))
    fusiones = [str(r) for r in ws.merged_cells.ranges]
    validaciones = [(dv, str(dv.sqref)) for dv in ws.data_validations.dataValidation]
    # Alto, nivel de outline y plegado viajan JUNTOS con la fila. El outline es
    # lo que pliega el rastro de la resolucion de material (Fase 4): si se
    # quedara en las filas viejas, el "+" del margen plegaria un bloque que ya
    # no esta ahi y el rastro quedaria abierto donde tenia que ir cerrado.
    dims = {r: (d.height, d.outline_level, d.hidden)
            for r, d in ws.row_dimensions.items()
            if d.height or d.outline_level or d.hidden}

    # 2. Se vacia la banda. Hay que deshacer las fusiones primero: una MergedCell
    # no admite escritura y sobrevivir a la mudanza con el rango antiguo dejaria
    # la hoja fusionando celdas que ya no son las suyas.
    for rango in fusiones:
        ws.unmerge_cells(rango)
    for fila in ws.iter_rows(max_col=ncols):
        for c in fila:
            c.value = None
            c.comment = None
            c.hyperlink = None
            c.style = "Normal"

    # 3. Se reescribe en el destino, con las formulas ya repuntadas.
    for r, col, valor, estilo, comentario, enlace in foto:
        nc = ws.cell(mapa.get(r, r), col)
        nc.value = remapear_referencias(valor, mapa)
        nc._style = estilo
        if comentario is not None:
            # Un Comment no se puede reasignar a otra celda (openpyxl lo ancla
            # a la suya al asignarlo): se clona.
            nc.comment = Comment(comentario.text, comentario.author)
        if enlace is not None:
            nc.hyperlink = enlace

    for rango in fusiones:
        ws.merge_cells(_remapear_rango(rango, mapa))
    for dv, sqref in validaciones:
        dv.sqref = MultiCellRange(
            " ".join(_remapear_rango(str(r), mapa) for r in
                     MultiCellRange(sqref).ranges))
    # Se limpian TODAS las filas de origen antes de escribir los destinos: dos
    # filas distintas pueden aterrizar donde antes habia otra cosa, y limpiar
    # sobre la marcha borraria lo que se acaba de poner.
    for r in dims:
        d = ws.row_dimensions[r]
        d.height, d.outline_level, d.hidden = None, 0, False
    for r, (alto, nivel, oculta) in dims.items():
        d = ws.row_dimensions[mapa.get(r, r)]
        d.height, d.outline_level, d.hidden = alto, nivel, oculta

    # --- Las columnas ocultas NO se mueven, pero SI apuntan a las que si -----
    # Ahi viven las listas de cascada materializadas y las auxiliares de
    # resolucion. Sus filas no son las de la tabla y no deben moverse; pero sus
    # formulas leen celdas de A..G —la clave de la cascada es
    # `=$D$109&"|"&$D$110&...`— y esas si se han movido. Sin este segundo pase
    # quedan apuntando a filas vacias y la CASCADA DEJA DE RESOLVER, en
    # silencio: el caso semilla no lo delata porque resuelve su material por la
    # celda 'Variante', que es una via de escape de la propia cascada.
    #
    # Solo se reescriben las REFERENCIAS: ninguna celda de estas columnas
    # cambia de sitio. Las referencias entre columnas ocultas ($W$109) y a otras
    # hojas (DB_B31_3!$A$4) no se tocan — la primera porque su columna esta
    # fuera de A..G y la segunda por el guardia del "!".
    for fila in ws.iter_rows(min_col=ncols + 1):
        for c in fila:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = remapear_referencias(c.value, mapa)
    return ws


def _remapear_rango(rango, mapa):
    """'A84:C84' -> 'A112:C112'. Un rango de una sola celda tambien vale."""
    rng = CellRange(rango)
    return str(CellRange(min_col=rng.min_col, max_col=rng.max_col,
                         min_row=mapa.get(rng.min_row, rng.min_row),
                         max_row=mapa.get(rng.max_row, rng.max_row)))


# ---------------------------------------------------------------------------
# Fase 5 — donde va un comentario y donde no
# ---------------------------------------------------------------------------
# Hasta ahora casi toda fila llevaba el MISMO comentario en el rotulo (col. A) y
# en la celda de valor. Duplicar no informa: el ingeniero pasa el raton por la
# celda que esta mirando, y un globo repetido solo tapa la hoja. La regla que
# pidio el ingeniero es por seccion:
#
#   - por omision, el comentario va SOLO en la columna de valor;
#   - en las secciones de CALCULO (cargas/espesor) y de RESULTADOS va tambien
#     en el rotulo, porque ahi el rotulo es un simbolo (F_m, w_min, S_w) y es
#     justo lo que hay que poder consultar;
#   - en VERIFICACIONES va ademas en la columna de Resultado, que es la que se
#     lee para decidir.
#
# No se edita call site por call site (serian ~200): se declara el rango de
# cada seccion y un pase final reparte lo que YA existe. No se inventa texto
# ninguno — cuando una columna obligada no trae comentario propio se le pone el
# de la fila, tomado de la columna que esa seccion declare como fuente.
#
# `cols` es el conjunto de columnas que DEBEN llevarlo (y ninguna otra lo lleva);
# `fuente` es de donde se copia si falta. En VERIFICACIONES la fuente es la
# columna de Resultado, no la de Requerido: el texto que describe la fila entera
# es "CUMPLE si ...", no "valor minimo exigido".
REGLAS_COMENTARIO_212 = (
    (5, 6, "D", "D"),          # Identificacion
    (10, 14, "D", "D"),        # Aplicacion y codigo de construccion
    (19, 36, "D", "D"),        # 1. Datos de entrada
    (40, 62, "DE", "D"),       # 2. Resolucion de material
    (64, 64, "A", "A"),        # nota fija de la cascada
    (68, 77, "D", "D"),        # 3. Parametros de calculo
    (81, 86, "D", "D"),        # 4. Geometria y propiedades derivadas
    (90, 98, "ADE", "D"),      # 5. Calculo de cargas y soldadura
    (102, 109, "AD", "D"),     # 6. Resultados del diseno
    (113, 118, "ADEF", "F"),   # 7. Verificaciones
    (119, 119, "AF", "F"),     # Dictamen global
    (122, 128, "B", "B"),      # 8. Especificaciones tecnicas
    (130, 130, "A", "A"),      # aviso de responsabilidad
)
REGLAS_COMENTARIO_206 = (
    (5, 6, "D", "D"),
    (10, 13, "D", "D"),
    (18, 33, "D", "D"),
    (37, 56, "D", "D"),        # 2. Resolucion de material (una sola columna)
    (58, 58, "A", "A"),
    (62, 65, "D", "D"),        # 3. Parametros de calculo
    (70, 71, "D", "D"),        # 4. Geometria del sleeve
    (76, 80, "ADE", "D"),      # 5. Calculo de espesor requerido
    (85, 91, "ADEF", "F"),     # 6. Verificaciones y avisos
    (94, 94, "AF", "F"),       # Dictamen global
)


def aplicar_reglas_de_comentario(ws, reglas):
    """Reparte los comentarios de cada seccion segun `reglas`. Ver arriba.

    Solo mueve y borra lo que ya existe. Nunca pone un comentario en una celda
    VACIA: asi una seccion de una sola columna de valor (el Art. 206) usa la
    misma regla que una de dos (el 212) sin fabricar globos sobre la nada.
    """
    for ini, fin, cols, fuente in reglas:
        obligadas = {ord(c) - 64 for c in cols}
        col_fuente = ord(fuente) - 64
        for fila in range(ini, fin + 1):
            propios = {c: ws.cell(fila, c).comment
                       for c in range(1, MOTOR_NCOLS + 1)
                       if ws.cell(fila, c).comment is not None}
            if not propios:
                continue
            # Texto de la fila: el de la columna que la seccion declara como
            # fuente; si esa no lo trae, el primero que haya (orden de columna).
            base = propios.get(col_fuente) or propios[min(propios)]
            for col in range(1, MOTOR_NCOLS + 1):
                celda = ws.cell(fila, col)
                # La cola de un rango fusionado no admite comentario propio
                # (openpyxl la deja de solo lectura) y ademas Excel muestra el
                # de la celda ancla en todo el rango: no hay nada que hacerle.
                if isinstance(celda, MergedCell):
                    continue
                if col not in obligadas or celda.value is None:
                    celda.comment = None
                elif celda.comment is None:
                    celda.comment = Comment(base.text, base.author)
    return ws


# ---------------------------------------------------------------------------
# Fase 6 — semaforo de aceptacion y dictamen global destacado
# ---------------------------------------------------------------------------
# Los textos que cuentan como ACEPTACION en cada fila de resultado. Cuatro de
# las seis verificaciones del 212 resuelven en CUMPLE/NO CUMPLE; las otras dos
# no son un pasa/no pasa sino una RUTA de reparacion, y ahi lo verde es la rama
# que deja seguir con el parche. Se declaran fila a fila, leidos de la propia
# formula del motor, para no suponer que toda celda de resultado dice "CUMPLE".
SEMAFORO_212 = {113: ("CUMPLE",), 114: ("CUMPLE",), 115: ("CUMPLE",),
                116: ("CUMPLE",), 117: ("Parche local",), 118: ("OK — parche",),
                161: ("CUMPLE",), 162: ("CUMPLE",)}
SEMAFORO_206 = {85: ("CUMPLE",), 86: ("CUMPLE",), 123: ("CUMPLE",)}
FILA_DICTAMEN_212, FILA_DICTAMEN_206 = 119, 94


def aplicar_semaforo_motor(ws, semaforo, fila_dictamen):
    """Semaforo en las celdas de Resultado y bloque propio para el dictamen."""
    for fila, favorables in semaforo.items():
        semaforo_resultado(ws, f"F{fila}:F{fila}", f"$F${fila}", favorables)

    # --- Dictamen global: bloque propio, no una fila mas de la tabla --------
    # Es la frase que se lee primero y la que se firma. Va en macrotipografia a
    # todo el ancho, separada de la tabla de arriba por la misma franja roja que
    # cierra las bandas de seccion —el limite duro entre dos zonas, dibujado y
    # no insinuado— y con la fila mas alta para que respire.
    for col in range(1, MOTOR_NCOLS + 1):
        c = ws.cell(fila_dictamen, col)
        c.fill = BAND_FILL
        c.font = DICTAMEN_BAD_FONT if col > 1 else Font(
            name=MACRO, size=16, color=PAPEL)
    franja(ws, fila_dictamen, 1, MOTOR_NCOLS, arriba=True)
    ws.row_dimensions[fila_dictamen].height = 26

    rango = f"A{fila_dictamen}:G{fila_dictamen}"
    ref = f"$F${fila_dictamen}"
    # APTO -> verde. ELIJA MATERIAL -> ambar: no es un fallo, es que todavia
    # falta una entrada, y pintarlo de rojo confundiria "aun no has elegido"
    # con "no cumple". Todo lo demas -REVISAR, PROHIBIDO, NO ELEGIBLE, FUERA DE
    # ALCANCE- es rojo, otra vez por complemento: lo imprevisto sale bloqueado.
    ws.conditional_formatting.add(rango, FormulaRule(
        formula=[f'{ref}="APTO"'], fill=CUMPLE_OK_FILL, font=DICTAMEN_OK_FONT))
    ws.conditional_formatting.add(rango, FormulaRule(
        formula=[f'LEFT({ref},14)="ELIJA MATERIAL"'],
        fill=DICTAMEN_ESPERA_FILL, font=DICTAMEN_ESPERA_FONT))
    ws.conditional_formatting.add(rango, FormulaRule(
        formula=[f'AND({ref}<>"",{ref}<>"APTO",LEFT({ref},14)<>"ELIJA MATERIAL")'],
        fill=CUMPLE_BAD_FILL, font=DICTAMEN_BAD_FONT))
    return ws


# ---------------------------------------------------------------------------
# Fase 7 — el motor entero cambia de sistema de unidades
# ---------------------------------------------------------------------------
# El conmutador no puede gobernar solo el bloque de material: la edicion US
# publica el esfuerzo admisible en ksi, y meter un ksi en una cadena que opera
# en MPa/mm da un numero sencillamente equivocado. O cambia de unidades TODO el
# motor, o el selector no puede tocar el calculo.
#
# La cadena entera es coherente en cualquiera de los dos sistemas, porque las
# ecuaciones del codigo son dimensionales y no llevan constantes de unidad:
#   t_req = P·D / (2·(S·E + P·Y))     ksi·in / ksi  -> in
#   F     = P·Dm/2                    ksi·in        -> kip/in
#   w_min = F/(E·Sa)                  (kip/in)/ksi  -> in
#   S_w   = P·Dm/(2T) + 3·P·Dm·e/T²   ksi           -> ksi
# Lo unico que hay que cambiar es (i) el rotulo de cada magnitud, (ii) las tres
# constantes que SI dependen del sistema y (iii) de que columna de B36 salen el
# OD y el espesor. Los valores que teclea el ingeniero NO se convierten: hay que
# volver a teclearlos, igual que en los cinco buscadores de cascada (regla 9).
#
# Los rotulos se declaran en una tabla, fila a fila, y los escribe un pase
# final. Editar cuarenta llamadas a lab() habria repartido por toda la funcion
# una decision que se lee de un golpe aqui.
_U = {"len": ("mm", "in"), "temp": ("°C", "°F"), "pres": ("kg/cm²", "psi"),
      "esf": ("MPa", "ksi"), "fuerza": ("N/mm", "kip/in"),
      "dens": ("kg/m³", "lb/in³"), "peso": ("kg", "lb"),
      "vol": ("m³", "ft³"), "pabs": ("MPa abs", "psia"),
      "energia": ("J", "ft·lb"), "masa": ("kg", "lb"),
      "dist": ("m", "ft"), "rscaled": ("m/kg^⅓", "ft/lb^⅓")}

UNIDADES_212 = {
    21: "len", 22: "len", 26: "temp", 27: "pres", 29: "pres", 30: "len",
    31: "len", 32: "len", 33: "len", 34: "len", 35: "len", 36: "len",
    68: "esf", 69: "esf", 70: "esf", 74: "dens", 77: "esf",
    81: "len", 82: "len", 83: "len", 84: "len", 85: "len",
    90: "pres", 91: "esf", 92: "fuerza", 93: "len", 94: "len",
    95: "esf", 96: "esf", 97: "esf",
    102: "len", 103: "esf", 104: "pres", 107: "len", 108: "len", 109: "peso",
    # Anexo de pasos del flujo (no se movio, pero si cambia de unidades).
    144: "fuerza", 145: "fuerza", 146: "fuerza", 147: "fuerza", 148: "fuerza",
    149: "fuerza", 150: "fuerza", 155: "len", 167: "len", 172: "len",
    184: "vol", 185: "pabs", 186: "pabs", 188: "rscaled",
    189: "energia", 190: "masa", 191: "dist",
}
UNIDADES_206 = {
    20: "len", 21: "len", 25: "temp", 26: "pres", 28: "pres", 29: "len",
    30: "len", 31: "len", 32: "len", 62: "esf", 70: "len", 71: "len",
    76: "pres", 77: "esf", 78: "len", 79: "len", 80: "len", 86: "len",
    113: "len", 121: "len", 140: "pres", 141: "pres",
}


def aplicar_unidades_motor(ws, tabla, sel):
    """Escribe el rotulo de unidad de cada fila como formula del selector."""
    for fila, clase in tabla.items():
        si, us = _U[clase]
        c = ws.cell(fila, 3)
        if c.value is None:        # fila que no existe en este motor: se avisa
            ISSUES.append(f"{ws.title}: la fila {fila} de la tabla de unidades "
                          f"no tiene rotulo de unidad; revise UNIDADES_*.")
            continue
        c.value = f'=IF({sel},"{si}","{us}")'
    return ws


def aplicar_leyenda_motor(ws, hasta_fila=None):
    """Pinta la leyenda de edicion sobre A..G de un motor de calculo.

    NO se etiqueta celda a celda en los ~200 sitios que escriben la hoja: el
    color se DERIVA del estado que la celda ya tiene, que es la unica forma de
    que la leyenda impresa no pueda mentir.

      celda desbloqueada  -> AMARILLO   (la escribe el ingeniero: `inp()` es lo
                            unico del libro que pone Protection(locked=False))
      formula             -> GRIS_2     (la calcula el libro)
      lo demas sin relleno propio -> PAPEL

    Se corre como ULTIMO paso de cada motor, despues de comentarios y
    subindices, y solo toca el relleno: valor, fuente, borde, validacion y
    comentario quedan intactos —por eso el *oracle* del 212, que compara
    valores, no se ve afectado—. Una celda que ya trajo relleno propio (banda,
    cabecera, boton, KPI, semaforo) se respeta: ya dice otra cosa.
    """
    # Hasta la ultima fila con contenido EN A..G, no hasta ws.max_row: las
    # listas de cascada materializadas viven en columnas ocultas a la derecha y
    # bajan cientos de filas mas que la tabla del motor. Pintar hasta ahi no se
    # veria (el sustrato de columna ya es papel) y solo multiplicaria las celdas
    # con estilo del .xlsm.
    fin = hasta_fila or _fin_tabla_motor(ws)
    for fila in ws.iter_rows(min_row=1, max_row=fin, max_col=MOTOR_NCOLS):
        for c in fila:
            # `_es_entrada_motor` es la MISMA funcion con la que el manifiesto de
            # reinicio (Fase 8) decide que celda limpiar, y excluye los botones
            # de navegacion por el hipervinculo. Compartirla es lo que impide que
            # la leyenda prometa amarillo donde el boton no borra, o al reves.
            if c.hyperlink is not None:
                continue
            if _es_entrada_motor(c):
                c.fill = MOTOR_IN_FILL
            elif isinstance(c.value, str) and c.value.startswith("="):
                c.fill = MOTOR_CALC_FILL
            elif c.fill is None or not c.fill.patternType:
                c.fill = MOTOR_LBL_FILL
            # Una celda con texto y SIN fuente propia se veia bien —heredaba la
            # mono del sustrato de columna— pero se colaba por la auditoria de
            # fuentes, que solo mira las celdas con estilo. Al darles relleno
            # aqui dejan de ser invisibles, asi que hay que fijarles tambien la
            # fuente del sistema o el libro pasaria a declarar Calibri.
            if c.value is not None and c.font.name not in (MONO, MACRO):
                c.font = DATA_F
    return ws


def construir_seccion7_material(
    ws, fila_base, b313, iid1a, iidb, fac_info, rangos,
    columnas=(("D", "Metal base"), ("E", "Collar / parche")),
    modo_cell="$D$11", temp_fuente_cell="$D$25",
    incluir_ej_ec=True, destino_st="la seccion de parametros de calculo",
    nota_extra_cascada="",
    titulo_banda="2. RESOLUCION DE MATERIAL — BASE DE DATOS ASME",
    banda_rotulo=True, mapa_citas=None,
    b313c=None, iid1ac=None, iidbc=None, unidad_cell=None,
):
    """Bloque de resolucion de material: cascada de 5 niveles + variante, con
    las listas materializadas en columnas ocultas (unica forma portable a
    Google Sheets), consulta sobre la banda compacta (sin formulas
    matriciales) e interpolacion de S(T) — y, si `incluir_ej_ec`, el lookup
    de Ej/Ec contra MAP_Factores. Construye el bloque a partir de `fila_base`
    para 1 o 2 columnas de material en paralelo (`columnas`), y ambos motores
    de calculo del libro (Art. 212 y Art. 206) lo comparten para no duplicar
    la cascada ni la interpolacion. `modo_cell`/`temp_fuente_cell` son las
    celdas —de la MISMA hoja— del selector de codigo y de la temperatura de
    evaluacion de cada motor; `destino_st` es solo texto informativo (a donde
    va el S(T) resuelto) para la nota de la celda, no altera ningun calculo.

    Devuelve las referencias absolutas que el llamador cablea en sus propias
    secciones: `por_columna[letra]` con `material_id`/`dictamen`/`s_t`/`tmax`,
    y, si `incluir_ej_ec`, `ej_ref`/`ec_ref`.
    """
    rb, ri = rangos["B313"], rangos["IID1A"]
    # Fase 7 — el eje SI <-> US. `bases` son las SEIS hojas de esfuerzo
    # admisible: las tres metricas y sus tres gemelas en U.S. Customary. El
    # indice de CHOOSE lleva +3 cuando el selector dice US, asi que "la misma
    # base en la otra edicion" es sumar tres y nada mas.
    #
    # Se hace con UN CHOOSE de seis ramas y no con un IF envolviendo cada
    # CHOOSE: un IF(cond, rangoA, rangoB) como argumento de MATCH exige entrada
    # matricial (CSE), que la regla 1 de diseno del libro prohibe. CHOOSE con
    # indice escalar entrega una REFERENCIA, que INDEX/MATCH consumen tal cual.
    us_ok = all(x is not None for x in (b313c, iid1ac, iidbc, unidad_cell))
    bases = [b313, iid1a, iidb] + ([b313c, iid1ac, iidbc] if us_ok else [])
    SEL = f'{unidad_cell}="SI"' if us_ok else "TRUE"
    IDC, TMAXC = CL["material_id"], CL["Temp. max. / limite"]
    BIC = CL["clave_bi"]

    def coln(cl_, cuales):
        """Columna `cl_` de las bases `cuales`, en un CHOOSE de N ramas."""
        return ('CHOOSE({i},' + ",".join(
            f'{b["sheet"]}!${cl_}${R_DATA}:${cl_}${b["last_row"]}'
            for b in cuales) + ')')

    IDS, TMAX = coln(IDC, bases), coln(TMAXC, bases)
    # La clave bilingue es lo unico que enlaza una fila metrica con su gemela
    # US: el material_id NO coincide entre ediciones (el tag es "A-1" frente a
    # "A-1C" y el tamano va en mm frente a in, y los dos entran en la clave).
    # Por eso la fila US no se busca por material_id sino en tres saltos:
    # fila SI -> clave bilingue -> fila US, igual que en los buscadores.
    BI_SI = coln(BIC, bases[:3])
    BI_US = coln(BIC, bases[3:]) if us_ok else None
    NAMES = 'CHOOSE({i},' + ",".join(f'"{b["sheet"]}"' for b in bases) + ')'
    pk = [packed_refs(b) for b in bases]
    NPTS = 'CHOOSE({i},' + ",".join(p["npts"] for p in pk) + ')'
    TANC = 'CHOOSE({i},' + ",".join(p["t_anchor"] for p in pk) + ')'
    VANC = 'CHOOSE({i},' + ",".join(p["v_anchor"] for p in pk) + ')'
    hl = get_column_letter
    M = modo_cell

    def lab(r, text, note=None, com=None):
        cl = ws.cell(r, 1, text)
        cl.font = Font(name=MONO, size=10, color=TINTA)
        if note:
            ws.cell(r, 7, note).font = Font(name=MONO, size=9, color=GRIS)
        if com:
            _nota(cl, com)

    def inp(cell, value="", com=None):
        c = ws[cell]
        c.value = value
        # Relleno AMARILLO: este bloque solo lo construyen los dos motores de
        # calculo, y en ellos la celda editable se distingue por el relleno
        # (ver el token AMARILLO y la leyenda de cada motor), no por el borde.
        c.font, c.fill, c.border = IN_F, MOTOR_IN_FILL, CAJA_CAMPO
        c.protection = Protection(locked=False)
        if com:
            _nota(c, com)

    F = fila_base

    def cita(r):
        """Numero de fila tal como lo vera el ingeniero en la hoja acabada.

        El bloque se escribe en su sitio historico y `remapear_filas` lo mueve
        despues (Fase 3), asi que una cita calculada sobre `F` apuntaria a la
        fila de ANTES de la mudanza — es decir, a una fila equivocada."""
        return (mapa_citas or {}).get(r, r)
    letras = [c for c, _ in columnas]
    if len(columnas) == 2:
        (c1, l1), (c2, l2) = columnas
        desc_cols = f"{l1} ({c1}) y {l2} ({c2})"
    else:
        (c1, l1), = columnas
        desc_cols = f"{l1} ({c1})"

    # Fase 3: el rotulo lo pone el LLAMADOR. El bloque dejo de ser la "Seccion
    # 7" del final y pasa a ser la 2, justo detras de los datos de entrada; y
    # cada motor escribe sus bandas con un estilo distinto —el 212 literal, en
    # mayus/minus mixtas, porque el *oracle* Rev0 las trae asi; el 206 con
    # rotulo(), en "[ MAYUSCULAS // ... ]"—, asi que una banda fija aqui iba a
    # desentonar en uno de los dos. `banda_rotulo` decide cual se aplica.
    ws.cell(F, 1, rotulo(titulo_banda) if banda_rotulo else titulo_banda)
    ws.cell(F, 1).font = Font(name=MACRO, size=11, color=PAPEL)
    # La banda de seccion 7 va a lo ancho de la tabla del motor (7 columnas) y
    # cierra con la franja roja, igual que las bandas de los buscadores. Aqui la
    # fila NO esta fusionada, asi que hay que recorrerla de verdad.
    for j2 in range(1, 8):
        ws.cell(F, j2).fill = BAND_FILL
    franja(ws, F, 1, 7)
    hdr_d = columnas[0][1] if columnas else ""
    hdr_e = columnas[1][1] if len(columnas) > 1 else ""
    for j2, h in enumerate(["Parametro", "", "Unidad", hdr_d, hdr_e,
                            "", "Referencia / Notas"], start=1):
        c = ws.cell(F + 1, j2, h.upper())
        c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
    com_modo = (f"Entrada: 'Interpolado' aplica la interpolacion lineal del "
               f"codigo entre T1 y T2 (fila {cita(F + 15)}/{cita(F + 16)}). "
               f"'Tabulado-conservador' adopta directamente el valor tabulado "
               f"en T2, sin interpolar. Rige el S(T) resuelto de {desc_cols} "
               f"(fila {cita(F + 21)}).")
    lab(F + 2, "Modo de lectura de S(T)   [MODO_S]", "Interpolado | Tabulado-conservador",
       com_modo)
    inp(f"D{F + 2}", "Interpolado", com_modo)
    dv_list(ws, f"D{F + 2}", '"Interpolado,Tabulado-conservador"', com_modo)
    lab(F + 3, "Temperatura de evaluacion", f"= {temp_fuente_cell}",
       f"Calculo: repite automaticamente la temperatura de evaluacion de esta "
       f"hoja ({temp_fuente_cell}). No se edita aqui.")
    ws.cell(F + 3, 4).value = f"={temp_fuente_cell}"
    # Los rotulos de unidad del bloque los decide la EDICION que se lee: la
    # metrica publica °C y MPa, la U.S. Customary °F y ksi. No se convierte
    # nada (regla 9): cambia de que tabla se lee, y el rotulo lo dice.
    u_temp = f'=IF({SEL},"°C","°F")' if us_ok else "°C"
    u_esf = f'=IF({SEL},"MPa","ksi")' if us_ok else "MPa"
    ws.cell(F + 3, 3).value = u_temp
    niveles = [(F + 4, "0 · Familia de material",
                f"Entrada: paso 0 de la cascada, para {desc_cols}. Elija la "
                f"familia de material de la lista desplegable. Habilita la "
                f"lista del paso 1 de esa misma columna."),
               (F + 5, "1 · Composicion nominal",
                "Entrada: paso 1 de la cascada, dependiente del paso 0 de la "
                "misma columna."),
               (F + 6, "2 · Forma de producto",
                "Entrada: paso 2 de la cascada, dependiente de los pasos 0-1 de la "
                "misma columna."),
               (F + 7, "3 · Especificacion (Spec. No.)",
                "Entrada: paso 3 de la cascada, dependiente de los pasos 0-2 de la "
                "misma columna."),
               (F + 8, "4 · Tipo / Grado",
                "Entrada: paso 4 de la cascada, dependiente de los pasos 0-3. Al "
                "completarlo ya se resuelve un material_id (salvo variantes, ver "
                "paso 5)."),
               (F + 9, "5 · Variante (clase / tamano) — opcional",
                "Entrada opcional: solo hace falta si el paso 4 deja mas de una fila "
                "posible. En blanco, se toma la primera variante encontrada.")]
    # La columna de notas decia "Lista desplegable en cascada" en las CINCO
    # filas. Repetir la misma frase cinco veces no informa: la convierte en
    # ruido y empuja fuera de la vista lo que si es propio de cada nivel. Se
    # dice una vez, en el nivel 0, y se declara que los de abajo dependen de el.
    for i, (r, t, com) in enumerate(niveles):
        lab(r, t, "Cascada de listas: cada nivel filtra el siguiente" if i == 0
            else None, com)
        for letra in letras:
            inp(f"{letra}{r}", com=com)

    # --- listas materializadas (columnas ocultas) --------------------------
    # K..X (o menos, si `columnas` trae una sola) son de uso exclusivo de esta
    # hoja: cada motor tiene su propio ws y su propio rango de columnas
    # ocultas, asi que dos motores en el mismo libro no colisionan aqui.
    FAM_COL = 11
    ws.cell(R_HDR, FAM_COL, "lista familia").font = SRC_F
    for k in range(1, 41):
        ws.cell(R_DATA + k - 1, FAM_COL).value = (
            f'=IF({M}=1,IFERROR(INDEX({rb["FAM"]},{k}),""),'
            f'IFERROR(INDEX({ri["FAM"]},{k}),""))')
    ws.column_dimensions[hl(FAM_COL)].hidden = True
    fam_ref = f'=${hl(FAM_COL)}${R_DATA}:${hl(FAM_COL)}${R_DATA + 39}'
    for letra in letras:
        dv_list(ws, f"{letra}{F + 4}", fam_ref, niveles[0][2])

    lv = [("C", "CK", "CV", F + 5, '{c}${r0}'.format(c="{c}", r0=F + 4)),
          ("F", "FK", "FV", F + 6, '{{c}}${r0}&"|"&{{c}}${r1}'.format(r0=F + 4, r1=F + 5)),
          ("S", "SK", "SV", F + 7,
           '{{c}}${r0}&"|"&{{c}}${r1}&"|"&{{c}}${r2}'.format(r0=F + 4, r1=F + 5, r2=F + 6)),
          ("G", "GK", "GV", F + 8,
           '{{c}}${r0}&"|"&{{c}}${r1}&"|"&{{c}}${r2}&"|"&{{c}}${r3}'
           .format(r0=F + 4, r1=F + 5, r2=F + 6, r3=F + 7)),
          ("V", "K4", "ID", F + 9,
           '{{c}}${r0}&"|"&{{c}}${r1}&"|"&{{c}}${r2}&"|"&{{c}}${r3}&"|"&{{c}}${r4}'
           .format(r0=F + 4, r1=F + 5, r2=F + 6, r3=F + 7, r4=F + 8))]
    niv_com = {F + 5: niveles[1][2], F + 6: niveles[2][2], F + 7: niveles[3][2],
              F + 8: niveles[4][2], F + 9: niveles[5][2]}
    col = FAM_COL + 1
    for letra, etiqueta in columnas:
        cl = f"${letra}"
        for tag, kk, vv, target, keyfmt in lv:
            key = keyfmt.format(c=cl)
            mx = max(1, min(int(max(rb.get("max" + tag, rb["maxV"]),
                                    ri.get("max" + tag, ri["maxV"]))), 250))
            L = hl(col)
            ws.cell(R_HDR, col, f"lista {tag} {etiqueta}").font = SRC_F
            for k in range(1, mx + 1):
                ws.cell(R_DATA + k - 1, col).value = (
                    f'=IF({M}=1,'
                    f'IF(COUNTIF({rb[kk]},{key})<{k},"",'
                    f'INDEX({rb[vv]},MATCH({key},{rb[kk]},0)+{k}-1)),'
                    f'IF(COUNTIF({ri[kk]},{key})<{k},"",'
                    f'INDEX({ri[vv]},MATCH({key},{ri[kk]},0)+{k}-1)))')
            ws.column_dimensions[L].hidden = True
            dv_list(ws, f'{letra}{target}',
                    f"=${L}${R_DATA}:${L}${R_DATA + mx - 1}", niv_com[target])
            col += 1

    AUX_LETTERS = ["W", "X", "Y", "Z"]
    aux = dict(zip(letras, AUX_LETTERS))
    for letra, ac in aux.items():
        cl = f"${letra}"
        ws[f"{ac}{F + 4}"] = (f'={cl}${F + 4}&"|"&{cl}${F + 5}&"|"&{cl}${F + 6}&"|"&'
                              f'{cl}${F + 7}&"|"&{cl}${F + 8}')
        ws[f"{ac}{F + 5}"] = (f'=IF({M}=1,IFERROR(MATCH(${ac}${F + 4},{rb["K4"]},0),0),'
                              f'IFERROR(MATCH(${ac}${F + 4},{ri["K4"]},0),0))')
        ws.column_dimensions[ac].hidden = True

    com_etiq = {
        F + 10: "Calculo: material_id resuelto por la cascada de esta columna "
             "(o el pegado a mano en la celda 'Variante' si la cascada no alcanza "
             "a identificar el material).",
        F + 11: "Calculo: hoja base de datos (B31.3, II-D 1A o II-D 1B/3) donde se "
             "encontro el material_id resuelto.",
        F + 12: f"Calculo: indice de la base de datos activa (1=B31.3, 2=II-D 1A, "
             f"3=II-D 1B/3), derivado del selector de aplicacion ({modo_cell}) "
             f"o del prefijo del material_id.",
        F + 13: "Calculo: numero de fila de la base de datos activa donde se localizo "
             "el material_id. Vacio si no se encuentra.",
        F + 14: "Calculo: cantidad de puntos tabulados (n_pts) que tiene esa fila en la "
             "banda compacta del codigo; determina el rango de la busqueda de T1/T2.",
        F + 15: f"Calculo: temperatura tabulada inmediatamente inferior (o igual) a la "
             f"temperatura de evaluacion (fila {cita(F + 3)}), tomada de la banda compacta.",
        F + 16: "Calculo: temperatura tabulada inmediatamente superior a la de "
             "evaluacion. Vacia si T1 es el ultimo punto tabulado.",
        F + 17: "Calculo: esfuerzo admisible S tabulado en T1, para el material "
             "resuelto.",
        F + 18: "Calculo: esfuerzo admisible S tabulado en T2. Vacio si T1 es el "
             "ultimo punto tabulado.",
        F + 19: "Calculo: temperatura maxima admisible o limite de aplicabilidad del "
             "material. Por encima de este valor, el Dictamen de rango marca FUERA "
             "DE RANGO (el codigo prohibe extrapolar).",
        F + 20: "Calculo: SIN MATERIAL SELECCIONADO si falta completar la cascada; "
             "MATERIAL NO ENCONTRADO si el material_id no aparece en la base; FUERA "
             "DE RANGO si no hay valor tabulado a esa T o si T supera la Temp. max.; "
             "OK en cualquier otro caso.",
        F + 21: f"Calculo: S(T) resuelto para esta columna — interpolacion lineal o "
             f"valor tabulado-conservador segun el Modo de lectura (fila {cita(F + 2)}), o "
             f"NA() si el Dictamen de rango no es OK. Alimenta {destino_st}.",
    }
    etiquetas = [(F + 10, "material_id resuelto", None), (F + 11, "Base de datos activa", None),
                 (F + 12, "Base ASME aplicada (1/2/3)", None), (F + 13, "Fila localizada en la base", None),
                 (F + 14, "Puntos tabulados de la fila", None),
                 (F + 15, "T1 — temperatura tabulada inferior", u_temp),
                 (F + 16, "T2 — temperatura tabulada superior", u_temp),
                 (F + 17, "S en T1", u_esf), (F + 18, "S en T2", u_esf),
                 (F + 19, "Temperatura maxima admisible", u_temp),
                 (F + 20, "Dictamen de rango", None),
                 (F + 21, "S(T) resuelto", u_esf)]
    for r, t, u in etiquetas:
        lab(r, t, com=com_etiq[r])
        if u:
            ws.cell(r, 3).value = u
    por_columna = {}
    for letra, _etiqueta in columnas:
        col2 = ord(letra) - 64
        L = hl(col2)
        ac = aux[letra]
        ws.cell(F + 10, col2).value = (
            f'=IF(${L}${F + 9}<>"",${L}${F + 9},IF(${ac}${F + 5}=0,"",'
            f'IF({M}=1,INDEX({rb["ID"]},${ac}${F + 5}),INDEX({ri["ID"]},${ac}${F + 5}))))')
        # Indice de base: 1/2/3 segun la base, +3 si el selector dice US. La
        # base la decide siempre el material_id METRICO (la cascada es unica y
        # se resuelve contra la edicion SI, igual que en los buscadores).
        base_i = f'IF({M}=1,1,IF(LEFT({L}{F + 10},2)="1A",2,3))'
        ws.cell(F + 12, col2).value = (
            f'={base_i}+IF({SEL},0,3)' if us_ok else f'={base_i}')
        ic = f"{L}{F + 12}"
        ic_si = f"MOD({ic}-1,3)+1" if us_ok else ic
        ws.cell(F + 11, col2).value = f'=IF({L}{F + 10}="","",{NAMES.format(i=ic)})'
        if us_ok:
            # fila SI -> clave bilingue -> fila US. Un material sin homologo
            # deja la fila US vacia, y de ahi en adelante todo el bloque se
            # comporta como si no hubiera seleccion: n_pts 0, S(T) NA() y el
            # dictamen lo dice con todas las letras. Nunca se cae a la fila
            # metrica por defecto, que daria un numero en la unidad equivocada.
            f_si = f'IFERROR(MATCH({L}{F + 10},{IDS.format(i=ic_si)},0),"")'
            ws[f"{ac}{F + 6}"] = (
                f'=IF({f_si}="","",INDEX({BI_SI.format(i=ic_si)},{f_si}))')
            bi = f"${ac}${F + 6}"
            ws[f"{ac}{F + 7}"] = (
                f'=IF({bi}="","",IFERROR(MATCH({bi},'
                f'{BI_US.format(i=ic_si)},0),""))')
            ws.cell(F + 13, col2).value = (
                f'=IF({SEL},IFERROR({f_si},""),${ac}${F + 7})')
        else:
            ws.cell(F + 13, col2).value = (
                f'=IFERROR(MATCH({L}{F + 10},{IDS.format(i=ic)},0),"")')
        FILA = f"{L}{F + 13}"
        ws.cell(F + 14, col2).value = f'=IF({FILA}="","",IFERROR(INDEX({NPTS.format(i=ic)},{FILA}),0))'
        NP = f"{L}{F + 14}"
        tr = f'OFFSET({TANC.format(i=ic)},{FILA}-1,0,1,MAX(1,{NP}))'
        vr = f'OFFSET({VANC.format(i=ic)},{FILA}-1,0,1,MAX(1,{NP}))'
        # $D{F+3}, no `temp_fuente_cell`: esa es la celda EXTERNA de la que
        # $D{F+3} copia su valor: el resto del bloque encadena sobre la copia
        # local, no sobre la externa, para no repetir la referencia externa
        # en cada formula.
        P1 = f'IFERROR(MATCH($D${F + 3},{tr},1),1)'
        ws.cell(F + 15, col2).value = f'=IF({FILA}="","",IFERROR(INDEX({tr},{P1}),""))'
        ws.cell(F + 16, col2).value = f'=IF({FILA}="","",IFERROR(INDEX({tr},{P1}+1),""))'
        ws.cell(F + 17, col2).value = f'=IF({FILA}="","",IFERROR(INDEX({vr},{P1}),""))'
        ws.cell(F + 18, col2).value = f'=IF({FILA}="","",IFERROR(INDEX({vr},{P1}+1),""))'
        ws.cell(F + 19, col2).value = f'=IFERROR(INDEX({TMAX.format(i=ic)},{FILA}),"")'
        sin_homologo = (
            f'IF(AND(NOT({SEL}),{FILA}=""),'
            f'"SIN EQUIVALENTE EN LA EDICION US",' if us_ok else "")
        ws.cell(F + 20, col2).value = (
            f'=IF({L}{F + 10}="","SIN MATERIAL SELECCIONADO",'
            + sin_homologo +
            f'IF({FILA}="","MATERIAL NO ENCONTRADO",'
            f'IF({L}{F + 17}="","FUERA DE RANGO (sin valor tabulado a esa T)",'
            f'IF(AND(ISNUMBER({L}{F + 19}),$D${F + 3}>{L}{F + 19}),'
            f'"FUERA DE RANGO (T > Temp. max.)","OK"))))'
            + (")" if us_ok else ""))
        ws.cell(F + 21, col2).value = (
            f'=IF({L}{F + 20}<>"OK",NA(),'
            f'IF($D${F + 3}<={L}{F + 15},{L}{F + 17},IF(OR({L}{F + 16}="",{L}{F + 18}=""),{L}{F + 17},'
            f'IF($D${F + 2}="Tabulado-conservador",{L}{F + 18},'
            f'{L}{F + 17}+({L}{F + 18}-{L}{F + 17})*($D${F + 3}-{L}{F + 15})/({L}{F + 16}-{L}{F + 15})))))')
        # S(T) resuelto: es el resultado del bloque, y por eso va en la
        # macrotipografia del sistema y no en negrita de dato.
        ws.cell(F + 21, col2).font = Font(name=MACRO, size=11, color=TINTA)
        for rr in range(F + 10, F + 22):
            _nota(ws.cell(rr, col2), com_etiq[rr])
        por_columna[letra] = {
            "material_id": f"${L}${F + 10}", "dictamen": f"${L}${F + 20}",
            "s_t": f"${L}${F + 21}", "tmax": f"${L}${F + 19}",
        }

    resultado = {"fila_banda": F, "fila_modo_s": F + 2, "fila_temp": F + 3,
                 "por_columna": por_columna}

    if incluir_ej_ec:
        fl, fa = fac_info["sheet"], fac_info["last_row"]
        com_ej = ("Entrada: elija la clave de junta longitudinal impresa en la Tabla "
                 "A-3 del B31.3 (tipo de junta/costura). Determina el factor de "
                 "eficiencia de junta longitudinal Ej (columna G de MAP_Factores) que "
                 "alimenta la ecuacion de espesor de pared requerido de esta hoja.")
        lab(F + 23, "Clave de junta longitudinal (Tabla A-3) -> Ej", "Lista desplegable de "
           "MAP_Factores", com_ej)
        inp(f"D{F + 23}", "A106 | Seamless pipe", com_ej)
        dv_list(ws, f"D{F + 23}", f"={fl}!$A${R_DATA}:$A${fa}", com_ej)
        ws.cell(F + 23, 3).value = "adimensional"
        ws.cell(F + 23, 5).value = (f'=IFERROR(INDEX({fl}!$G${R_DATA}:$G${fa},'
                                 f'MATCH(${hl(4)}${F + 23},{fl}!$A${R_DATA}:$A${fa},0)),1)')
        _nota(ws.cell(F + 23, 5), "Calculo: Ej por lookup de la clave elegida contra "
                              "la Tabla A-3 de MAP_Factores; 1 si no se encuentra.")
        com_ec = ("Entrada informativa: clave de fundicion (Tabla A-2 del B31.3) para "
                 "consultar su factor de eficiencia Ec. No alimenta ningun calculo de "
                 "esta hoja; se deja como referencia si algun material es de fundicion.")
        lab(F + 24, "Ec (fundicion, Tabla A-2) — informativo", com=com_ec)
        inp(f"D{F + 24}", com=com_ec)
        dv_list(ws, f"D{F + 24}", f"={fl}!$A${R_DATA}:$A${fa}", com_ec)
        ws.cell(F + 24, 3).value = "adimensional"
        ws.cell(F + 24, 5).value = (f'=IFERROR(INDEX({fl}!$G${R_DATA}:$G${fa},'
                                 f'MATCH(${hl(4)}${F + 24},{fl}!$A${R_DATA}:$A${fa},0)),"")')
        _nota(ws.cell(F + 24, 5), "Calculo: Ec por lookup de la clave elegida contra "
                              "la Tabla A-2 de MAP_Factores; vacio si no se encuentra.")
        resultado["ej_ref"] = f"${hl(5)}${F + 23}"
        resultado["ec_ref"] = f"${hl(5)}${F + 24}"
        nota_row = F + 26
    else:
        nota_row = F + 23

    nota_cascada = ws.cell(nota_row, 1, "La cascada filtra la base ASME por familia, "
                    "composicion nominal, forma de producto, especificacion y "
                    "tipo/grado. Si un material no aparece, localicelo en el "
                    "buscador correspondiente y pegue su material_id en la celda "
                    "'Variante'." + (f" {nota_extra_cascada}" if nota_extra_cascada else ""))
    nota_cascada.font = SRC_F
    _nota(nota_cascada, "Aviso fijo: como usar la cascada de esta seccion y que "
                        "hacer si un material no aparece en ella. No se edita.")

    # --- Fase 4: el rastro de la resolucion se pliega -----------------------
    # Once filas de trazabilidad (indice de base, fila localizada, puntos
    # tabulados, T1/T2, S en T1/T2, temperatura maxima) entre la cascada y el
    # resultado. Son el rastro que permite auditar de donde sale el S(T) y NO
    # se borran —Regla n.1—, pero abiertas empujan fuera de la pantalla
    # justamente lo que el ingeniero vino a ver.
    #
    # Se agrupan en un outline de Excel, plegado por defecto: el "+" del margen
    # las despliega. Se pliega de `material_id resuelto` a `Temperatura maxima`
    # (F+10..F+19) y se dejan SIEMPRE visibles las dos filas que cierran el
    # bloque: `Dictamen de rango` (F+20) y `S(T) resuelto` (F+21). El dictamen
    # no se pliega aunque la decision 4 del plan lo listara: es lo que BLOQUEA
    # el calculo, y una condicion de bloqueo escondida detras de un "+" es una
    # condicion que nadie ve.
    for r in range(F + 10, F + 20):
        dim = ws.row_dimensions[r]
        dim.outline_level = 1
        dim.hidden = True
    # summaryBelow: el "+" aparece junto a la fila-resumen, que aqui queda
    # DEBAJO del grupo (el dictamen y el S(T) resuelto).
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_view.showOutlineSymbols = True

    return resultado


# material_id del caso precargado (Linea 12"-CWS-46-032-B1), verificados contra
# DB_B31_3. Se siembran en la celda 'Variante' (paso 5) de la seccion 7, que es la
# via de escape declarada de la cascada (fila 115: D115=IF(D114<>"",D114,...)): un
# unico MATCH del material_id contra la columna A de la base, sin depender de las
# columnas ocultas de la cascada. El build ABORTA si el id no esta en la base.
SEED_ART212_BASE = 'A-1 | A106 | B | Pipe & tube | K03006 | P-1'      # tuberia (D)
SEED_ART212_COLLAR = 'A-1 | A516 | 70 | Plate, bar, shps., sheet | K02700 | P-1'  # (E)

# Ruta de los dos apendices del App. 501 en resources/ (reparados en la Fase 0.2
# del plan del Art. 212: la extraccion habia colapsado la capa de texto).
_ART_212_JSON = ("ASME PCC/pcc_2/p2_welded_repairs/"
                 "art_212_fillet_welded_patches/art_212.json")
_ART_206_JSON = ("ASME PCC/pcc_2/p2_welded_repairs/"
                 "art_206_full_encirclement_steel/art_206.json")
_APP_501_II = ("ASME PCC/pcc_2/p5_examination/art_501_pressure_tightness/app/"
               "app_501_ii/app_501_ii.json")
_APP_501_III = ("ASME PCC/pcc_2/p5_examination/art_501_pressure_tightness/app/"
                "app_501_iii/app_501_iii.json")


# Umbrales de LONGITUD que los dos motores comparan contra una entrada. Son
# normativos, y el codigo los imprime en las DOS unidades —"40 mm (1.5 in.)",
# "2.5 mm (3/32 in.)"—, asi que en modo US se LEEN, no se convierten (Regla n.1
# y regla 9). Cada entrada es (articulo, fragmento que lo ancla en el texto).
# Si el fragmento deja de aparecer, el build aborta: un umbral de aceptacion en
# la unidad equivocada convierte un "NO CUMPLE" en un "CUMPLE".
UMBRALES_PCC2 = {
    "212_filete_max": ("212", "nor 40 mm"),
    "212_separacion_g": ("212", "a minimum of 1.5 mm"),
    "212_separacion_max": ("212", "separated by more than 5 mm"),
    "212_espesor_examen": ("212", "greater than 25 mm"),
    "212_solape": ("212", "overlap sound base metal by at least 25 mm"),
    "206_longitud_min": ("206", "at least 100 mm"),
    "206_sobrepaso": ("206", "extend beyond the defect by at least 50mm"),
    "206_luz_radial": ("206", "radial gap of up to 2.5 mm"),
}
# "1 ∕ 16" y "3 ∕ 32" llegan como fraccion con el signo de division de Unicode.
_RE_UMBRAL = re.compile(
    r"([\d.]+)\s*mm\s*\(\s*(?:([\d.]+)|([\d.]+)\s*[∕/]\s*([\d.]+))\s*in\.?\s*\)")


def umbral(umbrales, clave, sel):
    """`IF(es_SI, valor metrico, valor US)` de un umbral leido de resources/."""
    si, us = umbrales[clave]
    return f"IF({sel},{si:g},{us:g})"


def leer_umbrales_pcc2(resources):
    """Los umbrales de longitud de PCC-2, en sus dos unidades impresas."""
    from pathlib import Path as _Path
    raiz = _Path(resources)
    textos = {}
    for art, ruta in (("212", _ART_212_JSON), ("206", _ART_206_JSON)):
        textos[art] = json.dumps(
            json.loads((raiz / ruta).read_text(encoding="utf-8")),
            ensure_ascii=False)
    fuera = {}
    for clave, (art, ancla) in UMBRALES_PCC2.items():
        i = textos[art].find(ancla)
        if i < 0:
            raise SystemExit(
                f"Fase 7: el umbral '{clave}' ya no aparece en el Art. {art} de "
                f"resources/ (ancla {ancla!r}). No se construye el modo US con "
                f"un umbral de aceptacion convertido de memoria.")
        m = _RE_UMBRAL.search(textos[art], i)
        if not m:
            raise SystemExit(
                f"Fase 7: el umbral '{clave}' no imprime su valor en pulgadas "
                f"junto al metrico; no se puede leer el par (Regla n.1).")
        us = float(m[2]) if m[2] else float(m[3]) / float(m[4])
        fuera[clave] = (float(m[1]), us)
    return fuera


def leer_energia_501(resources):
    """Lee de resources/ (App. 501-II y 501-III) los coeficientes de la energia
    almacenada y la distancia segura de una prueba neumatica (Paso 8 del Art.
    212). NINGUN valor sale de memoria: todos se leen del JSON reparado en la
    Fase 0.2 (Regla n.1). ABORTA si el apendice no esta reparado (texto
    colapsado o amendment ausente): sin fuente no se construye el Paso 8.

    Devuelve dict con:
      tnt_div_kg   divisor de la ec. (II-3): TNT = E / 4 266 920 (kg)
      blast_thr_J  umbral de la 501-III-1: R = 30 m si E <= 8 130 000 J
      blast_R_m    esa distancia fija (30 m)
      r_scaled_def Rscaled por defecto de la eq. (III-1) (20 m/kg^1/3)
      r_scaled_ops [(etiqueta, valor_si)] de la Tabla 501-III-1-1 para el dv_list
    """
    import re as _re
    from pathlib import Path as _Path
    raiz = _Path(resources)
    ii = json.loads((raiz / _APP_501_II).read_text(encoding="utf-8"))
    iii = json.loads((raiz / _APP_501_III).read_text(encoding="utf-8"))

    def _texto_ecuacion(doc, rotulo):
        for b in doc.get("blocks", []):
            if isinstance(b, dict) and b.get("type") == "equation" \
                    and rotulo in (b.get("text") or ""):
                return b["text"]
        return None

    # (II-3): TNT = E / 4 266 920 (kg). El divisor se lee del texto reparado.
    t_ii3 = _texto_ecuacion(ii, "(II-3)")
    if not t_ii3 or any(g in t_ii3 for g in "ÄÅÇÉÑÖ�"):
        raise SystemExit(
            "Art.212 Paso 8: la ec. (II-3) del App. 501-II no esta legible en "
            "resources/ (corra la Fase 0.2: completar_app_501_energia.py).")
    m = _re.search(r"E\s*/\s*([\d\s]+)", t_ii3)
    if not m:
        raise SystemExit(f"Art.212 Paso 8: no se pudo leer el divisor TNT de {t_ii3!r}")
    tnt_div = int(m.group(1).replace(" ", ""))

    # La Tabla 501-III-1-1 y sus reglas viven en el amendment de la Fase 0.2.
    amd = iii.get("extraction_amendments") or {}
    tabla = amd.get("tabla_501_iii_1_1")
    if not tabla:
        raise SystemExit(
            "Art.212 Paso 8: falta la Tabla 501-III-1-1 en resources/ "
            "(corra la Fase 0.2: completar_app_501_energia.py).")
    # Umbral y distancia fija del blast wave (501-III-1): "R = 30 m ... para
    # E <= 8 130 000 J". Se leen del texto de la regla, no de memoria.
    ub = tabla["umbral_blast"]
    m_thr = _re.search(r"E\s*<=\s*([\d\s]+)\s*J", ub)
    m_r = _re.search(r"R\s*=\s*([\d.]+)\s*m", ub)
    if not (m_thr and m_r):
        raise SystemExit(f"Art.212 Paso 8: no se pudo leer el umbral/distancia de {ub!r}")
    blast_thr = int(m_thr.group(1).replace(" ", ""))
    blast_r = float(m_r.group(1))
    # Rscaled por defecto (>= 20) de la regla.
    m_def = _re.search(r"([\d.]+)\s*m/kg", tabla["regla_por_defecto"])
    r_def = float(m_def.group(1)) if m_def else float(tabla["filas"][0][0])
    # Opciones de Rscaled por criterio para el dv_list (valor SI = columna 0).
    ops = []
    for fila in tabla["filas"]:
        val = float(fila[0])
        crit = " / ".join(c for c in (fila[2], fila[3]) if c and c != "...")
        ops.append((f"{val:g} ({crit})" if crit else f"{val:g}", val))

    # --- Los MISMOS coeficientes en U.S. Customary (Fase 7) -----------------
    # El App. 501 publica las dos ediciones y resources/ las trae enteras, asi
    # que el modo US del Paso 8 no obliga a convertir nada (regla 9): se lee lo
    # que imprime el codigo. Si faltara, se aborta igual que con la metrica —
    # un Paso 8 en US con constantes inventadas no vale nada.
    #   (II-5): TNT = E / 1,488,617 (lb)
    #   501-III-1: R = 100 ft para E <= 6 000 000 ft-lb
    #   Tabla 501-III-1-1, columna 1: Rscaled en ft/lb^(1/3)
    t_ii5 = _texto_ecuacion(ii, "(II-5)")
    if not t_ii5:
        raise SystemExit(
            "Art.212 Paso 8 (US): falta la ec. (II-5) del App. 501-II en "
            "resources/. El modo US no se construye con un divisor inventado.")
    m_us = _re.search(r"E\s*/\s*([\d\s,]+)", t_ii5)
    if not m_us:
        raise SystemExit(f"Art.212 Paso 8 (US): divisor TNT ilegible en {t_ii5!r}")
    tnt_div_us = int(m_us.group(1).replace(" ", "").replace(",", ""))

    m_thr_us = _re.search(r"\(([\d\s,]+)\s*ft-lb\)", ub)
    m_r_us = _re.search(r"\(([\d.]+)\s*ft\)", ub)
    if not (m_thr_us and m_r_us):
        raise SystemExit(f"Art.212 Paso 8 (US): umbral/distancia ilegibles en {ub!r}")
    blast_thr_us = int(m_thr_us.group(1).replace(" ", "").replace(",", ""))
    blast_r_us = float(m_r_us.group(1))

    m_def_us = _re.search(r"\(([\d.]+)\s*ft/lb", tabla["regla_por_defecto"])
    r_def_us = float(m_def_us.group(1)) if m_def_us else float(tabla["filas"][0][1])
    ops_us = []
    for fila in tabla["filas"]:
        val = float(fila[1])
        crit = " / ".join(c for c in (fila[2], fila[3]) if c and c != "...")
        ops_us.append((f"{val:g} ({crit})" if crit else f"{val:g}", val))

    return {"tnt_div_kg": tnt_div, "blast_thr_J": blast_thr, "blast_R_m": blast_r,
            "r_scaled_def": r_def, "r_scaled_ops": ops,
            "tnt_div_lb": tnt_div_us, "blast_thr_ftlb": blast_thr_us,
            "blast_R_ft": blast_r_us, "r_scaled_def_us": r_def_us,
            "r_scaled_ops_us": ops_us}


# --- Notacion de simbolos con subindice real (Fase 9, decision 5) -----------
# Los simbolos con guion bajo (F_m, S_w,m, T_s, w_min...) se muestran con el
# subindice REAL en vez del feo "_". Como el guion bajo aparece SOLO en los
# simbolos (los formulas y las bandas no lo usan), un unico paso barre las
# columnas de rotulo/simbolo y convierte cada token "base_sub" en un
# CellRichText con el subindice en vertAlign="subscript". El texto plano
# reconstruible (base + "_" + sub) sigue casando con el oracle Rev0, y el
# render visual (base+sub, sin "_") satisface el criterio de la decision 5.
_RE_SIMBOLO_SUB = re.compile(
    r'([A-Za-z][A-Za-z0-9]*)_([A-Za-z0-9,./áéíóúÁÉÍÓÚñ]+)')
_FF_TINTA = "FF" + TINTA  # aRGB para el InlineFont (color del sistema)


def sym(base, sub):
    """Un simbolo con subindice real: `base` normal + `sub` en subscript, ambos
    en la fuente MONO del sistema y color TINTA. Devuelve un CellRichText."""
    normal = InlineFont(rFont=MONO, sz=10, color=_FF_TINTA)
    baja = InlineFont(rFont=MONO, sz=8, color=_FF_TINTA, vertAlign="subscript")
    return CellRichText(TextBlock(normal, base), TextBlock(baja, sub))


def _aplicar_subindices(ws, cols=("B",)):
    """Convierte los simbolos 'base_sub' de las columnas `cols` en CellRichText
    con subindice real. Por defecto SOLO la columna B (la de simbolos aislados):
    la columna A es prosa/rotulos y de la Seccion 7 llega con guiones bajos que
    NO son subindices (rutas de resources/, tags como '[MODO_S]', nombres de
    campo), asi que no se toca. Deja intactas las celdas sin '_'. No toca
    formulas (viven en D-G) ni valores que no sean str."""
    normal = InlineFont(rFont=MONO, sz=10, color=_FF_TINTA)
    baja = InlineFont(rFont=MONO, sz=8, color=_FF_TINTA, vertAlign="subscript")
    for col in cols:
        for cell in ws[col]:
            v = cell.value
            if not isinstance(v, str) or "_" not in v:
                continue
            # Un simbolo es corto y no es formula ni prosa: la columna B tambien
            # aloja formulas (specs, B93/B95/B99: empiezan por "=") y textos
            # largos (B94/B96/B97/B98). Sus guiones bajos ("P_diseno" dentro de
            # una frase) NO son subindices. Se descartan por longitud y por "=".
            if v.startswith("=") or len(v) > 12:
                continue
            if not _RE_SIMBOLO_SUB.search(v):
                continue
            partes, ultimo = [], 0
            for m in _RE_SIMBOLO_SUB.finditer(v):
                if m.start() > ultimo:
                    partes.append(TextBlock(normal, v[ultimo:m.start()]))
                partes.append(TextBlock(normal, m.group(1)))
                partes.append(TextBlock(baja, m.group(2)))
                ultimo = m.end()
            if ultimo < len(v):
                partes.append(TextBlock(normal, v[ultimo:]))
            cell.value = CellRichText(*partes)


def build_parche_art212(wb, b313, iid1a, iidb, fac_info, rangos, b3610, b3619,
                        energia_501=None, b313c=None, iid1ac=None, iidbc=None,
                        umbrales=None):
    """Motor Art. 212 (parche soldado), 100% en codigo — desanclado del maestro
    Rev0 (Fase 4). Antes vivia heredado + corregido por integrate_motor/
    corregir_art212_fase1; ahora nace con new_sheet como Collar_PCC2_Art206.
    Reutiliza construir_seccion7_material (Seccion 7, compartida con el 206).
    Las formulas de las secciones 1-6 se transcriben del oracle
    parche_art212_ref.json (Regla n.1: no se reescriben de memoria) y
    TestParidadHojaParche las fija al caracter, SALVO el bloque dimensional
    (Tareas 7-8): NPS/cedula/OD/espesor y el selector de norma salen ahora de
    DB_B36_10/DB_B36_19 por cascada de listas (reglas 12/14), y sus celdas se
    declaran en DIVERGENCIAS_REEMPLAZADAS, verificadas por
    TestCascadaDimensionalArt212 en vez de por el oracle Rev0."""
    if umbrales is None:
        raise SystemExit(
            "build_parche_art212: faltan los umbrales de longitud de PCC-2 (leer_umbrales_pcc2). "
            "Un umbral de aceptacion no tiene valor por defecto.")
    UMBR = umbrales

    if MOTOR in wb.sheetnames:        # el maestro aun trae la hoja heredada
        del wb[MOTOR]
    ws = new_sheet(
        wb, MOTOR,
        "MOTOR DE CALCULO — PARCHE SOLDADO (ASME PCC-2 Art. 212)",
        'ASME PCC-2 Art. 212 (Fillet Welded Patches). El collar de encierro '
        'total (Art. 206) es un motor aparte. Caso precargado: Linea '
        '12"-CWS-46-032-B1 (U46).')
    # Columna A a 50: con 44, el rotulo mas largo de la hoja —«Presion de diseno
    # (maxima admisible / rating)»— se cortaba a media palabra contra la columna
    # de simbolo, que no esta vacia y por tanto no deja desbordar el texto. Visto
    # exportando la hoja a PDF en la Fase 11, no en openpyxl.
    autosize(ws, {"A": 50, "B": 8, "C": 14, "D": 16, "E": 16, "F": 16, "G": 46})
    # new_sheet() escribe A1/A2 con rotulo() (formato de banda: "[ ... ]" en
    # mayusculas). El titulo de ESTA hoja va literal, tal como lo capturo el
    # *oracle* desde el maestro Rev0 (corregir_art212_fase1 lo hacia igual,
    # sobreescribiendo el A1/A2 heredado del maestro) — por eso se reescribe
    # aqui sin pasar por rotulo(). Solo cambia el VALOR de la celda; la fuente
    # y el relleno de titulo que puso new_sheet (TITLE_F/TITLE_FILL, tokens
    # del sistema) se conservan.
    ws["A1"] = "MOTOR DE CALCULO — PARCHE SOLDADO (ASME PCC-2 Art. 212)"
    ws["A2"] = ("ASME PCC-2 Art. 212 (Fillet Welded Patches). El collar de encierro "
                "total (Art. 206) es un motor aparte. Caso precargado: Linea "
                "12\"-CWS-46-032-B1 (U46).")
    # new_sheet() no fusiona A1/A2 (las hojas de datos solo colorean la fila con
    # cabecera_hoja()); el maestro Rev0 SI las fusionaba A1:G1/A2:G2, y el
    # *oracle* lo capturo asi. Se replica aqui, no en new_sheet(), porque es
    # propio de esta hoja heredada, igual que la sobreescritura literal de
    # arriba.
    ws.merge_cells("A1:G1")
    ws.merge_cells("A2:G2")

    def lab(r, text, unidad=None, ref=None, com=None):
        cl = ws.cell(r, 1, text)
        cl.font = Font(name=MONO, size=10, color=TINTA)
        if unidad:
            ws.cell(r, 3).value = unidad
        if ref:
            ws.cell(r, 7, ref).font = Font(name=MONO, size=9, color=GRIS)
        if com:
            _nota(cl, com)

    def inp(cell, value="", com=None):
        c = ws[cell]
        c.value = value
        c.font, c.fill, c.border = IN_F, MOTOR_IN_FILL, CAJA_CAMPO
        c.protection = Protection(locked=False)
        if com:
            _nota(c, com)

    def calc(cell, formula, com=None):
        c = ws[cell]
        c.value = formula
        if com:
            _nota(c, com)
        return c

    # Esta hoja no usa un band()/header() que apliquen rotulo()/.upper(): el
    # *oracle* (capturado del maestro Rev0) trae el texto de banda y de
    # encabezado literal, con mayus/minus mixtas y sin los corchetes que
    # rotulo() añadiria — band()/header() "genericos" (mayusculas forzadas,
    # texto envuelto en "[ ... ]") no reproducirian esa celda letra a letra.
    # banda_literal()/encabezado() escriben el VALOR tal cual, con el mismo
    # estilo (fuente/relleno/franja o font/fill/border) que usarian los
    # genericos.
    def banda_literal(r, texto):
        ws.merge_cells(f"A{r}:G{r}")
        ws.cell(r, 1, texto)
        ws.cell(r, 1).font = Font(name=MACRO, size=11, color=PAPEL)
        for j in range(1, 8):
            ws.cell(r, j).fill = BAND_FILL
        franja(ws, r, 1, 7)

    def encabezado(r, pares):
        for j, h in pares:
            c = ws.cell(r, j, h)
            c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA

    # --- Seccion 7: cascada de material base + collar/parche ----------------
    construir_seccion7_material(
        ws, 105, b313, iid1a, iidb, fac_info, rangos,
        columnas=(("D", "Metal base"), ("E", "Collar / parche")),
        modo_cell="$D$11", temp_fuente_cell="$D$25", incluir_ej_ec=True,
        destino_st="D67/D68 de la seccion 3",
        nota_extra_cascada="El material se resuelve por esta cascada contra las "
        "bases auditadas (DB_B31_3 / DB_BPVC_IID); no hay lista corta de respaldo.",
        # Banda literal, sin rotulo(): las del 212 van en mayus/minus mixtas y
        # sin corchetes (el *oracle* Rev0 las trae asi). Ver banda_literal().
        titulo_banda="2.  RESOLUCIÓN DE MATERIAL",
        banda_rotulo=False, mapa_citas=MAPA_FILAS_212,
        b313c=b313c, iid1ac=iid1ac, iidbc=iidbc, unidad_cell=UNIDAD_212)

    # --- Banda IDENTIFICACION (fila 4) + identificacion (filas 5-6) ---------
    # A4 es la banda de seccion, fusionada A4:G4, texto literal del *oracle*
    # ("IDENTIFICACIÓN", sin numeral) transcrito sin pasar por rotulo() —
    # rotulo() lo envolveria en "[ ... ]" y forzaria mayusculas, y el *oracle*
    # no trae corchetes — mismo criterio que A1/A2. Van tambien aqui los
    # merges D5:F5/D6:F6 y los valores G5/G6 (revision/unidad del documento).
    banda_literal(4, "IDENTIFICACIÓN")

    lab(5, "Documento", ref="Rev.: 0",
       com="Entrada: identificador del documento de este calculo (numero de MC).")
    inp("D5", "MC-REP-U46-CWS-032",
        com="Entrada: identificador del documento de este calculo (numero de MC).")
    lab(6, "Componente / servicio", ref="Unidad: U46",
       com="Entrada: descripcion del componente y el servicio reparado.")
    inp("D6", 'Cuello de brida WN 12" Cl.150 · Agua de enfriamiento',
        com="Entrada: descripcion del componente y el servicio reparado.")
    # D5:F5/D6:F6 fusionados (el *oracle* los trae asi: D6 lleva un texto
    # largo que necesita el ancho de las tres columnas). El RELLENO de inp()
    # se hereda de la celda ancla al fusionar, pero el BORDE no (la misma
    # trampa que ya documenta franja() en este archivo) — se replica el
    # mismo CAJA_CAMPO en E/F para que la linea inferior cruce todo el campo.
    ws.merge_cells("D5:F5")
    ws.merge_cells("D6:F6")
    for celda in ("E5", "F5", "E6", "F6"):
        ws[celda].border = CAJA_CAMPO

    # --- Cableado Seccion 7 -> Secciones 1/3 (D39/D40/D44/F90) --------------
    # Transcrito literal del oracle:
    ws["D39"] = '=IF($E$125="OK",$E$126,NA())'
    ws["G39"] = "S(T) del collar — base de datos ASME (seccion 2)"
    ws["G39"].font = Font(name=MONO, size=9, color=GRIS)
    _nota(ws["D39"], "Calculo: trae el S(T) del collar/parche resuelto en la seccion "
                    "7 (E58) si su Dictamen de rango es OK; NA() si no. Alimenta "
                    "'Esf. admisible del collar' de la seccion 1 (D39 original del "
                    "maestro queda sustituido por este valor).")
    ws["D40"] = '=IF($D$125="OK",$D$126,NA())'
    ws["G40"] = "S(T) del metal base — base de datos ASME (seccion 2)"
    ws["G40"].font = Font(name=MONO, size=9, color=GRIS)
    _nota(ws["D40"], "Calculo: trae el S(T) del metal base resuelto en la seccion 2 "
                    "(D58) si su Dictamen de rango es OK; NA() si no.")
    ws["D44"] = "=$E$128"
    ws["G44"] = "Ej por lookup (B31.3 Tabla A-3) — seccion 2"
    ws["G44"].font = Font(name=MONO, size=9, color=GRIS)
    ws["D44"].font = Font(name=MONO, size=10, color=TINTA)
    ws["D44"].fill = PAPEL_FILL
    ws["D44"].protection = Protection(locked=True)
    _nota(ws["D44"], "Calculo: repite el Ej resuelto en la seccion 2 (E60) a partir "
                    "de la clave de junta longitudinal elegida en D128. No se edita "
                    "aqui.")
    # El dictamen distingue "sin material" de "fuera de rango": son estados
    # distintos y confundirlos haria leer un formulario recien abierto (cascada
    # vacia) como un material rechazado por temperatura. "ELIJA MATERIAL" invita
    # a completar la seccion 7; "FUERA DE RANGO" solo aparece cuando ya hay
    # material y la T supera lo que el codigo tabula.
    # F90 antepone la compuerta de elegibilidad del Paso 1 (D140, bloque nuevo
    # del anexo de flujo): si el Paso 1 arroja un estado BLOQUEANTE (PROHIBIDO
    # por servicio letal, NO ELEGIBLE por dano/grieta, o FUERA DE ALCANCE por
    # T > 345), el dictamen global es ese motivo, en rojo, sin evaluar los
    # CUMPLE de la seccion 5. El aviso de entalla (REVISAR, T < 0) NO bloquea:
    # es advertencia, deja seguir. Trazado: Art. 212-1/2 (bloques 6,11-13) y el
    # flujo aprobado del ingeniero (servicio letal).
    # F84-F87: verificaciones de la seccion 5. F161/F162: los dos topes del
    # filete del Paso 4 (Fase 4). Todos deben dar CUMPLE para APTO.
    ws["F90"] = ('=IF(OR(LEFT($D$140,9)="PROHIBIDO",LEFT($D$140,11)="NO ELEGIBLE",'
                 'LEFT($D$140,16)="FUERA DE ALCANCE"),$D$140,'
                 'IF(OR($D$125="SIN MATERIAL SELECCIONADO",'
                 '$E$125="SIN MATERIAL SELECCIONADO"),"ELIJA MATERIAL (Seccion 2)",'
                 'IF(OR($D$125<>"OK",$E$125<>"OK"),"REVISAR — MATERIAL FUERA DE RANGO",'
                 'IF(AND(F84="CUMPLE",F85="CUMPLE",F86="CUMPLE",F87="CUMPLE",'
                 'F161="CUMPLE",F162="CUMPLE"),'
                 '"APTO","REVISAR"))))')
    _nota(ws["F90"], "Calculo: DICTAMEN GLOBAL DEL DISEÑO. Primero la compuerta de "
                    "elegibilidad (Paso 1, D140): si es PROHIBIDO / NO ELEGIBLE / "
                    "FUERA DE ALCANCE, ese es el dictamen y no se evalua nada mas. "
                    "Si es ELEGIBLE (o el aviso de entalla, que no bloquea): ELIJA "
                    "MATERIAL (Seccion 2) si falta seleccionar; REVISAR — MATERIAL "
                    "FUERA DE RANGO si hay material fuera de rango de T; APTO solo si "
                    "ademas las 4 verificaciones de la seccion 5 dan CUMPLE; REVISAR "
                    "en cualquier otro caso.")

    # --- Caso precargado (seed por la celda Variante, con guardia Regla n.1)--
    db = wb["DB_B31_3"]
    ids = {db.cell(r, 1).value for r in range(R_DATA, db.max_row + 1)}
    for etiqueta, mid in (("base", SEED_ART212_BASE), ("collar", SEED_ART212_COLLAR)):
        if mid not in ids:
            raise SystemExit(
                f"Art.212: el material_id sembrado ({etiqueta}) no existe en "
                f"DB_B31_3: {mid!r} (Regla n.1).")
    ws["D114"] = SEED_ART212_BASE
    ws["E114"] = SEED_ART212_COLLAR
    _nota(ws["D114"], "Entrada (Variante): sembrada con el material_id del metal base "
                     "del caso precargado (A106 Gr.B). Es la via de escape de la "
                     "cascada; cambiela por el material de su caso o vacie y use la "
                     "cascada (pasos 0-4).")
    _nota(ws["E114"], "Entrada (Variante): sembrada con el material_id del collar/"
                     "parche del caso precargado (A516 Gr.70). Cambiela por el "
                     "material de su caso o vacie y use la cascada (pasos 0-4).")

    # --- Banda APLICACION Y CODIGO (fila 8) + encabezado (fila 9) -----------
    # El *oracle* trae, antes de su rango 10-14, esta banda (A8, fusionada
    # A8:G8, texto literal sin rotulo()) y esta fila de encabezado (A9:G9,
    # mayus/minus mixtas, incompatible con un header en mayusculas). Mismo
    # criterio que A1/A2: se escribe el VALOR literal con el estilo de
    # banda/encabezado, sin forzar mayusculas.
    banda_literal(8, "APLICACIÓN Y CÓDIGO DE CONSTRUCCIÓN")

    encabezado(9, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                   (4, "Valor"), (7, "Referencia / Notas")))

    # --- Aplicacion y codigo de construccion (filas 10-14) ------------------
    # Columna B (Simbolo) de estas cinco filas no la escribe ninguno de los
    # cinco helpers (lab/inp/calc/band/header): lab() solo cubre A+C+G. El
    # oracle trae ahi "MODO"/"kf" en D11/D14 y "—" de relleno en las demas,
    # asi que se escribe directo con el mismo font que usa lab() en columna A.
    com10 = ("Entrada: selector que conmuta el modo geometrico (tuberia, virola "
             "cilindrica o cabezal/esfera) y por tanto el codigo de construccion, "
             "la fuente del esfuerzo admisible y el factor kf de toda la hoja.")
    lab(10, "Aplicación / geometría", unidad="—", ref="Lista desplegable", com=com10)
    ws.cell(10, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D10", "Tubería (B31.3)", com10)
    dv_list(ws, "D10",
            '"Tubería (B31.3),Virola cilíndrica (VIII-1),Cabezal/esfera (VIII-1)"',
            com10)

    com11 = ("Calculo: MODO = 1 si D10 es 'Tuberia (B31.3)', 3 si es 'Cabezal/"
             "esfera (VIII-1)', 2 en cualquier otro caso (virola cilindrica).")
    lab(11, "Modo (1=tubería, 2=virola, 3=cabezal)", unidad="—",
        ref="Derivado del selector", com=com11)
    ws.cell(11, 2, "MODO").font = Font(name=MONO, size=10, color=TINTA)
    calc("D11", '=IF($D$10="Tubería (B31.3)",1,IF($D$10="Cabezal/esfera (VIII-1)",3,2))',
         com11)

    com12 = ("Calculo: ASME B31.3 si MODO=1 (tuberia); ASME BPVC VIII-1 en los "
             "demas casos (virola o cabezal/esfera).")
    lab(12, "Código de construcción", unidad="—", ref="Automático", com=com12)
    ws.cell(12, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    calc("D12", '=IF($D$11=1,"ASME B31.3","ASME BPVC VIII-1")', com12)

    com13 = ("Calculo: cita la tabla del codigo activo de la que sale el esfuerzo "
             "admisible S — Tabla A-1 del B31.3 si MODO=1, Tabla 1A de ASME II-D "
             "en los demas casos.")
    lab(13, "Fuente del esfuerzo admisible", unidad="—", ref="Automático", com=com13)
    ws.cell(13, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    calc("D13", '=IF($D$11=1,"B31.3 Tabla A-1","ASME II-D Tabla 1A")', com13)

    com14 = ("Calculo: factor de geometria de la ecuacion de membrana, kf = 0,25 "
             "para esfera (MODO=3) o 0,5 para cilindro (tuberia o virola).")
    lab(14, "Factor de geometría (membrana)", unidad="—",
        ref="Cil.=0,5 · Esfera=0,25", com=com14)
    ws.cell(14, 2, "kf").font = Font(name=MONO, size=10, color=TINTA)
    calc("D14", '=IF($D$11=3,0.25,0.5)', com14)

    # Fase 7: selector de sistema de unidades. Va en la banda de APLICACION,
    # por ENCIMA de los datos de entrada, porque gobierna las unidades de los
    # campos que se teclean mas abajo: un selector colocado debajo de los
    # campos que rotula es un selector que se descubre tarde.
    com15 = ("Entrada: sistema de unidades de LECTURA del codigo. Cambia de que "
    "edicion se lee el esfuerzo admisible -la metrica o la U.S. "
    "Customary-, NUNCA convierte un valor (regla 9: las dos ediciones "
    "son extracciones independientes de lo que cada una imprime). Los "
    "datos que usted teclea -presiones, dimensiones, temperatura- NO se "
    "convierten al cambiarlo: hay que volver a teclearlos en el sistema "
    "nuevo, igual que en los cinco buscadores de cascada.")
    lab(15, "Sistema de unidades", unidad="—", ref="SI (métrico) / US Customary",
        com=com15)
    ws.cell(15, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D15", "SI", com15)
    dv_list(ws, "D15", '"SI,US"', com15)

    # --- 1. Datos de entrada (filas 16-35) -----------------------------------
    # A16: texto NUEVO que reemplaza el heredado ("celdas azules sobre fondo
    # amarillo = editables", ver TEXTOS_HEREDADOS) por el que describe ESTE
    # sistema visual (CAJA_CAMPO = linea inferior, no relleno de color). El
    # *oracle* ya lo trae asi -se capturo con retonar_heredadas() aplicado
    # sobre el pipeline viejo- y se transcribe aqui letra a letra, incluidos
    # los espacios dobles. band() no sirve: aplica rotulo() (mayusculas,
    # colapsa espacios, envuelve en "[ ... ]") y el oracle trae texto mixto
    # con esos espacios deliberados -mismo problema que A1/A2-, asi que se
    # escribe con el mismo estilo (fuente/relleno/franja) que usa band() pero
    # con el VALOR literal, fusionando A16:G16 como declara el *oracle*
    # (fusionados).
    banda_literal(16, "1.  DATOS DE ENTRADA")

    # Encabezado de fila 17: mismo problema que un header en mayusculas
    # tendria en la fila 9 — el *oracle* trae mayusculas y minusculas mixtas
    # ("Parámetro", "Símbolo"...), asi que se escribe directo con el mismo
    # estilo de columna (HDR_F/HDR_FILL/BOX_FRANJA) sin forzar .upper().
    encabezado(17, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                    (4, "Valor"), (7, "Referencia / Notas")))

    # --- Bloque dimensional (Tareas 7-8): norma -> NPS -> cedula -> OD/espesor,
    # todo contra DB_B36_10/DB_B36_19 por cascada de listas (reglas 12/14). NPS,
    # cedula y OD/espesor conservan sus filas 18-21 del oracle Rev0 (no se corren,
    # asi ninguna formula de calculo aguas abajo cambia su referencia); lo que
    # cambia es su CONTENIDO, declarado en DIVERGENCIAS_REEMPLAZADAS. El selector
    # de norma ocupa la fila 22, antes descriptiva de material (ya retirada).
    com18 = ("Entrada: diametro nominal (NPS), del desplegable de DB_B36 segun la "
             "norma elegida en D22. No se teclea (regla 14). Se guarda como el NPS "
             "impreso del codigo, p.ej. '12 (300)'. Alimenta OD, espesor y la lista "
             "de cedulas de abajo.")
    lab(18, "Diámetro nominal", unidad="in", ref="Lista · DB_B36 (segun D22)", com=com18)
    ws.cell(18, 2, "NPS").font = Font(name=MONO, size=10, color=TINTA)
    inp("D18", "12 (300)", com18)

    com19 = ("Entrada: cedula (designador de Schedule) del NPS elegido, en la norma "
             "de D22, del desplegable de DB_B36. No se teclea (regla 14).")
    lab(19, "Cédula (Schedule)", unidad="—", ref="Lista · DB_B36 (segun D22 y NPS)",
        com=com19)
    ws.cell(19, 2, "SCH").font = Font(name=MONO, size=10, color=TINTA)
    inp("D19", "20", com19)

    com20 = ("Calculo: diametro exterior (OD, mm) por lookup de la clave NPS|cedula "
             "(D18|D19) contra DB_B36 de la norma elegida (D22). Vacio si el par no "
             "existe en la base.")
    lab(20, "Diámetro exterior", unidad="mm", ref="DB_B36 (auto, segun D22)", com=com20)
    ws.cell(20, 2, "OD").font = Font(name=MONO, size=10, color=TINTA)

    com21 = ("Calculo: espesor de pared (mm) por lookup de la clave NPS|cedula "
             "(D18|D19) contra DB_B36 de la norma elegida (D22). Vacio si el par no "
             "existe en la base.")
    lab(21, "Espesor de pared", unidad="mm", ref="DB_B36 (auto, segun D22)", com=com21)
    ws.cell(21, 2, "t").font = Font(name=MONO, size=10, color=TINTA)

    # Fila 22: selector de norma dimensional (nivel 0 de la cascada). Antes era el
    # material descriptivo de tuberia (lista fija prohibida por la regla 12), ya
    # retirado; ahora la fila la ocupa la norma. La norma se ELIGE, no se deriva del
    # material (ver _materializar_cascada_b36). La fila 23 (material del parche)
    # sigue retirada y vacia, declarada en DIVERGENCIAS_DECLARADAS.
    com22 = ("Entrada: norma dimensional de la tuberia. B36.10M (acero al carbono y "
             "de baja aleacion) o B36.19M (inoxidable, cedulas de la serie S). "
             "Gobierna las listas de NPS y cedula y el lookup de OD/espesor de arriba.")
    lab(22, "Norma dimensional", unidad="—", ref="ASME B36.10M / B36.19M", com=com22)
    ws.cell(22, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D22", "B36.10M", com22)

    casc = _materializar_cascada_b36(ws, b3610, b3619,
                                     fila_norma=22, fila_nps=18, fila_ced=19)
    idx = casc["idx"]
    clave_key = '$D$18&"|"&$D$19'
    calc("D20", f'=IFERROR(INDEX({_choose_b36(idx, b3610, b3619, "od_mm", ES_SI_212, "od_in")},'
                f'MATCH({clave_key},{_choose_b36(idx, b3610, b3619, "clave")},0)),"")',
         com20)
    calc("D21", f'=IFERROR(INDEX({_choose_b36(idx, b3610, b3619, "t_mm", ES_SI_212, "t_in")},'
                f'MATCH({clave_key},{_choose_b36(idx, b3610, b3619, "clave")},0)),"")',
         com21)

    com24 = ("Entrada informativa: fluido de servicio. No alimenta ningun "
             "calculo de esta hoja.")
    lab(24, "Fluido", unidad="—", com=com24)
    ws.cell(24, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D24", "Agua de enfriamiento", com24)

    com25 = ("Entrada: temperatura de operacion, en °C. Alimenta la "
             "Temperatura de evaluacion de la seccion 2 (D40) y por tanto "
             "todo el S(T) resuelto.")
    lab(25, "Temperatura de operación", unidad="°C", ref="Hoja de proceso", com=com25)
    ws.cell(25, 2, "T").font = Font(name=MONO, size=10, color=TINTA)
    inp("D25", 25, com25)

    com26 = ("Entrada: presion de operacion, en kg/cm². Es el caso "
             "'Operacion' evaluado en la seccion 3.")
    lab(26, "Presión de operación", unidad="kg/cm²", ref="Hoja de proceso", com=com26)
    ws.cell(26, 2, "P_op").font = Font(name=MONO, size=10, color=TINTA)
    inp("D26", 5, com26)

    # Fase 2 (modelo de presion de 2 casos). La fila 27 —"Presion de diseno
    # (tipica)"— SE RETIRA. El motor evaluaba tres presiones: operacion, un
    # "diseno tipico" intermedio y una "envolvente". Ese caso intermedio no lo
    # pide el codigo: 212-3.2 define una UNICA P = "internal design pressure"
    # para las ec. (1)/(2) —comprobado en
    # resources/ASME PCC/pcc_2/p2_welded_repairs/art_212_fillet_welded_patches/
    # art_212.json, Regla n.1— y el Art. 206, que comparte modelo de presion en
    # este libro, es explicito en 206-3.3: el espesor se dimensiona contra "the
    # maximum allowable design pressure". Mantener el caso intermedio ademas de
    # la maxima admisible dejaba dos columnas compitiendo por gobernar el t_req,
    # que es exactamente la ambiguedad que un motor de calculo no debe tener.
    # Las celdas de la fila 27 quedan en DIVERGENCIAS_DECLARADAS.
    com28 = ("Entrada: presion de diseno, en kg/cm² — la MAXIMA ADMISIBLE "
             "(rating), no un valor tipico intermedio. Es la que gobierna el "
             "espesor requerido y el esfuerzo de soldadura: 212-3.2 define una "
             "unica P = 'internal design pressure' para las ec. (1)/(2), y "
             "206-3.3 la nombra 'maximum allowable design pressure'. Es el caso "
             "'Diseno' de la seccion 3 y se compara contra la presion maxima "
             "admisible del parche en la verificacion de la seccion 5 (fila 117).")
    lab(28, "Presión de diseño (máxima admisible / rating)", unidad="kg/cm²",
        ref="Rating / máx. admisible", com=com28)
    ws.cell(28, 2, "P_dis").font = Font(name=MONO, size=10, color=TINTA)
    inp("D28", 20, com28)

    com29 = ("Entrada: espesor adoptado del parche o collar, en mm (debe ser "
             ">= espesor de pared). Alimenta la fuerza de membrana, el "
             "esfuerzo de soldadura y el peso estimado.")
    lab(29, "Espesor del parche/collar", unidad="mm", ref="Adoptado (≥ pared)", com=com29)
    ws.cell(29, 2, "T_c").font = Font(name=MONO, size=10, color=TINTA)
    inp("D29", 8, com29)

    com30 = ("Entrada: altura o dimension del parche, en mm, tomada del "
             "plano. Alimenta el peso estimado (seccion 4).")
    lab(30, "Altura / dimensión del parche", unidad="mm", ref="Del plano", com=com30)
    ws.cell(30, 2, "H").font = Font(name=MONO, size=10, color=TINTA)
    inp("D30", 166, com30)

    com31 = ("Entrada: luz radial entre el parche y el componente, en mm "
             "(<=2 mm tipico). Alimenta el radio de conformado (Rf) y el "
             "desarrollo de la media carcasa.")
    lab(31, "Luz radial parche–componente", unidad="mm", ref="≤ 2 mm", com=com31)
    ws.cell(31, 2, "luz").font = Font(name=MONO, size=10, color=TINTA)
    inp("D31", 1.5, com31)

    com32 = ("Entrada: cateto adoptado del filete perimetral, en mm. Se "
             "compara contra el filete minimo requerido en la verificacion "
             "de la seccion 5 (fila 112).")
    lab(32, "Cateto del filete perimetral", unidad="mm", ref="Adoptado", com=com32)
    ws.cell(32, 2, "w").font = Font(name=MONO, size=10, color=TINTA)
    inp("D32", 6, com32)

    com33 = ("Entrada: solape minimo del parche sobre metal sano, en mm, "
             "exigido por el Art. 212. Solo informativo en esta hoja.")
    lab(33, "Solape mínimo sobre metal sano", unidad="mm", ref="Art. 212", com=com33)
    ws.cell(33, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D33", 25, com33)

    com34 = ("Entrada: diametro del defecto, caracterizado por UT/PT, en mm. "
             "Solo informativo en esta hoja.")
    lab(34, "Diámetro del defecto", unidad="mm", ref="Caracterizar UT/PT", com=com34)
    ws.cell(34, 2, "d_def").font = Font(name=MONO, size=10, color=TINTA)
    inp("D34", 3, com34)

    com35 = ("Entrada: distancia del defecto a la discontinuidad mas cercana "
             "(cordon o boquilla), en mm. Se compara contra L_min en la "
             "verificacion de la seccion 5 (fila 116) para decidir entre "
             "refuerzo 360° y parche local.")
    lab(35, "Distancia defecto–discontinuidad", unidad="mm", ref="A cordón/boquilla", com=com35)
    ws.cell(35, 2, "L_def").font = Font(name=MONO, size=10, color=TINTA)
    inp("D35", 40, com35)

    # --- 2. Parametros de calculo (filas 37-48) ------------------------------
    # A37: banda de seccion, fusionada A37:G37 (confirmado en "fusionados" del
    # oracle), texto literal transcrito tal cual -sin pasar por rotulo(): no
    # lleva numeral "2." ni mayusculas forzadas, mismo criterio que A1/A2 y
    # A16 arriba (la banda/encabezado que precede inmediatamente un rango y
    # titula su seccion se escribe con el mismo criterio en toda la hoja).
    # Encabezado de fila 38: mismo patron que la fila 17 (Parametro/Simbolo/
    # Unidad/Valor/Referencia con mayus/minus mixtas), tampoco compatible con
    # un header en mayusculas.
    banda_literal(37, "3.  PARÁMETROS DE CÁLCULO")

    encabezado(38, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                    (4, "Valor"), (7, "Referencia / Notas")))

    # D39/D40 (y G39/G40) ya los escribio el cableado Seccion 7 -> Seccion
    # 1/3 mas arriba (junto con D44/G44/F90) — no se tocan; aqui solo A/B/C
    # de estas dos filas.
    lab(39, "Esf. admisible del collar", unidad="MPa")
    ws.cell(39, 2, "Sa_c").font = Font(name=MONO, size=10, color=TINTA)
    lab(40, "Esf. admisible del metal base", unidad="MPa")
    ws.cell(40, 2, "Sa_b").font = Font(name=MONO, size=10, color=TINTA)

    com41 = ("Calculo: esfuerzo admisible gobernante, el menor entre el del "
             "collar (D67) y el del metal base (D68).")
    lab(41, "Esf. admisible gobernante", unidad="MPa",
        ref="menor de collar / base", com=com41)
    ws.cell(41, 2, "Sa").font = Font(name=MONO, size=10, color=TINTA)
    calc("D41", "=MIN($D$39,$D$40)", com41)

    com42 = "Entrada: eficiencia de junta de filete del Art. 212, ec. 4."
    lab(42, "Eficiencia de junta de filete", unidad="—",
        ref="Art. 212 ec. 4", com=com42)
    ws.cell(42, 2, "E").font = Font(name=MONO, size=10, color=TINTA)
    inp("D42", 0.55, com42)

    com43 = "Entrada: factor Y de B31.3 Tabla 304.1.1."
    lab(43, "Factor Y (B31.3)", unidad="—",
        ref="B31.3 Tabla 304.1.1", com=com43)
    ws.cell(43, 2, "Y").font = Font(name=MONO, size=10, color=TINTA)
    inp("D43", 0.4, com43)

    # D44/G44 ya los escribio el cableado de la Seccion 7 mas arriba — no se
    # tocan; aqui solo A/B/C de esta fila.
    lab(44, "Eficiencia de junta long. (E)", unidad="—")
    ws.cell(44, 2, "E_j").font = Font(name=MONO, size=10, color=TINTA)

    com45 = ("Entrada: densidad del acero, en kg/m³, para el peso estimado "
             "(seccion 4).")
    lab(45, "Densidad del acero", unidad="kg/m³", com=com45)
    ws.cell(45, 2, "ρ").font = Font(name=MONO, size=10, color=TINTA)
    calc("D45", f'=IF({ES_SI_212},7850,0.2836)', com45 + " " + "Calculo (Fase 7): la constante depende del SISTEMA DE UNIDADES, no del "
        "caso, asi que la fija el selector y no se teclea. Dejarla editable "
        "obligaria a acordarse de cambiarla al conmutar, y olvidarlo daria un "
        "resultado plausible y equivocado.")

    com46 = "Entrada: factor de prueba hidrostatica, B31.3 345.4.2."
    lab(46, "Factor de prueba hidrostática", unidad="—",
        ref="B31.3 345.4.2", com=com46)
    ws.cell(46, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D46", 1.5, com46)

    com47 = ("Entrada: factor de conversion de presion de kg/cm² a MPa "
             "(1 kg/cm² = 0,0980665 MPa).")
    lab(47, "Conversión de presión", unidad="—",
        ref="1 kg/cm²=0,0980665 MPa", com=com47)
    ws.cell(47, 2, "kg/cm²→MPa").font = Font(name=MONO, size=10, color=TINTA)
    calc("D47", f'=IF({ES_SI_212},0.0980665,0.001)', com47 + " " + "Calculo (Fase 7): la constante depende del SISTEMA DE UNIDADES, no del "
        "caso, asi que la fija el selector y no se teclea. Dejarla editable "
        "obligaria a acordarse de cambiarla al conmutar, y olvidarlo daria un "
        "resultado plausible y equivocado.")

    com48 = ("Calculo: limite de esfuerzo para la verificacion de "
             "excentricidad (seccion 5), 1,5 veces el esfuerzo admisible "
             "gobernante Sa (Art. 212 ec. 5).")
    lab(48, "Límite de esfuerzo (excentricidad)", unidad="MPa",
        ref="Art. 212 ec. 5", com=com48)
    ws.cell(48, 2, "1,5·Sa").font = Font(name=MONO, size=10, color=TINTA)
    calc("D48", "=1.5*$D$41", com48)

    # --- 2. Geometria y propiedades derivadas (filas 52-57) ------------------
    # A50: banda de seccion, fusionada A50:G50 (confirmado en "fusionados" del
    # oracle), texto literal transcrito tal cual -sin pasar por rotulo(), con
    # el numeral "2." tal como lo imprime el oracle (aunque la banda de la
    # fila 37, "PARAMETROS DE CALCULO...", no lleve numeral: es lo que el
    # maestro trae impreso, Regla n.1) y el doble espacio entre "2." y
    # "GEOMETRIA" preservado letra a letra, mismo criterio que A1/A2, A16 y
    # A37 (la banda y el encabezado que preceden inmediatamente un rango y
    # titulan su seccion se escriben con el mismo criterio en toda la hoja).
    # Encabezado de fila 51: mismo patron que las filas 17/38 (mayus/minus
    # mixtas, no compatible con un header en mayusculas), pero con G51 =
    # "Formula / Referencia" en vez de "Referencia / Notas" -asi lo trae el
    # oracle para esta fila especifica, no se asume igual al resto-.
    banda_literal(50, "4.  GEOMETRÍA Y PROPIEDADES DERIVADAS")

    encabezado(51, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                    (4, "Valor"), (7, "Fórmula / Referencia")))

    # Los comentarios (com) de D52-D57 repiten el texto que ya trae el dict
    # "simples" de comentar_art212_base para estas mismas filas (52-57): no
    # se inventa contenido nuevo, se reusa el que documenta la formula tal
    # como esta en el maestro. comentar_art212_base() se llama una sola vez,
    # al final de la funcion, y sobreescribe estas mismas notas con el mismo
    # texto (idempotente); escribirlas aqui tambien deja la seccion legible
    # por si sola en el codigo, con el mismo criterio que las secciones
    # anteriores (com18, com22...).
    com52 = "Calculo: diametro a media pared, Dm = OD − t."
    lab(52, "Diámetro a media pared", unidad="mm", ref="Dm = OD − t", com=com52)
    ws.cell(52, 2, "Dm").font = Font(name=MONO, size=10, color=TINTA)
    calc("D52", "=$D$20-$D$21", com52)

    com53 = "Calculo: radio medio, Rm = Dm/2."
    lab(53, "Radio medio", unidad="mm", ref="Rm = Dm/2", com=com53)
    ws.cell(53, 2, "Rm").font = Font(name=MONO, size=10, color=TINTA)
    calc("D53", "=$D$52/2", com53)

    com54 = "Calculo: radio interior, Ri = OD/2 − t."
    lab(54, "Radio interior", unidad="mm", ref="Ri = OD/2 − t", com=com54)
    ws.cell(54, 2, "R_i").font = Font(name=MONO, size=10, color=TINTA)
    calc("D54", "=$D$20/2-$D$21", com54)

    # Fase 5: e incluye la separacion g del faying edge cuando g >= 1.5 mm
    # (212-4c, bloque [82]). g es una entrada nueva del Paso 5 (D167), distinta
    # de la luz radial de conformado D31. Con g < 1.5 (o g = 0, fit-up ajustado
    # del caso semilla) e = (T+t)/2 como antes. El rotulo/simbolo/unidad/ref
    # siguen anclados al oracle; solo cambia la formula.
    com55 = ("Calculo: excentricidad de la carga, e = (T_parche + t_pared + g)/2, "
             "donde g es la separacion en el borde (faying edge, D167) y solo se "
             "suma si g >= 1.5 mm (212-4c). Con fit-up ajustado (g=0) e = (T+t)/2.")
    lab(55, "Excentricidad de la carga", unidad="mm", ref="e = (T + t)/2", com=com55)
    ws.cell(55, 2, "e").font = Font(name=MONO, size=10, color=TINTA)
    calc("D55", f'=($D$29+$D$21+IF($D$167>={umbral(UMBR, "212_separacion_g", ES_SI_212)},'
         f'$D$167,0))/2', com55)

    com56 = ("Calculo: radio de conformado de la fibra media, Rf = OD/2 + luz "
             "+ T_parche/2; se usa en la deformacion por conformado (ec. 7).")
    lab(56, "Radio de conformado (fibra media)", unidad="mm",
        ref="Rf = OD/2 + luz + T/2", com=com56)
    ws.cell(56, 2, "Rf").font = Font(name=MONO, size=10, color=TINTA)
    calc("D56", "=$D$20/2+$D$31+$D$29/2", com56)

    # Fase 5: C_sw en forma literal de cilindro (ec. 5), sin el factor kf. Para
    # cilindro coincide con el valor anterior (kf=0.5); para esfera/cabezal
    # (D11=3) la ec.(5) no aplica -> NA (decision 3). Alimenta P_max (D74).
    com57 = ("Calculo: coeficiente C_sw tal que S_w = P·C_sw, en la forma literal "
             "de la ec. (5) para cilindro: C_sw = Dm/(2T)·(1+6e/T). En esfera/"
             "cabezal (D11=3) la ec. (5) no aplica y C_sw = NA (hand-off 212-3.4).")
    lab(57, "Coef. de esfuerzo por presión", unidad="MPa/MPa",
        ref="S_w = P·C_sw", com=com57)
    ws.cell(57, 2, "C_sw").font = Font(name=MONO, size=10, color=TINTA)
    calc("D57", "=IF($D$11=3,NA(),$D$52/(2*$D$29)*(1+6*$D$55/$D$29))", com57)

    # --- 3. Calculo de cargas y soldadura (filas 59-69) ----------------------
    # A59: banda de seccion, fusionada A59:G59 (confirmado en "fusionados" del
    # *oracle*), texto literal transcrito tal cual -sin pasar por rotulo(),
    # con el numeral "3." y el doble espacio antes de "CALCULO"- mismo
    # criterio que las bandas A16/A37/A50 anteriores (la banda y el
    # encabezado que preceden inmediatamente un rango y titulan su seccion
    # se escriben con el mismo criterio en toda la hoja). Encabezado de fila
    # 60: DOS columnas de valor propias (D/E = Operacion/Diseno, no una sola
    # "Valor" como en las filas 17/38/51), tampoco compatible con un header
    # generico -en mayusculas y con una sola columna de valor- asi que se
    # escribe directo con el mismo estilo (HDR_F/HDR_FILL/BOX_FRANJA).
    # Fase 2: eran TRES (Operacion/Diseno tipico/Envolvente). La columna F se
    # retira entera y E pasa a leer la presion de diseno maxima admisible
    # (D28): ver el comentario de la fila 27/28 en la Seccion 1.
    #
    # La fila 58 queda vacia (no aparece en el *oracle*: ni formula, ni
    # fusionado, ni validacion): separa la Seccion 2 (termina en fila 57) de
    # la banda de esta seccion.
    banda_literal(59, "5.  CÁLCULO DE CARGAS Y SOLDADURA  —  ASME PCC-2 Art. 212")

    encabezado(60, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                    (4, "Operación"), (5, "Diseño"), (7, "Referencia")))

    # Filas 61-69: los DOS casos de presion (Operacion/Diseno) en columnas D/E.
    # Cada fila repite el mismo comentario en las dos columnas de calculo y en
    # el rotulo de columna A (texto tomado de comentar_art212_base, sin inventar
    # contenido nuevo — mismo criterio que en com52-57; comentar_art212_base()
    # sobreescribe estas mismas notas de forma idempotente al llamarse al final
    # de la funcion).
    com61 = ("Calculo: repite, para este caso (Operacion / Diseno), la presion "
             "correspondiente de la seccion 1, en kg/cm². 'Diseno' es la maxima "
             "admisible (D28): 212-3.2 define una unica P de diseno.")
    lab(61, "Presión evaluada", unidad="kg/cm²", ref="Entradas §1", com=com61)
    ws.cell(61, 2, "P").font = Font(name=MONO, size=10, color=TINTA)
    calc("D61", "=$D$26", com61)
    calc("E61", "=$D$28", com61)

    com62 = ("Calculo: conversion a MPa de la presion de este caso, "
             "multiplicando por el factor de conversion D47.")
    lab(62, "Presión evaluada", unidad="MPa", ref="P·0,0980665", com=com62)
    ws.cell(62, 2, "P").font = Font(name=MONO, size=10, color=TINTA)
    calc("D62", "=D61*$D$47", com62)
    calc("E62", "=E61*$D$47", com62)

    com63 = ("Calculo: fuerza de membrana de este caso (ec. 1 del Art. 212): "
             "kf · P(MPa) · Dm.")
    lab(63, "Fuerza de membrana gobernante", unidad="N/mm",
        ref="ec.1: kf·P·Dm", com=com63)
    ws.cell(63, 2, "F_m").font = Font(name=MONO, size=10, color=TINTA)
    calc("D63", "=$D$14*D62*$D$52", com63)
    calc("E63", "=$D$14*E62*$D$52", com63)

    # Fase 4: w_min usa la fuerza gobernante F_max del Paso 2 (D150/E150/F150),
    # no F_m (D63). Para cilindro sin cargas externas F_max = F_CP = F_m, asi
    # que el caso semilla no cambia de valor; para esfera/cabezal F_max = NA()
    # (hand-off 212-3.2c), lo que bloquea w_min como debe (decision 3). El
    # rotulo A64, el simbolo B64, la unidad C64 y la referencia G64 siguen
    # anclados al oracle: solo cambia la formula (ec. 4, bloques [57],[59]).
    com64 = ("Calculo: cateto de filete minimo requerido para este caso (ec. 4, "
             "212-3.4): F_max / (E · Sa), con F_max del Paso 2 (fuerza "
             "gobernante) y E = 0.55. En esfera/cabezal F_max = NA() -> w_min "
             "no aplica (hand-off 212-3.2c).")
    lab(64, "Filete requerido", unidad="mm", ref="ec.4: F/(E·Sa)", com=com64)
    ws.cell(64, 2, "w_mín").font = Font(name=MONO, size=10, color=TINTA)
    calc("D64", "=D150/($D$42*$D$41)", com64)
    calc("E64", "=E150/($D$42*$D$41)", com64)

    com65 = ("Calculo: espesor de pared requerido para este caso, por B31.3 "
             "(modo tuberia) o por VIII-1 UG-27 (modo esfera/cilindro), "
             "segun la presion evaluada de este caso.")
    lab(65, "Espesor de pared requerido", unidad="mm",
        ref="B31.3 / VIII-1 UG-27", com=com65)
    ws.cell(65, 2, "t_req").font = Font(name=MONO, size=10, color=TINTA)
    calc("D65",
         '=IF($D$11=1,D62*$D$20/(2*($D$40*$D$44+D62*$D$43)),'
         'IF($D$11=3,D62*$D$54/(2*$D$40*$D$44-0.2*D62),'
         'D62*$D$54/($D$40*$D$44-0.6*D62)))', com65)
    calc("E65",
         '=IF($D$11=1,E62*$D$20/(2*($D$40*$D$44+E62*$D$43)),'
         'IF($D$11=3,E62*$D$54/(2*$D$40*$D$44-0.2*E62),'
         'E62*$D$54/($D$40*$D$44-0.6*E62)))', com65)

    # Fase 5: S_w se escribe LITERAL de la ec. (5) del 212-3.4c, con P*Dm
    # directamente en vez de via F_m/kf (decision 3). Para cilindro el valor no
    # cambia (F_m con kf=0.5 ya daba P*Dm/2 y 3*P*Dm*e/T^2); para esfera/cabezal
    # (D11=3) la ec.(5) no aplica -> NA, lo que bloquea S_w y su verificacion.
    # Los rotulos/simbolos/unidades/refs de las filas 66/67 siguen anclados al
    # oracle; solo cambian las formulas. D68 = D66+D67 no cambia (propaga NA).
    com66 = ("Calculo: componente de membrana de la ec. (5) del 212-3.4c: "
             "P(MPa)·Dm/(2T). En esfera/cabezal (D11=3) la ec. no aplica -> NA.")
    lab(66, "Esfuerzo soldadura — membrana", unidad="MPa",
        ref="ec.5 (memb.)=F/T", com=com66)
    ws.cell(66, 2, "S_w,m").font = Font(name=MONO, size=10, color=TINTA)
    calc("D66", "=IF($D$11=3,NA(),D62*$D$52/(2*$D$29))", com66)
    calc("E66", "=IF($D$11=3,NA(),E62*$D$52/(2*$D$29))", com66)

    com67 = ("Calculo: componente de flexion de la ec. (5) del 212-3.4c: "
             "3·P(MPa)·Dm·e/T². En esfera/cabezal (D11=3) -> NA.")
    lab(67, "Esfuerzo soldadura — flexión", unidad="MPa",
        ref="ec.5 (flex.)=6F·e/T²", com=com67)
    ws.cell(67, 2, "S_w,f").font = Font(name=MONO, size=10, color=TINTA)
    calc("D67", "=IF($D$11=3,NA(),3*D62*$D$52*$D$55/$D$29^2)", com67)
    calc("E67", "=IF($D$11=3,NA(),3*E62*$D$52*$D$55/$D$29^2)", com67)

    com68 = ("Calculo: esfuerzo de soldadura total de este caso (membrana + "
             "flexion); se compara contra el limite 1,5·Sa en la fila 97.")
    lab(68, "Esfuerzo soldadura — total", unidad="MPa", ref="ec.5: ≤ 1,5·Sa",
        com=com68)
    ws.cell(68, 2, "S_w").font = Font(name=MONO, size=10, color=TINTA)
    calc("D68", "=D66+D67", com68)
    calc("E68", "=E66+E67", com68)

    # Fila 69 no lleva simbolo propio: el *oracle* no declara B69 (a
    # diferencia de las filas 61-68, que si lo tienen) — se deja vacio.
    com69 = ("Calculo: CUMPLE si el esfuerzo de soldadura total de este "
             "caso (fila 96) no supera 1,5 veces el esfuerzo admisible "
             "gobernante (D76).")
    lab(69, "¿S_w ≤ 1,5·Sa?", unidad="—", ref="Verificación por presión",
        com=com69)
    calc("D69", '=IF(D68<=$D$48,"CUMPLE","NO CUMPLE")', com69)
    calc("E69", '=IF(E68<=$D$48,"CUMPLE","NO CUMPLE")', com69)

    # --- 4. Resultados del diseno (filas 71-80) ------------------------------
    # A71: banda de seccion, fusionada A71:G71 (confirmado en "fusionados"
    # del *oracle*), texto literal transcrito tal cual -sin pasar por
    # rotulo(), sin numeral extra ni doble espacio esta vez: es lo que el
    # *oracle* imprime- mismo criterio que las bandas anteriores. Encabezado
    # de fila 72: cinco columnas (Parametro/Simbolo/Unidad/Valor/"Fórmula /
    # Referencia" en G72, mismo patron que la fila 51 — no "Referencia /
    # Notas" de las filas 17/38), tampoco compatible con un header en
    # mayusculas.
    #
    # La fila 70 queda vacia (no aparece en el *oracle*): separa la Seccion 3
    # (termina en fila 69) de la banda de esta seccion, mismo patron
    # espaciador que la fila 58 delante de la Seccion 3.
    banda_literal(71, "6.  RESULTADOS DEL DISEÑO")

    encabezado(72, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                    (4, "Valor"), (7, "Fórmula / Referencia")))

    com73 = ("Calculo: distancia minima a una discontinuidad, L_min = "
             "2·RAIZ(Rm·t) (ec. 3). Por debajo de esta distancia el defecto "
             "exige refuerzo de 360° en vez de parche local (verificacion "
             "de la seccion 5, fila 116).")
    lab(73, "Distancia a la discontinuidad", unidad="mm",
        ref="ec.3: 2·√(Rm·t)", com=com73)
    ws.cell(73, 2, "L_mín").font = Font(name=MONO, size=10, color=TINTA)
    calc("D73", "=2*SQRT($D$53*$D$21)", com73)

    com74 = ("Calculo: presion maxima admisible del parche, en MPa: "
             "1,5·Sa / C_sw.")
    lab(74, "Presión máx. admisible del parche", unidad="MPa",
        ref="1,5·Sa / C_sw", com=com74)
    ws.cell(74, 2, "P_máx").font = Font(name=MONO, size=10, color=TINTA)
    calc("D74", "=$D$48/$D$57", com74)

    # G75 no lo declara el *oracle*: no se repite la referencia de D74 ahi
    # (Regla dura, tasks-3-8-common.md — se deja vacio, no se rellena).
    com75 = "Calculo: igual que D74, convertido a kg/cm² con el factor D47."
    lab(75, "Presión máx. admisible del parche", unidad="kg/cm²", com=com75)
    ws.cell(75, 2, "P_máx").font = Font(name=MONO, size=10, color=TINTA)
    calc("D75", "=$D$74/$D$47", com75)

    # B76 no lo declara el *oracle*: esta fila no tiene simbolo propio.
    com76 = ("Calculo: relacion entre la presion maxima admisible del "
             "parche y la presion de operacion (D26); entre mas alto, "
             "mayor margen.")
    lab(76, "Margen sobre la operación", unidad="×", ref="P_máx / P_op",
        com=com76)
    calc("D76", "=$D$75/$D$26", com76)

    # B77 no lo declara el *oracle*.
    # Fase 6: deformacion por conformado completa: coef*T/Rf*(1-Rf/Ro), con
    # coef = 75 para curvatura doble (esfera/cabezal, ec. 6 [71]) o 50 para
    # curvatura simple (tuberia/virola, ec. 7 [75]), y el factor (1-Rf/Ro) con
    # Ro = radio original (D172; en blanco = plano, Ro=inf -> factor 1) [73].
    # El caso semilla (cilindro, Ro plano) da 50*T/Rf como antes. El rotulo A77,
    # la unidad C77 y la ref G77 siguen anclados al oracle; solo cambia D77.
    com77 = ("Calculo: deformacion por conformado en frio (%): coef·T/Rf·(1-Rf/Ro), "
             "coef=75 doble curvatura (esfera/cabezal, ec.6) o 50 simple (tuberia/"
             "virola, ec.7); Ro=D172 (blanco = plano, factor 1). <=5% para no "
             "requerir PWHT post-conformado (212-3.5b).")
    lab(77, "Deformación por conformado", unidad="%",
        ref="ec.7: 50·T/Rf ≤ 5%", com=com77)
    calc("D77", '=IF($D$11=3,75,50)*$D$29/$D$56*IF($D$172="",1,1-$D$56/$D$172)',
         com77)

    com78 = ("Calculo: desarrollo de la media carcasa del collar, "
             "(π/2)·(OD + 2·luz + T_parche).")
    lab(78, "Desarrollo por media carcasa", unidad="mm",
        ref="(π/2)·(OD+2·luz+e)", com=com78)
    ws.cell(78, 2, "L").font = Font(name=MONO, size=10, color=TINTA)
    calc("D78", "=PI()/2*($D$20+2*$D$31+$D$29)", com78)

    # B79 no lo declara el *oracle*.
    com79 = ("Calculo: longitud real de corte de la plancha, igual al "
             "desarrollo menos 3 mm de luz de raiz.")
    lab(79, "Longitud de corte (media carcasa)", unidad="mm",
        ref="menos luz de raíz", com=com79)
    # Los 3 mm de luz de raiz NO son del codigo: son holgura de fabricacion.
    # Aun asi son una LONGITUD, y sin conmutar restaban 3 pulgadas en modo US
    # (se vio en el recalculo: el peso salia un 48 % corto).
    calc("D79", f'=$D$78-IF({ES_SI_212},3,{3 / 25.4:.4f})', com79)

    # B80 no lo declara el *oracle*.
    com80 = ("Calculo: peso estimado del parche/collar (dos mitades), a "
             "partir de la longitud de corte, la altura, el espesor y la "
             "densidad del acero.")
    lab(80, "Peso del parche/collar", unidad="kg", ref="2·L·H·T·ρ", com=com80)
    calc("D80", f'=2*$D$79*$D$30*$D$29*$D$45/IF({ES_SI_212},1000000000,1)', com80)

    # --- 5. Verificaciones (filas 82-90) --------------------------------
    # A82: banda de seccion, fusionada A82:G82 (confirmado en "fusionados"
    # del *oracle*), texto literal transcrito tal cual -sin pasar por
    # rotulo(), con el numeral "5." y el doble espacio antes del parentesis-
    # mismo criterio que las bandas anteriores. Encabezado de fila 83: CINCO
    # columnas propias (Verificacion/Requerido/Adoptado/Resultado/Criterio
    # en A/D/E/F/G, sin Simbolo ni Unidad en B/C — esta tabla no las tiene),
    # tampoco compatible con un header generico -en mayusculas y que escribe
    # columnas consecutivas desde la 1, sin poder saltar B/C-.
    banda_literal(82, "7.  VERIFICACIONES")

    encabezado(83, ((1, "Verificación"), (4, "Requerido"), (5, "Adoptado"),
                    (6, "Resultado"), (7, "Criterio")))

    # Filas 84-89: el rotulo de cada verificacion ocupa A:C fusionado (el
    # *oracle* no reparte Simbolo/Unidad en esta tabla); D/E/F llevan las
    # tres columnas Requerido/Adoptado/Resultado y G el criterio de
    # aceptacion impreso (se transcribe con lab(..., ref=...) para reusar el
    # mismo estilo GRIS9 que el resto de la hoja usa en columna G). Los
    # comentarios de D/E/F de estas seis filas los aplica
    # comentar_art212_base() al final: trae TRES textos distintos por fila
    # (Requerido/Adoptado/Resultado), no uno solo compartido como en las
    # secciones anteriores, asi que no hay un "com" unico que pasarle aqui a
    # lab()/calc() sin inventar contenido nuevo.
    # Fase 2: el MAX barre ahora D64:E64 (dos casos, no tres). Con la columna F
    # retirada, "=MAX(D64:F64)" seguiria dando el mismo numero -F64 esta vacia-
    # pero dejaria la hoja declarando un rango que ya no existe.
    lab(84, "Filete perimetral (cateto)", ref="w ≥ w_mín (diseño)")
    ws.merge_cells("A84:C84")
    calc("D84", "=MAX(D64:E64)")
    calc("E84", "=$D$32")
    calc("F84", '=IF(E84>=D84,"CUMPLE","NO CUMPLE")')

    lab(85, "Excentricidad de la soldadura (diseño)", ref="S_w(diseño) ≤ 1,5·Sa")
    ws.merge_cells("A85:C85")
    calc("D85", "=$D$48")
    calc("E85", "=E68")
    calc("F85", '=IF(E85<=D85,"CUMPLE","NO CUMPLE")')

    lab(86, "Conformado en frío", ref="≤ 5 % (ec. 7)")
    ws.merge_cells("A86:C86")
    calc("D86", 5)
    calc("E86", "=$D$77")
    calc("F86", '=IF(E86<=D86,"CUMPLE","NO CUMPLE")')

    # Fase 2: el t_req gobernante es el del caso de DISENO (maxima admisible),
    # que es lo que exige 206-3.3 y lo que 212-3.2 entiende por P. Antes leia la
    # columna Envolvente (F65), que era ese mismo concepto con otro nombre.
    lab(87, "Espesor de pared (diseño)", ref="t ≥ t_req")
    ws.merge_cells("A87:C87")
    calc("D87", "=E65")
    calc("E87", "=$D$21")
    calc("F87", '=IF(E87>=D87,"CUMPLE","NO CUMPLE")')

    lab(88, "Ubicación del defecto", ref="L_def vs L_mín")
    ws.merge_cells("A88:C88")
    calc("D88", "=$D$73")
    calc("E88", "=$D$35")
    calc("F88", '=IF(E88<D88,"Refuerzo 360°","Parche local")')

    lab(89, "Presión de diseño vs. parche", ref="P_dis ≤ P_máx del parche")
    ws.merge_cells("A89:C89")
    calc("D89", "=$D$75")
    calc("E89", "=$D$28")   # Fase 2: la presion de diseno vive ahora en D28
    calc("F89", '=IF(E89<=D89,"OK — parche","Migrar (Art.206)")')

    # Fila 90: DICTAMEN GLOBAL, fusionado A90:E90. F90 ya lo escribio el
    # cableado Seccion 7 mas arriba en esta funcion — no se reescribe aqui.
    # Mismo estilo que el dictamen global del Art. 206 (build_collar_art206:
    # lab() + override a Font MACRO/TINTA — emblematico pero sobre papel, no
    # sobre banda).
    lab(90, "DICTAMEN GLOBAL DEL DISEÑO")
    ws.cell(90, 1).font = Font(name=MACRO, size=11, color=TINTA)
    ws.merge_cells("A90:E90")

    # --- Especificaciones tecnicas: FUERA de esta hoja (Fase 9) --------------
    # La seccion vivia aqui, en las filas 93-99 (121-128 tras el remapeo), con
    # siete filas de parrafo largo fusionadas B:G. Era el bloque mas denso de la
    # hoja y el unico que no se lee mientras se calcula: se consulta cuando el
    # calculo ya esta cerrado y hay que redactar el procedimiento. Pasa a
    # Espec_PCC2_Art212, donde ademas cabe lo que aqui no cabia — la cita del
    # parrafo concreto de PCC-2 que sostiene cada especificacion.
    #
    # De la banda se conserva la fila (A121, fusionada A121:G121, ya declarada en
    # DIVERGENCIAS_REEMPLAZADAS desde la Fase 5) reescrita como LETRERO: quien
    # busque la seccion 8 tiene que encontrar adonde se fue. Pierde el numeral,
    # porque ya no es una seccion de esta hoja. Las siete filas de contenido
    # quedan VACIAS y declaradas en DIVERGENCIAS_DECLARADAS, sus fusionados
    # incluidos.
    banda_literal(92, "ESPECIFICACIONES TÉCNICAS  //  EN SU PROPIA PESTAÑA — "
                      "BOTÓN ARRIBA A LA DERECHA")

    # --- Aviso fijo (fila 101) -----------------------------------------------
    # Mismo estilo que los avisos fijos del Art. 206 (nota206/nota206b, mas
    # arriba en este archivo): fuente SRC_F (la misma del subtitulo A2), sin
    # banda ni fondo especial. Fusionado A101:G101 (confirmado en
    # "fusionados" del *oracle*).
    ws.merge_cells("A101:G101")
    aviso101 = ws.cell(
        101, 1,
        "Herramienta de ingeniería de referencia. Verificar entradas y resultados; complementar con WPS/PQR, ATS/JSA y registros del propietario. Cálculos según ASME PCC-2-2022 (Art. 212/206), ASME B31.3 y ASME BPVC VIII-1.")
    aviso101.font = SRC_F

    # === ANEXO — PASOS DEL FLUJO 212 (bloques nuevos, direcciones estables) ==
    # Los pasos del flujo aprobado que no cabian en las secciones 1-6 sin
    # desplazar cientos de referencias absolutas del oracle Rev0 se anaden aqui,
    # a partir de la fila 134 (la Seccion 7 termina en la 131). Cada bloque se
    # cablea a las celdas existentes (F90, verificaciones) por referencia; su
    # posicion fisica no importa para el calculo. Trazado a resources/ (Regla
    # n.1) y al flujo del ingeniero, fijado por cadena en test_dashboard.py.

    # --- PASO 1 — Elegibilidad y caracterizacion del dano (212-1 / 212-2) ----
    # Fuente: Art. 212 bloque [6] (T hasta 345 C; < nil-ductility -> tenacidad;
    # > 345 -> creep/fatiga), [11]-[13] (dano caracterizable; grietas solo si
    # arrestada + analisis FFS). El servicio letal -> Art. 201 lo fija el flujo
    # aprobado del ingeniero (212-2a remite a la Part 1 del estandar). El
    # dictamen D140 alimenta F90: un estado bloqueante detiene el diseno.
    banda_literal(134, "PASO 1 · ELEGIBILIDAD Y CARACTERIZACIÓN DEL DAÑO  "
                       "—  ASME PCC-2 Art. 212-1 / 212-2")
    encabezado(135, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                     (4, "Valor"), (7, "Referencia / Notas")))

    com136 = ("Entrada: mecanismo de dano caracterizado. El Art. 212 se aplica "
              "a adelgazamiento local (erosion, corrosion, perforacion "
              "traspasante) y danos locales similares (212-1c); no se usa si el "
              "dano no se puede caracterizar (212-2c).")
    lab(136, "Mecanismo de daño", unidad="—", ref="212-1(c) / 212-2(c)", com=com136)
    ws.cell(136, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D136", "Adelgazamiento local", com136)
    dv_list(ws, "D136",
            '"Adelgazamiento local,Erosion,Corrosion,Perforacion traspasante,'
            'Otro/no caracterizado"', com136)

    com137 = ("Entrada: ¿la tasa de dano (presente y futura) se conoce o se "
              "predice? El Art. 212-2(c) prohibe el metodo si el mecanismo, la "
              "extension o el dano futuro NO se pueden caracterizar.")
    lab(137, "¿Daño caracterizable (tasa conocida)?", unidad="—",
        ref="212-2(c) [11]", com=com137)
    ws.cell(137, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D137", "Si", com137)
    dv_list(ws, "D137", '"Si,No"', com137)

    com138 = ("Entrada: tipo de servicio. El servicio letal o de extrema "
              "peligrosidad esta PROHIBIDO para este metodo (flujo aprobado; "
              "212-2a remite a la Part 1 del estandar): usar Art. 201 (inserto "
              "a tope) o reemplazo de seccion.")
    lab(138, "Tipo de servicio", unidad="—", ref="Flujo · Part 1 (212-2a)",
        com=com138)
    ws.cell(138, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D138", "General", com138)
    dv_list(ws, "D138", '"General,Letal / extrema peligrosidad"', com138)

    com139 = ("Entrada: presencia de grietas o defectos tipo grieta. El Art. "
              "212-2(c) solo admite grietas si el crecimiento se ha detenido / "
              "arrestado / es predecible Y el efecto se evalua por analisis "
              "detallado (FFS, API 579); una grieta activa o no analizada lo "
              "hace NO ELEGIBLE.")
    lab(139, "Presencia de grietas", unidad="—", ref="212-2(c) [11]-[13]",
        com=com139)
    ws.cell(139, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D139", "No", com139)
    dv_list(ws, "D139",
            '"No,Si — arrestada + FFS,Si — activa/no analizada"', com139)

    com140 = ("Calculo: DICTAMEN DE ELEGIBILIDAD del Paso 1, en este orden: "
              "servicio letal -> PROHIBIDO (Art. 201); dano no caracterizable "
              "-> NO ELEGIBLE (212-2c); grieta activa/no analizada -> NO "
              "ELEGIBLE (212-2c); T de operacion > 345 -> FUERA DE ALCANCE "
              "(evaluar creep/fatiga, 212-1e); T < 0 -> REVISAR tenacidad a la "
              "entalla (cribado del nil-ductility, 212-1e; el limite real es la "
              "temperatura de nil-ductility del material); en otro caso "
              "ELEGIBLE. Los tres primeros y el de T>345 BLOQUEAN el dictamen "
              "global (F118); el de entalla solo avisa.")
    lab(140, "Dictamen de elegibilidad", unidad="—", ref="212-1/2 + flujo",
        com=com140)
    ws.cell(140, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    calc("D140",
         '=IF($D$138="Letal / extrema peligrosidad",'
         '"PROHIBIDO — servicio letal: usar Art. 201 o reemplazo de seccion '
         '(flujo 212 · 212-2a remite a Part 1)",'
         'IF($D$137="No","NO ELEGIBLE — dano no caracterizable (212-2c)",'
         'IF($D$139="Si — activa/no analizada",'
         '"NO ELEGIBLE — grieta activa o no analizada (212-2c)",'
         'IF($D$25>345,"FUERA DE ALCANCE — T > 345 (evaluar creep/fatiga 212-1e)",'
         'IF($D$25<0,"REVISAR — T < 0 (evaluar tenacidad a la entalla 212-1e)",'
         '"ELEGIBLE")))))', com140)

    # --- PASO 2 — Cargas de presion y externas combinadas (212-3.2) ----------
    # Fuente: ec.(1) F_CP = P*Dm/2 [bloque 29]; ec.(2) F_LP = P*Dm/4 [bloque 33,
    # fijada en la capa de texto en la Fase 0.1]; F_C = F_CP + F_CO [37];
    # F_L = F_LP + F_LO [39]; el filete usa F_A > F_C y F_L [59]; 212-3.1(a) [16]
    # obliga a evaluar flexion/torsion/viento/fatiga -> entran como F_CO/F_LO.
    # Para esfera/cabezal (D11=3), 212-3.2(c) [44] pide un calculo alternativo:
    # no se aplican (1)/(2), se declara el hand-off y F_max = NA() (decision 3
    # del ingeniero: literal, solo cilindro). Bloque nuevo: no toca el oracle.
    # F_max alimenta el filete en el Paso 4 (Fase 4 cablea w_min a F_max).
    banda_literal(142, "PASO 2 · CARGAS DE PRESIÓN Y EXTERNAS COMBINADAS  "
                       "—  ASME PCC-2 Art. 212-3.2")
    encabezado(143, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                     (4, "Operación"), (5, "Diseño"), (7, "Referencia")))

    com144 = ("Entrada: fuerza unitaria circunferencial por OTRAS cargas "
              "(flexion, torsion, viento, sismo), N/mm. 212-3.1(a) obliga a "
              "evaluarlas; 0 es un valor valido tecleado. Se suma a F_CP en los "
              "tres casos de presion.")
    lab(144, "Carga externa circunferencial", unidad="N/mm",
        ref="212-3.1(a) / 212-3.2(b)", com=com144)
    ws.cell(144, 2, "F_CO").font = Font(name=MONO, size=10, color=TINTA)
    inp("D144", 0, com144)

    com145 = ("Entrada: fuerza unitaria longitudinal por OTRAS cargas, N/mm "
              "(212-3.1a). 0 es un valor valido. Se suma a F_LP en los tres casos.")
    lab(145, "Carga externa longitudinal", unidad="N/mm",
        ref="212-3.1(a) / 212-3.2(b)", com=com145)
    ws.cell(145, 2, "F_LO").font = Font(name=MONO, size=10, color=TINTA)
    inp("D145", 0, com145)

    com146 = ("Calculo: fuerza unitaria circunferencial por presion (ec. 1, "
              "212-3.2a): F_CP = P(MPa)·Dm/2. Para cilindro; en esfera/cabezal "
              "(D11=3) no aplica -> ver F_max.")
    lab(146, "Fuerza circunf. por presión", unidad="N/mm", ref="ec.1: P·Dm/2",
        com=com146)
    ws.cell(146, 2, "F_CP").font = Font(name=MONO, size=10, color=TINTA)
    calc("D146", "=D62*$D$52/2", com146)
    calc("E146", "=E62*$D$52/2", com146)

    com147 = ("Calculo: fuerza unitaria longitudinal por presion (ec. 2, "
              "212-3.2a): F_LP = P(MPa)·Dm/4.")
    lab(147, "Fuerza longit. por presión", unidad="N/mm", ref="ec.2: P·Dm/4",
        com=com147)
    ws.cell(147, 2, "F_LP").font = Font(name=MONO, size=10, color=TINTA)
    calc("D147", "=D62*$D$52/4", com147)
    calc("E147", "=E62*$D$52/4", com147)

    com148 = ("Calculo: fuerza circunferencial total (ec. 212-3.2b): "
              "F_C = F_CP + F_CO.")
    lab(148, "Fuerza circunf. total", unidad="N/mm", ref="F_C = F_CP + F_CO",
        com=com148)
    ws.cell(148, 2, "F_C").font = Font(name=MONO, size=10, color=TINTA)
    calc("D148", "=D146+$D$144", com148)
    calc("E148", "=E146+$D$144", com148)

    com149 = ("Calculo: fuerza longitudinal total (ec. 212-3.2b): "
              "F_L = F_LP + F_LO.")
    lab(149, "Fuerza longit. total", unidad="N/mm", ref="F_L = F_LP + F_LO",
        com=com149)
    ws.cell(149, 2, "F_L").font = Font(name=MONO, size=10, color=TINTA)
    calc("D149", "=D147+$D$145", com149)
    calc("E149", "=E147+$D$145", com149)

    com150 = ("Calculo: fuerza gobernante del filete, F_max = MAX(F_C, F_L). El "
              "filete se dimensiona para que F_A la supere (ec. 4, 212-3.4). En "
              "esfera/cabezal (D11=3) las ec. (1)/(2) no aplican: F_max = NA() y "
              "se declara el hand-off 212-3.2(c) (calculo alternativo / analisis).")
    lab(150, "Fuerza gobernante del filete", unidad="N/mm",
        ref="F_max = MAX(F_C, F_L)", com=com150)
    ws.cell(150, 2, "F_max").font = Font(name=MONO, size=10, color=TINTA)
    calc("D150", "=IF($D$11=3,NA(),MAX(D148,D149))", com150)
    calc("E150", "=IF($D$11=3,NA(),MAX(E148,E149))", com150)

    com151 = ("212-3.2(c): para componentes esfericos, toriesfericos o "
              "elipsoidales (D11=3) se usan calculos de fuerza alternativos; "
              "este motor no los implementa y declara el hand-off a analisis. "
              "Las ec. (1)/(2) de arriba valen solo para cilindro (tuberia/"
              "virola).")
    calc("D151", '=IF($D$11=3,"HAND-OFF 212-3.2(c): esfera/cabezal — usar calculo '
                 'alternativo o analisis; F_max no aplica","cilindro: aplican ec. '
                 '(1) y (2)")', com151)
    ws.merge_cells("D151:G151")

    # --- PASO 3 — Proximidad a discontinuidades (212-3.3) --------------------
    # L_min = 2*sqrt(Rm*t) (ec. 3, bloque [47]) ya vive en D73, y la rama
    # 360°/local ya esta en F88 (3A/3B). Aqui se anade la rama 3C: el limite de
    # proximidad aplica TAMBIEN entre bordes de parches adyacentes [bloque 51].
    # La nota de esquinas redondeadas (R_min = 75 mm) es una RECOMENDACION DEL
    # FLUJO: el Art. 212 solo dice "rounded corners" (bloque [7]-f), no imprime
    # el 75 mm. Se anota como tal (Regla n.1). Solo se LEE D73; no se modifica.
    banda_literal(153, "PASO 3 · PROXIMIDAD A DISCONTINUIDADES  "
                       "—  ASME PCC-2 Art. 212-3.3")
    encabezado(154, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                     (4, "Valor"), (7, "Referencia / Notas")))

    com155 = ("Entrada: distancia entre el borde de soldadura de este parche y "
              "el del parche adyacente mas cercano, mm. Dejar en blanco si no "
              "hay parches adyacentes. El limite de proximidad L_min (D101) "
              "aplica tambien entre parches adyacentes (212-3.3 [51]).")
    lab(155, "Distancia a parches adyacentes", unidad="mm", ref="212-3.3 [51]",
        com=com155)
    ws.cell(155, 2, "L_adj").font = Font(name=MONO, size=10, color=TINTA)
    inp("D155", "", com155)

    com156 = ("Calculo: rama 3C. Si hay parche adyacente, su distancia debe ser "
              ">= L_min (D101); si no, reubicar. En blanco = sin parche adyacente.")
    lab(156, "¿Distancia a parche adyacente ≥ L_mín?", unidad="—",
        ref="212-3.3", com=com156)
    ws.cell(156, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    calc("D156",
         '=IF($D$155="","No aplica (sin parche adyacente)",'
         'IF($D$155>=$D$73,"OK","< L_min — reubicar (212-3.3)"))', com156)

    com157 = ("Nota (recomendacion del flujo, no del texto del 212): las "
              "esquinas del parche deben redondearse con R_min = 75 mm (3 in.) "
              "para no concentrar esfuerzos. El Art. 212 solo dice 'rounded "
              "corners' (bloque 7-f) sin imprimir el radio; el 75 mm lo aporta "
              "el flujo aprobado.")
    lab(157, "Esquinas redondeadas", ref="Flujo · 212 bloque 7-f", com=com157)
    d157 = calc("D157",
                "Esquinas redondeadas: R_min = 75 mm (3 in.) recomendado por el "
                "flujo; el Art. 212 (bloque 7-f) solo exige 'rounded corners'.",
                com157)
    d157.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D157:G157")

    # --- PASO 4 — Soldadura perimetral de filete: topes (212-3.4) -----------
    # w_min (D64:F64) ya usa F_max (arriba). Aqui los DOS topes de la NOTA del
    # bloque [60]: el cateto de diseno no debe exceder ni el menor espesor de
    # los materiales unidos (min(T_parche, t_pared)) ni 40 mm (1.5 in.). Los dos
    # entran al AND de F90. El bisel alternativo (garganta <= nominal) del
    # bloque [61] va como nota. w adoptado = D32; T_parche = D29; t_pared = D21.
    banda_literal(159, "PASO 4 · SOLDADURA DE FILETE, TOPES DE LA NOTA  "
                       "—  ASME PCC-2 Art. 212-3.4")
    encabezado(160, ((1, "Verificación"), (4, "Requerido"), (5, "Adoptado"),
                     (6, "Resultado"), (7, "Criterio")))

    # Los seis comentarios de estas dos filas los destapo la guia de uso de la
    # Fase 10: eran las unicas celdas de formula del 212 sin nota propia, asi que
    # su fila de la guia salia sin explicacion. El pase declara en ISSUES toda
    # celda sin comentario, de modo que un hueco asi ya no puede quedar callado.
    lab(161, "Filete ≤ menor espesor unido", ref="w ≤ min(T_parche, t_pared)")
    ws.merge_cells("A161:C161")
    calc("D161", "=MIN($D$29,$D$21)",
         "Calculo: primer tope de la NOTA de 212-3.4 — el menor de los espesores "
         "unidos: espesor del parche (D29) y espesor de pared del componente (D21).")
    calc("E161", "=$D$32",
         "Calculo: repite el cateto de filete adoptado en la seccion 1 (D32), que "
         "es el que se compara contra el tope.")
    calc("F161", '=IF(E161<=D161,"CUMPLE",'
                 '"NO CUMPLE — excede espesor menor (NOTA 212-3.4)")',
         "Calculo: veredicto del primer tope. La NOTA de 212-3.4 exige que el "
         "cateto de diseno no exceda el espesor del mas delgado de los materiales "
         "unidos. Entra al AND del dictamen global.")

    lab(162, "Filete ≤ 40 mm (1.5 in.)", ref="w ≤ 40 mm")
    ws.merge_cells("A162:C162")
    calc("D162", f'={umbral(UMBR, "212_filete_max", ES_SI_212)}',
         "Calculo: segundo tope de la NOTA de 212-3.4, leido del codigo en el "
         "sistema activo — PCC-2 lo imprime como 40 mm (1.5 in.) y no se convierte "
         "de una unidad a la otra (regla 9).")
    calc("E162", "=$D$32",
         "Calculo: repite el cateto de filete adoptado en la seccion 1 (D32).")
    calc("F162", '=IF(E162<=D162,"CUMPLE","NO CUMPLE — excede el tope de la NOTA 212-3.4")',
         "Calculo: veredicto del segundo tope. Entra al AND del dictamen global.")

    com163 = ("Nota (212-3.4b, bloque [61]): alternativamente el borde del filete "
              "puede biselarse para aumentar la garganta efectiva; en ningun caso "
              "la garganta efectiva debe exceder el espesor nominal del parche ni "
              "el del componente original.")
    lab(163, "Bisel alternativo", ref="212-3.4(b)", com=com163)
    d163 = calc("D163",
                "Bisel opcional (212-3.4b): garganta efectiva <= espesor nominal "
                "del parche o del componente. No excederlo.", com163)
    d163.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D163:G163")

    # --- PASO 5 — Excentricidad y separacion en el borde (212-3.4c / 212-4c) -
    # La excentricidad e (D55) y el esfuerzo de soldadura S_w (D66-D68) ya se
    # reescribieron arriba en forma literal de la ec.(5). Aqui va la ENTRADA de
    # la separacion g del faying edge (212-4c, bloque [82]): distinta de la luz
    # radial de conformado (D31), es el gap de fit-up de la soldadura. Si
    # g >= 1.5 mm, e la incluye (D55). El codigo exige ademas g <= 5 mm (fit-up),
    # que se verifica en el Paso 7. Y la declaracion del hand-off no-cilindro.
    banda_literal(165, "PASO 5 · EXCENTRICIDAD Y SEPARACIÓN EN EL BORDE  "
                       "—  ASME PCC-2 Art. 212-3.4c / 212-4c")
    encabezado(166, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                     (4, "Valor"), (7, "Referencia / Notas")))

    com167 = ("Entrada: separacion en el borde de la placa (faying edge), mm, "
              "medida en el fit-up de la soldadura. Es un dato de campo (se "
              "teclea). Distinta de la luz radial de conformado (D31). Si "
              "g >= 1.5 mm, se suma a la excentricidad e (212-4c); el codigo "
              "exige ademas g <= 5 mm (se verifica en el Paso 7). Fit-up "
              "ajustado del caso semilla: g = 0.")
    lab(167, "Separación en el borde (fit-up)", unidad="mm", ref="212-4(c) [82]",
        com=com167)
    ws.cell(167, 2, "g").font = Font(name=MONO, size=10, color=TINTA)
    inp("D167", 0, com167)

    com168 = ("La ec. (5) del 212-3.4c (S_w = P·Dm/2T + 3·P·Dm·e/T², e = (T+t+g)/2) "
              "aplica solo a CILINDRO (tuberia o virola). En esfera/cabezal "
              "(D11=3) S_w, sus componentes y C_sw son NA: usar analisis "
              "(hand-off 212-3.4). Decision 3 del ingeniero: literal, solo cilindro.")
    lab(168, "Aplicabilidad de la ec. (5)", ref="212-3.4c", com=com168)
    calc("D168",
         '=IF($D$11=3,"HAND-OFF: ec.(5) solo cilindro — esfera/cabezal por '
         'analisis; S_w = NA","cilindro: aplica ec.(5) con e = (T+t+g)/2")',
         com168)
    ws.merge_cells("D168:G168")

    # --- PASO 6 — Conformado en frio: curvatura simple/doble (212-3.5) -------
    # La deformacion (D77) ya usa coef 50/75 y el factor (1-Rf/Ro) (arriba). Aqui
    # va la ENTRADA del radio original Ro (D172): en blanco = plano (Ro=inf,
    # factor 1) [bloque 73]. El coeficiente lo elige la geometria: 75 doble
    # (esfera/cabezal, ec.6 [71]) o 50 simple (cilindro, ec.7 [75]). > 5% exige
    # PWHT post-conformado (212-3.5b [76]) — lo dictamina F86 (seccion 5).
    banda_literal(170, "PASO 6 · CONFORMADO EN FRÍO, CURVATURA SIMPLE / DOBLE  "
                       "—  ASME PCC-2 Art. 212-3.5")
    encabezado(171, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                     (4, "Valor"), (7, "Referencia / Notas")))

    com172 = ("Entrada: radio original de linea media del material del parche Ro, "
              "mm, ANTES de conformarlo. Dejar en blanco si parte de plancha "
              "plana (Ro = infinito -> factor (1-Rf/Ro) = 1) [212-3.5, bloque 73]. "
              "Dato de campo/plano (se teclea).")
    lab(172, "Radio original de línea media", unidad="mm",
        ref="∞ si plano (en blanco)", com=com172)
    ws.cell(172, 2, "Ro").font = Font(name=MONO, size=10, color=TINTA)
    inp("D172", "", com172)

    com173 = ("Rama de curvatura (coef del %Elong, D77): 75 para curvatura DOBLE "
              "(esfera/cabezal, ec. 6) o 50 para curvatura SIMPLE (tuberia/virola, "
              "ec. 7). Si %Elong > 5%, el parche exige PWHT post-conformado antes "
              "de instalar (212-3.5b); lo dictamina la verificacion F86.")
    lab(173, "Rama de curvatura", ref="212-3.5 ec.6/ec.7", com=com173)
    calc("D173",
         '=IF($D$11=3,"Curvatura DOBLE (coef 75, ec.6) — esfera/cabezal",'
         '"Curvatura SIMPLE (coef 50, ec.7) — tuberia/virola")', com173)
    ws.merge_cells("D173:G173")

    # --- PASO 7 — Fabricacion (212-4) — avisos -------------------------------
    # Avisos de fabricacion (no dictamen de diseno): la separacion de fit-up
    # (212-4c [82]: <=5 mm; >=1.5 -> e incluye g, ya en el Paso 5), el examen
    # MT/PT de bordes si T>25 mm (212-4a [80]) y la secuencia/preparacion/venteo
    # (212-4e/g [85],[88],[90]). Aditivo: no toca el oracle ni F90 (la separacion
    # es una constraint de ejecucion, no de aceptacion del diseno).
    banda_literal(175, "PASO 7 · FABRICACIÓN  —  ASME PCC-2 Art. 212-4")
    encabezado(176, ((1, "Aspecto"), (4, "Aviso"), (7, "Referencia")))

    com177 = ("Aviso: separacion de fit-up g (D167). El codigo exige g <= 5 mm "
              "(212-4c); si g >= 1.5 mm, la excentricidad e ya incluye g (Paso 5). "
              "Si g > 5 mm el fit-up no es admisible.")
    lab(177, "Separación de fit-up", ref="212-4(c) [82]", com=com177)
    # Los dos umbrales de 212-4c (maximo admisible y el que obliga a incluir g
    # en la excentricidad) se leen del codigo en las dos unidades. El TEXTO del
    # aviso deja de citar "5 mm" y "1.5 mm" a secas: en modo US esa cifra seria
    # falsa, y un aviso que miente sobre su propio umbral es peor que ninguno.
    g_max = umbral(UMBR, "212_separacion_max", ES_SI_212)
    g_min = umbral(UMBR, "212_separacion_g", ES_SI_212)
    calc("D177",
         f'=IF($D$167>{g_max},"SEPARACION sobre el maximo de 212-4c: fit-up no '
         f'admisible",IF($D$167>={g_min},"g en o sobre el minimo de 212-4c: e '
         f'incluye g — ver Paso 5","fit-up ajustado (g por debajo del minimo)"))',
         com177)
    ws.merge_cells("D177:G177")

    com178 = ("Aviso: si el parche es > 25 mm de espesor y el filete es menor que "
              "el espesor, los bordes de preparacion se examinan por MT/PT para "
              "detectar laminaciones (212-4a). Laminaciones = rechazo salvo "
              "reparacion o FFS (API 579).")
    lab(178, "Examen de bordes (laminaciones)", ref="212-4(a) [80]", com=com178)
    calc("D178",
         f'=IF($D$29>{umbral(UMBR, "212_espesor_examen", ES_SI_212)},'
         f'"Plancha por encima del espesor de 212-4: examinar bordes de preparacion por MT/PT '
         '(laminaciones), 212-4a","Plancha por debajo de ese espesor: sin examen de bordes")',
         com178)
    ws.merge_cells("D178:G178")

    com179 = ("Nota de fabricacion (212-4): corte termico -> esmerilar 1.5 mm de "
              "material calcinado (212-4a); preparar metal blanco en un ancho >= "
              "40 mm a cada lado del cordon (212-4e); costuras existentes bajo el "
              "parche esmeriladas a ras + MT/PT (212-4e); secuencia = costuras "
              "internas del parche primero, luego el perimetro (212-4e); prever "
              "venteo de gas durante el cierre / PWHT (212-4g).")
    lab(179, "Secuencia y preparación", ref="212-4(a)(e)(g)", com=com179)
    d179 = calc("D179",
                "Corte termico: esmerilar 1.5 mm. Metal blanco >= 40 mm a cada "
                "lado (212-4e). Costuras internas primero, luego perimetro "
                "(212-4e). Venteo de gas en el cierre (212-4g).", com179)
    d179.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D179:G179")

    # --- PASO 8 — NDE y prueba de hermeticidad (212-5/6 + App. 501) ----------
    # Selector de prueba (hidrostatica / neumatica). En neumatica se calcula la
    # energia almacenada E (ec. II-1 del App. 501-II, forma general en k), el
    # equivalente en TNT (ec. II-3) y la distancia segura R (ec. III-1 del App.
    # 501-III), y se avisa. TODOS los coeficientes (divisor TNT, umbral del
    # blast wave, distancia fija, valores de Rscaled) se LEEN de resources/ via
    # leer_energia_501() -Fase 0.2-, nunca de memoria (Regla n.1). En
    # hidrostatica el bloque neumatico es NA y se mantiene el criterio 1.5xP.
    if energia_501 is None:
        raise SystemExit("Art.212 Paso 8: energia_501 es None (pase los "
                         "coeficientes del App. 501 leidos de resources/).")
    tnt_div = energia_501["tnt_div_kg"]
    blast_thr = energia_501["blast_thr_J"]
    blast_r = energia_501["blast_R_m"]
    r_ops = energia_501["r_scaled_ops"]
    r_def = energia_501["r_scaled_def"]
    # Fase 7: los MISMOS coeficientes en U.S. Customary, leidos de resources/
    # (el App. 501 publica las dos ediciones). Ninguno se convierte: II-5 trae
    # su propio divisor de TNT y la 501-III-1 su propio umbral y distancia.
    tnt_div_us = energia_501["tnt_div_lb"]
    blast_thr_us = energia_501["blast_thr_ftlb"]
    blast_r_us = energia_501["blast_R_ft"]
    lista_rscaled = '"' + ",".join(f"{v:g}" for _, v in r_ops) + '"'

    banda_literal(181, "PASO 8 · NDE Y PRUEBA DE HERMETICIDAD  "
                       "—  ASME PCC-2 Art. 212-5 / 212-6 + App. 501")
    encabezado(182, ((1, "Parámetro"), (2, "Símbolo"), (3, "Unidad"),
                     (4, "Valor"), (7, "Referencia / Notas")))

    com183 = ("Entrada: tipo de prueba de hermeticidad. Hidrostatica (1.5xP de "
              "diseno, criterio del codigo de post-construccion) o Neumatica. La "
              "prueba neumatica exige precauciones de seguridad (212-6b): se "
              "calcula la energia almacenada y la distancia segura (App. 501).")
    lab(183, "Tipo de prueba", unidad="—", ref="212-6 / App. 501", com=com183)
    ws.cell(183, 2, "—").font = Font(name=MONO, size=10, color=TINTA)
    inp("D183", "Hidrostatica", com183)
    dv_list(ws, "D183", '"Hidrostatica,Neumatica"', com183)

    com184 = ("Entrada (solo neumatica): volumen total bajo presion de prueba, "
              "m3. Para recipiente, el volumen total; para tuberia, hasta 8 "
              "diametros por fallo (App. 501-II).")
    lab(184, "Volumen bajo presión", unidad="m³", ref="App. 501-II", com=com184)
    ws.cell(184, 2, "V").font = Font(name=MONO, size=10, color=TINTA)
    inp("D184", "", com184)

    com185 = ("Entrada (solo neumatica): presion ABSOLUTA de prueba, MPa abs "
              "(Pat de la ec. II-1). Es la presion manometrica de prueba mas la "
              "atmosferica.")
    lab(185, "Presión de prueba (abs)", unidad="MPa abs", ref="App. 501-II",
        com=com185)
    ws.cell(185, 2, "Pat").font = Font(name=MONO, size=10, color=TINTA)
    inp("D185", "", com185)

    com186 = ("Entrada (solo neumatica): presion atmosferica ABSOLUTA, MPa abs "
              "(Pa de la ec. II-1). El codigo usa 101 kPa = 0.101 MPa.")
    lab(186, "Presión atmosférica (abs)", unidad="MPa abs", ref="App. 501-II",
        com=com186)
    ws.cell(186, 2, "Pa").font = Font(name=MONO, size=10, color=TINTA)
    inp("D186", 0.101, com186)

    com187 = ("Entrada (solo neumatica): relacion de calores especificos k del "
              "fluido de prueba (aire/N2 = 1.4). La ec. II-1 es general en k; NO "
              "se usa la forma aire-only 2.5/0.286.")
    lab(187, "Calor específico del fluido (k)", unidad="—", ref="App. 501-II",
        com=com187)
    ws.cell(187, 2, "k").font = Font(name=MONO, size=10, color=TINTA)
    inp("D187", 1.4, com187)

    com188 = ("Entrada (solo neumatica): factor de consecuencia Rscaled de la "
              "Tabla 501-III-1-1 (m/kg^1/3). El codigo exige >= " + f"{r_def:g}"
              " (minimo recomendado). Valores por criterio: "
              + " · ".join(f"{lbl}" for lbl, _ in r_ops) + ".")
    lab(188, "Factor de consecuencia Rscaled", unidad="m/kg^⅓",
        ref="Tabla 501-III-1-1", com=com188)
    ws.cell(188, 2, "Rsc").font = Font(name=MONO, size=10, color=TINTA)
    inp("D188", r_def, com188)
    dv_list(ws, "D188", lista_rscaled, com188)

    com189 = ("Calculo (solo neumatica): energia almacenada E (J) por la ec. "
              "(II-1) del App. 501-II, general en k: E = [1/(k-1)]·Pat·V·"
              "[1-(Pa/Pat)^((k-1)/k)]. Pat se pasa a la unidad de fuerza/area "
              "del sistema: x1e6 (MPa->Pa) en metrico, x144 (psi->lb/ft²) en "
              "U.S. Customary — que es justo lo que convierte la ec. (II-1) en "
              "la (II-4) del codigo: 144/(1,4-1) = 360. NA en hidrostatica.")
    lab(189, "Energía almacenada", unidad="J", ref="App. 501-II ec.(II-1)",
        com=com189)
    ws.cell(189, 2, "E").font = Font(name=MONO, size=10, color=TINTA)
    calc("D189",
         f'=IF($D$183="Neumatica",(1/($D$187-1))*($D$185*IF({ES_SI_212},1000000,144))'
         f'*$D$184*(1-($D$186/$D$185)^(($D$187-1)/$D$187)),NA())', com189)

    com190 = ("Calculo (solo neumatica): equivalente en TNT (kg) por la ec. "
              "(II-3): TNT = E / " + f"{tnt_div}" + " (leido de resources/, "
              "App. 501-II). NA en hidrostatica.")
    lab(190, "Equivalente TNT", unidad="kg", ref="App. 501-II ec.(II-3)",
        com=com190)
    ws.cell(190, 2, "TNT").font = Font(name=MONO, size=10, color=TINTA)
    calc("D190", f'=IF($D$183="Neumatica",$D$189/'
     f'IF({ES_SI_212},{tnt_div},{tnt_div_us}),NA())', com190)

    com191 = ("Calculo (solo neumatica): distancia segura minima R (m) por la "
              "501-III-1: R = 30 m si E <= " + f"{blast_thr}" + " J; en otro caso "
              "R = Rscaled·(2·TNT)^(1/3), ec. (III-1). Umbral y distancia leidos "
              "de resources/ (App. 501-III). NA en hidrostatica.")
    lab(191, "Distancia segura mínima", unidad="m", ref="App. 501-III ec.(III-1)",
        com=com191)
    ws.cell(191, 2, "R").font = Font(name=MONO, size=10, color=TINTA)
    calc("D191",
         f'=IF($D$183="Neumatica",'
         f'IF($D$189<=IF({ES_SI_212},{blast_thr},{blast_thr_us}),'
         f'IF({ES_SI_212},{blast_r:g},{blast_r_us:g}),'
         f'$D$188*(2*$D$190)^(1/3)),NA())', com191)

    com192 = ("Dictamen del Paso 8: en neumatica, la distancia minima entre el "
              "personal y el equipo durante la prueba; ver Tabla 501-III-2-1 "
              "para distancia por fragmentos. En hidrostatica, presion de prueba "
              "= 1.5xP de diseno (D46xD28).")
    lab(192, "Dictamen de prueba", ref="212-6 / App. 501", com=com192)
    calc("D192",
         '=IF($D$183="Neumatica","NEUMATICA — distancia minima R = "&'
         'TEXT($D$191,"0.0")&" m; precaucion 212-6b; ver Tabla 501-III-2-1 '
         '(fragmentos)","HIDROSTATICA — presion de prueba "&TEXT($D$46*$D$28,'
         '"0.0")&" kg/cm2 (1.5xP de diseno); sin energia neumatica")', com192)
    ws.merge_cells("D192:G192")

    com193 = ("Nota NDE (212-5): examinar las soldaduras de union del parche al "
              "100% por MT o PT (212-5a [92]); las costuras entre piezas del "
              "parche por RT o UT en lo posible, o PT/MT multicapa (212-5c [94]); "
              "si hay PWHT, el END va despues del PWHT (212-5d). Criterio de "
              "aceptacion: el del codigo de construccion/post-construccion.")
    lab(193, "Examen no destructivo (NDE)", ref="212-5 [92],[94]", com=com193)
    d193 = calc("D193",
                "NDE (212-5): 100% MT/PT de las uniones del parche; RT/UT de las "
                "costuras entre piezas; END tras PWHT si aplica.", com193)
    d193.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D193:G193")

    # --- Comentarios de las Secciones 1-6 -----------------------------------
    # Se llama al final, cuando TODAS las celdas de las secciones 1-6 ya
    # existen en ws, para que comentar_art212_base() aplique sus notas de
    # verdad: su guardia interna (`if c.value is not None`) no salta en
    # silencio ninguna celda por pertenecer a una seccion todavia no escrita.
    comentar_art212_base(ws)

    # Fase 9: subindices reales en los simbolos (columnas A/B). Ultimo paso, con
    # todas las celdas ya escritas.
    _aplicar_subindices(ws)

    # Fase 3: la Seccion de Material sube justo detras de la Seccion 1. Se hace
    # aqui, con la hoja entera ya escrita (comentarios y subindices incluidos),
    # aplicando UN mapa de filas — ver _mapa_filas_212().
    remapear_filas(ws, MAPA_FILAS_212)

    # Fase 7: rotulos de unidad de todo el motor, con los rangos ya finales.
    aplicar_unidades_motor(ws, UNIDADES_212, ES_SI_212)

    # Fase 5: donde va un comentario y donde no. Se declara por seccion y se
    # reparte lo que ya existe; los rangos son los de DESPUES del remapeo.
    aplicar_reglas_de_comentario(ws, REGLAS_COMENTARIO_212)

    # Leyenda de color de celda + su aplicacion. Van al final, cuando ya existe
    # cada celda en su sitio definitivo: el pase DERIVA el color del estado real
    # (desbloqueada / formula / rotulo) y correrlo antes dejaria sin pintar lo
    # que falte.
    build_leyenda_motor(ws)
    aplicar_leyenda_motor(ws)

    # Fase 6: DESPUES de la leyenda, a proposito. El pase de leyenda pinta de
    # gris toda celda de formula, y el dictamen global lo es; pero el dictamen
    # no es un dato mas de la tabla sino la frase que se lee primero, asi que
    # su banda y su semaforo tienen que ganar.
    aplicar_semaforo_motor(ws, SEMAFORO_212, FILA_DICTAMEN_212)

    # Fase 8: manifiesto de entradas + boton de reinicio. Va al final, con las
    # filas ya remapeadas y la proteccion de cada celda ya definitiva: el
    # manifiesto es una lista de direcciones y se DERIVA de ese estado.
    build_reinicio_motor(ws)

    # Fase 9: atajo a las especificaciones tecnicas, que dejaron esta hoja.
    build_botones_documentos(ws, espec=ESPEC_212, instr=INSTR_212)

    # Fase 11: el area de impresion es A..G, no el area de uso (que llega a las
    # columnas ocultas). Se descubrio al exportar la hoja para revisarla.
    preparar_impresion(ws, MOTOR_NCOLS)

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
        18: "Entrada: diametro nominal (NPS), del desplegable de DB_B36 segun la "
            "norma dimensional elegida en D22. No se teclea (regla 14). Alimenta "
            "el lookup de OD y espesor de pared.",
        19: "Entrada: cedula (designador de Schedule) del NPS elegido, del "
            "desplegable de DB_B36 segun la norma de D22. No se teclea (regla 14).",
        20: "Calculo: diametro exterior (OD) por lookup del par (NPS|cedula) "
            "contra DB_B36 de la norma elegida (D22).",
        21: "Calculo: espesor de pared por lookup del par (NPS|cedula) contra "
            "DB_B36 de la norma elegida (D22).",
        22: "Entrada: norma dimensional de la tuberia (B36.10M acero al carbono / "
            "B36.19M inoxidable), nivel 0 de la cascada dimensional. Gobierna de "
            "que base (DB_B36_10 o DB_B36_19) leen NPS, cedula, OD y espesor.",
        24: "Entrada informativa: fluido de servicio. No alimenta ningun "
            "calculo de esta hoja.",
        25: "Entrada: temperatura de operacion, en °C. Alimenta la Temperatura "
            "de evaluacion de la seccion 2 (D40) y por tanto todo el S(T) "
            "resuelto.",
        26: "Entrada: presion de operacion, en kg/cm². Es el caso 'Operacion' "
            "evaluado en la seccion 3.",
        # Fase 2: la fila 27 ('Diseno tipico') se retiro; su entrada aqui se va
        # con ella. El guardia `if c.value is not None` del bucle la saltaria de
        # todas formas, pero dejarla escrita haria creer que la fila existe.
        28: "Entrada: presion de diseno, en kg/cm² — la MAXIMA ADMISIBLE "
            "(rating), no un valor tipico intermedio. Gobierna el espesor "
            "requerido y el esfuerzo de soldadura (212-3.2 define una unica P "
            "de diseno; 206-3.3 la llama 'maximum allowable design pressure'). "
            "Es el caso 'Diseno' de la seccion 3 y se compara contra la presion "
            "maxima admisible del parche en la verificacion de la fila 117.",
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
            "seccion 5 (fila 112).",
        33: "Entrada: solape minimo del parche sobre metal sano, en mm, "
            "exigido por el Art. 212. Solo informativo en esta hoja.",
        34: "Entrada: diametro del defecto, caracterizado por UT/PT, en mm. "
            "Solo informativo en esta hoja.",
        35: "Entrada: distancia del defecto a la discontinuidad mas cercana "
            "(cordon o boquilla), en mm. Se compara contra L_min en la "
            "verificacion de la seccion 5 (fila 116) para decidir entre "
            "refuerzo 360° y parche local.",
        41: "Calculo: menor entre el esfuerzo admisible del collar (D67) y el "
            "del metal base (D68) — rige el diseno por criterio conservador.",
        42: "Entrada: eficiencia de junta de filete adoptada (Art. 212 ec. 4); "
            "verificar contra el WPS/PQR calificado.",
        43: "Entrada: factor Y de la Tabla 304.1.1 del B31.3, segun material y "
            "temperatura; se ingresa a mano, no se interpola automaticamente.",
        45: "Entrada: densidad del acero adoptada para el calculo de peso del "
            "parche/collar (seccion 4).",
        46: "Entrada: factor multiplicador de la presion de diseno para la "
            "prueba hidrostatica (tipico 1,5x, B31.3 345.4.2). Alimenta la "
            "especificacion de prueba de hermeticidad (fila 127).",
        47: "Constante: factor de conversion de kg/cm² a MPa (1 kg/cm² = "
            "0,0980665 MPa). No cambiar salvo error de unidades.",
        48: "Calculo: 1,5 veces el esfuerzo admisible gobernante (D69) — limite "
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
            "seccion 5, fila 116).",
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

    # Fase 2: `trip` ya no son tres columnas sino DOS (D/E = Operacion/Diseno);
    # el nombre se conserva para no renombrar por estetica algo que el resto del
    # archivo cita, pero el bucle de abajo recorre (4, 5), no (4, 5, 6).
    trip = {
        61: "Calculo: repite, para este caso (Operacion / Diseno), la presion "
            "correspondiente de la seccion 1, en kg/cm². 'Diseno' es la maxima "
            "admisible (D28): 212-3.2 define una unica P de diseno.",
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
            "flexion); se compara contra el limite 1,5·Sa en la fila 97.",
        69: "Calculo: CUMPLE si el esfuerzo de soldadura total de este caso "
            "(fila 96) no supera 1,5 veces el esfuerzo admisible gobernante "
            "(D76).",
    }
    for r, com in trip.items():
        for col in (4, 5):
            c = ws.cell(r, col)
            if c.value is not None:
                _nota(c, com)

    verif = {
        84: ("Calculo: cateto de filete minimo requerido, el mayor de los dos "
             "casos de presion (fila 92).",
             "Entrada: cateto adoptado (repite D32).",
             "Calculo: CUMPLE si el cateto adoptado es mayor o igual al "
             "requerido."),
        85: ("Calculo: limite de esfuerzo por excentricidad, 1,5·Sa (repite "
             "D48).",
             "Calculo: esfuerzo de soldadura total adoptado para el diseno "
             "(repite E68, caso Diseno = presion maxima admisible).",
             "Calculo: CUMPLE si el esfuerzo adoptado no supera el limite."),
        86: ("Entrada: limite normativo de deformacion por conformado en "
             "frio, 5 % (ec. 7).",
             "Calculo: deformacion por conformado calculada (repite D77).",
             "Calculo: CUMPLE si la deformacion calculada no supera el 5 %."),
        87: ("Calculo: espesor de pared requerido, caso Diseno (repite E65, "
             "evaluado a la presion maxima admisible — 206-3.3).",
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
             "Entrada: presion de diseno adoptada, la maxima admisible "
             "(repite D28).",
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

    # Las siete filas de especificaciones tecnicas (93-99) ya no viven en esta
    # hoja: la Fase 9 las saco a Espec_PCC2_Art212. Su diccionario de
    # comentarios se retira con ellas y NO se deja "por si acaso": el bucle
    # estaba guardado por `if c.value is not None`, asi que un diccionario
    # huerfano no daria error — se quedaria callado, que es peor. El comentario
    # de cada especificacion vive ahora en su fila de la pestana.

    if ws.cell(101, 1).value is not None:
        _nota(ws.cell(101, 1), "Aviso fijo: alcance y limitaciones de la "
                              "herramienta. No se edita.")


COLLAR_MOTOR = "Collar_PCC2_Art206"

TIPO_A = "Type A (no contiene presion)"
TIPO_B = "Type B (contiene presion)"


def build_collar_art206(wb, b313, iid1a, iidb, fac_info, rangos, b3610, b3619,
                        b313c=None, iid1ac=None, iidbc=None, umbrales=None):
    """Motor Art. 206 (collar de encierro total, Type A y Type B), 100% en
    codigo — a diferencia de Parche_PCC2_Art212, no hereda nada del maestro
    Rev0. Reutiliza la cascada de material de construir_seccion7_material
    (compartida con build_parche_art212) y, desde la Tarea 9, la misma cascada
    dimensional norma -> NPS -> cedula -> OD/espesor del Art. 212: NPS/cedula/
    OD/espesor salen de DB_B36_10/DB_B36_19 (reglas 12/14), no de la lista fija
    ni del lookup corto de Datos_Ref.

    El Art. 206 no tiene ecuaciones propias de membrana/filete/conformado en
    frio (esas son del Art. 212): remite al codigo de construccion para el
    t_req (206-3.3) y solo anade reglas geometricas impresas en el propio
    articulo (206-3.1, 206-3.2, 206-3.4, 206-3.5), citadas de
    resources/ASME PCC/pcc_2/p2_welded_repairs/art_206_full_encirclement_steel.
    """
    if umbrales is None:
        raise SystemExit(
            "build_collar_art206: faltan los umbrales de longitud de PCC-2 (leer_umbrales_pcc2). "
            "Un umbral de aceptacion no tiene valor por defecto.")
    UMBR = umbrales

    ws = new_sheet(wb, COLLAR_MOTOR,
                   "MOTOR DE CALCULO — COLLAR DE ENCIERRO TOTAL (ASME PCC-2 Art. 206)",
                   "ASME PCC-2 Art. 206 (Full Encirclement Steel Reinforcing Sleeves), "
                   "Type A y Type B. Fuente: resources/ASME PCC/pcc_2/"
                   "p2_welded_repairs/art_206_full_encirclement_steel/art_206.json")
    # Titulo y subtitulo fusionados A:G, igual que en el Art. 212. Sin el merge,
    # el titulo se cortaba a media palabra —«[ MOTOR DE CALCULO // COLLAR DE
    # ENC»— en cuanto la celda de al lado dejaba de estar libre. Visto en el PDF
    # exportado de la Fase 11.
    ws.merge_cells(f"A1:{get_column_letter(MOTOR_NCOLS)}1")
    ws.merge_cells(f"A2:{get_column_letter(MOTOR_NCOLS)}2")
    # Columna A a 50: con 44, el rotulo mas largo de la hoja —«Presion de diseno
    # (maxima admisible / rating)»— se cortaba a media palabra contra la columna
    # de simbolo, que no esta vacia y por tanto no deja desbordar el texto. Visto
    # exportando la hoja a PDF en la Fase 11, no en openpyxl.
    autosize(ws, {"A": 50, "B": 8, "C": 14, "D": 16, "E": 16, "F": 16, "G": 46})

    def lab(r, text, unidad=None, ref=None, com=None):
        cl = ws.cell(r, 1, text)
        cl.font = Font(name=MONO, size=10, color=TINTA)
        if unidad:
            ws.cell(r, 3).value = unidad
        if ref:
            ws.cell(r, 7, ref).font = Font(name=MONO, size=9, color=GRIS)
        if com:
            _nota(cl, com)

    def inp(cell, value="", com=None):
        c = ws[cell]
        c.value = value
        c.font, c.fill, c.border = IN_F, MOTOR_IN_FILL, CAJA_CAMPO
        c.protection = Protection(locked=False)
        if com:
            _nota(c, com)

    def calc(cell, formula, com=None):
        c = ws[cell]
        c.value = formula
        if com:
            _nota(c, com)
        return c

    def band(r, texto):
        ws.cell(r, 1, rotulo(texto))
        ws.cell(r, 1).font = Font(name=MACRO, size=11, color=PAPEL)
        for j in range(1, 8):
            ws.cell(r, j).fill = BAND_FILL
        franja(ws, r, 1, 7)

    def header(r, cols):
        for j, h in enumerate(cols, start=1):
            c = ws.cell(r, j, h.upper())
            c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA

    # --- Seccion 7 primero: la cascada de material del sleeve --------------
    # Se llama antes de escribir el resto de la hoja para poder cablear
    # 'Parametros de calculo' contra sus referencias de salida, exactamente
    # como el orden de escritura no le importa a Excel (solo el contenido
    # final de cada celda).
    refs = construir_seccion7_material(
        ws, 75, b313, iid1a, iidb, fac_info, rangos,
        columnas=(("D", "Sleeve (Art. 206)"),),
        modo_cell="$D$11", temp_fuente_cell="$D$25", incluir_ej_ec=False,
        destino_st="D60 de '3. Parametros de calculo'",
        mapa_citas=MAPA_FILAS_206,
        b313c=b313c, iid1ac=iid1ac, iidbc=iidbc, unidad_cell=UNIDAD_206)
    s_t_sleeve = refs["por_columna"]["D"]["s_t"]
    dictamen_sleeve = refs["por_columna"]["D"]["dictamen"]

    # --- Identificacion ------------------------------------------------
    ws.cell(4, 1, rotulo("IDENTIFICACION"))
    ws.cell(4, 1).font, ws.cell(4, 1).fill = Font(name=MACRO, size=11, color=PAPEL), BAND_FILL
    for j in range(1, 8):
        ws.cell(4, j).fill = BAND_FILL
    franja(ws, 4, 1, 7)
    lab(5, "Documento", com="Entrada: numero de documento de la reparacion.")
    inp("D5", com="Entrada: numero de documento de la reparacion.")
    lab(6, "Componente / servicio",
       com="Entrada: identificacion del componente y su servicio.")
    inp("D6", com="Entrada: identificacion del componente y su servicio.")

    # --- Aplicacion y codigo de construccion ----------------------------
    band(8, "APLICACION Y CODIGO DE CONSTRUCCION")
    header(9, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com_app = ("Entrada: elija el tipo de componente portador. Determina el modo "
              "(D11) que rige el t_req de presion del sleeve Type B (seccion "
              "'Calculo de espesor requerido') igual que en Parche_PCC2_Art212.")
    lab(10, "Aplicacion / geometria del portador   [D10]", com=com_app)
    inp("D10", "Tuberia (B31.3)", com_app)
    dv_list(ws, "D10",
            '"Tuberia (B31.3),Virola cilindrica (VIII-1),Cabezal/esfera (VIII-1)"',
            com_app)
    lab(11, "Modo (1=tuberia, 2=virola, 3=cabezal/esfera)", com=com_app)
    calc("D11", '=IF($D$10="Tuberia (B31.3)",1,IF($D$10="Cabezal/esfera (VIII-1)",3,2))',
        com_app)
    lab(12, "Codigo de construccion")
    calc("D12", '=IF($D$11=1,"ASME B31.3","ASME BPVC VIII-1")',
         "Calculo: ASME B31.3 si MODO=1 (tuberia); ASME BPVC VIII-1 en los demas "
         "casos. El codigo activo decide de que tabla sale el esfuerzo admisible y "
         "cual es el criterio de aceptacion del examen.")
    lab(13, "Fuente del esfuerzo admisible")
    calc("D13", '=IF($D$11=1,"B31.3 Tabla A-1","ASME II-D Tabla 1A")',
         "Calculo: cita la tabla del codigo activo de la que sale el esfuerzo "
         "admisible S — Tabla A-1 del B31.3 si MODO=1, Tabla 1A de ASME II-D en los "
         "demas casos.")

    # Fase 7: selector de sistema de unidades. Aqui cabe en la fila 14 sin
    # mover nada -esta banda ya dejaba dos filas en blanco antes de la
    # siguiente-, asi que el 206 no necesita el desplazamiento que si hizo
    # falta en el 212.
    com_uni = ("Entrada: sistema de unidades de LECTURA del codigo. Cambia de que "
    "edicion se lee el esfuerzo admisible -la metrica o la U.S. "
    "Customary-, NUNCA convierte un valor (regla 9: las dos ediciones "
    "son extracciones independientes de lo que cada una imprime). Los "
    "datos que usted teclea -presiones, dimensiones, temperatura- NO se "
    "convierten al cambiarlo: hay que volver a teclearlos en el sistema "
    "nuevo, igual que en los cinco buscadores de cascada.")
    lab(14, "Sistema de unidades", "—", "SI (metrico) / US Customary", com_uni)
    inp("D14", "SI", com_uni)
    dv_list(ws, "D14", '"SI,US"', com_uni)

    # --- 1. Datos de entrada ---------------------------------------------
    # El parentetico describia el sistema visual anterior ("linea inferior =
    # editable"); con la leyenda de color de la Fase 1 seria falso.
    band(16, "1. DATOS DE ENTRADA")
    header(17, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    # --- Bloque dimensional (Tarea 9): norma -> NPS -> cedula -> OD/espesor, todo
    # contra DB_B36_10/DB_B36_19 por cascada de listas (reglas 12/14), igual que el
    # Art. 212. NPS/cedula/OD/espesor conservan sus filas 18-21 (no se corren: D45/
    # D46 referencian OD(D20) y D54/D64 referencian espesor(D21), y moverlas
    # reapuntaria esas formulas). El selector de norma no cabe adjunto —las filas
    # 18-32 estan todas ocupadas— asi que ocupa la fila 33, libre al final de la
    # seccion 1; misma decision que la fila 22 del Art. 212.
    com18 = ("Entrada: diametro nominal (NPS) del tubo portador, del desplegable de "
             "DB_B36 segun la norma elegida en D33. No se teclea (regla 14). Se "
             "guarda como el NPS impreso del codigo, p.ej. '12 (300)'. Alimenta OD, "
             "espesor y la lista de cedulas.")
    lab(18, "NPS del tubo portador", "in", "Lista · DB_B36 (segun D33)")
    inp("D18", "12 (300)", com18)
    com19 = ("Entrada: cedula (designador de Schedule) del NPS elegido, en la norma "
             "de D33, del desplegable de DB_B36. No se teclea (regla 14).")
    lab(19, "Cedula del tubo portador", None, "Lista · DB_B36 (segun D33 y NPS)")
    inp("D19", "20", com19)
    com20 = ("Calculo: diametro exterior (OD, mm) por lookup de la clave NPS|cedula "
             "(D18|D19) contra DB_B36 de la norma elegida (D33). Vacio si el par no "
             "existe en la base.")
    lab(20, "OD del tubo portador", "mm", "DB_B36 (auto, segun D33)")
    com21 = ("Calculo: espesor de pared (mm) por lookup de la clave NPS|cedula "
             "(D18|D19) contra DB_B36 de la norma elegida (D33). Vacio si el par no "
             "existe en la base.")
    lab(21, "t del tubo portador", "mm", "DB_B36 (auto, segun D33)")
    com_tipo = ("Entrada: 206-1.1.1 Type A —extremos NO soldados "
               "circunferencialmente, no contiene presion, actua como "
               "refuerzo—; 206-1.1.2 Type B —extremos soldados "
               "circunferencialmente, SI contiene presion—. Determina que "
               "regla de espesor rige (206-3.1 o 206-3.2/3.3).")
    lab(22, "Tipo de sleeve   [206-1.1.1 / 206-1.1.2]", com=com_tipo)
    inp("D22", TIPO_A, com_tipo)
    dv_list(ws, "D22", f'"{TIPO_A},{TIPO_B}"', com_tipo)
    com_ut = ("Entrada: 206-3.2 — un factor de eficiencia de junta longitudinal "
             "de 0,80 se aplica al t_req del Type B, salvo que la costura "
             "longitudinal del sleeve se examine 100% por ultrasonido, en cuyo "
             "caso el factor es 1,00. Ligado a 206-4.4 (la costura longitudinal "
             "va a tope con penetracion completa si hay filetes de cierre).")
    lab(23, "Soldadura longitudinal del sleeve examinada 100% UT?   [206-3.2]",
       com=com_ut)
    inp("D23", "No", com_ut)
    dv_list(ws, "D23", '"Si,No"', com_ut)
    com_circ = ("Entrada: 206-2.5 — el Type A puede no ser apto para defectos "
               "circunferenciales porque no resiste cargas axiales. Solo "
               "alimenta el aviso de la seccion de verificaciones; no bloquea "
               "el calculo.")
    lab(24, "Defecto circunferencial?   [206-2.5]", com=com_circ)
    inp("D24", "No", com_circ)
    dv_list(ws, "D24", '"Si,No"', com_circ)
    com25_206 = ("Entrada: temperatura de operacion del tubo portador. Es la "
                 "temperatura a la que se lee el esfuerzo admisible del collar en "
                 "la seccion de resolucion de material; no la publica ninguna "
                 "tabla, la da el servicio.")
    lab(25, "Temperatura de operacion", "°C", com=com25_206)
    inp("D25", 25, com25_206)
    com26_206 = ("Entrada: presion de operacion del tubo portador. Es el caso "
                 "'Operacion' de la seccion de calculo de espesor; el espesor "
                 "gobernante NO sale de ella sino de la presion de diseno "
                 "(206-3.3).")
    lab(26, "Presion de operacion", "kg/cm2", com=com26_206)
    inp("D26", 5, com26_206)
    # Fase 2 (modelo de presion de 2 casos, igual que el Art. 212). La fila 27
    # —"Presion de diseno tipica"— se retira: 206-3.3 dimensiona contra "the
    # maximum allowable design pressure" (resources/ASME PCC/pcc_2/
    # p2_welded_repairs/art_206_full_encirclement_steel/art_206.json, Regla n.1),
    # no contra un valor tipico intermedio. Tener las dos dejaba dos columnas
    # compitiendo por gobernar el T_s,min. La fila 27 queda vacia a proposito.
    # La REFERENCIA de la columna G va corta: el parrafo entero se salia del
    # ancho de la columna y se cortaba al imprimir. La explicacion completa vive
    # en el comentario de D28, que es donde la regla de la Fase 5 la pide.
    lab(28, "Presion de diseno (maxima admisible / rating)", "kg/cm2",
       "206-3.3 · maximum allowable design pressure")
    inp("D28", 20,
        "Entrada: 206-3.3 la llama 'maximum allowable design pressure'. Es la que "
        "gobierna el t_req de Type B, no un valor tipico intermedio.")
    com29_206 = ("Entrada: espesor nominal que usted adopta para el collar (T_s "
                 "de las Figs. 206-3.5-1/-2). La verificacion lo compara contra el "
                 "T_s,min gobernante; tambien decide el cateto del filete de "
                 "extremo, que cambia segun T_s <= 1,4 T_p o no.")
    lab(29, "Espesor adoptado del sleeve", "mm", com=com29_206)
    ws.cell(29, 2, "T_s").font = Font(name=MONO, size=10, color=TINTA)
    inp("D29", 8, com29_206)
    com30_206 = ("Entrada: longitud del defecto medida en campo, a lo largo del "
                 "eje del tubo. Fija la longitud minima del collar: 206-3.4 exige "
                 "que sobrepase el defecto por los dos extremos.")
    lab(30, "Longitud del defecto", "mm", com=com30_206)
    inp("D30", 100, com30_206)
    com31_206 = ("Entrada: luz radial entre collar y tubo portador (G de las "
                 "Figs. 206-3.5-1/-2), medida tras el ajuste. 206-4.1 pide en "
                 "general ajuste 'sin luz' y admite un maximo; la verificacion la "
                 "compara contra ese maximo, y G se suma al cateto del filete.")
    lab(31, "Luz radial sleeve-portador", "mm", com=com31_206)
    ws.cell(31, 2, "G").font = Font(name=MONO, size=10, color=TINTA)
    inp("D31", 1.5, com31_206)
    com32_206 = ("Entrada: longitud que usted adopta para el collar. La "
                 "verificacion la compara contra la minima que exige 206-3.4 "
                 "(longitud del defecto mas el sobrepaso por cada extremo, y nunca "
                 "menos del minimo absoluto del codigo).")
    lab(32, "Longitud adoptada del sleeve", "mm", com=com32_206)
    ws.cell(32, 2, "L_s").font = Font(name=MONO, size=10, color=TINTA)
    inp("D32", 250, com32_206)

    # Fila 33: selector de norma dimensional (nivel 0 de la cascada). Va al final
    # de la seccion 1 porque las filas 18-32 estan ocupadas y no pueden correrse
    # sin reapuntar las formulas de calculo; la norma se ELIGE, no se deriva del
    # material (ver _materializar_cascada_b36). Gobierna las listas de NPS/cedula
    # (D18/D19) y el lookup de OD/espesor (D20/D21) de arriba.
    com33 = ("Entrada: norma dimensional del tubo portador. B36.10M (acero al "
             "carbono y de baja aleacion) o B36.19M (inoxidable, cedulas de la "
             "serie S). Gobierna las listas de NPS y cedula y el lookup de OD/"
             "espesor de la seccion 1.")
    lab(33, "Norma dimensional", "—", "ASME B36.10M / B36.19M")
    inp("D33", "B36.10M", com33)

    casc = _materializar_cascada_b36(ws, b3610, b3619,
                                     fila_norma=33, fila_nps=18, fila_ced=19)
    idx = casc["idx"]
    clave_key = '$D$18&"|"&$D$19'
    calc("D20", f'=IFERROR(INDEX({_choose_b36(idx, b3610, b3619, "od_mm", ES_SI_206, "od_in")},'
                f'MATCH({clave_key},{_choose_b36(idx, b3610, b3619, "clave")},0)),"")',
         com20)
    calc("D21", f'=IFERROR(INDEX({_choose_b36(idx, b3610, b3619, "t_mm", ES_SI_206, "t_in")},'
                f'MATCH({clave_key},{_choose_b36(idx, b3610, b3619, "clave")},0)),"")',
         com21)

    # --- Parametros de calculo -------------------------------------------
    band(35, "3. PARAMETROS DE CALCULO")
    header(36, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    lab(37, "S(T) del sleeve", "MPa",
       "Calculo: trae el S(T) del sleeve resuelto en la seccion de resolucion "
       "de material si su Dictamen de rango es OK; NA() si no.")
    calc("D37", f'=IF({dictamen_sleeve}="OK",{s_t_sleeve},NA())',
         "Calculo: trae el S(T) del sleeve resuelto en la seccion de resolucion de "
         "material si su Dictamen de rango es OK; NA() si no. Un NA() aqui propaga "
         "hacia abajo a proposito: sin esfuerzo admisible no hay t_req que valga.")
    com_ej206 = ("Calculo: 206-3.2 — 0,80 salvo que D23='Si' (costura "
                "longitudinal del sleeve examinada 100% por UT), en cuyo caso "
                "1,00. NO es un lookup de la Tabla A-3 (a diferencia del Art. "
                "212): es el factor propio que el Art. 206 imprime para la "
                "junta longitudinal del sleeve.")
    lab(38, "Ej del sleeve   [206-3.2]", com=com_ej206)
    calc("D38", '=IF($D$23="Si",1,0.8)', com_ej206)
    lab(39, "Factor Y (B31.3 Tabla 304.1.1)", None,
       "Entrada manual: solo se usa si D11=1 (tuberia).")
    inp("D39", 0.4, "Entrada manual: solo se usa si D11=1 (tuberia).")
    lab(40, "kg/cm2 -> MPa")
    calc("D40", f'=IF({ES_SI_206},0.0980665,0.001)', "Calculo (Fase 7): la constante depende del SISTEMA DE UNIDADES, no del "
        "caso, asi que la fija el selector y no se teclea. Dejarla editable "
        "obligaria a acordarse de cambiarla al conmutar, y olvidarlo daria un "
        "resultado plausible y equivocado.")

    # --- Geometria del sleeve --------------------------------------------
    band(43, "4. GEOMETRIA DEL SLEEVE")
    header(44, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com_geo = ("Calculo: 206-3.3 remite al codigo de construccion sin fijar el "
              "diametro a usar; se deriva del tubo portador + luz radial + "
              "espesor adoptado del sleeve, mismo criterio que Rf en Parche_"
              "PCC2_Art212 (D81). Confirmar con el ingeniero antes de dar la "
              "hoja por cerrada — es el unico punto de esta hoja que no sale "
              "palabra por palabra de 206-3.2/3.3.")
    lab(45, "OD del sleeve", "mm", com=com_geo)
    calc("D45", "=$D$20+2*$D$31+2*$D$29", com_geo)
    lab(46, "Radio interior del sleeve", "mm", com=com_geo)
    calc("D46", "=$D$20/2+$D$31", com_geo)

    # --- Calculo de espesor requerido (Type B) ---------------------------
    band(49, "5. CALCULO DE ESPESOR REQUERIDO, Type B — 206-3.2/3.3")
    # Fase 2: dos casos (Operacion / Diseno), no tres. La columna F se retira y
    # E pasa a leer la presion de diseno maxima admisible (D28) — ver la nota de
    # la fila 27/28 de la seccion 1.
    header(50, ["Parametro", "", "Unidad", "Operacion", "Diseno",
               "", "Referencia / Notas"])
    com_pev = ("Calculo: la presion de cada caso, traida de la seccion 1 — "
               "columna Operacion de D26 y columna Diseno de D28 (la maxima "
               "admisible que exige 206-3.3). El t_req gobernante sale de la "
               "columna Diseno.")
    lab(51, "Presion evaluada", "kg/cm2", com=com_pev)
    calc("D51", "=$D$26", com_pev)
    calc("E51", "=$D$28", com_pev)
    com_pev2 = ("Calculo: la misma presion del caso, convertida a la unidad de "
                "esfuerzo con el factor de la seccion de parametros (D40). Las "
                "ecuaciones del codigo son dimensionales: la presion y el esfuerzo "
                "admisible tienen que entrar en la misma unidad.")
    lab(52, "Presion evaluada", "MPa", com=com_pev2)
    calc("D52", "=D51*$D$40", com_pev2)
    calc("E52", "=E51*$D$40", com_pev2)
    com_treq = ("Calculo: 206-3.3 — t_req por el codigo de construccion activo "
               "(D11), misma forma que D65 de Parche_PCC2_Art212 pero con el "
               "S(T) (D62), Ej (D63) y geometria (D70/D71) del SLEEVE, no del "
               "tubo portador. Solo aplica a Type B: Type A no contiene "
               "presion (206-3.1).")
    lab(53, "t_req del codigo (Type B)", "mm", com=com_treq)
    # Fase 2: el t_req Type B suma el sobreespesor de corrosion C.A. (D113,
    # 206-3.3: "corrosion allowances ... in accordance with the engineering
    # design"). El Type A (D54) NO lo lleva: 206-3.1 no es componente a presion.
    # Con C.A.=0 (default) el valor no cambia respecto de antes.
    calc("D53", '=(IF($D$11=1,D52*$D$45/(2*($D$37*$D$38+D52*$D$39)),'
                'IF($D$11=3,D52*$D$46/(2*$D$37*$D$38-0.2*D52),'
                'D52*$D$46/($D$37*$D$38-0.6*D52))))+$D$113', com_treq)
    calc("E53", '=(IF($D$11=1,E52*$D$45/(2*($D$37*$D$38+E52*$D$39)),'
                'IF($D$11=3,E52*$D$46/(2*$D$37*$D$38-0.2*E52),'
                'E52*$D$46/($D$37*$D$38-0.6*E52))))+$D$113', com_treq)
    lab(54, "T_s,min Type A", "mm", "Referencia / Notas: 206-3.1",
       "Calculo: 206-3.1 — dos tercios del espesor del tubo portador. No "
       "depende de la presion: Type A no es componente a presion.")
    calc("D54", "=2/3*$D$21",
        "Calculo: 206-3.1 — dos tercios del espesor del tubo portador.")
    com_gob = ("Calculo: si D22=Type A, rige 206-3.1 (D79); si Type B, rige el "
               "caso Diseno de 206-3.2/3.3 (E78), evaluado a la presion maxima "
               "admisible, que es la que el codigo nombra en 206-3.3.")
    lab(55, "T_s,min gobernante", "mm", com=com_gob)
    calc("D55", f'=IF($D$22="{TIPO_A}",$D$54,$E$53)', com_gob)

    # --- Verificaciones y avisos ------------------------------------------
    band(58, "6. VERIFICACIONES Y AVISOS")
    header(59, ["Parametro", "Requerido", "Adoptado", "Resultado", "", "",
               "Referencia / Notas"])
    lab(60, "Espesor del sleeve", com="Calculo: T_s adoptado (D29) contra T_s,min "
                                     "gobernante (D80).")
    calc("D60", "=$D$55")
    calc("E60", "=$D$29")
    calc("F60", '=IF(E60>=D60,"CUMPLE","NO CUMPLE")')
    lab(61, "Longitud del sleeve", "mm", "206-3.4",
       "Calculo: 206-3.4 — minimo 100 mm y ademas defecto + 2x50 mm de "
       "extension. L_s adoptado (D32) contra ese minimo.")
    calc("D61", f'=MAX({umbral(UMBR, "206_longitud_min", ES_SI_206)},'
         f'$D$30+2*{umbral(UMBR, "206_sobrepaso", ES_SI_206)})')
    calc("E61", "=$D$32")
    calc("F61", '=IF(E61>=D61,"CUMPLE","NO CUMPLE")')
    lab(63, "Aviso Type A + defecto circunferencial", None, "206-2.5")
    calc("F63", (f'=IF(AND($D$22="{TIPO_A}",$D$24="Si"),'
                 '"AVISO: Type A puede no ser apto para defectos '
                 'circunferenciales (206-2.5)","")'),
        "Calculo: 206-2.5 — Type A no resiste cargas axiales; el ingeniero "
        "decide si un defecto circunferencial exige Type B.")
    lab(64, "Filete de extremo (Type B)", None, "206-3.5")
    calc("F64", (f'=IF($D$22="{TIPO_B}",IF($D$29<=1.4*$D$21,'
                 '"Filete completo (206-3.5a)",'
                 '"Extremos as-is/achaflanados (206-3.5b)"),'
                 '"No aplica - Type A no lleva soldadura circunferencial '
                 '(206-1.1.1)")'),
        "Calculo: 206-3.5 — filete completo si T_s <= 1,4x el espesor nominal "
        "del tubo portador; si es mas grueso, extremos as-is o achaflanados.")
    nota206 = ws.cell(65, 1, "206-4.4: si hay filetes de cierre circunferenciales, "
                    "la costura longitudinal del sleeve va a tope con penetracion "
                    "completa (prever venteo); D23='Si' habilita Ej=1,00 (206-3.2).")
    nota206.font = SRC_F
    _nota(nota206, "Aviso fijo: 206-4.4, ligado al selector D23. No se edita.")
    nota206b = ws.cell(66, 1, "206-3.11: considerar la dilatacion termica "
                    "diferencial entre el tubo portador y el sleeve.")
    nota206b.font = SRC_F
    _nota(nota206b, "Aviso fijo: 206-3.11. No se edita.")

    # --- Dictamen global ----------------------------------------------------
    lab(69, "DICTAMEN GLOBAL DEL DISEÑO")
    ws.cell(69, 1).font = Font(name=MACRO, size=11, color=TINTA)
    com_dict = ("Calculo: DICTAMEN GLOBAL. ELIJA MATERIAL (Seccion de resolucion "
               "de material) si el sleeve aun no esta seleccionado en la "
               "cascada; REVISAR — MATERIAL FUERA DE RANGO si ya hay material "
               "pero quedo fuera de rango de temperatura; APTO solo si ademas "
               "las verificaciones de espesor y longitud dan CUMPLE; REVISAR "
               "en cualquier otro caso.")
    # F60 (espesor), F61 (longitud) y F123 (luz radial <= 2,5 mm del Paso 4,
    # 206-4.1): las tres deben dar CUMPLE para APTO.
    calc("F69", (f'=IF({dictamen_sleeve}="SIN MATERIAL SELECCIONADO",'
                 '"ELIJA MATERIAL (Seccion de resolucion de material)",'
                 f'IF({dictamen_sleeve}<>"OK","REVISAR — MATERIAL FUERA DE RANGO",'
                 'IF(AND(F60="CUMPLE",F61="CUMPLE",F123="CUMPLE"),"APTO","REVISAR")))'),
        com_dict)

    # === ANEXO — PASOS DEL FLUJO 206 (bloques nuevos, direcciones estables) ==
    # El 206 no tiene oracle, pero sus formulas internas usan referencias
    # absolutas (D45->D20, D53->D45/D46/D37/D38, F60->D55...). Para no
    # desplazarlas, los pasos del flujo que no cabian en las secciones de arriba
    # se anaden aqui, a partir de la fila 101 (la Seccion 7 termina en la 98).
    # Trazado a resources/ (art_206.json) y al flujo aprobado del ingeniero.

    # --- PASO 1 — Clasificacion y seleccion guiada de tipo (206-1 / 206-2) ---
    # El tipo (D22) sigue siendo decision del ingeniero; el motor RECOMIENDA y
    # AVISA si D22 contradice la recomendacion (advisory, NO bloquea F69: el
    # tipo es una decision de diseno, no una compuerta de elegibilidad). Fuente:
    # 206-1.1.1/1.1.2 [5],[6]; 206-2.3 [13]; 206-2.5 [17]; 206-2.6 [19];
    # 206-2.7 [21].
    band(101, "PASO 1 · CLASIFICACION Y SELECCION GUIADA DE TIPO — 206-1 / 206-2")
    header(102, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])

    com103 = ("Entrada: ¿el defecto fuga o puede llegar a fugar? Si es asi, el "
              "206-1.1.2 pide Type B (extremos soldados, contiene presion). El "
              "206-2.3 exige aislar una fuga activa antes de soldar.")
    lab(103, "¿Fuga o puede fugar?", None, "206-1.1.2 / 206-2.3", com=com103)
    inp("D103", "No", com103)
    dv_list(ws, "D103", '"Si,No"', com103)

    com104 = ("Entrada: ¿la reparacion debe reforzar la resistencia AXIAL del tubo "
              "(p.ej. junta girth defectuosa) o la tasa de dano NO esta clara? Si "
              "es asi, 206-1.1.1/206-2.5 empujan a Type B (el Type A no resiste "
              "cargas axiales ni danos evolutivos).")
    lab(104, "¿Refuerzo axial o tasa de dano no clara?", None,
        "206-1.1.1 / 206-2.5", com=com104)
    inp("D104", "No", com104)
    dv_list(ws, "D104", '"Si,No"', com104)

    com105 = ("Calculo: tipo recomendado por 206-1.1: Type B si hay fuga (D103) o "
              "refuerzo axial / tasa no clara (D104); Type A en otro caso. Es una "
              "recomendacion; el ingeniero decide en D22.")
    lab(105, "Tipo recomendado", None, "206-1.1.1 / 206-1.1.2", com=com105)
    calc("D105", f'=IF(OR($D$103="Si",$D$104="Si"),"{TIPO_B}","{TIPO_A}")', com105)

    com106 = ("Calculo: aviso si el tipo elegido (D22) difiere del recomendado. "
              "No bloquea: el tipo es decision de diseno del ingeniero.")
    lab(106, "Aviso de contradiccion de tipo", None, "206-1.1", com=com106)
    calc("D106",
         '=IF($D$22<>$D$105,"AVISO: el tipo elegido (D22) difiere del recomendado '
         'por 206-1.1 (fuga/axial) — confirmar","")', com106)

    com107 = ("Calculo: 206-2.6 — en Type A, evaluar la corrosion bajo manga por "
              "ingreso de humedad por los extremos no soldados; aplicar sellante/"
              "recubrimiento si aplica.")
    lab(107, "Corrosion bajo manga (Type A)", None, "206-2.6", com=com107)
    calc("D107",
         f'=IF($D$22="{TIPO_A}","206-2.6: evaluar corrosion bajo manga; '
         'sellante/recubrimiento si aplica","")', com107)

    com108 = ("Aviso fijo 206-2.7: una costura previa (girth/longitudinal) "
              "prominente puede impedir el fit-up; esmerilar + RT/UT, o fabricar "
              "la manga con bulge (Fig. 206-2.7-1).")
    lab(108, "Interferencia con costura previa", None, "206-2.7", com=com108)
    d108 = calc("D108",
                "206-2.7: costura previa prominente puede impedir el fit-up -> "
                "esmerilar + RT/UT o manga con bulge (Fig. 206-2.7-1).", com108)
    d108.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D108:G108")

    com109 = ("Calculo: 206-2.3 — si hay fuga activa en un Type B, aislar la fuga "
              "antes de soldar (ligado a 206-4.3, purga de N2 en fluidos "
              "inflamables).")
    lab(109, "Fuga activa (Type B)", None, "206-2.3", com=com109)
    calc("D109",
         f'=IF(AND($D$22="{TIPO_B}",$D$103="Si"),"206-2.3: aislar la fuga antes '
         'de soldar (ver 206-4.3 purga N2)","")', com109)

    # --- PASO 2 — Espesor requerido con sobreespesor de corrosion (206-3.3) --
    # El t_req Type B (D53/E53/F53) ya suma C.A. (arriba). Aqui va la ENTRADA
    # C.A. (206-3.3 [38]: "corrosion allowances ... in accordance with the
    # engineering design"). Default 0 (valor valido tecleado). No aplica al
    # Type A (206-3.1, no es componente a presion).
    band(111, "PASO 2 · ESPESOR REQUERIDO CON SOBREESPESOR DE CORROSION — 206-3.3")
    header(112, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com113 = ("Entrada: sobreespesor de corrosion C.A., mm, por el diseno de "
              "ingenieria (206-3.3). Se suma al t_req Type B (D78/E78/F53) en las "
              "tres columnas de presion; NO se aplica al Type A (206-3.1). Default "
              "0 es un valor valido tecleado.")
    lab(113, "Sobreespesor de corrosion C.A.", "mm", "206-3.3 [38]", com=com113)
    inp("D113", 0, com113)
    com114 = ("Nota: el Type B es componente a presion (206-3.2), por eso su t_req "
              "incluye C.A.; el Type A (206-3.1, 2/3 del espesor del portador) es "
              "refuerzo, no contiene presion, y no lo lleva.")
    lab(114, "Aplicabilidad del C.A.", None, "206-3.1 / 3.2", com=com114)
    d114 = calc("D114",
                "C.A. se suma al t_req del Type B (a presion); el Type A (2/3 Tp) "
                "no lo lleva.", com114)
    d114.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D114:G114")

    # --- PASO 3 — Dimensiones del sleeve (206-3.4) --------------------------
    # La verificacion de longitud (F61: L_s >= max(100, defecto + 2x50)) ya vive
    # en la seccion de verificaciones y es correcta. Aqui se traza explicitamente
    # y se anota el sobrepaso de 50 mm a cada lado (206-3.4 [40]).
    band(116, "PASO 3 · DIMENSIONES DEL SLEEVE — 206-3.4")
    com117 = ("Nota 206-3.4: el sleeve mide >= 100 mm (4 in.) y sobrepasa el "
              "defecto >= 50 mm (2 in.) a CADA lado. La verificacion de longitud "
              "(F86) usa L_s,min = max(100, longitud del defecto + 2x50).")
    lab(117, "Longitud y sobrepaso", None, "206-3.4 [40]", com=com117)
    d117 = calc("D117",
                "206-3.4: L_s >= 100 mm y >= defecto + 50 mm a cada lado. "
                "Verificado en F61.", com117)
    d117.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D117:G117")

    # --- PASO 4 — Cateto del filete de extremo y luz radial (206-3.5/206-4.1) -
    # El motor deja de solo elegir la RAMA (texto en F64) y calcula el CATETO w:
    # w = Ts + G (Ts <= 1.4 Tp) o 1.4 Tp + G (Ts > 1.4 Tp), Figs. 206-3.5-1/2
    # (fijadas en la Fase 0). Solo Type B (el Type A no lleva soldadura
    # circunferencial de extremo, 206-1.1.1). Ademas la luz radial G <= 2.5 mm
    # (206-4.1 [63]), que entra al AND de F69. Ts=D29, Tp=D21 (nominal), G=D31.
    band(119, "PASO 4 · CATETO DEL FILETE Y LUZ RADIAL — 206-3.5 / 206-4.1")
    header(120, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com121 = ("Calculo: cateto del filete de extremo w (Figs. 206-3.5-1/2, Fase 0). "
              "Type B: w = Ts + G si Ts <= 1.4 Tp, si no 1.4 Tp + G (chaflan "
              "opcional). Type A: no aplica (206-1.1.1, sin soldadura "
              "circunferencial). Ts = espesor del sleeve (D29), Tp = espesor del "
              "portador (D21), G = luz radial (D31).")
    lab(121, "Cateto del filete de extremo  w", "mm", "206-3.5 Figs. 1/2",
        com=com121)
    calc("D121",
         f'=IF($D$22="{TIPO_B}",IF($D$29<=1.4*$D$21,$D$29+$D$31,'
         f'1.4*$D$21+$D$31),"No aplica - Type A (206-1.1.1)")', com121)
    com122 = ("Nota (riesgo declarado): la leyenda de las Figs. 206-3.5 rotula "
              "Tp = 'carrier pipe required minimum wall thickness', mientras que "
              "el texto 206-3.5(a) compara Ts con 1.4x el espesor NOMINAL. Este "
              "motor usa el NOMINAL (D21, coherente con el texto de la rama). Si "
              "el diseno exige el minimo requerido, ajustar D21 o el cateto. "
              "Confirmar con el ingeniero.")
    lab(122, "Tp nominal vs mínimo requerido", None, "206-3.5 leyenda", com=com122)
    d122 = calc("D122",
                "Tp = NOMINAL (D21) para el cateto y la rama 1.4x, coherente con "
                "el texto 206-3.5(a). La leyenda de las figuras dice 'required "
                "minimum': confirmar con el ingeniero si aplica.", com122)
    d122.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D122:G122")
    com123 = ("Calculo: 206-4.1 — se admite 'no gap', y una luz radial de hasta "
              "2.5 mm (3/32 in.) maximo. G adoptado (D31) contra ese tope; entra "
              "al dictamen global (F94).")
    lab(123, "Luz radial G (206-4.1)", None, "206-4.1 [63]", com=com123)
    ws.merge_cells("A123:C123")
    calc("D123", f'={umbral(UMBR, "206_luz_radial", ES_SI_206)}',
         "Calculo: luz radial maxima admitida, leida del codigo en el sistema "
         "activo — 206-4.1 la imprime como 2.5 mm (3/32 in.) y no se convierte de "
         "una unidad a la otra (regla 9).")
    calc("E123", "=$D$31",
         "Calculo: repite la luz radial medida en la seccion 1 (D31), que es la que "
         "se compara contra el maximo.")
    calc("F123", '=IF(E123<=D123,"CUMPLE","NO CUMPLE — excede la luz maxima (206-4.1)")',
         "Calculo: veredicto de la luz radial. 206-4.1 pide en general un ajuste "
         "'sin luz' y admite el maximo de D123; por encima, el codigo advierte que "
         "pueden requerirse ajustes de tamano de soldadura y de tecnica. Entra al "
         "AND del dictamen global.")

    # --- PASO 5 — Presion externa, cavidades y bulging (206-3.6/3.7/3.9/3.10) -
    band(125, "PASO 5 · PRESION EXTERNA, CAVIDADES Y BULGING — 206-3.6/3.7/3.9/3.10")
    header(126, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com127 = ("Entrada: ¿el defecto es externo / hay perdida de pared externa? Si "
              "es asi, 206-3.7/3.9 piden rellenar las cavidades con material "
              "endurecible (epoxi) de resistencia a compresion adecuada para "
              "transferir la carga al sleeve.")
    lab(127, "¿Defecto externo?", None, "206-3.7 / 3.9", com=com127)
    inp("D127", "No", com127)
    dv_list(ws, "D127", '"Si,No"', com127)
    com128 = "Calculo: aviso de relleno endurecible si el defecto es externo."
    lab(128, "Relleno de cavidades", None, "206-3.7 / 3.9", com=com128)
    calc("D128",
         '=IF($D$127="Si","206-3.7/3.9: rellenar cavidades con material '
         'endurecible (epoxi) de resistencia a compresion adecuada","")', com128)
    com129 = ("Aviso fijo 206-3.6: considerar la presion externa sobre el tubo "
              "dentro del Type B; ajuste ceñido para transferir carga o rellenar "
              "el anular; si queda sin rellenar, verificar que el fluido estancado "
              "no cause corrosion.")
    lab(129, "Presion externa (Type B)", None, "206-3.6", com=com129)
    d129 = calc("D129",
                "206-3.6: presion externa sobre el tubo en Type B — ajuste ceñido "
                "o rellenar anular; verificar fluido estancado.", com129)
    d129.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D129:G129")
    com130 = ("Aviso fijo 206-3.9(b): reducir la presion de linea al instalar (se "
              "cruza con el rango 50-80% del Paso 7).")
    lab(130, "Reducir presion al instalar", None, "206-3.9(b)", com=com130)
    d130 = calc("D130",
                "206-3.9(b): reducir la presion de linea al instalar (ver rango "
                "50-80% del Paso 7).", com130)
    d130.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D130:G130")

    # --- PASO 6 — Fatiga y dilatacion diferencial (206-2.4/3.8/3.11) --------
    band(132, "PASO 6 · FATIGA Y DILATACION DIFERENCIAL — 206-2.4/3.8/3.11")
    header(133, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com134 = ("Entrada: ¿servicio con ciclos frecuentes de presion o gradientes "
              "termicos through-wall? Todo Type B se evalua a fatiga (206-3.8); "
              "los ciclos frecuentes la exigen (206-2.4).")
    lab(134, "¿Servicio ciclico?", None, "206-2.4 / 3.8", com=com134)
    inp("D134", "No", com134)
    dv_list(ws, "D134", '"Si,No"', com134)
    com135 = "Calculo: aviso de evaluacion de fatiga si el servicio es ciclico."
    lab(135, "Evaluacion de fatiga", None, "206-2.4 / 3.8", com=com135)
    calc("D135",
         '=IF($D$134="Si","206-2.4/3.8: requiere evaluacion de fatiga '
         '(VIII-2 / API 579-1/ASME FFS-1)","")', com135)
    com136 = ("Aviso fijo 206-3.11: considerar la dilatacion termica diferencial "
              "entre el tubo portador y el sleeve (ambos tipos). Refuerza nota206b.")
    lab(136, "Dilatacion diferencial", None, "206-3.11", com=com136)
    d136 = calc("D136",
                "206-3.11: considerar la dilatacion termica diferencial "
                "portador/sleeve (ambos tipos).", com136)
    d136.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D136:G136")

    # --- PASO 7 — Fabricacion y soldadura en servicio (206-4) ---------------
    band(138, "PASO 7 · FABRICACION Y SOLDADURA EN SERVICIO — 206-4")
    header(139, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com140 = ("Calculo: presion recomendada durante la instalacion del sleeve, "
              "entre 50% y 80% de la presion de operacion (206-4.5; API RP 2201). "
              "Minimo (50%).")
    lab(140, "Presion de instalacion (min 50%)", "kg/cm2", "206-4.5", com=com140)
    calc("D140", "=0.5*$D$26", com140)
    com141 = "Calculo: maximo del rango de presion de instalacion (80%), 206-4.5."
    lab(141, "Presion de instalacion (max 80%)", "kg/cm2", "206-4.5", com=com141)
    calc("D141", "=0.8*$D$26", com141)
    com142 = ("Calculo: si hay fuga (D103), purgar el anular con N2/gas inerte en "
              "fluidos inflamables antes de soldar (206-4.3).")
    lab(142, "Purga de N2 (fuga)", None, "206-4.3", com=com142)
    calc("D142",
         '=IF($D$103="Si","206-4.3: purgar el anular con N2/gas inerte en fluidos '
         'inflamables","")', com142)
    com143 = ("Aviso fijo de fabricacion: limpiar a metal blanco toda la "
              "circunferencia (206-4.1); el relleno no debe extruir a la soldadura "
              "(206-4.2); soldadura en servicio por Art. 210 (H2 en ZAC, ZAC dura, "
              "burn-through) (206-4.6); costuras longitudinales a tope penetracion "
              "completa + venteo si hay filetes de cierre (206-4.4).")
    lab(143, "Fabricacion y Art. 210", None, "206-4.1/4.2/4.4/4.6", com=com143)
    d143 = calc("D143",
                "206-4: metal blanco (4.1); relleno sin extruir (4.2); Art. 210 "
                "H2/ZAC/burn-through (4.6); longitudinales a tope + venteo (4.4).",
                com143)
    d143.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D143:G143")

    # --- PASO 8 — Examen (NDE) y prueba de hermeticidad (206-5 / 206-6) ------
    # Decision 3 del ingeniero: minima y fiel al 206 — selector + notas, SIN
    # ecuaciones de energia almacenada (el texto del 206 no las publica; no se
    # importa el motor neumatico del App. 501 del 212). El selector no gobierna
    # ningun calculo: solo materializa el requisito y la advertencia de presion.
    band(145, "PASO 8 · EXAMEN NO DESTRUCTIVO Y PRUEBA DE HERMETICIDAD — 206-5 / 206-6")
    header(146, ["Parametro", "", "Unidad", "Valor", "", "", "Referencia / Notas"])
    com147 = ("Entrada: tipo de prueba de hermeticidad del Type B (206-6, si el "
              "propietario la requiere): prueba del anular presurizado, prueba "
              "sensible de fugas (B31.3 345.8) o no requerida. Art. 501 da guia "
              "adicional. SIN ecuaciones de energia (el 206 no las publica).")
    lab(147, "Tipo de prueba de hermeticidad", None, "206-6", com=com147)
    inp("D147", "Prueba del anular", com147)
    dv_list(ws, "D147",
            '"Prueba del anular,Prueba sensible de fugas,No requerida"', com147)
    com148 = ("Calculo: aviso segun el tipo de prueba. En la del anular, la "
              "presion se elige tal que el tubo interno NO colapse (206-6a); "
              "remite al Art. 501.")
    lab(148, "Aviso de prueba", None, "206-6", com=com148)
    calc("D148",
         '=IF($D$147="Prueba del anular","206-6(a): presion de prueba tal que el '
         'tubo interno NO colapse; Art. 501 guia adicional",IF($D$147="Prueba '
         'sensible de fugas","206-6(b): prueba sensible de fugas (B31.3 345.8)",'
         '""))', com148)
    ws.merge_cells("D148:G148")
    com149 = ("Calculo: NDE por tipo. Type B (206-5.3): UT del portador; primer/"
              "ultimo pase MT/PT; NDE de las circunferenciales >= 24 h (>= 48 h si "
              "servicio con alta probabilidad de H2). Type A (206-5.2): VT de raiz "
              "+ PT/MT/UT de las longitudinales.")
    lab(149, "NDE por tipo", None, "206-5.2 / 5.3", com=com149)
    calc("D149",
         f'=IF($D$22="{TIPO_B}","206-5.3: UT del portador; primer/ultimo pase '
         'MT/PT; NDE de circunferenciales >=24 h (>=48 h si servicio con H2)",'
         '"206-5.2: Type A - VT de raiz + PT/MT/UT de longitudinales")', com149)
    ws.merge_cells("D149:G149")
    com150 = ("Aviso fijo 206-5.1: inspeccionar todos los fit-ups antes de soldar "
              "y examinar visualmente todas las soldaduras.")
    lab(150, "Examen visual (VT)", None, "206-5.1", com=com150)
    d150 = calc("D150",
                "206-5.1: inspeccionar todos los fit-ups antes de soldar; VT de "
                "todas las soldaduras.", com150)
    d150.font = Font(name=MONO, size=9, color=GRIS)
    ws.merge_cells("D150:G150")

    # Fase 9: subindices reales en los simbolos. En el 206 los simbolos (T_s,
    # L_s, T_p...) van embebidos en la prosa de la columna A, no en una columna
    # B propia como en el 212; el mismo paso los convierte token a token.
    _aplicar_subindices(ws)

    # Fase 3: la Seccion de Material sube justo detras de la Seccion 1, con el
    # mismo mecanismo que el 212 — ver _mapa_filas_206().
    remapear_filas(ws, MAPA_FILAS_206)
    aplicar_unidades_motor(ws, UNIDADES_206, ES_SI_206)
    aplicar_reglas_de_comentario(ws, REGLAS_COMENTARIO_206)

    # Misma leyenda de color que el Art. 212: es la misma convencion de edicion
    # y tiene que verse igual en los dos motores (ver LEYENDA_MOTOR).
    build_leyenda_motor(ws)
    aplicar_leyenda_motor(ws)
    aplicar_semaforo_motor(ws, SEMAFORO_206, FILA_DICTAMEN_206)

    # Fase 8: mismo manifiesto y mismo boton que el 212 (ver build_reinicio_motor).
    build_reinicio_motor(ws)

    # Fase 9: el 206 no tenia especificaciones tecnicas; ahora las tiene, en
    # pestana propia y construidas desde cero contra resources/.
    build_botones_documentos(ws, espec=ESPEC_206, instr=INSTR_206)

    preparar_impresion(ws, MOTOR_NCOLS)

    ws.protection.password = "0000"
    ws.protection.sheet = True


def retirar_datos_ref(wb):
    """Borra la hoja heredada Datos_Ref, que el maestro Rev0 aun trae.

    Tarea 10: ninguna hoja de datos viene ya del maestro. El esfuerzo admisible
    lo dan DB_B31_3 / DB_BPVC_IID y las dimensiones DB_B36_10/19, todas auditadas
    contra resources/ (regla 1). El bloque obsoleto de esfuerzos que vivia en las
    filas 40-47 se va con la hoja; el libro anterior queda en el historial de git.
    Mismo patron defensivo que build_parche_art212 usa con su propia hoja."""
    if "Datos_Ref" in wb.sheetnames:   # el maestro Rev0 todavia la trae
        del wb["Datos_Ref"]


# ---------------------------------------------------------------------------
# Fase 9 — Especificaciones tecnicas, una pestana por motor
# ---------------------------------------------------------------------------
# En el 212 esto vivia dentro de la hoja del motor, en siete filas de parrafo
# fusionadas B:G; en el 206 no existia. Sale a pestana propia por dos razones:
#
#   1. No se consulta mientras se calcula. Se lee cuando el calculo ya esta
#      cerrado y hay que redactar el procedimiento de la reparacion, asi que
#      ocupaba con el bloque mas denso de la hoja el sitio por el que se pasa
#      en cada iteracion del diseno.
#   2. La cita no cabia. Cada especificacion sale de un parrafo concreto de
#      PCC-2 y ahi no habia donde ponerlo: el texto se generaba correcto pero
#      sin decir de donde venia, que es justo lo que la Regla n.1 pide que se
#      pueda auditar. Aqui cada fila lleva su clausula en columna propia.
#
# ESTAS HOJAS SON PARTE DE SU MOTOR, no un segundo motor, y por eso leen sus
# celdas (`Parche_PCC2_Art212!$D$30`). La regla 13 —un motor no lee otro— separa
# motores de BUSQUEDA de motores de CALCULO para que ninguno dependa del estado
# de otro; aqui hay un solo motor repartido en dos pestanas del mismo articulo, y
# lo contrario seria peor: una especificacion que dijera un espesor distinto del
# que se calculo.
#
# El texto de cada fila TRANSCRIBE el requisito de su parrafo, con las unidades
# como el codigo las imprime —«5 mm (3/16 in.)»— y no convertidas (regla 9). Por
# eso estas hojas no llevan conmutador SI/US propio: lo que publica el codigo en
# las dos unidades a la vez no se conmuta. Lo que SI depende del sistema son las
# cifras que vienen del motor, y esas se rotulan con el selector del motor.
ESPEC_212, ESPEC_206 = "Espec_PCC2_Art212", "Espec_PCC2_Art206"
ESPEC_NCOLS = 7
ESPEC_ANCHO = {"A": 34, "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 30}
# Boton de ida y vuelta entre el motor y sus documentos. Van en H1:J1, H2:J2 y
# H3:J3 —las tres filas congeladas, a la derecha de la tabla— y NO en A..G: esa
# banda la compara celda a celda el *oracle* del 212, y las columnas K..AB de un
# motor estan ocultas (ahi viven las listas de cascada materializadas), asi que
# H, I y J son el unico sitio visible que queda fuera de la tabla.
COL_BTN_DOC_1, COL_BTN_DOC_2 = 8, 10
FILA_BTN_ESPEC, FILA_BTN_INSTR = 1, 2
TXT_BTN_ESPEC = "[ > ] ESPECIFICACIONES TECNICAS"
TXT_BTN_INSTR = "[ ? ] INSTRUCCIONES DE ESTE MOTOR"
TXT_BTN_MOTOR = "[ < ] IR AL MOTOR DE CALCULO"

AVISO_ESPEC = (
    "Herramienta de ingenieria de referencia. Cada fila transcribe el requisito "
    "del parrafo de ASME PCC-2 que cita a su derecha; verificar contra el codigo "
    "vigente y complementar con WPS/PQR, ATS/JSA y los registros del propietario. "
    "El criterio de aceptacion es el del codigo de construccion o "
    "post-construccion aplicable.")


def _espec_fila(ws, r, concepto, texto, cita, editable=False, com=None):
    """Una fila de especificacion: concepto | texto (B:F) | clausula citada."""
    a = ws.cell(r, 1, concepto)
    a.font = Font(name=MONO, size=10, bold=True, color=TINTA)
    a.alignment = Alignment(vertical="top", wrap_text=True, indent=1)

    v = _mrg(ws, r, 2, 6, texto)
    v.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    if editable:
        # AMARILLO esta acotado por nombre de hoja a los dos motores, asi que
        # aqui el campo editable se distingue como en un buscador: papel limpio
        # con la linea inferior de tinta. El relleno hay que ponerlo en toda la
        # cola del fusionado (el BORDE no se hereda del ancla; ver franja()).
        v.font = IN_F
        v.protection = Protection(locked=False)
        for c in range(2, 7):
            ws.cell(r, c).border = CAJA_CAMPO
    else:
        v.font = DATA_F

    g = ws.cell(r, 7, cita)
    g.font = SRC_F
    g.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    if com:
        _nota(v, com)

    # Excel no autoajusta el alto de una fila fusionada, asi que se calcula: B:F
    # suma 100 unidades de ancho y la mono de 9 pt entra a ~95 caracteres por
    # linea. Una fila corta de mas se lee; una corta de menos oculta el texto.
    # 100 caracteres por linea en B:F, medidos sobre el PDF exportado. Y en una
    # celda de FORMULA no se mide el fuente sino lo que se vera: los literales
    # entre comillas mas un hueco por cada TEXT(). Medir la formula entera daba
    # una fila de cinco lineas para un texto que ocupa dos.
    visible = texto if not str(texto).startswith("=") else (
        "".join(re.findall(r'"([^"]*)"', str(texto)))
        + " " * 10 * str(texto).count("TEXT("))
    lineas = max(1,
                 -(-len(str(visible)) // 100),
                 -(-len(str(cita)) // 30),
                 -(-len(str(concepto)) // 32))
    ws.row_dimensions[r].height = 13.5 * lineas + 4
    return r + 1


def build_especificaciones(wb, nombre, motor, titulo, subtitulo, bloques):
    """Hoja de especificaciones tecnicas de un motor de calculo.

    `bloques` es ((rotulo de banda, (filas,)), ...) y cada fila es
    (concepto, texto, clausula) o (concepto, texto, clausula, "EDITABLE").
    """
    if nombre in wb.sheetnames:
        del wb[nombre]
    ws = new_sheet(wb, nombre, titulo, subtitulo)
    autosize(ws, ESPEC_ANCHO)
    ws.merge_cells(f"A1:{get_column_letter(ESPEC_NCOLS)}1")
    ws.merge_cells(f"A2:{get_column_letter(ESPEC_NCOLS)}2")
    # La fila 3 se deja libre: ahi ancla link_volver() el boton de retorno de
    # toda hoja destino (ANCLA_VOLVER), y en H3:J3 va el atajo al motor.
    _boton(ws, 3, COL_BTN_DOC_1, COL_BTN_DOC_2, TXT_BTN_MOTOR, motor)
    _ocultar_columnas_clave(ws, (COL_BTN_DOC_1,))

    r = 5
    for rotulo_banda, filas in bloques:
        banda(ws, r, rotulo_banda, ESPEC_NCOLS)
        r += 1
        for j, h in ((1, "Concepto"), (2, "Especificación"), (7, "Cláusula citada")):
            c = ws.cell(r, j, h)
            c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
        for j in range(3, 7):          # la cabecera cierra a todo el ancho
            c = ws.cell(r, j)
            c.fill, c.border = HDR_FILL, BOX_FRANJA
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        r += 1
        for fila in filas:
            concepto, texto, cita = fila[:3]
            r = _espec_fila(ws, r, concepto, texto, cita,
                            editable=(len(fila) > 3 and fila[3] == "EDITABLE"))
        r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ESPEC_NCOLS)
    av = ws.cell(r, 1, AVISO_ESPEC)
    av.font = SRC_F
    av.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 30

    preparar_impresion(ws, ESPEC_NCOLS)
    ws.protection.password = "0000"
    ws.protection.sheet = True
    return ws


def _ref_motor(hoja, mapa, col, fila):
    """Referencia absoluta a una celda del motor, con la fila ya remapeada.

    Las filas se dan como las escribe el builder (ANTES del mapa de la Fase 3) y
    se traducen aqui con el mismo mapa que movio la hoja: escribir la direccion
    final a mano seria una copia del mapa que nadie volveria a revisar.
    """
    return f"{hoja}!${col}${mapa.get(fila, fila)}"


def _rot_unidad(hoja, celda_sel, clase):
    """Rotulo de unidad que sigue al selector SI/US del motor (Fase 7)."""
    si, us = _U[clase]
    return f'IF({hoja}!{celda_sel}="SI","{si}","{us}")'


def build_especificaciones_art212(wb):
    """Especificaciones tecnicas del Art. 212, citando 212-3.4, 212-4, 212-5,
    212-6 y el Art. 210. Todo el texto transcribe el parrafo que cita a su
    derecha, leido de resources/.../art_212.json (Regla n.1)."""
    d = lambda col, fila: _ref_motor(MOTOR, MAPA_FILAS_212, col, fila)
    ul = _rot_unidad(MOTOR, UNIDAD_212, "len")
    up = _rot_unidad(MOTOR, UNIDAD_212, "peso")
    upr = _rot_unidad(MOTOR, UNIDAD_212, "pres")

    # La plancha se identifica con el material_id que resolvio la cascada de la
    # Seccion 2 (columna del collar/parche), NO con la fila descriptiva de lista
    # fija: esa se retiro en la Tarea 7-8 y la formula heredada seguia
    # apuntandola, asi que imprimia "Plancha  de 8 mm" con el hueco en medio.
    # La fila "material_id resuelto" es F+10 del bloque de material, que el
    # builder escribe en la fila 105.
    mat = d("E", 105 + 10)
    metodo = (
        '="Reparación por parche de plancha con soldadura de filete perimetral '
        f'(ASME PCC-2, Art. 212). Aplicación: "&{d("D", 10)}&". Código de '
        f'construcción: "&{d("D", 12)}&". Material del parche: "&IF({mat}="","'
        f'(sin resolver en la Sección 2)",{mat})&". Espesor "&TEXT({d("D", 29)},'
        f'"0.0")&" "&{ul}&"; peso estimado "&TEXT({d("D", 80)},"0.0")&" "&{up}&"."')
    filetes = (
        '="Filete perimetral de cateto "&TEXT(' + d("D", 32) + ',"0.0")&" "&' + ul +
        '&" sobre metal sano, con solape mínimo de "&TEXT(' + d("D", 33) + ',"0")&'
        '" "&' + ul + '&" por extremo. El cateto se dimensiona con la ec. (4) del '
        '212-3.4(a), F_A = E·S_a·w_mín con E = 0,55, y la carga excéntrica con la '
        'ec. (5) del 212-3.4(c): las dos las verifica la Sección 7 del motor."')
    prueba = (
        '="Prueba de fuga del componente y del parche instalado conforme al código '
        'post-construcción aplicable (212-6b). Si es hidrostática, presión de '
        'prueba "&TEXT(' + d("D", 46) + "*" + d("D", 28) + ',"0.0")&" "&' + upr +
        '&" ("&TEXT(' + d("D", 46) + ',"0.0")&"× la presión de diseño). Se toman '
        'precauciones especiales de seguridad si la prueba es neumática."')

    bloques = (
        ("METODO Y ALCANCE  //  ASME PCC-2 Art. 212-1 / 212-2", (
            ("Método de reparación", metodo, "212-1(a) a (d) · se redacta con las "
             "entradas del motor"),
            ("Alcance del método", "El método cubre la selección, limitaciones de "
             "aplicación, diseño, fabricación, examen y prueba de parches de "
             "superficie soldados con filete a componentes que retienen presión. "
             "Se aplica típicamente a envolventes con adelgazamiento local de "
             "pared (incluido el traspasante) por erosión, corrosión y otros "
             "mecanismos locales, y es aplicable a envolventes cilíndricas, "
             "esféricas, planas y cónicas, además de otros componentes a presión.",
             "212-1(a) · 212-1(c) · 212-1(d)"),
            ("Limitaciones", "No se usa el método si el mecanismo de daño, su "
             "extensión o el daño futuro no se pueden caracterizar. Una grieta o "
             "defecto tipo grieta solo se admite si el crecimiento está detenido y "
             "hay evaluación de aptitud para el servicio conforme a "
             "API 579-1/ASME FFS-1. La compuerta de elegibilidad la resuelve el "
             "Paso 1 del anexo del motor, y un estado bloqueante detiene el "
             "dictamen global.", "212-2(c) · 212-2 [11] a [13]"),
        )),
        ("PREPARACION Y AJUSTE  //  212-4 FABRICATION", (
            ("Corte y preparación de bordes", "Los bordes de la plancha pueden "
             "cortarse por medios mecánicos (mecanizado, cizallado, esmerilado) o "
             "térmicos (oxicorte o corte por arco). Si se usan medios térmicos, se "
             "retira por esmerilado o mecanizado un mínimo de 1,5 mm (1/16 in.) de "
             "material adicional. Si la plancha es de más de 25 mm (1 in.) de "
             "espesor y el cateto del filete es menor que el espesor de la plancha, "
             "los bordes preparados se examinan por partículas magnéticas (MT) o "
             "líquidos penetrantes (PT) para detectar laminaciones; una laminación "
             "es causa de rechazo salvo que se repare o se acepte por evaluación de "
             "aptitud para el servicio según API 579-1/ASME FFS-1.", "212-4(a)"),
            ("Conformado y secciones partidas", "La plancha puede conformarse a la "
             "geometría requerida por cualquier proceso que no deteriore "
             "indebidamente sus propiedades mecánicas. Cuando el tamaño de la "
             "plancha o el acceso lo exijan, se admiten secciones partidas unidas "
             "por soldaduras de penetración completa. El límite de deformación por "
             "conformado en frío lo verifica el motor con el 212-3.5.",
             "212-4(b) · 212-3.5"),
            ("Ajuste (fit-up)", "Las partes a unir con filete se ajustan tan "
             "apretadas como sea practicable y en ningún caso con separación mayor "
             "de 5 mm (3/16 in.). Si la separación en el borde de apoyo de la "
             "plancha es de 1,5 mm (1/16 in.) o mayor, el tamaño del filete "
             "perimetral se recalcula sumando esa separación a la excentricidad e "
             "— es exactamente lo que hace el campo de separación g del motor.",
             "212-4(c) · 212-4c [82]"),
            ("Limpieza previa", "Se retiran pintura, cascarilla, óxido, líquidos y "
             "materia extraña de la zona de soldadura y de una franja no menor de "
             "40 mm (1 1/2 in.) a cada lado del cordón. En las áreas que quedarán "
             "cubiertas por la plancha, las costuras longitudinales o "
             "circunferenciales existentes deberían esmerilarse a ras del diámetro "
             "exterior y examinarse por MT o PT.", "212-4(e)(1) y (2)"),
            ("Secuencia de montaje", "La plancha se posiciona por cualquier método "
             "adecuado. Las costuras internas de la propia plancha se ejecutan "
             "primero y después se completa el cordón perimetral; se admiten grapas "
             "o cuñas para asegurar alineación y ajuste.",
             "212-4(e)(3) y (4)"),
            ("Venteo", "Para impedir la acumulación de presión de gas entre la "
             "plancha y la frontera de presión puede ser necesario ventear durante "
             "el cordón de cierre final y, si aplica, durante el tratamiento "
             "térmico post-soldadura. Si la plancha se diseñó para un defecto "
             "traspasante pero se instala antes de que la pared se perfore, el "
             "venteo se sella al terminar la soldadura y el PWHT.", "212-4(g)"),
        )),
        ("SOLDADURA  //  212-3.4 · 212-4(d) · Art. 210", (
            ("Filetes perimetrales", filetes,
             "212-3.4(a) ec. (4) · 212-3.4(c) ec. (5)"),
            ("Tope máximo del filete", "El tamaño de diseño del filete no excede el "
             "espesor del más delgado de los materiales unidos ni 40 mm (1,5 in.). "
             "Alternativamente el borde del cordón perimetral puede biselarse para "
             "aumentar la garganta efectiva, sin que esa garganta supere el espesor "
             "nominal de la plancha de reparación ni el espesor nominal original "
             "del componente. Los dos topes los verifica el Paso 4 del anexo del "
             "motor.", "212-3.4 NOTA · 212-3.4(b)"),
            ("Juntas de cierre", "Juntas a tope en V, penetración completa (C.J.P), "
             "raíz abierta sin respaldo. Raíz GTAW ER70S-6; relleno y peine SMAW "
             "E7018 bajo hidrógeno. Ángulo incluido 60°–75°, talón 1,5 mm, luz de "
             "raíz 2–3 mm. Ajustar al WPS calificado del proyecto.",
             "212-4(d) · WPS del proyecto", "EDITABLE"),
            ("Calificación de procedimientos y soldadores", "Los procedimientos, "
             "soldadores y operadores se califican conforme a los requisitos "
             "vigentes del código de construcción o post-construcción aplicable. Si "
             "no se especifica otra cosa, puede usarse ASME BPVC Sección IX para "
             "calificación de procedimiento y de desempeño. Para soldadura en "
             "servicio se consulta el Art. 210 y para tratamiento térmico en campo "
             "el Art. 214.", "212-4(d)"),
            ("Soldadura en servicio", "Electrodo bajo hidrógeno E7018 "
             "(Ø 2,4–3,2 mm); precalentamiento ≥ 100 °C; control del aporte térmico "
             "frente a perforación e hidrógeno. WPS calificado con el Apéndice "
             "Obligatorio 210-I; END diferido 24–72 h. Tramo drenado y "
             "despresurizado antes del cordón de cierre. Ajustar al procedimiento "
             "calificado del proyecto.", "212-4(d) · Art. 210", "EDITABLE"),
        )),
        ("EXAMEN  //  212-5 EXAMINATION", (
            ("Examen de las uniones del parche", "Las soldaduras de unión de la "
             "plancha se examinan conforme al código de construcción o "
             "post-construcción aplicable por MT o por PT, salvo que el método esté "
             "limitado por temperatura. Si el código no especifica procedimientos, "
             "el END se ejecuta con procedimientos escritos y calificados según "
             "ASME BPVC Sección V.", "212-5(a)"),
            ("Orejas de izaje y elementos temporales", "Si se usan orejas de izaje "
             "y se dejan en su sitio, sus soldaduras de unión se examinan por MT o "
             "PT. En toda ubicación donde se retiren orejas temporales, grapas "
             "soldadas o cuñas después de instalar la plancha, la zona de remoción "
             "se examina por MT o PT.", "212-5(b)"),
            ("Costuras entre piezas de la plancha", "Las soldaduras que unen "
             "secciones de plancha hechas de piezas separadas deberían contornearse "
             "en superficie y examinarse volumétricamente por radiografía o "
             "ultrasonido en la medida posible. Si no es practicable, se ejecutan "
             "exámenes PT o MT multicapa.", "212-5(c)"),
            ("END después del PWHT", "Si se requiere tratamiento térmico "
             "post-soldadura, el examen se realiza después de aplicarlo. El criterio "
             "de aceptación del examen es el del código de construcción o "
             "post-construcción aplicable.", "212-5(d) · 212-5(e)"),
            ("Alcance de END del proyecto", "100 % VT + 100 % PT/MT de juntas y "
             "filetes (criterio del código de construcción); UT de espesores bajo "
             "filetes y del área reparada; END diferido 24–72 h si la soldadura es "
             "en servicio (Art. 210). Ajustar al criterio de aceptación del código "
             "activo.", "212-5 · Art. 210", "EDITABLE"),
        )),
        ("PRUEBA Y CIERRE  //  212-6 TESTING · App. 501", (
            ("Prueba de hermeticidad", prueba, "212-6(a) · 212-6(b)"),
            ("Alternativas a la prueba de fuga", "Si el código post-construcción lo "
             "permite, el examen no destructivo puede ejecutarse como alternativa a "
             "la prueba de fuga. También puede hacerse una inspección de servicio "
             "inicial de todas las juntas soldadas una vez que el componente ha "
             "vuelto a su presión y temperatura normales de operación, si estas se "
             "habían reducido para soldar.", "212-6(c)"),
            ("Energía neumática", "Cuando la prueba de fuga es neumática se toman "
             "precauciones especiales de seguridad. La energía almacenada, su "
             "equivalente en TNT y la distancia mínima al personal las calcula el "
             "Paso 8 del anexo del motor con las ecuaciones del Apéndice "
             "Obligatorio 501-II/III, leídas de resources/.",
             "212-6(b) · App. 501-II / 501-III"),
            ("Orden de las actividades de cierre", "Las pruebas e inspecciones se "
             "ejecutan antes de reaplicar recubrimiento, aislamiento o forro. Las "
             "superficies metálicas expuestas deberían recubrirse de nuevo, si "
             "aplica, una vez completados todos los exámenes y pruebas.",
             "212-6(d) · 212-4(f)"),
            ("Recubrimiento", "ED-B-06.00 / PE-B-0600.01 Esquema N° 1: Sa 2½ "
             "(ISO 8501-1); imprimación epoxi-Al 70 µm + intermedia epoxi MIO "
             "110 µm + PU alifático 2×40 µm = 260 µm. Aplicación EC-B-53.00. "
             "Ajustar a la especificación del propietario.",
             "212-4(f) · especificación del propietario", "EDITABLE"),
        )),
    )
    return build_especificaciones(
        wb, ESPEC_212, MOTOR,
        "ESPECIFICACIONES TECNICAS — PARCHE SOLDADO (ASME PCC-2 Art. 212)",
        "Fabricacion, examen y prueba. Cada fila transcribe el parrafo de PCC-2 "
        "que cita a su derecha; las cifras las lee del motor Parche_PCC2_Art212.",
        bloques)


def build_especificaciones_art206(wb):
    """Especificaciones tecnicas del Art. 206 (no existian): 206-2, 206-3.5,
    206-4, 206-5 y 206-6, leidos de resources/.../art_206.json (Regla n.1)."""
    d = lambda col, fila: _ref_motor(COLLAR_MOTOR, MAPA_FILAS_206, col, fila)
    ul = _rot_unidad(COLLAR_MOTOR, UNIDAD_206, "len")

    metodo = (
        '="Collar de encierro total soldado sobre tuberia (ASME PCC-2, Art. 206). '
        f'Tipo: "&{d("D", 22)}&". Código de construcción: "&{d("D", 12)}&". Collar '
        f'de "&TEXT({d("D", 29)},"0.0")&" "&{ul}&" de espesor y "&TEXT('
        f'{d("D", 32)},"0")&" "&{ul}&" de longitud, con luz radial "&TEXT('
        f'{d("D", 31)},"0.0")&" "&{ul}&"."')
    # El cateto lo calcula el Paso del anexo (fila 121), que no se remapea.
    #
    # Esa celda NO siempre devuelve un numero: en Type A no hay cordon de cierre
    # y publica el texto "No aplica - Type A (206-1.1.1)". Un TEXT() a secas lo
    # dejaba pasar y la fila decia «w = No aplica - Type A (206-1.1.1) mm», con
    # una unidad pegada a una frase. Se distingue con ISNUMBER: el numero lleva
    # unidad, el dictamen se transcribe tal cual.
    w = d("D", 121)
    cateto = (
        '="Cateto del filete de extremo: "&IF(ISNUMBER(' + w + '),"w = "&TEXT('
        + w + ',"0.0")&" "&' + ul + ',' + w + ')&". Segun la Fig. 206-3.5-1 '
        '(w = T_s + G cuando T_s ≤ 1,4 T_p) o la Fig. 206-3.5-2 '
        '(w_máx = 1,4 T_p + G cuando T_s > 1,4 T_p, con chaflán opcional). '
        'G es la luz radial y T_p el espesor del tubo portador."')

    bloques = (
        ("METODO Y TIPO DE COLLAR  //  206-1 · 206-3.1 · 206-3.2 · 206-3.3", (
            ("Método de reparación", metodo, "206-1.1 · se redacta con las entradas "
             "del motor"),
            ("Type A frente a Type B", "El collar Type A refuerza el tubo portador "
             "y NO contiene presión: sus extremos no se sueldan al portador. El "
             "Type B sí contiene presión, lleva cordones circunferenciales de "
             "cierre en los extremos y se dimensiona con un espesor de pared igual "
             "o mayor que el requerido para la presión de diseño máxima admisible "
             "del tubo portador.", "206-1.1.1 · 206-1.1.2 · 206-3.1 · 206-3.2 · "
             "206-3.3"),
            ("Precauciones y limitaciones", "Defecto con fuga: exige Type B "
             "(206-2.3). Operación cíclica: valorar la fatiga de los cordones de "
             "cierre (206-2.4 y 206-3.8). Defecto circunferencial: el Type A no lo "
             "refuerza (206-2.5) y el motor lo advierte. Ver además corrosión bajo "
             "el collar (206-2.6), refuerzo de soldadura del portador (206-2.7), "
             "requisitos de tamaño del collar (206-2.8), soldadura (206-2.9) y "
             "material de aporte (206-2.10).", "206-2.1 a 206-2.10"),
            ("Dilatación térmica diferencial", "Si el collar y el tubo portador son "
             "de materiales con coeficientes de dilatación distintos, se considera "
             "la dilatación térmica diferencial.", "206-3.11"),
        )),
        ("INSTALACION Y AJUSTE  //  206-4.1 · 206-4.2 · 206-4.3", (
            ("Limpieza y ajuste", "Toda la circunferencia del tubo portador en la "
             "zona que cubrirá el collar se limpia a metal desnudo. Si se va a usar "
             "material de relleno endurecible, se aplica en todas las "
             "indentaciones, picaduras, huecos y depresiones. El collar se ajusta "
             "apretado alrededor del portador; puede usarse apriete mecánico con "
             "equipo hidráulico, pernos de arrastre u otros dispositivos. En "
             "general debería lograrse un ajuste «sin luz»; se admite una luz "
             "radial de hasta 2,5 mm (3/32 in.) máximo. Con collares de extremos "
             "soldados, una luz excesiva puede exigir ajustes del tamaño de "
             "soldadura y de la técnica del soldador, como pasadas de manteado.",
             "206-4.1"),
            ("Material de relleno del anular", "Si se usa relleno entre tubo y "
             "collar, se cuida que no se extruya a las zonas de soldadura: quemarlo "
             "durante la soldadura compromete la calidad del cordón. El exceso se "
             "retira antes de soldar. Bombear el relleno al anular después de "
             "soldar el collar en su sitio elimina el problema, siempre que las "
             "luces anulares sean bastante grandes para que el relleno fluya a "
             "todos los huecos.", "206-4.2 · 206-3.10"),
            ("Defecto con fuga", "En un defecto con fuga, el área del defecto se "
             "aísla antes de soldar. En líneas con contenido inflamable, el collar "
             "se purga con nitrógeno u otro gas inerte para impedir la formación de "
             "una mezcla combustible bajo el collar.", "206-4.3"),
        )),
        ("SOLDADURA  //  206-3.5 · 206-4.4 a 206-4.7 · Art. 210", (
            ("Costuras del collar y venteo", "Si se ejecutan cordones "
             "circunferenciales de filete en los extremos, las costuras "
             "longitudinales del collar se sueldan a tope con penetración completa "
             "(Fig. 206-1.1.2-1) y se prevé venteo durante el cordón de cierre "
             "final. Si no se ejecutan los cordones circunferenciales (Type A), las "
             "costuras longitudinales pueden ser una junta a tope en ranura o una "
             "junta solapada soldada con filete (Fig. 206-1.1.1-1).", "206-4.4"),
            ("Cateto del filete de extremo", cateto,
             "206-3.5 · Fig. 206-3.5-1 y -2"),
            ("Procedimiento y bajo hidrógeno", "El procedimiento de los cordones "
             "circunferenciales de filete debe ser adecuado a los materiales y a "
             "las condiciones de severidad de enfriamiento del cordón en la "
             "ubicación instalada, conforme al código de construcción o "
             "post-construcción. Debería usarse técnica de soldadura de bajo "
             "hidrógeno. Para costuras longitudinales sin fleje de respaldo, ver "
             "206-4.5.", "206-4.4"),
            ("Presión durante la reparación", "Se recomienda reducir la presión de "
             "operación del portador manteniendo el flujo mientras se ejecuta la "
             "reparación; ver API RP 2201 para recomendaciones de soldadura de "
             "tubería en servicio. La presión recomendada durante la instalación "
             "del collar está entre el 50 % y el 80 % de la presión de operación. "
             "La línea también puede sacarse de servicio para reparar, pero "
             "entonces hay que considerar la perforación por quemado.", "206-4.5"),
            ("Soldadura en servicio", "Se consulta el Art. 210. Como mínimo, la "
             "calificación del proceso de soldadura debe tener en cuenta (a) el "
             "potencial de agrietamiento inducido por hidrógeno en la zona afectada "
             "por el calor, por la velocidad de enfriamiento acelerada y el "
             "hidrógeno del ambiente de soldadura; (b) el riesgo de formar una ZAT "
             "inaceptablemente dura por la química del material base del collar y "
             "del tubo; y (c) la posible perforación del tubo.",
             "206-4.6 · Art. 210"),
            ("Calificación de procedimientos y soldadores", "Los procedimientos, "
             "soldadores y operadores se califican conforme al código "
             "post-construcción vigente. Si no se especifica otra cosa, se usa ASME "
             "BPVC Sección IX para calificación de procedimiento y de desempeño. La "
             "guía de precalentamiento y de tratamiento térmico post-soldadura, y "
             "la de soldadura en servicio cuando aplique, se toma del código de "
             "construcción o post-construcción aplicable.", "206-4.7"),
            ("Consumibles del proyecto", "Consumible compatible de resistencia "
             "igual o mayor que el metal base, técnica de bajo hidrógeno. Ajustar "
             "al WPS calificado del proyecto.",
             "206-2.10 · 206-4.4 · WPS del proyecto", "EDITABLE"),
        )),
        ("EXAMEN  //  206-5", (
            ("Examen visual", "Todos los ajustes del collar se inspeccionan antes "
             "de soldar. Las soldaduras se examinan visualmente.", "206-5.1"),
            ("Collar Type A", "En un collar Type A, la zona de raíz del cordón se "
             "examina visualmente durante la soldadura para verificar penetración y "
             "fusión adecuadas. Las costuras longitudinales se examinan por "
             "líquidos penetrantes, partículas magnéticas o ultrasonido una vez "
             "completadas.", "206-5.2"),
            ("Collar Type B", "En un collar Type B, el material base del tubo "
             "portador se examina por ultrasonido —espesor, grietas y posible "
             "laminación— en la zona donde se aplicarán los cordones "
             "circunferenciales. Si no se usa fleje de respaldo bajo la costura "
             "longitudinal, la zona bajo ella también se examina por ultrasonido "
             "antes de soldar. Las costuras longitudinales se inspeccionan al "
             "terminar. La primera y la última pasada de los cordones "
             "circunferenciales deberían examinarse por partículas magnéticas o "
             "líquidos penetrantes después de soldar.", "206-5.3"),
            ("Agrietamiento diferido", "Donde el agrietamiento diferido sea una "
             "preocupación, el examen no destructivo de los cordones "
             "circunferenciales no debería ejecutarse antes de 24 h de terminada la "
             "soldadura. Como alternativa, no antes de 48 h de terminada una "
             "soldadura en servicio cuando haya alta probabilidad de agrietamiento "
             "por hidrógeno.", "206-5.3"),
            ("Examen en proceso", "El propietario puede exigir examen visual «en "
             "proceso» completo de la instalación soldada del collar, tal como lo "
             "describe el para. 344.7 de ASME B31.3. Cuando se ejecuta, los "
             "resultados se documentan. Los exámenes los realiza personal que "
             "cumple los requisitos de calificación del código de construcción o "
             "post-construcción aplicable.", "206-5.4"),
            ("Procedimiento y criterio de END", "Los procedimientos de examen no "
             "destructivo deberían calificarse conforme al código de construcción o "
             "post-construcción aplicable; si ese código no fija requisitos de "
             "procedimiento, debería calificarse según ASME BPVC Sección V. El "
             "criterio de aceptación debería ser el del código aplicable, salvo que "
             "el propio Art. 206 dé criterios alternativos; donde no haya criterio, "
             "debería usarse el de ASME BPVC Sección VIII, División 1 o 2.",
             "206-5.5"),
        )),
        ("PRUEBA  //  206-6 TESTING · Art. 501", (
            ("Prueba de hermeticidad del anular", "Si el propietario lo requiere, "
             "debería ejecutarse una prueba de hermeticidad en collares Type B, por "
             "una de dos vías: (a) presurizar el anular entre el collar y el tubo "
             "portador conforme al código de construcción o post-construcción "
             "aplicable, con una presión de prueba elegida de modo que el tubo "
             "interior NO colapse por presión externa; o (b) ejecutar una prueba de "
             "fuga sensible como la describe el para. 345.8 de ASME B31.3 u otra "
             "norma nacional reconocida. El Art. 501 da guía adicional.",
             "206-6(a) · 206-6(b) · Art. 501"),
            ("Energía neumática", "Si la prueba se ejecuta con gas, la energía "
             "almacenada exige precauciones de seguridad y una distancia mínima al "
             "personal. El Paso 8 del anexo del motor las calcula con las "
             "ecuaciones del Apéndice Obligatorio 501-II/III, leídas de resources/.",
             "Art. 501-II / 501-III"),
        )),
    )
    return build_especificaciones(
        wb, ESPEC_206, COLLAR_MOTOR,
        "ESPECIFICACIONES TECNICAS — COLLAR DE ENCIERRO TOTAL (ASME PCC-2 Art. 206)",
        "Fabricacion, examen y prueba. Cada fila transcribe el parrafo de PCC-2 "
        "que cita a su derecha; las cifras las lee del motor Collar_PCC2_Art206.",
        bloques)


# ---------------------------------------------------------------------------
# Fase 10 — Instrucciones de uso, una pestana por motor
# ---------------------------------------------------------------------------
# El plan pide explicar «por cada fila o celda no bloqueada qué debe insertar el
# ingeniero» y «por cada celda de fórmula qué calcula y de qué cláusula sale».
# Son ~45 celdas de entrada y ~200 de formula por motor, y escribirlas a mano
# aqui seria una segunda copia de lo que el motor ya dice —la clase de copia que
# este libro ya vio divergir dos veces (HojasNavegables, la leyenda de color)—.
#
# Asi que la tabla NO se escribe: se DERIVA de la hoja del motor ya construida.
# Cada celda de un motor lleva, puesto por lab()/inp()/calc(), todo lo que esta
# pestana necesita: el rotulo en la columna A de su fila, la unidad en C, la
# referencia al codigo en G y un comentario de Excel que empieza por «Entrada:»
# o «Calculo:» y dice exactamente que hacer o que calcula. El tipo de celda no se
# declara: se lee del estado real —desbloqueada, con validacion de lista, o
# formula—, el mismo criterio con el que la leyenda pinta y el boton de reinicio
# limpia. Y el EJEMPLO de cada entrada es el valor del caso precargado, que es un
# caso real ya validado.
#
# Lo que si se escribe a mano es lo que el motor no puede decir de si mismo: para
# que sirve, que se hace en cada seccion, y como se leen la leyenda, el
# conmutador de unidades, el semaforo y el boton de reinicio.
INSTR_212, INSTR_206 = "Instruc_PCC2_Art212", "Instruc_PCC2_Art206"
INSTR_NCOLS = 7
INSTR_ANCHO = {"A": 10, "B": 40, "C": 13, "D": 24, "E": 24, "F": 24, "G": 30}
INSTR_HDR = ((1, "Celda"), (2, "Qué es"), (3, "Tipo"),
             (4, "Qué debe insertar / qué calcula"), (7, "Referencia del motor"))

AVISO_INSTR = (
    "Esta hoja se genera del propio motor: el rotulo, el tipo de celda, la "
    "explicacion y la referencia son los que la hoja del motor lleva en cada "
    "celda, no una copia escrita aparte. Si el motor cambia, esta hoja cambia con "
    "el. El EJEMPLO de cada entrada es el valor del caso precargado.")


def _celdas_con_lista(ws):
    """Coordenadas cubiertas por una validacion de lista en `ws`."""
    out = set()
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list":
            continue
        for rango in dv.sqref.ranges:
            for fila in range(rango.min_row, rango.max_row + 1):
                for col in range(rango.min_col, rango.max_col + 1):
                    out.add(f"{get_column_letter(col)}{fila}")
    return out


def _es_banda(ws, c):
    """True si la celda es la banda de una seccion del motor.

    Se reconoce por la ESTRUCTURA que le dio el builder —fusionada de A a G y
    con el relleno de banda—, no por su texto: el texto de las bandas cambia
    (numeral, parentetico, mayusculas) y un reconocimiento por cadena habria que
    perseguirlo cada vez que se renumera una seccion.
    """
    if c.column != 1 or c.value is None or c.row < 4:
        return False
    # Bloque de tinta + macrotipografia: es lo que distingue una banda de la fila
    # de encabezado, que comparte el relleno pero va en la mono del libro. NO se
    # exige que este fusionada: el 212 fusiona A:G sus bandas y el 206 no, y esa
    # diferencia es de estilo de cada motor, no de que cosa es una banda.
    if c.font is None or c.font.name != MACRO:
        return False
    return (c.fill is not None and c.fill.patternType
            and str(getattr(c.fill.fgColor, "rgb", ""))[-6:].upper() == TINTA)


def _guia_de_celdas(ws):
    """(banda, [(celda, rotulo, tipo, texto, referencia, ejemplo)]) del motor.

    Recorre la hoja de arriba abajo en A..G y reparte cada celda de entrada o de
    formula bajo la ultima banda de seccion que encontro. Devuelve las secciones
    en el orden en que se leen, que es el orden en que se rellena la hoja.
    """
    listas = _celdas_con_lista(ws)
    fin = _fin_tabla_motor(ws)
    secciones, actual = [], None
    sin_nota = []
    for fila in ws.iter_rows(min_row=1, max_row=fin, max_col=MOTOR_NCOLS):
        for c in fila:
            if _es_banda(ws, c):
                actual = (str(c.value), [])
                secciones.append(actual)
        if actual is None:
            continue
        rotulo_fila = ws.cell(fila[0].row, 1).value
        unidad = ws.cell(fila[0].row, 3).value
        referencia = ws.cell(fila[0].row, MOTOR_NCOLS).value
        for c in fila[3:MOTOR_NCOLS - 1]:          # columnas D, E y F: los valores
            if c.value is None or c.hyperlink is not None:
                continue
            if isinstance(c, MergedCell):
                continue
            formula = isinstance(c.value, str) and c.value.startswith("=")
            entrada = _es_entrada_motor(c)
            if not (formula or entrada):
                continue
            tipo = ("LISTA" if c.coordinate in listas
                    else "SE TECLEA" if entrada else "FORMULA")
            nota = c.comment.text if c.comment is not None else ""
            if not nota:
                sin_nota.append(f"{ws.title}!{c.coordinate}")
            # La direccion de la celda queda en la guia (columna «Celda»), y es
            # lo que distingue dos columnas de la misma fila cuando la seccion
            # tiene varios casos en paralelo: D es metal base u operacion, E es
            # collar/parche o diseno. El comentario de cada celda ya lo dice.
            rot = str(rotulo_fila or "").strip()
            # Tras la Fase 7 el rotulo de unidad de casi toda fila es la formula
            # del selector, `=IF(<sel>,"mm","in")`. Se muestran las DOS unidades,
            # que es la informacion util aqui: leer "[mm / in]" dice ademas que esa
            # fila cambia de unidad con el conmutador.
            u = str(unidad or "")
            if u.startswith("="):
                m = re.search(r'"([^"]*)"\s*,\s*"([^"]*)"\s*\)\s*$', u)
                u = f"{m.group(1)} / {m.group(2)}" if m else ""
            if u:
                rot = f"{rot}   [{u}]"
            # `ejemplo` es None en una formula (no se teclea) y el valor del caso
            # precargado en una entrada. Una entrada VACIA en el caso precargado
            # -los cinco niveles de la cascada de material, que el caso resuelve
            # por la via de escape- se queda como cadena vacia, y la guia lo dice
            # en vez de callarlo: «se elige de la lista» es una instruccion, «sin
            # ejemplo» es un hueco.
            actual[1].append((c.coordinate, rot, tipo, nota,
                              str(referencia or ""),
                              None if formula else c.value))
    if sin_nota:
        ISSUES.append(
            f"{ws.title}: {len(sin_nota)} celda(s) de entrada o formula sin "
            f"comentario propio; su fila de la guia de uso queda sin explicacion "
            f"({', '.join(sin_nota[:6])}).")
    return [(b, filas) for b, filas in secciones if filas]


def preparar_impresion(ws, ncols, hasta_fila=None):
    """Area de impresion y ajuste a lo ancho de una hoja de motor o documento.

    Sin esto la hoja se imprime -y se exporta a PDF- con el area de uso ENTERA,
    que llega hasta las columnas ocultas de listas materializadas y de claves de
    navegacion (la 100 del manifiesto de reinicio): al ajustar a una pagina de
    ancho, la tabla de A..G queda microscopica y el resto del folio en blanco.
    Se descubrio exportando la hoja y mirandola, que es la unica evidencia valida
    del aspecto en este libro.

    La fila 1 se repite en cada pagina: estas hojas son largas y una pagina 4 sin
    titulo no dice de que motor es.
    """
    fin = hasta_fila or max(
        (c.row for fila in ws.iter_rows(max_col=ncols) for c in fila
         if c.value is not None), default=1)
    ws.print_area = f"A1:{get_column_letter(ncols)}{fin}"
    ws.print_title_rows = "1:1"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "landscape"
    return ws


def _plano(ws, r, c, v):
    """Escribe TEXTO, aunque empiece por «=».

    La guia copia rotulos y referencias de la hoja del motor, y ahi hay celdas
    cuyo texto empieza por «=» —la columna de referencia dice cosas como
    «= $D$26» para explicar de donde sale un valor—. openpyxl lo escribiria como
    FORMULA y en la guia se evaluaria contra ESTA hoja: en vez de la explicacion
    se veria un numero de otra fila, o un error. Mismo criterio y mismo arreglo
    que `_txt_celda` en las hojas de la Seccion II.
    """
    cel = ws.cell(r, c)
    cel.value = "" if v is None else str(v)
    if cel.value[:1] in "=+-@":
        cel.data_type = "s"
    return cel


def _instr_parrafo(ws, r, rotulo, texto):
    """Fila de prosa: rotulo en A:B y parrafo en C:G."""
    a = _mrg(ws, r, 1, 2, rotulo)
    a.font = Font(name=MONO, size=10, bold=True, color=TINTA)
    a.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    v = _mrg(ws, r, 3, INSTR_NCOLS, texto)
    v.font = DATA_F
    v.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 13.5 * max(1, -(-len(str(texto)) // 105)) + 4
    return r + 1


def build_instrucciones_motor(wb, nombre, motor, titulo, subtitulo, prosa):
    """Guia de uso de un motor: prosa escrita + tabla derivada de la hoja.

    `prosa` es ((rotulo de banda, ((rotulo, parrafo), ...)), ...) y se escribe
    antes de la guia celda a celda.
    """
    if nombre in wb.sheetnames:
        del wb[nombre]
    ws = new_sheet(wb, nombre, titulo, subtitulo)
    autosize(ws, INSTR_ANCHO)
    ws.merge_cells(f"A1:{get_column_letter(INSTR_NCOLS)}1")
    ws.merge_cells(f"A2:{get_column_letter(INSTR_NCOLS)}2")
    _boton(ws, 3, COL_BTN_DOC_1, COL_BTN_DOC_2, TXT_BTN_MOTOR, motor)
    _ocultar_columnas_clave(ws, (COL_BTN_DOC_1,))

    r = 5
    for rotulo_banda, filas in prosa:
        banda(ws, r, rotulo_banda, INSTR_NCOLS)
        r += 1
        for rot, texto in filas:
            r = _instr_parrafo(ws, r, rot, texto)
        r += 1

    banda(ws, r, "GUIA CELDA A CELDA  //  DERIVADA DE LA HOJA DEL MOTOR",
          INSTR_NCOLS)
    r += 1
    n_entradas = n_formulas = 0
    for seccion, filas in _guia_de_celdas(wb[motor]):
        # La banda de la seccion del motor se repite aqui TAL CUAL, incluido su
        # numeral: es la unica forma de que el ingeniero sepa en que seccion de
        # la hoja esta la celda que esta leyendo.
        s = _mrg(ws, r, 1, INSTR_NCOLS, seccion)
        s.font = Font(name=MACRO, size=10, color=TINTA)
        s.fill = PatternFill("solid", fgColor=PAPEL_2)
        s.alignment = Alignment(vertical="center", indent=1)
        franja(ws, r, 1, INSTR_NCOLS)
        r += 1
        for j, h in INSTR_HDR:
            c = ws.cell(r, j, h)
            c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
        for j in range(5, 7):
            c = ws.cell(r, j)
            c.fill, c.border = HDR_FILL, BOX_FRANJA
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        r += 1
        for celda, rot, tipo, nota, ref, ejemplo in filas:
            a = ws.cell(r, 1, celda)
            a.font = Font(name=MONO, size=9, bold=True, color=TINTA)
            # Alineacion SUPERIOR en las tres columnas cortas: con el alto que
            # pide un parrafo de cuatro lineas, una celda centrada o al pie deja
            # la direccion flotando lejos de la fila que nombra. Visto en el PDF.
            a.alignment = Alignment(vertical="top")
            b = _plano(ws, r, 2, rot)
            b.font = DATA_F
            b.alignment = Alignment(vertical="top", wrap_text=True)
            t = ws.cell(r, 3, tipo)
            t.font = Font(name=MONO, size=9, bold=True,
                          color=TINTA if tipo == "FORMULA" else ROJO)
            t.alignment = Alignment(vertical="top")
            if ejemplo is None:
                texto = nota
            elif str(ejemplo).strip() == "":
                texto = (f"{nota}  ·  Ejemplo: vacia en el caso precargado; se "
                         f"rellena eligiendo de la lista desplegable.")
            else:
                texto = f"{nota}  ·  Ejemplo: {ejemplo}"

            v = _mrg(ws, r, 4, 6)
            _plano(ws, r, 4, texto)
            v.font = DATA_F
            v.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
            g = _plano(ws, r, 7, ref)
            g.font = SRC_F
            g.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
            # 78 caracteres por linea en D:F y 38 en B, medidos sobre el PDF
            # exportado. Con 72 sobraba aire; con 92 se cortaba la ultima linea,
            # que es peor: una instruccion a medias no se lee como una fila alta.
            ws.row_dimensions[r].height = 13.2 * max(
                1, -(-len(str(texto)) // 78), -(-len(str(rot)) // 38)) + 3
            n_entradas += tipo != "FORMULA"
            n_formulas += tipo == "FORMULA"
            r += 1
        r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=INSTR_NCOLS)
    av = ws.cell(r, 1, AVISO_INSTR)
    av.font = SRC_F
    av.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 30

    preparar_impresion(ws, INSTR_NCOLS)
    ISSUES.append(f"{nombre}: guia de uso derivada del motor — "
                  f"{n_entradas} celdas de entrada y {n_formulas} de formula.")
    ws.protection.password = "0000"
    ws.protection.sheet = True
    return ws


# Prosa comun a los dos motores: la leyenda de color, el conmutador de unidades,
# el semaforo y el boton de reinicio son el mismo mecanismo en los dos, y
# describirlos dos veces era garantizar que un dia dijeran cosas distintas.
def _prosa_comun(motor, celda_unidad, fila_dictamen):
    return (
        ("COMO SE LEE LA HOJA  //  COLOR, UNIDADES, SEMAFORO Y REINICIO", (
            ("Leyenda de color",
             "Cada celda declara con su relleno quien la rellena, y la leyenda "
             "esta impresa en la fila 3 de la hoja del motor: AMARILLO = la "
             "rellena usted, tecleando o eligiendo de una lista desplegable; GRIS "
             "= la calcula el libro y esta bloqueada; PAPEL = rotulo, unidad o "
             "referencia. El color no se pone celda a celda: se deriva del estado "
             "real de cada celda, asi que la leyenda no puede mentir."),
            ("Que se teclea y que se elige",
             "Si el dato esta tabulado en una base del libro —diametro nominal, "
             "cedula, espesor, material— se elige de una lista desplegable y no se "
             "teclea: esa es la regla 14 del proyecto, y existe para eliminar el "
             "error de tecleo como clase de fallo. Lo que se teclea es lo que "
             "ninguna tabla publica: presion y temperatura de operacion, "
             "dimensiones del defecto, sobreespesor de corrosion y medidas de la "
             "reparacion. La columna «Tipo» de la guia de abajo lo dice celda a "
             "celda."),
            ("Conmutador SI / US",
             f"La celda {celda_unidad} de la hoja del motor elige el sistema de "
             "unidades. Cambia DE QUE EDICION del codigo se lee —la metrica o la "
             "U.S. Customary— y los rotulos de unidad de toda la hoja; NUNCA "
             "convierte un valor, porque las dos ediciones son extracciones "
             "independientes de lo que cada una imprime. Los datos que usted ya "
             "tecleo no se convierten al cambiarlo: hay que volver a teclearlos en "
             "el sistema nuevo. Si un material no tiene homologo en la edicion US, "
             "el motor lo dice y bloquea el resultado en vez de dar un numero en "
             "la unidad equivocada."),
            ("Semaforo de aceptacion",
             "Cada verificacion se pinta sola: VERDE cuando cumple y ROJO cuando "
             "no. El rojo se pinta por complemento —todo lo que no sea el "
             "resultado favorable—, asi que un error de calculo o una celda vacia "
             "tambien salen en rojo en vez de pasar por buenos."),
            ("Dictamen global",
             f"La fila {fila_dictamen} de la hoja del motor es el dictamen del "
             "diseno completo, en bloque propio. Primero resuelve la compuerta de "
             "elegibilidad; si el material esta sin elegir o fuera de rango de "
             "temperatura lo dice; y solo da APTO si ademas TODAS las "
             "verificaciones cumplen. Un dictamen distinto de APTO se arregla "
             "cambiando el DISENO o las entradas, no el motor."),
            ("Boton de reiniciar entradas",
             "El boton [ RESET ] de la esquina superior derecha vacia todas las "
             "celdas de entrada de la hoja —las amarillas— y deja el motor en "
             "blanco para un caso nuevo. Pide confirmacion, no se puede deshacer, "
             "y no toca el formato, las listas desplegables ni las formulas. La "
             "lista de celdas que limpia la publica el propio motor en una columna "
             "oculta, asi que ninguna entrada nueva se le puede quedar fuera."),
            ("Que NO hace este libro",
             "No sustituye el criterio del ingeniero ni el codigo. Es una "
             "herramienta de calculo trazable: todo valor normativo sale de la "
             "extraccion auditada de la norma y cada resultado cita el parrafo del "
             "que sale. Verificar entradas y resultados, y complementar con "
             "WPS/PQR, ATS/JSA y los registros del propietario."),
        )),
    )


def build_instrucciones_art212(wb):
    prosa = (
        ("QUE HACE ESTE MOTOR  //  ASME PCC-2 Art. 212", (
            ("Para que sirve",
             "Dimensiona una reparacion por PARCHE DE PLANCHA soldado con filete "
             "perimetral sobre un componente que retiene presion, segun el "
             "Art. 212 de ASME PCC-2. Resuelve el espesor requerido por presion "
             "interna, la carga admisible del filete perimetral, el limite de "
             "conformado en frio de la plancha y las verificaciones de aceptacion, "
             "y ademas conduce el flujo completo de la reparacion en ocho pasos."),
            ("Cuando NO se usa",
             "Cuando el mecanismo de dano, su extension o el dano futuro no se "
             "pueden caracterizar; cuando hay grietas sin arresto ni evaluacion de "
             "aptitud para el servicio; y en servicio letal o de extrema "
             "peligrosidad. El Paso 1 del anexo resuelve esa compuerta y un estado "
             "bloqueante detiene el dictamen global."),
            ("Codigo de construccion",
             "El selector de aplicacion decide el codigo y con el la fuente del "
             "esfuerzo admisible: ASME B31.3 (Tabla A-1) para tuberia, o ASME BPVC "
             "Seccion VIII-1 con la Tabla 1A de la Seccion II-D para virola "
             "cilindrica y cabezal/esfera."),
            ("El orden en que se rellena",
             "De arriba abajo, y la hoja esta ordenada asi a proposito: 1 datos de "
             "entrada -> 2 resolucion del material (es otra entrada, la mas larga) "
             "-> 3 parametros -> 4 geometria -> 5 cargas y soldadura -> 6 "
             "resultados -> 7 verificaciones. Debajo, el anexo con los ocho pasos "
             "del flujo. Las especificaciones tecnicas viven en su propia pestana."),
            ("Que se hace en la seccion 2",
             "Se identifica el material del metal base y el del parche con una "
             "cascada de seis listas dependientes (familia -> composicion -> forma "
             "-> especificacion -> tipo/grado -> variante) contra las bases "
             "auditadas del libro, y el motor devuelve el esfuerzo admisible S a "
             "la temperatura de evaluacion. La celda «Variante» es una via de "
             "escape: si ya conoce el identificador del material, peguelo ahi. El "
             "rastro de como se resolvio (fila localizada, puntos tabulados, T1/T2, "
             "S en T1/T2, temperatura maxima) va PLEGADO: se abre con el «+» del "
             "margen izquierdo para auditarlo."),
        )),
    ) + _prosa_comun(MOTOR, UNIDAD_212.replace("$", ""), MAPA_FILAS_212.get(90, 90))
    return build_instrucciones_motor(
        wb, INSTR_212, MOTOR,
        "INSTRUCCIONES DE USO — PARCHE SOLDADO (ASME PCC-2 Art. 212)",
        "Que hace el motor, que se rellena en cada seccion y que calcula cada "
        "celda. La guia celda a celda se deriva de la propia hoja del motor.",
        prosa)


def build_instrucciones_art206(wb):
    prosa = (
        ("QUE HACE ESTE MOTOR  //  ASME PCC-2 Art. 206", (
            ("Para que sirve",
             "Dimensiona un COLLAR DE ENCIERRO TOTAL (full encirclement sleeve) "
             "soldado sobre tuberia, segun el Art. 206 de ASME PCC-2, en sus dos "
             "tipos: Type A, que refuerza el portador y no contiene presion, y "
             "Type B, que si la contiene y lleva cordones circunferenciales de "
             "cierre. Resuelve el espesor requerido, las dimensiones del collar y "
             "las verificaciones de aceptacion, y conduce el flujo en ocho pasos."),
            ("La eleccion del tipo es lo primero",
             "Un defecto con FUGA exige Type B (206-2.3), y un defecto "
             "circunferencial no lo refuerza un Type A (206-2.5). El motor guia la "
             "eleccion y avisa cuando el tipo elegido no corresponde al defecto "
             "declarado, pero el aviso NO decide por usted: es criterio de "
             "ingenieria."),
            ("Que espesor exige el codigo",
             "El Type B se dimensiona con un espesor de pared igual o mayor que el "
             "requerido para la PRESION DE DISENO MAXIMA ADMISIBLE del tubo "
             "portador (206-3.3), no para la presion de operacion. Por eso el motor "
             "evalua dos presiones —operacion y diseno— y el espesor gobernante "
             "sale de la de diseno. El Type A tiene ademas su propio minimo de dos "
             "tercios del espesor del portador."),
            ("El orden en que se rellena",
             "1 datos de entrada -> 2 resolucion del material -> 3 parametros -> 4 "
             "geometria del collar -> 5 espesor requerido -> 6 verificaciones y "
             "avisos, y debajo el anexo con los ocho pasos del flujo. Las "
             "especificaciones tecnicas viven en su propia pestana."),
            ("Que se hace en la seccion 2",
             "Igual que en el Art. 212: el material del collar se identifica con "
             "la cascada de seis listas dependientes contra las bases auditadas del "
             "libro, y el motor devuelve su esfuerzo admisible a la temperatura de "
             "evaluacion. El rastro de la resolucion va plegado y se abre con el "
             "«+» del margen."),
        )),
    ) + _prosa_comun(COLLAR_MOTOR, UNIDAD_206.replace("$", ""),
                     MAPA_FILAS_206.get(69, 69))
    return build_instrucciones_motor(
        wb, INSTR_206, COLLAR_MOTOR,
        "INSTRUCCIONES DE USO — COLLAR DE ENCIERRO TOTAL (ASME PCC-2 Art. 206)",
        "Que hace el motor, que se rellena en cada seccion y que calcula cada "
        "celda. La guia celda a celda se deriva de la propia hoja del motor.",
        prosa)


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
     "DB_B31_B1 / DB_B31_B1C — Tabla B-1 / B-1C del Apendice B: esfuerzo de diseno "
     "hidrostatico (HDS) de tuberia termoplastica y limites de temperatura recomendados. "
     "36 filas por edicion, con el HDS tabulado a 23, 38, 82 y 93 °C (73, 100, 180 y "
     "200 °F).\n"
     "Del Apendice B, este libro carga SOLO la Tabla B-1 / B-1C. Las Tablas B-2 y B-3 "
     "(listados de especificacion de RTR y RPM) y B-4, B-5 y B-6 (presiones admisibles "
     "de concreto, vidrio borosilicato y PEX-AL-PEX) NO estan cargadas, y el arbol lo "
     "dice con una tarjeta marcador. Para esas cinco tablas, el codigo impreso manda.\n"
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
     "Buscar_B31_B1 — Tabla B-1 / B-1C: esfuerzo de diseno hidrostatico (HDS) de tuberia "
     "TERMOPLASTICA, frente a la temperatura. Es el valor que entra como S en la eq. (26a) "
     "del para. A304.1.2, t = PD/(2S+P). Tres reglas de rango, todas del Capitulo VII y "
     "distintas de las de los metales: se INTERPOLA linealmente (para. A302.3.1(b)); por "
     "debajo de la primera temperatura tabulada se SOSTIENE ese HDS y no se extrapola "
     "(Nota (3) de la tabla, concordante con el para. A323.2.2(b)); y se BLOQUEA por "
     "arriba tanto en el limite maximo recomendado de las Notas (1) y (2) como en el "
     "ultimo punto tabulado (para. A323.2.1(a)), con aviso distinto para cada causa. "
     "Conmutador SI/US: cambia de HOJA (B-1 <-> B-1C), nunca convierte. Es lo UNICO del "
     "Apendice B que este libro carga: las Tablas B-2 a B-6 no tienen buscador ni hoja "
     "de datos aqui, y el arbol lo declara con una tarjeta marcador."),
    ("4. Conmutador de unidades SI / US",
     "Cada buscador tiene la celda 'Sistema de unidades'. En SI lee la tabla metrica del "
     "codigo (MPa, C); en US lee la tabla nativa en unidades inglesas (ksi, F): el B31.3 usa "
     "sus tablas companion 'C' y la II-D su edicion U.S. Customary 2025. NO hay conversiones. "
     "La seleccion se hace siempre sobre la edicion metrica y se enlaza con la US mediante "
     "una clave de identidad; si un material no tiene homologo, la ficha lo indica. Regla: no "
     "mezclar sistemas dentro de un mismo calculo. El motor opera siempre en SI."),
    ("5. El motor de calculo",
     "En la seccion 2 del modulo hay la misma cascada, por separado para el metal base y para "
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
    ("10. Convenciones de lectura del libro",
     "Calculo en SI por defecto. El libro entero va en un solo sistema visual: papel, tinta "
     "y UN acento rojo.\n"
     "· CAMPO QUE SE RELLENA: papel limpio dentro de una caja de tinta. Todos son listas "
     "desplegables.\n"
     "· UNICA CELDA QUE SE TECLEA (temperatura de consulta): el mismo campo, pero con la "
     "caja en ROJO. Es la unica caja roja de la hoja.\n"
     "· BANDA DE SECCION: bloque de tinta con el rotulo entre corchetes y franja roja al "
     "pie; separa las zonas de la hoja.\n"
     "· CIFRA DE RESULTADO: tipografia ancha, mucho mayor que el resto. Debajo, su unidad.\n"
     "· ROJO: siempre significa lo mismo — bloqueado, fuera de rango o aviso. Nunca es "
     "decoracion.\n"
     "· SEMAFORO DE SELECCION: bloque verde SELECCION COMPLETA / bloque ambar SELECCION "
     "INCOMPLETA. Verde = CUMPLE/APTO, rojo = NO CUMPLE/revisar.\n"
     "· GRIS APAGADO: no cargado en este libro, o metadato de procedencia.\n"
     "Hojas de calculo protegidas con contrasena preliminar 0000; las bases y los buscadores "
     "quedan sin proteger para poder filtrar y copiar."),
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
    ws["B2"] = "MOTOR DE CALCULO // ASME PCC (CODIGOS POST-CONSTRUCCION)"
    ws["B2"].font = Font(name=MACRO, size=14, color=TINTA)
    ws["B3"] = ("Reparacion, montaje e inspeccion de equipos a presion y tuberia segun la "
                "familia ASME PCC · Codigos de construccion: ASME B31.3-2024 y ASME BPVC "
                "Seccion VIII-1 con II-D 2025")
    ws["B3"].font = Font(name=MONO, size=9, color=GRIS)
    ws["B4"] = version_note
    ws["B4"].font = Font(name=MONO, size=10, bold=True, color=ROJO)
    r = 6
    for title, body in INSTRUCCIONES:
        t = ws.cell(r, 2, title.upper())
        t.font = Font(name=MONO, size=10, bold=True, color=PAPEL)
        t.fill = BAND_FILL
        t.alignment = Alignment(vertical="top", indent=1, wrap_text=True)
        c = ws.cell(r, 3, body)
        c.font = Font(name=MONO, size=10, color=TINTA)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = max(30, 13 * (body.count("\n") + 1 + len(body) // 110))
        r += 1
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 118
    ws.protection.password = "0000"
    ws.protection.sheet = True


def build_meta(wb, counts):
    # Sin revision escrita a mano en el rotulo: la anterior se quedo en
    # "Rev.2" mientras el libro iba por la 4, y una hoja de TRAZABILIDAD que
    # se equivoca sobre su propia version es lo contrario de trazable.
    ws = new_sheet(wb, "_meta", "TRAZABILIDAD DE LAS BASES DE DATOS — PLAN-DB-MAT-001",
                   "Cada base referencia el archivo de resources/ del que procede "
                   "(regla 5.4 de knowledge/claude.md). Al cambiar de edicion del codigo, "
                   "re-ejecutar build_db_materiales.py sobre los nuevos JSON.")
    n = write_headers(ws, ["Hoja", "Tabla", "Archivo fuente (resources/)", "Edicion",
                           "Sistema", "Filas", "Nota"])
    last = write_rows(ws, [([m["hoja"], m["tabla"], m["archivo"], m["edicion"],
                             m["sistema"], m["filas"], m["nota"]], {}) for m in META], n)
    r = last + 2
    ws.cell(r, 1, "VERIFICACION DE CONTEOS").font = Font(name=MACRO, size=10, color=TINTA)
    for k, v in counts.items():
        r += 1
        ws.cell(r, 1, k).font = DATA_F
        ws.cell(r, 3, v).font = DATA_F
    r += 2
    ws.cell(r, 1, "LIMITACIONES Y OBSERVACIONES").font = Font(name=MACRO, size=10, color=TINTA)
    for msg in ISSUES:
        r += 1
        c = ws.cell(r, 1, msg)
        c.font = DATA_F
        c.alignment = Alignment(wrap_text=True)
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
# Dashboard y capa de navegacion (Rev. 5 — arbol jerarquico)
# ---------------------------------------------------------------------------
# El libro se entrega como .xlsm: la unica hoja visible es el Dashboard y la
# macro conmuta la visibilidad del resto. Los estados se graban ademas en el
# archivo, de modo que quien abra con las macros bloqueadas siga sin ver
# ninguna base de datos.
DASH = "Dashboard"

# La clave de destino de cada boton se guarda oculta en (fila del boton,
# COL_CLAVE_BASE + columna del boton). Depende de la columna, y no solo de la
# fila, porque cada nivel pone tres tarjetas por banda: con una unica
# columna de claves las tres escribirian en la misma celda.
#
# La base es 66 (BN) porque BM es la columna mas alta que usa cualquier hoja
# navegable. No se usa N porque Parche_PCC2_Art212 ya ocupa K..P con sus
# listas de cascada.
# DEBE coincidir con COL_CLAVE_BASE de vba/mod_nav.vba.
COL_CLAVE_BASE = 66

# Celda del aviso de macros. DEBE coincidir con CELDA_AVISO de vba/mod_nav.vba.
FILA_AVISO = 4


# ---------------------------------------------------------------------------
# El arbol de navegacion: se declara UNA sola vez y todo lo demas se deriva
# ---------------------------------------------------------------------------
# El Dashboard conserva sus tres bandas por TIPO DE ARTEFACTO -motores de
# calculo, motores de busqueda y bases de datos-, que es la primera pregunta
# que se hace quien abre el libro: «que quiero hacer». Y solo DESPUES de esa
# separacion empieza la categorizacion, con el arbol con el que se CITA una
# norma:
#
#     PUBLICANTE > DISCIPLINA > CODIGO DE LA DISCIPLINA > STANDARD CONCRETO
#
# Cada banda tiene su PROPIA rama completa, y por eso la tarjeta ASME aparece
# tres veces: la del BPVC que lleva a los cinco buscadores de la Parte D no es
# la misma que la que lleva a las nueve hojas de datos de las Partes A, B y C.
# La separacion por tipo atraviesa todo el recorrido y un motor nunca se cruza
# con una hoja de datos. El precio son tres cascadas paralelas; lo que compra
# es que en cada pantalla todo lo que se ve es del mismo tipo.
#
# De este arbol se derivan, en preorden y sin escribir nada dos veces:
# HOJAS_NAV, NAVEGABLES, DESTINOS, PADRE, ROTULO y ANCLA_VOLVER. Anadir manana
# el B31.1 es anadir un Nodo aqui: no toca el constructor de hojas, ni el
# mecanismo de navegacion, ni -salvo la lista literal que el VBA tiene que
# repetir- una sola linea de VBA.
class Nodo(NamedTuple):
    """Un nivel del arbol de navegacion.

    `hoja`     nombre de la hoja NAV_* propia del nodo. Solo la tienen los
               nodos interiores (y el Dashboard, que es la raiz).
    `destino`  hoja final del libro a la que lleva la tarjeta. Solo la tienen
               las hojas del arbol. Un nodo tiene `hoja` o `destino`, nunca
               las dos: la tarjeta de un nodo interior lleva a su NAV_*.
    `corto`    rotulo con el que este nodo aparece en la miga de pan y en el
               boton VOLVER de sus hijos.
    `grupo`    rotulo de la banda bajo la que el PADRE agrupa esta tarjeta.
               Solo lo usa NAV_SEC_II, que reparte sus catorce destinos en
               tres bandas; en el resto se usa `banda` del padre.
    `banda`    rotulo de la banda unica bajo la que este nodo agrupa a los
               hijos que no declaran `grupo`.
    `cargado`  False -> tarjeta marcador: gris, sin hipervinculo y sin clave.
               Es un rotulo de documento -cero dato normativo, asi que no roza
               la Regla n.1- y existe para que un nivel con un solo hijo
               cargado siga explicando la taxonomia en vez de parecer una
               pantalla vacia, y para que se vea de un vistazo que falta.
    """
    titulo: str
    corto: str = ""
    subtitulo: str = ""
    lineas: tuple = ()
    hoja: str | None = None
    destino: str | None = None
    grupo: str = ""
    banda: str = ""
    hijos: tuple = ()
    cargado: bool = True
    nota: str = ""          # linea al pie de la hoja NAV_*, sin enlace


def _hoja_final(titulo, l1, l2, destino, grupo=""):
    """Tarjeta que abre una hoja real del libro: es una hoja del arbol."""
    return Nodo(titulo=titulo, lineas=(l1, l2), destino=destino, grupo=grupo)


def _marcador(titulo, l1, l2):
    """Tarjeta marcador: documento que existe en la norma y no en este libro."""
    return Nodo(titulo=titulo, lineas=(l1, l2), cargado=False)


# Sub-bandas de la hoja SEC. II. Es el unico nivel que agrupa, porque sus
# destinos no son homogeneos ni siquiera dentro de una misma banda: en BASES DE
# DATOS conviven el volcado integro de las Partes A, B y C con cinco hojas
# transversales a todas ellas. Sin las sub-bandas serian nueve tarjetas
# indistinguibles.
G_ABC = "PARTES A, B y C · volcado integro, una fila por fila impresa"
G_TRV = "TRANSVERSAL A LAS PARTES"

SIN_EXTRAER = "Sin extraccion en resources/"

# Rotulos de las tres bandas del Dashboard. El tipo de artefacto se decide
# ANTES que la norma: es la primera pregunta de quien abre el libro.
B_CALCULO = "1 · MOTORES DE CALCULO — dimensionan una reparacion"
B_BUSQUEDA = "2 · MOTORES DE BUSQUEDA — consultan un valor tabulado"
B_DATOS = "3 · BASES DE DATOS — especificaciones de material, sin buscador"
B_MANUAL = "4 · TRANSVERSAL A TODA NORMA"

T_CAL, T_BUS, T_DAT = "MOTORES DE CALCULO", "MOTORES DE BUSQUEDA", "BASES DE DATOS"


def _rama_bpvc(pre, tipo, sub_secii, lineas_secii, hijos_secii, nota=""):
    """La cascada PRESSURE VESSELS > BPVC > SEC. II, que dos bandas comparten.

    Los buscadores de la Parte D y las nueve hojas de datos de las Partes A, B
    y C viven en bandas distintas, asi que cada una necesita su propia rama de
    hojas NAV_* -no se cruzan por diseno-. Pero los tres niveles de arriba son
    literalmente la misma cita del codigo, y escribirlos dos veces era
    garantizar que un dia dijeran cosas distintas. Solo cambian el prefijo de
    las hojas, el rotulo del tipo y lo que cuelga del ultimo nivel.
    """
    return Nodo(
        titulo="PRESSURE VESSELS",
        corto="PRESSURE VESSELS",
        subtitulo=f"{tipo} · Recipientes a presion · elija el codigo",
        lineas=("Boiler and Pressure Vessel Code",
                "Materiales y reglas de diseno"),
        hoja=f"NAV_{pre}_PVESSELS",
        banda="CODIGOS DE LA DISCIPLINA",
        hijos=(
            Nodo(
                titulo="BPVC · BOILER AND PRESSURE VESSEL CODE",
                corto="BPVC",
                subtitulo=f"{tipo} · ASME BPVC · elija la seccion",
                lineas=("Seccion II materiales · Seccion VIII diseno",
                        "Solo la Seccion II esta cargada"),
                hoja=f"NAV_{pre}_BPVC",
                banda="SECCIONES DEL BPVC",
                hijos=(
                    Nodo(
                        titulo="SEC. II · MATERIALES",
                        corto="SEC. II",
                        subtitulo=f"{tipo} · BPVC Seccion II 2025",
                        lineas=lineas_secii,
                        hoja=f"NAV_{pre}_SEC_II",
                        banda=sub_secii,
                        hijos=hijos_secii,
                        nota=nota),
                    _marcador("SEC. VIII DIV. 1 · RECIPIENTES A PRESION",
                              "Reglas de diseno por formula", SIN_EXTRAER),
                )),
        ))


ARBOL = Nodo(
    titulo="MOTOR DE CALCULO ASME PCC          Rev. 4",
    corto="DASHBOARD",
    subtitulo="Ingenieria de reparacion · ASME PCC-2 · B31.3-2024 · BPVC VIII-1 "
              "con II-D 2025 · unidades SI",
    hoja=DASH,
    hijos=(
        # ---- 1 · MOTORES DE CALCULO -------------------------------------
        Nodo(
            titulo="ASME",
            corto="ASME · CALCULO",
            subtitulo=f"{T_CAL} · American Society of Mechanical Engineers · "
                      f"elija la disciplina",
            lineas=("Motores que dimensionan una reparacion",
                    "Hoy solo la familia Post Construction (PCC)"),
            hoja="NAV_CAL_ASME",
            grupo=B_CALCULO,
            banda="DISCIPLINAS CON MOTOR DE CALCULO",
            hijos=(
                Nodo(
                    titulo="REPARACIONES",
                    corto="REPARACIONES",
                    subtitulo=f"{T_CAL} · Reparacion de equipos en servicio · "
                              f"elija el codigo",
                    lineas=("Post Construction Code (PCC)",
                            "Reparacion, pernos y evaluacion de aptitud"),
                    hoja="NAV_CAL_REPARACION",
                    banda="CODIGOS DE LA DISCIPLINA",
                    hijos=(
                        Nodo(
                            titulo="PCC · POST CONSTRUCTION CODE",
                            corto="PCC",
                            subtitulo=f"{T_CAL} · Familia ASME PCC · elija el documento",
                            lineas=("PCC-1 pernos · PCC-2 reparacion · PCC-3 riesgo",
                                    "Solo PCC-2 esta cargado"),
                            hoja="NAV_CAL_PCC",
                            banda="DOCUMENTOS DE LA FAMILIA PCC",
                            hijos=(
                                Nodo(
                                    titulo="PCC-2 · REPARACION DE EQUIPOS A PRESION",
                                    corto="PCC-2",
                                    subtitulo=f"{T_CAL} · ASME PCC-2 · articulos "
                                              f"cargados en el libro",
                                    lineas=("Metodos de reparacion por articulo",
                                            "Art. 212 y Art. 206 cargados"),
                                    hoja="NAV_CAL_PCC2",
                                    banda="ARTICULOS CARGADOS",
                                    hijos=(
                                        # Fase 9: cada articulo pasa de ser una
                                        # tarjeta que abre el motor a ser un NIVEL
                                        # con los documentos de ese articulo. Es
                                        # lo que pedia el plan —las hojas nuevas
                                        # cuelgan del nodo del motor— y un nodo
                                        # tiene `hoja` o `destino`, nunca las dos:
                                        # para tener hijos, el articulo necesita
                                        # su propia hoja NAV_*. El motor sigue a
                                        # un clic desde su pantalla, y desde el
                                        # motor hay boton directo a cada
                                        # documento (no hay que subir para
                                        # cambiar de pestana).
                                        Nodo(
                                            titulo="ART. 212 · PARCHE DE PLANCHA",
                                            corto="ART. 212",
                                            subtitulo=f"{T_CAL} · ASME PCC-2 Art. 212 "
                                                      f"(Fillet Welded Patches)",
                                            lineas=("Parche con soldadura de filete",
                                                    "Motor, especificaciones e "
                                                    "instrucciones"),
                                            hoja="NAV_CAL_ART212",
                                            banda="DOCUMENTOS DE ESTE ARTICULO",
                                            hijos=(
                                                _hoja_final(
                                                    "MOTOR DE CALCULO",
                                                    "Diseno del parche, Art. 212",
                                                    "Tuberia B31.3 · virola BPVC VIII-1",
                                                    "Parche_PCC2_Art212"),
                                                _hoja_final(
                                                    "ESPECIFICACIONES TECNICAS",
                                                    "Fabricacion, examen y prueba",
                                                    "212-3.4 · 212-4 · 212-5 · 212-6",
                                                    ESPEC_212),
                                                _hoja_final(
                                                    "INSTRUCCIONES DE USO",
                                                    "Que se rellena y que calcula cada celda",
                                                    "Guia celda a celda del motor",
                                                    INSTR_212),
                                            )),
                                        Nodo(
                                            titulo="ART. 206 · COLLAR DE ENCIERRO TOTAL",
                                            corto="ART. 206",
                                            subtitulo=f"{T_CAL} · ASME PCC-2 Art. 206 "
                                                      f"(Full Encirclement Sleeves)",
                                            lineas=("Sleeve de refuerzo, Type A/B",
                                                    "Motor, especificaciones e "
                                                    "instrucciones"),
                                            hoja="NAV_CAL_ART206",
                                            banda="DOCUMENTOS DE ESTE ARTICULO",
                                            hijos=(
                                                _hoja_final(
                                                    "MOTOR DE CALCULO",
                                                    "Diseno del collar, Art. 206",
                                                    "Tuberia B31.3 · Type A y Type B",
                                                    "Collar_PCC2_Art206"),
                                                _hoja_final(
                                                    "ESPECIFICACIONES TECNICAS",
                                                    "Fabricacion, examen y prueba",
                                                    "206-2 · 206-4 · 206-5 · 206-6",
                                                    ESPEC_206),
                                                _hoja_final(
                                                    "INSTRUCCIONES DE USO",
                                                    "Que se rellena y que calcula cada celda",
                                                    "Guia celda a celda del motor",
                                                    INSTR_206),
                                            )),
                                        _marcador(
                                            "RESTO DE ARTICULOS DE PCC-2",
                                            "Manguitos, envolventes, obturaciones",
                                            SIN_EXTRAER),
                                    )),
                                _marcador("PCC-1 · APRIETE DE UNIONES BRIDADAS",
                                          "Montaje y apriete de pernos", SIN_EXTRAER),
                                _marcador("PCC-3 · INSPECCION BASADA EN RIESGO",
                                          "Planificacion de inspeccion", SIN_EXTRAER),
                            )),
                    )),
            )),

        # ---- 2 · MOTORES DE BUSQUEDA ------------------------------------
        Nodo(
            titulo="ASME",
            corto="ASME · BUSQUEDA",
            subtitulo=f"{T_BUS} · American Society of Mechanical Engineers · "
                      f"elija la disciplina",
            lineas=("Consulta de valores tabulados por el codigo",
                    "Tuberia B31.3 y materiales BPVC Seccion II"),
            hoja="NAV_BUS_ASME",
            grupo=B_BUSQUEDA,
            banda="DISCIPLINAS",
            hijos=(
                Nodo(
                    titulo="PIPING",
                    corto="PIPING",
                    subtitulo=f"{T_BUS} · Tuberia a presion · elija el codigo",
                    lineas=("Codigo B31 de tuberia a presion",
                            "Bridas, accesorios y valvulas B16"),
                    hoja="NAV_BUS_PIPING",
                    banda="CODIGOS DE LA DISCIPLINA",
                    hijos=(
                        Nodo(
                            titulo="B31 · CODIGO DE TUBERIA A PRESION",
                            corto="B31",
                            subtitulo=f"{T_BUS} · ASME B31 · elija la seccion",
                            lineas=("Una seccion por servicio de tuberia",
                                    "Solo B31.3 esta cargado"),
                            hoja="NAV_BUS_B31",
                            banda="SECCIONES DEL B31",
                            hijos=(
                                Nodo(
                                    titulo="B31.3 · PROCESS PIPING",
                                    corto="B31.3",
                                    subtitulo=f"{T_BUS} · ASME B31.3-2024 · tablas y "
                                              f"apendices cargados en el libro",
                                    lineas=("Tuberia de proceso · edicion 2024",
                                            "Apendices A, B y C y factores de calidad"),
                                    hoja="NAV_BUS_B31_3",
                                    banda="TABLAS Y APENDICES CARGADOS",
                                    hijos=(
                                        _hoja_final(
                                            "B31.3 · TABLAS A-1 y A-4",
                                            "Esfuerzo admisible S",
                                            "MPa (SI) y ksi (US)", "Buscar_B31_3"),
                                        _hoja_final(
                                            "B31.3 · APENDICE C",
                                            "Propiedades fisicas: dilatacion y modulo",
                                            "Metales y no metalicos · SI y US",
                                            "Buscar_Prop_B31_3"),
                                        _hoja_final(
                                            "B31.3 · TABLA B-1",
                                            "Esfuerzo de diseno hidrostatico HDS",
                                            "Tuberia termoplastica · MPa (SI) y ksi (US)",
                                            "Buscar_B31_B1"),
                                        _hoja_final(
                                            "B31.3 · TABLA A-2",
                                            "Factor de calidad de fundicion Ec",
                                            "Basico y con examen suplementario",
                                            "Buscar_Ec_A2"),
                                        _hoja_final(
                                            "B31.3 · TABLA A-3",
                                            "Factor de calidad de junta longitudinal Ej",
                                            "Por tipo de junta soldada", "Buscar_Ej_A3"),
                                        # Ultimo del nivel, a proposito: lo
                                        # cargado va primero y el marcador
                                        # cierra. Se retiro en la Rev. 4d por
                                        # decision de alcance; la extraccion
                                        # sigue intacta en resources/ y lo que
                                        # no se carga es la hoja.
                                        _marcador(
                                            "B31.3 · APENDICE B (B-2 a B-6)",
                                            "Presiones admisibles y listados de spec.",
                                            "Concreto, vidrio borosilicato y PEX-AL-PEX"),
                                    )),
                                _marcador("B31.1 · B31.4 · B31.5",
                                          "Potencia, hidrocarburos liquidos, refrigeracion",
                                          SIN_EXTRAER),
                                _marcador("B31.8 · B31.9 · B31.12",
                                          "Gas, servicios de edificio e hidrogeno",
                                          SIN_EXTRAER),
                            )),
                        _marcador("B16 · BRIDAS, ACCESORIOS Y VALVULAS",
                                  "B16.5, B16.9, B16.34 y familia", SIN_EXTRAER),
                    )),
                _rama_bpvc(
                    "BUS", T_BUS,
                    sub_secii="PARTE D · propiedades de diseno",
                    lineas_secii=("Parte D: esfuerzos admisibles y propiedades",
                                  "Cinco motores de busqueda"),
                    hijos_secii=(
                        _hoja_final("BPVC II-D · TABLA 1A",
                                    "Esfuerzo admisible S, ferrosos",
                                    "MPa (SI) y ksi (US)", "Buscar_BPVC_IID"),
                        _hoja_final("BPVC II-D · TABLAS 1B y 3",
                                    "Esfuerzo admisible S, no ferrosos",
                                    "MPa (SI) y ksi (US)", "Buscar_BPVC_IID_B"),
                        _hoja_final("BPVC II-D · TABLA U",
                                    "Resistencia a la traccion Su",
                                    "MPa (SI) y ksi (US)", "Buscar_Su"),
                        _hoja_final("BPVC II-D · TABLA Y-1",
                                    "Limite de fluencia Sy",
                                    "MPa (SI) y ksi (US)", "Buscar_Sy"),
                        _hoja_final("BPVC II-D · TM, TE y PRD",
                                    "Modulo E, dilatacion, Poisson y densidad",
                                    "por grupo de material · ver MAP_Grupo",
                                    "Buscar_Prop_IID"),
                    ),
                    # No es un marcador: las Partes A, B y C SI estan cargadas.
                    # Lo que no tienen es buscador, y por eso viven en la banda
                    # 3 del Dashboard. Decirlo aqui evita que se busquen en la
                    # rama equivocada; enlazarlas cruzaria las dos bandas, que
                    # es justo lo que esta separacion evita.
                    nota="Las Partes A, B y C no llevan buscador: su volcado integro "
                         "esta en la banda 3 · BASES DE DATOS del Dashboard."),
            )),

        # ---- 3 · BASES DE DATOS ------------------------------------------
        Nodo(
            titulo="ASME",
            corto="ASME · DATOS",
            subtitulo=f"{T_DAT} · American Society of Mechanical Engineers · "
                      f"elija la disciplina",
            lineas=("Volcado integro de especificaciones de material",
                    "Hojas de datos: sin buscador y sin una sola formula"),
            hoja="NAV_DAT_ASME",
            grupo=B_DATOS,
            banda="DISCIPLINAS",
            hijos=(
                _rama_bpvc(
                    "DAT", T_DAT,
                    sub_secii="",
                    lineas_secii=("Partes A, B y C: 379 especificaciones",
                                  "54 198 filas tal como estan impresas"),
                    hijos_secii=(
                        _hoja_final("SEC. II · PARTE A vol. 1",
                                    "Volcado integro de SA-6 a SA-450",
                                    "Una fila por fila impresa", "DB_SecII_A1",
                                    grupo=G_ABC),
                        _hoja_final("SEC. II · PARTE A vol. 2",
                                    "Volcado integro de SA-451 en adelante",
                                    "Una fila por fila impresa", "DB_SecII_A2",
                                    grupo=G_ABC),
                        _hoja_final("SEC. II · PARTE B", "No ferrosos: SB-",
                                    "Una fila por fila impresa", "DB_SecII_B",
                                    grupo=G_ABC),
                        _hoja_final("SEC. II · PARTE C",
                                    "Consumibles de soldadura: SFA-",
                                    "Una fila por fila impresa", "DB_SecII_C",
                                    grupo=G_ABC),
                        _hoja_final("SEC. II · CATALOGO",
                                    "379 entradas: specs, paginas PDF y folios",
                                    "Partes A (2 vol.), B y C · 2025",
                                    "CAT_SecII", grupo=G_TRV),
                        _hoja_final("SEC. II · INDICE DE TABLAS",
                                    "2 572 tablas logicas y su reparto",
                                    "Dice cuanto quedo tabulado y por que",
                                    "IDX_SecII_Tablas", grupo=G_TRV),
                        _hoja_final("SEC. II · NOTAS AL PIE",
                                    "Notas de tabla con su marcador",
                                    "Restringen lo que dice la tabla",
                                    "DB_SecII_Notas", grupo=G_TRV),
                        _hoja_final("SEC. II · QUIMICA",
                                    "Composicion normalizada por elemento",
                                    "Solo tablas de encabezado resuelto",
                                    "DB_SecII_Quimica", grupo=G_TRV),
                        _hoja_final("SEC. II · TRACCION",
                                    "Rm, Re, alargamiento y dureza",
                                    "Solo tablas de encabezado resuelto",
                                    "DB_SecII_Traccion", grupo=G_TRV),
                    ),
                    nota="El 45,8 % de las filas llega marcada AMBIGUA: conserva su "
                         "texto impreso ENTERO en la celda C01, pero no quedo "
                         "repartida en columnas. IDX_SecII_Tablas lo dice tabla a "
                         "tabla."),
            )),

        # ---- 4 · TRANSVERSAL --------------------------------------------
        _hoja_final("MANUAL DE USO", "Convenciones, alcance y limitaciones",
                    "Leer antes de calcular", "Instrucciones", grupo=B_MANUAL),
    ))


def _preorden(nodo):
    """Recorre el arbol en preorden. Es el orden en que se derivan las listas."""
    yield nodo
    for h in nodo.hijos:
        yield from _preorden(h)


def _mapa_padre(nodo, hoja_padre, out):
    propia = nodo.hoja or nodo.destino
    if propia and hoja_padre:
        out[propia] = hoja_padre
    for h in nodo.hijos:
        _mapa_padre(h, nodo.hoja or hoja_padre, out)


# Dashboard + las diez hojas NAV_*. Son las que construye build_arbol().
HOJAS_NAV = [n.hoja for n in _preorden(ARBOL) if n.hoja]

# Las hojas del arbol que ya existen en el libro (motores, buscadores y datos).
DESTINOS = [n.destino for n in _preorden(ARBOL) if n.destino]

# Hojas que el usuario puede llegar a abrir. Todo lo demas queda veryHidden.
# DEBE coincidir, en contenido y EN ORDEN, con HojasNavegables() de
# vba/mod_nav.vba. El Dashboard no entra: es la unica hoja visible.
NAVEGABLES = [n.hoja or n.destino for n in _preorden(ARBOL)
              if (n.hoja or n.destino) and n.hoja != DASH]

# Padre de cada hoja alcanzable. Sustituye a la antigua clave literal "VOLVER":
# con un arbol de cinco niveles, subir tiene que llevar al PADRE y no a la raiz,
# asi que la celda oculta guarda SIEMPRE el nombre de la hoja destino, se este
# bajando o subiendo. El VBA queda con una sola rama y deja de crecer con el
# arbol.
PADRE: dict[str, str] = {}
_mapa_padre(ARBOL, None, PADRE)

# Rotulo corto de cada hoja NAV_*, para la miga de pan y el boton de retorno.
ROTULO = {n.hoja: (n.corto or n.titulo) for n in _preorden(ARBOL) if n.hoja}

# Celda donde cada hoja DESTINO lleva su enlace de retorno. Se elige por hoja
# porque los tres tipos de layout difieren: los buscadores y el motor fusionan
# la fila 1 (y la 2) y congelan en A4, dejando la fila 3 libre; Instrucciones
# no congela y empieza en B2, dejando libre la fila 1. Las hojas NAV_* no
# entran aqui: build_nav() escribe su propia barra de acciones.
ANCLA_VOLVER = {n: (3, 1, 3) for n in DESTINOS}    # (fila, col_ini, col_fin)
ANCLA_VOLVER["Instrucciones"] = (1, 2, 3)
# Las nueve hojas de la Seccion II dejan libre la fila 3 a proposito (ver
# R_HDR_SECII): asi anclan el boton donde lo anclan los buscadores, sin pisar
# la fila de encabezados.

DASH_NCOLS = 12
DASH_ANCHO_COL = 15
BTN_FILL = PatternFill("solid", fgColor=TINTA)
BTN_F = Font(name=MONO, size=10, bold=True, color=PAPEL)
CARD_TIT_F = Font(name=MACRO, size=10, color=TINTA)
CARD_TXT_F = Font(name=MONO, size=9, color=TINTA)
# Aviso de macros: bloque de peligro macizo, tinta sobre el rojo del acento. Es
# el unico relleno rojo del libro, y por eso no se confunde con nada.
AVISO_ROJO_F = Font(name=MONO, size=11, bold=True, color=PAPEL)
AVISO_ROJO_FILL = PatternFill("solid", fgColor=ROJO)
PIE_F = Font(name=MONO, size=8, color=GRIS)
PIE_FILL = PatternFill("solid", fgColor=PAPEL_2)
KPI_AMBAR_F = Font(name=MACRO, size=16, color=AMBAR_TXT)

# Tarjeta marcador: trama de medio tono, sin hipervinculo y sin clave. El estado
# inactivo es lo unico que usa los dos grises del sistema.
MARK_FILL = PatternFill("solid", fgColor=PAPEL)
MARK_BAR_FILL = PatternFill("solid", fgColor=GRIS_2)
MARK_TIT_F = Font(name=MACRO, size=10, color=GRIS)
MARK_TXT_F = Font(name=MONO, size=9, color=GRIS)
MARK_BAR_F = Font(name=MONO, size=10, bold=True, color=GRIS)
MARK_BORDER = BOX
TXT_NO_CARGADO = "/// NO CARGADO EN ESTE LIBRO"

# Barra inferior de la tarjeta. El texto distingue el nivel: ENTRAR baja un
# nivel del arbol, ABRIR llega a la hoja final. Los signos son ASCII y no
# tipografia decorativa: sobreviven a cualquier codificacion y a cualquier
# fuente que Excel decida sustituir.
TXT_ENTRAR = ">>> ENTRAR"
TXT_ABRIR = ">>> ABRIR"
TXT_MANUAL = "[ ? ] MANUAL DE USO"

# Miga de pan: fila 4 de toda hoja NAV_*.
MIGA_FILL = PatternFill("solid", fgColor=PAPEL_2)
MIGA_F = Font(name=MONO, size=9, bold=True, color=ROJO, underline="single")
MIGA_AQUI_F = Font(name=MONO, size=9, bold=True, color=TINTA)

# Filas fijas de una hoja NAV_*. El Dashboard NO usa FILA_ACCIONES ni
# FILA_MIGA: es la raiz (no tiene padre) y su fila 4 la ocupa el aviso de
# macros, cuya direccion esta acoplada a CELDA_AVISO = "A4" del VBA.
FILA_ACCIONES = 3
FILA_MIGA = 4
FILA_PRIMERA_BANDA = 6


def _enlace(ws, cel, fila, columna, clave):
    """Hipervinculo senuelo + clave de destino en la celda oculta.

    El hipervinculo apunta siempre a Dashboard!A1 y no navega por si mismo:
    solo existe para que Excel dispare Workbook_SheetFollowHyperlink. Un
    hipervinculo directo a la hoja destino seria invalido, porque la hoja
    esta oculta cuando se hace clic.

    LLAMAR ANTES DE APLICAR NINGUN ESTILO a la celda: openpyxl le asigna el
    estilo "Hyperlink" (azul subrayado) si aun no tiene uno propio, y taparia
    el boton o la tarjeta.
    """
    texto = "" if cel.value is None else str(cel.value)
    cel.hyperlink = Hyperlink(ref=cel.coordinate, location=f"'{DASH}'!A1", display=texto)
    _celda_clave(ws, fila, columna).value = clave
    return cel


def _boton(ws, fila, c1, c2, texto, clave):
    """Celda-boton con hipervinculo inocuo y la clave de destino en COL_CLAVE."""
    b = _mrg(ws, fila, c1, c2, texto)
    _enlace(ws, b, fila, c1, clave)
    # El estilo se aplica DESPUES del hipervinculo (ver _enlace).
    b.font, b.fill = BTN_F, BTN_FILL
    b.alignment = Alignment(horizontal="center", vertical="center")
    return b


def _celda_clave(ws, fila, columna):
    """Celda oculta con la clave de destino del boton anclado en (fila, columna).

    Lleva la fuente de dato aunque no se vea nunca: asi ninguna celda con
    contenido del libro se queda fuera del sistema visual, y la auditoria de
    test_dashboard.py puede exigirlo sin excepciones que haya que recordar.
    """
    cel = ws.cell(fila, COL_CLAVE_BASE + columna)
    cel.font = DATA_F
    return cel


def _ocultar_columnas_clave(ws, cols_boton):
    for c in sorted({COL_CLAVE_BASE + c for c in cols_boton}):
        ws.column_dimensions[get_column_letter(c)].hidden = True


def _tarjeta(ws, fila, c1, ancho, titulo, lineas, clave,
             pie=TXT_ABRIR, cargado=True):
    """Tarjeta de 4 filas: titulo, dos lineas de detalle y la barra de accion.

    La tarjeta ENTERA es clicable, no solo la barra inferior: las cuatro filas
    llevan hipervinculo y su propia celda de clave, en (fila+i, COL_CLAVE_BASE
    + c1). Son cuatro celdas distintas -difieren en la fila-, asi que las
    claves siguen sin pisarse. La barra se conserva como senal visual de que el
    bloque navega, y su texto dice adonde: ENTRAR baja un nivel del arbol,
    ABRIR llega a la hoja final.

    `cargado=False` da la tarjeta marcador: gris, sin hipervinculo y sin clave,
    rotulada NO CARGADO EN ESTE LIBRO.
    """
    c2 = c1 + ancho - 1
    detalle = list(lineas[:2]) + [""] * (2 - len(lineas[:2]))
    # El titulo va en mayusculas —es el rotulo de la tarjeta, macrotipografia—
    # y las dos lineas de detalle se dejan tal como estan escritas: son la
    # unica prosa de la capa de navegacion y en versalitas se leen peor.
    textos = [titulo.upper(), detalle[0], detalle[1],
              pie if cargado else TXT_NO_CARGADO]

    for i, texto in enumerate(textos):
        cel = _mrg(ws, fila + i, c1, c2, texto)
        if cargado:
            _enlace(ws, cel, fila + i, c1, clave)
        if i == 0:
            cel.font = CARD_TIT_F if cargado else MARK_TIT_F
            cel.alignment = Alignment(vertical="center", indent=1)
        elif i < 3:
            cel.font = CARD_TXT_F if cargado else MARK_TXT_F
            cel.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
        else:
            cel.font = BTN_F if cargado else MARK_BAR_F
            cel.fill = BTN_FILL if cargado else MARK_BAR_FILL
            cel.alignment = Alignment(horizontal="center", vertical="center")

    fondo = CARD_FILL if cargado else MARK_FILL
    borde = CARD_BORDER if cargado else MARK_BORDER
    for r in range(fila, fila + 4):
        for c in range(c1, c2 + 1):
            cel = ws.cell(r, c)
            if r != fila + 3:
                cel.fill = fondo
            cel.border = borde
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


def _cabecera_nav(ws, nodo):
    """Titulo y subtitulo: filas 1 y 2 de toda hoja del arbol, Dashboard incluido."""
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TINTA
    for c in range(1, DASH_NCOLS + 1):
        ws.column_dimensions[get_column_letter(c)].width = DASH_ANCHO_COL

    # Cabecera: dos filas de tinta maciza, la segunda cerrada por la franja
    # roja. El relleno lo cubre la celda ancla de cada fila fusionada; la franja
    # se recorre celda a celda (ver franja()).
    t = _mrg(ws, 1, 1, DASH_NCOLS, nodo.titulo.upper())
    t.font, t.fill = Font(name=MACRO, size=16, color=PAPEL), TITLE_FILL
    t.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 32
    s = _mrg(ws, 2, 1, DASH_NCOLS, nodo.subtitulo)
    s.font, s.fill = Font(name=MONO, size=9, color=PAPEL), TITLE_FILL
    s.alignment = Alignment(vertical="center", indent=1)
    franja(ws, 2, 1, DASH_NCOLS)
    ws.row_dimensions[2].height = 16


def _texto_volver(hoja_padre):
    return ("<<< VOLVER AL DASHBOARD" if hoja_padre == DASH
            else f"<<< VOLVER A {ROTULO[hoja_padre].upper()}")


def _miga(ws, fila, ruta, nodo):
    """Miga de pan clicable: cada ancestro enlaza con su propia hoja.

    Sale gratis con el mismo mecanismo de clave que las tarjetas, y evita que
    subir cuatro niveles cueste cuatro clics. El ultimo segmento es el nodo
    actual: se rotula pero no se enlaza, porque ya se esta en el.
    """
    cols = []
    segmentos = [(n.hoja, ROTULO[n.hoja]) for n in ruta] + [(None, ROTULO[nodo.hoja])]
    for i, (hoja, rotulo) in enumerate(segmentos):
        c1 = 1 + i * 2
        cel = _mrg(ws, fila, c1, c1 + 1,
                   rotulo.upper() if i == 0 else f"/ {rotulo.upper()}")
        if hoja:
            _enlace(ws, cel, fila, c1, hoja)
            cel.font = MIGA_F
            cols.append(c1)
        else:
            cel.font = MIGA_AQUI_F
        cel.alignment = Alignment(vertical="center", indent=1)
    for c in range(1, DASH_NCOLS + 1):
        ws.cell(fila, c).fill = MIGA_FILL
    ws.row_dimensions[fila].height = 16
    return cols


def _tarjetas_de(ws, r, nodo):
    """Dibuja los hijos de `nodo` en bandas de tres tarjetas.

    Cada hijo va bajo la banda que declara en `grupo`; si no declara ninguna,
    bajo la banda unica del padre (`banda`). Devuelve las columnas ancla usadas
    -para ocultar sus columnas de clave- y la primera fila libre.
    """
    grupos: list[tuple[str, list]] = []
    for h in nodo.hijos:
        rotulo = h.grupo or nodo.banda
        if not grupos or grupos[-1][0] != rotulo:
            grupos.append((rotulo, []))
        grupos[-1][1].append(h)

    cols = set()
    for rotulo, hijos in grupos:
        if rotulo:
            banda(ws, r, rotulo, DASH_NCOLS)
            r += 1
        for i, h in enumerate(hijos):
            # El salto de fila va ANTES de dibujar: la primera terna usa la
            # fila que dejo banda().
            if i and i % 3 == 0:
                r += 5
            c1 = 1 + (i % 3) * 4
            cols.add(c1)
            _tarjeta(ws, r, c1, 4, h.titulo, h.lineas, h.hoja or h.destino,
                     pie=TXT_ENTRAR if h.hoja else TXT_ABRIR, cargado=h.cargado)
        r += 5
    return cols, r


def build_nav(wb, nodo, ruta):
    """Hoja NAV_* de un nodo interior. `ruta` son sus ancestros, raiz -> padre.

    Layout comun a todos los niveles del arbol: cabecera, barra de acciones,
    miga de pan y las tarjetas de sus hijos. El Dashboard comparte cabecera y
    tarjetas pero no barra ni miga: es la raiz -no tiene padre- y su fila 4 la
    ocupa el aviso de macros, acoplado a CELDA_AVISO = "A4" del VBA.
    """
    ws = wb.create_sheet(nodo.hoja)
    _cabecera_nav(ws, nodo)

    # Barra de acciones. El manual es transversal a las normas -no cuelga de
    # ningun codigo- y por eso esta siempre a un clic desde cualquier nivel,
    # ademas de tener su tarjeta en la raiz.
    padre = ruta[-1]
    _boton(ws, FILA_ACCIONES, 1, 3, _texto_volver(padre.hoja), padre.hoja)
    _boton(ws, FILA_ACCIONES, 10, 12, TXT_MANUAL, "Instrucciones")
    ws.row_dimensions[FILA_ACCIONES].height = 18

    cols = {1, 10} | set(_miga(ws, FILA_MIGA, ruta, nodo))
    ws.row_dimensions[FILA_MIGA + 1].height = 6

    cols_tarjetas, r = _tarjetas_de(ws, FILA_PRIMERA_BANDA, nodo)
    if nodo.nota:
        # Texto, sin enlace: dice donde esta lo que NO cuelga de esta rama.
        # Enlazarlo cruzaria dos bandas del Dashboard, que es justo lo que la
        # separacion por tipo de artefacto evita.
        n = _mrg(ws, r, 1, DASH_NCOLS, nodo.nota)
        n.font = SRC_F
        n.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
        ws.row_dimensions[r].height = 26
    _ocultar_columnas_clave(ws, cols | cols_tarjetas)
    return ws


def build_dashboard(wb, nodo, kpis, fecha):
    """Portada unica del libro y raiz del arbol.

    `kpis` es una lista de (titulo, valor, unidad, ambar). Los indicadores y el
    pie de responsabilidad se quedan aqui: el Dashboard sigue siendo la unica
    portada, aunque ya no sea un indice plano.
    """
    ws = wb.create_sheet(nodo.hoja)
    _cabecera_nav(ws, nodo)
    ws.row_dimensions[3].height = 6

    # Aviso de macros. Se graba en rojo: es el estado correcto para un archivo
    # en disco. Workbook_Open lo pasa a verde solo si las macros corren.
    av = _mrg(ws, FILA_AVISO, 1, DASH_NCOLS,
              "/// MACROS DESHABILITADAS - HABILITELAS PARA NAVEGAR ENTRE LOS MOTORES ///")
    av.font, av.fill = AVISO_ROJO_F, AVISO_ROJO_FILL
    av.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[FILA_AVISO].height = 22
    ws.row_dimensions[FILA_AVISO + 1].height = 6

    cols, r = _tarjetas_de(ws, FILA_PRIMERA_BANDA, nodo)
    _ocultar_columnas_clave(ws, cols)

    # --- estado del libro -------------------------------------------------
    banda(ws, r, "5 · ESTADO DEL LIBRO", DASH_NCOLS)
    r += 1
    # Cuatro KPI por banda: el Dashboard tiene DASH_NCOLS columnas y cada KPI
    # ocupa tres. El quinto se salia del ancho de la hoja y quedaba invisible.
    POR_BANDA = DASH_NCOLS // 3
    for i, (titulo, valor, unidad, ambar) in enumerate(kpis):
        if i and i % POR_BANDA == 0:
            r += 4
        c1 = 1 + (i % POR_BANDA) * 3
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
        # Las alturas se fijan DENTRO del bucle: con dos bandas de KPI, ponerlas
        # despues solo alcanzaba a la ultima y la primera quedaba encogida.
        ws.row_dimensions[r].height = 26
        ws.row_dimensions[r + 1].height = 34
        ws.row_dimensions[r + 2].height = 14
    r += 3
    n = _mrg(ws, r, 1, DASH_NCOLS,
             f"COMPILADO {fecha} // FUENTE UNICA: resources/ // valores cargados tal "
             f"como estan impresos en el codigo; SI y US son extracciones "
             f"independientes, nunca conversiones.")
    n.font = SRC_F
    n.alignment = Alignment(vertical="center", indent=1)
    r += 2

    # --- pie --------------------------------------------------------------
    p = _mrg(ws, r, 1, DASH_NCOLS,
             "Herramienta de ingenieria de referencia. Verificar entradas y resultados, y "
             "leer las notas del material, antes de emitir para construccion.")
    p.font, p.fill = PIE_F, PIE_FILL
    p.alignment = Alignment(vertical="center", indent=1)
    # El pie de responsabilidad cierra el Dashboard con la misma franja roja con
    # la que abren las bandas: es el limite inferior del documento.
    franja(ws, r, 1, DASH_NCOLS, arriba=True)
    ws.row_dimensions[r].height = 18
    return ws


def build_arbol(wb, kpis, fecha):
    """Crea el Dashboard y las diez hojas NAV_* recorriendo el arbol en preorden.

    Devuelve HOJAS_NAV: el orden en que main() las coloca en el libro.
    """
    def recorrer(nodo, ruta):
        if not nodo.hoja:                       # tarjeta destino o marcador
            return
        if nodo.hoja == DASH:
            build_dashboard(wb, nodo, kpis, fecha)
        else:
            build_nav(wb, nodo, ruta)
        for h in nodo.hijos:
            recorrer(h, ruta + [nodo])

    recorrer(ARBOL, [])
    return HOJAS_NAV


def link_volver(wb):
    """Escribe el enlace de retorno en cada hoja DESTINO del arbol.

    La clave que se graba es la hoja PADRE, no un literal "VOLVER": subir un
    nivel es navegar como cualquier otro, y por eso el VBA tiene una sola rama.
    Las hojas NAV_* no pasan por aqui: build_nav() escribe su propia barra de
    acciones, con el mismo mecanismo.

    La celda se desbloquea explicitamente: Parche_PCC2_Art212 e Instrucciones
    se entregan protegidas, y aunque Excel permite seguir un hipervinculo en
    celda bloqueada, dejarla desbloqueada evita depender de ese detalle.
    """
    for nombre in DESTINOS:
        if nombre not in wb.sheetnames:
            continue
        ws = wb[nombre]
        fila, c1, c2 = ANCLA_VOLVER[nombre]
        padre = PADRE[nombre]
        b = _boton(ws, fila, c1, c2, _texto_volver(padre), padre)
        b.protection = Protection(locked=False)
        _celda_clave(ws, fila, c1).protection = Protection(locked=False)
        _ocultar_columnas_clave(ws, (c1,))


# ---------------------------------------------------------------------------
# Retonado de lo que trae el maestro
# ---------------------------------------------------------------------------
# Una sola hoja no la construye este script: la trae maestro_con_macros.xlsm
# (Instrucciones). Su estilo es el de la Rev. 0 —azules corporativos, Arial,
# amarillo de entrada— y sin retonarla el libro tendria dos sistemas visuales a
# la vez. La plantilla NO se edita a mano (la genera make_vba_seed.py), asi que
# el mapeo se aplica al construir. (Datos_Ref tambien la traia el maestro, pero
# la Tarea 10 la retira del libro: su dato vivo esta en las bases auditadas.)
#
# Es una tabla de EQUIVALENCIA EXACTA, no una aproximacion por cercania de
# color: solo se traduce lo que estaba en el inventario de la plantilla. Asi es
# idempotente y no puede tocar por accidente una celda que ya nacio en el
# sistema nuevo —ningun color de la paleta nueva es clave de esta tabla—, y lo
# que quede sin traducir lo delata la auditoria (celdas_fuera_del_sistema).
HOJAS_HEREDADAS = ("Instrucciones",)

# Las doce hojas donde el ingeniero introduce datos. El guardia de verificar.py §5
# solo mira estas: una DB_*, MAP_* o NAV_* no lleva entradas y no le aplica.
HOJAS_DE_MOTOR = (MOTOR, COLLAR_MOTOR,
                  "Buscar_B31_3", "Buscar_BPVC_IID", "Buscar_BPVC_IID_B",
                  "Buscar_Su", "Buscar_Sy", "Buscar_Prop_IID",
                  "Buscar_Prop_B31_3", "Buscar_B31_B1",
                  "Buscar_Ec_A2", "Buscar_Ej_A3")

# Listas literales que SI pueden quedarse: son modos y booleanos del propio motor,
# no datos tabulados por un codigo. Todo lo demas que sea literal en una hoja de
# motor es una lista fija espejando una tabla — exactamente lo que la regla 12
# prohibe— y el guardia lo rechaza.
LITERALES_PERMITIDOS = {
    '"SI,US"': "Conmutador de unidades (regla 10). No es un dato tabulado.",
    '"Interpolado,Tabulado-conservador"': (
        "Modo de lectura de S(T): la interpolacion lineal la autoriza el propio "
        "codigo (p.ej. para. A302.3.1(b)); no es una tabla de valores."),
    '"Si,No"': "Booleano de examen UT / defecto circunferencial del Art. 206.",
    '"' + TIPO_A + ',' + TIPO_B + '"': (
        "Los dos tipos de sleeve que define el propio Art. 206 (206-1.1.1 / "
        "206-1.1.2)."),
    # D10 de los dos motores: conmuta MODO=1/2/3 (tuberia/virola/cabezal-esfera)
    # y por tanto el codigo de construccion y el kf de toda la hoja. Es un modo
    # del propio motor, no un dato tabulado por ningun codigo — misma clase que
    # el conmutador SI/US o Type A/Type B. Las dos hojas escriben el literal con
    # acentuacion distinta (divergencia cosmetica preexistente, no normativa: no
    # hay valor de codigo en juego), asi que entran las dos variantes.
    '"Tubería (B31.3),Virola cilíndrica (VIII-1),Cabezal/esfera (VIII-1)"': (
        "Selector de aplicacion/geometria del Art. 212 (D10): conmuta el modo "
        "geometrico y el codigo de construccion, no es un dato tabulado."),
    '"Tuberia (B31.3),Virola cilindrica (VIII-1),Cabezal/esfera (VIII-1)"': (
        "Idem, variante sin acentos usada en Collar_PCC2_Art206 (D10)."),
    # Entradas categoricas del Paso 1 del Art. 212 (compuerta de elegibilidad,
    # 212-1/212-2). No son materiales ni valores tabulados por ningun codigo:
    # son categorias del propio criterio de elegibilidad (mecanismo de dano,
    # tipo de servicio, estado de grietas), asi que su origen legitimo es una
    # lista de items, igual que el conmutador de modo o el booleano Si/No.
    '"Adelgazamiento local,Erosion,Corrosion,Perforacion traspasante,'
    'Otro/no caracterizado"': (
        "Mecanismo de dano del Paso 1 (Art. 212 D136): categoria de "
        "elegibilidad (212-1c/212-2c), no un material ni un dato tabulado."),
    '"General,Letal / extrema peligrosidad"': (
        "Tipo de servicio del Paso 1 (Art. 212 D138): categoria de elegibilidad "
        "(flujo aprobado; 212-2a remite a Part 1), no un dato tabulado."),
    '"No,Si — arrestada + FFS,Si — activa/no analizada"': (
        "Presencia de grietas del Paso 1 (Art. 212 D139): categoria de "
        "elegibilidad (212-2c [11]-[13]), no un dato tabulado."),
    '"Hidrostatica,Neumatica"': (
        "Tipo de prueba de hermeticidad del Paso 8 (Art. 212 D183): modo del "
        "motor (212-6), no un material ni un dato tabulado."),
    '"20,12,6,2"': (
        "Factor de consecuencia Rscaled del Paso 8 (Art. 212 D188): los cuatro "
        "valores de la Tabla 501-III-1-1, leidos de resources/ por "
        "leer_energia_501 (Fase 0.2). Lista de items con sus valores (el plan "
        "lo permite); si el codigo cambiara la tabla, esta lista y este literal "
        "cambian a la vez."),
    '"Prueba del anular,Prueba sensible de fugas,No requerida"': (
        "Tipo de prueba de hermeticidad del Paso 8 del Art. 206 (D147): las dos "
        "vias del 206-6 (a)/(b) mas 'No requerida'. Categoria de la prueba, no "
        "un material ni un dato tabulado."),
}

# Deuda SALDADA (Tarea 10). Las Tareas 7-9 repuntaron NPS y cedula de los dos
# motores contra DB_B36_10/DB_B36_19 — Art. 212 (MOTOR) en las Tareas 7-8, Art.
# 206 (COLLAR_MOTOR) en la Tarea 9 —: sus D18/D19 salen de DB_B36 por cascada y
# ninguno es ya una lista fija. El dict se conserva VACIO (no se elimina) porque
# el guardia de verificar.py §5 y test_dashboard lo consultan por nombre: vacio,
# el guardia no tolera ninguna lista fija, que es el estado definitivo buscado.
DEUDA_LISTA_FIJA = {}

# Celdas que el *oracle* del Art. 212 declara pero que el build ya NO reproduce a
# proposito. El oracle sigue siendo la hoja heredada tal como se capturo; esta lista
# es la unica forma declarada de apartarse de ella, y cada entrada lleva su motivo.
# Sin esto, la alternativa era regenerar el oracle desde el build nuevo — con lo que
# dejaria de ser un control independiente y pasaria a ser un volcado de si mismo.
# Dos clases de divergencia contra el oracle Rev0, con semantica distinta y
# disjunta (TestParidadHojaParche exige que ninguna celda este en las dos):
#   DIVERGENCIAS_DECLARADAS  -> la celda se RETIRA: el build la deja VACIA (None).
#   DIVERGENCIAS_REEMPLAZADAS -> la celda se REEMPLAZA por contenido nuevo a
#       proposito: sale de la igualdad contra el oracle, se EXIGE no-vacia, y su
#       correccion la prueba otro test nombrado (no el oracle Rev0). Sin esta
#       segunda clase, la unica forma de apartarse del oracle era vaciar la celda,
#       y un rediseño sancionado (lista fija -> cascada de base) no la vacia.
DIVERGENCIAS_DECLARADAS = {
    ("Parche_PCC2_Art212", "A24"): (
        "Fila retirada: el material del parche salia de una lista fija (regla 12); "
        "lo resuelve la cascada de la Seccion 7, no una lista fija. La fila 23 "
        "queda vacia (la 22 la ocupa ahora el selector de norma dimensional)."),
    ("Parche_PCC2_Art212", "B24"): "Idem A23: la fila entera se retira.",
    ("Parche_PCC2_Art212", "C24"): "Idem A23: la fila entera se retira.",
    ("Parche_PCC2_Art212", "D24"): "Idem A23: la fila entera se retira.",
    ("Parche_PCC2_Art212", "G24"): "Idem A23: la fila entera se retira.",
}

# --- Fase 2: el modelo de presion pasa de TRES casos a DOS ------------------
# El motor evaluaba Operacion / "Diseno tipico" / "Envolvente". Ese caso
# intermedio no lo pide el codigo: 212-3.2 define una UNICA P = "internal
# design pressure" para las ec. (1)/(2), y el Art. 206 -que comparte modelo de
# presion en este libro- es explicito en 206-3.3, "the maximum allowable design
# pressure" (los dos comprobados en resources/, Regla n.1). Con las dos
# columnas habia dos presiones compitiendo por gobernar el t_req, que es
# justo la ambiguedad que un motor de calculo no puede tener.
#
# Se retira la fila 27 entera y la columna F de la tabla de cargas (filas
# 60-69). La que sobrevive es la MAS conservadora: lo que se llamaba
# "Envolvente" pasa a ser el caso "Diseno", en la columna E.
_FASE2_RETIRA_PRESION = (
    "Fase 2: la fila 27 ('Presion de diseno tipica') se retira. 212-3.2 define "
    "una unica P de diseno y 206-3.3 la nombra 'maximum allowable design "
    "pressure': el caso intermedio no lo publica el codigo y competia con la "
    "maxima admisible por gobernar el t_req. La presion de diseno vive ahora en "
    "la fila 28.")
_FASE2_RETIRA_COL_F = (
    "Fase 2: la columna F ('Envolvente') se retira de la tabla de cargas. Su "
    "caso -la presion maxima admisible- no desaparece: es el que ahora ocupa la "
    "columna E, rotulada 'Diseno'. Lo que se elimino es el 'Diseno tipico' "
    "intermedio que ocupaba E, no la envolvente.")
# Las direcciones van ya en el layout de la Fase 3 (MAPA_FILAS_212): la fila
# 27 no se movio -esta por encima del bloque de material- y la tabla de cargas
# 60-69 bajo a 88-97.
DIVERGENCIAS_DECLARADAS.update({
    ("Parche_PCC2_Art212", c): _FASE2_RETIRA_PRESION
    for c in ("A28", "B28", "C28", "D28", "G28")
})
DIVERGENCIAS_DECLARADAS.update({
    ("Parche_PCC2_Art212", f"F{MAPA_FILAS_212[r]}"): _FASE2_RETIRA_COL_F
    for r in range(60, 70)
})

# Celdas del bloque dimensional (Tareas 7-8) que el build reemplaza a proposito:
# NPS/cedula/OD/espesor pasan a salir de DB_B36 por cascada (reglas 12/14) y la
# fila 22 pasa de material descriptivo (lista fija retirada) al selector de norma.
# Su correccion la prueba TestCascadaDimensionalArt212, no el oracle Rev0. Las
# celdas de rotulo/unidad/simbolo que NO cambian (A18-A21, B/C del bloque, B22,
# C22) siguen bajo el oracle y no se listan aqui.
DIVERGENCIAS_REEMPLAZADAS = {
    ("Parche_PCC2_Art212", "D19"): (
        "NPS: ya no es una lista fija con valor 12; sale del desplegable de DB_B36 "
        "segun la norma (D22) y se guarda como el NPS impreso del codigo."),
    ("Parche_PCC2_Art212", "G19"): "Idem D18: la referencia ahora apunta a DB_B36.",
    ("Parche_PCC2_Art212", "D20"): (
        "Cedula: la validacion pasa de lista fija a rango dependiente de la norma y "
        "el NPS (el valor por defecto '20' coincide con el oracle, pero el origen "
        "de la validacion cambia)."),
    ("Parche_PCC2_Art212", "G20"): "Idem D19: la referencia ahora apunta a DB_B36.",
    ("Parche_PCC2_Art212", "D21"): (
        "OD: lookup contra DB_B36 (clave NPS|cedula, edicion segun D22) en vez de la "
        "tabla corta de Datos_Ref, que se retira (regla 1: la fuente es resources/)."),
    ("Parche_PCC2_Art212", "G21"): "Idem D20: la referencia ahora apunta a DB_B36.",
    ("Parche_PCC2_Art212", "D22"): (
        "Espesor: lookup contra DB_B36 (clave NPS|cedula, edicion segun D22) en vez "
        "de Datos_Ref, que se retira."),
    ("Parche_PCC2_Art212", "G22"): "Idem D21: la referencia ahora apunta a DB_B36.",
    ("Parche_PCC2_Art212", "A23"): (
        "La fila 22 pasa de 'Material de tuberia' (retirado) al rotulo 'Norma "
        "dimensional': es el nivel 0 de la cascada, se elige B36.10M o B36.19M."),
    ("Parche_PCC2_Art212", "D23"): (
        "Valor del selector de norma dimensional (B36.10M por defecto); gobierna las "
        "listas de NPS/cedula y el lookup de OD/espesor. Antes era material descriptivo."),
    ("Parche_PCC2_Art212", "G23"): "Idem A22: referencia de la norma dimensional.",
    ("Parche_PCC2_Art212", "F119"): (
        "DICTAMEN GLOBAL: la Fase 1 antepone la compuerta de elegibilidad del "
        "Paso 1 (D140) y la Fase 4 anade los dos topes de filete (F161, F162) al "
        "AND de verificaciones. La forma nueva la fija test_paso4_filete; el "
        "oracle Rev0 no la cubre."),
    ("Parche_PCC2_Art212", "D93"): (
        "w_min (Operacion): la Fase 4 lo recablea de F_m (D63) a la fuerza "
        "gobernante F_max del Paso 2 (D150), ec. 4 del 212-3.4. Para cilindro sin "
        "cargas externas el valor no cambia; lo fija test_paso4_filete."),
    ("Parche_PCC2_Art212", "E93"): (
        "w_min (Diseno tipico): recableado a F_max (E150) en la Fase 4. Ver D64."),
    ("Parche_PCC2_Art212", "D84"): (
        "Excentricidad e: la Fase 5 le suma la separacion g del faying edge "
        "(D167) cuando g>=1.5 mm (212-4c). Con g=0 el valor no cambia. La fija "
        "test_paso5_excentricidad."),
    ("Parche_PCC2_Art212", "D86"): (
        "C_sw: la Fase 5 lo pasa a la forma literal de cilindro (sin kf) con NA "
        "en esfera (D11=3), ec.(5) 212-3.4c. Para cilindro no cambia de valor."),
    ("Parche_PCC2_Art212", "D95"): (
        "S_w membrana (Op): forma literal de la ec.(5), P*Dm/(2T), NA en esfera "
        "(Fase 5). Cilindro sin cambio de valor. Ver test_paso5_excentricidad."),
    ("Parche_PCC2_Art212", "E95"): "S_w membrana (Diseno): idem D66 (Fase 5).",
    ("Parche_PCC2_Art212", "D96"): (
        "S_w flexion (Op): forma literal de la ec.(5), 3*P*Dm*e/T^2, NA en "
        "esfera (Fase 5). Cilindro sin cambio de valor."),
    ("Parche_PCC2_Art212", "E96"): "S_w flexion (Diseno): idem D67 (Fase 5).",
    ("Parche_PCC2_Art212", "D106"): (
        "Deformacion por conformado: la Fase 6 le anade la rama simple/doble "
        "(coef 50/75 segun geometria, ec.7/ec.6) y el factor (1-Rf/Ro) con Ro "
        "(D172). Cilindro con plancha plana da 50*T/Rf como antes. La fija "
        "test_paso6_conformado."),
}

# --- Fase 2: lo que el colapso de presiones REEMPLAZA (no retira) -----------
# Su correccion la prueba TestModeloDePresionDosCasos, no el oracle Rev0.
DIVERGENCIAS_REEMPLAZADAS.update({
    ("Parche_PCC2_Art212", "A29"): (
        "Rotulo: 'Presion envolvente (cota superior)' pasa a 'Presion de diseno "
        "(maxima admisible / rating)'. Es el mismo numero con el nombre que le da "
        "el codigo (206-3.3), y ya no compite con un 'diseno tipico' retirado."),
    ("Parche_PCC2_Art212", "B29"): (
        "Simbolo: P_env pasa a P_dis. Es LA presion de diseno del motor, no una "
        "cota superior aparte de ella."),
    ("Parche_PCC2_Art212", "G29"): (
        "Referencia: 'Rating / envolvente' pasa a 'Rating / max. admisible', que "
        "es como la nombra 206-3.3."),
    ("Parche_PCC2_Art212", "E89"): (
        "Encabezado de columna: 'Diseno tipico' pasa a 'Diseno'. La columna E "
        "deja de llevar el caso intermedio y lleva la presion maxima admisible."),
    ("Parche_PCC2_Art212", "E90"): (
        "La columna de diseno lee ahora D28 (maxima admisible) en vez de D27 "
        "(tipico, retirado). Las filas 62-69 de esa columna no cambian de forma: "
        "encadenan desde E61, asi que siguen ancladas al oracle."),
    ("Parche_PCC2_Art212", "D113"): (
        "El MAX del filete requerido barre D64:E64 (dos casos) en vez de D64:F64. "
        "Mismo numero -F64 esta vacia- pero la hoja no puede declarar un rango "
        "que ya no existe."),
    ("Parche_PCC2_Art212", "G113"): (
        "Criterio: 'w >= w_min (envolvente)' pasa a '(diseno)', el nombre nuevo "
        "de ese mismo caso."),
    ("Parche_PCC2_Art212", "A116"): (
        "Rotulo: 'Espesor de pared (envolvente)' pasa a '(diseno)'."),
    ("Parche_PCC2_Art212", "D116"): (
        "El t_req gobernante se lee de E65 (caso Diseno) en vez de F65 "
        "(Envolvente, columna retirada). Es el mismo caso fisico: la presion "
        "maxima admisible que exige 206-3.3."),
    ("Parche_PCC2_Art212", "E118"): (
        "La presion de diseno contra la que se compara P_max del parche se lee "
        "de D28 en vez de D27 (retirada)."),
    # B128 (la prueba de hermeticidad) estuvo aqui en la Fase 2 —su presion de
    # prueba pasaba a leerse de D28—, y la Fase 9 la RETIRA de la hoja junto con
    # el resto de la seccion de especificaciones tecnicas. Una celda no puede
    # estar en las dos tablas (TestParidadHojaParche lo exige), asi que su
    # entrada se fue a DIVERGENCIAS_DECLARADAS, mas abajo.
})

# --- Fase 3: la Seccion de Material sube a la posicion 2 --------------------
# El movimiento de filas en si NO produce divergencias contra el *oracle*: el
# oracle se traslada con EL MISMO mapa (remapear_oracle_212.py), asi que las
# formulas siguen casando letra a letra en su direccion nueva. Lo que si
# diverge es el TEXTO de las bandas, porque la seccion de material pasa a ser
# la 2 y todo lo que venia detras se corre un numero.
_FASE3_RENUMERA = (
    "Fase 3: la Seccion de Material sube justo detras de los datos de entrada "
    "y pasa a ser la 2, asi que esta banda se renumera. Lo fija "
    "TestNumeracionDeSecciones.")
DIVERGENCIAS_REEMPLAZADAS.update({
    ("Parche_PCC2_Art212", "A38"): (
        "Banda de la Seccion de Material. Deja de ser la '7' del final y pasa a "
        "ser la '2'; y se escribe literal (banda_literal) en vez de con "
        "rotulo(), que la dejaba en '[ MAYUSCULAS // CON BARRAS ]' — el unico "
        "rotulo de esta hoja con ese formato, justo al lado de sus vecinas."),
    ("Parche_PCC2_Art212", "A66"): _FASE3_RENUMERA + (
        " Ademas pierde el parentetico '(constantes — editables)': con la "
        "leyenda de color de la Fase 1 el 'editables' lo dice el relleno."),
    ("Parche_PCC2_Art212", "A79"): _FASE3_RENUMERA,
    ("Parche_PCC2_Art212", "A88"): _FASE3_RENUMERA,
    ("Parche_PCC2_Art212", "A100"): _FASE3_RENUMERA,
    ("Parche_PCC2_Art212", "A111"): _FASE3_RENUMERA,
    ("Parche_PCC2_Art212", "A121"): _FASE3_RENUMERA + (
        " Fase 9: la seccion sale de esta hoja a Espec_PCC2_Art212 y con ella "
        "el numeral -ya no es una seccion de este motor-. La fila se conserva "
        "como LETRERO que dice adonde se fue: quien busque la seccion 8 aqui "
        "tiene que encontrar la pista, no un hueco."),
    ("Parche_PCC2_Art212", "A17"): (
        "La banda de datos de entrada pierde su parentetico '(campo con linea "
        "inferior = editable)': describia el sistema visual ANTERIOR y con la "
        "leyenda de color de la Fase 1 seria falso. Mismo criterio que "
        "TEXTOS_HEREDADOS con los azules del maestro Rev0."),
    ("Parche_PCC2_Art212", "G68"): (
        "La referencia remite ahora a la 'seccion 2' (antes 7), que es donde "
        "vive la resolucion de material tras la Fase 3."),
    ("Parche_PCC2_Art212", "G69"): "Idem G67: la seccion de material es la 2.",
    ("Parche_PCC2_Art212", "G73"): "Idem G67: la seccion de material es la 2.",
})

# Celdas que existen POR DISEÑO y que el *oracle* Rev0 nunca tuvo. Es una
# tercera clase, distinta de las dos de arriba y con motivo propio: una celda
# RETIRADA estaba en el oracle y se vacia; una REEMPLAZADA estaba en el oracle y
# cambia de contenido; estas no estaban. Meterlas en REEMPLAZADAS "porque
# funciona" habria dejado el diccionario diciendo algo falso —que el oracle las
# declara— y con el tiempo nadie sabria cual es cual.
# --- Fase 7: el bloque de material gana el eje de EDICION -------------------
# Todas las celdas de resolucion pasan a un CHOOSE de seis ramas (las tres bases
# metricas y sus tres gemelas U.S. Customary) y la fila se resuelve en tres
# saltos via clave bilingue. El valor en modo SI no cambia —lo prueba el
# recalculo del caso semilla en verificar.py §7 y §6e—, pero la formula si.
_FASE7_EDICION = (
    "Fase 7: la celda pasa a leer de la EDICION que marque el selector de "
    "unidades (D15). El CHOOSE crece de tres ramas a seis y la fila se resuelve "
    "por clave bilingue, porque el material_id NO coincide entre ediciones. En "
    "modo SI el valor es el mismo; lo fija TestConmutadorDeUnidades y lo "
    "recalcula verificar.py.")
_FASE7_UNIDAD = (
    "Fase 7: el rotulo de unidad lo decide la edicion que se lee —°C/MPa en la "
    "metrica, °F/ksi en la U.S. Customary—. No se convierte nada (regla 9): "
    "cambia de que tabla se lee, y el rotulo lo dice.")
DIVERGENCIAS_REEMPLAZADAS.update(
    {("Parche_PCC2_Art212", f"{col}{fila}"): _FASE7_EDICION
     for col in "DE" for fila in range(49, 59)})
DIVERGENCIAS_REEMPLAZADAS.update(
    {("Parche_PCC2_Art212", f"C{fila}"): _FASE7_UNIDAD
     for fila in (41, 53, 54, 55, 56, 57, 59)})

# Los rotulos de unidad de TODA la hoja pasan a ser formula del selector (el
# pase aplicar_unidades_motor). El *oracle* los trae como literal metrico.
DIVERGENCIAS_REEMPLAZADAS.update(
    {("Parche_PCC2_Art212", f"C{fila}"): _FASE7_UNIDAD
     for fila in UNIDADES_212})

# Las constantes y umbrales que dependen del SISTEMA, no del caso.
DIVERGENCIAS_REEMPLAZADAS.update({
    ("Parche_PCC2_Art212", "D74"): (
        "Fase 7: la densidad del acero deja de teclearse y la fija el selector "
        "(7850 kg/m³ / 0,2836 lb/in³). Dejarla editable obligaria a acordarse de "
        "cambiarla al conmutar, y olvidarlo daria un peso plausible y falso."),
    ("Parche_PCC2_Art212", "D76"): (
        "Fase 7: el factor de conversion de la presion de entrada a la unidad de "
        "calculo tambien lo fija el selector: kg/cm²->MPa en metrico, psi->ksi en "
        "U.S. Customary."),
    ("Parche_PCC2_Art212", "D109"): (
        "Fase 7: el divisor del peso (mm³·kg/m³ -> kg) no existe en U.S. "
        "Customary: in³·lb/in³ ya da libras. Conmuta entre 1e9 y 1."),
    ("Parche_PCC2_Art212", "D108"): (
        "Fase 7: los 3 mm de luz de raiz son una LONGITUD. Sin conmutar restaban "
        "3 pulgadas en modo US — se vio en el recalculo: el peso salia un 48 % "
        "corto. No son del codigo (es holgura de fabricacion), asi que aqui si se "
        "convierte, y se declara."),
})

CELDAS_NUEVAS_FUERA_DEL_ORACLE = {
    ("Parche_PCC2_Art212", "A15"): (
        "Fase 7: rotulo del selector de sistema de unidades. La banda de "
        "aplicacion y codigo gana una fila; el oracle Rev0 solo llegaba a la 14."),
    ("Parche_PCC2_Art212", "B15"): "Idem A15: simbolo de la fila del selector.",
    ("Parche_PCC2_Art212", "C15"): "Idem A15: unidad de la fila del selector.",
    ("Parche_PCC2_Art212", "D15"): (
        "Fase 7: selector SI / US, con su validacion de lista. Decide de que "
        "EDICION del codigo se lee el esfuerzo admisible; nunca convierte un "
        "valor (regla 9). Lo fija TestConmutadorDeUnidades."),
    ("Parche_PCC2_Art212", "G15"): "Idem A15: referencia de la fila del selector.",
}

# --- Fase 4: el rastro de la resolucion se pliega y se redacta mas llano ----
_FASE4_ETIQUETA = (
    "Fase 4: rotulo reescrito sin jerga de implementacion. La fila se pliega "
    "por defecto (es rastro de auditoria, no interfaz) y cuando el ingeniero la "
    "despliega tiene que poder leerla sin conocer el builder. El VALOR de la "
    "celda no cambia; lo fija TestDiagnosticoPlegable.")
DIVERGENCIAS_REEMPLAZADAS.update({
    ("Parche_PCC2_Art212", "A50"): _FASE4_ETIQUETA + " 'Indice de base' -> 'Base ASME aplicada'.",
    ("Parche_PCC2_Art212", "A51"): _FASE4_ETIQUETA + " 'Fila localizada' -> '... en la base'.",
    ("Parche_PCC2_Art212", "A52"): _FASE4_ETIQUETA + " 'n_pts / p1' -> 'Puntos tabulados de la fila'.",
    ("Parche_PCC2_Art212", "A57"): _FASE4_ETIQUETA + " Sin el '/ limite VIII-1' de mas.",
    # --- Fase 5: fuera el parentesis EXPLICATIVO de las bandas -------------
    # La cita al codigo se queda -es informacion normativa- pero sale del
    # parentesis; lo que se va es lo puramente aclaratorio. Una banda no tiene
    # que explicar como funciona la hoja: para eso esta la pestana de
    # instrucciones (Fase 10) y el comentario de cada celda.
    ("Parche_PCC2_Art212", "A8"): (
        "Fase 5: fuera el parentetico '(selector que conmuta S, t_req y la "
        "fuerza de membrana)'. Lo explica el comentario de la celda del "
        "selector, que es donde se mira."),
    ("Parche_PCC2_Art212", "A111"): (
        "Fase 3 (numeral 5 -> 7) y Fase 5: fuera el parentetico '(criterios de "
        "aceptacion)'. Una tabla titulada VERIFICACIONES no necesita que le "
        "digan que verifica."),
    ("Parche_PCC2_Art212", "A121"): (
        "Fase 5: fuera el parentetico '(generadas automaticamente a partir de "
        "las entradas)', ademas del numeral nuevo de la Fase 3."),
    ("Parche_PCC2_Art212", "A88"): (
        "Fase 3 (numeral) y Fase 5: la cita 'ASME PCC-2 Art. 212' se conserva "
        "-es normativa- pero sale del parentesis, que en el resto de la hoja "
        "significa 'aclaracion prescindible'."),
    ("Parche_PCC2_Art212", "G42"): (
        "Fase 4: la columna de notas decia 'Lista desplegable en cascada' en los "
        "CINCO niveles. Repetir la misma frase cinco veces no informa: la vuelve "
        "ruido y empuja fuera de la vista lo que si es propio de cada nivel. Se "
        "dice una vez aqui, en el nivel 0."),
})
# Los otros cuatro niveles se quedan SIN nota (retirada, no reemplazada).
DIVERGENCIAS_DECLARADAS.update({
    ("Parche_PCC2_Art212", f"G{r}"): (
        "Fase 4: nota de cascada repetida. La explicacion vive una sola vez en "
        "G41 (nivel 0) y esta fila la hereda; ver TestDiagnosticoPlegable.")
    for r in range(43, 48)
})

# --- Fase 9: las especificaciones tecnicas salen a su propia pestana ---------
# Las siete filas de parrafo (93-99 del oracle; 122-128 tras el mapa de la Fase
# 3) se retiran de la hoja del motor y renacen, con la cita del parrafo de PCC-2
# que sostiene cada una, en Espec_PCC2_Art212. Es el bloque mas denso de la hoja
# y el unico que no se consulta mientras se calcula.
#
# Se retiran el rotulo (columna A) y el parrafo (columna B, que estaba fusionado
# B:G). La banda A121 NO se retira: se reescribe como letrero y por eso sigue en
# DIVERGENCIAS_REEMPLAZADAS, arriba.
_FASE9_A_PESTANA = (
    "Fase 9: la seccion de especificaciones tecnicas sale de la hoja del motor "
    "a Espec_PCC2_Art212, donde cada especificacion cita el parrafo de PCC-2 que "
    "la sostiene (212-3.4, 212-4, 212-5, 212-6 y Art. 210) — la cita no cabia "
    "en una fila de esta hoja. La fila queda vacia; el letrero de A121 dice "
    "adonde se fue.")
DIVERGENCIAS_DECLARADAS.update({
    ("Parche_PCC2_Art212", f"{col}{MAPA_FILAS_212[r]}"): _FASE9_A_PESTANA
    for r in range(93, 100) for col in ("A", "B")
})
DIVERGENCIAS_REEMPLAZADAS.update({
    ("Parche_PCC2_Art212", "A3"): (
        "Fase 9: el boton de retorno dice '<<< VOLVER A ART. 212' en vez de "
        "'<<< VOLVER A PCC-2'. No es un cambio de texto: el articulo paso de ser "
        "una tarjeta que abre el motor a ser un NIVEL del arbol con los "
        "documentos de ese articulo (motor, especificaciones e instrucciones), asi "
        "que el padre de esta hoja cambio. El texto lo deriva _texto_volver() de "
        "PADRE; no se escribe a mano en ningun sitio."),
})

# Los rgb se comparan por sus SEIS digitos de color, sin el alfa: openpyxl
# devuelve "FF1A1A1A" en lo que leyo del maestro y "00050505" en lo que acaba de
# escribir este script a partir de una cadena de seis digitos. Comparar los ocho
# hacia pasar por «heredado» a un color de la paleta nueva.
def _rgb6(v):
    return v[-6:].upper() if isinstance(v, str) and len(v) >= 6 else None


# Relleno heredado -> relleno del sistema.
MAPA_RELLENO = {
    "16304F": TINTA,        # cabecera mas oscura del maestro
    "1F4E79": TINTA,        # banda de seccion
    "305496": TINTA,        # banda de subseccion / cabecera de tabla
    "F3F6FA": PAPEL_2,      # panel
    "DDE6F0": PAPEL_2,      # banda de fila alterna
    "FFFFFF": PAPEL,        # blanco puro -> papel sin blanquear
    "FFF7DC": PAPEL,        # celda de entrada: pasa a campo (caja de tinta)
}
# La celda de entrada del maestro se reconoce por su amarillo y es la unica que
# ademas cambia de BORDE: en este sistema un campo es papel dentro de una caja
# de tinta, no un relleno de color (ver IN_FILL / CAJA_CAMPO).
RELLENO_DE_CAMPO = "FFF7DC"

# (nombre, tamano, negrita, cursiva, color) heredado -> Font del sistema. El
# tamano se conserva salvo en los titulos, que pasan a la macrotipografia.
MAPA_FUENTE = {
    ("Arial", 14.0, True, False, "FFFFFF"): Font(name=MACRO, size=14, color=PAPEL),
    ("Arial", 13.0, True, False, "FFFFFF"): Font(name=MACRO, size=13, color=PAPEL),
    ("Arial", 12.0, True, False, "1A1A1A"): Font(name=MACRO, size=12, color=TINTA),
    ("Arial", 11.0, True, False, "FFFFFF"): Font(name=MACRO, size=11, color=PAPEL),
    ("Arial", 10.0, True, False, "FFFFFF"): Font(name=MONO, size=10, bold=True,
                                                 color=PAPEL),
    ("Arial", 9.0, True, False, "FFFFFF"): Font(name=MONO, size=9, bold=True,
                                                color=PAPEL),
    ("Arial", 9.0, False, True, "FFFFFF"): Font(name=MONO, size=9, color=PAPEL),
    ("Arial", 10.0, True, False, "1F4E79"): Font(name=MONO, size=10, bold=True,
                                                 color=TINTA),
    ("Arial", 9.0, True, False, "1F4E79"): Font(name=MONO, size=9, bold=True,
                                                color=TINTA),
    ("Arial", 10.0, True, False, "1A1A1A"): Font(name=MONO, size=10, bold=True,
                                                 color=TINTA),
    ("Arial", 9.0, True, False, "1A1A1A"): Font(name=MONO, size=9, bold=True,
                                                color=TINTA),
    ("Arial", 8.0, True, False, "1A1A1A"): Font(name=MONO, size=8, bold=True,
                                                color=TINTA),
    ("Arial", 8.0, True, False, "FFFFFF"): Font(name=MONO, size=8, bold=True,
                                                color=PAPEL),
    ("Arial", 10.0, False, False, "1A1A1A"): Font(name=MONO, size=10, color=TINTA),
    ("Arial", 9.0, False, False, "1A1A1A"): Font(name=MONO, size=9, color=TINTA),
    ("Arial", 8.0, False, False, "1A1A1A"): Font(name=MONO, size=8, color=TINTA),
    # Metadato y nota del maestro: gris y turquesa en cursiva. Los dos pasan a
    # la trama de medio tono, que es lo que este sistema usa para el metadato.
    ("Arial", 8.0, False, True, "7A7A7A"): Font(name=MONO, size=8, color=GRIS),
    ("Arial", 9.0, False, True, "7A7A7A"): Font(name=MONO, size=9, color=GRIS),
    ("Arial", 9.0, False, True, "0D7C7C"): Font(name=MONO, size=9, color=GRIS),
    # Texto de entrada del maestro (azul): pasa a tinta, como IN_F.
    ("Arial", 10.0, True, False, "0000CC"): IN_F,
    ("Arial", 9.0, True, False, "0000CC"): Font(name=MONO, size=9, bold=True,
                                                color=TINTA),
}
# Celda con estilo propio pero fuente por omision (Calibri 11 del tema): son
# celdas de dato del maestro y toman la fuente de dato del sistema.
FUENTE_POR_OMISION = ("Calibri", 11.0, False, False, None)

# Rotulos del maestro que DESCRIBEN el sistema visual y por tanto caducan con
# el: retonar la celda sin reescribir su texto dejaria al libro explicando una
# convencion que ya no existe. Se comprueba el texto esperado y, si no aparece,
# se declara en ISSUES en vez de reescribir a ciegas una celda que cambio.
TEXTOS_HEREDADOS = {
}

# Los colores que SI son del sistema, para no reportarlos como heredados.
_RGB_SISTEMA = {PAPEL, PAPEL_2, TINTA, TINTA_2, ROJO, GRIS, GRIS_2, VERDE,
                AMBAR, AMBAR_TXT}


def _clave_fuente(f):
    if f is None:
        return None
    rgb = getattr(f.color, "rgb", None) if f.color is not None else None
    return (f.name, f.sz, bool(f.b), bool(f.i), _rgb6(rgb))


def retonar_heredadas(wb):
    """Pasa al sistema visual las hojas que vienen del maestro sembrado.

    Devuelve (celdas tocadas, estilos heredados que no estaban en el mapa). Lo
    segundo se declara en ISSUES: un estilo sin traducir es una celda que se
    quedaria con el aspecto de la Rev. 0, y hay que verlo, no adivinarlo.
    """
    tocadas, sin_mapa = 0, Counter()
    for (hoja, ref), (esperado, nuevo) in TEXTOS_HEREDADOS.items():
        if hoja not in wb.sheetnames:
            continue
        cel = wb[hoja][ref]
        if cel.value == esperado:
            cel.value = nuevo
            tocadas += 1
        elif cel.value != nuevo:
            ISSUES.append(
                f"Retonado: {hoja}!{ref} ya no dice lo que declaraba el maestro, "
                f"asi que no se reescribio. Compruebe si sigue describiendo la "
                f"convencion visual antigua: {str(cel.value)[:80]!r}")
    for nombre in HOJAS_HEREDADAS:
        if nombre not in wb.sheetnames:
            continue
        ws = wb[nombre]
        for row in ws.iter_rows():
            for c in row:
                cambio = False
                relleno = c.fill
                rgb = _rgb6(getattr(relleno.fgColor, "rgb", None)
                            if relleno is not None and relleno.patternType else None)
                if rgb in MAPA_RELLENO:
                    era_campo = rgb == RELLENO_DE_CAMPO
                    c.fill = PatternFill("solid", fgColor=MAPA_RELLENO[rgb])
                    if era_campo:
                        c.border = CAJA_CAMPO
                    cambio = True
                elif rgb is not None and rgb not in _RGB_SISTEMA:
                    sin_mapa[f"{nombre}: relleno {rgb}"] += 1
                clave = _clave_fuente(c.font)
                if clave in MAPA_FUENTE:
                    c.font = MAPA_FUENTE[clave]
                    cambio = True
                elif clave == FUENTE_POR_OMISION:
                    c.font = DATA_F
                    cambio = True
                elif clave is not None and clave[0] not in (MONO, MACRO):
                    sin_mapa[f"{nombre}: fuente {clave}"] += 1
                if cambio:
                    tocadas += 1
    if sin_mapa:
        ISSUES.append(
            "Retonado de las hojas del maestro: "
            f"{len(sin_mapa)} estilos heredados sin equivalencia declarada "
            f"({', '.join(sorted(sin_mapa)[:6])}). Esas celdas conservan el "
            "aspecto de la Rev. 0; anada su equivalencia a MAPA_RELLENO / "
            "MAPA_FUENTE.")
    return tocadas, sin_mapa


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


# Layout de las dos bases dimensionales. El indice de columna se declara una sola
# vez y lo consumen el builder, las pruebas y verificar.py: si alguien inserta una
# columna, se entera todo el mundo a la vez.
#
# nps_orden, clave y clave_ced no son datos impresos por el codigo: son
# columnas auxiliares que alimentan la cascada dimensional de los motores
# (Tareas 7-8), precalculadas aqui como TEXTO porque una base no lleva formulas.
#   nps_orden  numera 1,2,3... la PRIMERA fila de cada bloque de NPS, para listar
#              los NPS sin repetirlos (nivel 1 de la cascada).
#   clave      = nps_impreso & "|" & designador, para el lookup de OD/espesor con
#              un unico MATCH (nivel de resultado).
#   clave_ced  = nps_impreso & "|" & <indice 1,2,3... del designador NO vacio
#              dentro del bloque de NPS>, y VACIA en las filas cuyo designador no
#              aplica (celda '...' del codigo). Deja listar las cedulas de un NPS
#              sin ofrecer opciones en blanco (nivel 2 de la cascada).
COL_B36 = {"nps_impreso": 1, "nps_in": 2, "dn_mm": 3, "designador": 4,
           "cedula": 5, "identificacion": 6, "od_in": 7, "od_mm": 8,
           "t_in": 9, "t_mm": 10, "peso_lb_ft": 11, "peso_kg_m": 12,
           "nps_orden": 13, "clave": 14, "clave_ced": 15}

CABECERA_B36 = ["NPS impreso", "NPS (in)", "DN (mm)", "Designador",
                "Schedule No.", "Identification", "OD (in)", "OD (mm)",
                "Espesor (in)", "Espesor (mm)", "Peso (lb/ft)", "Peso (kg/m)",
                "nps_orden (auxiliar)", "clave (auxiliar)", "clave_ced (auxiliar)"]

# Primera columna oculta de la cascada dimensional en la hoja de un motor
# (norma en COL_NORMA_B36, NPS en +1, cedula en +2, indice escalar de norma en
# +3). Y(25) en adelante: libre de las columnas ocultas de la Seccion 7, que
# llegan hasta la U(21) mas W/X de los indices auxiliares.
COL_NORMA_B36 = 25


def build_db_b36(wb, norma):
    """Base dimensional de una de las dos normas de tuberia.

    Una fila por (NPS, cedula), ordenada por NPS y luego por espesor para que cada
    NPS sea un bloque contiguo (regla 5) — lo exige la lista dependiente de cedula
    de los motores, que resuelve con COUNTIF/INDEX/MATCH y no con matrices.

    Las dos unidades vienen en columnas separadas porque la norma las publica en la
    MISMA celda: el conmutador de los motores cambia de columna, nunca convierte
    (regla 10).
    """
    cfg = {"B36.10M": (b36_dimensiones.RUTA_B3610, "DB_B36_10",
                       "ASME B36.10M-2022 — Welded and Seamless Wrought Steel Pipe"),
           "B36.19M": (b36_dimensiones.RUTA_B3619, "DB_B36_19",
                       "ASME B36.19M-2022 — Stainless Steel Pipe")}[norma]
    ruta, nombre, titulo = cfg
    filas = b36_dimensiones.cargar(ruta)
    # Orden: NPS ascendente y, dentro de cada NPS, espesor ascendente. El espesor
    # ordena mejor que la cedula porque la cedula es texto y mezcla numeros con
    # STD/XS/XXS y con 5S/10S: ordenarla como texto pondria "10" antes que "5".
    filas.sort(key=lambda f: (f["nps_in"] if f["nps_in"] is not None else 1e9,
                              f["t_mm"] if f["t_mm"] is not None else 1e9))
    ws = new_sheet(wb, nombre, titulo,
                   f"Fuente: resources/ASME B36/{ruta.parent.name}/{ruta.name}")
    for j, h in enumerate(CABECERA_B36, start=1):
        c = ws.cell(R_HDR, j, h)
        c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
    vistos = set()
    orden = 0
    nps_actual = None            # bloque de NPS en curso (contiguo por regla 5)
    ced_idx = 0                  # indice del designador NO vacio dentro del bloque
    max_ced = 0
    for i, f in enumerate(filas):
        r = R_DATA + i
        for clave, j in COL_B36.items():
            if clave in ("nps_orden", "clave", "clave_ced"):
                continue
            ws.cell(r, j, f[clave]).font = DATA_F
        if f["nps_impreso"] not in vistos:
            vistos.add(f["nps_impreso"])
            orden += 1
            ws.cell(r, COL_B36["nps_orden"], orden).font = DATA_F
        if f["nps_impreso"] != nps_actual:
            nps_actual = f["nps_impreso"]
            ced_idx = 0
        # clave_ced solo en las filas con designador; las '...' del codigo quedan
        # sin clave_ced (None) para no entrar como cedula en blanco al desplegable.
        if f["designador"]:
            ced_idx += 1
            max_ced = max(max_ced, ced_idx)
            ws.cell(r, COL_B36["clave_ced"],
                    f"{f['nps_impreso']}|{ced_idx}").font = DATA_F
        ws.cell(r, COL_B36["clave"], f"{f['nps_impreso']}|{f['designador']}").font = DATA_F
    autosize(ws, {"A": 14, "B": 10, "C": 10, "D": 14, "E": 13, "F": 13,
                  "G": 11, "H": 11, "I": 13, "J": 13, "K": 13, "L": 13,
                  "M": 14, "N": 20, "O": 20})
    # n_nps y max_ced dimensionan las listas de la cascada del motor sin numeros
    # magicos: se generan tantas filas de formula como NPS y como cedulas haya de
    # verdad, asi ninguna lista sale truncada (Tarea 7).
    return {"sheet": nombre, "last_row": R_DATA + len(filas) - 1,
            "n_nps": orden, "max_ced": max_ced}


def _rango_b36(info, clave):
    """Rango de columna completa de una base B36, para las formulas de la
    cascada dimensional (Tareas 7-9). Funcion de modulo para no duplicar la
    expresion entre _materializar_cascada_b36 y las formulas de OD/espesor."""
    L = get_column_letter(COL_B36[clave])
    return f"{info['sheet']}!${L}${R_DATA}:${L}${info['last_row']}"


def _choose_b36(idx, b3610, b3619, clave, sel=None, clave_us=None):
    """Selecciona la columna `clave` de la edicion elegida con CHOOSE(indice,...).

    CHOOSE con un indice ESCALAR (1 o 2) entrega una REFERENCIA de rango, que
    INDEX/MATCH consumen sin formula matricial — igual que col3() de la Seccion 7.
    Un IF(cond, rangoA, rangoB) como argumento de MATCH exigiria entrada matricial
    (CSE), que la regla 1 de diseno del libro prohibe."""
    if sel is None or clave_us is None:
        return (f"CHOOSE({idx},{_rango_b36(b3610, clave)},"
                f"{_rango_b36(b3619, clave)})")
    # Cuatro ramas: norma (1/2) + 2 si el selector dice US. El indice sigue
    # siendo ESCALAR, asi que CHOOSE entrega una referencia y no hace falta
    # entrada matricial (misma razon que arriba).
    return (f"CHOOSE({idx}+IF({sel},0,2),"
            f"{_rango_b36(b3610, clave)},{_rango_b36(b3619, clave)},"
            f"{_rango_b36(b3610, clave_us)},{_rango_b36(b3619, clave_us)})")


def _materializar_cascada_b36(ws, b3610, b3619, fila_norma, fila_nps, fila_ced):
    """Cascada norma -> NPS -> cedula de un motor, en columnas ocultas.

    La regla 2 prohibe una formula como origen de una validacion, asi que cada
    lista se calcula en una columna oculta de la propia hoja y la validacion
    apunta a ese rango literal — el mismo mecanismo que construir_seccion7_material.

    La norma se ELIGE, no se deriva del material: una tuberia inoxidable puede
    pedirse a cedulas B36.10M, asi que derivarla de la familia daria el espesor
    equivocado, y ademas acoplaria dos secciones del motor (regla 13).

    Devuelve {"idx": ref del indice escalar de norma, "ultima_col": ...} para que
    el llamador arme el lookup de OD/espesor con el mismo CHOOSE.
    """
    hl = get_column_letter
    n_nps = max(b3610["n_nps"], b3619["n_nps"])
    n_ced = max(b3610["max_ced"], b3619["max_ced"])
    col = COL_NORMA_B36
    L_norma, L_nps, L_ced, L_idx = (hl(col), hl(col + 1), hl(col + 2), hl(col + 3))

    # Nivel 0: las dos normas, literales en la hoja (no en el formula1 de la
    # validacion: asi el origen sigue siendo un rango). Anadir una tercera norma
    # manana es anadir una fila aqui, no tocar el motor.
    ws.cell(R_HDR, col, "lista norma B36").font = SRC_F
    for k, n in enumerate(("B36.10M", "B36.19M")):
        ws.cell(R_DATA + k, col, n).font = SRC_F
    ws.column_dimensions[L_norma].hidden = True
    dv_list(ws, f"D{fila_norma}", f"=${L_norma}${R_DATA}:${L_norma}${R_DATA + 1}")

    # Indice escalar 1/2 de la norma elegida, para el CHOOSE de todos los rangos.
    idx = f"${L_idx}${R_SRC}"
    ws.cell(R_SRC, col + 3, f'=IF($D${fila_norma}="B36.10M",1,2)').font = SRC_F
    ws.column_dimensions[L_idx].hidden = True

    # Nivel 1: NPS distintos de la norma elegida. El k-esimo es el de nps_orden=k
    # (numerado 1,2,3... en la primera fila de cada bloque contiguo).
    ws.cell(R_HDR, col + 1, "lista NPS B36").font = SRC_F
    for k in range(1, n_nps + 1):
        ws.cell(R_DATA + k - 1, col + 1,
                f'=IFERROR(INDEX({_choose_b36(idx, b3610, b3619, "nps_impreso")},'
                f'MATCH({k},{_choose_b36(idx, b3610, b3619, "nps_orden")},0)),"")'
                ).font = SRC_F
    ws.column_dimensions[L_nps].hidden = True
    dv_list(ws, f"D{fila_nps}", f"=${L_nps}${R_DATA}:${L_nps}${R_DATA + n_nps - 1}")

    # Nivel 2: cedulas (designador) del NPS elegido, en la norma elegida. El
    # k-esimo designador NO vacio del bloque se localiza por clave_ced = NPS&"|"&k,
    # asi que las celdas '...' del codigo no entran como opcion en blanco.
    ws.cell(R_HDR, col + 2, "lista cedula B36").font = SRC_F
    for k in range(1, n_ced + 1):
        ws.cell(R_DATA + k - 1, col + 2,
                f'=IFERROR(INDEX({_choose_b36(idx, b3610, b3619, "designador")},'
                f'MATCH($D${fila_nps}&"|"&{k},'
                f'{_choose_b36(idx, b3610, b3619, "clave_ced")},0)),"")').font = SRC_F
    ws.column_dimensions[L_ced].hidden = True
    dv_list(ws, f"D{fila_ced}", f"=${L_ced}${R_DATA}:${L_ced}${R_DATA + n_ced - 1}")
    return {"idx": idx, "ultima_col": col + 3}


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
    b1 = build_b1(res, wb, "SI")
    b1c = build_b1(res, wb, "US")
    verificar_paridad_b1(b1, b1c)
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
    sec2 = build_secii(res, wb)

    # listas simples de los buscadores por grupo
    wse, wste, wsprd = wb["DB_E"], wb["DB_TE_G"], wb["DB_PRD"]
    e_pairs = uniques(wse, e_si["last_row"], 3, 4)
    te_pairs = uniques(wste, te_g_si["last_row"], 3, 4)
    prd_pairs = uniques(wsprd, prd["last_row"], 3, 4)
    simples = [
        ("E_TABLA", sorted({k for k, _ in e_pairs})),
        ("E_K", [k for k, _ in e_pairs]), ("E_V", [v for _, v in e_pairs]),
        ("TE_TABLA", sorted({k for k, _ in te_pairs})),
        ("TE_K", [k for k, _ in te_pairs]), ("TE_V", [v for _, v in te_pairs]),
        ("PRD_TABLA", sorted({k for k, _ in prd_pairs if k})),
        ("PRD_K", [k for k, _ in prd_pairs]), ("PRD_V", [v for _, v in prd_pairs]),
    ]
    # El Apendice C entra en build_listas como una base mas: su contrato de las
    # 8 primeras columnas es identico al de STRESS_COLS, asi que la funcion no
    # necesita ni una linea de cambio.
    # La base de la Tabla B-1 entra aqui como una mas: su contrato de las 8
    # primeras columnas es el mismo (assert en APXB1_COLS), asi que build_listas
    # no necesita ni una linea de cambio. Solo entra la edicion SI: los rotulos
    # de la cascada son los nombres impresos en la metrica y el motor resuelve la
    # fila US por clave_bi, igual que hace Buscar_Prop_B31_3.
    rangos = build_listas(wb, [("B313", b313), ("IID1A", iid), ("IIDB", iidb),
                               ("SU", su), ("SY", sy), ("APXC", apxc),
                               ("APXB1", b1), ("A2EC", a2), ("A3EJ", a3)], simples)

    curvas = wb.create_sheet("_Curvas")
    curvas["A1"] = ("Datos auxiliares de las graficas. Hoja oculta: no editar. "
                    "Los valores proceden de las hojas DB por INDEX/MATCH.")
    curvas.sheet_state = "hidden"

    busc = [
        ("Buscar_B31_3", "BUSCADOR — ASME B31.3-2024, Apendice A (Tablas A-1 y A-4)",
         "B313", b313, b313c, "MPa", "ksi", "S admisible",
         # Este buscador reune DOS tablas con funciones distintas y hay que
         # decirlo antes de que el usuario elija: el admisible de un perno no es
         # el de un componente a presion. La ficha lo repite fila a fila.
         "Reune las DOS tablas de admisibles del Apendice A: la A-1 (A-1C en US) "
         "publica los ESFUERZOS BASICOS ADMISIBLES EN TRACCION de los metales "
         "—tuberia, placa, forja y fundicion—, y la A-4 (A-4C en US) publica los "
         "ESFUERZOS DE DISENO DE LA PERNERIA —pernos, esparragos y tuercas de "
         "union bridada—. La ficha del material dice de cual de las dos sale la "
         "fila resuelta y que publica esa tabla."),
        ("Buscar_BPVC_IID", "BUSCADOR — ASME BPVC II-D 2025, Tabla 1A (ferrosos)",
         "IID1A", iid, iidc, "MPa", "ksi", "S admisible",
         "Tabla 1A: esfuerzos admisibles de materiales FERROSOS."),
        ("Buscar_BPVC_IID_B", "BUSCADOR — ASME BPVC II-D 2025, Tablas 1B (no ferrosos) y 3",
         "IIDB", iidb, iidbc, "MPa", "ksi", "S admisible",
         "Reune dos tablas: la 1B publica los admisibles de los materiales NO "
         "FERROSOS y la 3 los de la PERNERIA. La ficha dice de cual sale la fila."),
        ("Buscar_Su", "BUSCADOR — Resistencia a la traccion Su (ASME BPVC II-D, Tabla U)",
         "SU", su, suc, "MPa", "ksi", "Su",
         "Tabla U: resistencia a la traccion minima especificada frente a la "
         "temperatura. No es un esfuerzo admisible."),
        ("Buscar_Sy", "BUSCADOR — Limite de fluencia Sy (ASME BPVC II-D, Tabla Y-1)",
         "SY", sy, syc, "MPa", "ksi", "Sy",
         "Tabla Y-1: limite de fluencia minimo especificado frente a la "
         "temperatura. No es un esfuerzo admisible."),
    ]
    for i, (nme, ttl, pref, m, u, us_, uu, vl, nta) in enumerate(busc):
        ctx = build_buscador(wb, curvas, nme, ttl, pref, rangos[pref], m, u,
                             us_, uu, vl, nota=nta)
        finish_buscador(ctx, wb, curvas, i)

    e_tab = sorted({k for k, _ in e_pairs})
    te_tab = sorted({k for k, _ in te_pairs})
    prd_tab = sorted({k for k, _ in prd_pairs if k})
    from collections import Counter as _Cnt
    mx_e = max(_Cnt(k for k, _ in e_pairs).values())
    mx_te = max(_Cnt(k for k, _ in te_pairs).values())
    mx_prd = max(_Cnt(k for k, _ in prd_pairs).values())
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
    build_buscador_b1(wb, curvas, rangos["APXB1"], b1, b1c)

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
             Font(name=MONO, size=10, bold=True, color=TINTA)),
            (f"para. 302.3.4(b), transcrito: {par34}",
             Font(name=MONO, size=9, color=TINTA)),
            (f"Table 302.3.4-1, Nota (1): {nota34}" if nota34 else
             "No se pudo leer la Nota (1) de la Tabla 302.3.4-1 en resources/.",
             Font(name=MONO, size=9, bold=True, color=AMBAR_TXT)),
            ("El factor Ej lo decide el TIPO DE JUNTA/COSTURA/EXAMEN de la fila "
             "abajo, no la especificacion de material seleccionada arriba: "
             "identifique cual de las diez filas describe su junta y lea su Ej. "
             "El codigo no imprime una correspondencia fila a fila entre la "
             "Tabla A-3 y esta tabla, asi que el motor no la infiere.",
             Font(name=MONO, size=9, bold=True, color=ROJO)),
        ] + bloques_tabla_ej(filas34)))

    b3610 = build_db_b36(wb, "B36.10M")
    b3619 = build_db_b36(wb, "B36.19M")

    # Coeficientes del Paso 8 (energia neumatica) leidos de resources/ (App.
    # 501, reparado en la Fase 0.2). Regla n.1: no salen de memoria. Aborta si
    # el apendice no esta reparado.
    energia_501 = leer_energia_501(a.resources)
    # Las tres bases US viajan a los motores igual que a los buscadores: el
    # conmutador de la Fase 7 lee de ellas, nunca convierte (regla 9/10).
    umbrales = leer_umbrales_pcc2(a.resources)
    build_parche_art212(wb, b313, iid, iidb, fac, rangos, b3610, b3619,
                        energia_501=energia_501, umbrales=umbrales,
                        b313c=b313c, iid1ac=iidc, iidbc=iidbc)
    build_collar_art206(wb, b313, iid, iidb, fac, rangos, b3610, b3619,
                        b313c=b313c, iid1ac=iidc, iidbc=iidbc,
                        umbrales=umbrales)
    # Fase 9: las especificaciones tecnicas de cada articulo, en pestana propia.
    # Van DESPUES de sus motores: leen celdas suyas y la hoja tiene que existir.
    build_especificaciones_art212(wb)
    build_especificaciones_art206(wb)
    # Fase 10: la guia de uso se DERIVA de la hoja del motor, asi que va
    # necesariamente despues de que el motor este entero y remapeado.
    build_instrucciones_art212(wb)
    build_instrucciones_art206(wb)
    retirar_datos_ref(wb)

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
        "Seccion II A/B/C -> DB_SecII_*":
            f"{sec2['tablas']} tablas -> {sec2['filas']} filas "
            f"({' · '.join(f'{k}={v}' for k, v in sorted(sec2['conf'].items()))})",
    }
    build_meta(wb, counts)
    rewrite_instrucciones(
        wb, "Rev. 4c — Entran al libro las tablas de la Seccion II, partes A, B y C: "
            "nueve hojas de DATOS (CAT_SecII, IDX_SecII_Tablas, los cuatro DB_SecII_* "
            "del volcado integro, DB_SecII_Notas y las normalizadas de quimica y "
            "traccion), navegables desde el Dashboard y sin una sola formula. El "
            "45,8 % de sus filas llega marcada AMBIGUA: no perdieron texto —va entero "
            "en la celda C01— pero NO quedaron repartidas en columnas, porque el JSON "
            "de la Seccion II no conserva la posicion de cada palabra dentro de la "
            "linea. IDX_SecII_Tablas lo dice tabla a tabla y el Dashboard publica el "
            "total. Para valores leidos de esas tablas, el PDF del codigo manda. "
            "MAP_Grupo queda cerrado: 0 casos abiertos, 106 declarados como limite de "
            "la fuente. Se mantiene el Dashboard unico de navegacion (libro con "
            "macros, .xlsm), la consulta por cascada y el libro sin funciones de "
            "matriz dinamica.")

    # Los conteos del Dashboard salen de las mismas variables que alimentan
    # `counts`, nunca escritos a mano. El indicador cuenta lo que el codigo NO
    # resuelve: la regla textual que debe aplicar el ingeniero y los materiales
    # para los que II-D no publica el dato. Las filas AUTO no entran: citan la
    # nota que las sostiene y son auditables 1:1.
    n_sin_resolver = map_stats.get(E_TEXTUAL, 0) + map_stats.get(E_SIN, 0)
    build_arbol(wb, [
        ("MATERIALES B31.3 · A-1 y A-4", b313["last_row"] - R_DATA + 1, "registros", False),
        ("MATERIALES II-D · TABLA 1A", iid["last_row"] - R_DATA + 1, "registros", False),
        ("MAP_Grupo SIN GRUPO NORMATIVO", n_sin_resolver,
         "filas · no usar E ni dilatacion", True),
        # Las filas AMBIGUAS no perdieron texto -van enteras en una celda- pero
        # NO quedaron repartidas en columnas. Se publica en ambar porque es la
        # limitacion que hay que tener delante al leer esas hojas.
        ("SEC. II · FILAS SIN TABULAR", sec2["conf"][secii.AMBIGUA],
         f"de {sec2['filas']} · texto integro, sin repartir", True),
    ], datetime.date.today().isoformat())

    link_volver(wb)

    # El Dashboard tiene que quedar en el indice 0; detras van las diez hojas
    # de navegacion, en el orden de preorden del arbol.
    order = [DASH] + HOJAS_NAV[1:] + [
             "Instrucciones", "Parche_PCC2_Art212", "Collar_PCC2_Art206",
             "Buscar_B31_3", "Buscar_BPVC_IID",
             "Buscar_BPVC_IID_B", "Buscar_Su", "Buscar_Sy", "Buscar_Prop_IID",
             "Buscar_Prop_B31_3", "Buscar_B31_B1", "Buscar_Ec_A2", "Buscar_Ej_A3",
             "DB_B31_3", "DB_B31_3C", "DB_BPVC_IID",
             "DB_BPVC_IIDC", "DB_BPVC_IID_B", "DB_BPVC_IID_BC", "DB_Su", "DB_SuC",
             "DB_Sy", "DB_SyC", "DB_E", "DB_EC", "DB_TE", "DB_TEC", "DB_TE_G", "DB_TE_GC",
             "DB_PRD", "DB_PRDC",
             "DB_B31_C", "DB_B31_CC", "DB_B31_B1", "DB_B31_B1C",
             "MAP_Factores", "DB_A2_Ec", "DB_A3_Ej",
             "DB_Ec_Incremento", "MAP_Grupo", "MAP_GrupoC",
             "Notas_Codigo"] + NAV_SECII + ["DB_Listas",
             "_meta", "_Curvas"]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + \
                 [s for s in wb._sheets if s.title not in order]

    # Sistema visual, al final y sobre el libro entero. El orden importa: primero
    # se traduce lo que trae el maestro y solo despues se papela, para que el
    # sustrato no tape un relleno heredado que aun habia que reconocer.
    retonadas, _ = retonar_heredadas(wb)
    hojas_papel = aplicar_sustrato(wb)
    ISSUES.append(f"Sistema visual Swiss Industrial Print: sustrato de papel en "
                  f"{hojas_papel} hojas (al nivel de columna, sin estilo por "
                  f"celda) y {retonadas} celdas del maestro retonadas.")

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
