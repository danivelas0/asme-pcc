"""Pruebas de la capa de navegacion (Dashboard + VBA) del entregable Rev. 4.

A diferencia de test_build_db.py, que solo ejerce funciones puras, estas
pruebas abren el .xlsm ya construido. Se saltan si no existe, para que el
paquete de pruebas siga corriendo en una copia limpia del repo.

    python build_db_materiales.py --resources ... --in ... --out ...\\Rev4.xlsm
    set MOTOR_XLSM=..\\..\\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
    python -m pytest test_dashboard.py -q
"""
from __future__ import annotations

import json
import os
import re
import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_db_materiales as B

AQUI = Path(__file__).resolve().parent
# scripts -> Base_Datos_Materiales_ASME -> outputs, donde vive el entregable.
POR_DEFECTO = AQUI.parents[1] / "Motor_de_Calculo_ASME_PCC_Rev4.xlsm"
RUTA = Path(os.environ.get("MOTOR_XLSM", POR_DEFECTO))

# Funciones de matriz dinamica prohibidas por la regla 1 del libro. La capa de
# navegacion no las necesita y no puede introducirlas por la puerta de atras.
DINAMICAS = ("_xlfn", "FILTER(", "XLOOKUP(", "UNIQUE(", "SORT(", "VSTACK(", "SEQUENCE(")


@pytest.fixture(scope="module")
def wb():
    if not RUTA.exists():
        pytest.skip(f"no existe {RUTA}; construya el libro antes de correr estas pruebas")
    return openpyxl.load_workbook(RUTA, keep_vba=True)


@pytest.fixture(scope="module")
def fuente_vba():
    d = AQUI / "vba"
    return {p.name: p.read_text(encoding="utf-8") for p in d.glob("*.vba")}


# ---------------------------------------------------------------------------
# Visibilidad
# ---------------------------------------------------------------------------
class TestVisibilidad:
    def test_dashboard_es_la_primera_hoja(self, wb):
        assert wb.sheetnames[0] == B.DASH

    def test_solo_el_dashboard_es_visible(self, wb):
        visibles = [s.title for s in wb.worksheets if s.sheet_state == "visible"]
        assert visibles == [B.DASH]

    def test_las_navegables_estan_hidden(self, wb):
        ocultas = {s.title for s in wb.worksheets if s.sheet_state == "hidden"}
        assert ocultas == set(B.NAVEGABLES)

    def test_el_resto_es_veryhidden(self, wb):
        # Se deriva del conjunto, no de un numero fijo: si el builder anade una
        # base nueva, la prueba sigue siendo correcta y la base sigue tapada.
        esperado = {s.title for s in wb.worksheets} - {B.DASH} - set(B.NAVEGABLES)
        very = {s.title for s in wb.worksheets if s.sheet_state == "veryHidden"}
        assert very == esperado

    def test_ninguna_base_que_alimente_un_motor_es_alcanzable(self, wb):
        """Las DB_* y MAP_* que alimentan un motor no pueden estar ni visibles
        ni en el menu Mostrar: son insumo auditado, no interfaz. Abrirlas
        dejaria leer un valor sin la cascada, sin el bloqueo por rango y sin la
        nota del material.

        Las nueve de la Seccion II son la excepcion, y es DECLARADA: ahi el
        entregable ES la hoja de datos —no alimentan ningun motor y no llevan
        formulas—, asi que ocultarlas no protegeria nada y las haria inutiles.
        """
        expuestas = [s.title for s in wb.worksheets
                     if s.sheet_state != "veryHidden"
                     and s.title not in B.NAV_SECII
                     and re.match(r"^(DB_|MAP_|Notas_Codigo|Datos_Ref|_)", s.title)]
        assert expuestas == []

    def test_las_hojas_de_seccion_ii_no_llevan_formulas(self, wb):
        """Son hojas de datos. Una formula aqui rompe la regla 1 del libro.

        Se mira el TIPO de la celda, no si el texto empieza por «=»: el codigo
        imprime celdas como «= 3.18 mm in any 1.524 m», que son texto y tienen
        que seguir siendolo. El builder las fuerza a `s` (ver `_txt_celda`)
        justo para eso; si un dia dejara de hacerlo, esta prueba lo veria.
        """
        for nombre in B.NAV_SECII:
            for fila in wb[nombre].iter_rows():
                for c in fila:
                    assert c.data_type != "f", \
                        f"{nombre}!{c.coordinate} guarda una formula"


# ---------------------------------------------------------------------------
# Botones y claves de destino
# ---------------------------------------------------------------------------
def _claves(ws):
    """(celda del boton, clave de destino) de cada hipervinculo de la hoja."""
    out = []
    for row in ws.iter_rows():
        for c in row:
            if c.hyperlink is not None:
                out.append((c, ws.cell(c.row, B.COL_CLAVE_BASE + c.column).value))
    return out


def _claves_del_arbol(wb):
    """Union de las claves de todas las hojas de navegacion (Dashboard + NAV_*)."""
    return {k for h in B.HOJAS_NAV for _, k in _claves(wb[h])}


class TestBotones:
    def test_el_arbol_enlaza_todas_las_hojas_navegables(self, wb):
        """Comparacion por CONJUNTO, no por lista ordenada: con la tarjeta
        clicable entera cada tarjeta produce cuatro claves iguales, y ademas
        una misma hoja se enlaza desde varios sitios (tarjeta, miga de pan,
        boton de retorno). Lo que importa es la cobertura, no la multiplicidad.

        El Dashboard entra en la union porque es la clave del boton de retorno
        de NAV_ASME, pero no esta en NAVEGABLES: es la unica hoja visible.
        """
        assert _claves_del_arbol(wb) == set(B.NAVEGABLES) | {B.DASH}

    def test_ninguna_clave_es_huerfana(self, wb):
        for k in _claves_del_arbol(wb):
            assert k in wb.sheetnames, f"la clave {k!r} no corresponde a ninguna hoja"

    def test_las_claves_no_se_pisan(self, wb):
        """Tres tarjetas comparten banda, y cada tarjeta ocupa cuatro filas:
        todas sus celdas de clave deben caer en (fila, columna) distintas."""
        for hoja in B.HOJAS_NAV:
            celdas = [(c.row, B.COL_CLAVE_BASE + c.column)
                      for c, _ in _claves(wb[hoja])]
            assert len(celdas) == len(set(celdas)), f"{hoja} pisa una celda de clave"

    def test_el_hipervinculo_apunta_a_destino_inocuo(self, wb):
        # Un hipervinculo directo a la hoja destino seria invalido: la hoja
        # esta oculta cuando se hace clic. El evento VBA hace la navegacion.
        for hoja in B.HOJAS_NAV:
            for c, _ in _claves(wb[hoja]):
                assert c.hyperlink.location == f"'{B.DASH}'!A1"

    def test_cada_navegable_tiene_enlace_de_retorno(self, wb):
        """Subir lleva al PADRE, no a la raiz: la clave del boton de retorno es
        el nombre de la hoja padre, igual que la de cualquier otra tarjeta."""
        for nombre in B.NAVEGABLES:
            claves = [k for _, k in _claves(wb[nombre])]
            assert B.PADRE[nombre] in claves, f"{nombre} no vuelve a su padre"

    def test_el_retorno_esta_desbloqueado_en_las_hojas_protegidas(self, wb):
        # Parche_PCC2_Art212, Collar_PCC2_Art206 e Instrucciones se entregan
        # protegidas.
        for nombre in ("Parche_PCC2_Art212", "Collar_PCC2_Art206", "Instrucciones"):
            ws = wb[nombre]
            assert ws.protection.sheet, f"{nombre} deberia entregarse protegida"
            botones = [c for c, k in _claves(ws) if k == B.PADRE[nombre]]
            assert botones and all(c.protection.locked is False for c in botones)

    def test_las_columnas_de_clave_estan_ocultas(self, wb):
        from openpyxl.utils import get_column_letter as L
        for nombre in [B.DASH] + B.NAVEGABLES:
            ws = wb[nombre]
            for c, _ in _claves(ws):
                col = L(B.COL_CLAVE_BASE + c.column)
                assert ws.column_dimensions[col].hidden, f"{nombre}!{col} no esta oculta"

    def test_las_claves_no_invaden_ningun_layout(self, wb):
        """La columna base debe quedar por encima de todo lo que usan las hojas."""
        for nombre in B.NAVEGABLES:
            ws = wb[nombre]
            usadas = [c.column for row in ws.iter_rows() for c in row
                      if c.value is not None and c.column <= B.COL_CLAVE_BASE]
            assert not usadas or max(usadas) < B.COL_CLAVE_BASE, (
                f"{nombre} usa columnas hasta {max(usadas)}, colisiona con las claves")


# ---------------------------------------------------------------------------
# El arbol de navegacion
# ---------------------------------------------------------------------------
# Aqui esta el valor real del rediseno: que el arbol declarado en Python sea
# el arbol que quedo grabado en el libro, que sea conexo (ningun motor
# huerfano) y que las tarjetas marcador no naveguen a ninguna parte.
def _nodos_con_hoja():
    """(nodo, ancestros) de cada nodo que tiene hoja propia, en preorden."""
    def rec(nodo, ruta):
        if not nodo.hoja:
            return
        yield nodo, ruta
        for h in nodo.hijos:
            yield from rec(h, ruta + [nodo])
    return list(rec(B.ARBOL, []))


class TestArbolDeNavegacion:
    def test_cada_hoja_enlaza_exactamente_a_sus_hijos_padre_y_manual(self, wb):
        """Ni de menos -un hijo inalcanzable- ni de mas -una hoja enlazada
        desde un nivel que no es el suyo, que es lo que el arbol elimina."""
        for nodo, ruta in _nodos_con_hoja():
            esperado = {h.hoja or h.destino for h in nodo.hijos if h.cargado}
            if ruta:                       # el Dashboard no tiene padre,
                esperado |= {n.hoja for n in ruta}   # ni miga de pan,
                esperado.add("Instrucciones")        # ni boton de manual
            assert {k for _, k in _claves(wb[nodo.hoja])} == esperado, nodo.hoja

    def test_todo_destino_se_alcanza_desde_el_dashboard(self, wb):
        """Conexidad: bajando por las claves desde la raiz se llega a las 31.

        Es la comprobacion que de verdad protege el rediseno: un motor que se
        quedase fuera del arbol seguiria existiendo en el libro, seguiria
        estando `hidden`, y no habria forma de abrirlo.
        """
        vistas, pila = set(), [B.DASH]
        while pila:
            hoja = pila.pop()
            if hoja in vistas:
                continue
            vistas.add(hoja)
            if hoja in B.HOJAS_NAV:
                pila.extend(k for _, k in _claves(wb[hoja]))
        assert set(B.NAVEGABLES) - vistas == set()

    def test_padre_no_tiene_ciclos_y_toda_cadena_acaba_en_el_dashboard(self):
        for nombre in B.NAVEGABLES:
            visto, cur = set(), nombre
            while cur != B.DASH:
                assert cur not in visto, f"ciclo de PADRE en {nombre}"
                visto.add(cur)
                cur = B.PADRE[cur]

    def test_las_tarjetas_marcador_no_llevan_hipervinculo_ni_clave(self, wb):
        """Son rotulos de documento -cero dato normativo- y no deben navegar.

        Se cuentan contra el arbol: si alguien marcase `cargado=True` en una
        norma que no esta en el libro, la tarjeta pasaria a tener clave y el
        conteo dejaria de cuadrar.
        """
        marcadores = [n for n in _todos(B.ARBOL) if not n.cargado]
        vistos = 0
        for hoja in B.HOJAS_NAV:
            ws = wb[hoja]
            for row in ws.iter_rows(min_col=1, max_col=B.DASH_NCOLS):
                for c in row:
                    if c.value != B.TXT_NO_CARGADO:
                        continue
                    vistos += 1
                    assert c.hyperlink is None, f"{hoja}!{c.coordinate}"
                    assert ws.cell(c.row, B.COL_CLAVE_BASE + c.column).value is None
        assert vistos == len(marcadores) > 0


def _todos(nodo):
    yield nodo
    for h in nodo.hijos:
        yield from _todos(h)


# ---------------------------------------------------------------------------
# Tabla B-1 del Apendice B: motor propio, y todas sus filas alcanzables
# ---------------------------------------------------------------------------
class TestTablaB1:
    """Lo que se rompio antes y no puede volver a romperse.

    Cuando B-1 vivia dentro de Buscar_NoMetalicos, la clave de seleccion era
    solo la designacion de material. Como el codigo publica PE2708 bajo tres
    especificaciones distintas con HDS distinto, 17 de las 36 filas quedaban
    INALCANZABLES y la ficha mezclaba campos de varias. Estas pruebas fijan que
    cada fila impresa tiene su propia clave.
    """

    def test_las_36_filas_de_cada_edicion_tienen_clave_propia(self, wb):
        for hoja in ("DB_B31_B1", "DB_B31_B1C"):
            w = wb[hoja]
            ids = [w.cell(r, 1).value for r in range(B.R_DATA, w.max_row + 1)]
            k4 = [w.cell(r, B.CB1["k4"]).value for r in range(B.R_DATA, w.max_row + 1)]
            assert len(ids) == 36, hoja
            assert len(set(ids)) == 36, f"{hoja}: material_id repetido"
            # k4 es la clave con la que la cascada resuelve la fila: si se
            # repitiera, MATCH tomaria siempre la primera y el resto del bloque
            # quedaria inalcanzable, que es exactamente el defecto de antes.
            assert len(set(k4)) == 36, f"{hoja}: clave de cascada repetida"

    def test_las_tres_filas_de_pe2708_son_distinguibles(self, wb):
        w = wb["DB_B31_B1"]
        cm = {w.cell(B.R_HDR, c).value: c for c in range(1, w.max_column + 1)}
        specs = {w.cell(r, cm["Spec. No. (ASTM)"]).value
                 for r in range(B.R_DATA, w.max_row + 1)
                 if w.cell(r, cm["Designacion de material"]).value == "PE2708"}
        assert specs == {"D2737", "D3035", "F714"}

    def test_b1_es_lo_unico_del_apendice_b_en_el_libro(self, wb):
        """Un dato, un motor.

        `Buscar_NoMetalicos` y su `DB_NoMetalicos` —las Tablas B-2 a B-6— se
        retiraron en la Rev. 4d por decision de alcance. Si volvieran, habria
        dos sitios distintos por donde entrar al Apendice B y podrian divergir
        sin que nada fallase.
        """
        assert "DB_NoMetalicos" not in wb.sheetnames
        assert "Buscar_NoMetalicos" not in wb.sheetnames
        b = [s for s in wb.sheetnames if "B31_B1" in s]
        assert sorted(b) == ["Buscar_B31_B1", "DB_B31_B1", "DB_B31_B1C"]

    def test_el_buscador_tiene_conmutador_de_unidades(self, wb):
        ws = wb["Buscar_B31_B1"]
        assert ws["D5"].value == "SI"
        origenes = [dv.formula1 for dv in ws.data_validations.dataValidation
                    if dv.sqref is not None and "D5" in str(dv.sqref)]
        assert origenes and origenes[0] == '"SI,US"'

    def test_el_buscador_no_lleva_matrices_dinamicas(self, wb):
        for row in wb["Buscar_B31_B1"].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for mala in DINAMICAS:
                        assert mala not in c.value, f"{c.coordinate}: {c.value}"

    def test_el_estado_bloquea_por_limite_antes_que_por_banda_tabulada(self):
        """El orden de las comprobaciones no es cosmetico.

        F441/CPVC4120-05 imprime limite maximo 93,3 °C y su ultimo HDS a 82 °C.
        Si la banda tabulada se comprobara antes que el limite recomendado, 90 y
        95 °C darian el mismo aviso y se perderia la distincion entre «el codigo
        no publica el dato» y «el material no se recomienda ahi».
        """
        f = B.formula_estado_b1("F", "N", "TMIN", "TMAX", "TPRI", "TULT", "T")
        assert f.index(B.EST_B1_BAJO) < f.index(B.EST_B1_ALTO_LIM) \
            < f.index(B.EST_B1_ALTO_TAB) < f.index(B.EST_B1_NOTA3)

    def test_el_valor_queda_bloqueado_en_cuanto_el_estado_dice_fuera_de_rango(self):
        f = B.formula_valor_b1("F", "EST", "VTAB")
        assert 'SEARCH("FUERA DE RANGO"' in f
        assert B.VAL_B1_BLOQUEADO in f


# ---------------------------------------------------------------------------
# La ficha dice QUE publica la tabla de la que sale la fila
# ---------------------------------------------------------------------------
class TestFuncionDeLaTabla:
    def test_a1_y_a4_tienen_descripcion_distinta(self):
        assert B.FUNCION_TABLA["A-1"] != B.FUNCION_TABLA["A-4"]
        assert "PERNERIA" in B.FUNCION_TABLA["A-4"]
        assert "METALES" in B.FUNCION_TABLA["A-1"]

    def test_las_ediciones_us_describen_lo_mismo_que_su_gemela(self):
        """A-1C no es otra tabla: es la misma en otras unidades. Describirlas
        distinto haria creer que publican cosas distintas."""
        for si, us in (("A-1", "A-1C"), ("A-4", "A-4C")):
            assert B.FUNCION_TABLA[si] == B.FUNCION_TABLA[us]

    def test_el_buscador_del_b31_3_muestra_la_funcion_en_la_ficha(self, wb):
        ws = wb["Buscar_B31_3"]
        etiquetas = {c.value for row in ws.iter_rows(max_col=B.NCOLS) for c in row
                     if isinstance(c.value, str)}
        assert "Funcion de la tabla" in etiquetas
        formulas = [c.value for row in ws.iter_rows(max_col=B.NCOLS) for c in row
                    if isinstance(c.value, str) and c.value.startswith("=")]
        assert any(B.FUNCION_TABLA["A-4"] in f for f in formulas), \
            "la ficha no cita la funcion de la Tabla A-4"
        assert any(B.FUNCION_TABLA["A-1"] in f for f in formulas)


# ---------------------------------------------------------------------------
# Portabilidad de las formulas del Dashboard
# ---------------------------------------------------------------------------
class TestPortabilidadDelDashboard:
    def test_sin_funciones_de_matriz_dinamica(self, wb):
        for row in wb[B.DASH].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for mala in DINAMICAS:
                        assert mala not in c.value, f"{c.coordinate}: {c.value}"

    def test_sin_validaciones_de_origen_no_portable(self, wb):
        for dv in wb[B.DASH].data_validations.dataValidation:
            if dv.formula1:
                assert not re.search(r"OFFSET|INDIRECT", dv.formula1, re.IGNORECASE)


# ---------------------------------------------------------------------------
# Sistema visual: un solo sustrato en las 70 hojas
# ---------------------------------------------------------------------------
# El libro entero va en Swiss Industrial Print (ver el bloque de tokens del
# builder). Lo que estas pruebas fijan no es el gusto sino la CONSISTENCIA: que
# no haya dos sistemas visuales a la vez, que es lo que pasaba mientras las tres
# hojas del maestro conservaban los azules y el Arial de la Rev. 0.
#
# Se auditan los ESTILOS EFECTIVAMENTE USADOS, resueltos por su indice: recorrer
# 2,4 millones de celdas y quedarse con sus style_id cuesta segundos, y evita el
# falso positivo de auditar styles.xml entero -que conserva entradas heredadas
# del maestro que ya no referencia ninguna celda-.
PALETA = {B.PAPEL, B.PAPEL_2, B.TINTA, B.TINTA_2, B.ROJO, B.GRIS, B.GRIS_2,
          B.VERDE, B.AMBAR, B.AMBAR_TXT, B.AMARILLO}
FUENTES_DEL_SISTEMA = {B.MONO, B.MACRO}

# El AMARILLO entra en la paleta del libro pero NO es de uso libre: marca la
# celda editable y solo existe en los dos motores de calculo (leyenda impresa en
# la fila 3 de cada uno). La lista blanca de arriba se resuelve por indice de
# estilo y no sabe de que hoja viene cada uno, asi que el alcance lo fija una
# prueba aparte, por nombre de hoja —mismo patron que el rojo del aviso—.
HOJAS_CON_AMARILLO = {"Parche_PCC2_Art212", "Collar_PCC2_Art206"}


def _rgb6(v):
    """Los seis digitos de color de un rgb, sin el alfa. openpyxl devuelve
    "FF1A1A1A" en lo que leyo del maestro y "00050505" en lo que escribio a
    partir de una cadena de seis digitos: son el mismo color."""
    return v[-6:].upper() if isinstance(v, str) and len(v) >= 6 else None


@pytest.fixture(scope="module")
def estilos_usados(wb):
    """(estilos de celda con contenido, estilos de toda celda con formato)."""
    con_texto, todos = set(), set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if not c.has_style:
                    continue
                todos.add(c.style_id)
                if c.value is not None:
                    con_texto.add(c.style_id)
    return con_texto, todos


class TestSistemaVisual:
    def test_ningun_relleno_fuera_de_la_paleta(self, wb, estilos_usados):
        _, todos = estilos_usados
        fuera = set()
        for sid in todos:
            fill = wb._fills[wb._cell_styles[sid].fillId]
            rgb = _rgb6(getattr(getattr(fill, "fgColor", None), "rgb", None))
            if fill.patternType and rgb and rgb not in PALETA:
                fuera.add(rgb)
        assert fuera == set(), f"rellenos fuera de la paleta: {sorted(fuera)}"

    def test_ningun_color_de_texto_fuera_de_la_paleta(self, wb, estilos_usados):
        con_texto, _ = estilos_usados
        fuera = set()
        for sid in con_texto:
            f = wb._fonts[wb._cell_styles[sid].fontId]
            rgb = _rgb6(getattr(f.color, "rgb", None) if f.color is not None else None)
            if rgb and rgb not in PALETA:
                fuera.add(rgb)
        assert fuera == set(), f"colores de texto fuera de la paleta: {sorted(fuera)}"

    def test_toda_celda_con_texto_usa_una_de_las_dos_fuentes(self, wb, estilos_usados):
        """Solo las celdas CON contenido: una celda vacia no tiene texto que
        mostrar, y las de relleno de una banda fusionada arrastran la fuente por
        omision del libro sin que se vea nunca."""
        con_texto, _ = estilos_usados
        fuera = set()
        for sid in con_texto:
            nombre = wb._fonts[wb._cell_styles[sid].fontId].name
            if nombre and nombre not in FUENTES_DEL_SISTEMA:
                fuera.add(nombre)
        assert fuera == set(), f"fuentes fuera del sistema: {sorted(fuera)}"

    def test_las_70_hojas_llevan_el_sustrato_de_papel(self, wb):
        """El sustrato se graba al nivel de COLUMNA, no celda a celda: es lo que
        permite papelar las 54 198 filas del volcado de la Seccion II sin
        multiplicar el tamano del archivo. Si una hoja se quedara sin el, se
        abriria en blanco de Excel en medio de un libro de papel."""
        for ws in wb.worksheets:
            dim = ws.column_dimensions.get("A")
            assert dim is not None and dim.fill is not None, ws.title
            assert _rgb6(getattr(dim.fill.fgColor, "rgb", None)) == B.PAPEL, ws.title
            assert dim.font is not None and dim.font.name == B.MONO, ws.title

    def test_el_aviso_de_macros_es_el_unico_relleno_rojo(self, wb):
        """El rojo del acento marca bloqueo, y como RELLENO solo aparece en el
        aviso de macros del Dashboard: si apareciera en mas sitios dejaria de
        significar algo.

        Basta con la celda ancla: el aviso esta fusionado a lo ancho y Excel
        aplica a todo el rango el formato de su esquina superior izquierda.
        """
        rojas = [(ws.title, c.coordinate)
                 for ws in wb.worksheets
                 for row in ws.iter_rows(max_col=B.DASH_NCOLS)
                 for c in row
                 if c.fill is not None and c.fill.patternType
                 and _rgb6(getattr(c.fill.fgColor, "rgb", None)) == B.ROJO]
        assert rojas == [(B.DASH, f"A{B.FILA_AVISO}")]

    def test_el_amarillo_solo_vive_en_los_dos_motores_de_calculo(self, wb):
        """El amarillo significa «aqui escribe usted». Fuera de un motor de
        calculo no hay nada que escribir, y si apareciera en un buscador o en
        una hoja de datos dejaria de significar eso.

        Se recorre con el mismo ancho acotado que el guardia del rojo: el
        amarillo solo puede nacer en la banda A..G de un motor.
        """
        hojas = {ws.title
                 for ws in wb.worksheets
                 for row in ws.iter_rows(max_col=B.DASH_NCOLS)
                 for c in row
                 if c.fill is not None and c.fill.patternType
                 and _rgb6(getattr(c.fill.fgColor, "rgb", None)) == B.AMARILLO}
        assert hojas <= HOJAS_CON_AMARILLO, sorted(hojas - HOJAS_CON_AMARILLO)

    def test_los_dos_motores_llevan_su_leyenda_de_color(self, wb):
        """La leyenda tiene que estar impresa en la hoja: un color que hay que
        adivinar no es una convencion, es un acertijo. Y cada muestra lleva el
        relleno que de verdad describe."""
        esperado = {celda: _rgb6(relleno.fgColor.rgb)
                    for celda, _, relleno in B.LEYENDA_MOTOR}
        for nombre in HOJAS_CON_AMARILLO:
            ws = wb[nombre]
            for celda, rgb in esperado.items():
                c = ws[celda]
                assert c.value, f"{nombre}!{celda} sin texto de leyenda"
                assert _rgb6(getattr(c.fill.fgColor, "rgb", None)) == rgb, \
                    f"{nombre}!{celda}"

    def test_toda_celda_editable_de_un_motor_va_en_amarillo(self, wb):
        """El reverso de la leyenda, y lo que la hace verdad: no basta con que
        lo amarillo sea editable —hay que comprobar que TODO lo editable esta
        amarillo—, o el ingeniero encontraria celdas que acepta Excel y que la
        hoja no anuncia. El boton de volver se entrega desbloqueado a proposito
        (ver reponer_botones_de_retorno) y se excluye por su hipervinculo.
        """
        for nombre in HOJAS_CON_AMARILLO:
            ws = wb[nombre]
            # La cola de un rango fusionado no se comprueba: no tiene contenido
            # propio y Excel pinta todo el rango con el formato de la celda
            # ancla (la trampa que ya documenta franja() en el builder). Ademas
            # openpyxl le arrastra la proteccion del ancla pero no su relleno,
            # asi que mirarla daria un falso positivo en cada boton.
            fuera = [c.coordinate
                     for row in ws.iter_rows(max_col=B.MOTOR_NCOLS)
                     for c in row
                     if not isinstance(c, openpyxl.cell.cell.MergedCell)
                     and c.hyperlink is None
                     and c.protection is not None
                     and c.protection.locked is False
                     and _rgb6(getattr(c.fill.fgColor, "rgb", None)) != B.AMARILLO]
            assert fuera == [], f"{nombre}: editables sin amarillo: {fuera}"

    def test_el_semaforo_conserva_sus_tres_estados(self, wb):
        """Retonados, no eliminados: el estado de la consulta es informacion de
        seguridad y se lee de un vistazo. Verde y ambar solo viven en el formato
        condicional del indicador de seleccion.

        El color de un formato diferencial se lee de `fgColor`: en un dxf
        openpyxl deja `bgColor` en negro y no es el que pinta.
        """
        ws = wb["Buscar_B31_3"]
        reglas = [r for rango in ws.conditional_formatting
                  for r in rango.rules if r.dxf is not None]
        rellenos = {_rgb6(getattr(r.dxf.fill.fgColor, "rgb", None))
                    for r in reglas if r.dxf.fill is not None}
        assert {B.VERDE, B.AMBAR} <= rellenos


# ---------------------------------------------------------------------------
# El proyecto VBA sobrevive al round-trip de openpyxl
# ---------------------------------------------------------------------------
class TestVba:
    def test_el_binario_vba_sigue_en_el_paquete(self, wb):
        assert "xl/vbaProject.bin" in zipfile.ZipFile(RUTA).namelist()

    def test_la_extension_es_xlsm(self):
        assert RUTA.suffix == ".xlsm"


# ---------------------------------------------------------------------------
# El desanclado total no puede cambiar el Art. 212
# ---------------------------------------------------------------------------
# parche_art212_ref.json es el *oracle*: la hoja 212 tal como la entregaba el
# maestro Rev0, congelada desde el ultimo build valido. Reconstruirla en codigo
# (build_parche_art212) debe reproducir cada formula, literal y validacion al
# caracter. Sin Excel aqui, esta es la unica red: si una transcripcion difiere,
# falla ahora y no en la mesa del ingeniero.
def cargar_oracle_parche():
    p = Path(__file__).resolve().parent / "parche_art212_ref.json"
    return json.loads(p.read_text(encoding="utf-8"))


def _plano(v):
    """Texto VISIBLE de una celda para comparar contra el oracle. La Fase 9
    convierte los simbolos 'base_sub' de la columna B en CellRichText con
    subindice real; el texto visible es la concatenacion de los runs SIN el '_'
    (p.ej. 'Fm'). Se compara asi (no reconstruyendo el '_') porque el round-trip
    del .xlsm no conserva el vertAlign de los runs; el oracle guarda ya la forma
    visible de esas celdas (ver actualizar_oracle_subindices en el commit). Una
    celda normal (str/num) se devuelve tal cual."""
    from openpyxl.cell.rich_text import CellRichText
    return str(v) if isinstance(v, CellRichText) else v


def _asserta_sin_guion_bajo(ws):
    """Decision 5: ningun simbolo VISIBLE de la columna B (la de simbolos) lleva
    un guion bajo — los subindices son reales (CellRichText). La columna A es
    prosa/rotulos (y de la Seccion 7 trae guiones bajos que no son subindices:
    rutas, tags '[MODO_S]'), asi que no se recorre. Un str con '_' en B es un
    simbolo sin convertir (fallo); un CellRichText no debe renderizar '_'."""
    from openpyxl.cell.rich_text import CellRichText
    hubo_rich = False
    for cell in ws["B"]:
        v = cell.value
        if isinstance(v, CellRichText):
            visible = str(v)
            assert "_" not in visible, f"{cell.coordinate}: {visible!r} lleva '_'"
            if len(list(v)) > 1:
                hubo_rich = True
        elif isinstance(v, str):
            # La columna B tambien aloja formulas (specs) y textos largos cuyo
            # '_' no es un subindice; se excluyen igual que en el converter.
            if v.startswith("=") or len(v) > 12:
                continue
            assert "_" not in v, f"{cell.coordinate}: simbolo sin convertir {v!r}"
    assert hubo_rich, "no se encontro ningun simbolo con subindice real"


class TestParidadHojaParche:
    HOJA = "Parche_PCC2_Art212"

    def test_las_dos_clases_de_divergencia_son_disjuntas(self):
        # Una celda no puede estar RETIRADA (debe estar vacia) y REEMPLAZADA (debe
        # tener contenido) a la vez: seria una contradiccion silenciosa.
        comunes = set(B.DIVERGENCIAS_DECLARADAS) & set(B.DIVERGENCIAS_REEMPLAZADAS)
        assert not comunes, comunes

    def test_todas_las_formulas_y_literales(self, wb):
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        for celda, esperado in oracle["formulas"].items():
            motivo = B.DIVERGENCIAS_DECLARADAS.get((self.HOJA, celda))
            if motivo:
                # Celda RETIRADA: se exige que este VACIA. Una divergencia
                # declarada que resulta traer otro valor es un error, no una
                # divergencia.
                assert ws[celda].value is None, f"{celda}: {motivo}"
                continue
            reempl = B.DIVERGENCIAS_REEMPLAZADAS.get((self.HOJA, celda))
            if reempl:
                # Celda REEMPLAZADA: sale de la igualdad contra el oracle (su
                # correccion la prueba TestCascadaDimensionalArt212), pero se
                # exige NO vacia — si quedo vacia, es una retirada disfrazada.
                assert ws[celda].value is not None, f"{celda}: {reempl}"
                continue
            assert _plano(ws[celda].value) == esperado, celda

    def test_toda_divergencia_declara_motivo(self):
        for d in (B.DIVERGENCIAS_DECLARADAS, B.DIVERGENCIAS_REEMPLAZADAS):
            for (hoja, celda), motivo in d.items():
                assert isinstance(motivo, str) and len(motivo) > 20, (hoja, celda)

    def test_validaciones_de_datos(self, wb):
        # Una validacion del ORACLE cuyo sqref cae entero en celdas divergentes
        # (retiradas o reemplazadas) ya no se exige tal cual. Y una validacion
        # REAL cuyo sqref cae entero en celdas REEMPLAZADAS es nueva a proposito
        # (el rango de la cascada) y su correccion la prueba
        # TestCascadaDimensionalArt212, asi que se excluye de la comparacion. Una
        # validacion parcialmente divergente seguiria exigiendose completa: aqui
        # no se da el caso, pero el guardia es honesto sobre ello.
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]

        def coords(sqref):
            rng = openpyxl.worksheet.cell_range.CellRange(sqref)
            return [f"{openpyxl.utils.get_column_letter(c)}{r}"
                    for r in range(rng.min_row, rng.max_row + 1)
                    for c in range(rng.min_col, rng.max_col + 1)]

        def todas_divergentes(sqref, tablas):
            return all(any((self.HOJA, c) in t for t in tablas) for c in coords(sqref))

        def en_anexo(sqref):
            # Una validacion del ANEXO DE PASOS DEL FLUJO (filas >= FILA_ANEXO)
            # es nueva por diseno: el oracle Rev0 solo cubre 1-129 y estas las
            # fijan las anclas por-paso (test_paso*). Se excluyen de la paridad.
            rng = openpyxl.worksheet.cell_range.CellRange(sqref)
            return rng.min_row >= B.FILA_ANEXO_FLUJO_212

        reales = sorted(
            ({"sqref": str(dv.sqref), "formula1": dv.formula1}
             for dv in ws.data_validations.dataValidation
             if not todas_divergentes(str(dv.sqref), (B.DIVERGENCIAS_REEMPLAZADAS,))
             and not en_anexo(str(dv.sqref))),
            key=lambda d: d["sqref"])
        esperadas = [
            v for v in oracle["validaciones"]
            if not todas_divergentes(v["sqref"],
                                     (B.DIVERGENCIAS_DECLARADAS, B.DIVERGENCIAS_REEMPLAZADAS))]
        assert reales == sorted(esperadas, key=lambda d: d["sqref"])

    def test_rangos_fusionados(self, wb):
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        # Los merges del anexo de pasos del flujo (bandas de las filas >=
        # FILA_ANEXO) son nuevos por diseno y salen de la paridad contra el
        # oracle Rev0 (que solo cubre 1-129).
        reales = sorted(
            str(r) for r in ws.merged_cells.ranges
            if openpyxl.worksheet.cell_range.CellRange(str(r)).min_row
            < B.FILA_ANEXO_FLUJO_212)
        assert reales == oracle["fusionados"]


# ---------------------------------------------------------------------------
# Tareas 7-8: el bloque dimensional del Art. 212 sale de DB_B36 por cascada
# ---------------------------------------------------------------------------
# Prueba lo que el oracle Rev0 ya no cubre (las celdas REEMPLAZADAS): la cascada
# norma -> NPS -> cedula y el lookup de OD/espesor contra DB_B36. Sin Excel no se
# recalcula el valor resuelto -eso lo hace verificar.py con el motor real-; aqui
# se fija la ESTRUCTURA: origen de rango (no lista fija), norma con sus dos
# opciones, y OD/espesor leyendo de las dos ediciones y no de Datos_Ref.
class TestCascadaDimensionalArt212:
    HOJA = "Parche_PCC2_Art212"

    def test_norma_nps_y_cedula_tienen_validacion_de_rango(self, wb):
        ws = wb[self.HOJA]
        origen = {}
        for dv in ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                origen[str(rng)] = str(dv.formula1 or "")
        # D22 = norma, D18 = NPS, D19 = cedula (la norma no corre las filas 18-21).
        for celda in ("D18", "D19", "D22"):
            f1 = origen.get(celda, "")
            assert f1.startswith("="), f"{celda}: origen {f1!r}, se esperaba un rango"
            assert not f1.startswith('"'), f"{celda}: sigue siendo una lista fija"

    def test_la_norma_ofrece_las_dos_y_solo_las_dos(self, wb):
        ws = wb[self.HOJA]
        col = B.COL_NORMA_B36
        vals = [ws.cell(r, col).value for r in range(B.R_DATA, B.R_DATA + 2)]
        assert vals == ["B36.10M", "B36.19M"]

    def test_od_y_espesor_leen_de_las_dos_ediciones(self, wb):
        ws = wb[self.HOJA]
        for celda in ("D20", "D21"):
            f = str(ws[celda].value or "")
            assert "DB_B36_10" in f and "DB_B36_19" in f, f"{celda}: {f!r}"
            assert "Datos_Ref" not in f, f"{celda}: aun lee de Datos_Ref"


class TestCascadaDimensionalArt206:
    """La misma cascada norma -> NPS -> cedula -> OD/espesor del Art. 212
    (Tareas 7-8), aplicada al Art. 206 (Tarea 9). El Art. 206 no tiene oracle
    Rev0 (nace 100% en codigo), asi que su correccion la fijan estos tests y no
    DIVERGENCIAS_*: toda la paridad que aqui se exige es contra DB_B36."""
    HOJA = "Collar_PCC2_Art206"

    def test_norma_nps_y_cedula_tienen_validacion_de_rango(self, wb):
        ws = wb[self.HOJA]
        origen = {}
        for dv in ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                origen[str(rng)] = str(dv.formula1 or "")
        # D33 = norma, D18 = NPS, D19 = cedula. La norma va en la fila 33 (libre al
        # final de la seccion 1) y no corre las filas 18-32: moverlas reapuntaria
        # las formulas de calculo aguas abajo (D45/D46 usan D20, D54/D64 usan D21),
        # mismo criterio que la fila 22 del Art. 212.
        for celda in ("D18", "D19", "D33"):
            f1 = origen.get(celda, "")
            assert f1.startswith("="), f"{celda}: origen {f1!r}, se esperaba un rango"
            assert not f1.startswith('"'), f"{celda}: sigue siendo una lista fija"

    def test_la_norma_ofrece_las_dos_y_solo_las_dos(self, wb):
        ws = wb[self.HOJA]
        col = B.COL_NORMA_B36
        vals = [ws.cell(r, col).value for r in range(B.R_DATA, B.R_DATA + 2)]
        assert vals == ["B36.10M", "B36.19M"]

    def test_od_y_espesor_leen_de_las_dos_ediciones(self, wb):
        ws = wb[self.HOJA]
        for celda in ("D20", "D21"):
            f = str(ws[celda].value or "")
            assert "DB_B36_10" in f and "DB_B36_19" in f, f"{celda}: {f!r}"
            assert "Datos_Ref" not in f, f"{celda}: aun lee de Datos_Ref"


# ---------------------------------------------------------------------------
# Las dos fuentes de verdad -Python y VBA- no pueden divergir
# ---------------------------------------------------------------------------
# La tabla de navegacion vive por duplicado: en build_db_materiales.py, que la
# graba en el archivo, y en vba/mod_nav.vba, que la reaplica al abrir. Si se
# separan, el libro se abre mostrando algo distinto de lo que se construyo.
class TestSincroniaPythonVba:
    def test_misma_lista_de_hojas_navegables(self, fuente_vba):
        # La lista se arma concatenando, no con Array(...): VBA no admite mas
        # de 25 continuaciones de linea y hoy hacen falta 31 nombres. Cada
        # nombre va en su linea, precedido del separador salvo el primero.
        src = fuente_vba["mod_nav.vba"]
        cuerpo = (src.split("Public Function HojasNavegables()", 1)[1]
                     .split("End Function", 1)[0])
        nombres = [t.lstrip("|") for t in re.findall(r'"([^"]+)"', cuerpo)]
        assert [n for n in nombres if n] == B.NAVEGABLES

    def test_misma_columna_base_de_claves(self, fuente_vba):
        m = re.search(r"COL_CLAVE_BASE\s+As\s+Long\s*=\s*(\d+)", fuente_vba["mod_nav.vba"])
        assert m and int(m.group(1)) == B.COL_CLAVE_BASE

    def test_misma_celda_de_aviso(self, fuente_vba):
        m = re.search(r'CELDA_AVISO\s+As\s+String\s*=\s*"([A-Z]+)(\d+)"',
                      fuente_vba["mod_nav.vba"])
        assert m and int(m.group(2)) == B.FILA_AVISO

    def test_mismo_texto_de_aviso_de_macros(self, wb, fuente_vba):
        """El VBA reescribe A4 al abrir. Si su texto no es el que grabo el
        builder, el aviso cambia de aspecto en cuanto se abre el libro —y era
        justo lo que pasaba con el rotulo antiguo, sin las barras ASCII—."""
        m = re.search(r'TXT_INACTIVAS As String = _\s*\n\s*"([^"]+)"',
                      fuente_vba["mod_nav.vba"])
        assert m, "no se encontro TXT_INACTIVAS en mod_nav.vba"
        assert wb[B.DASH].cell(B.FILA_AVISO, 1).value == m.group(1)

    def test_misma_hoja_de_inicio(self, fuente_vba):
        assert re.search(rf'HOJA_INICIO\s+As\s+String\s*=\s*"{B.DASH}"',
                         fuente_vba["mod_nav.vba"])

    def test_no_queda_la_clave_literal_de_retorno(self, fuente_vba):
        """La clave es SIEMPRE el nombre de la hoja destino, tambien al subir.

        Si alguien reintrodujera CLAVE_VOLVER en el VBA, el boton de retorno
        volveria a llevar a la raiz desde cualquier nivel, que es justo lo que
        el arbol elimina. La rutina que lo resolvia tampoco puede volver.
        """
        for nombre, texto in fuente_vba.items():
            # Se miran solo las lineas de codigo: los comentarios de cabecera
            # SI nombran lo que desaparecio, y explicar por que es util.
            codigo = "\n".join(l for l in texto.splitlines()
                               if not l.lstrip().startswith("'"))
            assert "CLAVE_VOLVER" not in codigo, nombre
            assert "VolverAlDashboard" not in codigo, nombre


# ---------------------------------------------------------------------------
# El lint de VBA que evita colgar Excel al guardar
# ---------------------------------------------------------------------------
class TestLintVba:
    def test_las_fuentes_pasan_el_lint(self, fuente_vba):
        import make_vba_seed as S
        for nombre, texto in fuente_vba.items():
            assert S.lint_vba(texto, nombre) == []

    def test_el_lint_detecta_declaracion_tras_rutina(self):
        import make_vba_seed as S
        malo = "\n".join(['Public Const A As String = "x"',
                          "Public Sub Uno()", "End Sub",
                          'Private Const B As String = "y"'])
        fallos = S.lint_vba(malo, "malo.vba")
        assert len(fallos) == 1 and "despues de la primera rutina" in fallos[0]

    def test_el_lint_detecta_el_exceso_de_continuaciones(self):
        """Pasarse de 25 no da un error legible: AddFromString deja el modulo
        VACIO y el sintoma que se ve luego es un "Sub o Function no definida"
        en un dialogo modal que cuelga Excel. Se atrapa aqui, en Python."""
        import make_vba_seed as S
        n = S.MAX_CONTINUACIONES + 2
        largo = ["Public Function F() As Variant", "    F = Array( _"]
        largo += [f'        "h{i}", _' for i in range(n)]
        largo += ['        "ultima")', "End Function"]
        fallos = S.lint_vba("\n".join(largo), "malo.vba")
        assert len(fallos) == 1 and "continuaciones" in fallos[0]

    def test_el_lint_no_confunde_lineas_logicas_consecutivas(self):
        """Cada linea logica cuenta sus propias continuaciones: muchas lineas
        de dos fisicas seguidas no pueden sumar hasta disparar el aviso."""
        import make_vba_seed as S
        cuerpo = ["Public Sub S()"]
        for i in range(S.MAX_CONTINUACIONES * 3):
            cuerpo += [f'    Debug.Print "a{i}", _', f'          "b{i}"']
        cuerpo.append("End Sub")
        assert S.lint_vba("\n".join(cuerpo), "bueno.vba") == []


# ---------------------------------------------------------------------------
# build_parche_art212(): prueba aislada, con stubs, sin Excel
# ---------------------------------------------------------------------------
# A diferencia de TestParidadHojaParche (que compara el .xlsm REAL, ya
# construido por build_parche_art212 desde main(), contra el oracle), esta
# clase llama a build_parche_art212() directamente sobre un libro de usar y
# tirar, sin resources/ ni el maestro real: no hay Excel en este entorno y el
# build completo tarda ~45 s, asi que un stub minimo es la forma rapida y
# aislada de probar la funcion sin pagar ese costo en cada corrida.
#
# Los *_stub() reproducen solo lo que construir_seccion7_material consume de
# verdad de cada dict (ver build_b313/build_iid/build_map_factores/
# build_listas, las funciones reales que main() usa para armarlos). La
# mayoria de los campos son bookkeeping de tamano de tabla (last_row,
# pack_t0/pack_v0/npts_col, CK/CV/FK/FV/SK/SV/GK/GV/FAM/K4/maxV...): no son
# valores normativos (no son un S admisible, un factor, una formula del
# codigo — Regla n.1 no aplica) y solo alimentan celdas D109-D113/D116-D124,
# fuera de ANCLAS_T2, asi que cualquier entero o rango con forma valida sirve.
#
# La UNICA excepcion es rangos["B313"]["ID"] / rangos["IID1A"]["ID"]: esos dos
# rangos SI quedan incrustados, letra a letra, dentro de las formulas D115 y
# E115 (material_id resuelto: INDEX(rb["ID"], ...) / INDEX(ri["ID"], ...)),
# que si estan en ANCLAS_T2. Para que la formula reproducida case con la del
# *oracle*, se fijan a los mismos dos rangos que el *oracle* ya capturo del
# .xlsm real (parche_art212_ref.json): "DB_B31_3!$A$4:$A$1291" (D115) y
# "DB_BPVC_IID!$A$4:$A$1802" (E115, via IID1A) — no son valores inventados,
# son los que ya se verificaron contra el libro real; aqui solo se reutilizan
# para que el stub no rompa el unico dato que el test compara.
def b313_stub():
    """Dict minimo de la base B31.3 (build_b313 real: sheet/last_row/pack_t0/
    pack_v0/npts_col). El nombre de hoja SI importa (tiene que casar con la
    hoja sembrada en TestBuildParcheContraOracle._construir()); last_row
    coincide con el que capturo el *oracle* (ver nota de modulo) y los demas
    son bookkeeping arbitrario, fuera del alcance de ANCLAS_T2."""
    return {"sheet": "DB_B31_3", "last_row": 1291,
            "pack_t0": 40, "pack_v0": 60, "npts_col": 31}


def iid1a_stub():
    """Idem para la Tabla 1A de II-D (build_iid real, edicion SI, tabla '1A')."""
    return {"sheet": "DB_BPVC_IID", "last_row": 1802,
            "pack_t0": 40, "pack_v0": 60, "npts_col": 31}


def iidb_stub():
    """Idem para las Tablas 1B/3 de II-D (build_iid real, tabla 'B'). last_row
    es arbitrario: iidb no alimenta ninguna celda de ANCLAS_T2."""
    return {"sheet": "DB_BPVC_IID_B", "last_row": 700,
            "pack_t0": 40, "pack_v0": 60, "npts_col": 31}


def fac_stub():
    """Idem para MAP_Factores (build_map_factores real). D128 (en ANCLAS_T2)
    es el VALOR de entrada por defecto del lookup de Ej ('A106 | Seamless
    pipe'), fijo en construir_seccion7_material y no derivado de este stub;
    sheet/last_row solo alimentan la formula de E128 y la validacion de
    datos, fuera del alcance de esta tarea."""
    return {"sheet": "MAP_Factores", "last_row": 50}


def rangos_stub():
    """Dict minimo de build_listas real, solo para B313 e IID1A (los dos
    unicos que construir_seccion7_material lee de `rangos`; iidb no entra por
    esta via). Ver la nota de modulo: "ID" es el unico campo que tiene que
    casar con el *oracle* letra a letra."""
    def bloque(id_range):
        return {
            "FAM": f"DB_Listas!$A${B.R_DATA}:$A${B.R_DATA + 9}",
            "CK": f"DB_Listas!$B${B.R_DATA}:$B${B.R_DATA + 9}",
            "CV": f"DB_Listas!$C${B.R_DATA}:$C${B.R_DATA + 9}",
            "maxC": 1,
            "FK": f"DB_Listas!$D${B.R_DATA}:$D${B.R_DATA + 9}",
            "FV": f"DB_Listas!$E${B.R_DATA}:$E${B.R_DATA + 9}",
            "maxF": 1,
            "SK": f"DB_Listas!$F${B.R_DATA}:$F${B.R_DATA + 9}",
            "SV": f"DB_Listas!$G${B.R_DATA}:$G${B.R_DATA + 9}",
            "maxS": 1,
            "GK": f"DB_Listas!$H${B.R_DATA}:$H${B.R_DATA + 9}",
            "GV": f"DB_Listas!$I${B.R_DATA}:$I${B.R_DATA + 9}",
            "maxG": 1,
            "ID": id_range,
            "K4": f"DB_Listas!$J${B.R_DATA}:$J${B.R_DATA + 9}",
            "maxV": 1,
        }
    return {
        "B313": bloque("DB_B31_3!$A$4:$A$1291"),
        "IID1A": bloque("DB_BPVC_IID!$A$4:$A$1802"),
    }


class TestBuildParcheContraOracle:
    """Construye Parche_PCC2_Art212 en un wb de usar y tirar y lo compara contra
    el *oracle*, sin regenerar el entregable. Cada bloque ANCLAS_SECCION_*
    (y ANCLAS_T2) enumera las celdas de esa seccion que la funcion ya debe
    reproducir letra a letra."""

    def _construir(self):
        import build_db_materiales as B
        wb = openpyxl.Workbook()
        # Bases minimas que la Seccion 7 referencia por nombre (Excel resuelve al
        # abrir; aqui solo deben existir para que el seed de la Fase 1 valide).
        for n in ("DB_B31_3", "Datos_Ref", "MAP_Factores", "DB_BPVC_IID",
                  "DB_BPVC_IID_B", "DB_B36_10", "DB_B36_19"):
            wb.create_sheet(n)
        # Sembrar en DB_B31_3 los dos material_id del caso precargado para que
        # el guardia de corregir_art212_fase1 (ahora en build_parche_art212) pase.
        db = wb["DB_B31_3"]
        db.cell(B.R_DATA, 1, B.SEED_ART212_BASE)
        db.cell(B.R_DATA + 1, 1, B.SEED_ART212_COLLAR)
        # Stubs de las dos bases dimensionales (Tareas 7-8): solo hacen falta el
        # nombre de hoja y los conteos que dimensionan las listas de la cascada.
        b36 = lambda sheet: {"sheet": sheet, "last_row": B.R_DATA + 9,
                             "n_nps": 5, "max_ced": 5}
        # Stub de los coeficientes del App. 501 (Paso 8) — en el motor real los
        # lee leer_energia_501() de resources/ (Fase 0.2). Aqui son fixtures del
        # test, como b313_stub, con los valores que imprime el codigo.
        energia = {"tnt_div_kg": 4266920, "blast_thr_J": 8130000, "blast_R_m": 30.0,
                   "r_scaled_def": 20.0,
                   "r_scaled_ops": [("20 (Glass windows)", 20.0),
                                    ("12 (Eardrum rupture / Concrete block panels)", 12.0),
                                    ("6 (Lung damage / Brick walls)", 6.0),
                                    ("2 (Fatal)", 2.0)]}
        B.build_parche_art212(wb, b313_stub(), iid1a_stub(), iidb_stub(),
                              fac_stub(), rangos_stub(),
                              b36("DB_B36_10"), b36("DB_B36_19"),
                              energia_501=energia)
        return wb["Parche_PCC2_Art212"]

    # F90 (DICTAMEN GLOBAL) sale de esta lista: la Fase 1 lo reescribe para
    # anteponer la compuerta de elegibilidad (Paso 1). Su forma nueva la fija
    # test_paso1_elegibilidad y su divergencia frente al oracle Rev0 se declara
    # en DIVERGENCIAS_REEMPLAZADAS.
    # Una celda anclada que ademas esta DECLARADA como divergencia no se
    # compara contra el oracle: se exige lo que su clase promete (vacia si esta
    # retirada, no vacia si esta reemplazada). Antes estas listas se limitaban a
    # comparar, asi que declarar una divergencia obligaba a BORRAR la celda de
    # la lista — y con ella su cobertura. Asi la celda sigue vigilada y en un
    # solo sitio se dice por que difiere.
    def _ancla(self, ws, oracle, celda):
        clave = (self.HOJA if hasattr(self, "HOJA") else "Parche_PCC2_Art212", celda)
        motivo = B.DIVERGENCIAS_DECLARADAS.get(clave)
        if motivo:
            assert ws[celda].value is None, f"{celda}: {motivo}"
            return
        reempl = B.DIVERGENCIAS_REEMPLAZADAS.get(clave)
        if reempl:
            assert ws[celda].value is not None, f"{celda}: {reempl}"
            return
        assert _plano(ws[celda].value) == oracle["formulas"][celda], celda

    ANCLAS_T2 = ("A1", "A2", "D5", "D6", "D47", "E47", "D58", "E58",
                 "D67", "D68", "D72", "D46", "E46", "D60")

    def test_anclas_de_la_tarea_2(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_T2:
            assert _plano(ws[celda].value) == oracle["formulas"][celda], celda

    # --- Fase 1: compuerta de elegibilidad (Paso 1, 212-1/212-2) -------------
    # Bloque nuevo en el anexo de pasos del flujo (filas 134+), fuera del rango
    # 1-129 del oracle: se fija por cadena aqui (Regla n.1: trazado a Art. 212 y
    # al flujo aprobado del ingeniero), no contra el oracle Rev0.
    # F90 evoluciona por fases; esta es su forma vigente. Fase 1 antepuso la
    # compuerta de elegibilidad; Fase 4 anadio los dos topes de filete (F161,
    # F162) al AND de las verificaciones.
    F90_CON_ELEGIBILIDAD = (
        '=IF(OR(LEFT($D$140,9)="PROHIBIDO",LEFT($D$140,11)="NO ELEGIBLE",'
        'LEFT($D$140,16)="FUERA DE ALCANCE"),$D$140,'
        'IF(OR($D$57="SIN MATERIAL SELECCIONADO",'
        '$E$57="SIN MATERIAL SELECCIONADO"),"ELIJA MATERIAL (Seccion 2)",'
        'IF(OR($D$57<>"OK",$E$57<>"OK"),"REVISAR — MATERIAL FUERA DE RANGO",'
        'IF(AND(F112="CUMPLE",F113="CUMPLE",F114="CUMPLE",F115="CUMPLE",'
        'F161="CUMPLE",F162="CUMPLE"),'
        '"APTO","REVISAR"))))')
    D140_DICTAMEN_ELEGIBILIDAD = (
        '=IF($D$138="Letal / extrema peligrosidad",'
        '"PROHIBIDO — servicio letal: usar Art. 201 o reemplazo de seccion '
        '(flujo 212 · 212-2a remite a Part 1)",'
        'IF($D$137="No","NO ELEGIBLE — dano no caracterizable (212-2c)",'
        'IF($D$139="Si — activa/no analizada",'
        '"NO ELEGIBLE — grieta activa o no analizada (212-2c)",'
        'IF($D$25>345,"FUERA DE ALCANCE — T > 345 (evaluar creep/fatiga 212-1e)",'
        'IF($D$25<0,"REVISAR — T < 0 (evaluar tenacidad a la entalla 212-1e)",'
        '"ELEGIBLE")))))')

    def test_paso1_elegibilidad(self):
        ws = self._construir()
        # Banda del bloque nuevo.
        assert str(ws["A134"].value).startswith("PASO 1"), ws["A134"].value
        # Cuatro entradas categoricas sembradas con el caso precargado.
        assert ws["D136"].value == "Adelgazamiento local"
        assert ws["D137"].value == "Si"
        assert ws["D138"].value == "General"
        assert ws["D139"].value == "No"
        # Dictamen de elegibilidad y F90 con la compuerta antepuesta.
        assert ws["D140"].value == self.D140_DICTAMEN_ELEGIBILIDAD, "D140"
        assert ws["F118"].value == self.F90_CON_ELEGIBILIDAD, "F118"
        # Las cuatro entradas llevan validacion de lista (no se teclean libres).
        origen = {}
        for dv in ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                origen[str(rng)] = str(dv.formula1 or "")
        for celda, contiene in (("D136", "Adelgazamiento"), ("D137", "Si,No"),
                                ("D138", "Letal"), ("D139", "arrestada")):
            assert celda in origen, f"{celda} sin validacion"
            assert contiene in origen[celda], f"{celda}: {origen[celda]!r}"

    # --- Fase 2: cargas de presion y externas combinadas (Paso 2, 212-3.2) ---
    # ec.(1) F_CP=P*Dm/2 [29], ec.(2) F_LP=P*Dm/4 [33/Fase 0.1], F_C=F_CP+F_CO
    # [37], F_L=F_LP+F_LO [39], F_max=MAX(F_C,F_L); esfera/cabezal (D11=3) ->
    # hand-off 212-3.2(c) [44], F_max=NA(). Bloque nuevo del anexo (filas 142+),
    # pura adicion: no toca ninguna celda del oracle.
    def test_paso2_cargas(self):
        ws = self._construir()
        assert str(ws["A142"].value).startswith("PASO 2"), ws["A142"].value
        # Entradas de cargas externas (un valor cada una, aplican a los 2 casos).
        assert ws["D144"].value == 0
        assert ws["D145"].value == 0
        # Fuerzas de presion por caso (Operacion D / Diseno E). La Fase 2 retiro
        # la tercera columna: ver TestModeloDePresionDosCasos.
        assert ws["D146"].value == "=D90*$D$80/2", "F_CP Op"
        assert ws["E146"].value == "=E90*$D$80/2", "F_CP Dis"
        assert ws["D147"].value == "=D90*$D$80/4", "F_LP Op"
        # Totales y gobernante.
        assert ws["D148"].value == "=D146+$D$144", "F_C Op"
        assert ws["D149"].value == "=D147+$D$145", "F_L Op"
        assert ws["D150"].value == "=IF($D$11=3,NA(),MAX(D148,D149))", "F_max Op"
        assert ws["E150"].value == "=IF($D$11=3,NA(),MAX(E148,E149))", "F_max Dis"

    # --- Fase 3: proximidad a discontinuidades (Paso 3, 212-3.3) -------------
    # L_min = 2*sqrt(Rm*t) (ec.3, ya en D73). Nuevo: 3C distancia a parches
    # adyacentes [51] y nota de esquinas redondeadas (R_min=75 mm, del flujo).
    def test_paso3_discontinuidad(self):
        ws = self._construir()
        assert str(ws["A153"].value).startswith("PASO 3"), ws["A153"].value
        assert ws["D156"].value == (
            '=IF($D$155="","No aplica (sin parche adyacente)",'
            'IF($D$155>=$D$101,"OK","< L_min — reubicar (212-3.3)"))'), "D156"
        # La nota de esquinas cita el 75 mm (recomendacion del flujo).
        assert "75" in str(ws["D157"].value), ws["D157"].value

    # --- Fase 4: filete perimetral con topes (Paso 4, 212-3.4) ---------------
    # w_min pasa a F_max (ec.4, E=0.55, bloques [57],[59]); topes NOTA [60]:
    # w <= min(T,t) y w <= 40 mm; ambos entran al AND de F90.
    def test_paso4_filete(self):
        ws = self._construir()
        # w_min (D64:E64) recablea a F_max (D150:E150), no F_m (D63). La
        # columna F se retiro en la Fase 2 (modelo de presion de dos casos).
        assert ws["D92"].value == "=D150/($D$70*$D$69)", "D92"
        # Bloque de topes.
        assert str(ws["A159"].value).startswith("PASO 4"), ws["A159"].value
        assert ws["D161"].value == "=MIN($D$29,$D$21)", "D161 req"
        assert ws["E161"].value == "=$D$32", "E161 adoptado"
        assert ws["F161"].value == (
            '=IF(E161<=D161,"CUMPLE",'
            '"NO CUMPLE — excede espesor menor (NOTA 212-3.4)")'), "F161"
        assert ws["F162"].value == (
            '=IF(E162<=D162,"CUMPLE","NO CUMPLE — excede 40 mm (NOTA 212-3.4)")'), "F162"
        assert ws["D162"].value == 40, "D162"

    # --- Fase 5: excentricidad literal con separacion g (Paso 5, 212-3.4c) ---
    # e = (T+t+g)/2 con g anadida si g>=1.5 (212-4c [82], g = fit-up del faying
    # edge, entrada nueva D167, distinta de la luz radial D31). S_w literal de
    # la ec.(5) [65,67]: P*Dm/(2T) + 3*P*Dm*e/T^2, NA en esfera (D11=3). Para
    # cilindro sin g el valor no cambia (el S_w anterior con kf=0.5 ya era la
    # ec.5); el cambio real es esfera->NA y el mecanismo de g.
    def test_paso5_excentricidad(self):
        ws = self._construir()
        assert str(ws["A165"].value).startswith("PASO 5"), ws["A165"].value
        assert ws["D167"].value == 0, "g default 0 (fit-up ajustado)"
        assert ws["D83"].value == (
            "=($D$29+$D$21+IF($D$167>=1.5,$D$167,0))/2"), "D55 e con g"
        assert ws["D85"].value == (
            "=IF($D$11=3,NA(),$D$80/(2*$D$29)*(1+6*$D$83/$D$29))"), "D57 C_sw"
        assert ws["D94"].value == (
            "=IF($D$11=3,NA(),D90*$D$80/(2*$D$29))"), "D66 membrana"
        assert ws["D95"].value == (
            "=IF($D$11=3,NA(),3*D90*$D$80*$D$83/$D$29^2)"), "D67 flexion"
        assert ws["E95"].value == (
            "=IF($D$11=3,NA(),3*E90*$D$80*$D$83/$D$29^2)"), "E67 flexion Dis"

    # --- Fase 6: conformado en frio simple/doble (Paso 6, 212-3.5) -----------
    # %Elong = coef*T/Rf*(1-Rf/Ro), coef=75 doble (esfera/cabezal, ec.6 [71]) o
    # 50 simple (cilindro, ec.7 [75]); Ro=infinito (plano) -> factor=1 [73].
    def test_paso6_conformado(self):
        ws = self._construir()
        assert str(ws["A170"].value).startswith("PASO 6"), ws["A170"].value
        assert ws["D172"].value == "", "Ro default blanco (plano)"
        assert ws["D105"].value == (
            '=IF($D$11=3,75,50)*$D$29/$D$84*IF($D$172="",1,1-$D$84/$D$172)'), "D105"

    # --- Fase 7: fabricacion (Paso 7, 212-4) — avisos ------------------------
    # Aviso de separacion (g>5 no admisible, g>=1.5 -> e incluye g, 212-4c),
    # aviso MT/PT si T_parche>25 mm (212-4a), y notas de fabricacion (secuencia,
    # prep 40 mm, corte termico, venteo). Aditivo: no toca oracle ni F90.
    def test_paso7_fabricacion(self):
        ws = self._construir()
        assert str(ws["A175"].value).startswith("PASO 7"), ws["A175"].value
        assert ws["D177"].value == (
            '=IF($D$167>5,"SEPARACION > 5 mm: fit-up no admisible (212-4c)",'
            'IF($D$167>=1.5,"g >= 1.5 mm: e incluye g (212-4c) — ver Paso 5",'
            '"fit-up ajustado (g < 1.5 mm)"))'), "D177"
        assert ws["D178"].value == (
            '=IF($D$29>25,"T > 25 mm: examinar bordes de preparacion por MT/PT '
            '(laminaciones), 212-4a","T <= 25 mm: sin examen de bordes por espesor")'), "D178"
        assert "40 mm" in str(ws["D179"].value), ws["D179"].value

    # --- Fase 8: NDE y prueba de hermeticidad (Paso 8, 212-5/6 + App.501) ----
    # Selector hidro/neumatica; en neumatica E (II-1 general en k), TNT (II-3),
    # R (III-1). Constantes leidas de resources/ (Fase 0.2, aqui via stub).
    def test_paso8_prueba(self):
        ws = self._construir()
        assert str(ws["A181"].value).startswith("PASO 8"), ws["A181"].value
        assert ws["D183"].value == "Hidrostatica", "default hidrostatica"
        assert ws["D188"].value == 20.0, "Rscaled default 20"
        assert ws["D189"].value == (
            '=IF($D$183="Neumatica",(1/($D$187-1))*($D$185*1000000)*$D$184*'
            '(1-($D$186/$D$185)^(($D$187-1)/$D$187)),NA())'), "D189 E"
        assert ws["D190"].value == (
            '=IF($D$183="Neumatica",$D$189/4266920,NA())'), "D190 TNT"
        assert ws["D191"].value == (
            '=IF($D$183="Neumatica",IF($D$189<=8130000,30,'
            '$D$188*(2*$D$190)^(1/3)),NA())'), "D191 R"
        # Selectores como lista.
        origen = {}
        for dv in ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                origen[str(rng)] = str(dv.formula1 or "")
        assert "Neumatica" in origen.get("D183", ""), origen.get("D183")
        assert origen.get("D188", "") == '"20,12,6,2"', origen.get("D188")

    # --- Fase 9: simbolos con subindice real (decision 5) -------------------
    def test_simbolos_sin_guion_bajo(self):
        _asserta_sin_guion_bajo(self._construir())

    # Aplicacion y codigo de construccion (filas 10-14). Cubre TODAS las
    # celdas que el oracle declara en ese rango: A/B/C/D/G de las cinco filas
    # (incluidas B10-B14 y C10-C14, no solo D10-D14 + los rotulos A10-A14).
    ANCLAS_SECCION_3 = (
        "A10", "B10", "C10", "D10", "G10",
        "A11", "B11", "C11", "D11", "G11",
        "A12", "B12", "C12", "D12", "G12",
        "A13", "B13", "C13", "D13", "G13",
        "A14", "B14", "C14", "D14", "G14",
    )

    def test_seccion_3(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_3:
            self._ancla(ws, oracle, celda)

    # Seccion 1, datos de entrada (filas 16-35). Incluye la banda A16 (texto
    # NUEVO que reemplaza TEXTOS_HEREDADOS, pero identico al que el *oracle*
    # ya capturo — ver build_parche_art212) y el encabezado de fila 17
    # (Parametro/Simbolo/Unidad/Valor/Referencia). Cubre TODAS las celdas que
    # el oracle declara en el rango 16-35: A/B/C/D/G de las 18 filas de datos
    # (B18-B35 y C18-C35 incluidas, no solo D18-D35 + A16 + G22/G23) — mismo
    # criterio de exhaustividad que ANCLAS_SECCION_3.
    ANCLAS_SECCION_4 = (
        "A16",
        "A17", "B17", "C17", "D17", "G17",
        "A18", "B18", "C18", "D18", "G18",
        "A19", "B19", "C19", "D19", "G19",
        "A20", "B20", "C20", "D20", "G20",
        "A21", "B21", "C21", "D21", "G21",
        "A22", "B22", "C22", "D22", "G22",
        "A23", "B23", "C23", "D23", "G23",
        "A24", "B24", "C24", "D24",
        "A25", "B25", "C25", "D25", "G25",
        "A26", "B26", "C26", "D26", "G26",
        # Fase 2: la fila 27 ("Presion de diseno tipica") se RETIRO entera y
        # A28/B28/G28 se reescriben ("Presion de diseno (maxima admisible /
        # rating)", simbolo P_dis). Salen del ancla contra el oracle: las
        # primeras estan en DIVERGENCIAS_DECLARADAS y las segundas en
        # DIVERGENCIAS_REEMPLAZADAS. Su forma nueva la fija
        # TestModeloDePresionDosCasos. C28/D28 (unidad y valor) no cambian.
        "C28", "D28",
        "A29", "B29", "C29", "D29", "G29",
        "A30", "B30", "C30", "D30", "G30",
        "A31", "B31", "C31", "D31", "G31",
        "A32", "B32", "C32", "D32", "G32",
        "A33", "B33", "C33", "D33", "G33",
        "A34", "B34", "C34", "D34", "G34",
        "A35", "B35", "C35", "D35", "G35",
    )

    def test_seccion_4(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_4:
            motivo = B.DIVERGENCIAS_DECLARADAS.get(("Parche_PCC2_Art212", celda))
            if motivo:
                assert ws[celda].value is None, f"{celda}: {motivo}"
                continue
            # Celda REEMPLAZADA por el bloque dimensional (Tareas 7-8): sale de la
            # igualdad contra el oracle y solo se exige no-vacia; su correccion la
            # prueba TestCascadaDimensionalArt212.
            if ("Parche_PCC2_Art212", celda) in B.DIVERGENCIAS_REEMPLAZADAS:
                assert ws[celda].value is not None, celda
                continue
            assert _plano(ws[celda].value) == oracle["formulas"][celda], celda

    # Seccion 2, esfuerzos admisibles y factores (filas 39-48). Incluye la
    # banda A37 (fusionada A37:G37, "PARAMETROS DE CALCULO (constantes -
    # editables)") y el encabezado de fila 38 (Parametro/Simbolo/Unidad/
    # Valor/Referencia): preceden inmediatamente el rango 39-48 y titulan
    # esta seccion, no una anterior. Cubre TODAS las celdas que el *oracle*
    # declara en el rango 37-48 — A/B/C/D/G de las diez filas de datos
    # (B39-B48 y C39-C48 incluidas, sin G45 porque el *oracle* no lo
    # declara), mismo criterio de exhaustividad que ANCLAS_SECCION_3 y
    # ANCLAS_SECCION_4. D39/D40/D44 ya los cubre ANCLAS_T2; se repiten aqui
    # para que esta seccion quede completa por si sola — redundante pero
    # correcto.
    ANCLAS_SECCION_5 = (
        "A65",
        "A66", "B66", "C66", "D66", "G66",
        "A67", "B67", "C67", "D67", "G67",
        "A68", "B68", "C68", "D68", "G68",
        "A69", "B69", "C69", "D69", "G69",
        "A70", "B70", "C70", "D70", "G70",
        "A71", "B71", "C71", "D71", "G71",
        "A72", "B72", "C72", "D72", "G72",
        "A73", "B73", "C73", "D73",
        "A74", "B74", "C74", "D74", "G74",
        "A75", "B75", "C75", "D75", "G75",
        "A76", "B76", "C76", "D76", "G76",
    )

    def test_seccion_5(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_5:
            self._ancla(ws, oracle, celda)

    # Geometria y propiedades derivadas (filas 52-57). Incluye la banda A50
    # (fusionada A50:G50, "2.  GEOMETRIA Y PROPIEDADES DERIVADAS", con el
    # numeral y el doble espacio tal como los imprime el *oracle*) y el
    # encabezado de fila 51 (Parametro/Simbolo/Unidad/Valor/"Formula /
    # Referencia" — distinto de "Referencia / Notas" de las filas 17/38):
    # preceden inmediatamente el rango 52-57 y titulan esta seccion, no una
    # anterior. Cubre TODAS las celdas que el *oracle* declara en el rango
    # 50-57 — A/B/C/D/G de las seis filas de datos, mismo criterio de
    # exhaustividad que ANCLAS_SECCION_3/4/5. El *oracle* no declara ninguna
    # validacion de datos en este rango.
    ANCLAS_SECCION_6 = (
        "A78",
        "A79", "B79", "C79", "D79", "G79",
        "A80", "B80", "C80", "D80", "G80",
        "A81", "B81", "C81", "D81", "G81",
        "A82", "B82", "C82", "D82", "G82",
        # D55 (e) y D57 (C_sw) salen del oracle: la Fase 5 les anade la
        # separacion g y las pasa a la forma literal de cilindro con NA en
        # esfera. Su forma nueva la fija test_paso5_excentricidad.
        "A83", "B83", "C83", "G83",
        "A84", "B84", "C84", "D84", "G84",
        "A85", "B85", "C85", "G85",
    )

    def test_seccion_6(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_6:
            self._ancla(ws, oracle, celda)

    # Seccion 3, calculo de cargas y soldadura (filas 61-69), mas Seccion 4,
    # resultados del diseno (filas 73-80). Nombrado distinto de
    # "test_seccion_7" a proposito: ese nombre ya lo usaria la "Seccion 7" de
    # cascada de material (construir_seccion7_material, filas 105+), que es
    # un numero de seccion del motor sin relacion con el rango de filas que
    # cubre este metodo (FILAS 61-80, no una "seccion 7"). Incluye la banda
    # A59 y el encabezado de fila 60 (con TRES columnas de valor propias:
    # Operacion/Diseno tipico/Envolvente en D/E/F, no una sola "Valor") y la
    # banda A71 y el encabezado de fila 72 (con G72="Formula / Referencia"):
    # preceden inmediatamente cada uno de los dos rangos y titulan sus
    # propias secciones, no una anterior. Cubre TODAS las celdas que el
    # *oracle* declara en 59-69 y 71-80 — comprobado fila por fila, sin
    # asumir que las tres columnas D/E/F existen en 61-69 (fila 69 no trae
    # B69) ni que A/B/C/D/G existen todas en 73-80 (75 no trae G, 76/77/79/80
    # no traen B). La fila 58 y la fila 70 no aparecen en el *oracle* (ni
    # formula, ni fusionado, ni validacion): quedan vacias, sin ancla.
    ANCLAS_FILAS_61_80 = (
        "A87",
        # Fase 2: la columna F (caso "Envolvente") se RETIRO de las filas 60-69
        # y E60/E61 cambian de contenido (E pasa a ser "Diseno" leyendo D28).
        # Las tres salen del ancla contra el oracle: F* esta en
        # DIVERGENCIAS_DECLARADAS y E60/E61 en DIVERGENCIAS_REEMPLAZADAS, y su
        # forma nueva la fija TestModeloDePresionDosCasos.
        "A88", "B88", "C88", "D88", "G88",
        "A89", "B89", "C89", "D89", "G89",
        "A90", "B90", "C90", "D90", "E90", "G90",
        "A91", "B91", "C91", "D91", "E91", "G91",
        # D64/E64 (w_min) salen del oracle: la Fase 4 los recablea a F_max
        # (D150) en vez de F_m (D63). Su forma nueva la fija test_paso4_filete y
        # su divergencia se declara en DIVERGENCIAS_REEMPLAZADAS. A/B/C/G64
        # (rotulo, simbolo, unidad, referencia) siguen anclados al oracle.
        "A92", "B92", "C92", "G92",
        "A93", "B93", "C93", "D93", "E93", "G93",
        # D66/D67 (componentes membrana y flexion de S_w) salen del oracle: la
        # Fase 5 los pasa a la forma literal de la ec.(5) (P*Dm directa, no via
        # F_m) con NA en esfera. D68 = D66+D67 no cambia de formula (propaga NA).
        "A94", "B94", "C94", "G94",
        "A95", "B95", "C95", "G95",
        "A96", "B96", "C96", "D96", "E96", "G96",
        "A97", "C97", "D97", "E97", "G97",
        "A99",
        "A100", "B100", "C100", "D100", "G100",
        "A101", "B101", "C101", "D101", "G101",
        "A102", "B102", "C102", "D102", "G102",
        "A103", "B103", "C103", "D103",
        "A104", "C104", "D104", "G104",
        # D77 (%Elong) sale del oracle: la Fase 6 le anade la rama simple/doble
        # (coef 50/75) y el factor (1-Rf/Ro). Lo fija test_paso6_conformado.
        "A105", "C105", "G105",
        "A106", "B106", "C106", "D106", "G106",
        "A107", "C107", "D107", "G107",
        "A108", "C108", "D108", "G108",
    )

    def test_filas_61_80(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_FILAS_61_80:
            self._ancla(ws, oracle, celda)

    # Cubre la Seccion 5 (verificaciones, filas 82-90), la Seccion 6
    # (especificaciones tecnicas, filas 93-99), el aviso fijo (fila 101) y
    # DOS gaps que quedaron sin asignar en secciones anteriores:
    #   (a) la banda IDENTIFICACION (A4) + G5/G6;
    #   (b) la banda APLICACION Y CODIGO (A8) + el encabezado de fila 9.
    # Incluye tambien la banda A82 y el encabezado de fila 83 (preceden el
    # rango 84-90 y titulan esta seccion, no una anterior) y la banda A92
    # (precede el rango 93-99; esta seccion no tiene fila de encabezado
    # propia — el *oracle* no la declara). F90 NO se ancla contra el oracle:
    # la Fase 1 lo reescribio (compuerta de elegibilidad) y lo fija
    # test_paso1_elegibilidad; aqui solo se comprueba su rotulo A90.
    ANCLAS_SECCION_8 = (
        "A4", "A5", "D5", "G5", "A6", "D6", "G6",
        "A8",
        "A9", "B9", "C9", "D9", "G9",
        "A110",
        "A111", "D111", "E111", "F111", "G111",
        # Fase 2: D84 (MAX sobre dos casos) y G84 (rotulo "diseno") salen
        # del ancla; su forma la fija TestModeloDePresionDosCasos.
        "A112", "E112", "F112",
        "A113", "D113", "E113", "F113", "G113",
        "A114", "D114", "E114", "F114", "G114",
        # Fase 2: A87/D87 pasan del caso "envolvente" al caso "diseno".
        "E115", "F115", "G115",
        "A116", "D116", "E116", "F116", "G116",
        # Fase 2: E89 lee D28 (diseno maxima admisible) en vez de D27.
        "A117", "D117", "F117", "G117",
        "A118",
        "A120",
        "A121", "B121",
        "A122", "B122",
        "A123", "B123",
        "A124", "B124",
        "A125", "B125",
        "A126", "B126",
        # Fase 2: B99 calcula la presion de prueba sobre D28 (diseno
        # maxima admisible) en vez de D27, retirada. Lo fija
        # TestModeloDePresionDosCasos.
        "A127",
        "A129",
    )

    def test_seccion_8(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_8:
            self._ancla(ws, oracle, celda)

    # --- Fase 2: el modelo de presion pasa de tres casos a dos --------------
    # Lo que el oracle Rev0 ya no cubre. 212-3.2 define una UNICA
    # P = "internal design pressure" para las ec. (1)/(2) y 206-3.3 la nombra
    # "maximum allowable design pressure": el "diseno tipico" intermedio no lo
    # publica ningun codigo y competia con la maxima admisible por gobernar el
    # t_req. Se retira la fila 27 y la columna F de la tabla de cargas; el caso
    # que sobrevive es el MAS conservador, ahora en la columna E.
    def test_la_fila_del_diseno_tipico_esta_retirada(self):
        ws = self._construir()
        for celda in ("A27", "B27", "C27", "D27", "G27"):
            assert ws[celda].value is None, f"{celda} deberia estar vacia"

    def test_la_presion_de_diseno_es_la_maxima_admisible(self):
        ws = self._construir()
        assert ws["A28"].value == "Presión de diseño (máxima admisible / rating)"
        assert _plano(ws["B28"].value) == "Pdis", ws["B28"].value
        assert ws["D28"].value == 20
        # encabezado() del 212 escribe el rotulo LITERAL (no pasa por .upper():
        # el oracle Rev0 trae mayus/minus mixtas en esta tabla).
        assert ws["E88"].value == "Diseño", ws["E88"].value
        assert ws["E89"].value == "=$D$28", ws["E89"].value

    def test_la_columna_envolvente_esta_retirada(self):
        """La tabla de cargas (60-69) y el Paso 2 (143-150) quedan en DOS
        columnas de valor. Si quedara una celda suelta en F, la hoja mostraria
        una tercera columna a medio calcular."""
        ws = self._construir()
        for fila in list(range(60, 70)) + list(range(143, 151)):
            assert ws[f"F{fila}"].value is None, f"F{fila}: {ws[f'F{fila}'].value}"

    def test_las_verificaciones_leen_el_caso_de_diseno(self):
        """El t_req y el filete gobernantes se toman del caso de diseno (E),
        que es la presion maxima admisible. Antes se tomaban de la columna
        'Envolvente' (F), que era ese mismo concepto con otro nombre: el valor
        no cambia, pero la hoja ya no declara un rango que no existe."""
        ws = self._construir()
        assert ws["D112"].value == "=MAX(D92:E92)", ws["D112"].value
        assert ws["D115"].value == "=E93", ws["D115"].value
        assert ws["E117"].value == "=$D$28", ws["E117"].value
        # La presion de prueba hidrostatica sale de la misma presion de diseno.
        # D74 es el factor de prueba hidrostatica tras la Fase 3 (era D46).
        assert "$D$74*$D$28" in str(ws["B127"].value), ws["B127"].value


class TestSemaforoDeAceptacion:
    """Fase 6: verde/rojo en la columna de Resultado y dictamen global en bloque.

    En un motor de calculo el criterio de aceptacion es informacion de
    seguridad: la diferencia entre cumple y no cumple tiene que leerse sin
    leerse. El rojo aparece aqui como relleno por primera vez fuera del aviso de
    macros, con el mismo significado que tiene en todo el libro: bloqueado.
    """

    CASOS = [("Parche_PCC2_Art212", B.SEMAFORO_212, B.FILA_DICTAMEN_212),
             ("Collar_PCC2_Art206", B.SEMAFORO_206, B.FILA_DICTAMEN_206)]

    @pytest.mark.parametrize("hoja,tabla,fila_dict", CASOS)
    def test_el_texto_favorable_existe_en_la_formula(self, wb, hoja, tabla, fila_dict):
        """El fallo silencioso mas probable de esta fase: que el texto que la
        regla considera favorable no sea LETRA POR LETRA el que escribe la
        formula. Una raya larga distinta basta para que no case nunca, y nadie
        se enteraria — la celda saldria roja siempre, que parece un resultado.
        """
        ws = wb[hoja]
        for fila, favorables in tabla.items():
            formula = str(ws.cell(fila, 6).value)
            for texto in favorables:
                assert f'"{texto}"' in formula, f"{hoja}!F{fila}: {texto!r}"

    @pytest.mark.parametrize("hoja,tabla,fila_dict", CASOS)
    def test_cada_resultado_tiene_sus_dos_reglas(self, wb, hoja, tabla, fila_dict):
        ws = wb[hoja]
        por_rango = {}
        for rango in ws.conditional_formatting:
            for r in rango.rules:
                if r.dxf is not None and r.dxf.fill is not None:
                    por_rango.setdefault(str(rango.sqref), []).append(
                        _rgb6(getattr(r.dxf.fill.fgColor, "rgb", None)))
        for fila in tabla:
            colores = por_rango.get(f"F{fila}")
            assert colores == [B.VERDE, B.ROJO], f"{hoja}!F{fila}: {colores}"

    @pytest.mark.parametrize("hoja,tabla,fila_dict", CASOS)
    def test_el_rojo_se_pinta_por_COMPLEMENTO(self, wb, hoja, tabla, fila_dict):
        """La regla roja es 'no es ninguno de los favorables', no una lista de
        textos malos. Asi un #N/A o un texto que nadie previo sale en ROJO, que
        es el lado seguro; enumerar lo malo dejaria lo imprevisto en blanco,
        indistinguible de 'aun no calculado'."""
        ws = wb[hoja]
        for rango in ws.conditional_formatting:
            if not str(rango.sqref).startswith("F"):
                continue
            for r in rango.rules:
                if (r.dxf is not None and r.dxf.fill is not None
                        and _rgb6(getattr(r.dxf.fill.fgColor, "rgb", None)) == B.ROJO):
                    assert "NOT(" in str(r.formula[0]), str(rango.sqref)

    @pytest.mark.parametrize("hoja,tabla,fila_dict", CASOS)
    def test_el_dictamen_global_es_un_bloque_propio(self, wb, hoja, tabla, fila_dict):
        ws = wb[hoja]
        assert "DICTAMEN GLOBAL" in str(ws.cell(fila_dict, 1).value), hoja
        # Banda de tinta a todo el ancho, en macrotipografia grande. La cola de
        # un rango fusionado se salta: Excel pinta todo el rango con el relleno
        # de la celda ANCLA, y openpyxl ni siquiera deja asignarselo (lo mismo
        # pasa en las demas bandas del libro — ver franja() en el builder, que
        # recorre el rango justamente porque el BORDE si necesita ponerse celda
        # a celda y el relleno no).
        for col in range(1, B.MOTOR_NCOLS + 1):
            c = ws.cell(fila_dict, col)
            if isinstance(c, openpyxl.cell.cell.MergedCell):
                assert c.border.top.style == "thick", c.coordinate
                continue
            assert _rgb6(getattr(c.fill.fgColor, "rgb", None)) == B.TINTA, c.coordinate
        # La franja roja de arriba separa el dictamen de la tabla: es el limite
        # duro entre dos zonas, dibujado y no insinuado.
        assert ws.cell(fila_dict, 1).border.top.style == "thick", hoja
        assert ws.cell(fila_dict, 1).font.name == B.MACRO
        assert ws.cell(fila_dict, 1).font.size >= 16
        assert ws.row_dimensions[fila_dict].height >= 20

    @pytest.mark.parametrize("hoja,tabla,fila_dict", CASOS)
    def test_elija_material_va_en_ambar_no_en_rojo(self, wb, hoja, tabla, fila_dict):
        """Falta una entrada, no falla una verificacion. Pintarlo de rojo
        confundiria 'aun no has elegido' con 'no cumple'."""
        ws = wb[hoja]
        rango = f"A{fila_dict}:G{fila_dict}"
        reglas = [r for rg in ws.conditional_formatting if str(rg.sqref) == rango
                  for r in rg.rules if r.dxf is not None and r.dxf.fill is not None]
        colores = {_rgb6(getattr(r.dxf.fill.fgColor, "rgb", None)): str(r.formula[0])
                   for r in reglas}
        assert set(colores) == {B.VERDE, B.AMBAR, B.ROJO}, f"{hoja}: {set(colores)}"
        assert "ELIJA MATERIAL" in colores[B.AMBAR], hoja
        assert "APTO" in colores[B.VERDE], hoja


class TestReglasDeComentario:
    """Fase 5: donde va un comentario y donde no.

    Antes casi toda fila llevaba el MISMO texto en el rotulo y en la celda de
    valor. Duplicar no informa: el ingeniero pasa el raton por la celda que esta
    mirando y un globo repetido solo tapa la hoja.
    """

    CASOS = [("Parche_PCC2_Art212", B.REGLAS_COMENTARIO_212),
             ("Collar_PCC2_Art206", B.REGLAS_COMENTARIO_206)]

    @pytest.mark.parametrize("hoja,reglas", CASOS)
    def test_ninguna_columna_lleva_comentario_de_mas(self, wb, hoja, reglas):
        ws = wb[hoja]
        sobran = []
        for ini, fin, cols, _fuente in reglas:
            permitidas = {ord(c) - 64 for c in cols}
            for fila in range(ini, fin + 1):
                for col in range(1, B.MOTOR_NCOLS + 1):
                    c = ws.cell(fila, col)
                    if c.comment is not None and col not in permitidas:
                        sobran.append(c.coordinate)
        assert sobran == [], f"{hoja}: {sobran}"

    @pytest.mark.parametrize("hoja,reglas", CASOS)
    def test_las_columnas_obligadas_lo_llevan(self, wb, hoja, reglas):
        """Y solo se exige donde hay algo escrito: una seccion de una sola
        columna de valor (el 206) usa la misma regla que una de dos (el 212)
        sin fabricar globos sobre celdas vacias."""
        ws = wb[hoja]
        faltan = []
        for ini, fin, cols, _fuente in reglas:
            for fila in range(ini, fin + 1):
                if not any(ws.cell(fila, c).comment is not None
                           for c in range(1, B.MOTOR_NCOLS + 1)):
                    continue        # fila sin explicacion: no se inventa una
                for letra in cols:
                    c = ws.cell(fila, ord(letra) - 64)
                    if c.value is not None and c.comment is None:
                        faltan.append(c.coordinate)
        assert faltan == [], f"{hoja}: {faltan}"

    def test_las_verificaciones_explican_tambien_el_resultado(self, wb):
        """La columna de Resultado es la que se lee para decidir, asi que lleva
        comentario propio.

        No se exige que diga "CUMPLE": dos de las seis verificaciones del 212 no
        resuelven en CUMPLE/NO CUMPLE sino en una RUTA ("Refuerzo 360" frente a
        "Parche local"; "OK - parche" frente a "Migrar (Art.206)"). Lo que si se
        exige es que el globo del Resultado NO sea el mismo que el del Requerido:
        si lo fuera, la columna que decide estaria explicada con el texto de otra.
        """
        for hoja, filas in (("Parche_PCC2_Art212", range(112, 118)),
                            ("Collar_PCC2_Art206", range(85, 87))):
            ws = wb[hoja]
            for fila in filas:
                res, req = ws.cell(fila, 6), ws.cell(fila, 4)
                assert res.comment is not None, f"{hoja}!F{fila}"
                if req.comment is not None and hoja == "Parche_PCC2_Art212":
                    assert res.comment.text != req.comment.text, f"{hoja}!F{fila}"

    def test_ninguna_banda_lleva_parentesis_explicativo(self, wb):
        """Fase 5: fuera el parentesis aclaratorio de las bandas de seccion. La
        cita al codigo se conserva —es normativa— pero fuera del parentesis, que
        en el resto de la hoja significa 'aclaracion prescindible'."""
        for hoja in ("Parche_PCC2_Art212", "Collar_PCC2_Art206"):
            ws = wb[hoja]
            con_parentesis = []
            # Desde la fila 3: A1/A2 son el TITULO y el subtitulo de la hoja,
            # y su parentesis es la cita del articulo del codigo — lo que esta
            # fase conserva a proposito, no un aclaratorio.
            for fila in range(3, ws.max_row + 1):
                c = ws.cell(fila, 1)
                # Una banda de seccion se reconoce por su relleno de tinta.
                if (c.fill is None or not c.fill.patternType
                        or _rgb6(getattr(c.fill.fgColor, "rgb", None)) != B.TINTA):
                    continue
                if isinstance(c.value, str) and "(" in c.value:
                    con_parentesis.append((c.coordinate, c.value))
            assert con_parentesis == [], f"{hoja}: {con_parentesis}"


class TestDiagnosticoPlegable:
    """Fase 4: el rastro de la resolucion de material va plegado por defecto.

    Once filas de trazabilidad entre la cascada y el resultado. No se borran
    —Regla n.1: son lo que permite auditar de donde sale el S(T)— pero abiertas
    empujan fuera de la pantalla lo que el ingeniero vino a ver.
    """

    # (hoja, primera y ultima fila del grupo, fila del dictamen, fila del S(T))
    BLOQUES = [("Parche_PCC2_Art212", 47, 56, 57, 58),
               ("Collar_PCC2_Art206", 45, 54, 55, 56)]

    @pytest.mark.parametrize("hoja,ini,fin,f_dict,f_st", BLOQUES)
    def test_el_rastro_va_agrupado_y_plegado(self, wb, hoja, ini, fin, f_dict, f_st):
        ws = wb[hoja]
        for r in range(ini, fin + 1):
            dim = ws.row_dimensions[r]
            assert dim.outline_level == 1, f"{hoja}!{r} sin agrupar"
            assert dim.hidden, f"{hoja}!{r} deberia nacer plegada"
        assert ws.sheet_properties.outlinePr.summaryBelow is True, hoja

    @pytest.mark.parametrize("hoja,ini,fin,f_dict,f_st", BLOQUES)
    def test_el_dictamen_y_el_resultado_NO_se_pliegan(self, wb, hoja, ini, fin,
                                                      f_dict, f_st):
        """El dictamen de rango es lo que BLOQUEA el calculo. Una condicion de
        bloqueo escondida detras de un '+' es una condicion que nadie ve."""
        ws = wb[hoja]
        assert str(ws.cell(f_dict, 1).value) == "Dictamen de rango", hoja
        assert str(ws.cell(f_st, 1).value) == "S(T) resuelto", hoja
        for r in (f_dict, f_st):
            dim = ws.row_dimensions[r]
            assert not dim.hidden, f"{hoja}!{r} no puede nacer plegada"
            assert not dim.outline_level, f"{hoja}!{r} no entra en el grupo"

    @pytest.mark.parametrize("hoja,ini,fin,f_dict,f_st", BLOQUES)
    def test_nada_del_rastro_se_perdio(self, wb, hoja, ini, fin, f_dict, f_st):
        """Plegar no es borrar: las once filas conservan su rotulo y su valor."""
        ws = wb[hoja]
        for r in range(ini, fin + 1):
            assert ws.cell(r, 1).value, f"{hoja}!A{r} sin rotulo"
            assert ws.cell(r, 4).value is not None, f"{hoja}!D{r} sin contenido"

    def test_la_nota_de_cascada_no_se_repite_cinco_veces(self, wb):
        """Decia 'Lista desplegable en cascada' en los cinco niveles. Repetir
        la misma frase cinco veces no informa: la vuelve ruido."""
        for hoja, ini in (("Parche_PCC2_Art212", 41), ("Collar_PCC2_Art206", 39)):
            ws = wb[hoja]
            notas = [str(ws.cell(r, 7).value or "") for r in range(ini, ini + 6)]
            assert notas[0], f"{hoja}: el nivel 0 tiene que explicar la cascada"
            assert all(not n for n in notas[1:]), f"{hoja}: {notas}"


class TestNumeracionDeSecciones:
    """Fase 3: la Seccion de Material sube justo detras de los datos de entrada.

    Es lo unico que el *oracle* no puede vigilar —se traslado con el mismo mapa,
    asi que casa igual de bien en el orden viejo que en el nuevo—, y es
    precisamente lo que el ingeniero pidio: leer la hoja en el orden en que se
    rellena. Se fija por el ORDEN de aparicion de las bandas, no solo por su
    texto: una hoja que las numerase bien pero las dejara descolocadas pasaria
    una prueba de texto y seguiria siendo la hoja vieja.
    """

    ORDEN_212 = (
        (16, "1.  DATOS DE ENTRADA"),
        (37, "2.  RESOLUCIÓN DE MATERIAL"),
        (65, "3.  PARÁMETROS DE CÁLCULO"),
        (78, "4.  GEOMETRÍA Y PROPIEDADES DERIVADAS"),
        (87, "5.  CÁLCULO DE CARGAS Y SOLDADURA"),
        (99, "6.  RESULTADOS DEL DISEÑO"),
        (110, "7.  VERIFICACIONES"),
        (120, "8.  ESPECIFICACIONES TÉCNICAS"),
    )
    ORDEN_206 = (
        (16, "1. DATOS DE ENTRADA"),
        (35, "2. RESOLUCION DE MATERIAL"),
        (60, "3. PARAMETROS DE CALCULO"),
        (68, "4. GEOMETRIA DEL SLEEVE"),
        (74, "5. CALCULO DE ESPESOR REQUERIDO"),
        (83, "6. VERIFICACIONES Y AVISOS"),
    )

    @pytest.mark.parametrize("hoja,orden", [("Parche_PCC2_Art212", ORDEN_212),
                                            ("Collar_PCC2_Art206", ORDEN_206)])
    def test_las_bandas_van_numeradas_y_en_orden(self, wb, hoja, orden):
        ws = wb[hoja]
        for fila, prefijo in orden:
            texto = str(ws.cell(fila, 1).value or "")
            # El 206 escribe sus bandas con rotulo(): "[ MAYUSCULAS // ... ]".
            assert prefijo.upper() in texto.upper(), f"{hoja}!A{fila}: {texto!r}"
        filas = [f for f, _ in orden]
        assert filas == sorted(filas), "las bandas no van en orden creciente"

    def test_el_material_va_pegado_a_los_datos_de_entrada(self, wb):
        """Una sola fila en blanco entre la Seccion 1 y la de Material: si se
        colara un bloque en medio, dejaria de leerse de un tiron."""
        for hoja, fin_seccion1, banda_material in (("Parche_PCC2_Art212", 35, 37),
                                                   ("Collar_PCC2_Art206", 33, 35)):
            ws = wb[hoja]
            hueco = ws.cell(banda_material - 1, 1).value
            assert hueco is None, f"{hoja}!A{banda_material - 1}: {hueco!r}"
            assert ws.cell(fin_seccion1, 1).value is not None, hoja

    def test_ninguna_banda_conserva_el_numeral_viejo(self, wb):
        """La Seccion de Material era la '7' y estaba al final. Si ese rotulo
        sobreviviera en cualquier celda, la hoja se contradiria a si misma."""
        for hoja in ("Parche_PCC2_Art212", "Collar_PCC2_Art206"):
            ws = wb[hoja]
            malas = [c.coordinate for fila in ws.iter_rows(max_col=B.MOTOR_NCOLS)
                     for c in fila
                     if isinstance(c.value, str)
                     and ("7. RESOLUCION" in c.value.upper()
                          or "SECCION 7" in c.value.upper())]
            assert malas == [], f"{hoja}: {malas}"


class TestBuildCollarArt206:
    """Construye Collar_PCC2_Art206 en un wb de usar y tirar y fija por cadena
    las formulas nuevas de los pasos del flujo (el 206 no tiene oracle: se
    verifica por verificar.py §1-5 book-wide + §6f). Cada test_paso* traza a
    resources/ (Regla n.1) y al flujo aprobado del ingeniero."""

    def _construir(self):
        import build_db_materiales as B
        wb = openpyxl.Workbook()
        for n in ("DB_B31_3", "Datos_Ref", "MAP_Factores", "DB_BPVC_IID",
                  "DB_BPVC_IID_B", "DB_B36_10", "DB_B36_19"):
            wb.create_sheet(n)
        b36 = lambda sheet: {"sheet": sheet, "last_row": B.R_DATA + 9,
                             "n_nps": 5, "max_ced": 5}
        B.build_collar_art206(wb, b313_stub(), iid1a_stub(), iidb_stub(),
                              fac_stub(), rangos_stub(),
                              b36("DB_B36_10"), b36("DB_B36_19"))
        return wb["Collar_PCC2_Art206"]

    # --- Fase 1: seleccion guiada de tipo (Paso 1, 206-1 / 206-2) ------------
    def test_paso1_seleccion_tipo(self):
        import build_db_materiales as B
        ws = self._construir()
        assert str(ws["A101"].value).find("PASO 1") >= 0, ws["A101"].value
        assert ws["D103"].value == "No"  # fuga
        assert ws["D104"].value == "No"  # axial / tasa no clara
        assert ws["D105"].value == (
            f'=IF(OR($D$103="Si",$D$104="Si"),"{B.TIPO_B}","{B.TIPO_A}")'), "D105"
        assert ws["D106"].value == (
            '=IF($D$22<>$D$105,"AVISO: el tipo elegido (D22) difiere del '
            'recomendado por 206-1.1 (fuga/axial) — confirmar","")'), "D106"
        # Avisos de 206-2.
        assert ws["D107"].value == (
            f'=IF($D$22="{B.TIPO_A}","206-2.6: evaluar corrosion bajo manga; '
            'sellante/recubrimiento si aplica","")'), "D107"
        assert ws["D109"].value == (
            f'=IF(AND($D$22="{B.TIPO_B}",$D$103="Si"),"206-2.3: aislar la fuga '
            'antes de soldar (ver 206-4.3 purga N2)","")'), "D109"

    # --- Fase 2: t_req con sobreespesor de corrosion (Paso 2, 206-3.3) -------
    def test_paso2_treq_con_ca(self):
        ws = self._construir()
        assert "PASO 2" in str(ws["A111"].value), ws["A111"].value
        assert ws["D113"].value == 0, "C.A. default 0"
        # t_req Type B suma C.A. (D113) en las dos columnas (la Fase 2 retiro la
        # tercera; ver TestModeloDePresionDosCasos).
        for celda in ("D78", "E78"):
            assert str(ws[celda].value).endswith("+$D$113"), f"{celda}: {ws[celda].value}"
        # Type A (2/3 Tp) NO lleva C.A.
        assert "$D$113" not in str(ws["D79"].value), "D54 no debe llevar C.A."

    # --- Fase 2 (UX): el modelo de presion de dos casos, tambien aqui --------
    # 206-3.3 dimensiona contra "the maximum allowable design pressure"
    # (resources/.../art_206.json, Regla n.1), no contra un tipico intermedio.
    def test_modelo_de_presion_dos_casos(self):
        ws = self._construir()
        assert ws["D27"].value is None, "la fila del diseno tipico se retiro"
        assert ws["E76"].value == "=$D$28", ws["E76"].value
        for fila in (76, 77, 78):
            assert ws[f"F{fila}"].value is None, f"F{fila}: {ws[f'F{fila}'].value}"
        # El T_s,min gobernante de Type B pasa a leer el caso Diseno (E78).
        assert ws["D80"].value == f'=IF($D$22="{B.TIPO_A}",$D$79,$E$78)', ws["D80"].value

    # --- Fase 3: dimensiones del sleeve (Paso 3, 206-3.4) -------------------
    def test_paso3_dimensiones(self):
        ws = self._construir()
        assert "PASO 3" in str(ws["A116"].value), ws["A116"].value
        # F61 (longitud) ya existe y es correcto; el Paso 3 lo traza + nota 50 mm.
        assert ws["D86"].value == "=MAX(100,$D$30+2*50)", "D61 L_s,min"
        assert "50" in str(ws["D117"].value), ws["D117"].value

    # --- Fase 4: cateto del filete y luz radial (Paso 4, 206-3.5 / 206-4.1) --
    def test_paso4_filete_cateto(self):
        import build_db_materiales as B
        ws = self._construir()
        assert "PASO 4" in str(ws["A119"].value), ws["A119"].value
        assert ws["D121"].value == (
            f'=IF($D$22="{B.TIPO_B}",IF($D$29<=1.4*$D$21,$D$29+$D$31,'
            f'1.4*$D$21+$D$31),"No aplica - Type A (206-1.1.1)")'), "D121 cateto"
        # Verificacion de luz G <= 2.5 mm y su cableado a F69.
        assert ws["D123"].value == 2.5, "D123 req"
        assert ws["E123"].value == "=$D$31", "E123 adoptado"
        assert ws["F123"].value == (
            '=IF(E123<=D123,"CUMPLE","NO CUMPLE — excede 2,5 mm (206-4.1)")'), "F123"
        assert "F123" in str(ws["F94"].value), "F69 debe incluir F123"

    # --- Fase 5: presion externa, cavidades y bulging (206-3.6/7/9/10) ------
    def test_paso5_cavidades(self):
        ws = self._construir()
        assert "PASO 5" in str(ws["A125"].value), ws["A125"].value
        assert ws["D127"].value == "No"  # defecto externo
        assert ws["D128"].value == (
            '=IF($D$127="Si","206-3.7/3.9: rellenar cavidades con material '
            'endurecible (epoxi) de resistencia a compresion adecuada","")'), "D128"

    # --- Fase 6: fatiga por operacion ciclica (206-2.4/3.8/3.11) ------------
    def test_paso6_fatiga(self):
        ws = self._construir()
        assert "PASO 6" in str(ws["A132"].value), ws["A132"].value
        assert ws["D134"].value == "No"  # servicio ciclico
        assert ws["D135"].value == (
            '=IF($D$134="Si","206-2.4/3.8: requiere evaluacion de fatiga '
            '(VIII-2 / API 579-1/ASME FFS-1)","")'), "D135"

    # --- Fase 7: fabricacion y soldadura en servicio (206-4) ----------------
    def test_paso7_fabricacion(self):
        ws = self._construir()
        assert "PASO 7" in str(ws["A138"].value), ws["A138"].value
        assert ws["D140"].value == "=0.5*$D$26", "P_instal min 50%"
        assert ws["D141"].value == "=0.8*$D$26", "P_instal max 80%"
        assert ws["D142"].value == (
            '=IF($D$103="Si","206-4.3: purgar el anular con N2/gas inerte en '
            'fluidos inflamables","")'), "D142 purga N2"

    # --- Fase 8: NDE y prueba de hermeticidad (206-5 / 206-6) ---------------
    def test_paso8_nde_prueba(self):
        import build_db_materiales as B
        ws = self._construir()
        assert "PASO 8" in str(ws["A145"].value), ws["A145"].value
        origen = {}
        for dv in ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                origen[str(rng)] = str(dv.formula1 or "")
        assert "Prueba del anular" in origen.get("D147", ""), origen.get("D147")
        assert ws["D149"].value == (
            f'=IF($D$22="{B.TIPO_B}","206-5.3: UT del portador; primer/ultimo '
            'pase MT/PT; NDE de circunferenciales >=24 h (>=48 h si servicio '
            'con H2)","206-5.2: Type A - VT de raiz + PT/MT/UT de longitudinales")'), "D149"

    # --- Fase 9: simbolos con subindice real (decision 5) -------------------
    def test_simbolos_sin_guion_bajo_206(self):
        _asserta_sin_guion_bajo(self._construir())


class TestValidacionesBloqueantes:
    """Un desplegable que no rechaza lo que no esta en su lista es una sugerencia,
    no una restriccion — y el error de tecleo sigue siendo posible. Con
    showErrorMessage=False (como estaba hasta esta tarea) Excel aceptaba en
    silencio cualquier valor escrito a mano."""

    def test_toda_validacion_de_lista_bloquea(self, wb):
        flojas = []
        for ws in wb.worksheets:
            for dv in ws.data_validations.dataValidation:
                if dv.type != "list":
                    continue
                if not dv.showErrorMessage or dv.errorStyle != "stop":
                    flojas.append(f"{ws.title}!{sorted(str(r) for r in dv.sqref.ranges)[0]}")
        assert flojas == [], flojas


class TestBasesDimensionalesB36:
    HOJAS = ("DB_B36_10", "DB_B36_19")

    def test_existen_y_no_llevan_formulas(self, wb):
        for h in self.HOJAS:
            ws = wb[h]
            for fila in ws.iter_rows(min_row=1):
                for c in fila:
                    assert not (isinstance(c.value, str) and c.value.startswith("=")), \
                        f"{h}!{c.coordinate}: una base no lleva formulas"

    def test_orden_contiguo_por_nps(self, wb):
        # Regla 5: la cascada dependiente usa COUNTIF/INDEX/MATCH sobre bloques
        # contiguos. Si un NPS aparece en dos tramos separados, la lista de cedulas
        # de ese NPS sale truncada y nadie se entera.
        ws = wb["DB_B36_10"]
        col = B.COL_B36["nps_impreso"]
        vistos, ultimo = set(), None
        for r in range(B.R_DATA, ws.max_row + 1):
            v = ws.cell(r, col).value
            if v != ultimo:
                assert v not in vistos, f"NPS {v!r} reaparece fuera de su bloque"
                vistos.add(v)
                ultimo = v

    def test_caso_semilla_presente(self, wb):
        ws = wb["DB_B36_10"]
        cn, cc, ct = (B.COL_B36["nps_in"], B.COL_B36["cedula"], B.COL_B36["t_mm"])
        hit = [r for r in range(B.R_DATA, ws.max_row + 1)
               if ws.cell(r, cn).value == 12.0 and str(ws.cell(r, cc).value) == "20"]
        assert len(hit) == 1
        assert ws.cell(hit[0], ct).value == 6.35

    def test_nps_orden_marca_solo_la_primera_fila_del_bloque(self, wb):
        # nps_orden alimenta la lista de NPS sin repetir (Tarea 7): vale 1,2,3...
        # en la primera fila de cada bloque contiguo y vacio en el resto.
        ws = wb["DB_B36_10"]
        col_nps, col_orden = B.COL_B36["nps_impreso"], B.COL_B36["nps_orden"]
        vistos, marcados = set(), []
        for r in range(B.R_DATA, ws.max_row + 1):
            v = ws.cell(r, col_nps).value
            orden = ws.cell(r, col_orden).value
            if v not in vistos:
                vistos.add(v)
                assert orden is not None, f"fila {r} (primera de {v!r}) sin nps_orden"
                marcados.append(orden)
            else:
                assert orden is None, f"fila {r} ({v!r} repetido) no deberia llevar nps_orden"
        assert marcados == list(range(1, len(marcados) + 1))

    def test_clave_combina_nps_impreso_y_designador(self, wb):
        ws = wb["DB_B36_10"]
        col_nps, col_des, col_clave = (B.COL_B36["nps_impreso"], B.COL_B36["designador"],
                                       B.COL_B36["clave"])
        for r in (B.R_DATA, B.R_DATA + 1):
            esperada = f"{ws.cell(r, col_nps).value}|{ws.cell(r, col_des).value}"
            assert ws.cell(r, col_clave).value == esperada

    def test_clave_ced_numera_designadores_no_vacios_por_nps(self, wb):
        # clave_ced = nps_impreso|<indice 1,2,3... del designador NO vacio dentro
        # del bloque de NPS>, y vacia (None al recargar) en las filas '...'. Es lo
        # que deja listar las cedulas de un NPS sin ofrecer opciones en blanco.
        for h in self.HOJAS:
            ws = wb[h]
            col_nps, col_des, col_cc = (B.COL_B36["nps_impreso"],
                                        B.COL_B36["designador"], B.COL_B36["clave_ced"])
            nps_actual, idx = None, 0
            for r in range(B.R_DATA, ws.max_row + 1):
                nps = ws.cell(r, col_nps).value
                des = ws.cell(r, col_des).value
                cc = ws.cell(r, col_cc).value
                if nps != nps_actual:
                    nps_actual, idx = nps, 0
                if des:
                    idx += 1
                    assert cc == f"{nps}|{idx}", f"{h}!fila {r}: {cc!r}"
                else:
                    assert cc is None, f"{h}!fila {r}: designador vacio con clave_ced {cc!r}"


class TestDatosRefRetirada:
    """Cierre de la Tarea 10: la hoja heredada Datos_Ref sale del libro. Su dato
    vivo (esfuerzos admisibles y dimensiones) esta en las bases auditadas
    (DB_B31_3, DB_BPVC_IID, DB_B36_10/19); el libro anterior queda en git."""

    def test_la_hoja_no_existe(self, wb):
        assert "Datos_Ref" not in wb.sheetnames

    def test_ninguna_formula_la_menciona(self, wb):
        restos = []
        for ws in wb.worksheets:
            for fila in ws.iter_rows():
                for c in fila:
                    if isinstance(c.value, str) and "Datos_Ref" in c.value:
                        restos.append(f"{ws.title}!{c.coordinate}")
        assert restos == [], restos

    def test_solo_queda_una_hoja_heredada(self):
        import build_db_materiales as B
        assert B.HOJAS_HEREDADAS == ("Instrucciones",)

    def test_la_deuda_de_listas_fijas_esta_saldada(self):
        import build_db_materiales as B
        assert B.DEUDA_LISTA_FIJA == {}
