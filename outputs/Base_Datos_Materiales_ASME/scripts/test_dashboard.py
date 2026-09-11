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
          B.VERDE, B.AMBAR, B.AMBAR_TXT}
FUENTES_DEL_SISTEMA = {B.MONO, B.MACRO}


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


class TestParidadHojaParche:
    HOJA = "Parche_PCC2_Art212"

    def test_todas_las_formulas_y_literales(self, wb):
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        for celda, esperado in oracle["formulas"].items():
            motivo = B.DIVERGENCIAS_DECLARADAS.get((self.HOJA, celda))
            if motivo:
                # Celda que el build ya no reproduce a proposito. Se exige que
                # este VACIA: una divergencia declarada que resulta traer otro
                # valor es un error, no una divergencia.
                assert ws[celda].value is None, f"{celda}: {motivo}"
                continue
            assert ws[celda].value == esperado, celda

    def test_toda_divergencia_declara_motivo(self):
        for (hoja, celda), motivo in B.DIVERGENCIAS_DECLARADAS.items():
            assert isinstance(motivo, str) and len(motivo) > 20, (hoja, celda)

    def test_validaciones_de_datos(self, wb):
        # Igual que las formulas: una validacion cuyo sqref cae ENTERO dentro
        # de celdas declaradas en DIVERGENCIAS_DECLARADAS ya no la reproduce
        # el build (la fila entera se retiro) y se excluye del oracle. Una
        # validacion parcialmente divergente seguiria exigiendose completa:
        # aqui no se da el caso, pero el guardia es honesto sobre ello.
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        reales = sorted(
            ({"sqref": str(dv.sqref), "formula1": dv.formula1}
             for dv in ws.data_validations.dataValidation),
            key=lambda d: d["sqref"])
        esperadas = []
        for v in oracle["validaciones"]:
            rng = openpyxl.worksheet.cell_range.CellRange(v["sqref"])
            coords = [f"{openpyxl.utils.get_column_letter(c)}{r}"
                      for r in range(rng.min_row, rng.max_row + 1)
                      for c in range(rng.min_col, rng.max_col + 1)]
            if all((self.HOJA, c) in B.DIVERGENCIAS_DECLARADAS for c in coords):
                continue
            esperadas.append(v)
        assert reales == sorted(esperadas, key=lambda d: d["sqref"])

    def test_rangos_fusionados(self, wb):
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        assert sorted(str(r) for r in ws.merged_cells.ranges) == oracle["fusionados"]


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
                  "DB_BPVC_IID_B"):
            wb.create_sheet(n)
        # Sembrar en DB_B31_3 los dos material_id del caso precargado para que
        # el guardia de corregir_art212_fase1 (ahora en build_parche_art212) pase.
        db = wb["DB_B31_3"]
        db.cell(B.R_DATA, 1, B.SEED_ART212_BASE)
        db.cell(B.R_DATA + 1, 1, B.SEED_ART212_COLLAR)
        B.build_parche_art212(wb, b313_stub(), iid1a_stub(), iidb_stub(),
                              fac_stub(), rangos_stub())
        return wb["Parche_PCC2_Art212"]

    ANCLAS_T2 = ("A1", "A2", "D5", "D6", "D115", "E115", "D126", "E126",
                 "D39", "D40", "D44", "F90", "D114", "E114", "D128")

    def test_anclas_de_la_tarea_2(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_T2:
            assert ws[celda].value == oracle["formulas"][celda], celda

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
            assert ws[celda].value == oracle["formulas"][celda], celda

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
        "A27", "B27", "C27", "D27", "G27",
        "A28", "B28", "C28", "D28", "G28",
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
            assert ws[celda].value == oracle["formulas"][celda], celda

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
        "A37",
        "A38", "B38", "C38", "D38", "G38",
        "A39", "B39", "C39", "D39", "G39",
        "A40", "B40", "C40", "D40", "G40",
        "A41", "B41", "C41", "D41", "G41",
        "A42", "B42", "C42", "D42", "G42",
        "A43", "B43", "C43", "D43", "G43",
        "A44", "B44", "C44", "D44", "G44",
        "A45", "B45", "C45", "D45",
        "A46", "B46", "C46", "D46", "G46",
        "A47", "B47", "C47", "D47", "G47",
        "A48", "B48", "C48", "D48", "G48",
    )

    def test_seccion_5(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_5:
            assert ws[celda].value == oracle["formulas"][celda], celda

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
        "A50",
        "A51", "B51", "C51", "D51", "G51",
        "A52", "B52", "C52", "D52", "G52",
        "A53", "B53", "C53", "D53", "G53",
        "A54", "B54", "C54", "D54", "G54",
        "A55", "B55", "C55", "D55", "G55",
        "A56", "B56", "C56", "D56", "G56",
        "A57", "B57", "C57", "D57", "G57",
    )

    def test_seccion_6(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_6:
            assert ws[celda].value == oracle["formulas"][celda], celda

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
        "A59",
        "A60", "B60", "C60", "D60", "E60", "F60", "G60",
        "A61", "B61", "C61", "D61", "E61", "F61", "G61",
        "A62", "B62", "C62", "D62", "E62", "F62", "G62",
        "A63", "B63", "C63", "D63", "E63", "F63", "G63",
        "A64", "B64", "C64", "D64", "E64", "F64", "G64",
        "A65", "B65", "C65", "D65", "E65", "F65", "G65",
        "A66", "B66", "C66", "D66", "E66", "F66", "G66",
        "A67", "B67", "C67", "D67", "E67", "F67", "G67",
        "A68", "B68", "C68", "D68", "E68", "F68", "G68",
        "A69", "C69", "D69", "E69", "F69", "G69",
        "A71",
        "A72", "B72", "C72", "D72", "G72",
        "A73", "B73", "C73", "D73", "G73",
        "A74", "B74", "C74", "D74", "G74",
        "A75", "B75", "C75", "D75",
        "A76", "C76", "D76", "G76",
        "A77", "C77", "D77", "G77",
        "A78", "B78", "C78", "D78", "G78",
        "A79", "C79", "D79", "G79",
        "A80", "C80", "D80", "G80",
    )

    def test_filas_61_80(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_FILAS_61_80:
            assert ws[celda].value == oracle["formulas"][celda], celda

    # Cubre la Seccion 5 (verificaciones, filas 82-90), la Seccion 6
    # (especificaciones tecnicas, filas 93-99), el aviso fijo (fila 101) y
    # DOS gaps que quedaron sin asignar en secciones anteriores:
    #   (a) la banda IDENTIFICACION (A4) + G5/G6;
    #   (b) la banda APLICACION Y CODIGO (A8) + el encabezado de fila 9.
    # Incluye tambien la banda A82 y el encabezado de fila 83 (preceden el
    # rango 84-90 y titulan esta seccion, no una anterior) y la banda A92
    # (precede el rango 93-99; esta seccion no tiene fila de encabezado
    # propia — el *oracle* no la declara). F90 ya lo cubre
    # test_anclas_de_la_tarea_2; se repite aqui para que la seccion quede
    # completa por si sola.
    ANCLAS_SECCION_8 = (
        "A4", "A5", "D5", "G5", "A6", "D6", "G6",
        "A8",
        "A9", "B9", "C9", "D9", "G9",
        "A82",
        "A83", "D83", "E83", "F83", "G83",
        "A84", "D84", "E84", "F84", "G84",
        "A85", "D85", "E85", "F85", "G85",
        "A86", "D86", "E86", "F86", "G86",
        "A87", "D87", "E87", "F87", "G87",
        "A88", "D88", "E88", "F88", "G88",
        "A89", "D89", "E89", "F89", "G89",
        "A90", "F90",
        "A92",
        "A93", "B93",
        "A94", "B94",
        "A95", "B95",
        "A96", "B96",
        "A97", "B97",
        "A98", "B98",
        "A99", "B99",
        "A101",
    )

    def test_seccion_8(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_SECCION_8:
            assert ws[celda].value == oracle["formulas"][celda], celda


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
