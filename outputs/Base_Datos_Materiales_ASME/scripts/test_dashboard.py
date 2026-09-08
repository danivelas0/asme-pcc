"""Pruebas de la capa de navegacion (Dashboard + VBA) del entregable Rev. 4.

A diferencia de test_build_db.py, que solo ejerce funciones puras, estas
pruebas abren el .xlsm ya construido. Se saltan si no existe, para que el
paquete de pruebas siga corriendo en una copia limpia del repo.

    python build_db_materiales.py --resources ... --in ... --out ...\\Rev4.xlsm
    set MOTOR_XLSM=..\\..\\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
    python -m pytest test_dashboard.py -q
"""
from __future__ import annotations

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
        # Parche_PCC2_Art212 e Instrucciones se entregan protegidas.
        for nombre in ("Parche_PCC2_Art212", "Instrucciones"):
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
# El proyecto VBA sobrevive al round-trip de openpyxl
# ---------------------------------------------------------------------------
class TestVba:
    def test_el_binario_vba_sigue_en_el_paquete(self, wb):
        assert "xl/vbaProject.bin" in zipfile.ZipFile(RUTA).namelist()

    def test_la_extension_es_xlsm(self):
        assert RUTA.suffix == ".xlsm"


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
