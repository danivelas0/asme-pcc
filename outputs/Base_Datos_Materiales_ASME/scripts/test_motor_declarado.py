# test_motor_declarado.py
import pytest
import motor_declarado as M


def _motor_minimo():
    """Dos filas en una seccion: una entrada y un calculo que la usa."""
    return M.Motor(
        articulo="999",
        hoja="Prueba_PCC2_Art999",
        titulo="MOTOR DE PRUEBA",
        fuente="ASME PCC/pcc_2/p2_welded_repairs/art_999/art_999.json",
        secciones=(
            M.Seccion("1. DATOS DE ENTRADA", filas=(
                M.Fila("P", "Presion de diseno", magnitud="pres", tipo=M.ENTRADA,
                       ejemplo=20,
                       cita=M.Cita("art_999.json", bloque=12, clausula="999-3.2")),
                M.Fila("dos_P", "El doble", magnitud="pres", tipo=M.FORMULA,
                       formula="={P}*2",
                       cita=M.Cita("art_999.json", bloque=12, clausula="999-3.2")),
            )),
        ),
    )


def test_cada_fila_recibe_una_direccion_en_orden():
    dirs = M.resolver_direcciones(_motor_minimo())
    assert dirs["P"] == "$D$6"
    assert dirs["dos_P"] == "$D$7"


def test_la_formula_se_escribe_por_nombre_y_sale_por_direccion():
    dirs = M.resolver_direcciones(_motor_minimo())
    assert M.sustituir_nombres("={P}*2", dirs) == "=$D$6*2"


def test_un_nombre_que_no_existe_aborta():
    dirs = M.resolver_direcciones(_motor_minimo())
    with pytest.raises(SystemExit, match="no existe"):
        M.sustituir_nombres("={NO_EXISTE}+1", dirs)


def test_una_fila_de_calculo_sin_cita_aborta(tmp_path):
    motor = _motor_minimo()._replace(secciones=(
        M.Seccion("1. DATOS", filas=(
            M.Fila("x", "Sin procedencia", tipo=M.FORMULA, formula="=1+1"),
        )),))
    with pytest.raises(SystemExit, match="sin procedencia"):
        M.comprobar_procedencia(motor, tmp_path)


def test_una_cita_a_un_bloque_que_no_existe_aborta(tmp_path):
    import json
    d = tmp_path / "ASME PCC" / "pcc_2" / "p2_welded_repairs" / "art_999"
    d.mkdir(parents=True)
    (d / "art_999.json").write_text(
        json.dumps({"blocks": [{"type": "paragraph", "text": "uno"}]}),
        encoding="utf-8")
    motor = _motor_minimo()   # su cita apunta al bloque 12, que no existe
    with pytest.raises(SystemExit, match="bloque 12"):
        M.comprobar_procedencia(motor, tmp_path)


def _helpers_de_prueba():
    from openpyxl.styles import Protection

    def banda(ws, fila, texto):
        ws.cell(fila, 1, texto)

    def rotulo(ws, fila, texto, simbolo, unidad, referencia="", comentario=""):
        ws.cell(fila, 1, texto)
        if referencia:
            ws.cell(fila, 7, referencia)

    def entrada(ws, celda, valor):
        ws[celda] = valor
        ws[celda].protection = Protection(locked=False)

    def calculo(ws, celda, formula):
        ws[celda] = formula

    return M.Helpers(banda=banda, rotulo=rotulo, entrada=entrada, calculo=calculo)


def test_la_hoja_sale_con_las_bandas_y_las_filas_en_orden():
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    res = M.emitir_tabla(ws, _motor_minimo(), _helpers_de_prueba())
    assert ws["A5"].value == "1. DATOS DE ENTRADA"
    assert ws["A6"].value == "Presion de diseno"
    assert ws["D6"].value == 20                 # el ejemplo del caso precargado
    assert ws["A7"].value == "El doble"
    assert ws["D7"].value == "=$D$6*2"          # la formula, ya con direcciones
    assert res["ultima_fila"] == 7


def test_la_clausula_llega_a_la_columna_de_referencia():
    """La procedencia no se queda en la tabla de trazabilidad: el ingeniero
    tiene que ver de que clausula sale el numero SIN salir de la fila."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    M.emitir_tabla(ws, _motor_minimo(), _helpers_de_prueba())
    assert ws["G6"].value == "999-3.2"


def test_la_entrada_nace_desbloqueada_y_el_calculo_no():
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    M.emitir_tabla(ws, _motor_minimo(), _helpers_de_prueba())
    assert ws["D6"].protection.locked is False   # entrada
    assert ws["D7"].protection.locked is not False


def _motor_dos_secciones():
    """Dos secciones con varias filas cada una -no una sola, como en
    `_motor_minimo`-, para blindar que `resolver_direcciones()` y
    `emitir_tabla()` recorren el MISMO arbol. Hoy coinciden por casualidad:
    ninguna prueba anterior lo comprueba con mas de una seccion, que es
    exactamente el patron de "dos fuentes de verdad" que este proyecto marca
    como peligroso en otras partes (ver la capa de navegacion en el CLAUDE.md
    del repo)."""
    cita = M.Cita("art_999.json", bloque=12, clausula="999-3.2")
    return M.Motor(
        articulo="999",
        hoja="Prueba_PCC2_Art999",
        titulo="MOTOR DE PRUEBA",
        fuente="ASME PCC/pcc_2/p2_welded_repairs/art_999/art_999.json",
        secciones=(
            M.Seccion("1. DATOS DE ENTRADA", filas=(
                M.Fila("P", "Presion", tipo=M.ENTRADA, ejemplo=20),
                M.Fila("D", "Diametro", tipo=M.ENTRADA, ejemplo=300),
                M.Fila("dosP", "El doble de P", tipo=M.FORMULA,
                       formula="={P}*2", cita=cita),
            )),
            M.Seccion("2. RESULTADOS", filas=(
                M.Fila("suma", "P + D", tipo=M.FORMULA,
                       formula="={P}+{D}", cita=cita),
                M.Fila("otra", "suma + P", tipo=M.FORMULA,
                       formula="={suma}+{P}", cita=cita),
            )),
        ),
    )


def test_emitir_tabla_escribe_en_la_direccion_que_resolver_direcciones_devuelve():
    """Con dos secciones y varias filas: la direccion que `emitir_tabla` USA
    para escribir cada fila tiene que ser exactamente la que
    `resolver_direcciones` DEVUELVE para esa clave -incluida una formula de la
    segunda seccion que cita una clave de la primera-. Sin esta prueba, las
    dos funciones podian divergir en cuanto un motor real tuviera mas de una
    seccion y nadie lo notaria hasta ver un #REF! en una celda."""
    import openpyxl
    motor = _motor_dos_secciones()
    dirs = M.resolver_direcciones(motor)                      # fuente nº 1

    wb = openpyxl.Workbook()
    ws = wb.active
    M.emitir_tabla(ws, motor, _helpers_de_prueba())            # fuente nº 2

    # Las direcciones en si: la segunda seccion sigue contando filas donde la
    # primera las dejo, y no reinicia en su propia banda.
    assert dirs == {"P": "$D$6", "D": "$D$7", "dosP": "$D$8",
                    "suma": "$D$10", "otra": "$D$11"}

    esperado = {"P": 20, "D": 300, "dosP": "=$D$6*2",
                "suma": "=$D$6+$D$7", "otra": "=$D$10+$D$6"}
    for clave, valor in esperado.items():
        direccion = dirs[clave]
        assert ws[direccion].value == valor, (
            f"{clave}: resolver_direcciones dice {direccion}, pero ese no es "
            f"el valor que emitir_tabla escribio ahi")


def _motor_con_pasos():
    """Una seccion normal seguida del anexo de pasos: dos Paso, cada uno con
    sus propias filas -incluida una formula que cita una clave de la seccion
    y otra que cita una clave de un paso anterior-, para blindar que
    resolver_direcciones() y emitir_tabla() tambien coinciden cuando el
    arbol incluye motor.pasos (Tarea 6 de la skill motor_pcc2)."""
    cita = M.Cita("art_999.json", bloque=12, clausula="999-3.2")
    return M.Motor(
        articulo="999",
        hoja="Prueba_PCC2_Art999",
        titulo="MOTOR DE PRUEBA",
        fuente="ASME PCC/pcc_2/p2_welded_repairs/art_999/art_999.json",
        secciones=(
            M.Seccion("1. DATOS DE ENTRADA", filas=(
                M.Fila("P", "Presion", tipo=M.ENTRADA, ejemplo=20),
            )),
        ),
        pasos=(
            M.Paso(1, "ELEGIBILIDAD", "999-1", filas=(
                M.Fila("elegible", "Compuerta", tipo=M.FORMULA,
                       formula="={P}*0", cita=cita),
            )),
            M.Paso(2, "CARGAS", "999-2", filas=(
                M.Fila("carga", "Carga total", tipo=M.FORMULA,
                       formula="={P}+{elegible}", cita=cita),
            )),
        ),
    )


def test_emitir_tabla_escribe_en_la_direccion_que_resolver_direcciones_devuelve_con_pasos():
    """Mismo blindaje que la prueba anterior, pero con motor.pasos: el anexo
    de pasos tiene que salir DETRAS de la ultima seccion, con una banda por
    paso, y las direcciones que asigna resolver_direcciones() tienen que ser
    las mismas donde emitir_tabla() escribio de verdad -incluida la formula
    del Paso 2 que cita una clave declarada en el Paso 1."""
    import openpyxl
    motor = _motor_con_pasos()
    dirs = M.resolver_direcciones(motor)                       # fuente n.1

    wb = openpyxl.Workbook()
    ws = wb.active
    M.emitir_tabla(ws, motor, _helpers_de_prueba())             # fuente n.2

    # Fila 5: banda de la seccion. Fila 6: "P". Fila 7: banda del Paso 1.
    # Fila 8: "elegible". Fila 9: banda del Paso 2. Fila 10: "carga".
    assert dirs == {"P": "$D$6", "elegible": "$D$8", "carga": "$D$10"}
    assert ws["A7"].value == "PASO 1 . ELEGIBILIDAD  --  999-1"
    assert ws["A9"].value == "PASO 2 . CARGAS  --  999-2"

    esperado = {"P": 20, "elegible": "=$D$6*0", "carga": "=$D$6+$D$8"}
    for clave, valor in esperado.items():
        direccion = dirs[clave]
        assert ws[direccion].value == valor, (
            f"{clave}: resolver_direcciones dice {direccion}, pero ese no es "
            f"el valor que emitir_tabla escribio ahi")
