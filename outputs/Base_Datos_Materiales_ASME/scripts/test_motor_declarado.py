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
