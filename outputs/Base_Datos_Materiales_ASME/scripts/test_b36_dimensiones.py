"""Pruebas del parser de B36.10M / B36.19M. Las filas de muestra estan copiadas
LITERALES de la extraccion (pagina 13 de B36.10M): no se inventan valores."""
import b36_dimensiones as B


def test_parte_celda_de_doble_unidad():
    assert B.partir_doble_unidad("0.405 (10.29)") == (0.405, 10.29)
    assert B.partir_doble_unidad("0.068 (1.73)") == (0.068, 1.73)


def test_doble_unidad_ausente_no_inventa():
    assert B.partir_doble_unidad("...") == (None, None)
    assert B.partir_doble_unidad("") == (None, None)


def test_doble_unidad_con_nota_al_pie_no_se_descarta():
    # B36.19M pega la referencia de nota al pie tras el parentesis. El valor
    # sigue siendo el que imprime el codigo; solo se descarta la referencia.
    assert B.partir_doble_unidad("0.049 (1.24) [Note (1)]") == (0.049, 1.24)
    assert B.partir_doble_unidad("0.188 (4.78) [Notes (1), (2)]") == (0.188, 4.78)


def test_parte_el_nps_fraccionario():
    assert B.partir_nps("1/8 (6)") == ("1/8 (6)", 0.125, 6.0)
    assert B.partir_nps("12 (300)") == ("12 (300)", 12.0, 300.0)
    assert B.partir_nps("2 1/2 (65)") == ("2 1/2 (65)", 2.5, 65.0)


def test_designador_combina_las_dos_designaciones():
    # El codigo publica una tubería por numero de cedula, por identificacion, por
    # las dos, o por una sola. El designador las junta sin perder ninguna.
    assert B.designador("STD", "40") == "40 (STD)"
    assert B.designador("XS", "80") == "80 (XS)"
    assert B.designador("...", "10") == "10"
    assert B.designador("XXS", "...") == "XXS"


def test_no_convierte_el_marcador_del_codigo_a_vacio():
    # '...' es 'no aplica' IMPRESO por la norma (regla 9). La fila lo conserva.
    filas = B.cargar(B.RUTA_B3610)
    xxs = [f for f in filas if f["nps_impreso"] == "1/8 (6)" and f["identificacion"] == "XXS"]
    assert len(xxs) == 1
    assert xxs[0]["cedula"] == "..."
    assert xxs[0]["designador"] == "XXS"


def test_caso_semilla_del_art_212():
    # NPS 12, Sch 20 -> espesor 6.35 mm. Es el caso precargado del motor Art. 212:
    # si esta fila no resuelve, el motor deja de reproducir su propio caso.
    filas = B.cargar(B.RUTA_B3610)
    f = [x for x in filas if x["nps_in"] == 12.0 and x["cedula"] == "20"]
    assert len(f) == 1
    assert f[0]["t_mm"] == 6.35
    assert f[0]["od_mm"] == 323.8


def test_b3619_no_trae_columna_de_identificacion():
    filas = B.cargar(B.RUTA_B3619)
    assert filas, "la extraccion de B36.19M no produjo filas"
    assert all(f["identificacion"] == "" for f in filas)
    assert any(f["cedula"].endswith("S") for f in filas), "faltan las cedulas 5S/10S/40S/80S"
