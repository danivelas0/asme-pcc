# -*- coding: utf-8 -*-
"""Pruebas de secii_tablas — reconstruccion de las tablas de la Seccion II.

Mitad de las aserciones van contra los JSON reales de `resources/`: son
regresiones ancladas a filas concretas del codigo, no a un fixture inventado.
Cada una detiene un modo de fallo distinto, y el modo va anotado al lado.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import secii_tablas as S

RES = Path(__file__).resolve().parents[3] / "resources"
HAY_RESOURCES = (RES / S.SEC_II).is_dir()
saltar = pytest.mark.skipif(not HAY_RESOURCES,
                            reason="no estan las partes A/B/C de la Seccion II")


def tablas(parte, spec):
    ruta = next(S.archivos_de(RES, parte, [spec]))
    return S.tablas_de(S.cargar_spec(ruta))


# ---------------------------------------------------------------------------
# Texto: lo que separa y lo que no
# ---------------------------------------------------------------------------
class TestTexto:
    def test_cada_etiqueta_separa(self):
        # Quitar las etiquetas a lo bruto pegaria las palabras del Caption de
        # SB-111, que llega con cada palabra en su propio <b>.
        assert S.texto("<b>TABLE</b><b>1</b><b>Chemical</b><b>Requirements</b>") == \
            "TABLE 1 Chemical Requirements"

    def test_el_superindice_va_pegado(self):
        # El marcador de nota se imprime pegado al valor: '0.25A' es UNA celda.
        # Si se separa, la fila pasa de 3 fichas de valor a 6 y deja de cuadrar
        # con las columnas de la tabla.
        assert S.texto("Carbon, max 0.25<i><sup>A</sup></i> 0.30<i><sup>B</sup></i>") \
            == "Carbon, max 0.25A 0.30B"
        assert S.texto("Chromium, max<i><sup>C</sup></i> 0.40") == "Chromium, maxC 0.40"


class TestFichas:
    def test_el_espacio_no_siempre_separa_celdas(self):
        # '48 000' es un separador de millares y '[330]' la unidad SI que el
        # propio codigo imprime junto al valor: las tres son UNA ficha.
        assert S.fichas("Tensile strength, min, psi [MPa] 48 000 [330]")[-1] == \
            "48 000 [330]"
        assert S.fichas("a 1 1⁄2 b") == ["a", "1 1⁄2", "b"]

    def test_reconoce_los_valores_que_imprime_el_codigo(self):
        for v in ("0.035", "0.27–0.93", "0.25A", "48 000 [330]", "1.00", ". . ."):
            assert S.es_valor(v), v
        for t in ("Manganese", "Carbon,", "max", "Grade"):
            assert not S.es_valor(t), t

    def test_cuenta_las_fichas_de_valor_del_final(self):
        assert S.n_valores_finales("Manganese 0.27–0.93 0.29–1.06 0.29–1.06") == 3
        assert S.n_valores_finales("Grade A Grade B Grade C") == 0


# ---------------------------------------------------------------------------
# Filas y columnas
# ---------------------------------------------------------------------------
def _l(x0, y0, x1, y1, t="x"):
    return dict(id=f"L{x0}-{y0}", html=t, bbox=[x0, y0, x1, y1],
                pagina="1", folio="1", tabla="T")


class TestAgrupamiento:
    def test_dos_filas_consecutivas_no_se_funden(self):
        # Es el fallo que colapsaba una tabla entera en una fila: entre el pie
        # de una fila y la cabeza de la siguiente hay solo 2 pt, dentro de la
        # tolerancia si esta se mide sobre el borde. Sobre el CENTRO no.
        filas = S.agrupar_filas([_l(0, 562.3, 100, 569.3), _l(0, 571.3, 100, 577.5)])
        assert len(filas) == 2

    def test_un_superindice_no_parte_la_celda(self):
        # SB-111: '99.99' y 'min^A' son contiguos en x y la caja del segundo
        # sube ~1 pt por el superindice. Son la misma fila.
        filas = S.agrupar_filas([_l(0, 100.0, 20, 106.0), _l(21, 99.0, 40, 105.5)])
        assert len(filas) == 1 and len(filas[0]) == 2

    def test_la_banda_ancha_no_tapa_los_huecos_entre_columnas(self):
        # Un rotulo que cruza varias columnas, si entra en la proyeccion en x,
        # fusiona la tabla entera en una sola banda.
        filas = [[_l(0, 10, 300, 16, "rotulo que cruza todo"),
                  _l(0, 10, 40, 16)],
                 [_l(0, 20, 40, 26), _l(100, 20, 140, 26), _l(200, 20, 240, 26)]]
        assert len(S.bandas_x(filas)) >= 3

    def test_la_banda_se_parte_cuando_una_fila_la_desmiente(self):
        # Una celda de la segunda fila cruza el hueco que separa las dos
        # primeras columnas, y la proyeccion en x las fusiona. Que sean dos
        # columnas lo demuestra la PRIMERA fila, que tiene un Line a cada lado
        # del hueco: la banda se parte por ahi.
        filas = [[_l(0, 10, 60, 16), _l(100, 10, 160, 16), _l(300, 10, 360, 16)],
                 [_l(50, 20, 110, 26), _l(300, 20, 360, 26)]]
        antes = S.bandas_x(filas)
        assert len(antes) == 2                     # las dos primeras, fusionadas
        despues = S.refinar_bandas(antes, filas)
        assert len(despues) == 3
        assert despues[0][1] == despues[1][0] == 80.0   # corte en medio del hueco


class TestRepartoPorConteo:
    def test_parte_desde_la_derecha(self):
        assert S.repartir_por_conteo("Manganese 0.27–0.93 0.29–1.06 0.29–1.06", 4) == \
            ["Manganese", "0.27–0.93", "0.29–1.06", "0.29–1.06"]

    def test_no_parte_si_el_conteo_no_cuadra(self):
        # Nunca se fuerza: si las fichas de valor no son las que pide la tabla,
        # la fila se queda entera y se declara AMBIGUA.
        assert S.repartir_por_conteo("Grade A Grade B Grade C", 4) is None
        assert S.repartir_por_conteo("Manganese 0.27–0.93", 4) is None


# ---------------------------------------------------------------------------
# Regresiones ancladas al codigo
# ---------------------------------------------------------------------------
@saltar
class TestContraElCodigo:
    def test_sa_106_tabla_1_se_reparte_entera(self):
        # Los dos modos de fila conviven en esta tabla: la cabecera llega en un
        # Line que no se deja partir y las once filas de datos si.
        t = tablas("bpvc_ii_a_1", "SA-106")[0]
        assert t["titulo"] == "TABLE 1 Chemical Requirements"
        assert t["ncols"] == 4
        por = {f["celdas"][0]: f for f in t["filas"] if len(f["celdas"]) == 4}
        assert por["Manganese"]["celdas"] == \
            ["Manganese", "0.27–0.93", "0.29–1.06", "0.29–1.06"]
        # Los marcadores de nota siguen pegados a su valor.
        assert por["Carbon, max"]["celdas"][1:] == ["0.25A", "0.30B", "0.35B"]
        # Y el rotulo conserva el suyo.
        assert "Chromium, maxC" in por
        assert not S.verificar_sin_perdida(t["filas"])

    def test_sb_111_el_caption_no_pega_las_palabras(self):
        ts = tablas("bpvc_ii_b", "SB-111")
        assert any(t["titulo"] == "TABLE 1 Chemical Requirements" for t in ts)

    def test_sb_111_la_celda_partida_en_dos_line_no_se_pierde(self):
        # '99.99' y 'min^A' son dos Line de la misma celda. Se exige que el
        # texto siga entero en una sola celda, no repartido entre dos columnas.
        ts = tablas("bpvc_ii_b", "SB-111")
        celdas = [c for t in ts for f in t["filas"] for c in f["celdas"]]
        assert any("99.99 minA" in c for c in celdas)

    def test_sfa_5_9_une_las_cuatro_paginas_de_su_tabla_1(self):
        # Es el caso de continuacion de la parte C: `Table 1 (Continued)`. Si no
        # se unen, la cabecera —que solo se imprime en la primera pagina— no
        # llega a las otras tres.
        ts = tablas("bpvc_ii_c", "SFA-5.9")
        t = ts[0]
        assert t["titulo"].startswith("Table 1 Chemical Composition")
        assert [int(p) for p in t["paginas"]] == [356, 357, 358, 359]
        assert t["continuada"]

    def test_sfa_5_9_las_notas_de_la_pagina_359_aparecen_una_vez(self):
        # Marker emite pares Text/Footnote con el mismo bbox y el contenido
        # ALTERNA entre uno y otro: sin deduplicar se pierde la mitad o se
        # duplica la otra mitad.
        t = tablas("bpvc_ii_c", "SFA-5.9")[0]
        textos = [n["texto"] for n in t["notas"]]
        assert len(textos) == len(set(textos))
        assert any("Single values shown are maximum percentages" in x
                   for x in textos)
        assert any("Up to 20% of the amount of Nb can be replaced by Ta" in x
                   for x in textos)

    def test_sa_6_titula_sus_tablas_sueltas(self):
        # El 32 % de las tablas del drop va SUELTA, sin TableGroup ni Caption
        # hermano: el titulo esta en el bloque de encima.
        ts = tablas("bpvc_ii_a_1", "SA-6")
        sueltas = [t for t in ts if t["suelta"]]
        assert len(sueltas) >= 20
        assert any(t["titulo"].startswith("TABLE 2 Permitted Variations in Weight")
                   for t in ts if any(int(p) == 71 for p in t["paginas"]))

    def test_sin_perdida_en_una_muestra_de_las_cuatro_partes(self):
        # La garantia de toda esta capa: reparte el texto impreso y no anade ni
        # quita un caracter. Se ejerce sobre mas de 50 tablas de las 4 partes.
        n = 0
        for parte, specs in (("bpvc_ii_a_1", ["SA-106", "SA-6", "SA-240"]),
                             ("bpvc_ii_a_2", ["SA-479", "SA-516"]),
                             ("bpvc_ii_b", ["SB-111", "SB-265"]),
                             ("bpvc_ii_c", ["SFA-5.9", "SFA-5.5"])):
            for spec in specs:
                try:
                    ts = tablas(parte, spec)
                except StopIteration:
                    continue
                for t in ts:
                    n += 1
                    assert not S.verificar_sin_perdida(t["filas"]), \
                        f"{parte}/{spec} tabla {t['n']}"
        assert n >= 50

    def test_los_huecos_de_la_parte_c_siguen_declarados(self):
        # 101 bloques `Form` con html vacio y sin hijos. No se rellenan: se
        # declaran. Si un dia dejan de contarse, es que alguien los relleno.
        total = 0
        for ruta in S.archivos_de(RES, "bpvc_ii_c"):
            total += S.huecos_de(S.cargar_spec(ruta))["Form sin contenido (parte C)"]
        assert total == 101
