# -*- coding: utf-8 -*-
"""Pruebas unitarias de db_lib / build_db_materiales (pytest)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import db_lib as L
import build_db_materiales as B


class TestNormalizacion:
    def test_guiones_tipograficos(self):
        assert L.txt("SA–516") == "SA-516"
        assert L.txt("SA−516") == "SA-516"

    def test_elipsis_asme(self):
        assert L.clean("… Seamless pipe") == "Seamless pipe"
        assert L.clean("   ") is None

    def test_search_key(self):
        assert L.search_key("SA–516", "70", None) == "SA-516 70"


class TestTemperaturas:
    def test_miles_con_coma(self):
        assert L.temp_to_number("1,000") == 1000

    def test_signo_menos_unicode(self):
        assert L.temp_to_number("−200") == -200
        assert L.temp_to_number("–125") == -125

    def test_no_numerico(self):
        assert L.temp_to_number("abc") is None

    def test_num(self):
        assert L.num("13.8") == 13.8
        assert L.num(None) is None


class TestClaves:
    def test_material_id_omite_vacios(self):
        assert L.build_material_id(["A106", "B", None, "", "P-1"]) == "A106 | B | P-1"

    def test_unicidad(self):
        out, coll = L.make_unique(["X", "X", "Y"], [1, 2, 3])
        assert len(set(out)) == 3 and len(coll) == 1

    def test_disambiguacion_por_notas(self):
        out, st = L.disambiguate(["A", "A", "B"], [("Notas", ["G4", "W6", None])],
                                 [1, 2, 3])
        assert len(set(out)) == 3 and st["Notas"] == 2 and st["linea"] == 0

    def test_disambiguacion_en_cascada(self):
        out, st = L.disambiguate(["A", "A"], [("Notas", [None, None]),
                                              ("Rm", [415, 485])], ["1.1", "1.2"])
        assert len(set(out)) == 2 and st["Rm"] == 2 and st["linea"] == 0

    def test_clave_bilingue(self):
        k = L.bilingual_key("A106", "B", "Pipe", "K03006", None, 1, 2)
        assert k.endswith("#2") and "A106" in k


class TestArtefactosDeExtraccion:
    def test_separa_nombre_y_valor_fusionados(self):
        nombre, extra = B.fix_merged_ident({"material_200": "N02200 222", "values": {}})
        assert nombre == "N02200" and extra == {"200": 222.0}

    def test_deja_intacto_lo_que_no_lo_es(self):
        nombre, extra = B.fix_merged_ident({"material": "Carbon steel"})
        assert nombre is None and extra == {}


class TestMapeoDeGrupos:
    """El grupo de propiedades sale de las Notas del codigo, nunca de una regla propia."""

    def test_no_queda_ninguna_heuristica_de_composicion(self):
        # La Rev. 2 asignaba grupo con expresiones regulares sobre la composicion
        # impresa. Producia falsos (`mo\b` capturaba austeniticos y aleaciones de
        # niquel). Si alguien la reintroduce, esta prueba lo detiene.
        assert not hasattr(B, "FERROUS_RULES")

    def test_clave_ignora_espacios_y_guiones_tipograficos(self):
        assert B.comp_key("18Cr–10Ni–Cb") == B.comp_key("18Cr - 10Ni - Cb")
        assert B.comp_key("16Cr–12Ni–2Mo") == "16CR-12NI-2MO"

    def test_clave_distingue_composiciones_parecidas(self):
        # El fallo de la Rev. 2: `8ni` casaba dentro de `18Ni`.
        assert B.comp_key("18Cr–18Ni–2Si") != B.comp_key("8Ni")
        assert B.comp_key("9Ni") != B.comp_key("18Cr–19Ni")

    def test_regla_textual_se_reconoce(self):
        assert B.REGLA_TEXTUAL.search("9Cr–Mo, including variations thereof")
        assert not B.REGLA_TEXTUAL.search("16Cr–12Ni–2Mo")

    def test_stem_de_regla_textual(self):
        assert B._stem_textual("9Cr–Mo, including variations thereof") == "9CR"

    def test_no_existe_estado_propuesta(self):
        # Una conjetura junto a datos normativos es peor que un hueco declarado.
        estados = {B.E_UNS, B.E_NOTA, B.E_TEXTUAL, B.E_SIN}
        assert not any("PROPUESTA" in e.upper() for e in estados)


class TestNotasDeGrupoEnResources:
    """Las Notas de TM-1 y TE-1 deben estar completas en resources/."""

    @staticmethod
    def _notas(archivo):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" /
                "bpvc_ii_d_metric_2025" / archivo)
        with open(ruta, encoding="utf-8") as fh:
            return json.load(fh).get("note_members", [])

    def test_tm1_trae_los_diez_grupos(self):
        notas = self._notas("table_tm_1.json")
        grupos = {n["grupo"] for n in notas if n["tipo"] == "grupo"}
        # A..J son diez grupos; la Nota (8) y la (9) definen ambas el Grupo H.
        assert grupos == {f"Material Group {L}" for L in "ABCDEFGHIJ"}

    def test_te1_trae_los_cuatro_grupos(self):
        notas = self._notas("table_te_1.json")
        grupos = {n["grupo"] for n in notas if n["tipo"] == "grupo"}
        assert grupos == {f"Group {i}" for i in (1, 2, 3, 4)}

    def test_se_conserva_la_duplicacion_impresa_de_la_nota_8(self):
        # ASME imprime (8) y (9) con identico contenido. Regla 9: los valores se
        # cargan tal como estan impresos, no se deduplican.
        notas = self._notas("table_tm_1.json")
        h = [n for n in notas if n.get("grupo") == "Material Group H"]
        assert {n["nota"] for n in h} == {"(8)", "(9)"}
        assert h[0]["miembros"] == h[1]["miembros"]

    def test_la_nota_1_no_perdio_la_segunda_columna(self):
        # Se imprime a dos columnas; la extraccion previa guardaba solo 4 de 8.
        notas = self._notas("table_tm_1.json")
        n1 = next(n for n in notas if n["nota"] == "(1)")
        assert len(n1["miembros"]) == 8
        assert "Mn–V" in n1["miembros"]

    def test_la_nota_2_recupera_la_continuacion_de_pagina(self):
        # El Grupo B arranca en el folio 1188 y termina en el 1189.
        notas = self._notas("table_tm_1.json")
        n2 = next(n for n in notas if n["nota"] == "(2)")
        assert len(n2["miembros"]) == 21
        assert "7Ni" in n2["miembros"]

    def test_los_austeniticos_no_caen_en_baja_aleacion(self):
        # El falso concreto que motivo el cambio: 16Cr-12Ni-2Mo (tipo 316) es
        # Grupo G en TM-1 y Grupo 3 en TE-1, no un acero de baja aleacion.
        tm = {m for n in self._notas("table_tm_1.json")
              if n.get("grupo") == "Material Group G" for m in n["miembros"]}
        te = {m for n in self._notas("table_te_1.json")
              if n.get("grupo") == "Group 3" for m in n["miembros"]}
        assert "16Cr–12Ni–2Mo" in tm and "16Cr–12Ni–2Mo" in te


class TestFormulas:
    def test_cascada_es_offset_clasico(self):
        f = B.cascade_formula("V", "K", "$C$5")
        assert f.startswith("=OFFSET(") and "COUNTIF(K,$C$5)" in f
        assert "_xlfn" not in f and "FILTER" not in f

    def test_banda_compacta_es_offset(self):
        refs = {"t_anchor": "H!$A$4", "v_anchor": "H!$B$4"}
        tr, vr = B.packed_rows(refs, "$C$1", "$C$2")
        assert tr.startswith("OFFSET(H!$A$4,$C$1-1,0,1,MAX(1,$C$2))")
        assert "MATCH(TRUE" not in tr and "LOOKUP(" not in tr

    def test_valor_no_extrapola(self):
        f = B.interp_value("A1", "B1", "C1", "D1", "$C$10", "$C$11")
        assert 'IF(OR(C1="",D1=""),B1' in f   # sin punto superior -> ultimo tabulado
        assert 'IF($C$10<=A1,B1' in f          # por debajo del minimo -> minimo tabulado

    def test_sin_funciones_de_matriz_dinamica(self):
        f = B.interp_value("A1", "B1", "C1", "D1", "$C$10", "$C$11")
        for bad in ("_xlfn", "FILTER(", "XLOOKUP(", "UNIQUE(", "SORT("):
            assert bad not in f
