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
