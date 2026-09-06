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
    def _notas(archivo, edicion="bpvc_ii_d_metric_2025"):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" /
                edicion / archivo)
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


class TestNotasEnLasDosEdiciones:
    """SI y US son extracciones independientes: ambas deben traer sus Notas.

    El libro construye bandas de propiedades en las dos ediciones (DB_E / DB_EC,
    DB_TE / DB_TEC), asi que dejar una sin notas deja media tabla sin forma
    trazable de saber a que grupo pertenece un material.
    """

    # staticmethod() explicito: al reasignar la funcion como atributo de clase
    # se pierde el descriptor y Python volveria a inyectar `self` como 1er arg.
    _N = staticmethod(TestNotasDeGrupoEnResources._notas)
    US = "bpvc_ii_d_customary_2025"

    def test_la_edicion_us_trae_sus_notas(self):
        assert {n["grupo"] for n in self._N("table_tm_1.json", self.US)
                if n["tipo"] == "grupo"} == {f"Material Group {L}" for L in "ABCDEFGHIJ"}
        assert {n["grupo"] for n in self._N("table_te_1.json", self.US)
                if n["tipo"] == "grupo"} == {f"Group {i}" for i in (1, 2, 3, 4)}

    def test_la_us_no_duplica_el_grupo_h(self):
        # La duplicacion es una errata EXCLUSIVA de la edicion metrica.
        h = [n for n in self._N("table_tm_1.json", self.US)
             if n.get("grupo") == "Material Group H"]
        assert len(h) == 1 and h[0]["nota"] == "(8)"

    def test_la_numeracion_va_corrida_entre_ediciones(self):
        # Consecuencia de la errata: a partir del Grupo H la metrica numera uno
        # mas que la US. Dar por hecho que coinciden romperia el mapeo.
        si = [n["nota"] for n in self._N("table_tm_1.json")
              if n.get("grupo") == "Material Group J"]
        us = [n["nota"] for n in self._N("table_tm_1.json", self.US)
              if n.get("grupo") == "Material Group J"]
        assert si == ["(11)"] and us == ["(10)"]

    def test_cada_edicion_cita_la_nota_que_su_tabla_referencia(self):
        # La metrica trae el Grupo H en las Notas (8) y (9) pero su tabla apunta
        # a la (9); la US lo trae solo en la (8). Citar la nota huerfana
        # remitiria al lector a algo que la tabla no referencia.
        import json
        def tabla(ed):
            ruta = (Path(__file__).resolve().parents[3] / "resources" / ed /
                    "table_tm_1.json")
            with open(ruta, encoding="utf-8") as fh:
                return B.notas_referenciadas(json.load(fh))
        assert tabla("bpvc_ii_d_metric_2025")["Material Group H"] == "(9)"
        assert tabla(self.US)["Material Group H"] == "(8)"

    def test_la_pertenencia_es_la_misma_en_ambas_ediciones(self):
        # Los valores difieren (MPa vs ksi) pero el reparto de materiales en
        # grupos no: si difiriera, una de las dos extracciones estaria mal.
        for archivo in ("table_tm_1.json", "table_te_1.json"):
            si = {n["grupo"]: set(n["miembros"]) for n in self._N(archivo)
                  if n["tipo"] == "grupo"}
            us = {n["grupo"]: set(n["miembros"]) for n in self._N(archivo, self.US)
                  if n["tipo"] == "grupo"}
            assert si == us, archivo


class TestBarrasDeFraccion:
    """II-D imprime «C-1/2Mo» y el Apendice A del B31.3 «C-1∕2Mo» (U+2215)."""

    def test_las_barras_se_pliegan(self):
        assert B.comp_key("C–1∕" + "2Mo") == B.comp_key("C-1/2Mo")
        assert B.comp_key("1Cr–1⁄4Mo") == B.comp_key("1Cr-1/4Mo")

    def test_sin_esto_el_contraste_entre_tablas_fallaria_en_silencio(self):
        # No lanza excepcion: simplemente no encontraria nada. Por eso hay prueba.
        import unicodedata
        assert unicodedata.normalize("NFKC", "∕") != "/"


class TestComposicionPrestada:
    """Filas que no imprimen composicion: se recupera por UNS, nunca como AUTO."""

    NOTAS = {"C-1/2MO": [("TM-1", "(1)", "Material Group A")]}

    def test_recupera_cuando_el_uns_tiene_una_sola_composicion(self):
        idx = {"K11522": {"C–1∕2Mo": ["DB_B31_3"]}}
        r = B._composicion_prestada(idx, "K11522", "", self.NOTAS)
        assert r is not None
        comp, hojas, origenes = r
        assert hojas == ["DB_B31_3"] and origenes[0][2] == "Material Group A"

    def test_no_elige_cuando_el_codigo_da_dos_composiciones(self):
        idx = {"K11522": {"C–1∕2Mo": ["DB_B31_3"], "C–1Mo": ["DB_Su"]}}
        assert B._composicion_prestada(idx, "K11522", "", self.NOTAS) is None

    def test_no_se_usa_si_la_fila_ya_trae_composicion_propia(self):
        idx = {"K11522": {"C–1∕2Mo": ["DB_B31_3"]}}
        assert B._composicion_prestada(idx, "K11522", "18CR-8NI", self.NOTAS) is None

    def test_el_motivo_distingue_por_que_fallo(self):
        dos = {"K11522": {"C–1∕2Mo": ["DB_B31_3"], "C–1Mo": ["DB_Su"]}}
        assert "mas de una" in B._motivo_sin_composicion(dos, "K11522")
        una = {"K11522": {"C–1∕2Mo": ["DB_B31_3"]}}
        assert "no figura en ninguna Nota" in B._motivo_sin_composicion(una, "K11522")
        assert "no aparece con" in B._motivo_sin_composicion({}, "K11522")


class TestDecisionesDelIngeniero:
    """La vuelta al motor de lo que el ingeniero decide."""

    @staticmethod
    def _archivo(tmp_path, decisiones):
        import json
        p = tmp_path / "decisiones_map_grupo.json"
        p.write_text(json.dumps({"decisiones": decisiones}, ensure_ascii=False),
                     encoding="utf-8")
        return p

    def test_sin_archivo_no_falla(self, tmp_path):
        assert B.cargar_decisiones(tmp_path / "no_existe.json") == ({}, {})

    def test_ignora_las_entradas_de_plantilla_sin_rellenar(self, tmp_path):
        p = self._archivo(tmp_path, [{"composicion": "18Cr–8Ni", "grupo_tm": "",
                                      "grupo_te": ""}])
        assert B.cargar_decisiones(p) == ({}, {})

    def test_ancla_por_composicion_y_por_uns(self, tmp_path):
        p = self._archivo(tmp_path, [
            {"composicion": "9Cr–1Mo–V", "grupo_tm": "Material Group E"},
            {"uns": "S30400", "grupo_te": "Group 3"},
        ])
        comp, uns = B.cargar_decisiones(p)
        assert B.comp_key("9Cr-1Mo-V") in comp
        assert "S30400" in uns

    def test_el_estado_validado_es_distinto_de_auto(self):
        # Quien audite el libro tiene que poder separar lo que dice el codigo de
        # lo que decidio una persona.
        assert B.E_VALIDADO not in (B.E_UNS, B.E_NOTA)
        assert "AUTO" not in B.E_VALIDADO

    def test_la_firma_deja_constancia_aunque_falte(self):
        assert B._firma({}) == "sin firma, sin fecha"
        assert B._firma({"validado_por": "DV", "fecha": "2026-09-06"}) == \
            "DV, 2026-09-06"


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
