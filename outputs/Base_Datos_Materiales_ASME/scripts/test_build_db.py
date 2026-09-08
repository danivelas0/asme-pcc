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
    def _notas(archivo, edicion="ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025"):
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
    US = "ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025"

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
        assert tabla("ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025")["Material Group H"] == "(9)"
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
            {"composicion": "9Cr–1Mo–V", "grupo_tm": "Material Group E",
             "validado_por": "DV", "fecha": "2026-09-06"},
            {"uns": "S30400", "grupo_te": "Group 3",
             "validado_por": "DV", "fecha": "2026-09-06"},
        ])
        comp, uns = B.cargar_decisiones(p)
        assert B.comp_key("9Cr-1Mo-V") in comp
        assert "S30400" in uns

    def test_el_nombre_es_opcional(self, tmp_path):
        # `validado_por` se registra si esta, pero no se exige: lo que separa una
        # decision de un dato del codigo es el ESTADO de la fila, no que alguien
        # haya escrito su nombre en un JSON.
        p = self._archivo(tmp_path, [
            {"uns": "S30400", "grupo_tm": "Material Group G", "validado_por": ""},
            {"composicion": "18Cr–8Ni", "grupo_tm": "Material Group G",
             "validado_por": "DV", "fecha": "2026-09-06"},
        ])
        comp, uns = B.cargar_decisiones(p)
        assert len(uns) == 1 and len(comp) == 1

    def test_el_estado_validado_es_distinto_de_auto(self):
        # Quien audite el libro tiene que poder separar lo que dice el codigo de
        # lo que decidio una persona.
        assert B.E_VALIDADO not in (B.E_UNS, B.E_NOTA)
        assert "AUTO" not in B.E_VALIDADO

    def test_deja_constancia_aunque_falte_el_nombre(self):
        assert B._firma({}) == "sin nombre, sin fecha"
        assert B._firma({"validado_por": "DV", "fecha": "2026-09-06"}) == \
            "DV, 2026-09-06"


class TestColumnasNombradasTE1:
    """TE-1 publica dilatacion en columnas que se autodescriben, no solo por Grupo."""

    @staticmethod
    def _cols(edicion="ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025"):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" / edicion /
                "table_te_1.json")
        with open(ruta, encoding="utf-8") as fh:
            return B.columnas_nombradas_te1(json.load(fh))

    def test_reparte_los_titulos_con_varias_composiciones(self):
        c = self._cols()
        # «12Cr, 12Cr-1Al, 13Cr, and 13Cr-4Ni Steels» son cuatro designaciones.
        for k in ("12CR", "12CR-1AL", "13CR", "13CR-4NI", "15CR", "17CR", "27CR"):
            assert k in c, k

    def test_no_arrastra_la_palabra_steels_a_la_clave(self):
        # «9Cr-1Mo Steels» daba la clave «9CR-1MOSTEELS», que no casa con nada:
        # la columna quedaba inservible sin dar ningun error.
        c = self._cols()
        assert "9CR-1MO" in c and not any("STEELS" in k for k in c)

    def test_separa_la_condicion_del_titulo(self):
        c = self._cols()
        _, cond = c["9CR-1MO"]
        assert cond and "91" in cond          # «Including Grades 9, 91, 911, and 92»
        assert c["15CR"][1] is None           # esta no condiciona nada

    def test_no_duplica_los_grupos_numerados(self):
        # Los Grupos 1..4 llegan por Nota; no deben entrar tambien por columna.
        assert not any("GROUP" in k for k in self._cols())

    def test_las_dos_ediciones_dan_las_mismas_designaciones(self):
        # La metrica rotula «7% Nickel Steel» y la US «7Ni Steels»: es la misma
        # columna. Sin unificarlo, el mismo material tendria dilatacion en una
        # hoja y no en la otra.
        si, us = self._cols(), self._cols("ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025")
        assert set(si) == set(us)
        for k in si:
            assert (si[k][1] or "") == (us[k][1] or ""), k

    def test_la_condicion_de_tratamiento_es_pegajosa(self):
        # La extraccion trunca una de las tres columnas del 17Cr-4Ni-4Cu y la
        # truncada pierde el «Condition 1075». Si esa variante ganase, se
        # asignaria dilatacion eligiendo a ciegas entre dos tratamientos con
        # valores distintos.
        c = self._cols()
        assert c["17CR-4NI-4CU"][1].startswith("Condition")


class TestCondicionesTratamientoTE1:
    """columnas_nombradas_te1 colapsa el tratamiento a UNA condicion por
    composicion (pegajosa, ver arriba); condiciones_tratamiento_te1 no
    colapsa nada, para poder resolver la fila contra su propio tratamiento
    impreso (Clase/Cond./Temple)."""

    @staticmethod
    def _cond(edicion="ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025"):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" / edicion /
                "table_te_1.json")
        with open(ruta, encoding="utf-8") as fh:
            return B.condiciones_tratamiento_te1(json.load(fh))

    def test_conserva_las_dos_condiciones_del_17cr_4ni_4cu(self):
        c = self._cond()
        assert set(c["17CR-4NI-4CU"]) == {"1075", "1150"}
        assert "Condition 1075" in c["17CR-4NI-4CU"]["1075"]
        assert "Condition 1150" in c["17CR-4NI-4CU"]["1150"]

    def test_no_incluye_columnas_sin_tratamiento_numerico(self):
        # «9Cr-1Mo Steels (Including Grades 9, 91, 911, and 92)» condiciona por
        # GRADO, no por tratamiento termico: no debe aparecer aqui.
        c = self._cond()
        assert "9CR-1MO" not in c

    def test_las_dos_ediciones_dan_las_mismas_condiciones(self):
        si, us = self._cond(), self._cond("ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025")
        assert set(si["17CR-4NI-4CU"]) == set(us["17CR-4NI-4CU"])


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

    def test_el_apendice_c_bloquea_en_los_dos_extremos(self):
        # Los cinco buscadores de esfuerzos sostienen el ultimo valor tabulado
        # porque el tope se lo pone la columna «Temp. max.» del codigo. El
        # Apendice C no publica esa columna: si aqui se sostuviera el extremo se
        # estaria extrapolando, que es lo que prohibe la regla 4.
        f = B.formula_estado_apxc("$X$1", "$X$2", "$X$3", "$X$4", "$X$5", "$D$11")
        assert f'IF($D$11<$X$4,"{B.EST_BAJO}"' in f
        assert f'IF($D$11>$X$5,"{B.EST_ALTO}"' in f
        assert f'IF($X$2="{B.TIPO_PUNTO}","{B.EST_PUNTO}"' in f
        for bad in ("_xlfn", "FILTER(", "XLOOKUP(", "UNIQUE(", "SORT("):
            assert bad not in f

    def test_el_valor_del_apendice_c_se_bloquea_fuera_de_rango(self):
        f = B.formula_valor_apxc("$X$1", "$X$2", "$X$3", "$X$4", "$X$5",
                                 "$X$6", "$X$7")
        assert f'ISNUMBER(SEARCH("FUERA DE RANGO",$X$6)),"{B.VAL_BLOQUEADO}"' in f
        # El factor de escala se aplica a la VISTA, en las dos ramas.
        assert "$X$3*$X$5" in f and "$X$7*$X$5" in f


class TestApendiceCUnificado:
    """La base DB_B31_C: contrato de columnas, factores y tipado de celda."""

    def test_las_ocho_primeras_columnas_son_las_de_stress_cols(self):
        # Es lo que permite que build_listas y las secciones 2 y 4 de
        # verificar.py funcionen sobre esta base sin tocar una linea.
        assert B.APXC_COLS[:8] == B.STRESS_COLS[:8]
        assert B.APXC_COLS[-1] == "n_pts" and B.CC["n_pts"] == B.N_IDENT_C

    def test_el_factor_se_lee_del_json_y_no_esta_codificado(self):
        # Confundir 10^3 con 10^6 son tres ordenes de magnitud en el modulo E.
        assert B._factor_escala({"operacion": "multiplicar", "exponente": 3,
                                 "impreso": "x10^3"}) == (1000, "x10^3")
        assert B._factor_escala({"operacion": "multiplicar", "exponente": 6})[0] == 10 ** 6
        assert B._factor_escala({"operacion": "dividir", "exponente": 6})[0] == 10 ** -6
        assert B._factor_escala({"operacion": "ninguna", "exponente": 0})[0] == 1
        assert B._factor_escala(None)[0] == 1

    def test_las_referencias_de_nota_salen_del_rotulo_impreso(self):
        assert B._notas_ref("Group 1 carbon and low alloy steels [Note (2)]") == "(2)"
        assert B._notas_ref("Thermoplastics [Note (1)]", "x [Note (1)]") == "(1)"
        assert B._notas_ref("Carbon steels") == ""

    def test_el_separador_de_cascada_no_puede_aparecer_en_un_segmento(self):
        # El separador de las claves k1..k4 es '|': un rotulo que lo llevase
        # partiria la cascada en build_listas sin dar ningun error.
        assert "|" not in B._seg("A|B")
        assert B._seg(None, B.SIN_GRUPO) == B.SIN_GRUPO

    def test_la_lista_blanca_de_divergencias_si_us_es_cerrada(self):
        # El enlace SI/US del Apendice C es posicional. La unica divergencia de
        # nombre admitida es la del 'Type 309' (punto en C-3, coma en C-3C).
        assert len(B.DIVERGENCIAS_NOMBRE_C) == 1
        (prop, si), (us, motivo) = next(iter(B.DIVERGENCIAS_NOMBRE_C.items()))
        assert prop == "C3" and si == "TYPE309.23CR-12NI" and us == "TYPE309,23CR-12NI"
        assert motivo

    def test_normaliza_espacios_y_guiones_al_comparar_ediciones(self):
        # '29Cr–7Ni– 2Mo–N' (C-1C) y '29Cr–7Ni–2Mo–N' (C-1) son el mismo
        # material: la unica diferencia es un espacio del salto de linea.
        assert (B._norm_nombre_c("5Cr–1Mo and 29Cr–7Ni– 2Mo–N steels") ==
                B._norm_nombre_c("5Cr–1Mo and 29Cr–7Ni–2Mo–N steels"))
        assert B._norm_nombre_c("Type 310, 25Cr–20Ni") == \
            B._norm_nombre_c("Type 310, 25Cr-20Ni")

    def test_carga_las_cuatro_tablas_con_su_reparto_impreso(self):
        from db_lib import Resources
        res = Resources(str(Path(__file__).resolve().parents[3] / "resources"))
        for si, tags in ((True, ("C-1", "C-3")), (False, ("C-1C", "C-3C"))):
            filas = B._carga_apendice_c(res, si)
            from collections import Counter
            n = Counter(d["prop"] for d in filas)
            assert n[B.PROP_C1] == 52 and n[B.PROP_C2] == 44
            assert n[B.PROP_C3] == 75 and n[B.PROP_C4] == 30
            assert len(filas) == 201
            # Curva frente a punto: C-2 y C-4 no dependen de la temperatura.
            for d in filas:
                if d["prop"] in (B.PROP_C2, B.PROP_C4):
                    assert d["tipo"] == B.TIPO_PUNTO and not d["vals"]
                else:
                    assert d["tipo"] == B.TIPO_CURVA and d["vals"]
            tablas = {d["tabla"] for d in filas}
            assert tags[0] in tablas and tags[1] in tablas

    def test_los_intervalos_de_texto_no_se_convierten_en_numero(self):
        from db_lib import Resources
        res = Resources(str(Path(__file__).resolve().parents[3] / "resources"))
        filas = {(d["tabla"], d["material"]): d
                 for d in B._carga_apendice_c(res, True)}
        # C-2 publica '16–23.5' para el vidrio-epoxi bobinado: es un intervalo,
        # no un numero. Parsearlo seria interpretar, no transcribir (regla 9).
        d = filas[("C-2", "Glass–epoxy, filament-wound")]
        assert d["vnum"] is None and d["vtxt"] == "16–23.5"
        # C-4 publica '8 275–13 100' para el vidrio-epoxi centrifugado.
        d = filas[("C-4", "Epoxy–glass, centrifugally cast")]
        assert d["vnum"] is None and d["vtxt"]
        # Y un valor numerico sigue siendo numerico.
        d = filas[("C-4", "Acetal")]
        assert d["vnum"] == 2830 and d["vtxt"] is None

    def test_los_factores_de_escala_no_se_confunden_entre_ediciones(self):
        from db_lib import Resources
        res = Resources(str(Path(__file__).resolve().parents[3] / "resources"))
        for si, esperado in ((True, 10 ** 3), (False, 10 ** 6)):
            mod = [d for d in B._carga_apendice_c(res, si)
                   if d["prop"] == B.PROP_C3]
            assert {d["factor"] for d in mod} == {esperado}
            assert all("Multiply Tabulated Values" in d["factor_txt"] for d in mod)
        # C-2 divide, no multiplica; C-1 lleva el 10^-6 dentro de la unidad.
        filas = B._carga_apendice_c(res, True)
        assert {d["factor"] for d in filas if d["prop"] == B.PROP_C2} == {10 ** -6}
        assert {d["factor"] for d in filas if d["prop"] == B.PROP_C1} == {1}

    def test_el_dedup_colapsa_el_rotulo_repetido_por_la_extraccion(self):
        assert B._dedup_frase("Gray iron Gray iron") == "Gray iron"
        assert B._dedup_frase("Carbon steels") == "Carbon steels"

    def test_la_variante_separa_las_filas_que_el_codigo_repite(self):
        from db_lib import Resources
        res = Resources(str(Path(__file__).resolve().parents[3] / "resources"))
        pfa = [d for d in B._carga_apendice_c(res, True)
               if d["tabla"] == "C-2" and "perfluoroalkoxy" in d["material"]]
        # Tres filas con el mismo nombre y rangos de validez distintos: el rango
        # es lo que las separa, asi que es la variante de la cascada.
        assert len(pfa) == 3
        assert len({d["variante"] for d in pfa}) == 3
        assert all(d["variante"] != B.VAR_UNICA for d in pfa)
        # Y los coeficientes A/B de C-1 se rotulan con su magnitud impresa.
        c1 = [d for d in B._carga_apendice_c(res, True) if d["tabla"] == "C-1"]
        assert {d["variante"] for d in c1} == set(B.VARIANTE_COEF.values())
        assert all(d["definicion"].startswith(("A =", "B =")) for d in c1)


# ---------------------------------------------------------------------------
# La capa de metadatos del Apendice C del B31.3
# ---------------------------------------------------------------------------
# Los VALORES del Apendice C siempre estuvieron bien; lo que faltaba era todo lo
# que rodea al numero, y sin eso el motor de propiedades no puede ni rotularse:
#   - C-1/C-1C no declaraban la unidad de sus coeficientes A y B;
#   - las Notas (2)..(6) de C-1 se imprimen a tres columnas y se habian leido por
#     filas, dejando los miembros de los Grupos 1..4 intercalados;
#   - C-2 y C-3 perdieron la indentacion del impreso y cinco filas quedaron
#     colgando de un subgrupo que no es el suyo ("Gray iron" como austenitico).
# `completar_apendice_c.py` lo repara desde el OCR del escaneo. Estas pruebas
# impiden que una reextraccion futura vuelva a dejarlo a medias.
class TestApendiceCCompletado:
    """Metadatos del Apendice C, contra los folios impresos 408-428."""

    @staticmethod
    def _tabla(archivo):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" / "ASME B31" /
                "ASME B31.3" / "APPEX" / "appendix_c" / archivo)
        with open(ruta, encoding="utf-8") as fh:
            return json.load(fh)

    def test_c1_declara_la_unidad_de_sus_coeficientes(self):
        # Folios 408 y 414. Sin esto el KPI del motor muestra un numero desnudo.
        si = self._tabla("table_c_1.json")["coefficient_defs"]
        us = self._tabla("table_c_1c.json")["coefficient_defs"]
        assert si["A"]["unidad"] == "10^-6 mm/mm/°C"
        assert si["B"]["unidad"] == "mm/m"
        assert us["A"]["unidad"] == "10^-6 in./in./°F"
        assert us["B"]["unidad"] == "in./100 ft"
        assert "20°C" in si["referencia"] and "70°F" in us["referencia"]

    def test_las_notas_de_c1_traen_sus_miembros(self):
        notas = {n["nota"]: n for n in self._tabla("table_c_1.json")["note_members"]}
        assert {k: len(v["miembros"]) for k, v in notas.items()} == {
            "(2)": 51, "(3)": 6, "(4)": 13, "(5)": 14, "(6)": 18}
        assert notas["(2)"]["grupo"] == "Group 1"
        # La Nota (6) no define un Grupo: enumera aleaciones de aluminio por UNS.
        assert notas["(6)"]["tipo"] == "alias_uns" and notas["(6)"]["grupo"] is None

    def test_las_notas_se_leyeron_por_columnas(self):
        # Leidas por filas, el segundo miembro seria el primero de la 2a columna.
        # El impreso lista la 1a columna entera antes de pasar a la 2a.
        m = self._tabla("table_c_1.json")["note_members"][0]["miembros"]
        assert m[:3] == ["Carbon steel", "C-Mn-Cb", "C-Mn-Si-Cb"]
        assert m[17] == "1Cr-1/2Mo-V"     # inicio de la 2a columna
        assert m[34] == "1/2Ni-1/2Cr-1/4Mo-V"   # inicio de la 3a

    def test_las_dos_ediciones_listan_los_mismos_miembros(self):
        # A diferencia de TM-1 de la II-D, el Apendice C NO tiene errata de
        # numeracion: C-1 y C-1C imprimen las mismas Notas (folios 412 y 418).
        si = self._tabla("table_c_1.json")["note_members"]
        us = self._tabla("table_c_1c.json")["note_members"]
        assert [n["nota"] for n in si] == [n["nota"] for n in us]
        assert [n["miembros"] for n in si] == [n["miembros"] for n in us]

    def test_c2_reparte_sus_44_filas_como_el_impreso(self):
        from collections import Counter
        filas = self._tabla("table_c_2.json")["rows"]
        assert Counter(r.get("material_group") for r in filas) == Counter({
            "Thermoplastics": 38,
            "Reinforced Thermosetting Resins and Reinforced Plastic Mortars": 5,
            "Other Nonmetallic Materials": 1})
        # El titulo llegaba partido por la mitad ("and Reinforced Plastic Mortars").
        assert all(not (r.get("material_group") or "").startswith("and ") for r in filas)
        # Ninguna fila sin grupo: nueve lo perdieron al cruzar el corte de pagina.
        assert all(r.get("material_group") for r in filas)

    def test_ninguna_fila_cuelga_de_un_subgrupo_ajeno(self):
        # Las cinco del folio 419-421 mas las siete del 420 van A RAS del margen.
        a_ras = {"Polybutylene PB 2110", "Polyether, chlorinated",
                 "Polyphenylene POP 2125", "Poly(vinylidene fluoride)",
                 "Poly(tetrafluoroethylene)", "Poly(perfluoroalkoxy alkane)"}
        for r in self._tabla("table_c_2.json")["rows"]:
            if r["material_description"] in a_ras:
                assert r.get("material_subgroup") is None, r["material_description"]
        for f in ("table_c_3.json", "table_c_3c.json"):
            for r in self._tabla(f)["rows"]:
                if r["material"] in ("Gray iron",
                                     "Straight chromium stainless steels "
                                     "(12Cr, 17Cr, 27Cr)"):
                    assert r.get("material_subgroup") is None, (f, r["material"])
        # Y "Austenitic stainless steels:" cubre solo los seis Type ...
        aust = [r["material"] for r in self._tabla("table_c_3.json")["rows"]
                if r.get("material_subgroup") == "Austenitic stainless steels:"]
        assert len(aust) == 6 and all(m.startswith("Type ") for m in aust)

    def test_los_cont_d_no_son_un_grupo_distinto(self):
        for f in ("table_c_3.json", "table_c_3c.json"):
            grupos = {r.get("material_group") for r in self._tabla(f)["rows"]}
            assert not any("Cont" in (g or "") for g in grupos), f

    def test_los_factores_de_escala_estan_declarados(self):
        # Confundir 10^3 con 10^6 en el modulo E son tres ordenes de magnitud.
        esperado = {"table_c_2.json": ("dividir", 6),
                    "table_c_3.json": ("multiplicar", 3),
                    "table_c_3c.json": ("multiplicar", 6)}
        for f, (op, exp) in esperado.items():
            sf = self._tabla(f)["scale_factor"]
            assert (sf["operacion"], sf["exponente"]) == (op, exp), f
            assert "impreso" in sf, f   # debe trazar al encabezado del codigo

    def test_la_enmienda_queda_declarada(self):
        for f in ("table_c_1.json", "table_c_1c.json", "table_c_2.json",
                  "table_c_3.json", "table_c_3c.json", "table_c_4.json"):
            am = self._tabla(f)["extraction_amendments"]
            assert am["folios_impresos"] and am["fuente"]
            # Regla 9: la enmienda no puede haber tocado valores ni nombres.
            assert "valores y los nombres" in am["alcance"]


# ---------------------------------------------------------------------------
# La identificacion y la estructura del Apendice B del B31.3
# ---------------------------------------------------------------------------
# Los VALORES del Apendice B estaban bien; lo que fallaba era como quedaron
# identificadas y estructuradas las filas. El caso peligroso es B-5: la cabecera
# se desalineo y el MAXIMO de 232 C quedo guardado en un campo llamado `c`, de
# modo que quien lo leyese como temperatura minima se llevaba 232 C de minimo en
# un tubo de vidrio. `completar_apendice_b.py` lo repara desde el impreso.
class TestApendiceBCorregido:
    """Identificacion y estructura del Apendice B, contra los folios 399-405."""

    @staticmethod
    def _tabla(archivo):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" / "ASME B31" /
                "ASME B31.3" / "APPEX" / "appendix_b" / archivo)
        with open(ruta, encoding="utf-8") as fh:
            return json.load(fh)

    def test_b1_separa_la_especificacion_de_la_designacion(self):
        # La extraccion pegaba la elipsis del codigo y la designacion de tuberia
        # en un mismo campo. B-1C nunca lo hizo: es defecto solo de B-1.
        filas = {r["material_designation"]: r for r in self._tabla("table_b_1.json")["rows"]}
        assert filas["ABS"]["astm_spec_no"] is None
        assert filas["ABS"]["pipe_designation"] == "PR"
        for m in ("PP-R", "PP-RCT"):
            assert filas[m]["astm_spec_no"] == "F2389"
            assert filas[m]["pipe_designation"] == "PR"

    def test_la_spec_partida_en_dos_lineas_se_recompone(self):
        # El indice la escribe "F2788/F2788M"; con el espacio no emparejaba.
        for f in ("table_b_1.json", "table_b_1c.json"):
            specs = {r["astm_spec_no"] for r in self._tabla(f)["rows"]}
            assert "F2788/F2788M" in specs, f
            assert not any(s and " " in s for s in specs), f

    def test_b3_lista_sus_seis_especificaciones(self):
        # El impreso es una rejilla de 3 columnas: se lee por columnas y quedan
        # en orden ascendente, igual que en el indice.
        filas = self._tabla("table_b_3.json")["rows"]
        assert [r["spec_nos_astm_except_as_noted"] for r in filas] == [
            "D2517", "D2996", "D2997", "D3517", "D3754", "AWWA C950"]
        assert self._tabla("table_b_3.json")["row_count"] == 6

    def test_b4_y_b5_reparten_minimo_y_maximo(self):
        for f in ("table_b_4.json", "table_b_5.json"):
            d = self._tabla(f)
            assert [c for c in d["columns"] if "°" in c] == [
                "Minimum °C", "Minimum °F", "Maximum °C", "Maximum °F"], f
            for r in d["rows"]:
                assert "c" not in r and "f" not in r, f
                assert {"minimum_c", "minimum_f", "maximum_c", "maximum_f"} <= set(r), f

    def test_b5_guarda_los_232C_como_maximo(self):
        # Es el fallo que hacia dano: 232 C es el MAXIMO del vidrio borosilicato.
        for r in self._tabla("table_b_5.json")["rows"]:
            assert r["maximum_c"] == 232 and r["maximum_f"] == 450
            assert r["minimum_c"] is None and r["minimum_f"] is None

    def test_b6_recompone_el_material_de_dos_lineas(self):
        filas = self._tabla("table_b_6.json")["rows"]
        mats = [r["material"] for r in filas]
        assert "Metal insert fittings for PE-AL-PE systems" in mats
        assert "Metal insert fittings for" not in mats   # nombre truncado
        assert "PE-AL-PE systems" not in mats            # segunda linea suelta
        # La tabla solo publica limites maximos.
        assert all("f" not in r and "maximum_f" in r for r in filas)

    def test_el_indice_trae_sus_notas(self):
        d = self._tabla("spec_index_b.json")
        assert d["entry_count"] == len(d["entries"]) == 28
        texto = " ".join(d["notes"])
        assert "A326.4" in texto            # Nota (1)
        assert "fiberglass RTR" in texto    # Nota (2), normativa

    def test_las_rarezas_del_codigo_quedan_declaradas(self):
        # No se corrigen (regla 9), pero no pueden pasar desapercibidas.
        assert any("D2846" in o for o in
                   self._tabla("table_b_1.json")["observaciones_del_codigo"])
        assert any("100 psi" in o for o in
                   self._tabla("table_b_6.json")["observaciones_del_codigo"])

    def test_la_enmienda_queda_declarada(self):
        for f in ("table_b_1.json", "table_b_1c.json", "table_b_3.json",
                  "table_b_4.json", "table_b_5.json", "table_b_6.json",
                  "spec_index_b.json"):
            am = self._tabla(f)["extraction_amendments"]
            assert am["folios_impresos"] and am["fuente"]
            assert "valores NO se tocan" in am["alcance"]


# ---------------------------------------------------------------------------
# La costura de las paginas enfrentadas del Apendice A
# ---------------------------------------------------------------------------
# A-1 y A-1C se imprimen a doble pagina: identificacion a la izquierda, numero de
# linea y rejilla de esfuerzos a la derecha. Al fusionarlas se colaron tres
# defectos, y uno perdia dato: en el bloque de niquel de A-1C la palabra "Metal"
# del titulo de banda se pego al encabezado "200" y el esfuerzo a 200 F acabo en
# un campo de identificacion inventado, fuera de la curva.
class TestApendiceACorregido:
    """Desplazamientos de columna del Apendice A, contra los folios 166-398."""

    @staticmethod
    def _tabla(archivo):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" / "ASME B31" /
                "ASME B31.3" / "APPEX" / "appendix_a" / archivo)
        with open(ruta, encoding="utf-8") as fh:
            return json.load(fh)

    def test_a1c_no_pierde_el_punto_de_200F(self):
        d = self._tabla("table_a_1c.json")
        assert not any("metal_200" in r for r in d["rows"])
        assert "Metal 200" not in (d.get("columns_identification") or [])
        # 197 filas lo tenian fuera de la curva y otras 21 lo habian perdido con
        # su pagina izquierda: las 218 del bloque de niquel tienen que tenerlo.
        niquel = [r for r in d["rows"] if r.get("block") == 6]
        con200 = [r for r in niquel if (r.get("values") or {}).get("200") is not None]
        assert len(niquel) == 218 and len(con200) == 218, (len(niquel), len(con200))
        # Y la curva tiene que decrecer: 200 F no puede valer menos que 300 F.
        for r in con200:
            v = r["values"]
            if v.get("300") is not None:
                assert v["200"] >= v["300"], (r.get("spec_no"), v["200"], v["300"])

    def test_a3_separa_clase_y_descripcion(self):
        filas = self._tabla("table_a_3.json")["rows"]
        assert all(r.get("description") for r in filas), "ninguna descripcion vacia"
        # La elipsis del codigo significa "sin valor": no puede quedar en el campo.
        assert not any("…" in str(r.get("class_or_type") or "") for r in filas)
        # Las tres separaciones que no eran mecanicas (folios 366 y 368).
        def fila(spec, desc):
            return next(r for r in filas
                        if r.get("spec_no") == spec and r.get("description") == desc)
        assert fila("A53", "Seamless pipe")["class_or_type"] == "Type S"
        assert fila("B675", "Welded pipe")["class_or_type"] == "All"
        assert fila("API 5L", "Electric welded pipe")["class_or_type"] is None
        # Y el primer registro del codigo: API 5L, sin clase, Ej = 1.00
        assert fila("API 5L", "Seamless pipe")["ej"] == 1.0

    def test_ninguna_elipsis_suelta_en_la_identificacion(self):
        campos = ("material", "nominal_composition", "product_form", "spec_no",
                  "type_grade", "uns_no", "class_condition_temper", "p_no")
        for f in ("table_a_1.json", "table_a_1c.json", "table_a_4.json",
                  "table_a_4c.json"):
            for r in self._tabla(f)["rows"]:
                for k in campos:
                    v = r.get(k)
                    if isinstance(v, str):
                        assert "…" not in v, (f, k, v)

    def test_la_enmienda_queda_declarada(self):
        for f in ("table_a_1.json", "table_a_1c.json", "table_a_3.json",
                  "table_a_4.json", "table_a_4c.json"):
            am = self._tabla(f)["extraction_amendments"]
            assert am["folios_impresos"] and am["fuente"]
            assert "Ningun valor tabulado se altera" in am["alcance"]

    def test_no_queda_ninguna_pagina_derecha_huerfana(self):
        # En 21 lineas del bloque 6 de A-1C la fusion de paginas enfrentadas
        # perdio la pagina IZQUIERDA entera. Se recuperaron leyendo los folios
        # 336, 338, 340, 342, 344 y 346. Ninguna fila puede volver a quedarse sin
        # identificacion ni sin su punto de 200 F.
        d = self._tabla("table_a_1c.json")
        b6 = [r for r in d["rows"] if r.get("block") == 6]
        assert not [r for r in b6 if not r.get("spec_no")]
        assert not [r for r in b6 if (r.get("values") or {})
                    and "200" not in r["values"]]
        assert all(r.get("min_temp_to_100") is not None for r in b6)

    def test_las_21_recuperadas_llevan_lo_que_imprime_el_folio(self):
        d = {(r.get("block"), r.get("line_no")): r for r in
             self._tabla("table_a_1c.json")["rows"]}
        # Folio 344: la linea 203 de A-1C NO es la 203 de A-1 -las ediciones no
        # numeran igual-, por eso cada fila se leyo en su propia pagina US.
        r = d[(6, 203)]
        assert (r["spec_no"], r["uns_no"], r["class_condition_temper"]) == \
            ("B649", "N08031", "Annealed")
        assert r["min_temp_to_100"] == 26.7 and r["values"]["200"] == 26.7
        # Folio 336: B804 en <=3/16 corta en 800 F donde B675 y B690 llegan a 900.
        assert d[(6, 64)]["max_temp_f"] == 900 and d[(6, 66)]["max_temp_f"] == 800
        assert d[(6, 66)]["values"]["300"] == 29.9
        # Folio 346: la unica fundicion del grupo, con su grado.
        assert (d[(6, 216)]["spec_no"], d[(6, 216)]["type_grade"]) == ("A494", "CX2MW")
        # La discrepancia entre ediciones queda declarada, no armonizada.
        texto = " ".join(self._tabla("table_a_1c.json")["observaciones_del_codigo"])
        assert "numeran igual" in texto and "N08367" in texto

    def test_el_uns_no_arrastra_la_clase(self):
        # El folio 336 imprime "UNS No." y "Class/Condition/Temper" en columnas
        # separadas; en 186 filas del bloque de niquel quedaron pegadas.
        import re
        for f in ("table_a_1.json", "table_a_1c.json"):
            for r in self._tabla(f)["rows"]:
                v = r.get("uns_no")
                if isinstance(v, str) and v.strip():
                    assert re.fullmatch(r"[A-Z]\d{5}", v.strip()), (f, v)

    def test_a2_sigue_intacta(self):
        # A-2 se cotejo fila a fila contra el folio 364 y no tenia ni un defecto:
        # esta prueba es la que avisara si una reextraccion futura la estropea.
        d = self._tabla("table_a_2.json")
        assert d["row_count"] == len(d["rows"]) == 24
        f = {r["spec_no"]: r for r in d["rows"]}
        assert f["A47"]["ec"] == 1.0 and f["A47"]["material_group"] == "Iron"
        assert f["A451"]["ec"] == 0.9 and f["A451"]["notes"] == "(4), (5)"
        assert f["B26, Temper F"]["ec"] == 1.0
        assert f["B367"]["material_group"] == "Titanium and Titanium Alloy"
        # Las Notas (1)..(5) y las dos GENERAL NOTES, repartidas en dos bloques.
        texto = " ".join(d["notes"])
        for marca in ("(1)", "(2)", "(3)", "(4)", "(5)", "GENERAL NOTES"):
            assert marca in texto, marca


# ---------------------------------------------------------------------------
# Factores de calidad Ec (Tabla A-2) y Ej (Tabla A-3)
# ---------------------------------------------------------------------------
# Ec y Ej multiplican directamente el esfuerzo admisible en t = PD/(2(SE+PY)).
# Lo que estas pruebas fijan no es solo el numero: es la distincion entre la
# Nota (4) -«can be enhanced by supplementary examination»- y la Nota (5)
# -«applicable only when proper supplementary examination has been performed»-,
# que dicen cosas opuestas y comparten formato.
class TestFactoresDeCalidad:

    @staticmethod
    def _res():
        from db_lib import Resources
        return Resources(str(Path(__file__).resolve().parents[3] / "resources"))

    @staticmethod
    def _tabla(archivo, sub="APPEX/appendix_a"):
        import json
        ruta = (Path(__file__).resolve().parents[3] / "resources" / "ASME B31" /
                "ASME B31.3")
        for parte in sub.split("/"):
            ruta = ruta / parte
        with open(ruta / archivo, encoding="utf-8") as fh:
            return json.load(fh)

    @staticmethod
    def _base(cual):
        import openpyxl
        from db_lib import Resources
        res = Resources(str(Path(__file__).resolve().parents[3] / "resources"))
        return B.build_factores(res, openpyxl.Workbook(), cual)

    def test_troceo_de_notas_en_orden_y_no_por_cualquier_parentesis(self):
        # El propio texto del codigo cita «para. 302.3.1(a)» y «Table A-1
        # (Table A-1C)»: trocear por cualquier «(n)» partiria las notas por la
        # mitad. Solo valen los marcadores que aparecen EN ORDEN.
        notas = B.notas_numeradas(self._tabla("table_a_2.json")["notes"])
        assert set(notas) == {"1", "2", "3", "4", "5"}
        assert notas["1"].startswith("Specifications are ASTM")
        assert "can be enhanced by supplementary examination" in notas["4"]
        assert "applicable only when proper supplementary examination" in notas["5"]
        assert "302.3.3(c)" in notas["4"]

    def test_admite_incremento_se_deriva_de_las_notas_citadas(self):
        por_spec = {}
        for x in self._base("A-2")["recs"]:
            por_spec.setdefault(x["spec"], []).append(x)
        # A395: Nota (4) -> el factor 0,80 es un minimo y puede subirse.
        a395 = por_spec["A395"][0]
        assert a395["factor"] == 0.8 and a395["admite"] == B.INC_SI
        # A426: Nota (5) -> su 1,00 YA supone el examen; no admite incremento.
        a426 = por_spec["A426"][0]
        assert a426["factor"] == 1.0 and a426["admite"] == B.INC_NO
        # A451: el codigo cita las DOS notas. No es un caso teorico: es la unica
        # fila de A-2 que lo hace, y tratarla como si solo tuviera la (5)
        # ocultaria que su 0,90 tambien puede subirse.
        a451 = por_spec["A451"][0]
        assert a451["factor"] == 0.9 and a451["admite"] == B.INC_SI_YA
        # A47: sin (4) ni (5) -> el codigo no se pronuncia.
        a47 = por_spec["A47"][0]
        assert a47["factor"] == 1.0 and a47["admite"] == B.INC_SILENCIO
        # El texto de las notas viaja con la fila, no hay que ir a buscarlo.
        assert "supplementary examination" in a395["texto"]

    def test_el_factor_ej_lo_decide_el_tipo_de_junta(self):
        por = {(x["spec"], x["desc"]): x["factor"] for x in self._base("A-3")["recs"]}
        assert por[("API 5L", "Seamless pipe")] == 1.0
        assert por[("API 5L", "Continuous welded (furnace butt welded) pipe")] == 0.6
        # Para el mismo A312 el codigo publica cuatro factores distintos.
        a312 = {d: f for (s, d), f in por.items() if s == "A312"}
        assert sorted(a312.values()) == [0.8, 0.85, 1.0, 1.0]
        assert a312["Electric fusion welded pipe, double butt seam"] == 0.85
        assert a312["Electric fusion welded pipe, single butt seam"] == 0.8

    def test_el_corte_de_pagina_no_crea_grupos_falsos(self):
        # "Copper and Copper Alloy (Cont'd)" es el mismo grupo reimpreso al
        # pasar de folio. Si sobreviviese, la cascada ofreceria dos grupos
        # donde el codigo publica uno.
        recs = self._base("A-3")["recs"]
        grupos = {x["grupo"] for x in recs}
        assert not any("Cont" in g for g in grupos)
        assert len(grupos) == 8
        # Y queda declarado fila a fila, no en silencio.
        assert any(x["detalle"] and "corte de pagina" in x["detalle"] for x in recs)

    def test_la_tabla_302_3_3_1_publica_los_seis_examenes(self):
        d = self._tabla("table_302_3_3_1.json", "CHAPTERS/tables")
        assert [r["column_2"] for r in d["rows"]] == [0.85, 0.85, 0.95, 0.9, 1.0, 1.0]
        assert d["rows"][4]["column_1"] == "(1) and (3)(a) or (3)(b)"

    def test_el_archivo_canonico_de_la_302_3_3_1_esta_declarado(self):
        # Existen table_302_3_3_1.json y table_table_302_3_3_1.json. Sin declarar
        # cual manda, el motor citaria una fuente ambigua.
        am = self._tabla("table_302_3_3_1.json",
                         "CHAPTERS/tables")["extraction_amendments"]
        assert am["archivo_canonico"] == "table_302_3_3_1.json"
        assert am["duplicado_descartado"]["archivo"] == "table_table_302_3_3_1.json"
        assert len(am["duplicado_descartado"]["sha256"]) == 64
        dup = self._tabla("table_table_302_3_3_1.json", "CHAPTERS/tables")
        assert dup["superseded_by"]["archivo"] == "table_302_3_3_1.json"

    def test_los_rotulos_de_columna_son_derivados_y_lo_dicen(self):
        # El impreso no se capturo y el PDF no esta en el repositorio: los
        # rotulos salen del para. 302.3.3(c) y llevan su procedencia. El campo
        # `header` sigue en null porque sigue siendo verdad que falta.
        cols = self._tabla("table_302_3_3_1.json", "CHAPTERS/tables")["columns"]
        assert all(c["header"] is None for c in cols)
        assert "supplementary examination" in cols[0]["header_derivado"]
        assert "Ec" in cols[1]["header_derivado"]
        for c in cols:
            assert "302.3.3(c)" in c["fuente_derivacion"]

    def test_el_incremento_de_ej_existe_y_su_tabla_se_reconstruyo(self):
        # Tercera pregunta de la Fase 0: si el codigo publica un mecanismo
        # analogo para Ej. Lo publica —para. 302.3.4(b)— y completar_tabla_302_3_4.py
        # reconstruyo el cuerpo de la Tabla 302.3.4-1 (colapsado en la
        # extraccion original) desde el folio impreso: el motor ya puede
        # ofrecer las diez filas en vez de declarar un hueco.
        par, nota1, filas = B.texto_incremento_ej(self._res())
        assert "higher joint quality factors" in par
        assert nota1 and "not permitted to increase" in nota1
        assert len(filas) == 10
        factores = [f["factor_ej"] for f in filas]
        assert factores[:5] == [0.60, 0.85, 0.80, 0.90, 1.00]
        assert sum(1 for f in filas if f.get("note")) == 2   # juntas 1 y 2

    def test_bloques_tabla_ej_no_esconde_ninguna_fila(self):
        _, _, filas = B.texto_incremento_ej(self._res())
        bloques = B.bloques_tabla_ej(filas)
        assert len(bloques) == 10
        assert all("Ej = " in t for t, _ in bloques)
        assert sum(1 for t, _ in bloques if "NO admite incremento" in t) == 2

    def test_el_conmutador_si_us_degenerado_esta_rotulado(self):
        # Regla 10: todo motor lleva conmutador SI/US. Ec y Ej son
        # adimensionales y el codigo publica una sola tabla, asi que aqui es una
        # celda fija: se cumple el espiritu sin inventar un interruptor.
        assert "ADIMENSIONAL" in B.TXT_ADIMENSIONAL
        assert "una sola tabla" in B.TXT_ADIMENSIONAL

    def test_el_contrato_de_columnas_se_mantiene(self):
        assert B.FACT_COLS[:8] == B.STRESS_COLS[:8]
        assert B.FACT_COLS[-1] == "n_pts" and B.CF["n_pts"] == B.N_IDENT_F


# ---------------------------------------------------------------------------
# El doble prefijo de CHAPTERS/tables y el hueco de la Tabla 302.3.4-1
# ---------------------------------------------------------------------------
# 66 archivos que son 32 pares, y en 11 de ellos el contenido DIFIERE. Citar "la
# Tabla X de resources/" sin decir cual de los dos archivos no es auditable.
class TestTablasCanonicas:

    RUTA = (Path(__file__).resolve().parents[3] / "resources" / "ASME B31" /
            "ASME B31.3" / "CHAPTERS" / "tables")

    @classmethod
    def _json(cls, nombre):
        import json
        with open(cls.RUTA / nombre, encoding="utf-8") as fh:
            return json.load(fh)

    @classmethod
    def _pares(cls):
        import declarar_tablas_canonicas as D
        return D.pares_de(cls.RUTA)

    def test_los_32_pares_estan_declarados(self):
        import declarar_tablas_canonicas as D
        pares = self._pares()
        assert len(pares) == 32
        for simple, doble in pares:
            dec = self._json(simple).get("canonical_declaration")
            assert dec, simple
            assert dec["archivo_canonico"] == simple
            assert dec["clase"] in D.CLASES
            assert dec["duplicado_descartado"]["archivo"] == doble
            assert len(dec["duplicado_descartado"]["sha256"]) == 64
            sup = self._json(doble).get("superseded_by")
            assert sup and sup["archivo"] == simple, doble

    def test_el_criterio_sigue_resolviendo_todos_los_pares(self):
        # Si una reextraccion futura cambia uno de los dos archivos de un par de
        # forma que el criterio mecanico ya no lo resuelva, esto lo detiene antes
        # de que nadie cite el equivocado. table_302_3_4_1.json esta EXCLUIDO a
        # proposito: su canonico se reconstruyo desde el folio impreso, no desde
        # su gemelo, asi que clasificar() nunca lo resolveria mecanicamente.
        import declarar_tablas_canonicas as D
        for simple, doble in self._pares():
            if simple in D.EXCLUIDOS_DEL_CRITERIO_MECANICO:
                assert (self._json(simple)["canonical_declaration"]["clase"]
                        == D.EXTERNA)
                continue
            clase = D.clasificar(self._json(simple), self._json(doble))
            assert clase is not None, f"{simple} vs {doble}"
            assert clase == self._json(simple)["canonical_declaration"]["clase"]

    def test_las_dos_herramientas_no_discrepan(self):
        # completar_tabla_302_3_3.py y declarar_tablas_canonicas.py afirman el
        # mismo hecho sobre el mismo par: tienen que afirmarlo igual.
        d = self._json("table_302_3_3_1.json")
        assert (d["extraction_amendments"]["archivo_canonico"]
                == d["canonical_declaration"]["archivo_canonico"]
                == "table_302_3_3_1.json")

    def test_la_302_3_4_1_ya_no_declara_hueco(self):
        # completar_tabla_302_3_4.py reconstruyo el cuerpo desde el folio
        # impreso que aporto el ingeniero. El hueco se retira del archivo -no
        # se acumula junto al arreglo- y queda la constancia de como se cerro.
        d = self._json("table_302_3_4_1.json")
        assert d.get("extraction_gap") is None, \
            "table_302_3_4_1.json sigue declarando el hueco ya cerrado"
        am = d.get("extraction_amendments")
        assert am and am["script"] == "completar_tabla_302_3_4.py"
        assert len(am["fuente"]["sha256"]) == 64

    def test_el_cuerpo_de_la_302_3_4_1_esta_reconstruido(self):
        # El sintoma que declaraba el hueco -el factor dentro del encabezado de
        # columna- ya no debe aparecer, y las 10 filas tienen que estar.
        import re
        d = self._json("table_302_3_4_1.json")
        cab = " ".join(str((c or {}).get("header") or "") for c in d["columns"])
        assert not re.search(r"Factor,\s*Ej\s*0\.\d", cab), (
            "la Tabla 302.3.4-1 volvio a colapsar el cuerpo dentro del "
            "encabezado")
        assert d["row_count"] == 10 == len(d["rows"])
        assert [r["factor_ej"] for r in d["rows"][:5]] == [0.60, 0.85, 0.80, 0.90, 1.00]
