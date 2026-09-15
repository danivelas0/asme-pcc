"""Motor DECLARADO del Art. 204 de ASME PCC-2 -- Welded Leak Box Repair.

Alcance de esta primera entrega: SOLO la rama TE + CAPS (orden del ingeniero,
2026-09-15; ver outputs/plans/plan_motor_art204.md seccion 5). La rama "codo +
tapas planas" se declara en una entrega posterior.

El Art. 204 NO publica ni una ecuacion propia (barrido completo: 106 bloques,
67 paragraph + 38 section_header + 1 figure, CERO equation, CERO table). Delega
el dimensionamiento al codigo de construccion (204-3.5/204-3.6), asi que las
filas de calculo de este motor citan el B31.3 (multi-fuente, `Cita.fuente`),
no el propio 204. Lo que el 204 SI aporta son compuertas, verificaciones,
avisos y el procedimiento -y eso si sale de su propio JSON.
"""
from __future__ import annotations

import motor_declarado as MD

FUENTE_204 = ("ASME PCC/pcc_2/p2_welded_repairs/"
              "art_204_welded_leak_box/art_204.json")
FUENTE_B313 = "ASME B31/ASME B31.3/CHAPTERS/chapter_02.json"
FUENTE_TABLA_Y = "ASME B31/ASME B31.3/CHAPTERS/tables/table_304_1_1_1.json"
FUENTE_TABLA_CIERRE = "ASME B31/ASME B31.3/CHAPTERS/tables/table_304_4_1_1.json"


def _c204(bloque, clausula):
    return MD.Cita(archivo="art_204.json", bloque=bloque, clausula=clausula)


def _cb313(bloque, clausula):
    return MD.Cita(archivo="B31.3 cap. II", bloque=bloque, clausula=clausula,
                    fuente=FUENTE_B313)


def _cY(bloque, clausula):
    return MD.Cita(archivo="B31.3 Tabla 304.1.1-1", bloque=bloque,
                    clausula=clausula, fuente=FUENTE_TABLA_Y)


def _cCierre(bloque, clausula):
    return MD.Cita(archivo="B31.3 Tabla 304.4.1-1", bloque=bloque,
                    clausula=clausula, fuente=FUENTE_TABLA_CIERRE)


# ---------------------------------------------------------------------------
# 1. Aplicacion y codigo
# ---------------------------------------------------------------------------
SEC_1 = MD.Seccion("1. APLICACION Y CODIGO", filas=(
    MD.Fila("unidad", "Sistema de unidades", tipo=MD.LISTA, ejemplo="SI",
            lista=MD.Lista(opciones=("SI", "US")),
            comentario="Conmutador SI <-> US de TODO el motor (regla 10)."),
    MD.Fila("modo", "Codigo de aplicacion", tipo=MD.LISTA, ejemplo="PCC2-204",
            lista=MD.Lista(opciones=("PCC2-204",)),
            comentario="Selector de codigo de la seccion de resolucion de "
            "material."),
))

# ---------------------------------------------------------------------------
# 2. Datos de entrada
# ---------------------------------------------------------------------------
SEC_2 = MD.Seccion("2. DATOS DE ENTRADA", filas=(
    MD.Fila("temperatura", "Temperatura de evaluacion", magnitud="temp",
            tipo=MD.ENTRADA, ejemplo=20,
            comentario="Alimenta la seccion de resolucion de material."),
    MD.Fila("tipo_caja", "Tipo de caja", tipo=MD.LISTA,
            ejemplo="No estructural",
            lista=MD.Lista(opciones=("No estructural", "Estructural"),
                           cita=_c204(7, "204-1(f)")),
            comentario="204-1(f): no estructural (contiene la fuga) o "
            "estructural (refuerza y mantiene unido el componente danado). "
            "Enumeracion cerrada que publica el propio codigo, con cita "
            "auditada (Decision 2 del plan)."),
    MD.Fila("Tp", "Espesor medido del componente portador", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=6.35,
            comentario="Dato de campo. Alimenta el aviso de pared delgada "
            "del Art. 210 (204-4.4)."),
    MD.Fila("hay_grieta", "El defecto es una grieta?", tipo=MD.ENTRADA,
            ejemplo="No",
            comentario="204-2.2: si el defecto es una grieta, la reparacion "
            "solo es elegible por una de las 4 vias (a)-(d) de abajo. "
            "Teclee Si / No."),
    MD.Fila("via_a", "(a) causa de la grieta eliminada, sin crecimiento "
            "previsto", tipo=MD.ENTRADA, ejemplo="No",
            cita=_c204(14, "204-2.2(a)")),
    MD.Fila("via_b", "(b) FFS: crecimiento aceptable en la vida de diseno",
            tipo=MD.ENTRADA, ejemplo="No", cita=_c204(15, "204-2.2(b)")),
    MD.Fila("via_c", "(c) grieta circunferencial, caja estructural disenada "
            "para separacion total", tipo=MD.ENTRADA, ejemplo="No",
            cita=_c204(16, "204-2.2(c)")),
    MD.Fila("via_d", "(d) caja encapsula por completo un vent/drain "
            "agrietado", tipo=MD.ENTRADA, ejemplo="No",
            cita=_c204(17, "204-2.2(d)")),
    MD.Fila("soldar_knuckle", "Se soldara al nudillo de un cabezal formado?",
            tipo=MD.ENTRADA, ejemplo="No",
            comentario="204-3.6: prohibido sin calificar el diseno por "
            "analisis o prueba, con fatiga incluida."),
))

# ---------------------------------------------------------------------------
# 3. Resolucion de material -- dos columnas: metal base (componente
#    portador, tuberia existente) y material de la caja (D 204-3.1: debe ser
#    compatible con el fluido, presion y temperatura).
# ---------------------------------------------------------------------------
MATERIAL = MD.Material(
    columnas=(("D", "Metal base (componente portador)"),
              ("E", "Material de la caja (envolvente)")),
    incluir_ej_ec=False,
    semilla="",
)

# ---------------------------------------------------------------------------
# 4. Envolvente: seleccion del accesorio -- rama TE + CAPS. Cascada
#    dimensional norma -> NPS -> cedula contra DB_B36_10/DB_B36_19 (regla 14),
#    de donde salen OD y espesor de pared de la te. El rating del componente
#    listado (para. 303) entra como entrada declarada, HAND-OFF a B16.9 (no
#    extraible -- ver huecos declarados): 303 autoriza calcular con 304 (la
#    Seccion 6), asi que el rating NO esta en la ruta critica.
# ---------------------------------------------------------------------------
SEC_4 = MD.Seccion("4. ENVOLVENTE - SELECCION DEL ACCESORIO (TE + CAPS)", filas=(
    MD.Fila("norma_te", "Norma dimensional de la te", tipo=MD.LISTA,
            ejemplo="B36.10M", lista=MD.Lista(cascada="B36"),
            comentario="Se elige, no se deriva del material: una te "
            "inoxidable puede pedirse a cedulas B36.10M."),
    MD.Fila("nps_te", "NPS de la te", tipo=MD.LISTA, ejemplo="2 (50)",
            lista=MD.Lista(cascada="B36")),
    MD.Fila("ced_te", "Cedula de la te", tipo=MD.LISTA, ejemplo="40 (STD)",
            lista=MD.Lista(cascada="B36")),
    MD.Fila("od_te", "OD de la te", magnitud="len", tipo=MD.FORMULA,
            comentario="Calculo: lookup INDEX/MATCH contra DB_B36_10/19 por "
            "NPS+cedula, con el conmutador SI/US (lo escribe el chasis, "
            "regla 14)."),
    MD.Fila("esp_te", "Espesor de pared de la te", magnitud="len",
            tipo=MD.FORMULA,
            comentario="Calculo: idem od_te, columna de espesor."),
    MD.Fila("rating_te", "Rating presion-temperatura del accesorio listado",
            magnitud="pres", tipo=MD.ENTRADA, ejemplo=0,
            comentario="HAND-OFF (hueco 2bis): ASME B16.9 no esta en "
            "resources/ -PDF escaneado sin capa de texto, sin OCR instalado-. "
            "Se teclea del catalogo del fabricante. para. 303 autoriza "
            "calcular con 304 (Seccion 6 de este motor), asi que este dato "
            "NO esta en la ruta critica del dictamen."),
))

# ---------------------------------------------------------------------------
# 5. Parametros de calculo
# ---------------------------------------------------------------------------
SEC_5 = MD.Seccion("5. PARAMETROS DE CALCULO", filas=(
    MD.Fila("P", "Presion de diseno", magnitud="pres", tipo=MD.ENTRADA,
            ejemplo=1.0,
            comentario="204-3.5/204-3.6: se disena para las condiciones de "
            "diseno del componente reparado, segun el codigo de construccion "
            "(B31.3, rama Te+caps)."),
    MD.Fila("CA", "Sobreespesor de corrosion", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=0, cita=_c204(44, "204-3.7")),
    MD.Fila("Y_304", "Factor Y (B31.3 Tabla 304.1.1-1)", tipo=MD.ENTRADA,
            ejemplo=0.4, cita=_cY(1, "304.1.2, Tabla 304.1.1-1"),
            comentario="Entrada citada a la tabla, no base auditada (Fase "
            "posterior segun el plan: DB_B31_Y). El ingeniero lee el valor "
            "de la fila de su material y temperatura."),
    MD.Fila("Efac", "Factor de calidad E de la conexion", tipo=MD.ENTRADA,
            ejemplo=1.0,
            comentario="HAND-OFF: se lee de Buscar_Ec_A2 / Tabla A-1 del "
            "B31.3 para el material de la caja (columna E de la seccion de "
            "material), y se teclea aqui (regla 13: un motor de calculo no "
            "lee celdas de un motor de busqueda)."),
    MD.Fila("W_junta", "Factor de reduccion de resistencia de junta soldada",
            tipo=MD.ENTRADA, ejemplo=1.0,
            comentario="para. 302.3.5(e) del B31.3. Se teclea: no hay tabla "
            "de W en resources/ para este alcance."),
))

# ---------------------------------------------------------------------------
# 6. Calculo por para. 304 -- respaldo y verificacion. Regla nº1: TODA cifra
#    normativa citada existe en B31.3 chapter_02.json (bloques confirmados por
#    lectura real: 297 = eq.(3a), 401 = eq.(6), 407 = eq.(7), 409 = eq.(8)).
# ---------------------------------------------------------------------------
SEC_6 = MD.Seccion("6. CALCULO POR PARA. 304 - RESPALDO Y VERIFICACION", filas=(
    MD.Fila("t_req_te", "Espesor requerido del cuerpo de la te",
            magnitud="len", tipo=MD.FORMULA,
            formula="={P}*{od_te}/(2*({S@E}*{Efac}*{W_junta}"
            "+{P}*{Y_304}))",
            cita=_cb313(297, "304.1.2(a), eq. (3a)"),
            comentario="t = PD/(2(SEW+PY)). S es el esfuerzo admisible del "
            "material de la caja (columna E) a la temperatura de "
            "evaluacion, resuelto por la cascada de material."),
    MD.Fila("chk_espesor_te", "Espesor real de la te vs. requerido",
            tipo=MD.FORMULA,
            # IFERROR: t_req_te propaga #N/A (a proposito, mismo patron que
            # el S_t del sleeve en Collar_PCC2_Art206) mientras la columna E
            # de material no tenga un material resuelto -no hay semilla para
            # el material de la caja, solo para el metal base-. Sin este
            # IFERROR el caso semilla dejaba un #N/A crudo en la celda de
            # veredicto, que ningun semaforo reconoce (ni CUMPLE, ni un texto
            # de aviso): verificar.py §12 lo destapo pintando gris "sin
            # calcular" en vez del rojo que promete la Verificacion.
            formula='=IFERROR(IF({esp_te}>={t_req_te}+{CA},"CUMPLE",'
            '"NO CUMPLE - 304.1.2(a): espesor de la te insuficiente"),'
            '"REVISAR - seleccione el material de la caja (seccion de '
            'resolucion de material, columna E)")',
            cita=_cb313(297, "304.1.2(a), eq. (3a)")),
    MD.Fila("d1", "Diametro de la abertura en el cierre (recortada)",
            magnitud="len", tipo=MD.ENTRADA, ejemplo=25,
            comentario="Dato de campo/diseno: diametro del corte en el cap."),
    MD.Fila("beta", "Angulo entre ejes de la derivacion y la te (grados)",
            tipo=MD.ENTRADA, ejemplo=90,
            cita=_cb313(399, "304.3.3(a)")),
    MD.Fila("t_h", "Espesor del cierre en la zona de refuerzo (hand-off)",
            magnitud="len", tipo=MD.ENTRADA, ejemplo=8,
            comentario="HAND-OFF (hueco 2bis/2): 304.4.1(b) remite a UG-32/"
            "UG-33 (Tabla 304.4.1-1) para el espesor del cap, fuera de "
            "resources/ (Sec. VIII no extraida). Se teclea del calculo "
            "externo o del catalogo del fabricante."),
    MD.Fila("A1", "Area de refuerzo requerida", tipo=MD.FORMULA,
            formula="={t_h}*{d1}*(2-SIN(RADIANS({beta})))",
            cita=_cb313(401, "304.3.3(b), eq. (6)")),
    MD.Fila("d2", "Ancho de refuerzo disponible en el cierre (medido)",
            magnitud="len", tipo=MD.ENTRADA, ejemplo=50,
            cita=_cb313(407, "304.3.3(c), eq. (7)")),
    MD.Fila("T_h_real", "Espesor real instalado del cierre", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=10,
            cita=_cb313(407, "304.3.3(c), eq. (7)")),
    MD.Fila("A2", "Area disponible en el cierre", tipo=MD.FORMULA,
            formula="=(2*{d2}-{d1})*({T_h_real}-{t_h}-{CA})",
            cita=_cb313(407, "304.3.3(c), eq. (7)")),
    MD.Fila("L4", "Longitud de refuerzo disponible de la conexion",
            magnitud="len", tipo=MD.ENTRADA, ejemplo=30,
            cita=_cb313(409, "304.3.3(c), eq. (8)")),
    MD.Fila("A3", "Area disponible en la conexion (te)", tipo=MD.FORMULA,
            formula="=2*{L4}*({esp_te}-{t_req_te}-{CA})/SIN(RADIANS({beta}))",
            cita=_cb313(409, "304.3.3(c), eq. (8)")),
    MD.Fila("chk_refuerzo", "Area disponible (A2+A3) vs. requerida (A1)",
            tipo=MD.FORMULA,
            # IFERROR: A3 hereda el #N/A de t_req_te sin material resuelto en
            # la columna E -mismo motivo que en chk_espesor_te-.
            formula='=IFERROR(IF({A2}+{A3}>={A1},"CUMPLE",'
            '"NO CUMPLE - 304.4.2(e)/304.3.3(a): refuerzo insuficiente en '
            'la abertura del cierre"),'
            '"REVISAR - seleccione el material de la caja (seccion de '
            'resolucion de material, columna E)")',
            cita=_cb313(497, "304.4.2(e) -> 304.3.3")),
))

# ---------------------------------------------------------------------------
# 7. Verificaciones y avisos
# ---------------------------------------------------------------------------
SEC_7 = MD.Seccion("7. VERIFICACIONES Y AVISOS", filas=(
    MD.Fila("elegibilidad_grieta", "Elegibilidad por grieta (204-2.2)",
            tipo=MD.FORMULA,
            formula='=IF({hay_grieta}<>"Si","CUMPLE",'
            'IF(OR({via_a}="Si",{via_b}="Si",{via_c}="Si",{via_d}="Si"),'
            '"CUMPLE","NO ELEGIBLE - GRIETA: ninguna via (a)-(d) de 204-2.2 '
            'aplica"))',
            cita=_c204(13, "204-2.2")),
    MD.Fila("knuckle", "Compuerta: soldadura al nudillo sin calificar",
            tipo=MD.FORMULA,
            formula='=IF({soldar_knuckle}="Si",'
            '"NO PERMITIDO - 204-3.6: no soldar al nudillo del cabezal '
            'formado sin calificar el diseno por analisis o prueba",'
            '"CUMPLE")',
            cita=_c204(42, "204-3.6")),
    MD.Fila("aviso_pared_delgada", "Aviso: pared delgada (Art. 210)",
            tipo=MD.FORMULA,
            formula='=IF({Tp}<={UMBRAL:210_pared_delgada},'
            '"AVISO: pared delgada - extremar precauciones de soldadura en '
            'servicio (204-4.4, remite al Art. 210)","-")',
            cita=_c204(79, "204-4.4")),
    MD.Fila("aviso_reducir_aporte", "Aviso: considerar reducir aporte "
            "termico", tipo=MD.FORMULA,
            formula='=IF({Tp}<={UMBRAL:210_pared_reducir_aporte},'
            '"AVISO: pared por debajo del umbral - puede hacer falta bajar '
            'el aporte termico (204-4.4, remite al Art. 210)","-")',
            cita=_c204(79, "204-4.4")),
    MD.Fila("empuje_axial", "Empuje axial por separacion circunferencial "
            "total", tipo=MD.ENTRADA, ejemplo="No",
            comentario="204-3.9(a): considere el empuje axial, salvo que la "
            "resistencia remanente al fin de vida baste para renunciar. "
            "Teclee Si cuando este considerado o justificada la renuncia."),
    MD.Fila("chk_empuje_axial", "Verificacion: empuje axial considerado",
            tipo=MD.FORMULA,
            formula='=IF({empuje_axial}="Si","CUMPLE",'
            '"REVISAR - 204-3.9(a): considere el empuje axial por '
            'separacion circunferencial total o justifique la renuncia. '
            'HAND-OFF: sin ecuacion publicada para la resistencia del '
            'filete perimetral frente a este empuje (hueco 3)")',
            cita=_c204(53, "204-3.9(a)")),
    MD.Fila("colapso_externo", "Colapso por presion externa (sellante / "
            "prueba)", tipo=MD.ENTRADA, ejemplo="No",
            comentario="204-3.12/204-6.3: verificado externamente contra "
            "UG-28..UG-30 (Sec. VIII, fuera de resources/). Teclee Si "
            "cuando ese calculo externo este hecho."),
    MD.Fila("chk_colapso_externo", "Verificacion: colapso externo "
            "verificado", tipo=MD.FORMULA,
            formula='=IF({colapso_externo}="Si","CUMPLE",'
            '"REVISAR - 204-3.12/204-6.3: verifique el colapso por presion '
            'externa del componente reparado (UG-28..UG-30, hueco 4)")',
            cita=_c204(63, "204-3.12")),
    MD.Fila("metodo_prueba", "Metodo de prueba adoptado", tipo=MD.LISTA,
            ejemplo="Prueba hidrostatica",
            lista=MD.Lista(opciones=(
                "Prueba de fuga en servicio", "Prueba hidrostatica",
                "Prueba neumatica", "Prueba de fuga sensible"),
                cita=_c204(99, "204-6.2")),
            comentario="204-6.2(a)-(d): los cuatro metodos que el codigo "
            "admite, elegidos segun riesgo (Decision 2 del plan)."),
))

# ---------------------------------------------------------------------------
# 8. Dictamen
# ---------------------------------------------------------------------------
VERIFICACIONES = (
    MD.Verificacion(clave="elegibilidad_grieta", rotulo="Elegibilidad por "
                    "grieta", requerido="", adoptado="",
                    criterio="una de las 4 vias de 204-2.2 aplica",
                    favorables=("CUMPLE",), avisos=(),
                    cita=_c204(13, "204-2.2")),
    MD.Verificacion(clave="knuckle", rotulo="No soldar al nudillo sin "
                    "calificar", requerido="", adoptado="",
                    criterio="no se suelda al nudillo sin calificacion",
                    favorables=("CUMPLE",), avisos=(),
                    cita=_c204(42, "204-3.6")),
    MD.Verificacion(clave="chk_espesor_te", rotulo="Espesor de la te",
                    requerido="{t_req_te}", adoptado="{esp_te}",
                    criterio="esp_te >= t_req_te + CA",
                    favorables=("CUMPLE",), avisos=(),
                    cita=_cb313(297, "304.1.2(a), eq. (3a)")),
    MD.Verificacion(clave="chk_refuerzo", rotulo="Refuerzo de la abertura",
                    requerido="{A1}", adoptado="{A2}+{A3}",
                    criterio="A2 + A3 >= A1",
                    favorables=("CUMPLE",), avisos=(),
                    cita=_cb313(497, "304.4.2(e) -> 304.3.3")),
    MD.Verificacion(clave="chk_empuje_axial", rotulo="Empuje axial",
                    requerido="", adoptado="",
                    criterio="empuje considerado o renuncia justificada",
                    favorables=("CUMPLE",), avisos=(),
                    cita=_c204(53, "204-3.9(a)")),
    MD.Verificacion(clave="chk_colapso_externo", rotulo="Colapso externo",
                    requerido="", adoptado="",
                    criterio="colapso por presion externa verificado",
                    favorables=("CUMPLE",), avisos=(),
                    cita=_c204(63, "204-3.12")),
)

FILA_DICTAMEN = MD.Fila(
    "dictamen", "Dictamen global del articulo", tipo=MD.FORMULA,
    cita=_c204(39, "204-3.5/204-3.6"),
    comentario="Calculo: DICTAMEN GLOBAL. Lo compone el chasis desde las "
    "compuertas y verificaciones declaradas.")

DICTAMEN = MD.Dictamen(
    compuertas=("elegibilidad_grieta", "knuckle"),
    verificaciones=("chk_espesor_te", "chk_refuerzo", "chk_empuje_axial",
                    "chk_colapso_externo"),
)

SEC_8 = MD.Seccion("8. DICTAMEN", filas=(FILA_DICTAMEN,))

# ---------------------------------------------------------------------------
# 9. Pasos del flujo (anexo, B.4 del plan)
# ---------------------------------------------------------------------------
def _paso(numero, titulo, clausula, clave, rotulo_check):
    return MD.Paso(numero=numero, titulo=titulo, clausula=clausula, filas=(
        MD.Fila(clave, rotulo_check, tipo=MD.ENTRADA, ejemplo="No"),
    ))

PASOS = (
    _paso(1, "PELIGROS Y ELEGIBILIDAD", "204-2.2, 204-2.4", "p1_ok",
          "Revision de peligros hecha y elegibilidad confirmada?"),
    _paso(2, "CARACTERIZACION DEL COMPONENTE Y DEL DEFECTO", "204-3.11",
          "p2_ok", "Longitud a metal sano y espesor actual medidos en "
          "campo?"),
    _paso(3, "MATERIAL DE LA CAJA", "204-3.1, 204-3.4", "p3_ok",
          "Material compatible y con tenacidad verificada?"),
    _paso(4, "DISENO: CARGAS Y ESPESOR", "204-3.5 a 204-3.9", "p4_ok",
          "Espesor, refuerzo y cargas transitorias verificados?"),
    _paso(5, "VENTS, DRENAJES Y SELLANTE", "204-3.10, 204-3.12, 204-3.13",
          "p5_ok", "Vents/drenajes dispuestos y colapso por sellante "
          "considerado?"),
    _paso(6, "PREPARACION E INSTALACION", "204-4.1, 204-4.2", "p6_ok",
          "Superficie preparada e izado/instalacion planificados?"),
    _paso(7, "SOLDADURA, EN SERVICIO Y PWHT", "204-4.3 a 204-4.6", "p7_ok",
          "WPS/soldador calificados y PWHT conforme al codigo?"),
    _paso(8, "EXAMEN Y PRUEBA", "204-5, 204-6", "p8_ok",
          "NDE y prueba de presion/fuga completadas?"),
)

# ---------------------------------------------------------------------------
# 10. Especificaciones tecnicas (pestana propia)
# ---------------------------------------------------------------------------
ESPECIFICACIONES = (
    ("DISENO", (
        MD.Especificacion(
            concepto="Codigo de construccion aplicable",
            texto="204-3.5/204-3.6: la caja y sus soldaduras se disenan "
            "para las condiciones de diseno del componente reparado, "
            "siguiendo el codigo de construccion o post-construccion "
            "aplicable (aqui, B31.3, rama Te + caps).",
            clausula="204-3.5, 204-3.6", editable=False,
            cita=_c204(39, "204-3.5")),
        MD.Especificacion(
            concepto="Espesor del cierre (cap)",
            texto="Cierres no cubiertos por para. 303: se disenan segun "
            "ASME BPVC Sec. VIII Div. 1 (eq. 13, t_m = t + c), remitido por "
            "la Tabla 304.4.1-1 al parrafo de UG-32/UG-33 segun la forma "
            "del cierre (p.ej. UG-32(d)/UG-33(d) para un cabezal "
            "elipsoidal). HAND-OFF: Sec. VIII no esta en resources/.",
            clausula="304.4.1(b), Tabla 304.4.1-1", editable=False,
            cita=_cCierre(1, "Tabla 304.4.1-1 - Ellipsoidal")),
        MD.Especificacion(
            concepto="Longitud minima de la caja",
            texto="204-3.11: la caja debe extenderse a una zona sana del "
            "componente reparado, sin cuantificar. Se verifica en campo "
            "contra la distancia a metal sano medida; no se emite una "
            "formula de longitud de atenuacion (hueco 5).",
            clausula="204-3.11", editable=False,
            cita=_c204(61, "204-3.11")),
        MD.Especificacion(
            concepto="Rating del accesorio listado (ASME B16.9)",
            texto="para. 303: un componente listado en la Tabla 326.1.1-1 "
            "se considera apto por su rating presion-temperatura. B16.9 no "
            "esta en resources/ (PDF escaneado sin capa de texto, sin OCR "
            "instalado): el rating se lee del catalogo del fabricante "
            "(hueco 2bis). No esta en la ruta critica: 303 autoriza "
            "calcular con 304 (Seccion 6 de este motor).",
            clausula="para. 303", editable=False,
            cita=_c204(39, "204-3.5")),
        MD.Especificacion(
            concepto="Figura 204-1-1",
            texto="Fotografia de una caja soldada sobre una te, sin acotar. "
            "No publica ningun dato dimensional (hueco 6).",
            clausula="Fig. 204-1-1", editable=False,
            cita=_c204(11, "Fig. 204-1-1")),
    )),
    ("FABRICACION", (
        MD.Especificacion(
            concepto="Preparacion de la superficie",
            texto="El componente al que se suelda la caja debe estar libre "
            "de depositos de corrosion sueltos, suciedad, pintura, "
            "aislamiento, mastics y otros recubrimientos en la vicinidad de "
            "la soldadura.",
            clausula="204-4.1", editable=False,
            cita=_c204(71, "204-4.1")),
        MD.Especificacion(
            concepto="Soldadura en servicio",
            texto="Instalacion y soldadura fuera o en servicio; para "
            "soldadura en servicio, la calificacion del procedimiento debe "
            "atender temperatura de precalentamiento, velocidad de "
            "enfriamiento y riesgo de burn-through (ASME PCC-2 Art. 210). "
            "Ver avisos de pared delgada en la Seccion 7.",
            clausula="204-4.4", editable=False,
            cita=_c204(79, "204-4.4")),
        MD.Especificacion(
            concepto="PWHT",
            texto="El precalentamiento y el tratamiento termico "
            "postsoldadura deben cumplir el codigo de construccion o "
            "post-construccion aplicable, salvo que una evaluacion de "
            "aptitud para el servicio justifique una desviacion.",
            clausula="204-4.6", editable=False,
            cita=_c204(83, "204-4.6")),
    )),
    ("EXAMEN Y PRUEBA", (
        MD.Especificacion(
            concepto="Criterio de omision de NDE de penetracion total",
            texto="Puede omitirse la confirmacion de penetracion total si "
            "se justifica por las condiciones de servicio, por ejemplo con "
            "esfuerzos calculados bajos (menos de la mitad del esfuerzo "
            "admisible de diseno a la temperatura de operacion) o bajo "
            "riesgo de corrosion en resquicio en la junta caja-componente.",
            clausula="204-5.2", editable=False,
            cita=_c204(88, "204-5.2")),
        MD.Especificacion(
            concepto="Prueba de presion o fuga",
            texto="El tipo de prueba se determina por riesgo (probabilidad "
            "y consecuencia de fallo). Se debe considerar el potencial de "
            "colapso por presion externa de la tuberia portadora al "
            "especificar la presion de prueba (ver Seccion 7).",
            clausula="204-6.1, 204-6.3", editable=False,
            cita=_c204(97, "204-6.1")),
    )),
)

# ---------------------------------------------------------------------------
# 11. El Motor completo
# ---------------------------------------------------------------------------
MOTOR_ART_204 = MD.Motor(
    articulo="204",
    hoja="Caja_PCC2_Art204",
    titulo="MOTOR DE CALCULO -- ASME PCC-2 ART. 204 (WELDED LEAK BOX REPAIR)",
    fuente=FUENTE_204,
    corto="CAJA DE FUGA SOLDADA",
    descripcion="Reparacion por caja de fuga soldada (leak box) sobre un "
    "componente a presion, rama Te + caps.",
    alcance="Art. 204 delega el dimensionamiento al codigo de construccion "
    "aplicable (B31.3, rama Te + caps). No cubre la rama codo + tapas "
    "planas ni componentes bajo BPVC Sec. VIII (bloqueado y declarado).",
    clausulas_espec="204-3.5 · 204-3.6 · 304.4.1 · 304.4.2 · 204-3.11",
    aplicacion=True,
    casos=("Diseno",),
    secciones=(SEC_1, SEC_2, SEC_4, SEC_5, SEC_6, SEC_7, SEC_8),
    dimensionales=(
        MD.Dimensional(norma="norma_te", nps="nps_te", cedula="ced_te",
                       od="od_te", espesor="esp_te"),
    ),
    material=MATERIAL,
    verificaciones=VERIFICACIONES,
    dictamen=DICTAMEN,
    pasos=PASOS,
    especificaciones=ESPECIFICACIONES,
)
