"""Plantilla de declaracion de un motor PCC-2, para la skill `motor_pcc2`.

ADVERTENCIA — esto es MATERIAL DE REFERENCIA, no un motor real:

- El motor que de verdad rige el Art. 206 (`Collar_PCC2_Art206`) sigue escrito a
  mano en `build_collar_art206()`, dentro de
  `outputs/Base_Datos_Materiales_ASME/scripts/build_db_materiales.py`. Es lo unico
  de ese articulo verificado celda a celda contra su *oracle* y contra
  `verificar.py` §6f. Este archivo NO lo sustituye, no lo migra y no se importa
  desde el builder.
- Las citas de este ejemplo (`Cita(archivo=..., bloque=N, clausula=...)`) SI son
  reales: se verificaron leyendo
  `resources/ASME PCC/pcc_2/p2_welded_repairs/art_206_full_encirclement_steel/art_206.json`
  con Python antes de escribir este archivo (Regla n.1: ni siquiera un ejemplo de
  referencia inventa un numero de bloque). El bloque al final de este archivo lo
  vuelve a comprobar en caliente si lo ejecutas.
- El articulo de ejemplo NO reproduce el Art. 206 completo (~45 filas reales,
  conmutador SI/US con seis bases, cascada de material, anexo de 8 pasos...): es
  un recorte pequeno pensado para enseniar la FORMA de una declaracion, no para
  ser copiado literal a `motores/art_XXX.py`.

Como usar esto en la Fase 2 (Declarar) de la skill:
  1. Copia la forma (Motor / Seccion / Fila / Verificacion / Dictamen / Paso /
     Especificacion), nunca los valores.
  2. Cambia cada `Cita` por bloques REALES del `art_XXX.json` del articulo que
     estas declarando -leidos en la Fase 1, confirmados por el ingeniero-.
  3. No declares una Fila de tipo FORMULA sin su Cita: `comprobar_procedencia()`
     la rechaza con SystemExit antes de escribir una sola celda (ver mas abajo).
"""
from __future__ import annotations

# En un motor real, este archivo vive dentro de `motores/`, al lado de
# `motor_declarado.py` (un nivel arriba), y `import motor_declarado` funciona
# sin nada especial porque el builder ya corre con `scripts/` en `sys.path`.
# Aqui, como este archivo vive en la skill y no en `motores/`, se agrega esa
# ruta a mano SOLO para poder ejecutar este ejemplo de forma aislada
# (ver el bloque de autocomprobacion, al final).
try:
    import motor_declarado as MD
except ModuleNotFoundError:
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[4] /
                           "outputs" / "Base_Datos_Materiales_ASME" / "scripts"))
    import motor_declarado as MD


# ---------------------------------------------------------------------------
# 0. La fuente. UN archivo de resources/ por motor -- todas las Cita de este
#    Motor describen bloques de ESTE mismo JSON. `comprobar_procedencia()` no
#    admite un archivo distinto por fila: el articulo entero vive en un solo
#    documento de resources/, y mezclar fuentes dentro de un mismo motor
#    perderia la propiedad de "un motor, un JSON, un `--pdfs` para auditarlo".
# ---------------------------------------------------------------------------
FUENTE_206 = ("ASME PCC/pcc_2/p2_welded_repairs/"
              "art_206_full_encirclement_steel/art_206.json")


def _cita(bloque, clausula, comentario=""):
    """Atajo para no repetir `archivo=FUENTE_206` en cada Cita del ejemplo.

    El campo `Cita.archivo` es SOLO lo que ve el ingeniero en la columna de
    referencia (regla del proyecto: la procedencia se escribe en tres formas,
    y esta es una de ellas). La ruta que de verdad valida
    `comprobar_procedencia()` contra el disco es siempre `motor.fuente`, nunca
    `Cita.archivo` -- por eso aqui basta con un nombre corto y legible.
    """
    return MD.Cita(archivo="art_206.json", bloque=bloque, clausula=clausula)


# ---------------------------------------------------------------------------
# 1. Filas reservadas (CLAVES_RESERVADAS en motor_declarado.py): "unidad",
#    "modo", "temperatura", "dictamen". Un motor con `motor.material is not
#    None` los exige los cuatro -- build_motor_declarado() aborta con un
#    SystemExit legible si falta alguno, citando que pase lo necesitaba.
#    Ojo con los defectos (corregido 2026-09-13: aqui se decia que
#    `material is not None` y `aplicacion=True` eran "los dos valores por
#    defecto", y el de `material` es None): `Motor.material` vale None por
#    defecto -un motor sin cascada de material no exige `modo` ni
#    `temperatura`- y `Motor.aplicacion` vale True. Y desde 2026-09-13
#    `aplicacion=False` YA NO exime de declarar `modo` cuando hay material:
#    antes se pasaba el literal "$D$11" -el default de los dos motores
#    escritos a mano-, que en una hoja declarada es una fila cualquiera.
#
#    NOTA verificada contra el codigo: `MD.LISTA` existe como tipo, pero HOY
#    ningun pase de build_db_materiales.py lo distingue de `MD.ENTRADA` --las
#    dos llegan igual a `helpers.entrada()` dentro de `emitir_tabla()`--. Una
#    Fila `tipo=MD.LISTA` documenta la intencion (esto se elige de una lista,
#    regla 12/14 del catalogo de reglas del libro) pero la `DataValidation`
#    real (el desplegable con SI/US, o con los codigos de aplicacion) NO se
#    escribe sola: hay que anadirla aparte, con el mismo mecanismo que ya usan
#    los dos motores a mano, hasta que un pase transversal la automatice.
# ---------------------------------------------------------------------------
FILAS_APLICACION = MD.Seccion("1. APLICACION Y CODIGO", filas=(
    MD.Fila("unidad", "Sistema de unidades", tipo=MD.LISTA, ejemplo="SI",
            comentario="Conmutador SI <-> US de TODO el motor (regla 10 del "
            "catalogo). Ver la nota de arriba: hoy esto no trae su "
            "DataValidation puesta sola."),
    MD.Fila("modo", "Codigo de aplicacion", tipo=MD.LISTA, ejemplo="PCC2-206",
            comentario="Selector de codigo de la seccion de resolucion de "
            "material (regla 12/14: material siempre de una base, nunca "
            "tecleado)."),
))

FILAS_ENTRADA = MD.Seccion("2. DATOS DE ENTRADA", filas=(
    MD.Fila("temperatura", "Temperatura de evaluacion", magnitud="temp",
            tipo=MD.ENTRADA, ejemplo=20,
            comentario="Alimenta la seccion de resolucion de material "
            "(temp_fuente_cell de construir_seccion7_material)."),
    MD.Fila("Tp", "Espesor nominal de la tuberia portadora", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=6.35),
    MD.Fila("Ts", "Espesor del sleeve", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=9.5),
    MD.Fila("G", "Luz radial de ajuste (fit-up gap)", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=1.5,
            comentario=_cita(63, "206-4.1").clausula + ": el codigo exige "
            "cerrar esta luz antes de soldar; se teclea porque es un dato de "
            "campo, no de tabla (regla 14)."),
    MD.Fila("w_adoptado", "Cateto de filete adoptado en el diseno",
            magnitud="len", tipo=MD.ENTRADA, ejemplo=9.5),
))


# ---------------------------------------------------------------------------
# 2. Filas de calculo. Las dos primeras son el caso real de "ecuacion vacia
#    con imagen" (ver el diseno de la skill y reglas_del_libro.md): los
#    bloques 72 y 77 de art_206.json llegaron con `text` inservible y la
#    formula solo en `fig/fig_206_3_5_1.png` / `fig/fig_206_3_5_2.png`. Se
#    leyeron esos PNG con la herramienta Read -NUNCA un PDF externo, NUNCA
#    memoria del modelo- y el texto recuperado quedo grabado en
#    `extraction_amendments` del propio JSON (script, fecha, SHA-256 de la
#    imagen fuente). Esta Cita apunta exactamente a esos dos bloques.
# ---------------------------------------------------------------------------
FILAS_CALCULO = MD.Seccion("3. CALCULO DEL CATETO DE FILETE (206-3.5)", filas=(
    MD.Fila("w_delgado", "Cateto - Ts <= 1.4 x Tp (filete completo)",
            magnitud="len", tipo=MD.FORMULA, formula="={Ts}+{G}",
            cita=_cita(72, "206-3.5(a), Fig. 206-3.5-1"),
            comentario="Texto recuperado de fig/fig_206_3_5_1.png: "
            "'w = Ts + G'. Bloque 72 de art_206.json."),
    MD.Fila("w_grueso", "Cateto maximo - Ts > 1.4 x Tp (chaflan opcional)",
            magnitud="len", tipo=MD.FORMULA, formula="=1.4*{Tp}+{G}",
            cita=_cita(77, "206-3.5(b), Fig. 206-3.5-2"),
            comentario="Texto recuperado de fig/fig_206_3_5_2.png: "
            "'w_max = 1.4 Tp + G'. Bloque 77 de art_206.json."),
    MD.Fila("w_requerido", "Cateto de filete requerido",
            magnitud="len", tipo=MD.FORMULA,
            formula="=IF({Ts}<=1.4*{Tp},{w_delgado},{w_grueso})",
            cita=_cita(43, "206-3.5"),
            comentario="La eleccion no es una tercera ecuacion: es la propia "
            "estructura de 206-3.5. El bloque 43 imprime solo la frase que "
            "introduce la eleccion ('...shall be as follows:'); las dos "
            "condiciones que la resuelven -- (a) Ts<=1.4 Tp, (b) Ts>1.4 Tp -- "
            "estan en los bloques 44 y 45, no en el 43 (corregido: Ronda de "
            "arreglo 1, Hallazgo 3 -- la cita al 43 es valida, pero no hay "
            "que darle a entender que el (a)/(b) vive ahi)."),
    MD.Fila("chk_cateto", "Verificacion: cateto adoptado >= requerido",
            tipo=MD.FORMULA,
            formula='=IF({w_adoptado}>={w_requerido},"CUMPLE","NO CUMPLE")',
            cita=_cita(43, "206-3.5")),
    # Umbral de DOBLE UNIDAD, el segundo de los tres casos costosos: el
    # codigo imprime "100 mm (4 in.)" en el bloque 40. Las dos cifras se
    # guardan TAL CUAL -nunca 100/25.4 ni 4*25.4-, porque la relacion exacta
    # entre las dos es del codigo, no una conversion a recalcular. El motor
    # real resuelve este patron con `leer_umbrales_pcc2()` (build_db_materiales.py):
    # lee las DOS cifras impresas junto a un ancla de texto y ABORTA el build
    # si el codigo deja de imprimir el par. Aqui, para el ejemplo, se deja
    # solo la SI: un motor real necesitaria una segunda Fila (o una FORMULA
    # que elija segun {unidad}) para la cifra US, nunca una division.
    MD.Fila("L_min_SI", "Longitud minima del sleeve (SI)", magnitud="len",
            tipo=MD.FORMULA, formula="=100",
            cita=_cita(40, "206-3.4"),
            comentario="Codigo: '100 mm (4 in.)'. La cifra US (4 in.) es un "
            "dato INDEPENDIENTE, no 100/25.4 -- ver leer_umbrales_pcc2()."),
))


# ---------------------------------------------------------------------------
# 3. Verificaciones y Dictamen.
#
#    NOTA verificada contra el codigo: de los cinco campos de `Verificacion`
#    (`clave`, `rotulo`, `requerido`, `adoptado`, `criterio`, mas `favorables`/
#    `avisos`/`cita`), HOY solo `clave`/`favorables`/`avisos` los lee
#    `_semaforo_de()` en build_db_materiales.py -- es lo unico que pinta el
#    semaforo. `rotulo`/`requerido`/`adoptado`/`criterio`/`cita` se declaran
#    para documentar la verificacion (y para que una tarea futura los
#    consuma), pero si quieres que el ingeniero VEA el valor requerido y el
#    adoptado en la hoja, tienen que existir tambien como Fila propia dentro
#    de alguna Seccion -como "w_requerido" y "w_adoptado" arriba-, tal como
#    hacen el 212 y el 206 escritos a mano.
#    Desde 2026-09-13, lo que SI hace el chasis con esos dos campos es
#    comprobar que sus {nombres} existan como fila del motor
#    (`comprobar_nombres_declarativos()`): asi, el dia que alguien renombre
#    `w_requerido`, la declaracion aborta en vez de quedarse describiendo una
#    fila que ya no esta. La llave sigue permitida; lo que no hay es emision.
#
#    NOTA verificada contra el codigo: `Dictamen.compuertas` no lo lee ningun
#    pase todavia (es "tarea de quien componga la formula del dictamen a
#    partir de ellas, todavia sin escribir" -- comentario textual de
#    `_semaforo_de()`). `Dictamen.verificaciones` SI lo lee `verificar.py`
#    §12, punto 4: recalcula cada celda que nombra y exige que el texto del
#    dictamen sea "APTO" si y solo si TODAS dan "CUMPLE". Pero la FORMULA de
#    la fila reservada "dictamen" no la compone ningun pase: hay que
#    escribirla a mano, como se ve abajo, y que sea CONSISTENTE con lo que
#    `Dictamen.verificaciones` promete -si diverge, verificar.py §12 lo
#    encuentra, no lo previene.
# ---------------------------------------------------------------------------
VERIFICACIONES = (
    MD.Verificacion(
        clave="chk_cateto", rotulo="Cateto de filete (206-3.5)",
        requerido="{w_requerido}", adoptado="{w_adoptado}",
        criterio="w_adoptado >= w_requerido",
        favorables=("CUMPLE",), avisos=(),
        cita=_cita(43, "206-3.5")),
)

FILA_DICTAMEN = MD.Fila(
    "dictamen", "Dictamen global del articulo (ejemplo)", tipo=MD.FORMULA,
    formula='=IF({chk_cateto}="CUMPLE","APTO","REVISAR")',
    cita=_cita(43, "206-3.5"),
    comentario="Tiene que ser CONSISTENTE con Dictamen.verificaciones de "
    "abajo: si aqui se anade una verificacion nueva al AND, hay que "
    "anadirla tambien alli, o verificar.py §12 (punto 4) reporta la "
    "divergencia.")

DICTAMEN_EJEMPLO = MD.Dictamen(
    compuertas=(),                    # sin compuertas en este ejemplo
    verificaciones=("chk_cateto",),   # lo que verificar.py §12 recalcula
)

FILAS_VEREDICTO = MD.Seccion("4. DICTAMEN", filas=(FILA_DICTAMEN,))


# ---------------------------------------------------------------------------
# 4. El anexo de pasos del flujo (Paso). Se emite DESPUES de la ultima
#    Seccion -nunca intercalado-, con su propia banda por paso
#    ("PASO N . TITULO - clausula", ver `_titulo_paso()` en
#    motor_declarado.py). Un Paso puede llevar sus propias Filas de calculo,
#    con la MISMA exigencia de Cita que cualquier Seccion.
# ---------------------------------------------------------------------------
PASO_1 = MD.Paso(
    numero=1, titulo="INSTALACION", clausula="206-4.1",
    filas=(
        MD.Fila("gap_ok", "Luz radial dentro de tolerancia?", tipo=MD.FORMULA,
                formula='=IF({G}<=2.5,"CUMPLE","REVISAR")',
                # Bloque 63, NO 62: el 62 es el section_header "206-4.1
                # Installation" -un rotulo, que no publica ningun valor- y la
                # luz radial de "2.5 mm (3/32 in.) maximum" la imprime el
                # parrafo 63. Desde 2026-09-13 comprobar_procedencia() rechaza
                # una Cita a un bloque section_header por este mismo caso.
                cita=_cita(63, "206-4.1")),
    ))


# ---------------------------------------------------------------------------
# 5. Especificaciones tecnicas de la pestana propia del articulo (Fase 5,
#    ya automatica via build_documentos_declarados: esta tupla es todo lo
#    que hay que declarar, la pestana la construye el chasis).
#
#    `Especificacion.texto` admite {nombres} de verdad desde 2026-09-13:
#    `MD.texto_de_especificacion()` los sustituye por la direccion CALIFICADA
#    CON LA HOJA del motor ("='Collar_PCC2_Art206'!$D$40"), porque la pestana
#    es otra hoja y un "$D$40" a secas apuntaria dentro de ella. Un nombre que
#    no existe aborta. Antes el texto se copiaba CRUDO y un "{t_req}" salia
#    impreso con las llaves.
# ---------------------------------------------------------------------------
ESPECIFICACIONES_EJEMPLO = (
    ("FABRICACION", (
        MD.Especificacion(
            concepto="Longitud minima del sleeve",
            texto="100 mm (4 in.), y sobrepasar el defecto al menos 50 mm "
            "(2 in.)",
            clausula="206-3.4", editable=False,
            cita=_cita(40, "206-3.4")),
    )),
)


# ---------------------------------------------------------------------------
# 6. El Material. Se declara aparte de `secciones` -el chasis inserta su
#    propio bloque DESPUES de escribir todas las Secciones y Pasos del motor
#    (ver build_motor_declarado(): `construir_seccion7_material()` se llama
#    con `res["ultima_fila"] + 2`, es decir, al final de la hoja, no en medio).
#
#    LIMITE VERIFICADO CONTRA EL CODIGO, no supuesto: `resolver_direcciones()`
#    solo indexa `motor.secciones` y `motor.pasos` (via `_bloques()`). Las
#    celdas que arma `construir_seccion7_material()` -material_id, S(T)
#    resuelto, Temp. max., dictamen de rango- NO entran a ese indice, asi que
#    NINGUNA Fila de una Seccion declarada puede citarlas por nombre (un
#    `{s_admisible}` en una formula fallaria con "no existe como fila del
#    motor", el mismo guardia de `sustituir_nombres()`). Si el articulo que
#    estas declarando necesita el esfuerzo admisible del material dentro de
#    una ecuacion (el caso tipico de un t_req = PD/(2(SE+PY))), hoy eso es
#    una extension del chasis que todavia no existe -- se declara el hueco al
#    ingeniero en la Fase 1/2 en vez de fingir que {S} funciona.
#
#    Este ejemplo evita el problema a proposito: el calculo del cateto de
#    filete (206-3.5) no necesita el esfuerzo admisible del material, asi que
#    `material=` de abajo se deja solo como referencia de como se declara.
# ---------------------------------------------------------------------------
MATERIAL_EJEMPLO = MD.Material(
    columnas=(("D", "Carrier pipe / sleeve"),),
    incluir_ej_ec=False,
    semilla="",     # "" = sin caso semilla precargado en la cascada
)


# ---------------------------------------------------------------------------
# 7. El Motor completo.
#
#    LIMITE #4 VERIFICADO CONTRA EL CODIGO (Ronda de arreglo 1, Hallazgo 2) --
#    `casos=` ABAJO NO HACE LO QUE PARECE. El diseno lo describe como "1 o 2
#    columnas de caso" (Operacion / Diseno, el patron heredado del 212 y del
#    206 escritos a mano, donde Operacion vive en D y Diseno en E). El chasis
#    NO reproduce eso: `emitir_tabla()` fija la columna de valor UNA SOLA VEZ,
#    fuera del bucle de filas (`col = chr(ord("A") + COL_PRIMER_CASO - 1)` en
#    motor_declarado.py), y escribe TODA fila en esa unica columna (D). En
#    todo el repositorio, `motor.casos` se lee en un solo sitio
#    (`_reglas_de_comentario_de()`, build_db_materiales.py:10061) y solo para
#    acotar en cuantas columnas se ADMITE un comentario -no para generar
#    ninguna columna de valor nueva-. Comprobado ejecutando `emitir_tabla()`
#    sobre este mismo `MOTOR_206_EJEMPLO` con `casos=("Operacion","Diseno")`:
#    las 23 filas de la hoja resultante tienen su valor en D: la columna E
#    queda vacia en las 23.
#
#    Esto NO es un defecto a arreglar aqui: es una decision del chasis
#    -`resolver_direcciones()` devuelve UNA direccion por clave, y de eso
#    dependen `sustituir_nombres()`, las pruebas del chasis y el semaforo de
#    la §12-, pero declarar `casos=("Operacion","Diseno")` como si el motor
#    fuera a evaluar las dos presiones en columnas separadas -el patron que
#    cualquiera que conozca el 212/206 esperaria- es enganoso: no pasa nada
#    de eso. Un articulo que de verdad necesite dos columnas de caso hoy
#    tiene que declarar dos Fila independientes por caso (p. ej. "P_oper" y
#    "P_diseno", cada una con su propia clave) y decidir a mano en que
#    columna de la hoja quiere verlas -otra extension del chasis que todavia
#    no existe-, en vez de confiar en que `casos=` la resuelva.
# ---------------------------------------------------------------------------
MOTOR_206_EJEMPLO = MD.Motor(
    articulo="206",
    hoja="Ejemplo_PCC2_Art206",     # en un motor real: f"...Art{articulo}"
    titulo="EJEMPLO DE REFERENCIA — NO ES EL MOTOR REAL DEL ART. 206",
    fuente=FUENTE_206,
    corto="EJEMPLO DE REFERENCIA",
    descripcion="Recorte pedagogico del Art. 206, solo el cateto de filete.",
    alcance="NO USAR: el motor real es Collar_PCC2_Art206, escrito a mano.",
    clausulas_espec="206-3.4 · 206-3.5 · 206-4.1",
    aplicacion=True,
    casos=("Operacion", "Diseno"),   # ver LIMITE #4 arriba: NO crea la columna E
    secciones=(FILAS_APLICACION, FILAS_ENTRADA, FILAS_CALCULO, FILAS_VEREDICTO),
    material=MATERIAL_EJEMPLO,
    verificaciones=VERIFICACIONES,
    dictamen=DICTAMEN_EJEMPLO,
    pasos=(PASO_1,),
    especificaciones=ESPECIFICACIONES_EJEMPLO,
)


# ---------------------------------------------------------------------------
# Autocomprobacion opcional: valida que las Cita de este ejemplo de verdad
# existen en resources/, exactamente lo que hace comprobar_procedencia() por
# cada motor real. Este archivo NO se importa desde el builder ni entra a
# MOTORES_DECLARADOS -es solo una plantilla-, asi que este bloque es la unica
# forma de probarlo por su cuenta:
#
#   cd outputs/Base_Datos_Materiales_ASME/scripts
#   python ../../../.agents/skills/motor_pcc2/referencias/plantilla_declaracion.py
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    from pathlib import Path

    tabla = MD.comprobar_procedencia(
        MOTOR_206_EJEMPLO,
        Path(__file__).resolve().parents[4] / "resources")
    print(f"{len(tabla)} citas de calculo, todas resueltas contra resources/:")
    for clave, clausula, archivo, bloque in tabla:
        print(f"  {clave:14s} {clausula:24s} {archivo}#{bloque}")
