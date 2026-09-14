"""Plantilla de declaracion de un motor PCC-2, para la skill `motor_pcc2`.

Que es esto, y que NO es:

- Es la FORMA de una declaracion: que tipos hay, como se encadenan y que
  guardias los vigilan. La forma es MARCO y vale igual para los 37 articulos.
- NO es el contenido de ningun articulo. Las ecuaciones, las variables, las
  condiciones, los umbrales, las verificaciones y las especificaciones salen
  del JSON del articulo que se este declarando, leido en resources/. No se
  copian de aqui y no se sacan de la memoria del modelo (Regla n.1).
- Los valores de este archivo son de UN articulo real (el 206) por una sola
  razon: `comprobar_procedencia()` valida toda `Cita` contra un JSON de
  verdad, asi que un ejemplo con citas inventadas ni siquiera se podria
  ejecutar. Son ilustracion, no procedimiento.
- El motor que de verdad rige el Art. 206 sigue escrito a mano en
  `build_collar_art206()` y es lo unico de ese articulo verificado celda a
  celda contra su *oracle*. Este archivo no lo sustituye, no lo migra y no se
  importa desde el builder.

Como se usa al declarar un articulo:
  1. Se copia la FORMA (Motor / Seccion / Fila / Lista / Verificacion /
     Dictamen / Paso / Especificacion), nunca los valores.
  2. Cada `Cita` apunta a un bloque REAL del `art_XXX.json` del articulo que se
     declara, leido en el barrido del articulo.
  3. Una Fila de tipo FORMULA sin su Cita no construye el libro:
     `comprobar_procedencia()` la rechaza con SystemExit antes de escribir una
     sola celda.
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
#    Motor describen bloques de ESE mismo JSON. `comprobar_procedencia()` no
#    admite un archivo distinto por fila: el articulo entero vive en un solo
#    documento de resources/, y mezclar fuentes dentro de un mismo motor
#    perderia la propiedad de "un motor, un JSON, un `--pdfs` para auditarlo".
# ---------------------------------------------------------------------------
FUENTE = ("ASME PCC/pcc_2/p2_welded_repairs/"
          "art_206_full_encirclement_steel/art_206.json")


def _cita(bloque, clausula):
    """Atajo para no repetir `archivo=` en cada Cita.

    El campo `Cita.archivo` es SOLO lo que ve el ingeniero en la columna de
    referencia (la procedencia se escribe en tres formas, y esta es una de
    ellas). La ruta que de verdad valida `comprobar_procedencia()` contra el
    disco es siempre `motor.fuente`, nunca `Cita.archivo` -- por eso aqui
    basta con un nombre corto y legible.
    """
    return MD.Cita(archivo="art_206.json", bloque=bloque, clausula=clausula)


# ---------------------------------------------------------------------------
# 1. Filas reservadas (CLAVES_RESERVADAS en motor_declarado.py): "unidad",
#    "modo", "temperatura", "dictamen". Un motor con `motor.material is not
#    None` las exige las cuatro -- build_motor_declarado() aborta con un
#    SystemExit legible si falta alguna, citando que pase la necesitaba.
#    `Motor.material` vale None por defecto -un motor sin cascada de material
#    no exige `modo` ni `temperatura`- y `Motor.aplicacion` vale True.
#
#    TODA Fila `tipo=LISTA` declara de donde salen sus items (`Lista`), y el
#    chasis ABORTA si no lo hace. Dos formas, y solo dos:
#      - `Lista(hoja="DB_...", columna="<rotulo de la cabecera>")`: el
#        desplegable lo publica una base de datos del libro. Es la forma
#        normal y la que exigen las reglas 12 y 14.
#      - `Lista(opciones=(...))`: enumeracion cerrada que publica el propio
#        marco. SOLO en las claves reservadas -- fuera de ellas el chasis la
#        rechaza, porque una lista tecleada a mano es lo que la regla 12
#        prohibe.
#    En los dos casos el chasis materializa los items en una columna oculta de
#    la hoja y apunta la validacion a ESE rango (regla 2: nunca una formula
#    como origen).
# ---------------------------------------------------------------------------
FILAS_APLICACION = MD.Seccion("1. APLICACION Y CODIGO", filas=(
    MD.Fila("unidad", "Sistema de unidades", tipo=MD.LISTA, ejemplo="SI",
            lista=MD.Lista(opciones=("SI", "US")),
            comentario="Conmutador SI <-> US de TODO el motor (regla 10). Es "
            "una enumeracion del marco: ninguna base la tabula."),
    MD.Fila("modo", "Codigo de aplicacion", tipo=MD.LISTA, ejemplo="PCC2-206",
            lista=MD.Lista(opciones=("PCC2-206",)),
            comentario="Selector de codigo de la seccion de resolucion de "
            "material. Enumeracion del marco, igual que el conmutador."),
))

FILAS_ENTRADA = MD.Seccion("2. DATOS DE ENTRADA", filas=(
    MD.Fila("temperatura", "Temperatura de evaluacion", magnitud="temp",
            tipo=MD.ENTRADA, ejemplo=20,
            comentario="Alimenta la seccion de resolucion de material "
            "(temp_fuente_cell de construir_seccion7_material)."),
    # Un valor que una base del libro TABULA se elige, nunca se teclea
    # (regla 14). El desplegable se ata a la base y a la columna, por el
    # rotulo impreso de su cabecera.
    MD.Fila("NPS", "Diametro nominal de la tuberia", tipo=MD.LISTA,
            ejemplo='2"',
            lista=MD.Lista(hoja="DB_B36_10", columna="NPS impreso"),
            comentario="Se elige de DB_B36_10 (ASME B36.10M), no se teclea: "
            "la base es la que lo publica (reglas 12 y 14)."),
    # Un dato de SERVICIO o de INSPECCION no lo tabula ninguna base, y por eso
    # se teclea. La linea es esa: si el libro tiene una tabla con el valor,
    # desplegable; si es un dato de campo, entrada manual.
    MD.Fila("Tp", "Espesor medido de la tuberia portadora", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=6.35),
    MD.Fila("Ts", "Espesor de la envolvente", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=9.5),
    MD.Fila("G", "Luz radial de ajuste", magnitud="len",
            tipo=MD.ENTRADA, ejemplo=1.5,
            comentario="206-4.1: el codigo exige cerrar esta luz antes de "
            "soldar; se teclea porque es un dato de campo, no de tabla."),
    MD.Fila("w_adoptado", "Cateto de filete adoptado en el diseno",
            magnitud="len", tipo=MD.ENTRADA, ejemplo=9.5),
))


# ---------------------------------------------------------------------------
# 2. Filas de calculo. Las dos primeras ilustran el caso de "ecuacion vacia
#    con imagen": el bloque llega con `text` inservible y la formula solo en
#    un PNG de `fig/`. Esa imagen ES parte de resources/, asi que se lee con
#    la herramienta Read y de ahi sale la formula -- nunca de un PDF externo
#    ni de la memoria del modelo-, y el texto recuperado se graba en el JSON
#    con su `extraction_amendments`.
# ---------------------------------------------------------------------------
FILAS_CALCULO = MD.Seccion("3. CALCULO DEL CATETO DE FILETE", filas=(
    MD.Fila("w_delgado", "Cateto con envolvente delgada", magnitud="len",
            tipo=MD.FORMULA, formula="={Ts}+{G}",
            cita=_cita(72, "206-3.5(a), Fig. 206-3.5-1"),
            comentario="Texto recuperado de fig/fig_206_3_5_1.png: "
            "'w = Ts + G'. Bloque 72 de art_206.json."),
    MD.Fila("w_grueso", "Cateto maximo con envolvente gruesa", magnitud="len",
            tipo=MD.FORMULA, formula="=1.4*{Tp}+{G}",
            cita=_cita(77, "206-3.5(b), Fig. 206-3.5-2"),
            comentario="Texto recuperado de fig/fig_206_3_5_2.png: "
            "'w_max = 1.4 Tp + G'. Bloque 77 de art_206.json."),
    MD.Fila("w_requerido", "Cateto de filete requerido", magnitud="len",
            tipo=MD.FORMULA,
            formula="=IF({Ts}<=1.4*{Tp},{w_delgado},{w_grueso})",
            cita=_cita(43, "206-3.5"),
            comentario="La eleccion no es una tercera ecuacion: es la propia "
            "estructura del parrafo. El bloque 43 imprime la frase que la "
            "introduce; las dos condiciones que la resuelven estan en los "
            "bloques 44 y 45."),
    MD.Fila("chk_cateto", "Verificacion: cateto adoptado >= requerido",
            tipo=MD.FORMULA,
            formula='=IF({w_adoptado}>={w_requerido},"CUMPLE","NO CUMPLE")',
            cita=_cita(43, "206-3.5")),
))


# ---------------------------------------------------------------------------
# 2b. LOS UMBRALES NO SE TECLEAN AQUI. Nunca.
#
#     Un umbral normativo -"100 mm (4 in.)", "2.5 mm (3/32 in.)"- es un valor
#     del codigo, y un `formula="=100"` lo saca de la memoria de quien declara:
#     cumple la Regla n.1 en la cita y la incumple en el numero. Ademas, en
#     modo US ese 100 no se convierte en nada, y un umbral de aceptacion en la
#     unidad equivocada convierte un "NO CUMPLE" en un "CUMPLE".
#
#     El mecanismo existe y es POR ARTICULO -- registrarlo es un paso
#     obligatorio de cada articulo que se declare:
#
#       1. Se localizan en el JSON del articulo los umbrales que el codigo
#          imprime en DOS unidades, y se anota el fragmento de texto que
#          ancla a cada uno.
#       2. Se registra cada umbral en `UMBRALES_PCC2` de
#          `build_db_materiales.py` como  clave: (articulo, ancla).
#          `leer_umbrales_pcc2()` lee las DOS cifras impresas -nunca convierte
#          una en la otra- y el build ABORTA si el ancla deja de aparecer.
#          La ruta del JSON la resuelve `_json_de_articulo()` desde
#          `Motor.fuente`, asi que un articulo declarado no pide tocar nada
#          mas.
#       3. En la hoja, el umbral entra con `umbral(umbrales, clave, es_si)`,
#          que emite `IF(es_SI, metrico, US)`.
#
#     El diccionario ya trae entradas de los dos motores escritos a mano; se
#     miran como ejemplo de la FORMA de una entrada, no como el procedimiento.
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# 3. Verificaciones y Dictamen.
#
#    De los campos de `Verificacion`, el chasis emite `clave`/`favorables`/
#    `avisos` -son los que pintan el semaforo- y usa `favorables`/`avisos`
#    tambien para componer el dictamen. `requerido`, `adoptado`, `criterio` y
#    `cita` documentan la verificacion y NINGUN pase los escribe en una celda:
#    si quiere verlos en la hoja, declarelos como Fila propia -como
#    "w_requerido" y "w_adoptado" arriba-. Lo que si se comprueba es que sus
#    {nombres} existan (`comprobar_nombres_declarativos`).
#
#    LA FORMULA DEL DICTAMEN NO SE ESCRIBE. La compone el chasis
#    (`formula_dictamen`) desde `Dictamen.compuertas` y
#    `Dictamen.verificaciones`, en tres capas: las compuertas bloquean y su
#    propio texto pasa a ser el dictamen; despues el material; al final el AND
#    de las verificaciones. La fila `dictamen` va SIN formula -- si trae una,
#    el chasis aborta-, y toda clave citada tiene que estar declarada como
#    fila Y como `Verificacion`. Las que entran al AND publican "CUMPLE" y
#    solo eso: es el literal contra el que verificar.py §12 juzga si el
#    dictamen se contradice.
# ---------------------------------------------------------------------------
VERIFICACIONES = (
    MD.Verificacion(
        clave="chk_cateto", rotulo="Cateto de filete",
        requerido="{w_requerido}", adoptado="{w_adoptado}",
        criterio="w_adoptado >= w_requerido",
        favorables=("CUMPLE",), avisos=(),
        cita=_cita(43, "206-3.5")),
)

FILA_DICTAMEN = MD.Fila(
    "dictamen", "Dictamen global del articulo", tipo=MD.FORMULA,
    cita=_cita(43, "206-3.5"),
    comentario="Calculo: DICTAMEN GLOBAL. Lo compone el chasis desde las "
    "compuertas y las verificaciones declaradas; no se teclea.")

DICTAMEN = MD.Dictamen(
    compuertas=(),                    # sin compuertas en este ejemplo
    verificaciones=("chk_cateto",),   # lo que entra al AND
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
                # OJO: ese 2.5 tecleado es exactamente lo que la nota 2b
                # prohibe. En un motor real la cifra entra con
                # `umbral(umbrales, "206_luz_radial", es_si)`, leida de
                # resources/ en sus dos unidades. Se deja asi a la vista
                # porque este archivo no recibe `umbrales`.
                #
                # Bloque 63, NO 62: el 62 es el section_header del parrafo -un
                # rotulo, que no publica ningun valor- y la luz radial la
                # imprime el parrafo 63. `comprobar_procedencia()` rechaza una
                # Cita a un bloque `section_header` por este mismo caso.
                cita=_cita(63, "206-4.1")),
    ))


# ---------------------------------------------------------------------------
# 5. Especificaciones tecnicas de la pestana propia del articulo. Esta tupla
#    es todo lo que hay que declarar: la pestana la construye el chasis.
#
#    `Especificacion.texto` admite {nombres}: `MD.texto_de_especificacion()`
#    los sustituye por la direccion CALIFICADA CON LA HOJA del motor, porque
#    la pestana es otra hoja y un "$D$40" a secas apuntaria dentro de ella. Un
#    nombre que no existe aborta.
#
#    El texto transcribe el requisito con las unidades COMO EL CODIGO LAS
#    IMPRIME ("100 mm (4 in.)"): estas hojas no llevan conmutador propio.
# ---------------------------------------------------------------------------
ESPECIFICACIONES = (
    ("FABRICACION", (
        MD.Especificacion(
            concepto="Longitud minima de la envolvente",
            texto="100 mm (4 in.), y sobrepasar el defecto al menos 50 mm "
            "(2 in.)",
            clausula="206-3.4", editable=False,
            cita=_cita(40, "206-3.4")),
    )),
)


# ---------------------------------------------------------------------------
# 6. El Material. Se declara aparte de `secciones` -el chasis inserta su
#    propio bloque DESPUES de escribir todas las Secciones y Pasos del motor-.
#
#    LIMITE VERIFICADO CONTRA EL CODIGO: `resolver_direcciones()` solo indexa
#    `motor.secciones` y `motor.pasos` (via `_bloques()`). Las celdas que arma
#    `construir_seccion7_material()` -material_id, S(T) resuelto, Temp. max.,
#    dictamen de rango- NO entran a ese indice, asi que ninguna Fila declarada
#    puede citarlas por nombre: un `{S}` en una formula fallaria con "no
#    existe como fila del motor". Si el articulo que se declara necesita el
#    esfuerzo admisible DENTRO de una ecuacion, eso es una extension del
#    chasis que todavia no existe: se declara el hueco con su clausula y su
#    bloque, no se finge que {S} funciona.
#
#    Lo que el chasis SI hace con el material, sin que haya que declarar nada
#    mas: lo antepone al dictamen -sin material resuelto no hay APTO-.
# ---------------------------------------------------------------------------
MATERIAL = MD.Material(
    columnas=(("D", "Metal base"),),
    incluir_ej_ec=False,
    semilla="",     # "" = sin material precargado en la cascada
)


# ---------------------------------------------------------------------------
# 7. El Motor completo.
#
#    `casos` lleva UN solo caso. El chasis emite UNA columna de valor, y
#    `comprobar_casos()` ABORTA si se declaran dos: el reparto de la hoja por
#    casos de presion no es generico y se resuelve articulo por articulo, con
#    el articulo delante. Antes se aceptaban dos en silencio y la segunda
#    columna salia vacia, que es el peor de los resultados posibles: una hoja
#    plausible e incompleta.
# ---------------------------------------------------------------------------
MOTOR_EJEMPLO = MD.Motor(
    articulo="206",
    hoja="Ejemplo_PCC2_Art206",     # en un motor real: f"...Art{articulo}"
    titulo="EJEMPLO DE REFERENCIA — NO ES EL MOTOR REAL DEL ART. 206",
    fuente=FUENTE,
    corto="EJEMPLO DE REFERENCIA",
    descripcion="Recorte ilustrativo: solo la FORMA de una declaracion.",
    alcance="NO USAR: el motor real es Collar_PCC2_Art206, escrito a mano.",
    clausulas_espec="206-3.4 · 206-3.5 · 206-4.1",
    aplicacion=True,
    casos=("Operacion",),
    secciones=(FILAS_APLICACION, FILAS_ENTRADA, FILAS_CALCULO,
               FILAS_VEREDICTO),
    material=MATERIAL,
    verificaciones=VERIFICACIONES,
    dictamen=DICTAMEN,
    pasos=(PASO_1,),
    especificaciones=ESPECIFICACIONES,
)


# ---------------------------------------------------------------------------
# Autocomprobacion opcional: valida que las Cita de este ejemplo de verdad
# existen en resources/ y que la declaracion pasa los guardias del marco --
# exactamente lo que hace el builder por cada motor real. Este archivo NO se
# importa desde el builder ni entra a MOTORES_DECLARADOS -es una plantilla-,
# asi que este bloque es la unica forma de probarlo por su cuenta:
#
#   cd outputs/Base_Datos_Materiales_ASME/scripts
#   python ../../../.agents/skills/motor_pcc2/referencias/plantilla_declaracion.py
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    from pathlib import Path

    MD.comprobar_casos(MOTOR_EJEMPLO)
    MD.comprobar_listas(MOTOR_EJEMPLO)
    MD.comprobar_dictamen(MOTOR_EJEMPLO)
    MD.comprobar_nombres_declarativos(MOTOR_EJEMPLO)
    tabla = MD.comprobar_procedencia(
        MOTOR_EJEMPLO,
        Path(__file__).resolve().parents[4] / "resources")
    print(f"{len(tabla)} citas de calculo, todas resueltas contra resources/:")
    for clave, clausula, archivo, bloque in tabla:
        print(f"  {clave:14s} {clausula:24s} {archivo}#{bloque}")
    print("\nDictamen compuesto por el chasis:")
    print("  " + MD.formula_dictamen(MOTOR_EJEMPLO))
