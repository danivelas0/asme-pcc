"""Chasis de un motor de calculo DECLARADO.

Un motor nuevo no se escribe celda a celda: se DECLARA -que secciones tiene, que
filas, que se teclea y que se calcula- y este modulo lo convierte en hoja. La
diferencia que importa es que las formulas se escriben POR NOMBRE ({P}, {OD}) y
las direcciones las asigna el chasis: la clase de error que obligo al mapa de
filas de la Fase 3 -y que dejo tres defectos latentes- deja de existir.

No importa build_db_materiales: el builder importa este modulo, no al reves.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import NamedTuple

# Tipos de fila. Gobiernan el estilo, si la celda es editable y si entra en el
# manifiesto de reinicio.
ENTRADA = "entrada"          # se teclea
LISTA = "lista"              # se elige de una lista materializada
FORMULA = "formula"          # la calcula el libro

# Layout de la tabla de un motor: rotulo en A, simbolo en B, unidad en C, el
# primer caso en D. Es el de los dos motores que ya existen, y se respeta para
# que las cinco hojas se lean igual.
COL_ROTULO, COL_SIMBOLO, COL_UNIDAD, COL_PRIMER_CASO = 1, 2, 3, 4
FILA_PRIMERA_BANDA = 5       # 1-2 titulo y subtitulo, 3 acciones, 4 libre

# Claves de fila reservadas para pases transversales del libro. Los pases que
# gobiernan el selector de unidades, el selector de codigo, la temperatura de
# evaluacion y la celda del dictamen exigiran estas claves en tareas posteriores.
# Se declaran aqui para que dos tareas no las escriban distinto.
CLAVES_RESERVADAS = ("unidad", "modo", "temperatura", "dictamen")

# Las claves reservadas que el MARCO publica de verdad como ENUMERACION, y las
# unicas donde se admite `Lista(opciones=...)` tecleada.
#
# No son las cuatro reservadas, y confundirlas costaba caro: `CLAVES_RESERVADAS`
# es la lista de claves que los pases transversales exigen, no la de las que son
# una enumeracion del marco. De las cuatro, solo dos lo son -el conmutador SI/US
# (`unidad`) y el selector de codigo (`modo`), que ninguna base del libro tabula-.
# Las otras dos NO: `temperatura` es un dato de servicio que se TECLEA, y
# `dictamen` es el veredicto global, que lo COMPONE el chasis
# (`formula_dictamen`). Con la puerta abierta a las cuatro se podia declarar
#     Fila("dictamen", ..., tipo=LISTA, lista=Lista(opciones=("APTO","REVISAR")))
# y los tres guardias pasaban: la celda del veredicto quedaba editable, amarilla,
# con desplegable y dentro del manifiesto de reinicio -quien usara el motor
# elegiria a mano su propio dictamen-. Se cierra por las dos vias: aqui, y con
# `comprobar_dictamen`, que exige que esa fila sea de tipo FORMULA.
CLAVES_CON_ENUMERACION = ("unidad", "modo")

# Textos del dictamen que compone el chasis. Viven aqui -no en el builder-
# porque los consume tambien verificar.py §12 al juzgar si el dictamen
# contradice a sus verificaciones: un literal escrito dos veces es un literal
# que un dia diverge, y el guardia se quedaria mudo justo donde mas caro sale.
TXT_APTO = "APTO"
TXT_REVISAR = "REVISAR"
TXT_ELIJA_MATERIAL = "ELIJA MATERIAL (seccion de material)"
TXT_MATERIAL_FUERA_DE_RANGO = "REVISAR - MATERIAL FUERA DE RANGO"
# El dictamen de rango que publica la seccion de material del libro. Sus dos
# estados son los que el chasis antepone al AND de verificaciones.
TXT_SIN_MATERIAL = "SIN MATERIAL SELECCIONADO"
TXT_MATERIAL_OK = "OK"
# El unico veredicto favorable que admite una `Verificacion` que entra al AND
# del dictamen. No es una preferencia de estilo: verificar.py §12 punto 4
# compara cada veredicto contra este literal para decidir si el dictamen se
# contradice, asi que una verificacion con otro favorable dejaria al guardia
# midiendo una cosa distinta de la que el motor calcula.
VEREDICTO_CUMPLE = "CUMPLE"


class Cita(NamedTuple):
    """De donde sale este valor. Sin esto, la celda no se construye.

    `archivo` es una ETIQUETA legible ("art_204.json", "B31.3 cap. II"), para
    que el ingeniero vea de un vistazo de que documento sale el dato. La ruta
    que de verdad se abre y se valida es `fuente` si esta, y `motor.fuente` si
    no -que es el caso normal: casi todas las citas de un motor describen
    bloques del articulo del motor-.

    `fuente` existe desde 2026-09-15 y es lo que hace posible declarar un
    articulo que DELEGA el calculo. El Art. 204 no publica ni una ecuacion:
    remite ocho veces al "applicable construction or post-construction code",
    asi que sus filas de calculo citan el B31.3, no el 204. Sin este campo, un
    motor solo podia citar un JSON, y el unico modo de declarar el 204 habria
    sido mentir en la cita o no citarlo.
    """
    archivo: str             # ETIQUETA legible del documento
    bloque: int              # indice del bloque, segun bloques_citables()
    clausula: str            # lo que se imprime en la columna de referencia
    fuente: str = ""         # ruta real en resources/; "" = la del motor


class Lista(NamedTuple):
    """De donde salen los items del desplegable de una Fila `tipo=LISTA`.

    Se declara de UNA de dos formas, nunca de las dos ni de ninguna:

    - `hoja` + `columna`: el desplegable lo publica una BASE DE DATOS del
      libro (`DB_B36_10`, `DB_B31_3`, `DB_BPVC_IID`...), nombrando la columna
      por el rotulo impreso en su cabecera. Es la forma normal y la que exigen
      las reglas 12 y 14 del libro: toda variable que el libro tabula se elige
      de la base que la tabula, nunca se teclea ni se copia a un literal.
    - `opciones`: una enumeracion CERRADA que no existe en ninguna base. Se
      admite en dos sitios y solo en dos (ver `comprobar_listas`):
      (i) las claves de `CLAVES_CON_ENUMERACION` -el sistema de unidades y el
      selector de codigo-, que publica el propio MARCO;
      (ii) cualquier otra clave **acompanada de `cita`**, cuando la enumeracion
      la publica el CODIGO -los cuatro metodos de prueba de `para. 204-6.2`, o
      la caja estructural / no estructural de `para. 204-1(f)`-.
      Fuera de esos dos casos, una lista tecleada a mano es exactamente lo que
      la regla 12 prohibe.

      Lo que la regla 12 ataca es la lista "que no se audita, no se actualiza
      si la base cambia, y puede ofrecer un material que la base ni siquiera
      admite". Una enumeracion con `cita` SI se audita:
      `comprobar_procedencia()` exige que el bloque citado exista y que no sea
      un `section_header`. LIMITE DECLARADO: comprueba que el bloque publique
      texto, NO que cada opcion aparezca literalmente en el -el libro esta en
      espanol y el codigo en ingles-. Eso va en el comentario de la celda.

    En los dos casos el chasis MATERIALIZA los items en una columna oculta de
    la propia hoja y apunta la validacion a ese RANGO, nunca a una formula ni
    a un literal en `formula1` (regla 2 del libro).
    """
    hoja: str = ""           # base del libro: "DB_B36_10", "DB_B31_3"...
    columna: str = ""        # rotulo impreso de la columna, en la cabecera
    opciones: tuple = ()     # enumeracion cerrada: marco, o codigo con `cita`
    cita: Cita | None = None  # obligatoria si `opciones` sale del codigo
    # Tercera forma: un nivel de una CASCADA dimensional (ver `Dimensional`).
    # Los items no salen de una columna suelta sino del nivel anterior -las
    # cedulas que existen PARA ESE NPS-, asi que no los materializa el helper
    # generico: los escribe el constructor de la cascada, que es el que sabe
    # encadenarlos. Aqui solo se marca para que el helper no la toque y para
    # que la declaracion diga de donde sale, que es lo que exige la regla 12.
    cascada: str = ""        # hoy solo "B36"


class Fila(NamedTuple):
    clave: str               # con la que otras filas la nombran: {clave}
    rotulo: str
    magnitud: str = ""       # clave de _U: len, pres, esf, temp... "" = discreta
    tipo: str = FORMULA
    formula: str = ""        # con {nombres}, no con direcciones
    simbolo: str = ""
    ejemplo: object = None   # valor del caso precargado, si es entrada
    comentario: str = ""
    cita: Cita | None = None
    lista: Lista | None = None   # obligatorio -y exclusivo- de `tipo=LISTA`


class Dimensional(NamedTuple):
    """Una cascada norma -> NPS -> cedula contra `DB_B36_10` / `DB_B36_19`.

    Es la regla 14 aplicada a la geometria: NPS, cedula, OD y espesor estan
    TABULADOS, asi que se eligen de la base que los tabula y no se teclean. El
    mecanismo ya existia y estaba probado (`_materializar_cascada_b36`), pero
    cableado a los dos motores escritos a mano: a una sola tanda de columnas
    ocultas y a la columna D. Un motor DECLARADO no podia usarlo.

    El Art. 204 necesita DOS: una para el componente portador y otra para la
    envolvente -la caja de te + caps, que es como se fabrican en planta-, y
    cada una con su propio NPS y su propia cedula. De ahi que esto sea una
    tupla en `Motor.dimensionales` y no un campo suelto.

    `norma`, `nps` y `cedula` nombran filas `tipo=LISTA` con
    `Lista(cascada="B36")`. `od` y `espesor` son opcionales y nombran filas
    `tipo=FORMULA` **sin formula propia**: la escribe el chasis con el mismo
    CHOOSE de la cascada -que es quien conoce la norma elegida y el conmutador
    SI/US-, y `comprobar_dimensionales()` exige que lleguen vacias para que
    nadie escriba ahi una formula creyendo que se usa.
    """
    norma: str               # clave de la fila del selector de norma B36
    nps: str                 # clave de la fila de NPS
    cedula: str              # clave de la fila de cedula
    od: str = ""             # clave de la fila que publica el OD
    espesor: str = ""        # clave de la fila que publica el espesor de pared


class Seccion(NamedTuple):
    titulo: str
    filas: tuple = ()


class Paso(NamedTuple):
    """Un bloque del anexo del flujo: el recorrido de la reparacion.

    Misma forma que una Seccion -titulo de banda + filas nombradas- pero
    aparte, porque el anexo de pasos no es una entrada mas del calculo: es la
    conduccion del procedimiento completo (Tarea 6 de la skill motor_pcc2).
    """
    numero: int
    titulo: str
    clausula: str
    filas: tuple = ()


class Especificacion(NamedTuple):
    """Una fila de la pestana de especificaciones tecnicas del articulo."""
    concepto: str
    # Texto fijo, o formula con {nombres}: los sustituye de verdad
    # texto_de_especificacion(), con la direccion calificada por hoja.
    texto: str
    clausula: str
    editable: bool = False
    cita: Cita | None = None


class Material(NamedTuple):
    """La seccion de resolucion de material, que es siempre la misma."""
    columnas: tuple = (("D", "Metal base"),)
    incluir_ej_ec: bool = False
    semilla: str = ""        # material_id del caso precargado, o ""


class Verificacion(NamedTuple):
    """Un veredicto del motor.

    De sus campos, el chasis SOLO emite `clave`/`favorables`/`avisos` (son los
    que pintan el semaforo). `requerido`, `adoptado`, `criterio` y `cita` se
    declaran para documentar la verificacion y NINGUN pase los escribe en una
    celda: si quiere verlos en la hoja, declarelos como Fila propia. Lo que si
    se comprueba es que sus {nombres} existan (comprobar_nombres_declarativos).
    """
    clave: str
    rotulo: str
    requerido: str           # {nombres} validados, pero HOY no se emiten
    adoptado: str            # {nombres} validados, pero HOY no se emiten
    criterio: str            # el texto de la columna G: "t >= t_req"
    favorables: tuple = ("CUMPLE",)
    avisos: tuple = ()
    cita: Cita | None = None


class Dictamen(NamedTuple):
    """Que entra en el veredicto global. La FORMA la pone el chasis.

    La forma del dictamen es identica en todo motor -las compuertas bloquean
    antes de nada, el material se resuelve despues, y al final las
    verificaciones entran a un AND-, asi que es MARCO y la compone
    `formula_dictamen()`. Lo que cambia por articulo es QUE entra, y eso es lo
    que declaran estos dos campos.

    - `compuertas`: claves de filas cuyo texto BLOQUEA el dictamen. Una
      compuerta tiene que estar declarada ademas como `Verificacion`, porque
      es de ahi de donde sale que textos suyos NO bloquean (`favorables` y
      `avisos`). Cuando una bloquea, su propio texto pasa a ser el dictamen:
      "no elegible por servicio letal" dice mas que un "REVISAR" pelado.
    - `verificaciones`: claves que entran al AND final. Cada una tiene que
      estar declarada como `Verificacion` con `favorables=("CUMPLE",)`.
    """
    compuertas: tuple = ()   # claves de filas cuyo texto bloquea
    verificaciones: tuple = ()   # claves que entran al AND


class Motor(NamedTuple):
    """Un motor de calculo entero, declarado.

    Se define COMPLETO desde el principio aunque las tareas siguientes vayan
    rellenando su uso: un campo que aparece a mitad de plan es un campo que dos
    tareas escriben distinto.
    """
    articulo: str            # "211"
    hoja: str                # "Recargue_PCC2_Art211"
    titulo: str
    fuente: str              # ruta del JSON, relativa a resources/
    corto: str = ""          # "RECARGUE POR SOLDADURA", para la tarjeta
    descripcion: str = ""    # primera linea de la tarjeta
    alcance: str = ""        # segunda linea de la tarjeta (NUNCA estado)
    clausulas_espec: str = ""    # "211-4 · 211-5 · 211-6"
    aplicacion: bool = True  # lleva selector de geometria y codigo
    # UN solo caso de valor. El reparto de la hoja en varias columnas de
    # presion NO es generico y no se resuelve por adelantado (decision del
    # ingeniero, 2026-09-13): el chasis emite una columna y `comprobar_casos()`
    # aborta si se declaran dos, en vez de aceptarlas en silencio y dejar la
    # segunda columna vacia -que es lo que hacia hasta hoy-.
    casos: tuple = ("Operacion",)
    secciones: tuple = ()
    dimensionales: tuple = ()    # (Dimensional, ...) — cascadas NPS/cedula
    material: Material | None = None
    verificaciones: tuple = ()
    dictamen: Dictamen | None = None
    pasos: tuple = ()
    especificaciones: tuple = () # ((grupo, (Especificacion, ...)), ...)


class Helpers(NamedTuple):
    """Las cuatro funciones de estilo que el chasis NO define.

    El chasis no importa build_db_materiales -seria una dependencia circular y
    ademas lo ataria a este libro-: recibe los helpers y los llama. Asi las
    pruebas del chasis corren sin el sistema visual, y el libro real le pasa los
    suyos, que son los mismos que usan los motores escritos a mano.
    """
    banda: object
    rotulo: object
    entrada: object
    calculo: object
    # Materializa el desplegable de una Fila `tipo=LISTA` y apunta su
    # validacion al rango. Va con defecto None -y no como quinto campo
    # obligatorio- para que las pruebas del chasis que no declaran ninguna
    # LISTA sigan construyendo un Helpers de cuatro campos; `emitir_tabla()`
    # aborta con un mensaje claro si hace falta y no esta.
    lista: object = None


_RE_NOMBRE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]*)\}")


def _titulo_paso(paso):
    """Rotulo de banda de un Paso, con el mismo patron que ya usan los anexos
    escritos a mano del 212 y el 206: "PASO N . TITULO - clausula". Los
    separadores son los mismos caracteres que esos dos anexos (punto medio y
    raya, no tilde: la regla de texto sin tildes del libro es sobre vocales
    acentuadas y enies, no sobre estos separadores)."""
    return f"PASO {paso.numero} · {paso.titulo} — {paso.clausula}"


def _bloques(motor):
    """(titulo_de_banda, filas) de cada bloque del motor, en el orden en que
    se lee la hoja: primero las Secciones y despues los Pasos del anexo del
    flujo.

    Una Seccion y un Paso son la MISMA FORMA para lo que necesitan
    resolver_direcciones(), emitir_tabla() y comprobar_procedencia(): un
    titulo de banda y una tupla de filas nombradas. Centralizar el recorrido
    aqui -en vez de que cada una de esas tres funciones repita su propio
    `for seccion in motor.secciones` y ahora tambien su propio
    `for paso in motor.pasos`- es lo que garantiza que no puedan divergir: es
    exactamente la clase de "dos fuentes de verdad" que
    test_emitir_tabla_escribe_en_la_direccion_que_resolver_direcciones_devuelve
    existe para impedir, aplicada ahora a un tercer consumidor del mismo
    arbol. Un motor sin pasos (`motor.pasos == ()`, el valor por defecto) no
    aporta ningun bloque aqui: la ausencia de anexo es el caso normal, no un
    error.
    """
    for seccion in motor.secciones:
        yield seccion.titulo, seccion.filas
    for paso in motor.pasos:
        yield _titulo_paso(paso), paso.filas


# Alias PUBLICO del recorrido. El chasis no es su unico consumidor: el builder
# deriva de el las unidades, las reglas de comentario y el semaforo, y las
# pruebas del libro recorren lo mismo. Un guion bajo que no impide nada solo
# invita a escribir un `for s in motor.secciones` paralelo -que es justo el
# defecto que se cerro el 2026-09-13: tres consumidores ignoraban motor.pasos-.
bloques = _bloques


def resolver_direcciones(motor):
    """clave -> direccion absoluta, asignando las filas en orden de lectura."""
    dirs, fila = {}, FILA_PRIMERA_BANDA
    col = chr(ord("A") + COL_PRIMER_CASO - 1)
    for _titulo, filas in _bloques(motor):
        fila += 1                        # la banda del bloque (seccion o paso)
        for f in filas:
            dirs[f.clave] = f"${col}${fila}"
            fila += 1
    return dirs


# --------------------------------------------------------------------------
# Los cuatro valores que publica la SECCION DE MATERIAL, y no una fila.
#
# `resolver_direcciones()` solo indexa `motor.secciones` y `motor.pasos`, asi
# que hasta 2026-09-15 una formula declarada NO podia usar el esfuerzo
# admisible que resuelve la cascada: la plantilla de la skill lo decia
# explicitamente ("si el articulo necesita {S} dentro de una ecuacion, eso es
# una extension del chasis que todavia no existe").
#
# El Art. 204 la necesita: delega el espesor de la envolvente al B31.3, y la
# eq. (3a) del `para. 304.1.2` es `t = PD/(2(SEW + PY))` -- sin S no hay
# ecuacion. Se resuelve como ya se resolvia el dictamen: el bloque de material
# se construye DESPUES de la tabla, asi que su direccion no existe cuando se
# emiten las formulas; se emite un CENTINELA y `resolver_material()` lo cambia
# por la direccion real en cuanto el bloque existe. Dos pases sobre la misma
# hoja, no dos formulas parecidas en dos sitios.
#
# La sintaxis es `{S}` para la primera columna de material y `{S@E}` para la
# columna E, que es como se declara un motor con metal base y material de la
# caja por separado.
NOMBRES_MATERIAL = {
    "S": "s_t",                      # esfuerzo admisible a la temperatura
    "Tmax": "tmax",                  # temperatura maxima del material
    "material_id": "material_id",    # el id resuelto por la cascada
    "dictamen_material": "dictamen",  # OK / el motivo del bloqueo
}
_RE_MATERIAL = re.compile(
    r"\{(" + "|".join(NOMBRES_MATERIAL) + r")(?:@([A-Z]))?\}")


def _centinela(nombre, columna):
    """Marca que `resolver_material()` cambiara por una direccion real.

    Lleva `@` a proposito: Excel no admite ese caracter en un nombre definido
    ni en una referencia, asi que un centinela que sobreviviera a los dos pases
    no se puede confundir con una formula valida -- se ve, y revienta, en vez
    de calcular en silencio contra la celda equivocada.
    """
    return f"__MAT@{nombre}@{columna}__"


_RE_UMBRAL = re.compile(r"\{UMBRAL:([A-Za-z0-9_]+)\}")


def _centinela_umbral(clave):
    """Marca que `resolver_umbrales()` cambiara por `IF(es_SI, si, us)`.

    Mismo patron que `_centinela()` para el material: un umbral normativo -"4.8
    mm (0.188 in.)"- no se puede teclear en la declaracion (seria sacarlo de la
    memoria de quien declara, y en modo US no se convertiria en nada). Pero
    `UMBRALES_PCC2` solo existe en build_db_materiales y se lee de resources/
    en tiempo de build, no cuando se declara el Motor: se emite un centinela y
    build_motor_declarado lo cambia por la expresion real -leida del codigo, en
    sus dos unidades impresas- en cuanto conoce los umbrales del libro.
    """
    return f"__UMB@{clave}__"


def nombres_de_umbral(motor):
    """Claves {UMBRAL:...} que el motor pide, para que build_motor_declarado
    sepa que resolver y pueda avisar si una clave no esta en UMBRALES_PCC2."""
    claves = set()
    for _titulo, filas in _bloques(motor):
        for f in filas:
            claves.update(_RE_UMBRAL.findall(f.formula or ""))
    return claves


def resolver_umbrales(ws, reemplazos, ultima_fila, ncols):
    """Cambia los centinelas de umbral por la expresion `IF(es_SI, si, us)`.

    `reemplazos` es {clave: expresion}, ya compuesta por quien SI conoce
    `UMBRALES_PCC2` y `umbral()` (build_db_materiales, para no crear un import
    circular con este modulo). Mismo patron de dos pases que `resolver_material`.
    """
    if not reemplazos:
        return 0
    cambios = {_centinela_umbral(c): e for c, e in reemplazos.items()}
    tocadas = 0
    for fila in ws.iter_rows(min_row=1, max_row=ultima_fila, max_col=ncols):
        for celda in fila:
            v = celda.value
            if not isinstance(v, str) or "__UMB@" not in v:
                continue
            for marca, destino in cambios.items():
                v = v.replace(marca, destino)
            celda.value = v
            tocadas += 1
    return tocadas


def sustituir_nombres(formula, direcciones, columna_material=""):
    """Cambia {clave} por su direccion. Un nombre que no existe ABORTA.

    Sin este guardia, una referencia a una fila que se renombro produciria un
    #REF! en la hoja, que es un error que nadie mira hasta que alguien firma un
    calculo con el.

    Los cuatro nombres de `NOMBRES_MATERIAL` son la excepcion: no son filas del
    motor y su direccion todavia no existe cuando se emite la formula, asi que
    salen como centinela y los resuelve `resolver_material()`. `{UMBRAL:...}`
    es la misma idea aplicada a un umbral normativo: lo resuelve
    `resolver_umbrales()`.
    """
    formula = _RE_UMBRAL.sub(lambda m: _centinela_umbral(m.group(1)), formula)

    def _material(m):
        return _centinela(m.group(1), m.group(2) or columna_material or "D")

    formula = _RE_MATERIAL.sub(_material, formula)

    def _uno(m):
        clave = m.group(1)
        if clave not in direcciones:
            raise SystemExit(
                f"motor_declarado: la formula cita {{{clave}}}, que no existe "
                f"como fila del motor. Formula: {formula!r}")
        return direcciones[clave]

    return _RE_NOMBRE.sub(_uno, formula)


def nombres_de_material(motor):
    """(nombre, columna) que el motor pide a la seccion de material."""
    fuera = set()
    for _titulo, filas in _bloques(motor):
        for f in filas:
            col = motor.material.columnas[0][0] if motor.material else "D"
            for nombre, columna in _RE_MATERIAL.findall(f.formula or ""):
                fuera.add((nombre, columna or col))
    return fuera


def comprobar_material_citable(motor):
    """Aborta si una formula pide {S} y el motor no tiene cascada de material.

    Un centinela sin quien lo resuelva llega a la hoja tal cual y Excel lo
    rechaza: mejor decirlo al declarar, con el nombre de la fila delante.
    """
    pedidos = nombres_de_material(motor)
    if not pedidos:
        return
    if motor.material is None:
        raise SystemExit(
            f"motor_declarado: {motor.hoja} usa "
            f"{', '.join(sorted('{' + n + '}' for n, _c in pedidos))} en una "
            f"formula, pero no declara `material`. Esos nombres los publica la "
            f"seccion de resolucion de material; sin ella no existen.")
    columnas = {letra for letra, _et in motor.material.columnas}
    for nombre, columna in sorted(pedidos):
        if columna not in columnas:
            raise SystemExit(
                f"motor_declarado: {motor.hoja} cita {{{nombre}@{columna}}}, "
                f"pero su seccion de material no tiene columna {columna!r} "
                f"(tiene {', '.join(sorted(columnas))}).")


def comprobar_dimensionales(motor):
    """Aborta si una cascada dimensional no encaja con las filas que nombra."""
    filas = {f.clave: f for _t, fs in _bloques(motor) for f in fs}
    declaradas = set()
    for d in motor.dimensionales:
        for campo in ("norma", "nps", "cedula"):
            clave = getattr(d, campo)
            f = filas.get(clave)
            if f is None:
                raise SystemExit(
                    f"motor_declarado: la cascada dimensional de {motor.hoja} "
                    f"nombra {clave!r} como su nivel {campo!r}, y no existe "
                    f"como fila del motor.")
            if f.tipo != LISTA or f.lista is None or f.lista.cascada != "B36":
                raise SystemExit(
                    f"motor_declarado: la fila {clave!r} de {motor.hoja} es el "
                    f"nivel {campo!r} de una cascada dimensional, asi que "
                    f"tiene que ser `tipo=LISTA` con "
                    f"`Lista(cascada=\"B36\")`. Es {f.tipo!r}.")
            if clave in declaradas:
                raise SystemExit(
                    f"motor_declarado: la fila {clave!r} de {motor.hoja} es "
                    f"nivel de dos cascadas dimensionales a la vez.")
            declaradas.add(clave)
        for campo in ("od", "espesor"):
            clave = getattr(d, campo)
            if not clave:
                continue
            f = filas.get(clave)
            if f is None:
                raise SystemExit(
                    f"motor_declarado: la cascada dimensional de {motor.hoja} "
                    f"publica su {campo!r} en {clave!r}, que no existe como "
                    f"fila del motor.")
            if f.tipo != FORMULA:
                raise SystemExit(
                    f"motor_declarado: la fila {clave!r} de {motor.hoja} "
                    f"recibe el {campo!r} de una cascada dimensional: tiene "
                    f"que ser `tipo=FORMULA`. Es {f.tipo!r}.")
            if f.formula:
                raise SystemExit(
                    f"motor_declarado: la fila {clave!r} de {motor.hoja} trae "
                    f"formula propia ({f.formula!r}) y ademas recibe el "
                    f"{campo!r} de una cascada dimensional. La escribe el "
                    f"chasis -es quien conoce la norma elegida y el conmutador "
                    f"SI/US-, asi que la declarada no se usaria: dejela vacia.")
    # El reverso: una fila marcada como nivel de cascada que ninguna
    # `Dimensional` recoge se quedaria SIN desplegable y sin que nada lo diga.
    for _t, fs in _bloques(motor):
        for f in fs:
            if (f.tipo == LISTA and f.lista is not None and f.lista.cascada
                    and f.clave not in declaradas):
                raise SystemExit(
                    f"motor_declarado: la fila {f.clave!r} de {motor.hoja} dice "
                    f"salir de la cascada {f.lista.cascada!r}, pero ninguna "
                    f"`Dimensional` de `motor.dimensionales` la nombra: se "
                    f"quedaria sin desplegable y sin decirlo.")


def resolver_material(ws, motor, refs, ultima_fila, ncols):
    """Cambia los centinelas de material por las direcciones que ya existen.

    Segundo pase, hermano del que RECOMPONE el dictamen: el bloque de material
    se construye despues de la tabla, asi que este es el primer momento en que
    `refs["por_columna"][letra]["s_t"]` tiene un valor.
    """
    pedidos = nombres_de_material(motor)
    if not pedidos:
        return 0
    cambios = {}
    for nombre, columna in pedidos:
        destino = refs["por_columna"][columna][NOMBRES_MATERIAL[nombre]]
        cambios[_centinela(nombre, columna)] = destino
    tocadas = 0
    for fila in ws.iter_rows(min_row=1, max_row=ultima_fila, max_col=ncols):
        for celda in fila:
            v = celda.value
            if not isinstance(v, str) or "__MAT@" not in v:
                continue
            for marca, destino in cambios.items():
                v = v.replace(marca, destino)
            celda.value = v
            tocadas += 1
    return tocadas


def bloques_citables(doc):
    """Los bloques direccionables de un JSON de `resources/`, normalizados.

    `Cita.bloque` es un INDICE, y hasta 2026-09-15 solo tenia sentido en los
    JSON de PCC-2, que traen una lista plana `blocks`. Un motor que delega el
    dimensionamiento al codigo de construccion -el Art. 204 lo hace ocho veces-
    tiene que poder citar el B31.3, cuyo JSON no usa esa clave: trae `preamble`
    y un arbol `sections`, cada una con `paragraphs` y `sections` anidadas. Y
    sus tablas son otro archivo mas, con `columns`/`rows`.

    Aqui se normalizan las tres formas a una sola lista de `{type, text}`. El
    orden es el del documento y es DETERMINISTA, que es la unica propiedad que
    de verdad importa: si el aplanado cambiara, TODAS las citas a ese archivo
    se re-apuntarian en silencio a otro parrafo. `test_motor_declarado.py` fija
    el indice de un parrafo conocido para que eso falle en vez de pasar.
    """
    # 1. PCC-2 y todo lo extraido con marker: lista plana, ya normalizada.
    if isinstance(doc.get("blocks"), list):
        return doc["blocks"]

    # 2. Tabla suelta del B31.3: el rotulo primero, despues una entrada por
    #    fila impresa. Asi `bloque=0` es el titulo -que NO publica un valor y
    #    el guardia de abajo rechaza- y `bloque=n` es la fila n-1.
    if isinstance(doc.get("rows"), list) and doc.get("columns") is not None:
        fuera = [{"type": "section_header",
                  "text": f"{doc.get('table_id', '')} {doc.get('title', '')}".strip()}]
        fuera += [{"type": "table_row", "text": json.dumps(r, ensure_ascii=False)}
                  for r in doc["rows"]]
        return fuera

    # 3. Capitulo del B31.3: preambulo y arbol de secciones, en preorden.
    if isinstance(doc.get("sections"), list):
        fuera = []

        def _parrafos(nodo):
            for p in nodo.get("paragraphs") or ():
                fuera.append({"type": p.get("kind") or "paragraph",
                              "text": p.get("text") or ""})

        def _rama(nodo):
            for s in nodo.get("sections") or ():
                fuera.append({
                    "type": "section_header",
                    "text": f"{s.get('id', '')} {s.get('heading', '')}".strip()})
                _parrafos(s)
                _rama(s)

        for p in doc.get("preamble") or ():
            fuera.append({"type": p.get("kind") or "paragraph",
                          "text": p.get("text") or ""})
        _parrafos(doc)
        _rama(doc)
        return fuera

    return []


def _citas_del_motor(motor):
    """(clave, etiqueta, cita, obligatoria) de todo lo que el motor puede citar.

    `clave` es lo que va a la tabla de trazabilidad -la misma que tenia antes,
    para no cambiarle la forma a quien la consuma-; `etiqueta` es la
    descripcion que sale en el mensaje de error, que sin ella no diria si el
    fallo esta en una fila, en una verificacion o en una especificacion.

    Las `Fila` de calculo son las unicas donde la cita es OBLIGATORIA -sin ella
    no hay numero trazable-, pero `Verificacion` y `Especificacion` tambien la
    declaran, y hasta 2026-09-15 NADIE las comprobaba: una cita a un bloque
    inexistente en una especificacion pasaba en verde. Se validan igual cuando
    estan; seguir sin declararlas sigue siendo legal.
    """
    # Las filas od/espesor de una cascada dimensional son FORMULA pero SIN
    # formula propia (comprobar_dimensionales lo exige): las escribe el
    # chasis con el lookup contra DB_B36_10/19, cuya procedencia ya audita
    # verificar.py §11 fila a fila. Exigirles ademas una Cita a un bloque de
    # texto seria pedir una trazabilidad que no les corresponde -el dato no
    # sale de un parrafo, sale de una base de datos ya auditada-.
    sin_cita_exigible = {d.od for d in motor.dimensionales if d.od} | \
        {d.espesor for d in motor.dimensionales if d.espesor}
    for _titulo, filas in _bloques(motor):
        for f in filas:
            if f.tipo == FORMULA:
                yield (f.clave, f"fila {f.clave!r}", f.cita,
                       f.clave not in sin_cita_exigible)
            elif f.cita is not None:
                yield f.clave, f"fila {f.clave!r}", f.cita, False
            # Una enumeracion que publica el codigo se cita igual (ver Lista).
            if f.lista is not None and f.lista.cita is not None:
                yield (f"{f.clave}.lista", f"lista de {f.clave!r}",
                       f.lista.cita, False)
    for v in motor.verificaciones:
        if v.cita is not None:
            yield v.clave, f"verificacion {v.clave!r}", v.cita, False
    for _grupo, especs in motor.especificaciones:
        for e in especs:
            if e.cita is not None:
                yield (e.concepto, f"especificacion {e.concepto!r}",
                       e.cita, False)


def comprobar_procedencia(motor, resources):
    """Tabla de trazabilidad del motor. Aborta si algo no se puede rastrear.

    Un numero que nadie puede rastrear no se construye: es la Regla n.1 hecha
    mecanismo. Se comprueba ademas que el bloque citado EXISTE en el JSON, el
    mismo guardia que la s.9 de verificar.py hace con MAP_Grupo.

    La cita se valida contra SU PROPIO archivo (`Cita.archivo`), no contra
    `motor.fuente`. Hasta 2026-09-15 `archivo` era decorativo -solo se imprimia-
    y todo se validaba contra la fuente unica del motor; un articulo que delega
    el calculo a su codigo de construccion no se puede declarar asi. `archivo`
    vacio sigue significando "la fuente del motor", que es el caso normal.
    """
    resources = Path(resources)
    cache, tabla = {}, []
    for clave, etiqueta, cita, obligatoria in _citas_del_motor(motor):
        if cita is None:
            if obligatoria:
                raise SystemExit(
                    f"motor_declarado: la {etiqueta} de {motor.hoja} es un "
                    f"calculo sin procedencia. Un numero que nadie puede "
                    f"rastrear no se construye (Regla n.1).")
            continue
        ruta = resources / (cita.fuente or motor.fuente)
        if ruta not in cache:
            if not ruta.exists():
                raise SystemExit(
                    f"motor_declarado: la {etiqueta} de {motor.hoja} cita "
                    f"{cita.archivo!r}, cuya ruta ({cita.fuente or motor.fuente}) "
                    f"no existe en {resources}. Todo lo que cita un motor tiene "
                    f"que estar en resources/.")
            cache[ruta] = bloques_citables(
                json.loads(ruta.read_text(encoding="utf-8")))
        bloques = cache[ruta]
        donde = cita.archivo or cita.fuente or motor.fuente
        if not 0 <= cita.bloque < len(bloques):
            raise SystemExit(
                f"motor_declarado: la {etiqueta} de {motor.hoja} cita el bloque "
                f"{cita.bloque} de {donde}, que tiene {len(bloques)} bloques.")
        # Un ROTULO no publica un valor. El bloque existe -el guardia de arriba
        # pasa- pero `section_header` es el titulo del parrafo, no el parrafo:
        # citarlo es la forma silenciosa de que una cita parezca valida y
        # apunte al sitio donde el dato NO esta (caso real: "206-4.1
        # Installation" es el bloque 62 y la luz radial de 2,5 mm la imprime el
        # 63). Se rechaza en vez de admitirse, que es lo que exige la Regla n.1:
        # la trazabilidad tiene que llevar al texto que publica el numero.
        if str(bloques[cita.bloque].get("type", "")) == "section_header":
            raise SystemExit(
                f"motor_declarado: la {etiqueta} de {motor.hoja} cita el bloque "
                f"{cita.bloque} de {donde}, que es un section_header "
                f"({bloques[cita.bloque].get('text', '')!r}). Un rotulo no "
                f"publica un valor: cite el bloque que imprime el dato.")
        tabla.append((clave, cita.clausula, donde, cita.bloque))
    return tabla


def direcciones_externas(motor):
    """clave -> direccion CALIFICADA CON LA HOJA del motor.

    `resolver_direcciones()` devuelve `$D$12`, que solo vale dentro de la propia
    hoja del motor. La pestana de especificaciones tecnicas es otra hoja y SI
    lee las celdas de su motor -son dos pestanas del mismo artefacto, no dos
    motores (regla 13)-, asi que ahi un `$D$12` a secas apuntaria a la fila 12
    de la pestana: el peor error posible, uno que parece un resultado. Se
    califica con el nombre de la hoja, entrecomillado porque un nombre con
    caracteres raros lo exige y con comillas de mas nunca sobra.
    """
    hoja = motor.hoja.replace("'", "''")
    return {c: f"'{hoja}'!{d}" for c, d in resolver_direcciones(motor).items()}


def texto_de_especificacion(motor, especificacion):
    """El texto de una Especificacion con sus {nombres} ya sustituidos.

    `Especificacion.texto` se documenta como "texto fijo, o formula con
    {nombres}" y hasta 2026-09-13 se copiaba CRUDO a la pestana: una
    especificacion con `{t_req}` imprimia la llave literal, sin abortar. La
    promesa se cierra con mecanismo y no con un veto, porque el mecanismo ya
    existe entero -`sustituir_nombres()` mas la calificacion por hoja- y es lo
    que hace util a esta pestana: que la especificacion tecnica cite el espesor
    que el motor acaba de calcular en vez de repetirlo a mano y poder mentir.
    Un nombre que no existe aborta con el mensaje de `sustituir_nombres()`.
    """
    return sustituir_nombres(especificacion.texto, direcciones_externas(motor))


def comprobar_nombres_declarativos(motor):
    """Aborta si un {nombre} de `Verificacion` no existe como fila del motor.

    `Verificacion.requerido` y `.adoptado` se documentan como "formula con
    {nombres}" y HOY no los lee ningun pase (solo `clave`/`favorables`/`avisos`
    alimentan el semaforo). No se prohibe la llave -escribir la verificacion en
    terminos de las filas del motor es justo como se documenta, y como la usa
    la plantilla de la skill- pero tampoco se deja sin mecanismo: se comprueba
    que cada nombre citado EXISTE. Asi, el dia que alguien renombre la fila
    `w_requerido`, la declaracion que la nombraba aborta con un mensaje legible
    en vez de quedarse describiendo una fila que ya no esta.

    Lo que NO se hace es fingir que esas dos cadenas acaban en una celda: no
    hay pase que las emita, y eso sigue declarado en la plantilla de la skill.
    """
    dirs = resolver_direcciones(motor)
    for v in motor.verificaciones:
        for campo in ("requerido", "adoptado"):
            for nombre in _RE_NOMBRE.findall(str(getattr(v, campo))):
                if nombre not in dirs:
                    raise SystemExit(
                        f"motor_declarado: la Verificacion {v.clave!r} de "
                        f"{motor.hoja} cita {{{nombre}}} en su campo "
                        f"{campo!r}, que no existe como fila del motor.")
            if "{" in str(getattr(v, campo)) and not _RE_NOMBRE.search(
                    str(getattr(v, campo))):
                raise SystemExit(
                    f"motor_declarado: la Verificacion {v.clave!r} de "
                    f"{motor.hoja} trae una llave sin cerrar o con un nombre "
                    f"invalido en {campo!r}: {getattr(v, campo)!r}.")


def comprobar_casos(motor):
    """Aborta si el motor declara mas de un caso de valor.

    El chasis emite UNA columna de valor: `emitir_tabla()` fija la columna
    fuera del bucle de filas y `resolver_direcciones()` devuelve UNA direccion
    por clave. `Motor.casos` aceptaba dos en silencio y la segunda columna
    salia vacia, sin aviso -el peor de los dos resultados posibles: una hoja
    plausible e incompleta-.

    No se construye un mecanismo generico de columnas por caso a proposito
    (decision del ingeniero, 2026-09-13): el reparto por casos no es agnostico
    y se resuelve articulo por articulo, con el articulo delante. Lo unico que
    debe hacer el marco es no fingir que lo resuelve.
    """
    if len(motor.casos) != 1:
        raise SystemExit(
            f"motor_declarado: {motor.hoja} declara {len(motor.casos)} casos "
            f"({', '.join(map(str, motor.casos)) or 'ninguno'}). El chasis "
            f"emite UNA columna de valor: el reparto de la hoja por casos de "
            f"presion NO es generico y se resuelve articulo por articulo, no "
            f"por adelantado. Declare un solo caso, y si este articulo "
            f"necesita de verdad dos columnas, parese aqui y decidalo con el "
            f"articulo delante.")


def comprobar_listas(motor):
    """Aborta si un desplegable no sale de una base de datos del libro.

    Instruccion del ingeniero (2026-09-13): "todo material debe ser extraido
    de las bases de datos existentes, siempre". Es la regla 12/14 del libro
    elevada a absoluta, y hasta hoy el chasis no podia ni cumplirla: una Fila
    `tipo=LISTA` llegaba a `helpers.entrada()` igual que una ENTRADA y no
    generaba ninguna validacion, asi que un motor declarado no podia ofrecer
    un desplegable aunque quisiera.

    La unica excepcion son las enumeraciones que publica el PROPIO MARCO y que
    no existen en ninguna base -el sistema de unidades y el selector de
    codigo-, acotadas por nombre a `CLAVES_CON_ENUMERACION`. El guardia bloquea
    **cualquier** clave con `Lista(opciones=...)` que no este en esa constante.
    Acotarlas por nombre y no por criterio es deliberado: "esto no es un dato
    tabulado" no se puede comprobar leyendo la declaracion, y dejarlo a juicio
    de quien declara reabre justo la puerta que la instruccion cierra.
    """
    for _titulo, filas in _bloques(motor):
        for f in filas:
            if f.tipo != LISTA:
                if f.lista is not None:
                    raise SystemExit(
                        f"motor_declarado: la fila {f.clave!r} de {motor.hoja} "
                        f"declara una Lista pero su tipo es {f.tipo!r}. Una "
                        f"lista solo gobierna una fila `tipo=LISTA`.")
                continue
            if f.lista is None:
                raise SystemExit(
                    f"motor_declarado: la fila {f.clave!r} de {motor.hoja} es "
                    f"`tipo=LISTA` y no dice de que base sale. Todo "
                    f"desplegable se ata a una base de datos del libro "
                    f"-Lista(hoja='DB_...', columna='<rotulo de la "
                    f"cabecera>')-: una lista de items tecleada a mano no se "
                    f"audita, no se actualiza si la base cambia y puede "
                    f"ofrecer un valor que la base ni admite (reglas 12 y 14).")
            if f.lista.cascada:
                # Los items de un nivel dependen del nivel de arriba, asi que
                # no es una columna que se pueda copiar: la escribe el
                # constructor de la cascada. Aqui solo se exige que la fila no
                # diga ademas otra procedencia, que seria decir dos cosas.
                if f.lista.hoja or f.lista.columna or f.lista.opciones:
                    raise SystemExit(
                        f"motor_declarado: la Lista de {f.clave!r} en "
                        f"{motor.hoja} dice que sale de la cascada "
                        f"{f.lista.cascada!r} y ademas declara hoja, columna u "
                        f"opciones. Una Lista es una cosa o la otra.")
                if f.lista.cascada != "B36":
                    raise SystemExit(
                        f"motor_declarado: la Lista de {f.clave!r} en "
                        f"{motor.hoja} pide la cascada {f.lista.cascada!r}, "
                        f"que no existe. Hoy solo hay 'B36'.")
                continue
            de_base = bool(f.lista.hoja or f.lista.columna)
            if de_base and f.lista.opciones:
                raise SystemExit(
                    f"motor_declarado: la fila {f.clave!r} de {motor.hoja} "
                    f"declara a la vez una base y una enumeracion. Una Lista "
                    f"es una cosa o la otra.")
            if de_base:
                if not (f.lista.hoja and f.lista.columna):
                    raise SystemExit(
                        f"motor_declarado: la Lista de {f.clave!r} en "
                        f"{motor.hoja} declara hoja o columna, pero no las "
                        f"dos. Hace falta la base y el rotulo de su columna.")
                continue
            if not f.lista.opciones:
                raise SystemExit(
                    f"motor_declarado: la Lista de {f.clave!r} en "
                    f"{motor.hoja} llega vacia: ni base ni enumeracion.")
            if f.clave in CLAVES_CON_ENUMERACION:
                if f.lista.cita is not None:
                    raise SystemExit(
                        f"motor_declarado: la Lista de {f.clave!r} en "
                        f"{motor.hoja} trae `cita`, pero {f.clave!r} es una "
                        f"enumeracion del MARCO -el conmutador SI/US y el "
                        f"selector de codigo no los publica ningun articulo-. "
                        f"Una cita ahi afirmaria una procedencia que no existe.")
            elif f.lista.cita is None:
                raise SystemExit(
                    f"motor_declarado: la fila {f.clave!r} de {motor.hoja} "
                    f"declara una enumeracion tecleada "
                    f"({', '.join(map(str, f.lista.opciones))}) sin `cita`. "
                    f"Una enumeracion se admite en dos sitios: "
                    f"{' y '.join(CLAVES_CON_ENUMERACION)} -las que publica el "
                    f"propio marco- o cualquier clave cuya enumeracion la "
                    f"publique el CODIGO, y entonces tiene que decir DONDE "
                    f"(Lista(opciones=..., cita=Cita(...))). Sin una de las "
                    f"dos cosas, sale de una base del libro (reglas 12 y 14).")
            # El `ejemplo` es el valor que el chasis escribe en la celda, y una
            # celda con validacion "detener" cuyo valor precargado no esta en
            # la lista es un formulario que nace invalido: Excel no lo rechaza
            # -la validacion solo actua al teclear- asi que nadie lo ve hasta
            # que se toca. Es la misma clase de fallo que nombra la regla 12
            # ("puede ofrecer un valor que la base ni admite"), aplicada al
            # valor precargado. La rama de base se comprueba en el builder,
            # que es quien conoce los items materializados.
            if f.ejemplo not in (None, "") and \
                    str(f.ejemplo) not in [str(o) for o in f.lista.opciones]:
                raise SystemExit(
                    f"motor_declarado: la fila {f.clave!r} de {motor.hoja} "
                    f"trae ejemplo={f.ejemplo!r}, que no esta entre los items "
                    f"de su enumeracion "
                    f"({', '.join(map(repr, f.lista.opciones))}).")


def comprobar_dictamen(motor):
    """Aborta si el dictamen declarado y la fila `dictamen` no se corresponden.

    Tres cosas, y las tres existen porque el chasis COMPONE la formula del
    veredicto global (`formula_dictamen`) en vez de exigir que se teclee:

    1. La fila reservada `dictamen` y `Motor.dictamen` van juntas o no van. Una
       fila de dictamen sin declaracion no tendria con que componerse, y una
       declaracion sin fila no tendria donde escribirse.
    1b. La fila `dictamen` es de tipo FORMULA. Este guardia miraba solo
       `fila.formula`, que en una fila ENTRADA o LISTA esta vacia, asi que
       pasaba sin ver nada: se podia declarar el veredicto global como
       desplegable y la celda que dice si la reparacion cumple quedaba
       editable, con lista y dentro del manifiesto de reinicio -el usuario
       elegiria a mano su propio dictamen-. El dictamen no es un dato de
       entrada: lo CALCULA el libro, y se dice aqui por el tipo, no por la
       ausencia de formula.
    2. La fila `dictamen` no trae formula propia: la pone el chasis. Una
       formula tecleada al lado de una declaracion es la forma segura de que
       las dos digan cosas distintas -y verificar.py §12 lo encontraria
       despues, en vez de prevenirlo-.
    3. Toda clave citada -compuerta o verificacion- esta declarada como fila Y
       como `Verificacion`: de la `Verificacion` salen los textos que NO
       bloquean. Y la que entra al AND tiene que tener `favorables` igual a
       ("CUMPLE",), porque es contra ese literal exacto contra el que la §12
       juzga si el dictamen se contradice.
    """
    dirs = resolver_direcciones(motor)
    fila = next((f for _t, filas in _bloques(motor) for f in filas
                 if f.clave == "dictamen"), None)
    if (motor.dictamen is None) != (fila is None):
        raise SystemExit(
            f"motor_declarado: {motor.hoja} declara "
            f"{'Motor.dictamen sin fila `dictamen`' if fila is None else 'la fila `dictamen` sin Motor.dictamen'}"
            f". Los dos van juntos: el chasis compone la formula del veredicto "
            f"global desde Motor.dictamen y la escribe en esa fila.")
    if fila is not None and fila.tipo != FORMULA:
        raise SystemExit(
            f"motor_declarado: la fila `dictamen` de {motor.hoja} se declara "
            f"con tipo={fila.tipo!r}. El veredicto global lo CALCULA el libro "
            f"(lo compone el chasis desde Motor.dictamen): declarado como "
            f"{ENTRADA!r} o {LISTA!r}, la celda que dice si la reparacion "
            f"cumple quedaria editable, con desplegable y dentro del "
            f"manifiesto de reinicio. Declarela tipo={FORMULA!r}, sin formula.")
    if motor.dictamen is None:
        return
    if fila.formula:
        raise SystemExit(
            f"motor_declarado: la fila `dictamen` de {motor.hoja} trae una "
            f"formula tecleada ({fila.formula!r}). La compone el chasis desde "
            f"Motor.dictamen (compuertas + verificaciones): declare QUE entra, "
            f"no COMO se escribe.")
    por_clave = {v.clave: v for v in motor.verificaciones}
    for clase, claves in (("compuerta", motor.dictamen.compuertas),
                          ("verificacion", motor.dictamen.verificaciones)):
        for clave in claves:
            if clave not in dirs:
                raise SystemExit(
                    f"motor_declarado: el dictamen de {motor.hoja} cita la "
                    f"{clase} {clave!r}, que no existe como fila del motor.")
            if clave not in por_clave:
                raise SystemExit(
                    f"motor_declarado: el dictamen de {motor.hoja} cita la "
                    f"{clase} {clave!r}, que no esta declarada como "
                    f"Verificacion. De ahi salen los textos que cuentan como "
                    f"favorables y como aviso; sin ellos el chasis no puede "
                    f"componer la formula sin inventarlos.")
            if clase == "verificacion" and \
                    tuple(por_clave[clave].favorables) != (VEREDICTO_CUMPLE,):
                raise SystemExit(
                    f"motor_declarado: la Verificacion {clave!r} de "
                    f"{motor.hoja} entra al AND del dictamen con favorables="
                    f"{tuple(por_clave[clave].favorables)!r}. Las que entran "
                    f"al AND publican {VEREDICTO_CUMPLE!r} y solo eso: es el "
                    f"literal contra el que verificar.py §12 juzga si el "
                    f"dictamen contradice a sus verificaciones, y con otro "
                    f"favorable ese guardia mediria otra cosa.")


def formula_dictamen(motor, celdas_material=()):
    """La formula del veredicto global, compuesta desde la declaracion.

    Tres capas, en este orden y por este motivo:

    1. Las COMPUERTAS. Si alguna bloquea, su propio texto ES el dictamen y no
       se evalua nada mas -un "no elegible por servicio letal" dice mas que un
       "REVISAR" pelado, y evaluar verificaciones de un caso que el codigo ni
       admite seria contestar a una pregunta que no se hizo-.
    2. El MATERIAL, si el motor lo lleva. Falta de seleccion y fuera de rango
       son estados distintos: confundirlos hace leer un formulario recien
       abierto como un material rechazado por temperatura.
    3. El AND de las VERIFICACIONES.

    `celdas_material` son las celdas del "Dictamen de rango" de la seccion de
    material, una por columna. Llegan por parametro y no se adivinan aqui: el
    chasis no sabe donde las escribio el libro.

    Es la MISMA funcion la que compone la formula con y sin material -el
    builder la llama por segunda vez cuando ya conoce esas celdas-, no dos
    formulas parecidas: si la forma del dictamen cambiara, cambia en un solo
    sitio.
    """
    dirs = resolver_direcciones(motor)
    por_clave = {v.clave: v for v in motor.verificaciones}
    if motor.dictamen.verificaciones:
        cond = ",".join(f'{dirs[c]}="{VEREDICTO_CUMPLE}"'
                        for c in motor.dictamen.verificaciones)
        nucleo = f'IF(AND({cond}),"{TXT_APTO}","{TXT_REVISAR}")'
    else:
        # Sin verificaciones no hay nada que pueda fallar, y esa es tambien la
        # lectura de verificar.py §12 (`all([])` es verdadero): el dictamen
        # tiene que decir APTO, o el guardia y el motor discreparian.
        nucleo = f'"{TXT_APTO}"'
    if celdas_material:
        sin = ",".join(f'{m}="{TXT_SIN_MATERIAL}"' for m in celdas_material)
        mal = ",".join(f'{m}<>"{TXT_MATERIAL_OK}"' for m in celdas_material)
        nucleo = (f'IF(OR({sin}),"{TXT_ELIJA_MATERIAL}",'
                  f'IF(OR({mal}),"{TXT_MATERIAL_FUERA_DE_RANGO}",{nucleo}))')
    # Se recorre al reves porque cada compuerta ENVUELVE a lo anterior: la
    # primera declarada tiene que quedar por fuera, que es donde se evalua
    # antes.
    for clave in reversed(tuple(motor.dictamen.compuertas)):
        v, d = por_clave[clave], dirs[clave]
        # Un aviso NO bloquea y se compara por prefijo, igual que lo hace el
        # semaforo del libro (`_sem`): el texto del aviso suele llevar detras
        # el motivo, y exigir igualdad exacta lo convertiria en bloqueo.
        ok = ([f'{d}="{x}"' for x in v.favorables] +
              [f'LEFT({d},{len(a)})="{a}"' for a in v.avisos])
        nucleo = f'IF(NOT(OR({",".join(ok)})),{d},{nucleo})'
    return "=" + nucleo


def emitir_tabla(ws, motor, helpers):
    """Escribe el motor completo a una hoja de Excel.

    Produce: dict con 'direcciones' (clave -> $col$fila) y 'ultima_fila' (int).

    La procedencia llega a la HOJA, no solo a la tabla de trazabilidad: la
    clausula a la columna de referencia y la explicacion al comentario de la
    celda de valor. Son las otras dos de las tres formas que pide el diseno.
    """
    # Los tres guardias del marco corren AQUI y no solo desde el builder: esta
    # es la unica puerta por la que una declaracion se convierte en hoja, y un
    # guardia que se pueda esquivar llamando a otra funcion no es un guardia.
    comprobar_casos(motor)
    comprobar_listas(motor)
    comprobar_dictamen(motor)
    dirs = resolver_direcciones(motor)
    col = chr(ord("A") + COL_PRIMER_CASO - 1)
    fila = FILA_PRIMERA_BANDA
    # Recorre el mismo arbol que resolver_direcciones() -via _bloques()-, asi
    # que el anexo de pasos (si el motor declara alguno) sale detras de la
    # ultima Seccion, con su propia banda, sin que esta funcion tenga que
    # saber nada de Paso mas alla de que tambien es (titulo, filas).
    for titulo, filas in _bloques(motor):
        helpers.banda(ws, fila, titulo)
        fila += 1
        for f in filas:
            helpers.rotulo(ws, fila, f.rotulo, f.simbolo, f.magnitud,
                           referencia=f.cita.clausula if f.cita else "",
                           comentario=f.comentario)
            celda = f"{col}{fila}"
            if f.tipo == FORMULA:
                # La fila reservada `dictamen` no trae formula propia
                # (comprobar_dictamen lo exige): la COMPONE el chasis desde lo
                # que el motor declara que entra al veredicto.
                expr = (formula_dictamen(motor) if f.clave == "dictamen"
                        else sustituir_nombres(
                            f.formula, dirs,
                            motor.material.columnas[0][0]
                            if motor.material else ""))
                helpers.calculo(ws, celda, expr)
            else:
                helpers.entrada(ws, celda, f.ejemplo)
                if f.tipo == LISTA:
                    if helpers.lista is None:
                        raise SystemExit(
                            f"motor_declarado: {motor.hoja} declara la fila "
                            f"{f.clave!r} como LISTA y los helpers recibidos "
                            f"no traen `lista`. Sin el, la celda quedaria "
                            f"editable a mano y sin desplegable, que es "
                            f"justo lo que las reglas 12 y 14 prohiben.")
                    helpers.lista(ws, celda, f.lista)
            fila += 1
    return {"direcciones": dirs, "ultima_fila": fila - 1}
