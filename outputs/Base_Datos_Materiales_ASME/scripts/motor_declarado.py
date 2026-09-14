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


class Cita(NamedTuple):
    """De donde sale este valor. Sin esto, la celda no se construye."""
    archivo: str             # relativo a resources/
    bloque: int              # indice del bloque en el JSON
    clausula: str            # lo que se imprime en la columna de referencia


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
    """Como se compone el veredicto global."""
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
    casos: tuple = ("Operacion", "Diseno")
    secciones: tuple = ()
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


def sustituir_nombres(formula, direcciones):
    """Cambia {clave} por su direccion. Un nombre que no existe ABORTA.

    Sin este guardia, una referencia a una fila que se renombro produciria un
    #REF! en la hoja, que es un error que nadie mira hasta que alguien firma un
    calculo con el.
    """
    def _uno(m):
        clave = m.group(1)
        if clave not in direcciones:
            raise SystemExit(
                f"motor_declarado: la formula cita {{{clave}}}, que no existe "
                f"como fila del motor. Formula: {formula!r}")
        return direcciones[clave]

    return _RE_NOMBRE.sub(_uno, formula)


def comprobar_procedencia(motor, resources):
    """Tabla de trazabilidad del motor. Aborta si algo no se puede rastrear.

    Un numero que nadie puede rastrear no se construye: es la Regla n.1 hecha
    mecanismo. Se comprueba ademas que el bloque citado EXISTE en el JSON, el
    mismo guardia que la s.9 de verificar.py hace con MAP_Grupo.
    """
    resources = Path(resources)
    cache, tabla = {}, []
    for _titulo, filas in _bloques(motor):
        for f in filas:
            if f.tipo != FORMULA:
                continue
            if f.cita is None:
                raise SystemExit(
                    f"motor_declarado: la fila {f.clave!r} de {motor.hoja} es un "
                    f"calculo sin procedencia. Un numero que nadie puede "
                    f"rastrear no se construye (Regla n.1).")
            ruta = resources / motor.fuente
            if ruta not in cache:
                if not ruta.exists():
                    raise SystemExit(
                        f"motor_declarado: no existe {ruta}. La fuente del motor "
                        f"tiene que estar en resources/.")
                cache[ruta] = json.loads(ruta.read_text(encoding="utf-8"))
            bloques = cache[ruta].get("blocks", [])
            if not 0 <= f.cita.bloque < len(bloques):
                raise SystemExit(
                    f"motor_declarado: {f.clave!r} cita el bloque "
                    f"{f.cita.bloque} de {f.cita.archivo}, que tiene "
                    f"{len(bloques)} bloques.")
            # Un ROTULO no publica un valor. El bloque existe -el guardia de
            # arriba pasa- pero `section_header` es el titulo del parrafo, no
            # el parrafo: citarlo es la forma silenciosa de que una cita
            # parezca valida y apunte al sitio donde el dato NO esta (caso
            # real: "206-4.1 Installation" es el bloque 62 y la luz radial de
            # 2,5 mm la imprime el 63). Se rechaza en vez de admitirse, que es
            # lo que exige la Regla n.1: la trazabilidad tiene que llevar al
            # texto que publica el numero.
            tipo_bloque = str(bloques[f.cita.bloque].get("type", ""))
            if tipo_bloque == "section_header":
                raise SystemExit(
                    f"motor_declarado: {f.clave!r} cita el bloque "
                    f"{f.cita.bloque} de {f.cita.archivo}, que es un "
                    f"section_header ({bloques[f.cita.bloque].get('text', '')!r}). "
                    f"Un rotulo no publica un valor: cite el bloque que "
                    f"imprime el dato.")
            tabla.append((f.clave, f.cita.clausula, f.cita.archivo, f.cita.bloque))
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


def emitir_tabla(ws, motor, helpers):
    """Escribe el motor completo a una hoja de Excel.

    Produce: dict con 'direcciones' (clave -> $col$fila) y 'ultima_fila' (int).

    La procedencia llega a la HOJA, no solo a la tabla de trazabilidad: la
    clausula a la columna de referencia y la explicacion al comentario de la
    celda de valor. Son las otras dos de las tres formas que pide el diseno.
    """
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
                helpers.calculo(ws, celda, sustituir_nombres(f.formula, dirs))
            else:
                helpers.entrada(ws, celda, f.ejemplo)
            fila += 1
    return {"direcciones": dirs, "ultima_fila": fila - 1}
