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
    texto: str               # texto fijo, o formula con {nombres}
    clausula: str
    editable: bool = False
    cita: Cita | None = None


class Material(NamedTuple):
    """La seccion de resolucion de material, que es siempre la misma."""
    columnas: tuple = (("D", "Metal base"),)
    incluir_ej_ec: bool = False
    semilla: str = ""        # material_id del caso precargado, o ""


class Verificacion(NamedTuple):
    clave: str
    rotulo: str
    requerido: str           # formula con {nombres}
    adoptado: str            # formula con {nombres}
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
            tabla.append((f.clave, f.cita.clausula, f.cita.archivo, f.cita.bloque))
    return tabla


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
