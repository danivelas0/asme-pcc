# Skill `motor_pcc2` — plan de implementación

> **Para ejecutores:** usar `superpowers:subagent-driven-development` o
> `superpowers:executing-plans`. Casillas `- [ ]` para seguimiento tarea por tarea.
>
> **NO EJECUTAR sin orden explícita del ingeniero** (regla general de `outputs/plans/`).

**Objetivo:** que `/motor_pcc2 <artículo>` construya el motor de cálculo de cualquier
artículo de ASME PCC-2 dentro del libro existente, leyendo el artículo de `resources/`,
con una sola parada de criterio de ingeniería y sin que nadie escriba una dirección de
celda.

**Arquitectura:** el motor nuevo nace **declarativo**. La skill escribe una declaración
por artículo (`motores/art_XXX.py`) y un **chasis** nuevo (`motor_declarado.py`) la
convierte en hoja, reutilizando los pases transversales que ya existen en
`build_db_materiales.py`. Las fórmulas se escriben **por nombre** (`{P}`, `{OD}`) y el
chasis resuelve las direcciones. El 212 y el 206 **no se tocan**.

**Tech stack:** Python 3.14, openpyxl 3.1.5, pytest, VBA (proyecto del libro), Excel para
`verificar.py`.

**Spec:** `outputs/plans/diseno_skill_motor_pcc2.md`

## Restricciones globales

Todas las tareas las heredan. Copiadas del proyecto, sin reinterpretar:

- **Regla nº 1:** ningún valor normativo sale de la memoria del modelo. Todo sale de
  `resources/`, y toda celda de cálculo lleva su archivo, bloque y cláusula.
- **Cero funciones de matriz dinámica:** nada de `FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`,
  `VSTACK`, `_xlfn`. Solo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`.
- **Validación de datos:** solo rango literal o lista de ítems; nunca una fórmula como
  origen. Las listas dependientes se materializan en columnas ocultas.
- **No extrapolar:** por encima del último punto tabulado, dictamen bloqueado.
- **SI y US son extracciones independientes**, nunca conversiones (regla 9). Todo motor
  lleva conmutador (regla 10).
- **Unidades visibles** en toda variable.
- **Comentarios solo en las columnas de valor** (D, E, F).
- **Dos decimales** en toda celda cuya fila declare una magnitud; lo discreto se queda.
- **Todo veredicto lleva semáforo**, y el relleno de un dxf se escribe en `fgColor` **y**
  `bgColor`.
- **Las tarjetas del árbol no muestran texto de estado.**
- **`snake_case`** en todo archivo nuevo. Comentarios y documentos **en español**.
- **`Parche_PCC2_Art212` y `Collar_PCC2_Art206` salen idénticos** de todo el proceso.

## Estructura de archivos

| Archivo | Responsabilidad |
|---|---|
| `scripts/motor_declarado.py` **(nuevo)** | El chasis: los tipos de la declaración, la resolución de nombres a direcciones, los guardias de procedencia y `build_motor_declarado()`. Librería pura: no importa el builder, el builder la importa a ella. |
| `scripts/motores/__init__.py` **(nuevo)** | Registro de los motores declarados: `MOTORES_DECLARADOS`. |
| `scripts/motores/art_XXX.py` **(nuevo, uno por artículo)** | La declaración de un artículo. Lo único que escribe la skill por artículo. |
| `scripts/test_motor_declarado.py` **(nuevo)** | Pruebas del chasis, puras: sin Excel y sin libro. |
| `scripts/build_db_materiales.py` | Importa el chasis, construye los motores declarados y los añade al árbol. **No se reescribe nada de lo que ya hay.** |
| `scripts/verificar.py` | Sección nueva, genérica para todo motor declarado. |
| `scripts/test_dashboard.py` | Clase nueva, genérica para todo motor declarado. |
| `scripts/vba/mod_nav.vba` | Una línea por hoja nueva en `HojasNavegables()`. |
| `.agents/skills/motor_pcc2/SKILL.md` **(nuevo)** | Las nueve fases y la parada. |
| `.agents/skills/motor_pcc2/referencias/reglas_del_libro.md` **(nuevo)** | El catálogo de reglas que el libro ya pagó. |
| `.agents/skills/motor_pcc2/referencias/plantilla_declaracion.py` **(nuevo)** | La declaración comentada, con un ejemplo. |

---

## Tarea 1 — Los tipos de la declaración y la resolución de nombres

**Archivos:**
- Crear: `outputs/Base_Datos_Materiales_ASME/scripts/motor_declarado.py`
- Crear: `outputs/Base_Datos_Materiales_ASME/scripts/test_motor_declarado.py`

**Interfaces:**
- Produce: `Cita`, `Fila`, `Seccion`, `Motor`, `resolver_direcciones(motor) -> dict[str, str]`,
  `sustituir_nombres(formula, direcciones) -> str`.

- [ ] **Paso 1: escribir la prueba que falla**

```python
# test_motor_declarado.py
import pytest
import motor_declarado as M


def _motor_minimo():
    """Dos filas en una seccion: una entrada y un calculo que la usa."""
    return M.Motor(
        articulo="999",
        hoja="Prueba_PCC2_Art999",
        titulo="MOTOR DE PRUEBA",
        fuente="asme_pcc/pcc_2/p2_welded_repairs/art_999/art_999.json",
        secciones=(
            M.Seccion("1. DATOS DE ENTRADA", filas=(
                M.Fila("P", "Presion de diseno", magnitud="pres", tipo=M.ENTRADA,
                       ejemplo=20,
                       cita=M.Cita("art_999.json", bloque=12, clausula="999-3.2")),
                M.Fila("dos_P", "El doble", magnitud="pres", tipo=M.FORMULA,
                       formula="={P}*2",
                       cita=M.Cita("art_999.json", bloque=12, clausula="999-3.2")),
            )),
        ),
    )


def test_cada_fila_recibe_una_direccion_en_orden():
    dirs = M.resolver_direcciones(_motor_minimo())
    assert dirs["P"] == "$D$6"
    assert dirs["dos_P"] == "$D$7"


def test_la_formula_se_escribe_por_nombre_y_sale_por_direccion():
    dirs = M.resolver_direcciones(_motor_minimo())
    assert M.sustituir_nombres("={P}*2", dirs) == "=$D$6*2"


def test_un_nombre_que_no_existe_aborta():
    dirs = M.resolver_direcciones(_motor_minimo())
    with pytest.raises(SystemExit, match="no existe"):
        M.sustituir_nombres("={NO_EXISTE}+1", dirs)
```

- [ ] **Paso 2: correrla y ver que falla**

Ejecutar: `python -m pytest test_motor_declarado.py -q`
Esperado: FALLA con `ModuleNotFoundError: No module named 'motor_declarado'`.

- [ ] **Paso 3: escribir el mínimo que la hace pasar**

```python
# motor_declarado.py
"""Chasis de un motor de calculo DECLARADO.

Un motor nuevo no se escribe celda a celda: se DECLARA -que secciones tiene, que
filas, que se teclea y que se calcula- y este modulo lo convierte en hoja. La
diferencia que importa es que las formulas se escriben POR NOMBRE ({P}, {OD}) y
las direcciones las asigna el chasis: la clase de error que obligo al mapa de
filas de la Fase 3 -y que dejo tres defectos latentes- deja de existir.

No importa build_db_materiales: el builder importa este modulo, no al reves.
"""
from __future__ import annotations

import re
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
    material: object = None      # Material o None
    verificaciones: tuple = ()
    dictamen: object = None      # Dictamen o None
    pasos: tuple = ()
    especificaciones: tuple = () # ((grupo, (Especificacion, ...)), ...)


_RE_NOMBRE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]*)\}")


def resolver_direcciones(motor):
    """clave -> direccion absoluta, asignando las filas en orden de lectura."""
    dirs, fila = {}, FILA_PRIMERA_BANDA
    col = chr(ord("A") + COL_PRIMER_CASO - 1)
    for seccion in motor.secciones:
        fila += 1                        # la banda de la seccion
        for f in seccion.filas:
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
```

- [ ] **Paso 4: correrla y ver que pasa**

Ejecutar: `python -m pytest test_motor_declarado.py -q`
Esperado: 3 passed.

- [ ] **Paso 5: commit**

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/motor_declarado.py \
        outputs/Base_Datos_Materiales_ASME/scripts/test_motor_declarado.py
git commit -m "Chasis de motor declarado: tipos y formulas por nombre"
```

---

## Tarea 2 — El guardia de procedencia

**Archivos:**
- Modificar: `scripts/motor_declarado.py`
- Modificar: `scripts/test_motor_declarado.py`

**Interfaces:**
- Consume: `Motor`, `Fila`, `Cita` (Tarea 1).
- Produce: `comprobar_procedencia(motor, resources) -> list[str]` — devuelve la tabla de
  trazabilidad; **aborta** si una fila de cálculo no la trae o si la cita no existe en el
  JSON que dice.

- [ ] **Paso 1: escribir la prueba que falla**

```python
def test_una_fila_de_calculo_sin_cita_aborta(tmp_path):
    motor = _motor_minimo()._replace(secciones=(
        M.Seccion("1. DATOS", filas=(
            M.Fila("x", "Sin procedencia", tipo=M.FORMULA, formula="=1+1"),
        )),))
    with pytest.raises(SystemExit, match="sin procedencia"):
        M.comprobar_procedencia(motor, tmp_path)


def test_una_cita_a_un_bloque_que_no_existe_aborta(tmp_path):
    import json
    d = tmp_path / "asme_pcc" / "pcc_2" / "p2_welded_repairs" / "art_999"
    d.mkdir(parents=True)
    (d / "art_999.json").write_text(
        json.dumps({"blocks": [{"type": "paragraph", "text": "uno"}]}),
        encoding="utf-8")
    motor = _motor_minimo()   # su cita apunta al bloque 12, que no existe
    with pytest.raises(SystemExit, match="bloque 12"):
        M.comprobar_procedencia(motor, tmp_path)
```

- [ ] **Paso 2: correrlas y ver que fallan**

Ejecutar: `python -m pytest test_motor_declarado.py -q`
Esperado: FALLA con `AttributeError: module 'motor_declarado' has no attribute 'comprobar_procedencia'`.

- [ ] **Paso 3: escribir el mínimo que las hace pasar**

```python
import json
from pathlib import Path


def comprobar_procedencia(motor, resources):
    """Tabla de trazabilidad del motor. Aborta si algo no se puede rastrear.

    Un numero que nadie puede rastrear no se construye: es la Regla n.1 hecha
    mecanismo. Se comprueba ademas que el bloque citado EXISTE en el JSON, el
    mismo guardia que la §9 de verificar.py hace con MAP_Grupo.
    """
    resources = Path(resources)
    cache, tabla = {}, []
    for seccion in motor.secciones:
        for f in seccion.filas:
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
```

- [ ] **Paso 4: correrlas y ver que pasan**

Ejecutar: `python -m pytest test_motor_declarado.py -q`
Esperado: 5 passed.

- [ ] **Paso 5: commit**

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/motor_declarado.py \
        outputs/Base_Datos_Materiales_ASME/scripts/test_motor_declarado.py
git commit -m "Chasis: el guardia de procedencia, la Regla n.1 hecha mecanismo"
```

---

## Tarea 3 — Emitir la hoja: secciones, filas y estilos

**Archivos:**
- Modificar: `scripts/motor_declarado.py`
- Modificar: `scripts/test_motor_declarado.py`

**Interfaces:**
- Consume: `resolver_direcciones`, `sustituir_nombres` (Tarea 1).
- Produce: `emitir_tabla(ws, motor, helpers) -> dict` con `{"direcciones": …,
  "ultima_fila": …}`. `helpers` es el paquete de funciones de estilo que pasa el builder
  (ver Tarea 4), para que el chasis **no importe** `build_db_materiales`.

- [ ] **Paso 1: escribir la prueba que falla**

```python
def test_la_hoja_sale_con_las_bandas_y_las_filas_en_orden():
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    res = M.emitir_tabla(ws, _motor_minimo(), _helpers_de_prueba())
    assert ws["A5"].value == "1. DATOS DE ENTRADA"
    assert ws["A6"].value == "Presion de diseno"
    assert ws["D6"].value == 20                 # el ejemplo del caso precargado
    assert ws["A7"].value == "El doble"
    assert ws["D7"].value == "=$D$6*2"          # la formula, ya con direcciones
    assert res["ultima_fila"] == 7


def test_la_clausula_llega_a_la_columna_de_referencia():
    """La procedencia no se queda en la tabla de trazabilidad: el ingeniero
    tiene que ver de que clausula sale el numero SIN salir de la fila."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    M.emitir_tabla(ws, _motor_minimo(), _helpers_de_prueba())
    assert ws["G6"].value == "999-3.2"


def test_la_entrada_nace_desbloqueada_y_el_calculo_no():
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    M.emitir_tabla(ws, _motor_minimo(), _helpers_de_prueba())
    assert ws["D6"].protection.locked is False   # entrada
    assert ws["D7"].protection.locked is not False
```

`_helpers_de_prueba()` devuelve un `Helpers` con funciones mínimas que solo escriben el
valor, para que la prueba del chasis no dependa del sistema visual del libro:

```python
def _helpers_de_prueba():
    from openpyxl.styles import Protection

    def banda(ws, fila, texto):
        ws.cell(fila, 1, texto)

    def rotulo(ws, fila, texto, simbolo, unidad, referencia="", comentario=""):
        ws.cell(fila, 1, texto)
        if referencia:
            ws.cell(fila, 7, referencia)

    def entrada(ws, celda, valor):
        ws[celda] = valor
        ws[celda].protection = Protection(locked=False)

    def calculo(ws, celda, formula):
        ws[celda] = formula

    return M.Helpers(banda=banda, rotulo=rotulo, entrada=entrada, calculo=calculo)
```

- [ ] **Paso 2: correrlas y ver que fallan**

Ejecutar: `python -m pytest test_motor_declarado.py -q`
Esperado: FALLA con `AttributeError: module 'motor_declarado' has no attribute 'Helpers'`.

- [ ] **Paso 3: escribir el mínimo que las hace pasar**

```python
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


def emitir_tabla(ws, motor, helpers):
    dirs = resolver_direcciones(motor)
    col = chr(ord("A") + COL_PRIMER_CASO - 1)
    fila = FILA_PRIMERA_BANDA
    for seccion in motor.secciones:
        helpers.banda(ws, fila, seccion.titulo)
        fila += 1
        for f in seccion.filas:
            # La procedencia llega a la HOJA, no solo a la tabla de
            # trazabilidad: la clausula a la columna de referencia y la
            # explicacion al comentario de la celda de valor. Son las otras dos
            # de las tres formas que pide el diseno.
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
```

- [ ] **Paso 4: correrlas y ver que pasan**

Ejecutar: `python -m pytest test_motor_declarado.py -q`
Esperado: 7 passed.

- [ ] **Paso 5: commit**

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/motor_declarado.py \
        outputs/Base_Datos_Materiales_ASME/scripts/test_motor_declarado.py
git commit -m "Chasis: emite la tabla del motor sin depender del builder"
```

---

## Tarea 4 — Enchufar el chasis al libro, con los pases transversales

**Archivos:**
- Modificar: `scripts/build_db_materiales.py` (añadir `build_motor_declarado()`; **no se
  reescribe nada de lo existente**)
- Crear: `scripts/motores/__init__.py`
- Modificar: `scripts/test_dashboard.py`

**Interfaces:**
- Consume: `emitir_tabla`, `comprobar_procedencia` (Tareas 2-3).
- Produce: `build_motor_declarado(wb, motor, b313, iid1a, iidb, fac_info, rangos, b3610,
  b3619, *, b313c, iid1ac, iidbc, umbrales, resources) -> Worksheet`.
- Produce: `MOTORES_DECLARADOS: tuple[Motor, ...]` en `motores/__init__.py`.

- [ ] **Paso 1: añadir el registro, vacío**

```python
# motores/__init__.py
"""Los motores de calculo DECLARADOS del libro.

Uno por articulo de ASME PCC-2. El Art. 212 y el Art. 206 NO estan aqui: se
escribieron a mano antes del chasis y son lo unico verificado celda a celda
contra su oracle, asi que se quedan como estan (ver el diseno de la skill).
"""
MOTORES_DECLARADOS = ()
```

- [ ] **Paso 2: escribir la prueba que falla**

```python
# test_dashboard.py
class TestMotoresDeclarados:
    """Todo motor declarado cumple lo mismo que los dos escritos a mano."""

    def _hojas(self, wb):
        from motores import MOTORES_DECLARADOS
        return [(m, wb[m.hoja]) for m in MOTORES_DECLARADOS if m.hoja in wb.sheetnames]

    def test_todos_estan_en_el_libro_y_en_la_navegacion(self, wb):
        from motores import MOTORES_DECLARADOS
        for m in MOTORES_DECLARADOS:
            assert m.hoja in wb.sheetnames, m.hoja
            assert m.hoja in B.NAVEGABLES, m.hoja

    def test_cada_entrada_es_editable_y_entra_en_el_reinicio(self, wb):
        for m, ws in self._hojas(wb):
            man = set(TestReinicioDeEntradas()._manifiesto(ws))
            for seccion in m.secciones:
                for f in seccion.filas:
                    if f.tipo == "formula":
                        continue
                    celda = B.direccion_de(m, f.clave)
                    assert ws[celda].protection.locked is False, f"{m.hoja}!{celda}"
                    assert celda.replace("$", "") in man, f"{m.hoja}!{celda}"
```

- [ ] **Paso 3: correrla y ver que falla**

Ejecutar: `python -m pytest test_dashboard.py -k MotoresDeclarados -q`
Esperado: PASA en vacío (no hay motores declarados todavía). Es el estado correcto: la
prueba empieza a morder en la Tarea 10.

- [ ] **Paso 4: escribir `build_motor_declarado`**

```python
# build_db_materiales.py, junto a build_collar_art206
import motor_declarado as MD
from motores import MOTORES_DECLARADOS


def _helpers_del_libro(ws):
    """Los helpers de estilo del libro, para el chasis.

    Son los MISMOS gestos que usan los dos motores escritos a mano: banda de
    tinta con franja roja, rotulo en mono, campo con linea inferior. El chasis
    no los conoce; se los pasamos.
    """
    def banda(w, fila, texto):
        w.merge_cells(start_row=fila, start_column=1, end_row=fila,
                      end_column=MOTOR_NCOLS)
        c = w.cell(fila, 1, texto)
        c.font, c.fill = Font(name=MACRO, size=11, color=PAPEL), BAND_FILL
        for j in range(1, MOTOR_NCOLS + 1):
            w.cell(fila, j).fill = BAND_FILL
        franja(w, fila, 1, MOTOR_NCOLS)

    def rotulo(w, fila, texto, simbolo, magnitud, referencia="", comentario=""):
        c = w.cell(fila, 1, texto)
        c.font = Font(name=MONO, size=10, color=TINTA)
        if simbolo:
            w.cell(fila, 2, simbolo).font = Font(name=MONO, size=10, color=TINTA)
        if referencia:
            w.cell(fila, MOTOR_NCOLS, referencia).font = Font(
                name=MONO, size=9, color=GRIS)
        if comentario:
            # Solo en la columna de VALOR (F9 del 2026-09-13); el barrido final
            # de aplicar_reglas_de_comentario lo confirma.
            _nota(w.cell(fila, COLS_VALOR_MOTOR[0]), comentario)

    def entrada(w, celda, valor):
        c = w[celda]
        c.value = "" if valor is None else valor
        c.font, c.fill, c.border = IN_F, MOTOR_IN_FILL, CAJA_CAMPO
        c.protection = Protection(locked=False)

    def calculo(w, celda, formula):
        w[celda].value = formula

    return MD.Helpers(banda=banda, rotulo=rotulo, entrada=entrada, calculo=calculo)


def direccion_de(motor, clave):
    """Direccion de una fila declarada. La usan las pruebas y verificar.py."""
    return MD.resolver_direcciones(motor)[clave]


def build_motor_declarado(wb, motor, b313, iid1a, iidb, fac_info, rangos,
                          b3610, b3619, *, b313c, iid1ac, iidbc, umbrales,
                          resources):
    """Construye la hoja de un motor DECLARADO y le aplica los pases del libro."""
    MD.comprobar_procedencia(motor, resources)      # Regla n.1: antes de nada
    if motor.hoja in wb.sheetnames:
        del wb[motor.hoja]
    ws = new_sheet(wb, motor.hoja, motor.titulo, f"Fuente: resources/{motor.fuente}")
    autosize(ws, {"A": 50, "B": 8, "C": 14, "D": 16, "E": 16, "F": 16, "G": 46})
    ws.merge_cells(f"A1:{get_column_letter(MOTOR_NCOLS)}1")
    ws.merge_cells(f"A2:{get_column_letter(MOTOR_NCOLS)}2")
    res = MD.emitir_tabla(ws, motor, _helpers_del_libro(ws))
    build_leyenda_motor(ws)
    aplicar_leyenda_motor(ws)
    build_reinicio_motor(ws)
    preparar_impresion(ws, MOTOR_NCOLS)
    ws.protection.password = "0000"
    ws.protection.sheet = True
    return ws
```

- [ ] **Paso 5: llamarlo en `main()`**

```python
    # build_db_materiales.py, junto a build_collar_art206(...)
    for motor in MOTORES_DECLARADOS:
        build_motor_declarado(wb, motor, b313, iid, iidb, fac, rangos,
                              b3610, b3619, b313c=b313c, iid1ac=iidc,
                              iidbc=iidbc, umbrales=umbrales,
                              resources=a.resources)
```

- [ ] **Paso 6: construir y comprobar que el libro NO cambia**

Ejecutar:
```powershell
python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
```
Esperado: el mismo número de pruebas en verde que antes de la tarea. Con
`MOTORES_DECLARADOS` vacío, el libro tiene que salir **igual**: es la red que protege al
212 y al 206.

- [ ] **Paso 7: commit**

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/
git commit -m "El chasis entra al libro, con el registro vacio: nada cambia todavia"
```

---

## Tarea 5 — Material, unidades, semáforo y dictamen en el chasis

**Archivos:**
- Modificar: `scripts/motor_declarado.py` (tipos `Material`, `Verificacion`, `Dictamen`)
- Modificar: `scripts/build_db_materiales.py` (`build_motor_declarado`)
- Modificar: `scripts/test_motor_declarado.py`

**Interfaces:**
- Consume: `construir_seccion7_material`, `aplicar_unidades_motor`,
  `aplicar_semaforo_motor`, `aplicar_dos_decimales`, `build_manifiesto_cascada`,
  `_sem` (todas ya existen en `build_db_materiales.py`).
- Produce: `Material`, `Verificacion`, `Dictamen` en `motor_declarado.py`; y en
  `construir_seccion7_material`, una clave nueva `"ultima_fila"` en lo que devuelve.

- [ ] **Paso 1: que la sección de material diga dónde termina**

`construir_seccion7_material` ya calcula `nota_row` (`F + 26` con Ej/Ec, `F + 23` sin
ellos). Devolver esa fila es lo único que le falta al chasis para seguir escribiendo
debajo:

```python
    resultado = {"fila_banda": F, "fila_modo_s": F + 2, "fila_temp": F + 3,
                 "ultima_fila": nota_row,          # <-- nuevo
                 "por_columna": por_columna}
```

- [ ] **Paso 2: la prueba de que el 212 y el 206 no se enteran**

Ejecutar: `python -m pytest test_dashboard.py -q`
Esperado: el mismo número en verde. Añadir una clave al diccionario devuelto no cambia
ninguna celda; si algo falla aquí, es que alguien dependía del tamaño del diccionario.

- [ ] **Paso 3: los tipos nuevos del chasis**

```python
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
```

- [ ] **Paso 4: engancharlos en `build_motor_declarado`**

```python
    # despues de emitir la tabla, y antes de la leyenda
    if motor.material is not None:
        refs = construir_seccion7_material(
            ws, res["ultima_fila"] + 2, b313, iid1a, iidb, fac_info, rangos,
            columnas=motor.material.columnas,
            modo_cell=direccion_de(motor, "modo") if motor.aplicacion else "$D$11",
            temp_fuente_cell=direccion_de(motor, "temperatura"),
            incluir_ej_ec=motor.material.incluir_ej_ec,
            destino_st="la seccion de parametros de calculo",
            mapa_citas={},                  # el chasis no remapea filas
            b313c=b313c, iid1ac=iid1ac, iidbc=iidbc,
            unidad_cell=direccion_de(motor, "unidad"))
        if motor.material.semilla:
            sembrar_cascada(ws, wb["DB_B31_3"], motor.material.semilla,
                            motor.material.columnas[0][0],
                            res["ultima_fila"] + 2)
    aplicar_unidades_motor(ws, _unidades_de(motor), es_si)
    aplicar_reglas_de_comentario(ws, _reglas_de_comentario_de(motor))
    semaforo = _semaforo_de(motor)
    aplicar_semaforo_motor(ws, semaforo, fila_dictamen)
    aplicar_dos_decimales(ws, semaforo)
    build_manifiesto_cascada(ws, motor.material.columnas, fila_material, {},
                             direccion_de(motor, "unidad"))
```

Las tres tablas **se derivan** de la declaración, no se escriben:

```python
def _unidades_de(motor):
    """fila -> clase de magnitud, para aplicar_unidades_motor."""
    dirs = MD.resolver_direcciones(motor)
    return {int(dirs[f.clave].split("$")[-1]): f.magnitud
            for s in motor.secciones for f in s.filas if f.magnitud}


def _reglas_de_comentario_de(motor):
    """Una regla por seccion, siempre a las columnas de VALOR (F9 2026-09-13)."""
    dirs = MD.resolver_direcciones(motor)
    reglas, cols = [], "".join(chr(ord("A") + c - 1) for c in COLS_VALOR_MOTOR)
    for s in motor.secciones:
        if not s.filas:
            continue
        filas = [int(dirs[f.clave].split("$")[-1]) for f in s.filas]
        reglas.append((min(filas), max(filas), cols[:len(motor.casos)], "D"))
    return tuple(reglas)


def _semaforo_de(motor):
    """celda -> (favorables, avisos), con _sem, para aplicar_semaforo_motor."""
    dirs = MD.resolver_direcciones(motor)
    return {dirs[v.clave].replace("$", ""): _sem(v.favorables, v.avisos)
            for v in motor.verificaciones}
```

`es_si` es la condición del selector —`f'{direccion_de(motor, "unidad")}="SI"'`—,
`fila_dictamen` la fila del dictamen y `fila_material` la que devolvió
`construir_seccion7_material`. Las tres salen de la misma resolución de direcciones: en
el chasis **no hay ninguna fila escrita a mano**.

- [ ] **Paso 5: correr todo**

Ejecutar: `python -m pytest test_motor_declarado.py test_dashboard.py -q`
Esperado: todo en verde, y el libro sin cambios (registro aún vacío).

- [ ] **Paso 6: commit**

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/
git commit -m "Chasis: material, unidades, semaforo y dictamen derivados de la declaracion"
```

---

## Tarea 6 — Pasos del flujo, especificaciones e instrucciones

**Archivos:**
- Modificar: `scripts/motor_declarado.py` (tipos `Paso`, `Especificacion`)
- Modificar: `scripts/build_db_materiales.py`

**Interfaces:**
- Consume: `build_especificaciones`, `build_instrucciones_motor`,
  `build_botones_documentos`, `ocultar_filas_en_blanco` (ya existen).
- Produce: `Paso`, `Especificacion`; y `build_documentos_declarados(wb, motor)`.

- [ ] **Paso 1: los tipos**

```python
class Paso(NamedTuple):
    """Un bloque del anexo del flujo: el recorrido de la reparacion."""
    numero: int
    titulo: str
    clausula: str
    filas: tuple = ()


class Especificacion(NamedTuple):
    concepto: str
    texto: str               # texto fijo, o formula con {nombres}
    clausula: str
    editable: bool = False
    cita: Cita | None = None
```

- [ ] **Paso 2: emitirlos**

Los pasos se escriben con el mismo `emitir_tabla`, detrás de la última sección y con una
banda propia por paso. Las dos pestañas reutilizan las funciones que ya existen:

```python
def build_documentos_declarados(wb, motor):
    """Las dos pestanas del articulo, desde la declaracion."""
    bloques = tuple(
        (grupo, tuple((e.concepto, e.texto, e.clausula) +
                      (("EDITABLE",) if e.editable else ())
                      for e in specs))
        for grupo, specs in motor.especificaciones)
    build_especificaciones(wb, f"Espec_PCC2_Art{motor.articulo}", motor.hoja,
                           f"ESPECIFICACIONES TECNICAS — {motor.titulo}",
                           f"Fuente: resources/{motor.fuente}", bloques)
    build_instrucciones_motor(wb, f"Instruc_PCC2_Art{motor.articulo}", motor.hoja,
                              f"INSTRUCCIONES DE USO — {motor.titulo}",
                              "Que hace el motor y que se rellena en cada seccion.",
                              _prosa_de(motor))
    build_botones_documentos(wb[motor.hoja],
                             espec=f"Espec_PCC2_Art{motor.articulo}",
                             instr=f"Instruc_PCC2_Art{motor.articulo}")
```

`_prosa_de` arma la parte escrita a mano de las instrucciones **desde la propia
declaración**, y la común —color, unidades, semáforo, reinicio— la trae `_prosa_comun`,
que ya existe:

```python
def _prosa_de(motor):
    """Lo que el motor no puede decir de si mismo, desde la declaracion."""
    propia = (
        ("QUE HACE ESTE MOTOR  //  ASME PCC-2 Art. " + motor.articulo, (
            ("Para que sirve", motor.descripcion),
            ("Codigo y alcance", motor.alcance),
            ("De donde sale cada numero",
             f"Todo valor normativo de esta hoja sale de resources/{motor.fuente}, "
             f"y cada celda de calculo cita su clausula en la columna de notas."),
            ("El orden en que se rellena",
             " -> ".join(s.titulo for s in motor.secciones)),
        )),
    )
    return propia + _prosa_comun(motor.hoja, direccion_de(motor, "unidad"),
                                 _fila_dictamen(motor))
```

- [ ] **Paso 3: correr y commitear**

Ejecutar: `python -m pytest test_motor_declarado.py test_dashboard.py -q`

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/
git commit -m "Chasis: pasos del flujo y las dos pestanas del articulo"
```

---

## Tarea 7 — Navegación genérica

**Archivos:**
- Modificar: `scripts/build_db_materiales.py` (`ARBOL`)
- Modificar: `scripts/vba/mod_nav.vba`
- Modificar: `scripts/test_dashboard.py`

**Interfaces:**
- Consume: `MOTORES_DECLARADOS`, `Nodo`, `_hoja_final`.
- Produce: los nodos de artículo derivados del registro.

- [ ] **Paso 1: derivar los nodos del registro**

```python
def _nodo_de_articulo(motor):
    """El articulo es un NIVEL con sus tres documentos, igual que el 212 y el 206."""
    return Nodo(
        titulo=f"ART. {motor.articulo} · {motor.corto}",
        corto=f"ART. {motor.articulo}",
        subtitulo=f"{T_CAL} · ASME PCC-2 Art. {motor.articulo}",
        lineas=(motor.descripcion, "Motor, especificaciones e instrucciones"),
        hoja=f"NAV_CAL_ART{motor.articulo}",
        banda="DOCUMENTOS DE ESTE ARTICULO",
        hijos=(
            _hoja_final("MOTOR DE CALCULO", motor.descripcion, motor.alcance, motor.hoja),
            _hoja_final("ESPECIFICACIONES TECNICAS", "Fabricacion, examen y prueba",
                        motor.clausulas_espec, f"Espec_PCC2_Art{motor.articulo}"),
            _hoja_final("INSTRUCCIONES DE USO", "Que se rellena y que calcula cada celda",
                        "Guia celda a celda del motor",
                        f"Instruc_PCC2_Art{motor.articulo}"),
        ))
```

y en `ARBOL`, detrás de los dos nodos escritos a mano:

```python
                                        *[_nodo_de_articulo(m)
                                          for m in MOTORES_DECLARADOS],
```

**Ninguna línea de estado** en `lineas` (regla del F9).

- [ ] **Paso 2: la prueba de sincronía con el VBA**

`TestSincroniaPythonVba::test_misma_lista_de_hojas_navegables` ya compara
`HojasNavegables()` con `NAVEGABLES`. Al añadir un motor declarado, la prueba falla hasta
que se añaden sus tres líneas al VBA **en preorden**. Eso es lo que se quiere: el VBA es la
única lista que no se deriva.

- [ ] **Paso 3: correr y commitear**

Ejecutar: `python -m pytest test_dashboard.py -q`

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/
git commit -m "Navegacion: el arbol deriva un nivel por articulo declarado"
```

---

## Tarea 8 — La sección de `verificar.py` para motores declarados

**Archivos:**
- Modificar: `scripts/verificar.py`

**Interfaces:**
- Consume: `MOTORES_DECLARADOS`, `direccion_de`, `comprobar_procedencia`.
- Produce: la sección `§12. Motores declarados`, sumada al total.

- [ ] **Paso 1: escribir la sección**

Comprueba cuatro cosas, y ninguna de ellas es la ingeniería:

```python
    log("## 12. Motores declarados (maquinaria, no ingenieria)")
    log("")
    dec_bad = 0
    for motor in B.MOTORES_DECLARADOS:
        ws = wb0[motor.hoja]
        # 1. El caso semilla recalcula sin errores de Excel.
        for fila in ws.iter_rows(max_col=B.MOTOR_NCOLS):
            for c in fila:
                if isinstance(c.value, str) and c.value.startswith("#"):
                    dec_bad += 1
                    log(f"| {motor.hoja}!{c.coordinate} | {c.value} | FALLO |")
        # 2. Toda cita existe en el JSON que dice (aborta si no).
        MD.comprobar_procedencia(motor, RES)

        # 3. Todo veredicto PINTA. Se reutiliza el mismo recorrido de la §6h: la
        # regla puede estar escrita y no pintar -pasó con todo el libro- y eso
        # solo lo dice Excel.
        for celda, (favorables, avisos) in B._semaforo_de(motor).items():
            r = hws.Range(celda)
            texto = str(r.Text).strip()
            esperado = (B.VERDE if texto in favorables
                        else B.AMBAR if any(texto.startswith(a) for a in avisos)
                        else B.ROJO if texto else None)
            if esperado is None:
                continue
            real = _rgb(r.DisplayFormat.Interior.Color)
            if real != esperado:
                dec_bad += 1
                log(f"| {motor.hoja}!{celda} | {texto[:22]} | {real} | "
                    f"{esperado} | FALLO |")

        # 4. El dictamen no contradice a sus propias verificaciones. No se
        # compara contra el literal "APTO" -una decision de ingenieria mueve ese
        # literal y no es un fallo-, sino contra lo que implican los criterios
        # que su propio AND consulta, igual que la §7.
        veredictos = [str(hws.Range(B.direccion_de(motor, c).replace("$", "")).Text
                          ).strip()
                      for c in motor.dictamen.verificaciones]
        implica_apto = all(v in ("CUMPLE",) for v in veredictos)
        dictado = str(hws.Range(B.direccion_de(motor, "dictamen").replace("$", "")
                                ).Text).strip()
        if implica_apto != (dictado == "APTO"):
            dec_bad += 1
            log(f"| {motor.hoja} dictamen | {dictado} | contradice a "
                f"{veredictos} | — | FALLO |")
```

- [ ] **Paso 2: sumarla al total**

```python
    total = (... + dec_bad)
```

y su fila en la tabla de resultado. **Sin esto la sección imprime y no bloquea**, que es
el fallo que ya tuvo `us_bad`.

- [ ] **Paso 3: correr y commitear**

Ejecutar: `python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm`
Esperado: 0 fallos, con la sección 12 en la tabla.

```bash
git add outputs/Base_Datos_Materiales_ASME/scripts/verificar.py
git commit -m "verificar.py §12: la maquinaria de todo motor declarado"
```

---

## Tarea 9 — La skill

**Archivos:**
- Crear: `.agents/skills/motor_pcc2/SKILL.md`
- Crear: `.agents/skills/motor_pcc2/referencias/reglas_del_libro.md`
- Crear: `.agents/skills/motor_pcc2/referencias/plantilla_declaracion.py`

- [ ] **Paso 1: `SKILL.md`**

Con el frontmatter que exige el formato:

```markdown
---
name: motor_pcc2
description: Use when building a calculation engine for an ASME PCC-2 article in this repo — reads the article from resources/, stops for the engineer to decide which equations govern, then generates the declaration, the sheet, the tabs, the tests and the verification. Triggers on "motor del articulo", "nuevo motor PCC-2", "/motor_pcc2 <n>".
---
```

y el cuerpo: las nueve fases, la parada de la Fase 1, el formato del inventario que se le
presenta al ingeniero, y los tres casos de `resources/` que ya costaron caro.

- [ ] **Paso 2: las referencias**

`reglas_del_libro.md` es el catálogo, no una paráfrasis: las 14 reglas de diseño, la Regla
nº 1, las tres del flujo de la cascada, el semáforo con `bgColor`, los comentarios solo en
columna de valor, el ancla del VML, los dos decimales, y las trampas de openpyxl ya
pagadas (relleno de fusionado, `_txt_celda`, tamaño de comentario al clonar).

- [ ] **Paso 3: commitear**

```bash
git add .agents/skills/motor_pcc2/
git commit -m "Skill motor_pcc2: las nueve fases y el catalogo de reglas del libro"
```

---

## Tarea 10 — El primer motor real, de punta a punta

**Archivos:**
- Crear: `scripts/motores/art_XXX.py` (el artículo lo elige el ingeniero)
- Modificar: `scripts/motores/__init__.py`
- Modificar: `scripts/vba/mod_nav.vba`

> **Esta tarea NO se puede ejecutar sola.** Su Fase 1 es la parada de criterio de
> ingeniería: qué ecuaciones gobiernan, qué se teclea y qué se verifica. El ejecutor
> presenta el inventario y **espera**.

- [ ] **Paso 1: correr la skill sobre el artículo elegido y presentar el inventario**
- [ ] **Paso 2: el ingeniero marca qué entra** ← parada
- [ ] **Paso 3: generar `motores/art_XXX.py` y registrarlo**
- [ ] **Paso 4: añadir sus tres hojas a `HojasNavegables()` en preorden**
- [ ] **Paso 5: construir, `pytest`, `verificar.py`**
- [ ] **Paso 6: exportar la hoja a PDF y mirarla** — es la única evidencia válida del
      aspecto en este libro
- [ ] **Paso 7: comprobar que el 212 y el 206 salen idénticos**

```powershell
python -c "import openpyxl; a=openpyxl.load_workbook('antes.xlsm'); b=openpyxl.load_workbook('..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm'); ..."
```
Comparar celda a celda A..G de las dos hojas contra una copia del libro anterior. Cero
diferencias. Es el criterio de aceptación nº 3.

- [ ] **Paso 8: commit**

---

## Tarea 11 — Un artículo sin ecuaciones

> Criterio de aceptación nº 4 del diseño. Sin esta tarea, el chasis solo estaría
> probado contra artículos de diseño por fórmula, que son la minoría de los 38.

**Archivos:**
- Crear: `scripts/motores/art_214.py` (o el artículo de procedimiento que elija el
  ingeniero: 214 tratamiento térmico en campo, 310 hot bolting, 307 enderezado)
- Modificar: `scripts/motores/__init__.py`, `scripts/vba/mod_nav.vba`

- [ ] **Paso 1: correr la skill y presentar el inventario** — se espera que la Fase 1
      diga, con todas las letras, que **el artículo no publica ecuaciones**, y que lo que
      trae son valores tabulados (temperaturas, rampas, tiempos de permanencia) y
      requisitos
- [ ] **Paso 2: el ingeniero marca qué entra** ← parada
- [ ] **Paso 3: declarar el motor SIN sección de cálculo**, con el peso en los pasos y en
      las verificaciones —que siguen teniendo semáforo y dictamen—
- [ ] **Paso 4: comprobar que el chasis no exige lo que no hay**

```bash
python -m pytest test_motor_declarado.py test_dashboard.py -q
```
Un motor sin `verificaciones` de fórmula no puede abortar por el guardia de procedencia:
el guardia solo mira las filas de tipo `FORMULA`, y aquí no hay. Si aborta, el guardia
está mal escrito y hay que arreglarlo **aquí**, no en el artículo.

- [ ] **Paso 5: construir, verificar y mirar el PDF**
- [ ] **Paso 6: commit**

---

## Qué encontró la autorrevisión del plan

Antes de darlo por bueno se revisó contra el diseño —cobertura, marcadores de posición y
consistencia de tipos— y salieron **cinco fallos del propio plan**. Se dejan escritos
porque son la clase de fallo que un plan esconde bien:

1. **`Motor` se definía con cinco campos** en la Tarea 1 y las Tareas 5-7 usaban campos
   que no existían (`material`, `especificaciones`, `corto`…). Un campo que aparece a
   mitad de plan es un campo que dos tareas escriben distinto. Ahora se define completo
   desde el principio.
2. **Tres funciones derivadas se citaban sin definirse** (`_unidades_de`, `_semaforo_de`,
   `_reglas_de_comentario_de`). Eran justo las que hacen que las tablas del motor salgan
   de la declaración en vez de escribirse a mano: sin su cuerpo, el plan pedía magia.
3. **`_prosa_de` igual**, en la Tarea 6.
4. **La sección 12 de `verificar.py` dejaba dos comprobaciones como comentario**
   (`# 3. Todo veredicto pinta`, `# 4. El dictamen no contradice…`). Un plan con un
   comentario donde va el código no es un plan: es una nota. Y esas dos son precisamente
   las que este libro aprendió a no dar por supuestas.
5. **La procedencia llegaba a la tabla de trazabilidad pero no a la hoja.** El diseño pide
   las tres formas —columna de referencia, comentario y tabla— y solo estaba la tercera.

Y un hueco de cobertura: ninguna tarea comprobaba el criterio de aceptación nº 4, que un
artículo **sin ecuaciones** produzca un motor útil. Es la **Tarea 11**.

## Notas de riesgo

- **La Tarea 4 es la bisagra.** Mientras `MOTORES_DECLARADOS` esté vacío, el libro tiene
  que salir byte a byte igual. Si en ese punto algo cambia, es que el chasis tocó algo
  que no era suyo — parar y mirar, no seguir.
- **La Tarea 5 toca una función compartida** (`construir_seccion7_material`) que usan los
  dos motores escritos a mano. El cambio es aditivo —una clave más en el diccionario que
  devuelve— y su prueba es que el resto del libro no se entera.
- **La Tarea 10 no es mecánica.** Es donde se descubre si el chasis cubre un artículo de
  verdad. Lo que no encaje se resuelve declarando una sección de filas libres, no
  forzando el artículo dentro del chasis.
