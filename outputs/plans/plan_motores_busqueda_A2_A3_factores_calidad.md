# Dos motores de búsqueda para los factores de calidad — ASME B31.3, Tablas A-2 y A-3

**Documento:** PLAN-BUSC-A2A3-001 · Rev. 0 · 2026-09-07
**Entregable objetivo:** `Motor_de_Calculo_ASME_PCC_Rev3.xlsm` → **43 hojas** (hoy 39)
**Fuente única de verdad:** `resources/ASME B31/ASME B31.3/APPEX/appendix_a/`

---

## ✅ EJECUTADO — 2026-09-07

Ejecutado **después** del plan del Apéndice C, como el propio plan recomienda. El
entregable es `Motor_de_Calculo_ASME_PCC_Rev4.xlsm` con **43 hojas** y **12
navegables** — la composición que el plan preveía (39 + 4 + 1 − 2 = 42) da 43
porque hizo falta una hoja más de la prevista: `DB_Ec_Incremento`, la Tabla
302.3.3-1 como base propia. Meter sus seis filas en `DB_Listas` habría sido más
barato, pero esa hoja está rotulada «no es fuente normativa» y los seis factores Ec
sí lo son.

`verificar.py` devuelve **0 fallos**; §6c recalcula 11 casos en Excel.

### Fase 0 — las tres preguntas, respondidas

1. **Duplicado de la Tabla 302.3.3-1.** Confirmado y resuelto:
   `completar_tabla_302_3_3.py` declara `table_302_3_3_1.json` como canónico y
   registra el SHA-256 del descartado, que además queda marcado con
   `superseded_by`. **Hallazgo adicional:** el patrón es más amplio de lo que el
   plan suponía —**32 pares** con doble prefijo en `CHAPTERS/tables/`, de los que
   **21 son idénticos y 11 no**—, y en los distintos el de **prefijo simple** trae
   las filas ya recompuestas mientras el de prefijo doble conserva los fragmentos
   del corte de línea (`table_302_3_3_2.json` tiene 8 filas legibles frente a las 19
   partidas de su gemelo). Los otros 31 pares **no se tocan**: queda censado en la
   enmienda para que se decida aparte.
2. **Encabezados perdidos.** Recuperados del propio código, del para. 302.3.3(c),
   leído de `chapter_02.json`. Van como `header_derivado` con la frase de la que
   salen; `header` **sigue en null**, porque sigue siendo verdad que el impreso no
   se capturó y el PDF no está en el repo.
3. **¿Existe equivalente para `Ej`?** **Sí**, y no es una tabla aparte: son filas de
   la propia Tabla 302.3.4-1, según el para. 302.3.4(b). **Pero su extracción está
   inservible**: el cuerpo de la tabla se colapsó dentro de los encabezados de
   columna. El motor de A-3 **declara el hueco**, transcribe el párrafo y la Nota
   (1), y remite al folio impreso. No ofrece número.

### Una corrección al plan, tomada de `resources/`

El plan da por hecho que **A451** cita solo la Nota (5) y que su bloque de
incremento debe salir **desactivado**. El código imprime `"(4), (5)"`: cita **las
dos**. Es la única fila de A-2 que lo hace. Tratarla como si solo tuviera la (5)
ocultaría que su 0,90 **también puede subirse**. El motor contempla los cuatro
estados —solo (4), solo (5), las dos, ninguna— derivados de las notas que cita la
fila, y la comprobación manual del plan queda así: `A451` da **0,90** con el aviso
de que su factor ya supone examen **y** con el bloque de incremento **activo**.

También se normaliza el `(Cont'd)` de los rótulos de grupo de A-3 —«Copper and
Copper Alloy (Cont'd)» es el mismo grupo reimpreso al pasar de folio—, declarándolo
fila a fila en la columna `Detalle del incremento`. Sin eso la cascada ofrecería 10
grupos donde el código publica 8.

---

## Contexto

`Ec` y `Ej` son los dos factores de calidad que entran en la ecuación de diseño por
presión del B31.3:

$$t = \frac{P\,D}{2\,(S\,E + P\,Y)}$$

donde `E` es **`Ec`** si el componente es una **fundición** (Tabla A-2) y **`Ej`** si es un
**tubo con junta longitudinal soldada** (Tabla A-3). Equivocar el factor mueve el espesor
requerido hasta un 40 % —de `E = 1,00` en un tubo sin costura a `E = 0,60` en uno soldado
a tope en horno—, así que no es un dato accesorio: es un multiplicador directo del
esfuerzo admisible.

**Qué hay hoy.** Las dos tablas están cargadas en la hoja `MAP_Factores` (151 filas = 24 de
A-2 + 127 de A-3), con especificación, clase/tipo, descripción, factor, notas y grupo de
material. El motor del Art. 212 ya las consulta y sustituyó el `E_j = 1` fijo.

**Qué falta.** No hay forma de *buscar* en ellas. Para saber qué `Ej` corresponde a un
A671 Clase 22 hay que abrir una hoja oculta de 151 filas mezcladas y filtrar a mano. Y,
sobre todo, **el libro no dice lo más importante que dicen esas tablas**: que el factor
publicado es un **mínimo** que puede subirse con examen suplementario.

**Alcance.** Dos motores independientes, uno por tabla, como se pidió. No comparten
cascada porque no responden a la misma pregunta: A-2 se busca por **especificación de
fundición**; A-3, por **especificación y tipo de junta soldada**.

### Estado de los datos, verificado en esta sesión

| Tabla | Filas | Estado |
|---|---:|---|
| **A-2** (Ec) | 24 | **Cotejada fila a fila contra el folio impreso 364: 0 defectos.** Grupos, specs, descripciones, factores y las Notas (1)–(5) completas |
| **A-3** (Ej) | 127 | Cotejada contra los folios 366, 367 y 368. Tenía **38 filas con `Class (or Type)` y `Description` fusionadas**; corregidas por `completar_apendice_a.py` y protegidas por `TestApendiceACorregido` |

Las dos entran a este plan ya auditadas contra el código.

---

## Fase 0 — Tres huecos que cerrar antes de construir (bloqueante)

1. **La Tabla 302.3.3-1 está duplicada en `resources/`.** Existen
   `CHAPTERS/tables/table_302_3_3_1.json` y `CHAPTERS/tables/table_table_302_3_3_1.json`
   con contenido idéntico (mismo `table_id`, mismas 6 filas), y el mismo par para
   `k302_3_3_1`. Es un artefacto de extracción —el prefijo `table_` aplicado dos veces—.
   Hay que quedarse con uno y **declarar cuál**, o el motor citará una fuente ambigua.
2. **Esa tabla perdió sus encabezados de columna**: `columns` trae
   `[{"key": "column_1", "header": null}, {"key": "column_2", "header": null}]`. Sin ellos
   no se puede rotular qué es cada columna. Se recuperan del folio impreso, como se hizo
   con el Apéndice C.
3. **Falta confirmar si existe una tabla equivalente para `Ej`.** La Nota (2) de A-2 remite
   a 302.3.3(c) y a la Tabla 302.3.3-1 para factores **de fundición**. Hay que verificar en
   el capítulo 2 si el B31.3 publica un mecanismo análogo para la junta longitudinal
   (para. 302.3.4) y, si lo publica, en qué tabla. **Si no existe, el motor de A-3 debe
   decirlo explícitamente** en vez de callar: es justo la clase de vacío que la regla nº 1
   obliga a declarar.

Se resuelve con `scripts/completar_tabla_302_3_3.py`, hermano de los tres
`completar_apendice_*.py` ya en el repo: idempotente, cita el folio, no toca valores.

---

## Los dos motores

### `Buscar_Ec_A2` — Factor de calidad de fundición

**Cascada de 3 niveles.** La tabla es pequeña (24 filas) y sus grupos son los que imprime
el código:

| Nivel | Contenido | Opciones |
|---|---|---:|
| 0 · Grupo del código | `Iron`, `Carbon Steel`, `Low and Intermediate Alloy Steel`, `Stainless Steel`, `Copper and Copper Alloy`, `Nickel and Nickel Alloy`, `Aluminum Alloy`, `Titanium and Titanium Alloy` | 8 |
| 1 · Especificación | A47, A48, A126… B367 | 1–8 por grupo |
| 2 · Descripción | «Gray iron castings», «Centrifugally cast pipe»… | 1–2 |

El nivel 2 solo desambigua donde el código repite la especificación en dos grupos
(**A352 aparece en `Carbon Steel` y en `Low and Intermediate Alloy Steel`**, ambas con
`Ec = 0,80`).

### `Buscar_Ej_A3` — Factor de calidad de junta longitudinal

**Cascada de 4 niveles**, porque aquí el factor lo decide el **tipo de junta**, no el
material:

| Nivel | Contenido | Opciones |
|---|---|---:|
| 0 · Grupo del código | 8 grupos, de `Carbon Steel` a `Aluminum Alloy` | 8 |
| 1 · Especificación | API 5L, A53, A106… B210 | — |
| 2 · Clase o tipo | `Type S`/`E`/`F`, `All`, `DW`, `SW`, `12, 22, 32, 42, 52`… o `(sin clase)` | 1–3 |
| 3 · Descripción de la junta | «Seamless pipe», «Electric resistance welded pipe», «Electric fusion welded pipe, 100% radiographed»… | 1–4 |

Es el nivel 3 el que mueve el factor: para el mismo A312 el código publica **1,00**
(sin costura), **1,00** (EFW 100 % radiografiada), **0,85** (doble butt) y **0,80**
(butt simple).

---

## Lo que estos motores hacen distinto — y por qué

**1 · No hay temperatura, ni interpolación, ni curva.** `Ec` y `Ej` son escalares: el
código publica un número por fila, no una función de `T`. Por tanto estos dos motores
**no** reutilizan `interp_value`, ni la banda compacta, ni `append_packed`, ni el
`ScatterChart`. Son la mitad de máquina que los otros siete, y el plan no finge lo
contrario. Lo que sí reutilizan: `new_sheet`, `banda`, `_mrg`, `campo`, `dv_list`,
`_nota`, el patrón de lista dependiente `IF(COUNTIF…)/INDEX/MATCH` de
`build_db_materiales.py:1861-1863`, y la capa de navegación entera.

**2 · El conmutador de unidades es degenerado, y hay que decirlo.** La regla 10 de
`CLAUDE.md` obliga a que todo motor lleve conmutador SI ↔ US. Pero **`Ec` y `Ej` son
adimensionales y el código publica una sola tabla para los dos sistemas**: no hay A-2C ni
A-3C. Aplicar la regla al pie de la letra aquí inventaría una distinción que el código no
hace. Resolución: en el sitio donde los demás motores ponen el conmutador, estos ponen una
celda **fija y no editable** que dice *«FACTOR ADIMENSIONAL — identico en SI y en US
(el codigo publica una sola tabla)»*, con el mismo formato de banda. Se cumple el espíritu
de la regla —el usuario ve siempre en qué sistema está leyendo y nunca un valor
convertido— sin fabricar un interruptor que no gobierna nada.

**3 · El factor publicado es un MÍNIMO, y ese es el dato que hoy falta.** Es la razón de
peso para construir estos motores. Las Notas de A-2 dicen literalmente:

> **(4)** *This casting quality factor **can be enhanced** by supplementary examination in
> accordance with para. 302.3.3(c) and Table 302.3.3-1. The **higher factor** from Table
> 302.3.3-1 **may be substituted** for this factor in pressure design equations.*
>
> **(5)** *This casting quality factor is applicable **only when** proper supplementary
> examination has been performed.*

Así que cada motor lleva un bloque **`3 · FACTOR INCREMENTADO POR EXAMEN SUPLEMENTARIO`**:

- Para `Ec`, una lista con las **6 combinaciones de examen de la Tabla 302.3.3-1** y el
  factor que cada una habilita —0,85 · 0,85 · 0,95 · 0,90 · 1,00 · 1,00—, y el factor
  resultante junto al básico.
- El bloque se **desactiva y lo dice** cuando la fila seleccionada no cita la Nota (4):
  `A426` y `A451` llevan Nota (5), que es la condición inversa —su factor **ya supone** el
  examen y no se puede subir más—. Confundir las dos notas es un error de diseño con
  consecuencia directa en el espesor.
- Para `Ej`, el bloque existe solo si la Fase 0 confirma que el código publica un mecanismo
  equivalente. **Si no lo publica, el bloque dice que no lo publica.**

**4 · Bloque normativo con las notas verbatim.** Cada motor cierra con las notas que cita
la fila seleccionada, transcritas del código, y la advertencia fija de que el factor entra
en `t = PD / (2(SE + PY))` y no en ningún otro sitio.

---

## Layout (12 columnas A..L, como el resto)

```
┌──────────────────────────────────────────────────────────────────────────────┐
│ FACTOR DE CALIDAD DE FUNDICION Ec — ASME B31.3-2024, TABLA A-2               │ NAVY
│ Valores tal como estan impresos · folio 364                                  │
│ ◂ VOLVER AL DASHBOARD                                                        │ fila 3
├──────────────────────────────────────────────────────────────────────────────┤
│ 1 · SELECCION DE LA FUNDICION                                                │ banda BLUE
│   FACTOR ADIMENSIONAL — identico en SI y en US   (celda fija, no editable)   │ D5
│   0 · Grupo del codigo      [ .................. ]                           │ D6
│   1 · Especificacion        [ .................. ]                           │ D7
│   2 · Descripcion           [ .................. ]     [ SELECCION COMPLETA ]│ D8 · G8:I8
├──────────────────────────────────────────────────────────────────────────────┤
│ 2 · FACTOR BASICO                                                            │
│  ┌───────────────┬───────────────┬───────────────┬───────────────┐           │
│  │     0,80      │  adimensional │   Tabla A-2   │   (3), (4)    │           │ KPI
│  │  Ec BASICO    │   sin unidad  │   folio 364   │ notas citadas │           │
│  └───────────────┴───────────────┴───────────────┴───────────────┘           │
├──────────────────────────────────────────────────────────────────────────────┤
│ 3 · FACTOR INCREMENTADO POR EXAMEN SUPLEMENTARIO   (Tabla 302.3.3-1)         │
│   Examen realizado        [ .................. ]   ← 6 combinaciones          │
│   Ec aplicable            [   1,00   ]  ← solo si la fila cita la Nota (4)   │
│   ⚠ Nota (5): este factor YA supone el examen; no admite incremento          │
├──────────────────────────────────────────────────────────────────────────────┤
│ 4 · NOTAS DEL CODIGO — transcritas, las que cita la fila                     │
├──────────────────────────────────────────────────────────────────────────────┤
│ 5 · TRAZABILIDAD — tabla, folio y archivo de resources/                      │
│     Recordatorio: E entra en t = PD / (2(SE + PY)). En ningun otro sitio.    │
└──────────────────────────────────────────────────────────────────────────────┘
```

`Buscar_Ej_A3` es el mismo layout con un nivel más de cascada y la columna
`Clase / Tipo` añadida a la ficha.

---

## Bases

Dos hojas nuevas, `DB_A2_Ec` (24 filas) y `DB_A3_Ej` (127 filas), con las **8 primeras
columnas idénticas a `STRESS_COLS`** —`material_id, Tabla, k0..k4, clave_bi`— para que
`build_listas` y las secciones §2 y §4 de `verificar.py` funcionen **sin tocar una línea**,
igual que se hizo en el plan del Apéndice C. Los niveles no usados (`k3`, `k4` en A-2;
`k4` en A-3) se rellenan con `(no aplica)`.

> **`MAP_Factores` se queda como está.** Contiene los mismos 151 registros y la consulta el
> motor del Art. 212, que está verificado por el caso semilla. Duplicar 151 filas en el
> libro es más barato que tocar un consumidor auditado. Consolidar las tres hojas en una
> sola queda como mejora posterior, no como parte de este plan.

Columnas propias, tras las 8 de contrato: `Grupo impreso · Spec. No. · Clase o tipo ·
Descripcion · Factor · Notas citadas · Admite incremento · Tabla · Folio`.

`Admite incremento` se deriva **de las notas que cita la fila**, no de una heurística:
`(4)` → sí; `(5)` → no, ya lo supone; sin ninguna de las dos → el código no se pronuncia.

---

## Navegación

| Sitio | Cambio |
|---|---|
| `build_db_materiales.py:3187` `NAVEGABLES` | + `Buscar_Ec_A2`, + `Buscar_Ej_A3` (pasa de 9 a 11) |
| `scripts/vba/mod_nav.vba:63-72` | los mismos literales **y en la misma posición** — `TestSincroniaPythonVba` compara listas ordenadas |
| `build_db_materiales.py:3365` `buscadores` | dos tarjetas nuevas. Se pasa de 9 a **11 tarjetas**: la rejilla 3 × 3 se queda corta y hay que ampliar a **4 filas (3+3+3+2)**, subiendo el bloque de KPI y el pie |
| `build_db_materiales.py:3637` `order` | insertar las 4 hojas nuevas |
| `make_vba_seed.py` | **reejecutar**: al entregable viaja `vbaProject.bin`, no el `.vba` |

Resultado: **43 hojas** — 1 visible, 11 `hidden`, 31 `veryHidden`.

> Este plan y el del Apéndice C (`plan_motor_busqueda_apendice_c_b31_3.md`) tocan las
> mismas cuatro listas. Si se ejecutan los dos, el recuento compone: 39 + 4 (este) + 1
> − 2 (aquel) = **42 hojas** y 12 navegables. Conviene ejecutar uno y reconstruir antes de
> empezar el otro.

---

## Verificación

- `verificar.py` §1: `DB_A2_Ec` = 24 filas, `DB_A3_Ej` = 127.
- §2 y §4: añadir las dos hojas a `bases` — sin más cambios, gracias al contrato de las
  8 primeras columnas.
- §3b: `audita_factores()` nueva — compara fila a fila el factor y las notas de cada hoja
  contra `table_a_2.json` y `table_a_3.json`.
- §5 y §8: automáticos.
- **§6 no aplica**: no hay interpolación que recalcular. En su lugar, una comprobación
  propia de que el factor mostrado es exactamente el de la fila, sin redondeo.

`test_build_db.py::TestFactoresDeCalidad` — casos tomados del folio: `A47 → 1,00`;
`A395 → 0,80` con Nota (4); `A426 → 1,00` con Nota (5) y **sin** incremento admisible;
`A451 → 0,90`; `API 5L` sin costura `→ 1,00` y soldado en horno `→ 0,60`; `A312` con sus
cuatro descripciones `→ 1,00 / 1,00 / 0,85 / 0,80`; y que las 6 filas de la Tabla 302.3.3-1
dan `0,85 · 0,85 · 0,95 · 0,90 · 1,00 · 1,00`.

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts
python completar_tabla_302_3_3.py --resources ..\..\..\resources
python make_vba_seed.py
python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm
python -m pytest test_build_db.py test_dashboard.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm
```

`verificar.py` debe seguir devolviendo **0**: los 271 473 valores auditados, la contigüidad,
cero matriz dinámica y el caso semilla `Sa` 138 MPa **APTO** no pueden moverse.

Comprobación manual: seleccionar `Stainless Steel → A451 → Centrifugally cast pipe` y ver
**0,90** con la advertencia de la Nota (5) y el bloque de incremento **desactivado**;
`Iron → A395` y ver **0,80** con el bloque **activo** y `1,00` alcanzable con el examen
`(1) and (3)(a) or (3)(b)`.

---

## Orden de ejecución

1. Fase 0: desduplicar y completar la Tabla 302.3.3-1; resolver si existe equivalente para `Ej`.
2. `build_a2_ec` y `build_a3_ej` → las dos bases, con el contrato de columnas.
3. `build_buscador_factor(...)` — una sola función parametrizada por tabla; los dos motores
   comparten layout y solo difieren en el número de niveles y en la columna `Clase / Tipo`.
4. Navegación: `NAVEGABLES`, VBA, dos tarjetas, `order`, rejilla del Dashboard a 4 filas;
   reejecutar `make_vba_seed.py`.
5. `verificar.py` (§1, §2/§4, §3b) y `test_build_db.py`.
6. `CLAUDE.md` —43 hojas, la excepción declarada a la regla 10 para factores
   adimensionales— e `Instrucciones`.

## Riesgos

| Riesgo | Mitigación |
|---|---|
| Presentar el factor básico como si fuera el definitivo | El bloque 3 es parte del núcleo, no un extra; y el estado «no admite incremento» se deriva de la nota citada, no de una suposición |
| Confundir la Nota (4) con la (5) —incrementable frente a «ya lo supone»— | Test dedicado sobre A395 (4) y A426/A451 (5) |
| Citar una fuente ambigua por el duplicado de la Tabla 302.3.3-1 | Fase 0 bloqueante: se elige un archivo y se declara cuál |
| Inventar un conmutador SI/US que no gobierna nada | Celda fija y rotulada, no un desplegable; documentado como excepción de la regla 10 |
| La rejilla del Dashboard se desborda con 11 tarjetas | Se amplía a 4 filas en el mismo cambio; `test_dashboard.py` comprueba que los botones cubren las hojas navegables |

---

*Herramienta de ingeniería de referencia. Verificar entradas y resultados antes de emitir
para construcción.*
