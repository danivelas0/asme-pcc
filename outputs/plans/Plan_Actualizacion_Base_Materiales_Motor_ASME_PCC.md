# Plan de Actualización — Base de Datos de Materiales
## Motor de Cálculo ASME PCC (`Motor_de_Calculo_ASME_PCC.xlsx`)

**Documento:** PLAN-DB-MAT-001 · Rev. 0
**Fecha:** 2026-09-06
**Autor:** Daniel Velásquez
**Ámbito:** Reemplazar la tabla de esfuerzos admisibles embebida (6 aceros al carbono, T ≤ 40 °C) por dos bases de datos de materiales completas y dependientes de temperatura —una por código— que alimenten por *lookup* a todos los motores de cálculo del libro.
**Unidades:** SI (MPa, mm, °C). **Fuente única de verdad:** carpeta `resources/` (regla §5 de `knowledge/claude.md`).

---

## 0. Resumen ejecutivo y recomendación

El libro actual resuelve el esfuerzo admisible `Sa` con un `INDEX/MATCH` plano sobre 6 materiales a una sola temperatura. Esto limita el motor a agua de enfriamiento a temperatura ambiente y obliga a editar la tabla a mano para cada caso nuevo. La actualización propuesta construye, a partir de los JSON ya extraídos en `resources/`, **dos bases de datos independientes**:

- **`DB_B31_3`** — esfuerzos admisibles básicos en tracción, ASME B31.3-2024, Apéndice A (Tabla A-1 / A-1C + pernería A-2).
- **`DB_BPVC_IID`** — máximos esfuerzos admisibles `S`, ASME BPVC Sección II-D (Métrico) 2025, Tabla 1A (ferrosos) + 1B (no ferrosos) + pernería Tabla 3.

Cada base entrega `S(T)` **interpolado linealmente** entre temperaturas tabuladas, con un **conmutador** que permite forzar el valor tabulado conservador (decisión confirmada). El selector de código existente (`$D$11`) sigue gobernando de qué base lee cada motor.

Este ciclo incluye además la **Fase 2 de propiedades de material** (confirmada): resistencia a la tracción `Su`, fluencia `Sy`, módulo `E`, dilatación térmica y Poisson/densidad, todas desde `resources/`, para habilitar los cálculos que las requieren (p. ej. PCC-2 Art. 202/205 y análisis de dilatación). El alcance del B31.3 abarca los **Apéndices A, B y C completos** (metales y no metales) y reemplaza el factor `E_j` fijo por *lookup* (A-2/A-3).

Cada base es **interactiva y dual SI/US**: una hoja-consulta con **buscador dinámico** (Excel 365/2021) permite escribir una especificación, conmutar entre **unidades SI y US Customary** y ver al instante todos los datos del material tal como los organiza el código, ahorrando el recorrido manual de miles de filas (§6, §3.4). Ambos códigos tienen doble unidad **nativa**: B31.3 con sus tablas "C" y la II-D con las carpetas métrica y `bpvc_ii_d_customary_2025`. La pestaña `Instrucciones` se reescribe para explicar en detalle cómo usar cada base (§7).

**Recomendación de ubicación (tu pregunta abierta): alojar ambas bases como hojas nuevas dentro del mismo libro** (`DB_B31_3`, `DB_BPVC_IID`), no como archivos externos. Razones:

1. **Portabilidad y trazabilidad.** El libro viaja con cada memoria de cálculo; una base externa introduce enlaces que se rompen al mover/compartir el archivo —especialmente sobre rutas de *Google Drive Streaming*— y complican la auditoría (regla §4 de `knowledge/claude.md`: todo coeficiente debe trazar a su archivo de `resources/`).
2. **Volumen manejable.** El alcance del *lookup* de esfuerzos es ~3.200 filas de material × ~30 columnas de temperatura. Excel lo maneja con holgura; el peso estimado del libro sube a ~3–6 MB, aceptable para un motor maestro.
3. **Mitigación del costo.** Las hojas-DB se escriben **solo con valores** (sin fórmulas ni funciones volátiles) para no penalizar el recálculo; los motores las consultan con `INDEX/MATCH` no volátil.

Fallback documentado: si el libro se vuelve pesado o quieres versionar la DB por separado, migrar las hojas-DB a un archivo `DB_Materiales_ASME.xlsx` referenciado, sin cambiar el esquema (Fase 6, opcional).

---

## 1. Diagnóstico del estado actual

**Libro:** 3 hojas — `Instrucciones`, `Datos_Ref`, `Parche_PCC2_Art212`.

**`Datos_Ref`, bloque B (filas 41–47):** esfuerzos admisibles `S` en MPa para 6 materiales (A106 Gr.B, A516 Gr.70, A105, A285 Gr.C, A333 Gr.6, A53 Gr.B), dos columnas: `Sa B31.3` (Tabla A-1, factor 3,0) y `Sa VIII-1` (II-D Tabla 1A, factor 3,5). **Válido solo a T ≤ 40 °C.**

**Motor `Parche_PCC2_Art212` (celdas D39/D40):**

```
D39 (Sa_c) = IF($D$11=1, INDEX(Datos_Ref!$B$42:$B$47, MATCH($D$23,Datos_Ref!$A$42:$A$47,0)),
                          INDEX(Datos_Ref!$C$42:$C$47, MATCH($D$23,Datos_Ref!$A$42:$A$47,0)))
D40 (Sa_b) = ídem con $D$22 (material del metal base)
D41 (Sa)   = MIN(D39,D40)
```

**Limitaciones a resolver:**

- Sin dependencia de temperatura: `D25 = T` (°C) no entra en el cálculo de `Sa`.
- Universo de 6 materiales; sin aceros inoxidables, aleados, no ferrosos ni pernería.
- Clave de *lookup* = nombre de material simple; no distingue forma de producto, espesor, clase/temple ni P-No., que en las tablas reales diferencian filas con el mismo *spec/grade*.
- No hay control de temperatura máxima admisible por material (riesgo de extrapolar fuera de rango).

---

## 2. Objetivo y alcance de datos

### 2.1 Fuentes (ya disponibles en `resources/`, verificadas)

**ASME B31.3-2024 — Apéndice A** (`resources/ASME B31/ASME B31.3/APPEX/`):

| Archivo | Contenido | Filas | Uso en la DB |
|---|---|---|---|
| `appendix_a/table_a_1.json` | Esfuerzos admisibles básicos en tracción, metales (SI) | 1.146 | **Núcleo `DB_B31_3`** |
| `appendix_a/table_a_1c.json` | Companion A-1C (continuación) | — | Fusionar con A-1 |
| `appendix_a/table_a_4.json` | **Design Stress Values for Bolting** (SI), por T | 142 | Pernería `DB_B31_3` |
| `appendix_a/table_a_2.json` | Factor de calidad de fundición `Ec` | 24 | `MAP_Factores` (§4.5) |
| `appendix_a/table_a_3.json` | Factor de calidad de junta long. soldada `Ej` | 127 | `MAP_Factores` (§4.5) |
| `appendix_c/table_c_1.json` | Dilatación térmica, metales (SI), por T | 52 | Propiedades B31.3 (Fase 2) |
| `appendix_c/table_c_3.json` | Módulo de elasticidad `E`, metales (SI), por T | 75 | Propiedades B31.3 (Fase 2) |
| `appendix_c/table_c_2.json` | Dilatación térmica, **no metales** | 44 | Propiedades no metálicas (Fase 2) |
| `appendix_c/table_c_4.json` | Módulo de elasticidad, **no metales** | 30 | Propiedades no metálicas (Fase 2) |
| `appendix_b/table_b_1.json` … `table_b_6.json` | **Apéndice B** — esf. diseño hidrostático (HDS) y presión admisible, tuberías **no metálicas** (termoplásticos/RTR) | 36 (B-1) + B-2…B-6 | `DB_B31_B` (no metálicos) |
| `appendix_b/spec_index_b.json` | Índice de especificaciones no metálicas | — | Validación de cobertura |
| `appendix_a/notes_tables_a_1_a_1c.json` / `notes_tables_a_4_a_4c.json` | Notas A-1/A-1C y A-4 | — | Columna `notas` (trazabilidad) |
| `appendix_a/spec_index_a.json` | Índice de especificaciones (190 entradas) | 190 | Validación de cobertura |

> **Corrección aplicada a `knowledge/claude.md`** ✅ (resources prevalece, regla §5): la pernería del B31.3 es **Tabla A-4** (no A-2 = `Ec`, A-3 = `Ej`); el **Apéndice B** son esfuerzos de diseño hidrostático de **tuberías no metálicas** (termoplásticos/RTR), no "factores de fatiga en soldadura". El knowledge ya fue actualizado con el detalle de A/B/C.
>
> **Alcance ampliado (confirmado):** se incluyen los **Apéndices B y C completos** —metales y no metales— para dar cobertura total al B31.3. Las bases no metálicas (`DB_B31_B`, C-2/C-4) quedan disponibles para líneas plásticas/RTR aunque el foco actual sea metálico.
>
> **Unidades duales SI/US, nativas en ambos códigos (confirmado):**
> - **B31.3** — SI + companion "C": `table_a_1c`, `table_a_4c`, `table_b_1c`, `table_c_1c`, `table_c_3c` (+ notas `a_1_a_1c` / `a_4_a_4c`).
> - **BPVC II-D** — dos carpetas nativas: `resources/bpvc_ii_d_metric_2025/` (°C, MPa) y **`resources/bpvc_ii_d_customary_2025/`** (°F, ksi), añadida por el usuario. Misma estructura (1808 filas en 1A, campos con sufijo `_ksi`/`_in`). **El gap de US en la II-D queda cerrado — ya no se usa conversión.**
>
> Los valores se cargan tal como están impresos en cada edición (no por conversión). Material con especificación dual (`SA-516/SA-516M`) aplica indistintamente en ambas ediciones.

**ASME BPVC II-D 2025 — dual: Métrico (`resources/bpvc_ii_d_metric_2025/`) y U.S. Customary (`resources/bpvc_ii_d_customary_2025/`, misma estructura, °F/ksi):**

| Archivo | Contenido | Filas | Uso en la DB |
|---|---|---|---|
| `table_1a.json` | Máx. esf. admisible `S`, **ferrosos** (Secc. I, III-2/3, VIII-1, XII) | 1.808 | **Núcleo `DB_BPVC_IID`** |
| `table_1b.json` | Máx. esf. admisible `S`, **no ferrosos** | 1.419 | Núcleo `DB_BPVC_IID` |
| `table_3.json` | Esf. admisible de pernería | 243 | Pernería `DB_BPVC_IID` |
| `notes_table_1a.json` / `notes_table_1b.json` | Notas (84 / 91 ítems) | — | Columna `notas` |

### 2.2 Alcance por fases (ambas fases confirmadas en este ciclo)

- **Fase 1 — Esfuerzos admisibles `S(T)`:** B31.3 A-1/A-1C y II-D 1A/1B, más pernería A-2 / Tabla 3. Es lo que consumen los motores hoy (`Sa_c`, `Sa_b`, `t_req`).
- **Fase 2 — Propiedades de material (confirmada dentro del ciclo):** de la misma fuente `resources/`:
  - **Resistencia a la tracción `Su`** — Tabla U (2.484 filas; usada por PCC-2 Art. 202/205 en `P = 2·t·S_act/D`).
  - **Límite de fluencia `Sy`** — Tabla Y-1 (2.475 filas).
  - **Módulo de elasticidad `E`** — II-D TM-1…5 (VIII-1) · B31.3 Apéndice **C-3** (tubería).
  - **Dilatación térmica** — II-D TE-1…5 (VIII-1) · B31.3 Apéndice **C-1** (tubería).
  - **Poisson y densidad** — Tabla PRD (115 filas).
  - **Factores de calidad B31.3** `Ej`/`Ec` — Apéndice A-3 / A-2 (§4.5; en realidad Fase 1, feeding `t_req`).

  Dos patrones de *lookup* distintos (ver §4.4): `Su` y `Sy` se indexan **por material** (se unen a la DB de esfuerzos por `material_id`); `E`, dilatación y Poisson/densidad se tabulan **por grupo de material**, por lo que requieren una **tabla de mapeo material→grupo**.

  **Consistencia de código en propiedades:** los cálculos de **tubería B31.3** toman `E` y dilatación del **Apéndice C del B31.3** (C-1 dilatación, C-3 módulo, metales); los de **recipiente VIII-1** los toman de la **II-D** (TE/TM/PRD). El motor conmuta la fuente igual que conmuta `S` (por `$D$11`). `Su`/`Sy` para B31.3, cuando se requieran, se toman de la II-D (Tabla U / Y-1), que el B31.3 referencia como base de resistencia.

> Nota de alcance: "todos los materiales" se interpreta como **todas las filas** de las tablas de esfuerzo admisible listadas (no un subconjunto curado). No se cargan aquí las tablas de otros ámbitos (2A/2B Div. 2 Clase 1, 5A/5B Div. 2 Clase 2, 6A–6D Secc. IV) salvo que se confirme su necesidad, para no mezclar bases de admisibles de secciones que el motor PCC no usa. Quedan como ampliación trivial si se requieren.

---

## 3. Decisiones de arquitectura

### 3.1 Ubicación
Hojas nuevas en el mismo libro (§0). Nombres: `DB_B31_3`, `DB_BPVC_IID`. Hoja auxiliar `DB_Listas` para rangos de validación (spec, grade, forma) y parámetros del *lookup*.

### 3.2 Clave de material (resuelve la ambigüedad *spec+grade* repetido)
Cada fila de las tablas reales puede repetir *spec/grade* variando **forma de producto, tamaño, clase/temple y P-No.** Por ello:

- Se agrega a cada base una columna **`material_id`** = clave única legible, p. ej.:
  `A106 | Gr.B | Smls. pipe | P-No.1` · `A516 | Gr.70 | Plate | ≤... | P-No.1`.
- El usuario **no** teclea material libre. Se implementa **validación en cascada** en las hojas de motor: `Spec. No.` → `Type/Grade` → `Forma/Tamaño/Clase` (listas dependientes con `INDIRECT`/`OFFSET` sobre `DB_Listas`), que resuelven a un único `material_id`.
- El *lookup* de `S(T)` se hace por `MATCH(material_id)` contra la columna clave de la base — determinista y auditable.

### 3.3 Modelo de temperatura (confirmado: interpolado con conmutador)
- Rejilla de temperatura = encabezados de columna de cada base (B31.3: 40/65/100…; II-D 1A: 40/65/100…900).
- Celda de control **`MODO_S`** en cada motor: `Interpolado` (por defecto) | `Tabulado-conservador`.
- **Interpolado:** lineal entre las dos temperaturas tabuladas que enmarcan `T`:
  `S = S1 + (S2 − S1)·(T − T1)/(T2 − T1)`, con `T1` = mayor tabulada ≤ T y `T2` = siguiente.
- **Tabulado-conservador:** toma `S` en `T2` (temperatura tabulada inmediatamente superior).
- **Bordes:** si `T` < mínima tabulada → usa el valor mínimo tabulado; si `T` > `Máx. Temp.` del material o la celda tabulada es `null` (no permitido) → el motor devuelve `#N/D controlado` con dictamen "Material fuera de rango a T" y bloquea el resultado (no extrapola). Regla de código: nunca extrapolar por encima de la temperatura máxima admisible.

---

### 3.4 Sistema de unidades (dual SI/US, confirmado)
- Cada base almacena las **dos versiones** de la tabla (SI y su companion "C" en US) como bandas o tablas emparejadas, con una columna `sistema` = `SI` | `US`.
- Cada buscador (§6) tiene un **conmutador SI ↔ US** que decide qué banda se lee y con qué encabezados de temperatura (°C vs. °F) y unidades (MPa vs. ksi; mm/mm/°C vs. in/in/°F; etc.).
- **Cálculos:** el motor sigue operando en SI por defecto (regla §1 del knowledge); el toggle US es para **consulta y extracción**. Si el usuario trabaja en US, se convierte a SI antes de calcular. No se mezclan sistemas en un mismo cálculo.
- **Fidelidad de impresión:** los valores se guardan tal como están impresos en cada tabla; SI y US son extracciones independientes (no conversiones) tanto en B31.3 (tablas "C") como en II-D (carpetas métrica y customary).
- **Factores de escala en propiedades:** al cargar C-3/C-3C el valor real de `E` se obtiene multiplicando el tabulado por 10³ (SI → MPa) o 10⁶ (US → psi); el script aplica y documenta el factor. La dilatación C-1/C-1C se modela con sus **dos filas por grupo**: fila **A** (coef. medio, 10⁻⁶/°C) y fila **B** (expansión lineal acumulada, mm/m), como en el código.

## 4. Esquema de las bases de datos

Cada hoja-DB tiene una **banda de identificación** (izquierda, columnas fijas) y una **banda de valores** `S` por temperatura (derecha, una columna por °C). Solo valores; sin fórmulas.

### 4.1 `DB_B31_3` (desde `table_a_1.json` + `table_a_1c.json`)

| Col. | Campo | Origen JSON |
|---|---|---|
| A | `material_id` (clave única) | derivado (spec+grade+forma+P-No.) |
| B | Composición nominal | `nominal_composition` |
| C | Forma de producto | `product_form` |
| D | Spec. No. | `spec_no` |
| E | Tipo/Grado | `type_grade` |
| F | UNS No. | `uns_no` |
| G | Clase/Condición/Temple | `class_condition_temper` |
| H | P-No. | (cols. de identificación) |
| I | Temp. mín., °C | `min_temp_c` |
| J | Resist. tracción mín., MPa | `min_tensile_strgth_mpa` |
| K | Fluencia mín., MPa | `min_yield_strgth_mpa` |
| L | Temp. máx., °C | `max_temp_c` |
| M | Notas | `notes` → cruzar con `notes_tables_a_1_a_1c.json` |
| N | `S` a ≤40 °C | `min_temp_to_40` |
| O…AV | `S` a 65,100,125,…,°C | `values{}` (34 columnas) |

### 4.2 `DB_BPVC_IID` (desde `table_1a.json` + `table_1b.json`)

| Col. | Campo | Origen JSON |
|---|---|---|
| A | `material_id` (clave única) | derivado |
| B | Composición nominal | `nominal_composition` |
| C | Forma de producto | `product_form` |
| D | Spec. No. | `spec_no` |
| E | Tipo/Grado | `type_grade` |
| F | UNS / Alloy desig. | `alloy_desig_uns_no` |
| G | Clase/Condición/Temple | `class_condition_temper` |
| H | Tamaño/Espesor, mm | `size_thickness_mm` |
| I | P-No. / Group No. | `p_no`, `group_no` |
| J | Resist. tracción mín., MPa | `min_tensile_strength_mpa` |
| K | Fluencia mín., MPa | `min_yield_strength_mpa` |
| L | Aplicabilidad I / III / VIII-1 / XII | `i`, `limits_iii`, `viii_1`, `xii` (`NP`=no permitido) |
| M | Gráfico presión externa | `external_pressure_chart_no` |
| N | Notas | `notes` → cruzar con `notes_table_1a/1b.json` |
| O…AT | `S` a 40,65,100,…,900 °C | `values{}` (32 columnas) |

> Los valores `null` (elipsis ASME = no admisible a esa temperatura) se conservan como celda vacía; el motor los trata como límite superior de temperatura (§3.3, bordes).

### 4.3 Pernería (bandas separadas dentro de cada DB o sub-hoja)
`DB_B31_3` ← `table_a_2.json`; `DB_BPVC_IID` ← `table_3.json`. Mismo patrón id + `S(T)`. Un selector `tipo=pernería` en el futuro módulo de bridas/PCC-1 conmutará a esta banda.

### 4.4 Bases de propiedades (Fase 2)

**Patrón A — indexado por material (`material_id`), interpolable en T.** Se cargan como bandas adicionales o hojas gemelas que comparten la clave `material_id` con `DB_BPVC_IID`:

| Hoja/banda | Fuente | Filas | Campos de valor |
|---|---|---|---|
| `DB_Su` (tracción) | `table_u.json` | 2.484 | `Su(T)` a 40,100,150,…,°C (30 cols) |
| `DB_Sy` (fluencia) | `table_y_1.json` | 2.475 | `Sy(T)` a 40,65,100,…,°C (35 cols) |

Identificación: `nominal_composition`, `product_form`, `spec_no`, `type_grade`, `alloy_desig_uns_no`, `class_condition_temper`, `size_thickness_mm` (idéntica a §4.2 → une por `material_id`). Interpolación y bordes según §3.3.

**Patrón B — indexado por grupo de material (no por spec/grado).** Requiere la tabla de mapeo **`MAP_Grupo`** (material → grupo TM / grupo TE / grupo PRD):

| Hoja | Fuente | Estructura | Clave de *lookup* |
|---|---|---|---|
| `DB_E` (módulo) | `table_tm_1…5.json` | 19+18+47+36+15 filas por **grupo de material** × T (−200…°C) | grupo de material + T |
| `DB_TE` (dilatación) | `table_te_1…5.json` | Indexada por **temperatura** × columnas de grupo (coef. A=instantáneo, B=medio, C) | grupo + T + tipo de coef. |
| `DB_PRD` (Poisson/densidad) | `table_prd.json` | 115 filas por `material`/`material_group`, valor único (sin T) | grupo de material |

> `MAP_Grupo` se construye una vez asociando cada `material_id` a su grupo TM/TE/PRD (p. ej. "aceros al carbono C ≤ 0,30 %", "aceros inoxidables austeníticos"). Es el único artefacto nuevo no derivable 1:1 del JSON; se validará material por material contra las cabeceras de grupo de TM/TE (§10).

**Propiedades lado B31.3 (Apéndice C).** Hojas gemelas indexadas por material/UNS, por temperatura:

| Hoja | Fuente | Filas | Contenido |
|---|---|---|---|
| `DB_C_dilatacion` | `table_c_1.json` | 52 | Coef. dilatación térmica de metales, por T |
| `DB_C_modulo` | `table_c_3.json` | 75 | Módulo `E` de metales, por T |

Se consultan cuando el código activo es B31.3; para VIII-1 se usan `DB_E`/`DB_TE` (II-D). Interpolación y bordes según §3.3.

### 4.5 Factores de calidad B31.3 — `MAP_Factores` (Fase 1)

Pequeñas tablas de consulta que hoy están hardcodeadas en el motor:

| Hoja/banda | Fuente | Filas | Alimenta |
|---|---|---|---|
| `Ec` (fundición) | `table_a_2.json` | 24 | Factor de calidad de fundición por `spec_no` |
| `Ej` (soldadura long.) | `table_a_3.json` | 127 | Factor de calidad de junta por `spec_no`+tipo/clase |

Reemplazan el valor fijo `D44 (E_j)=1` ✅ (confirmado): el motor leerá `Ej` (y `Ec` si aplica fundición) por la especificación del tubo, mejorando la trazabilidad de `t_req`. Requiere una lista desplegable de "tipo de junta/tubo" ligada a la especificación seleccionada.

### 4.6 Bases no metálicas B31.3 (Apéndices B y C-2/C-4)

| Hoja | Fuente | Contenido |
|---|---|---|
| `DB_B31_B` | `table_b_1…b_6.json` | Esf. de diseño hidrostático (HDS) y presión admisible de tuberías termoplásticas/RTR, por temperatura (23/38/82/93 °C, etc.) |
| `DB_C_dilat_nm` | `table_c_2.json` | Dilatación térmica de no metales |
| `DB_C_modulo_nm` | `table_c_4.json` | Módulo de elasticidad de no metales |

Se cargan como bases independientes con su propio buscador (§6). No alimentan por defecto el motor de parche metálico; quedan disponibles para futuros módulos de tubería no metálica y para consulta directa.

---

## 5. Motor de *lookup* (fórmulas objetivo)

Reemplazo de `D39`/`D40` en `Parche_PCC2_Art212` (y patrón replicable a futuros motores). Se define un **bloque de resolución de material** por cada material a evaluar (collar y metal base):

1. `fila = MATCH(material_id, <DB>!$A:$A, 0)` — DB según `$D$11` (1 → `DB_B31_3`, resto → `DB_BPVC_IID`).
2. `col_T1 = MATCH(T, <fila de encabezados de temperatura>, 1)` → temperatura tabulada ≤ T; `col_T2 = col_T1+1`.
3. `T1,T2,S1,S2 = INDEX(...)` sobre esa fila.
4. Control de borde: `IF(T > Temp.máx, NA_controlado, ...)`; `IF(S1 o S2 = "", cap a última tabulada no vacía)`.
5. Salida según `MODO_S`:
   `S = IF(MODO_S="Tabulado-conservador", S2, S1 + (S2-S1)*(T-T1)/(T2-T1))`

`Sa_c` y `Sa_b` invocan el bloque con `material_id` del collar y del metal base; `Sa = MIN(Sa_c, Sa_b)` (sin cambios aguas abajo). Toda la cadena `t_req`, `S_w`, verificaciones queda intacta.

> Implementación: para mantener las fórmulas legibles y auditables se usarán celdas intermedias nombradas (T1, T2, S1, S2, col_T) en una zona "resolución de material" del motor, en lugar de una megafórmula anidada.

---

## 6. Consulta interactiva (buscador dinámico por base de datos)

**Objetivo (confirmado):** que cualquiera de las bases pueda **buscarse y leerse en segundos**, extrayendo cualquier dato tal como está organizado en el código, sin recorrer miles de filas a mano. Mecanismo elegido: **fórmulas dinámicas de Excel 365/2021** (`FILTER`, `XLOOKUP`, `SORT`, `UNIQUE`, arrays derramados) — sin macros, el libro sigue en `.xlsx`.

### 6.1 Patrón de hoja de consulta
Cada base tiene una **hoja-consulta** asociada (`Buscar_B31_3`, `Buscar_BPVC_IID`, `Buscar_B31_B`, y análogas para propiedades), con esta estructura:

- **Celda de búsqueda** (texto libre): el usuario teclea parte de una especificación, grado, UNS o composición (p. ej. `A106`, `304`, `SA-516`, `F11501`).
- **Conmutador de unidades SI ↔ US** (§3.4): cambia la banda leída, los encabezados (°C/°F, MPa/ksi) y las unidades de la ficha. **Nativo en ambos códigos:** B31.3 usa sus tablas "C"; II-D usa la carpeta `bpvc_ii_d_customary_2025`.
- **Filtros de columna** (listas desplegables): Código/Tabla, forma de producto, P-No./grupo, aplicabilidad (I/III/VIII-1/XII), rango de temperatura.
- **Resultados derramados:** `FILTER(<DB>, (búsqueda en varias columnas) * (filtros))` devuelve todas las filas coincidentes; se ordenan con `SORT`. Si no hay coincidencias, mensaje "sin resultados".
- **Ficha del material seleccionado:** al elegir un `material_id` de los resultados, un bloque con `XLOOKUP` muestra automáticamente la fila completa —identificación + curva `S(T)` completa + `Su`, `Sy`, `E`, dilatación, notas y trazabilidad de tabla/edición—. Incluye el valor **interpolado a una temperatura que el usuario teclee** (reusa el motor de §5).

### 6.2 "Extraer cualquier dato según el código"
La ficha respeta la organización original de las tablas ASME: banda de identificación a la izquierda (como en el código) y banda de valores por temperatura a la derecha, con encabezados idénticos a los del estándar (40, 65, 100… °C, o °F en modo US). Botonera de copia rápida (celdas preparadas para copiar/pegar a una memoria de cálculo). Todo se actualiza al instante al cambiar la búsqueda o el sistema de unidades, sin recalcular el resto del libro (funciones no volátiles).

**Datos derivados en la ficha:** para B31.3, el buscador muestra también el **esfuerzo de diseño combinado `S·E`** (con `Ej`/`Ec` de A-3/A-2, §4.5). Se marca cuando un valor de A-1C está en **cursiva** (>2/3 del límite elástico) o **negrita** (90% del límite elástico) si esa tipografía viene capturada en la fuente; si el JSON no la trae, se indica como limitación y se añade una columna-bandera en una iteración posterior.

### 6.3 Requisito de versión
`FILTER`/`XLOOKUP`/`SORT`/`UNIQUE` requieren **Excel 365 o 2021**. Si el archivo se abriera en una versión anterior, las celdas de consulta mostrarían `#¡NOMBRE?`; las bases de datos y el motor de cálculo seguirían funcionando (no dependen de arrays dinámicos). Se documenta esta condición en Instrucciones (§7).

## 7. Guía de uso en la pestaña «Instrucciones»

La pestaña `Instrucciones` se amplía para explicar **en detalle cada base de datos y su buscador**. Contenido:

1. **Mapa de bases:** qué contiene cada hoja (`DB_B31_3`, `DB_BPVC_IID`, `DB_B31_B`, propiedades, factores), a qué código/tabla corresponde y su factor de seguridad.
2. **Cómo buscar (paso a paso, con capturas):** dónde teclear, cómo usar los filtros, cómo leer los resultados derramados y la ficha del material.
3. **Conmutador de unidades SI ↔ US:** cómo cambiar de sistema, qué tablas nativas usa el B31.3 ("C") y por qué la II-D en US es conversión etiquetada; advertencia de no mezclar sistemas en un cálculo.
4. **Cómo se elige la fuente por código:** B31.3 → Apéndice A/C; VIII-1 → II-D 1A/1B + TE/TM/PRD; cómo el selector `$D$11` conmuta todo.
5. **Temperatura:** explicación del modo interpolado vs. tabulado-conservador (`MODO_S`) y del bloqueo fuera de rango.
6. **Factores `Ej`/`Ec` y `S·E`:** cómo el motor los toma de A-3/A-2, qué elegir en la lista de tipo de junta y cómo se usa `S·E` en el diseño por presión.
7. **Trazabilidad:** cada dato cita su tabla/edición de `resources/`; cómo auditarlo.
8. **Limitaciones y requisitos:** Excel 365/2021 para el buscador; II-D-US por conversión; tipografía cursiva/negrita de A-1C; qué hacer en versiones previas.
9. **Convenciones del proyecto:** SI por defecto en cálculo, celdas editables (azules), protección de hoja (contraseña provisional `0000`).

Formato: tabla de contenidos navegable, lenguaje claro y ejemplos concretos (p. ej. "buscar A106 Gr.B a 250 °C en VIII-1").

## 8. Impacto en el libro existente

- `Datos_Ref` bloque B (filas 40–47): se **deprecia** y se conserva como respaldo histórico (marcar "OBSOLETO — ver DB_B31_3/DB_BPVC_IID"), sin borrar, para no romper trazabilidad de memorias ya emitidas.
- Bloque A de `Datos_Ref` (dimensiones B36.10M): **sin cambios**.
- `Parche_PCC2_Art212`: se añaden celdas de `material_id` + `MODO_S` + zona de resolución; se reconectan `D39`/`D40`. Se re-aplica protección de hoja (solo celdas de entrada azules editables; contraseña provisional `0000`, según convención del proyecto).
- Listas desplegables de `D22`/`D23` (material) pasan de la lista corta de `Datos_Ref` a la cascada sobre `DB_Listas`.

---

## 9. Plan de ejecución por fases

**Fase 0 — Preparación (0,5 día)**
Copia de respaldo del libro; congelar `resources/` como fuente; confirmar edición vigente (B31.3-2024, II-D 2025) contra los `*_index.json`.

**Fase 1 — Construcción de las bases (script Python, 1–2 días)**
Script `build_db_materiales.py` (openpyxl + json) que:
1. Lee los JSON de `resources/` (nunca memoria del modelo — regla §5.1).
2. Normaliza cada fila → `material_id` único (con desambiguación por forma/tamaño/clase/P-No.).
3. Escribe `DB_B31_3` y `DB_BPVC_IID` (solo valores), banda id + banda `S(T)`.
4. Genera `DB_Listas` (spec únicos, grados por spec, formas por spec+grade) para la cascada.
5. Registra en una columna/hoja `_meta` la trazabilidad (archivo fuente, tabla, edición) por regla §4.

**Fase 2 — Motor de *lookup* e integración de esfuerzos (1 día)**
Insertar zona de resolución + `MODO_S`; reconectar `D39/D40`; cascada de validación; reproteger hoja.

**Fase 3 — Bases de propiedades (confirmada en el ciclo, 1,5–2 días)**
3a. Cargar `DB_Su` y `DB_Sy` (Patrón A, unión por `material_id`) — mismo script, mismas rutinas de interpolación.
3b. Cargar `DB_E`, `DB_TE`, `DB_PRD` (Patrón B, II-D) y construir `MAP_Grupo` (material→grupo).
3c. Cargar `DB_C_dilatacion` y `DB_C_modulo` (Apéndice C metales, lado B31.3) y las bases no metálicas `DB_B31_B`, `DB_C_dilat_nm`, `DB_C_modulo_nm` (Apéndices B y C-2/C-4).
3d. Exponer las propiedades en los motores donde aplican, conmutando la fuente por código (B31.3 → Apéndice C; VIII-1 → II-D): `S_act`=`Su(T)` en Art. 202/205; `E` y dilatación para análisis futuros.

> `MAP_Factores` (`Ej`/`Ec`, §4.5) se construye en la **Fase 2 del motor** (junto con los esfuerzos) por ser Fase 1 conceptual: alimenta `t_req` directamente.

**Fase 4 — Consulta interactiva + Instrucciones (1,5 día)**
Construir las hojas-consulta con buscador dinámico (§6) para cada base y reescribir la pestaña `Instrucciones` con la guía detallada (§7).

**Fase 5 — Verificación y QA (1 día)** — ver §10 (incluye validación del mapeo de grupos y pruebas del buscador).

**Fase 6 — Entrega y documentación**
Nota de versión; commit de la memoria de cambios.

**Fase 7 — (Opcional) Migración a DB externa** si el peso del libro lo justifica.

---

## 10. Verificación y trazabilidad (protocolo de aceptación)

1. **Conteo:** filas cargadas por base = filas del JSON fuente (1.146 A-1; 1.808 1A; 1.419 1B; etc.). Sin pérdidas ni duplicados de `material_id`.
2. **Puntos de control (spot-check) contra el código:** verificar ≥10 materiales/temperaturas conocidos contra el PDF/JSON — p. ej. A106 Gr.B a 40, 200, 400 °C en ambas bases; A516 Gr.70; un inoxidable austenítico (304/316) a alta T; un no ferroso de 1B.
3. **Regresión del caso semilla:** el collar 12"-CWS-46-032-B1 debe reproducir el resultado actual cuando `T=25 °C` y `MODO_S` en el modo equivalente, dentro de tolerancia por la nueva fuente de `S`.
4. **Interpolación:** validar `S(T)` interpolado vs. cálculo manual en 3 puntos intermedios; validar que `MODO_S=Tabulado` devuelve exactamente el valor de la temperatura superior.
5. **Bordes:** `T` > `Temp.máx` → dictamen "fuera de rango" y bloqueo; `T` bajo mínimo → valor mínimo tabulado.
6. **Pruebas unitarias** del script (pytest): normalización de claves, manejo de `null`, unicidad de `material_id`.
7. **Auditoría de trazabilidad:** cada base apunta a su archivo `resources/` (regla §4).
8. **Buscador (§6):** probar que una búsqueda parcial (`A106`, `304`, `SA-516`) devuelve todas las filas esperadas; que los filtros combinan correctamente; que la ficha por `XLOOKUP` trae la fila completa e interpola a la T tecleada; y comportamiento "sin resultados".
9. **Unidades SI/US (§3.4):** conmutar y verificar que B31.3 lee la tabla nativa "C" y la II-D la carpeta customary, con encabezados y unidades correctos (°F, ksi); que el factor ×10³/×10⁶ de C-3/C-3C se aplica bien; spot-check de un mismo material en ambos sistemas (p. ej. SA-516 Gr.70 en MPa/°C vs ksi/°F), confirmando aplicabilidad NP/SPT y coherencia de la curva.

---

## 11. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| Colisión de `material_id` (spec+grade repetido) | Clave compuesta con forma/tamaño/clase/P-No.; test de unicidad en QA (§10.6) |
| Peso del libro / recálculo lento | Hojas-DB solo valores; `INDEX/MATCH` no volátil; fallback externo (Fase 7) |
| Buscador `#¡NOMBRE?` en Excel < 2021 | El buscador usa arrays dinámicos (Excel 365/2021); bases y motor no dependen de ellos; se documenta en Instrucciones (§7) |
| Extrapolación fuera de rango de T | Bloqueo por `Temp.máx` y `null` (§3.3) |
| Notas de código con restricciones (soldadura, `NP`, servicios) | Columna `notas` + aplicabilidad I/III/VIII-1/XII visible; el usuario valida la nota antes de emitir |
| Cambio de edición del código | Hoja `_meta` con edición/fecha; re-ejecutar script sobre nuevos JSON |
| Ruptura de memorias previas | `Datos_Ref` B se conserva como OBSOLETO, no se borra |

---

## 12. Entregables

1. `Motor_de_Calculo_ASME_PCC.xlsx` actualizado — hojas de datos: `DB_B31_3` (A-1/A-1C + pernería A-4), `DB_BPVC_IID` (1A/1B + pernería T.3), `MAP_Factores` (`Ej`/`Ec`), `DB_Listas`, `DB_Su`, `DB_Sy`, `DB_E`, `DB_TE`, `DB_PRD`, `MAP_Grupo`, `DB_C_dilatacion`, `DB_C_modulo`, y bases no metálicas `DB_B31_B`, `DB_C_dilat_nm`, `DB_C_modulo_nm`; hojas-consulta con buscador dinámico (`Buscar_*`); pestaña `Instrucciones` reescrita; motor reconectado.
2. `build_db_materiales.py` + pruebas unitarias (modular, documentado, SI) — guardado en `outputs/`.
3. Reporte de verificación (conteos + spot-checks + regresión del caso semilla + validación de mapeo de grupos + pruebas del buscador).
4. Nota de versión en `Instrucciones`.
5. `knowledge/claude.md` corregido y actualizado ✅ (ya entregado): pernería A-4, Apéndice B no metálicos, y regla de unidades dual SI/US.

---

### Estado de decisiones (todas confirmadas)
- **Ubicación:** ✅ hojas en el mismo libro.
- **Alcance II-D:** ✅ Tablas 1A/1B (+ pernería T.3) + propiedades U/Y-1/TE/TM/PRD. Fuera: 2A/2B, 4, 5A/5B, 6A–6D.
- **Alcance B31.3:** ✅ Apéndice A (A-1/A-1C, pernería A-4, factores `Ec`/`Ej` A-2/A-3) **+ Apéndices B y C completos** (metales y no metales).
- **Factores `Ej`/`Ec`:** ✅ reemplazan el `E_j=1` fijo por *lookup* (§4.5).
- **Fase 2 (propiedades):** ✅ incluida en este ciclo.
- **Consulta interactiva:** ✅ buscador con fórmulas dinámicas Excel 365/2021 (§6); guía detallada por base en `Instrucciones` (§7).
- **Unidades duales SI/US:** ✅ cada buscador conmuta SI ↔ US (§3.4), **nativo en ambos códigos**: B31.3 tablas "C" y II-D carpeta `bpvc_ii_d_customary_2025` (añadida por el usuario). Sin conversiones.
- **`knowledge/claude.md`:** ✅ corregido (A-4/Apéndice B), regla de unidades dual, II-D US integrada y guía de lectura de la Subparte 1 (NP/SPT, creep, interpolación, no extrapolar) — entregado.

**Alcance totalmente cerrado, sin gaps abiertos. Listo para ejecutar la Fase 1 a tu orden.**

*Herramienta de ingeniería de referencia. Verificar entradas y resultados antes de emitir para construcción.*
