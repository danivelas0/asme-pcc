# Rediseño del motor Art. 212 para igualar el flujo de proceso — Plan de implementación

> **Para ejecutores:** SUB-SKILL REQUERIDA: usar `superpowers:subagent-driven-development`
> (recomendado) o `superpowers:executing-plans` para ejecutar tarea por tarea. Los pasos
> usan casillas (`- [ ]`) para el seguimiento.

**Objetivo:** Que la hoja `Parche_PCC2_Art212` (motor de cálculo del Art. 212) reproduzca
**los 8 pasos** del flujo de proceso aprobado por el ingeniero
(`outputs/plans/flows/flujo_proceso_art_212.pdf` y su transcripción `.md`), cerrando los
huecos detectados en la verificación: compuerta de elegibilidad, cargas externas, topes de
soldadura, término de separación en la excentricidad, curvatura doble y energía neumática.

**Arquitectura:** `build_parche_art212()` (hoy 100 % en código, hermana de
`build_collar_art206`) se **reestructura por pasos 1-8 explícitos**. Toda fórmula o
coeficiente nuevo se **traza a `resources/`** (Regla nº 1). Dos coeficientes no estaban como
texto: la **ec. (2)** del Art. 212 vive en `resources/` **como imagen** (bloque `figure` con
`image:`) y se recupera leyendo ese PNG del propio `resources/`; la **ec. (II-1)** del App.
501-II es un **vacío real de `resources/`** (bloque de ecuación vacío, sin imagen, en los dos
espejos) que se repara re-extrayendo ese apéndice. La Fase 0 los cierra **sin folio externo**
(`resources/` es la fuente; regla registrada en `CLAUDE.md`). El *oracle*
`parche_art212_ref.json` deja de ser una copia de la
hoja heredada y pasa a **re-baselinarse desde la hoja nueva** una vez el ingeniero la valida
en Excel; durante el rediseño, cada tarea fija sus fórmulas nuevas con aserciones de cadena
contra lo que este plan especifica (trazado a `resources/`), y `verificar.py §6/§7` +
el F9 del ingeniero son el árbitro final en Excel real.

**Stack:** Python 3 + openpyxl; pytest; el builder `build_db_materiales.py`; Excel de
Windows para la validación final (`verificar.py §6+`, F9 del caso semilla). openpyxl
`CellRichText` + `InlineFont(vertAlign="subscript")` para los símbolos con subíndice.

**Spec:** `outputs/plans/flows/flujo_proceso_art_212.md` (flujo aprobado) + la verificación
de esta conversación. Fuente normativa única: `resources/ASME PCC/pcc_2/.../
article_212_fillet_welded_patches.json` (Art. 212, ed. 2022) y `.../app_501_ii`, `.../
app_501_iii` (App. 501). Decisiones de alcance del ingeniero (2026-09-10):
1. Cargas externas: **entradas completas** F_CO/F_LO → F_C, F_L, F_max = max(F_C, F_L).
2. Pasos 7-8: implementar **recálculo de e con la separación g** y **energía neumática** (App. 501).
3. Ec. (5) de excentricidad: **literal del código (sólo cilindro)**; esfera/cabezal → hand-off declarado del 212-3.2(c) / análisis, sin escalar por kf.
4. **Refactor por pasos 1-8 explícitos.**
5. Notación: **sin guiones bajos en los símbolos visibles** — subíndices reales (texto enriquecido).

## Global Constraints

Copiadas de `asme-pcc/CLAUDE.md` («Reglas de diseño del libro» + Regla nº 1). Todas las
tareas las incluyen implícitamente:

- **Regla nº 1 — ningún valor normativo sale de memoria.** Cada fórmula/coeficiente nuevo
  cita el bloque de `resources/` del que procede. Los dos que hoy son imagen se extraen a
  JSON en la Fase 0 **antes** de que un motor los consuma. Prohibido codificar la energía
  neumática con la forma aire-only `2.5`/`0.286` de `knowledge/claude.md` (es derivada): el
  código publica la ec. (II-1) en función de `k` (calor específico del fluido de prueba).
- **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`,
  `VSTACK`, `_xlfn`. Sólo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`. `verificar.py §5` falla si aparece una.
- **Un texto no se guarda como fórmula** (`="…"` > 255 car. → `_xlfn._LONGTEXT`). Usar el
  normalizador existente; no introducir literales `="..."` largos nuevos.
- **Validación de datos: sólo rango literal o lista de ítems**, nunca fórmula como origen.
  `dv_list(ws, cell, '"a,b,c"', com)` para categóricos; listas dependientes materializadas
  en columnas ocultas para lo que salga de una base.
- **Regla 12/14 — todo material y todo valor tabulado se elige de una base (`DB_*`), no se
  teclea.** Las entradas categóricas nuevas del Paso 1 (mecanismo de daño, tipo de servicio)
  **no** están tabuladas en ninguna base: son lista de ítems literal (permitido; no son
  materiales). Las de proceso/campo (T, P, cargas externas, dimensiones del defecto,
  separación) se teclean (rule 14). El material sigue saliendo de la cascada de la Sección 7.
- **Regla 13 — los motores no se cruzan.** El motor 212 lee sólo de `DB_*`, nunca de un buscador.
- **snake_case** en todo archivo nuevo. **Español, SI por defecto** (MPa, mm, °C).
- **Entrega parchando sobre Rev4** por el builder, no una Rev5.
- **Sistema visual Swiss Industrial Print:** sólo tokens del sistema (`PAPEL`, `TINTA`, `ROJO`,
  `MONO`, `MACRO`, `CAJA_CAMPO`, `CAJA_TECLEO`, etc.). `TestSistemaVisual` recorre las celdas
  con formato y falla ante cualquier color/fuente fuera de paleta. La celda que se teclea es
  el único rectángulo rojo cerrado (`CAJA_TECLEO`); las de lista, línea inferior (`CAJA_CAMPO`).
- **Sincronía Python↔VBA:** si el refactor cambia el nº de filas navegables o `NAVEGABLES`,
  actualizar `mod_nav.vba` y `HojasNavegables()` a la vez (`TestSincroniaPythonVba`). Este
  refactor **no** añade hojas navegables nuevas (todo ocurre dentro de `Parche_PCC2_Art212`),
  así que no debería tocar la capa de navegación; confirmarlo.

---

## Estructura de archivos

- `scripts/completar_art_212_eq2.py` — **Create.** Extrae la ec. (2) `F_LP = P·D_m/4` del
  Art. 212 (hoy bloque `figure` nº 33, sin capa de texto) al JSON, con `extraction_amendments`
  y SHA-256 del folio. Idempotente; sólo escribe metadatos + el valor de la ecuación.
- `scripts/completar_app_501_energia.py` — **Create.** Extrae las ec. (II-1)/(II-2) de energía
  almacenada del App. 501-II (bloque imagen nº 3) y confirma la Tabla 501-III-1-1 de `R_scaled`
  al JSON, con `extraction_amendments` y SHA-256.
- `scripts/build_db_materiales.py` — **Modify.** Reestructura `build_parche_art212()`
  (6165-7034) por pasos 1-8; añade el helper de símbolos con subíndice `sym()`; nuevas
  secciones (elegibilidad, cargas externas, energía neumática). Lee de las bases y de los JSON
  de la Fase 0. `main()` (9085) no cambia de firma.
- `scripts/comentar_art212_base.py` lógica embebida (`comentar_art212_base`, 7036): **Modify.**
  Ampliar el dict de notas a las filas nuevas.
- `scripts/parche_art212_ref.json` — **Regenerate (al final).** Deja de ser copia de la hoja
  heredada; se re-baselina desde la hoja nueva tras la validación en Excel del ingeniero.
- `scripts/_dump_parche_ref.py` — **Reuse.** Regenera el *oracle* desde el `.xlsm` validado.
- `scripts/test_dashboard.py` — **Modify.** `TestBuildParcheContraOracle` pasa a fijar las
  fórmulas **nuevas** por paso (anclas por sección); `TestParidadHojaParche` se re-arma al
  final contra el *oracle* re-baselinado.
- `scripts/verificar.py` — **Modify.** §7 (regresión del caso semilla) recalcula el dictamen
  con las fórmulas nuevas; nuevas comprobaciones §6e (topes de soldadura, e con g, energía
  neumática) que recalculan en Excel la misma expresión que emite el motor (patrón §6d).

---

## Fase 0 — Cerrar los dos huecos de Regla nº 1 **dentro de `resources/`** (prerrequisito)

> **`resources/` es la fuente; no se pide ningún folio externo** (regla registrada en
> `CLAUDE.md` tras la Regla nº 1). Los dos coeficientes que hoy no están como texto se
> resuelven así: la ec. (2) del Art. 212 **sí vive en `resources/` como imagen** y se
> recupera leyéndola; la ec. (II-1) del App. 501 **no está en `resources/` ni como texto ni
> como imagen** (bloque de ecuación vacío, sin `image`, en los dos espejos) — es un **vacío
> real** que se repara re-extrayendo ese apéndice. Estado ya verificado en el diseño.

### Tarea 0.1 — Recuperar la ec. (2) `F_LP = P·D_m/4` **de la imagen en `resources/`** ✅ dato ya confirmado

**Files:** Create `scripts/completar_art_212_eq2.py`; Modify el JSON del Art. 212 (los dos
espejos: `p2_welded_repairs/.../art_212.json` y `part_2_welded_repairs/.../
article_212_fillet_welded_patches.json`).

**Fuente (resources/):** el bloque 33 es `type: figure` con
`image: figures/article_212_fillet_welded_patches_diagram_3_2.png`. Leída esa imagen con la
herramienta Read, imprime literalmente **`F_LP = P·D_m/4` (2)**. El dato ya está confirmado
desde `resources/`; esta tarea solo lo fija en la capa de texto del JSON para que sea
citable/trazable sin volver a abrir la imagen.

**Interfaces:** Produce en cada espejo un `text` para la ec. (2) = `"F_LP = P·Dm/4 (2)"` y un
`extraction_amendments` con la procedencia (bloque 33, la ruta del PNG de `resources/`, nota
«recuperada de la imagen del propio `resources/`, no de un folio externo»).

- [x] **Step 1: Confirmar el ancla.** `python -c` que imprime `blocks[33]` de ambos espejos →
  `type=figure`, `image=…diagram_3_2.png`, `text` vacío. (Verificado en el diseño y en ejecución.)
- [x] **Step 2: Escribir `completar_art_212_eq2.py`** — idempotente, `--resources`; **lee el
  PNG referenciado por el propio bloque** (no un PDF), escribe `text="F_LP = P·Dm/4 (2)"` con
  su `extraction_amendments`, no toca ningún otro valor. Aborta si el bloque 33 ya no apunta a
  ese PNG (defensa ante reorganización de `resources/`). Imagen leída y confirmada: `F_LP = P·Dm/4 (2)`.
- [x] **Step 3: Correr una vez por espejo** y comprobar con `git diff` que sólo cambia el
  bloque de la ec. (2) y su metadato. Confirmado: 28 ins / 4 del, solo el bloque 33 y el amendment.
- [x] **Step 4: Commit** — `Art. 212: fija la ec. (2) F_LP=P·Dm/4 desde la imagen de resources/ (Regla n.1)`.

### Tarea 0.2 — Reparar el **vacío real** de la energía almacenada (App. 501-II eq. II-1/II-2) en `resources/`

**Files:** Create `scripts/completar_app_501_energia.py`; Modify `app_501_ii.json` (y su
espejo `part_5_...`); leer/validar Tabla 501-III-1-1 en `app_501_iii.json`.

**Fuente (resources/) — PREMISA CORREGIDA EN EJECUCIÓN (2026-09-11).** El diseño del plan
afirmaba que los bloques de la ec. (II-1)/(II-2) estaban **vacíos y sin `image`** (vacío real).
Al ejecutar se comprobó que es **falso**: los bloques **sí traen texto**, pero la extracción
está **colapsada** — el PDF de PCC-2 es born-digital y su fuente matemática privada mapeó los
corchetes y operadores a letras Latin-1 (`Ä Å Ç É Ñ Ö ×`), no a `U+FFFD`. `figures = 0` (no
hay imagen que leer) y el otro espejo (`part_5_…`) **no contiene** el App. 501-II. Las hermanas
aire-only (II-2, II-4) tienen el mismo defecto; las de TNT (II-3, II-5) llegaron legibles.

> **Decisión del ingeniero (2026-09-11): «Léelo directamente de la carpeta de standards».**
> El PDF de PCC-2 sí está en `…/0-CODES/PCC - POST CONSTRUCTION CODE/ASME PCC-2 REPAIR OF
> PRESSURE EQUIPMENT AND PIPING.pdf` (la ruta del `CLAUDE.md` usa el usuario `dvelasquez`; en
> esta máquina es `User`). Se leyó **solo para corregir/verificar** el JSON (uso permitido por
> la Regla nº 1); nunca se añade al repo (copyright ASME) ni alimenta un motor directamente.
> Las páginas 306-308 (folios 281-282) se renderizaron con PyMuPDF y se leyeron; de ahí salen
> las formas limpias. Camino (a) cumplido **sin re-extraer con marker**: la corrección puntual
> de la capa de texto es suficiente y más segura que re-correr el extractor.

- [x] **Step 1: Confirmar el estado real** en `resources/`. Hecho: no es vacío sino **texto
  colapsado**; sin imagen; sin espejo `part_5`. Premisa del diseño corregida (arriba).
- [x] **Step 2 (camino a):** `completar_app_501_energia.py` fija en la capa de texto las ec.
  (II-1) `E = [1/(k-1)]*Pat*V*[1 - (Pa/Pat)^((k-1)/k)]`, (II-2) `2.5*…^0.286` y (II-4)
  `360*…^0.286` **tal como las imprime el código** (la general en `k`, no la aire-only como
  única forma), con procedencia (PDF, folios, SHA-256 `ab8e7b6…`) en `extraction_amendments`.
  Idempotente y defensivo (localiza por rótulo y tipo). Ningún valor se altera.
- [x] **Step 3: Leer la Tabla 501-III-1-1** (`R_scaled`) del PDF y dejar constancia estructurada
  en `app_501_iii.json` (4 filas: 20/50 vidrio, 12/30 tímpano-bloques, 6/15 pulmón-ladrillo,
  2/5 fatal) + reglas (`R_scaled ≥ 20 m/kg^⅓`; `R = 30 m` para `E ≤ 8 130 000 J`) y la ec.
  (III-1) `R = Rscaled*(2*TNT)^(1/3)`, para que la Fase 8 los ofrezca por lista con trazabilidad.
- [x] **Step 4: Commit** — `App. 501: repara la extraccion colapsada de la energia almacenada (II-1/II-2/II-4, III-1) desde el PDF (Regla n.1)`.

---

## Fase 1 — Paso 1: compuerta de elegibilidad y caracterización del daño (212-1/2)

**Files:** Modify `build_parche_art212()` — insertar una sección nueva **antes** de la
Sección 1 de datos (o como sub-bloque de ella), y su dictamen; ampliar `comentar_art212_base`.
Modify `test_dashboard.py`.

**Fuente (resources/, Art. 212):** bloques [6] (T > nil-ductility hasta **345 °C**; < 0 °C
tenacidad a la entalla; > 345 °C creep/fatiga), [4]/[11] (mecanismo de daño = adelgazamiento
local/erosión/corrosión/traspasante; **no usar** si el daño no se caracteriza), [11]-[13]
(grietas: sólo si crecimiento arrestado/predecible + FFS), y el límite de servicio letal /
uso del Art. 201 (flujo, cruzado con Part 1 del estándar — bloque [10] remite a Part 1).

**Interfaces:** Produce las celdas de entrada del Paso 1 y una celda de **dictamen de
elegibilidad** que el DICTAMEN GLOBAL (F90) consume: si no es elegible, el resultado global
se bloquea con el motivo, en rojo, antes de cualquier cálculo.

- [x] **Step 1: test (falla).** `test_paso1_elegibilidad` en `TestBuildParcheContraOracle`
  (construye con stubs, sin Excel): afirma banda A134, D136-D139 sembradas, D140 y F90 nuevos
  por cadena, y validación de lista en las cuatro entradas. Corrió → FAIL (A134 vacía). ✔
- [x] **Step 2: implementar las entradas del Paso 1.** Bloque nuevo en el **anexo de pasos del
  flujo (filas 134-140)**, direcciones estables (decisión del ingeniero 2026-09-11: bloques
  nuevos, no reordenar). D136 mecanismo, D137 caracterizable, D138 servicio, D139 grietas, con
  `dv_list`. Los 3 literales categóricos nuevos entran en `LITERALES_PERMITIDOS` (no son
  materiales; regla 12/14 los permite como el selector D10). `sym()` es de la Fase 9 (aún no).
- [x] **Step 3: dictamen de elegibilidad** en D140 (IF anidado, solo IF/OR): letal→PROHIBIDO
  (Art. 201, flujo); no caracterizable→NO ELEGIBLE (212-2c); grieta activa→NO ELEGIBLE (212-2c);
  T>345→FUERA DE ALCANCE (212-1e); T<0→REVISAR entalla (cribado nil-ductility, 212-1e); else
  ELEGIBLE. **Trazado a `resources/`:** el 345 °C es del bloque [6]; el 0 °C es cribado del
  nil-ductility (no umbral del código), anotado honestamente; el letal→Art. 201 es del flujo.
- [x] **Step 4: cablear a F90.** F90 antepone `LEFT($D$140,·)` para PROHIBIDO/NO ELEGIBLE/
  FUERA DE ALCANCE (bloquean); el aviso de entalla no bloquea. Declarada en
  `DIVERGENCIAS_REEMPLAZADAS`. Verificado en Excel real: caso semilla (ELEGIBLE) → **APTO**.
- [x] **Step 5: test PASS + commit.** pytest 237 passed; `verificar.py` **0 fallos** (§5b y §7
  incluidos). Commit `Art. 212 Paso 1: compuerta de elegibilidad (212-1/2)`.

---

## Fase 2 — Paso 2: cargas de presión y externas combinadas (212-3.2)

**Files:** Modify `build_parche_art212()` (Sección de cargas, hoy filas 59-69). Modify
`test_dashboard.py`, `verificar.py §6e`.

**Fuente (resources/):** ec. (1) `F_CP = P·D_m/2` (bloque [29]); ec. (2) `F_LP = P·D_m/4`
(Fase 0.1); `F_C = F_CP + F_CO`, `F_L = F_LP + F_LO` (bloques [37],[39]); `F_A > F_C y F_L`
(bloque [59]); 212-3.2(c) alternativas para esfera/toriesférico/elipsoidal (bloque [44]);
212-3.1(a) exige evaluar flexión/torsión/viento/fatiga (bloque [16]).

**Interfaces:** El motor deja de usar `F_m = kf·P·Dm` como fuerza única. Calcula `F_CP`,
`F_LP` (cilindro), suma `F_CO`/`F_LO` (entradas nuevas, default 0 no — **entradas completas**:
el ingeniero las teclea; 0 es un valor válido tecleado, no un default oculto) → `F_C`, `F_L`,
`F_max = MAX(F_C, F_L)`. `F_max` alimenta el filete (Paso 4). Para geometría no cilíndrica, el
motor **declara** el hand-off del 212-3.2(c) en vez de aplicar (1)/(2).

- [x] **Step 1: test (falla).** `test_paso2_cargas` afirma A142, D144/D145 (F_CO/F_LO=0),
  F_CP=`=D62*$D$52/2`, F_LP=`=D62*$D$52/4`, F_C=`=D146+$D$144`, F_L=`=D147+$D$145`,
  F_max=`=IF($D$11=3,NA(),MAX(D148,D149))`. → FAIL, luego PASS.
- [x] **Step 2: entradas externas.** F_CO (D144) y F_LO (D145) en el anexo (Paso 2), `inp`
  tecleables default 0, ref 212-3.1(a)/212-3.2(b). Van en el anexo, no en la Sección 1, para no
  desplazar el oracle (misma decisión de maquetación).
- [x] **Step 3: fuerzas de presión (cilindro, por columna Op/Diseño/Env).** F_CP (146) = P·Dm/2
  [29]; F_LP (147) = P·Dm/4 [33, Fase 0.1]. Filas triples D/E/F.
- [x] **Step 4: totales y gobernante.** F_C (148), F_L (149), F_max (150) por columna.
- [x] **Step 5: hand-off no cilíndrico.** F_max = `IF($D$11=3,NA(),MAX(...))`: solo esfera/
  cabezal (D11=3) es no-cilindro (virola D11=2 SÍ es cilindro — corregido respecto al «≠1» del
  plan). D151 declara el hand-off 212-3.2(c). Decisión 3: literal, solo cilindro.
- [x] **Step 6: `verificar.py §6e`** (nueva sección) recalcula F_CP/F_LP/F_C/F_L/F_max en Excel
  para el caso semilla desde su re-derivación en Python (P·Dm/2 etc.) sobre las entradas que la
  hoja recalculó. 15 casos OK. pytest 238 · verificar.py **0 fallos**. Commit.
  **Nota:** el cableado de w_min a F_max NO es de esta fase (es Fase 4 Step 2); Fase 2 solo
  CREA F_max, por eso no toca ninguna celda del oracle.

---

## Fase 3 — Paso 3: proximidad a discontinuidades (212-3.3)

**Files:** Modify `build_parche_art212()` (hoy L_min en D73, ubicación en D88). Modify tests.

**Fuente (resources/):** ec. (3) `L_min = 2·(R_m·t)^(1/2)` (bloque [47]); aplica a boquilla
**y** a parches adyacentes (bloques [50],[51]); < L_min → refuerzo 360° penetración completa
o análisis (bloque [53]); esquinas redondeadas (flujo: `R_min = 75 mm`; bloque [7] «rounded
corners», el 75 mm es del flujo — citar como recomendación del flujo, no del texto del 212).

**Interfaces:** Mantiene `L_min` y la rama 3A/3B (parche local vs refuerzo 360°). Añade la
distancia a parches adyacentes (3C) y la nota de esquinas redondeadas.

- [x] **Step 1: test (falla).** `test_paso3_discontinuidad`: A153, D156 dictamen 3C, nota de 75 mm.
- [x] **Step 2: 3C.** Anexo (Paso 3, filas 153-156): D155 «Distancia a parches adyacentes» [mm]
  (blanco = sin adyacente); D156 = `IF($D$155="","No aplica...",IF($D$155>=$D$73,"OK","< L_min
  — reubicar (212-3.3)"))`. Lee L_min de D73 (no lo modifica). Cita bloque [51].
- [x] **Step 3: nota de esquinas.** D157 (texto plano, no fórmula): R_min = 75 mm — **anotado
  como recomendación del flujo**, no del texto del 212 (que solo dice «rounded corners», bloque
  [7]-f). Regla n.1 respetada.
- [x] **Step 4: PASS + commit.** pytest 239 · verificar.py **0 fallos**.

---

## Fase 4 — Paso 4: soldadura perimetral de filete con topes (212-3.4A)

**Files:** Modify `build_parche_art212()` (w_min D64, verificación D84). Modify tests, `§6e`.

**Fuente (resources/):** ec. (4) `w_min = F_A/(E·S_a)`, `E = 0.55` (bloques [57],[59]); NOTA:
`w ≤ min(T, t)` y `w ≤ 40 mm (1.5")` (bloque [60]); bisel alternativo, garganta ≤ nominal
(bloque [61]).

**Interfaces:** `w_min` pasa a usar `F_max` (Fase 2) en vez de `F_m`. Se añaden las dos
verificaciones de tope que hoy faltan.

- [x] **Step 1: test (falla).** `test_paso4_filete`: D64/F64 = F_max, topes F161/F162. → PASS.
- [x] **Step 2: `w_min` desde `F_max`.** D64/E64/F64 = `=D150/($D$42*$D$41)` (F_max, no F_m). A/B/
  C/G64 siguen anclados; D/E/F64 en `DIVERGENCIAS_REEMPLAZADAS`. Seed sin cambio (cilindro).
- [x] **Step 3: topes.** Anexo Paso 4 (filas 159-162): D161=`MIN($D$29,$D$21)`, F161 dictamen
  «≤ menor espesor»; D162=40, F162 dictamen «≤ 40 mm». Bloque [60].
- [x] **Step 4: nota de bisel.** D163 (texto): garganta efectiva ≤ nominal (212-3.4b, [61]).
- [x] **Step 5: cablear a F90.** AND incluye F161, F162. §6e verifica w_min=F_max/(E·Sa) en
  Excel (3 casos OK). pytest 240 · verificar.py **0 fallos** · seed → APTO (topes CUMPLE).

---

## Fase 5 — Paso 5: excentricidad literal de cilindro, con separación g (212-3.4C / 212-4c)

**Files:** Modify `build_parche_art212()` (e en D55, S_w en D66-D68). Modify tests, `§6e`.

**Fuente (resources/):** ec. (5) `S_w = P·D_m/(2T) + 3·P·D_m·e/T²`, `S_w ≤ 1.5·S_a`,
`e = (T + t)/2` (bloques [65],[67]); 212-4(c): si la separación en el borde ≥ **1.5 mm**,
recalcular el filete **sumando la separación a la excentricidad** → `e = (T + t + g)/2`
(bloque [82]).

**Interfaces:** `e` pasa a incluir la separación `g` cuando `g ≥ 1.5 mm` (corrige el defecto
detectado). `S_w` se escribe **literal de la ec. (5)** (cilindro), sin el factor `kf`. En modo
no cilíndrico, `S_w` = NA() y la verificación declara «ec. (5) aplica sólo a cilindro (212-3.4);
esfera/cabezal → análisis» (decisión 3).

- [x] **Step 1: test (falla).** `test_paso5_excentricidad`: e con g, S_w literal, NA en esfera. PASS.
- [x] **Step 2: `e` con umbral de g.** **Resuelto el riesgo del plan: D31 (luz radial de
  conformado, alimenta Rf/desarrollo) NO es el gap de faying edge.** El flujo (línea 74)
  confirma que `g` es la separación de fit-up. Se añade una **entrada nueva `g` (D167)**,
  distinta de D31, default 0. `D55 = ($D$29+$D$21+IF($D$167>=1.5,$D$167,0))/2`. Bloque [82].
- [x] **Step 3: `S_w` literal de cilindro.** D66=`P·Dm/(2T)`, D67=`3·P·Dm·e/T²` directas (sin kf);
  D68=D66+D67 sin cambio de fórmula. C_sw (D57) también literal. Para cilindro coinciden con el
  valor anterior (kf=0.5 ya daba la ec.5); NA en esfera.
- [x] **Step 4: hand-off no cilíndrico.** D66/D67/D57 = `IF($D$11=3,NA(),...)` (esfera; virola
  D11=2 SÍ es cilindro). D168 declara el hand-off. Decisión 3.
- [x] **Step 5: `verificar.py §6e`** recalcula e (7.175) y S_w (62.08/124.16/248.32) en Excel,
  exactos. pytest 241 · verificar.py **0 fallos** · seed → APTO.

> **CORRECCIÓN de la nota del plan: el caso semilla NO cambia de valor.** El plan preveía un
> cambio por (a) ec.5 literal sin kf y (b) g. Pero (a) para cilindro la ec.5 literal es
> **idéntica** al S_w anterior (F_m con kf=0.5 = P·Dm/2 y 6·F_m·e/T² = 3·P·Dm·e/T²), y (b) `g`
> es entrada nueva con default 0 (no se reutiliza el 1.5 de D31). Resultado: `verificar.py §7`
> sigue en APTO con los mismos valores, **sin re-baseline del oracle**. El cambio de valor real
> es solo en esfera (kf=0.25 → ahora NA/hand-off, que es lo correcto por decisión 3).

---

## Fase 6 — Paso 6: conformado en frío, simple y doble, con factor (1 − R_f/R_o) (212-3.5)

**Files:** Modify `build_parche_art212()` (deformación en D77, verificación D86). Modify tests.

**Fuente (resources/):** ec. (7) simple `(50·T/R_f)·(1 − R_f/R_o) ≤ 5%` (bloque [75]); ec. (6)
doble `(75·T/R_f)·(1 − R_f/R_o) ≤ 5%` (bloque [71]); `R_o = ∞` si originalmente plano
(bloque [73]); > 5% → PWHT post-conformado (bloque [76]).

**Interfaces:** La deformación pasa de `50·T/Rf` (hardcode simple, sin `(1−Rf/Ro)`) a la
fórmula completa, con **rama simple/doble** según geometría, y entrada `R_o`.

- [x] **Step 1: test (falla).** `test_paso6_conformado`: Ro en D172, D77 con coef 50/75 y factor. PASS.
- [x] **Step 2: entrada `R_o`.** D172 «Radio original de línea media» [mm], blanco = plano (Ro=∞,
  factor 1) [73]. `factor = IF($D$172="",1,1-$D$56/$D$172)`.
- [x] **Step 3: coef por geometría.** `coef = IF($D$11=3,75,50)` (esfera/cabezal doble [71];
  tubería/virola simple [75]). D173 declara la rama.
- [x] **Step 4: deformación y dictamen.** D77 = `IF($D$11=3,75,50)*$D$29/$D$56*IF($D$172="",1,
  1-$D$56/$D$172)`. El dictamen ≤5% ya existe en F86 (verificación sección 5); >5% → PWHT
  (212-3.5b), ya en el AND de F90 vía F86. Seed: %Elong=2.389 (sin cambio).
- [x] **Step 5: PASS + commit.** §6e recalcula %Elong en Excel. pytest 242 · verificar.py **0**.

---

## Fase 7 — Paso 7: fabricación (212-4)

**Files:** Modify `comentar_art212_base` y la Sección 6 de especificaciones. (El único ítem
calculable —`e` con `g`— ya se cerró en la Fase 5.)

**Fuente (resources/):** corte térmico → esmerilar 1.5 mm (bloque [80]); T > 25 mm → MT/PT
por laminaciones (bloque [80]); prep. de superficie 40 mm a cada lado (bloque [85]); costuras
existentes esmeriladas a ras + MT/PT (bloque [86]); separación ≤ 5 mm (bloque [82]); secuencia
(costuras internas primero, luego perímetro) (bloque [88]); venteo de gas (bloque [90]).

- [x] **Step 1: aviso condicional de MT/PT por espesor.** D178 = `IF($D$29>25,"...MT/PT
  (laminaciones), 212-4a","...")`. Anexo Paso 7 (fila 178).
- [x] **Step 2: aviso de separación.** D177 = aviso de g (>5 no admisible; ≥1.5 → e incluye g,
  ver Paso 5). 212-4(c) [82]. Se deja como **aviso** (constraint de fabricación, no de
  aceptación de diseño), fiel al plan; no entra al AND de F90.
- [x] **Step 3: notas de fabricación citando `resources/`.** D179 (anexo): corte térmico 1.5 mm
  [80], metal blanco 40 mm a cada lado [85], secuencia costuras internas→perímetro [88], venteo
  [90]. **Se añaden en el anexo** (no se reescriben las celdas B94-B99 del oracle) para mantener
  direcciones estables (estrategia elegida). pytest 243 · verificar.py **0 fallos**.

---

## Fase 8 — Paso 8: NDE y prueba de hermeticidad, con energía neumática (212-5/6 + App. 501)

> **Depende de la Fase 0.2.** Si el vacío de `resources/` de la ec. (II-1) no se repara
> (camino a), esta fase se **difiere** y se documenta como límite (hidrostática como caso base).

**Files:** Modify `build_parche_art212()` (nueva sub-sección de prueba). Modify tests, `§6e`.

**Fuente (resources/, tras Fase 0.2):** NDE 100 % MT/PT del perímetro (bloque [92]); RT/UT de
costuras del parche (bloque [94]); ec. (II-1)/(II-2) energía almacenada en `k`; (II-3)
`TNT = E/4 266 920` kg (App. 501-II bloque [13]); (III-1) `R = R_scaled·(2·TNT)^(1/3)`
(App. 501-III bloque [4]); Tabla 501-III-1-1 de `R_scaled`; hidrostática según código de
post-construcción (bloque [99]).

**Interfaces:** Selector de tipo de prueba (hidrostática/neumática). En neumática, el motor
calcula E, TNT y la distancia segura R, y avisa; en hidrostática, mantiene el cálculo actual
(1.5×P_diseño). Todo dato de la ec. (II-1) sale del JSON de la Fase 0.2, nunca de memoria.

- [x] **Step 1: test (falla).** `test_paso8_prueba`: selector, E/TNT/R, R_scaled dv_list. PASS.
- [x] **Step 2: entradas de prueba.** Anexo Paso 8 (filas 181-193): D183 selector, D184 V [m³],
  D185 Pat [MPa abs], D186 Pa [MPa abs]=0.101, D187 k=1.4, D188 R_scaled (dv_list de la Tabla
  501-III-1-1 leída de resources/, criterios en el comentario).
- [x] **Step 3: fórmulas.** E (D189) = ec.(II-1) **general en k** (`[1/(k-1)]·Pat·V·[1-(Pa/Pat)^
  ((k-1)/k)]`, Pat→Pa ×1e6); TNT (D190) = E/4 266 920; R (D191) = 30 m si E≤8 130 000 J, si no
  `R_scaled·(2·TNT)^(1/3)`. **Regla n.1: todos los coeficientes los lee `leer_energia_501()` de
  `resources/`** (App. 501, Fase 0.2); el build ABORTA si el apéndice no está reparado. Se
  añadió `energia_501=` (keyword) a `build_parche_art212`; main() lo pasa, el test un stub.
- [x] **Step 4: dictamen** D192 (neumática → distancia R + Tabla 501-III-2-1; hidrostática →
  1.5×P). Bloque neumático = NA() en hidrostática (default). NDE D193 (212-5 [92],[94]).
- [x] **Step 5: `verificar.py §6e`** recalcula E/TNT/R en Excel (qa forzado a neumática, V=100,
  Pat=5 → rama eq(III-1)): E=840 MJ, TNT=196.9 kg, R=146.6 m, exactos. pytest 244 · verificar.py
  **0 fallos**. **Requirió corregir la Fase 0.2**: (II-3)/(II-5) estaban reordenadas ('TNT =
  (kg) E …'); se normalizaron a la forma canónica para que el motor lea el divisor.

---

## Fase 9 — Notación de símbolos con subíndice (decisión 5)

**Files:** Modify `build_db_materiales.py` (helper `sym()`); aplicarlo en los rótulos de
símbolo (columna B) de todas las secciones del 212. Modify `TestSistemaVisual` si hace falta.

**Interfaces:** `sym(base, sub)` → `CellRichText` con el subíndice en `InlineFont(vertAlign=
"subscript")`, fuente `MONO`. Reemplaza `"F_m"`, `"P_op"`, `"C_sw"`, `"w_mín"`, etc. por su
forma con subíndice real. No cambia ninguna fórmula (los símbolos de columna B son rótulos,
no referencias).

- [ ] **Step 1: test (falla).** `test_simbolos_sin_guion_bajo`: recorre la columna B del 212 y
  afirma que **ninguna** celda de símbolo contiene `"_"` en su texto plano y que las que llevan
  subíndice son `CellRichText`. → FAIL.
- [ ] **Step 2: helper `sym()`** con `CellRichText`/`TextBlock`/`InlineFont(vertAlign="subscript")`.
- [ ] **Step 3: aplicar** a los símbolos de todas las secciones (F_CP→F con subíndice CP, etc.).
- [ ] **Step 4: comprobar `TestSistemaVisual`** (el rich text usa `MONO`; si la prueba no
  contempla `CellRichText`, ampliar su recorrido para leer los runs). PASS + commit.

---

## Fase 10 — Re-baselinar el *oracle*, regenerar, verificar y entregar

**Files:** Regenerate `parche_art212_ref.json`; Modify `test_dashboard.py`
(`TestParidadHojaParche` re-armado), `verificar.py §7`; Modify `CLAUDE.md` (estado del motor);
Modify este plan (estado final).

- [ ] **Step 1: build completo.** `python build_db_materiales.py --resources … --in … --out
  ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm`. Sin abortos.
- [ ] **Step 2: los tres gates sin recálculo.** `pytest test_build_db.py test_dashboard.py
  test_secii_tablas.py -q` → verde (las anclas por paso de `TestBuildParcheContraOracle`
  cubren las fórmulas nuevas; `TestSistemaVisual` y `TestSincroniaPythonVba` pasan).
- [ ] **Step 3: entregar para F9 en Excel.** `SendUserFile` del `.xlsm`. Pedir al ingeniero:
  F9 del caso semilla con las fórmulas nuevas, y confirmar los valores esperados (que **cambian**
  respecto de Rev4 por la ec. 5 literal y la posible g). Registrar los nuevos valores de
  referencia del caso semilla.
- [ ] **Step 4: re-armar el *oracle* y `verificar.py §7`.** Con la hoja validada, correr
  `_dump_parche_ref.py` para regenerar `parche_art212_ref.json` desde el `.xlsm` validado, y
  actualizar los valores esperados de `verificar.py §7` (dictamen del caso semilla). Re-armar
  `TestParidadHojaParche` contra el nuevo *oracle*.
- [ ] **Step 5: `verificar.py` completo en Windows** (§1-10, incl. §6e y §7 nuevos) → 0 fallos.
- [ ] **Step 6: documentar.** Actualizar `CLAUDE.md` (sección del motor 212: ahora 8 pasos
  explícitos; cita las Fases 0). Marcar este plan como ejecutado. Commit.

---

## Verificación del plan contra el Art. 212 (evidencia de aprobación)

Trazado de cada cláusula/ecuación del Art. 212 (y App. 501) a la tarea que la implementa,
cruzado contra `resources/` (Regla nº 1). El flujo del ingeniero es fiel al código en las 7
ecuaciones y en la regla de la separación; este plan cierra la brecha entre el flujo y el motor.

| Cláusula / ec. (resources/) | Requisito | Tarea | Estado tras el plan |
|---|---|---|---|
| 212-1(e) [bloque 6] | T entre nil-ductility y 345 °C; <0 entalla; >345 creep/fatiga | Fase 1 | Compuerta que bloquea/avisa por T |
| 212-2(c) [11-13] | daño caracterizable; grietas sólo con FFS | Fase 1 | Dictamen de elegibilidad |
| Servicio letal (flujo + Part 1 [10]) | prohibido → Art. 201 | Fase 1 | Bloqueo con motivo |
| 212-3.1(a) [16] | evaluar flexión/torsión/viento/fatiga | Fase 2 | Entradas F_CO/F_LO |
| 212-3.2 ec. (1) [29] | F_CP = P·Dm/2 | Fase 2 | Literal cilindro |
| 212-3.2 ec. (2) [fig 33, imagen en resources/] | **F_LP = P·Dm/4** | **Fase 0.1 + Fase 2** | Recuperada de la imagen de `resources/`, fijada en texto, luego codificada |
| 212-3.2(b) [37],[39] | F_C=F_CP+F_CO; F_L=F_LP+F_LO; F_max | Fase 2 | Formados y gobernante |
| 212-3.2(c) [44] | alternativas esfera/toriesférico | Fase 2/5 | Hand-off declarado (decisión 3) |
| 212-3.3 ec. (3) [47] | L_min = 2·(Rm·t)^½ | Fase 3 (ya existía) | Conservado |
| 212-3.3 [51],[53] | parches adyacentes; 360° si < L_min | Fase 3 | 3C nuevo; 3B ya existía |
| 212-3.4 ec. (4) [57],[59] | w_min = F_max/(E·Sa), E=0.55 | Fase 4 | Usa F_max |
| 212-3.4 NOTA [60] | w ≤ min(T,t) y ≤ 40 mm | Fase 4 | Dos verificaciones nuevas |
| 212-3.4(b) [61] | bisel, garganta ≤ nominal | Fase 4 | Nota |
| 212-3.4(c) ec. (5) [65],[67] | S_w = PDm/2T + 3PDm·e/T² ≤ 1.5Sa; e=(T+t)/2 | Fase 5 | Literal cilindro |
| 212-4(c) [82] | g ≥ 1.5 mm → e = (T+t+g)/2 | Fase 5 | **Defecto corregido** |
| 212-3.5 ec. (6)/(7) [71],[75] | doble 75 / simple 50, ×(1−Rf/Ro) ≤ 5% | Fase 6 | Rama + factor completos |
| 212-3.5(b) [76] | > 5% → PWHT | Fase 6 | Dictamen |
| 212-4 [80],[85],[86],[88],[90] | fabricación, MT/PT>25mm, prep, secuencia, venteo | Fase 7 | Notas + aviso condicional |
| 212-5 [92],[94] | NDE 100% MT/PT; RT/UT costuras | Fase 8 | Notas |
| 212-6 [99] | prueba según código; neumática con precaución | Fase 8 | Selector |
| App. 501-II (II-1) [vacío en resources/] | energía almacenada en k | **Fase 0.2 + Fase 8** | Vacío real: reparar resources/ (a) o diferir (b) |
| App. 501-II (II-3) [13] | TNT = E/4 266 920 | Fase 8 | Texto disponible |
| App. 501-III (III-1) [4] + Tabla III-1-1 | R = R_scaled·(2TNT)^⅓ | Fase 8 | Texto/tabla disponibles |

**Huecos de `resources/` que el plan cierra antes de codificar (Regla nº 1):** la ec. (2) del
Art. 212 (imagen presente en `resources/` → recuperada leyendo el PNG) y la ec. (II-1)/(II-2)
del App. 501-II (**vacío real** de `resources/`, sin texto ni imagen → reparar re-extrayendo,
o diferir la Fase 8). Sin cerrarlas, sus tareas no arrancan. En ningún caso se pide un folio
externo: `resources/` es la fuente.

---

## Verificación — qué se puede aquí y qué no

- **Aquí (sin recálculo):** `pytest` (tres suites), build sin abortos, `verificar.py §1-5`,
  y las anclas por paso de `TestBuildParcheContraOracle` (igualdad de cadena de las fórmulas
  nuevas contra lo que este plan especifica, trazado a `resources/`).
- **Sólo en Excel de Windows (ingeniero):** `verificar.py §6e/§7` (recálculo real de las
  fórmulas nuevas y del caso semilla), F9 del caso precargado, y la revisión visual de la hoja
  exportada a PDF/PNG (openpyxl miente sobre bordes de rango fusionado). El *oracle* se
  re-baselina **después** de esa validación.

## Riesgos

- **Regla nº 1 en los dos coeficientes que no eran texto.** La ec. (2) del Art. 212 **ya está
  resuelta desde `resources/`** (imagen del bloque 33). La ec. (II-1) del App. 501 es un
  **vacío real de `resources/`**: la Fase 8 solo arranca por el camino (a) —reparar el JSON
  re-extrayendo el apéndice— o se difiere por el camino (b) (hidrostática como caso base). No
  se pide folio externo; la fuente es siempre `resources/`.
- **El caso semilla cambia de valor.** La ec. (5) literal (sin kf) y la posible inclusión de g
  mueven `S_w`, `P_máx` y quizá el dictamen. `verificar.py §7` y el *oracle* deben re-baselinarse
  tras la validación del ingeniero — **no** antes. No cerrar el plan hasta esa confirmación.
- **`luz` (D31) reutilizada como `g`.** Comprobar el significado geométrico actual de `luz` en
  `Rf` (D56) y el desarrollo (D78) antes de reinterpretarla como separación de faying edge; si
  son físicamente distintas, separarlas en dos entradas para no acoplar la excentricidad al radio
  de conformado.
- **`CellRichText` y `TestSistemaVisual`.** El recorrido de la prueba visual debe leer los runs
  del rich text; si sólo mira `cell.font`, ampliarlo, o los subíndices podrían escapar al control.
- **Volumen del cambio.** Es un refactor amplio de una hoja validada; el orden por fases (cada
  una un bloque de filas con su ancla) permite aceptar/rechazar por separado y mantener el
  `.xlsm` construible entre fases.

## Self-review (hecho)

- **Cobertura del flujo:** los 8 pasos tienen fase (tabla de trazado arriba). Los 6 huecos de
  la verificación (elegibilidad, cargas externas, topes de filete, g en e, curvatura doble,
  energía neumática) tienen tarea.
- **Regla nº 1:** ningún coeficiente nuevo sale de memoria; los dos que son imagen se extraen en
  la Fase 0 antes de codificarse. Prohibida la forma aire-only de la energía.
- **Consistencia:** `build_parche_art212(wb, b313, iid1a, iidb, fac_info, rangos)` mantiene su
  firma; `F_max` (Fase 2) es la entrada única de los Pasos 4-5; `e` (Fase 5) alimenta S_w.
- **Sin placeholders:** cada fórmula nueva se da en forma matemática con su bloque de
  `resources/`; la cadena Excel exacta la produce la implementación y la fija `verificar.py §6e`.
