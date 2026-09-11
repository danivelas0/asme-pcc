# Rediseño del motor Art. 206 para igualar el flujo de proceso — Plan de implementación

> **Para ejecutores:** SUB-SKILL REQUERIDA: usar `superpowers:subagent-driven-development`
> (recomendado) o `superpowers:executing-plans` para ejecutar tarea por tarea. Los pasos
> usan casillas (`- [x]`) para el seguimiento. **NO EJECUTAR sin orden explícita del
> ingeniero** (este plan se entrega para revisión, no para correr).

**Objetivo:** Que la hoja `Collar_PCC2_Art206` (motor de cálculo del Art. 206, hoy en
`build_collar_art206()`) reproduzca **los 8 pasos** del flujo de proceso aportado por el
ingeniero (`outputs/plans/flows/flujo_proceso_art_206.md`, transcrito del PDF
`asmepcc2articulo206flujo.pdf`), cerrando los huecos detectados al cotejar el motor actual
contra ese flujo: **cateto del filete de extremo no calculado**, **luz radial `G ≤ 2,5 mm`
sin verificar**, **sobreespesor de corrosión C.A. ausente del `t_req` Type B**, **selección
de tipo sin guía**, y las **precauciones/avisos** del 206-2/3/4/5/6 que el flujo pide y la
hoja aún no trae (corrosión bajo-manga, interferencia con costura previa/bulge, fatiga,
soldadura en servicio Art. 210, reducción de presión 50-80 %, purga N₂, relleno de cavidades,
NDE diferido, prueba de hermeticidad del anular).

**Arquitectura:** `build_collar_art206()` (7265-7565, 100 % en código, hermana de
`build_parche_art212`) se **reestructura por pasos 1-8 explícitos**, conservando la
reutilización de `construir_seccion7_material()` (Sección de resolución de material,
compartida con el 212 — su comportamiento **no cambia**). Toda regla o número nuevo se
**traza a `resources/`** (Regla nº 1). El Art. 206 **no tiene ecuaciones propias**
(`counts.equations = 0`): casi todo su diseño está en la **capa de texto** del JSON
(⅔, 100 mm, 50 mm, 1,4×, 0,80/1,00, 2,5 mm, 50-80 %, C.A. por diseño). El **único** dato que
no es texto —el cateto del filete `w = T_s + G` / `w_máx = 1,4·T_p + G`— vive **como imagen**
en `resources/` (`fig_206_3_5_1.png` / `fig_206_3_5_2.png`) y **ya se recuperó leyendo ambos
PNG** (confirmado en el diseño de este plan). La Fase 0 lo fija en la capa de texto del JSON
para citarlo/trazarlo sin reabrir la imagen. **No hay ningún vacío real** en `resources/`
para el Art. 206 (a diferencia del 212, que sí tenía uno en el App. 501).

**Decisiones de alcance del ingeniero (2026-09-10, esta conversación):**
1. **Selección de tipo A/B: recomendación guiada** (como la compuerta del 212). Entradas
   ¿fuga/puede fugar? y ¿reduce resistencia axial / tasa de daño no clara?; el motor
   **recomienda** el tipo y **avisa** si `D22` contradice la recomendación. `D22` sigue siendo
   la decisión del ingeniero (advisory, **no** bloqueo: el tipo es una decisión de diseño del
   206, no una compuerta de elegibilidad como en el 212).
2. **Sobreespesor de corrosión C.A.: se añade** como entrada tecleable y se **suma al `t_req`
   Type B** (206-3.3). Cambia el valor del caso semilla → re-validar en Excel.
3. **Prueba de hermeticidad (206-6): mínima, fiel al 206** — selector prueba del anular /
   prueba sensible de fugas + nota de presión limitada (que el tubo interno no colapse) y
   remisión al Art. 501. **Sin** ecuaciones de energía almacenada: el texto del 206 no las
   publica (Regla nº 1); **no** se importa el motor neumático del App. 501 del 212.
4. **Notación: subíndices reales** en los símbolos visibles (helper `sym()`), como el 212.

**Stack:** Python 3 + openpyxl; pytest; el builder `build_db_materiales.py`; Excel de Windows
para la validación final (`verificar.py`, F9 del caso semilla). openpyxl `CellRichText` +
`InlineFont(vertAlign="subscript")` para los símbolos con subíndice.

**Spec:** `outputs/plans/flows/flujo_proceso_art_206.md` (flujo transcrito) + el cotejo motor
↔ flujo de esta conversación. Fuente normativa única: `resources/ASME PCC/pcc_2/
p2_welded_repairs/art_206_full_encirclement_steel/art_206.json` y su espejo
`part_2_welded_repairs/article_206_full_encirclement_steel/…json` (Art. 206), más las figuras
`fig_206_3_5_1.png` / `fig_206_3_5_2.png`. Art. 210 (`…/art_210…json`) solo para citar sus
cláusulas en los avisos de soldadura en servicio, **no** alimenta ningún cálculo.

## Global Constraints

Copiadas de `asme-pcc/CLAUDE.md` («Reglas de diseño del libro» + Regla nº 1). Todas las
tareas las incluyen implícitamente:

- **Regla nº 1 — ningún valor normativo sale de memoria.** Cada regla/número nuevo cita el
  bloque de `resources/` del que procede. Las dos fórmulas de cateto que hoy son imagen se
  fijan en el JSON (Fase 0) **antes** de que el motor las consuma. Prohibido tomar cualquier
  número del 206 de `knowledge/claude.md` (derivado) o de memoria.
- **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`,
  `VSTACK`, `_xlfn`. Sólo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`. `verificar.py §5` falla si
  aparece una.
- **Un texto no se guarda como fórmula** (`="…"` > 255 car. → `_xlfn._LONGTEXT`). Usar el
  normalizador existente (`normalizar_textos_como_formula`); no introducir literales `="..."`
  largos nuevos. Los avisos de aviso condicional (`IF(cond,"texto…","")`) deben quedar por
  debajo del límite o materializarse como celda de texto con `IF` corto.
- **Validación de datos: sólo rango literal o lista de ítems**, nunca fórmula como origen.
  `dv_list(ws, cell, '"a,b,c"', com)` para categóricos.
- **Regla 12/14 — todo material y todo valor tabulado se elige de una base (`DB_*`), no se
  teclea.** Las entradas categóricas nuevas del Paso 1 (¿fuga?, ¿axial?, ¿servicio cíclico?,
  ¿defecto externo?) **no** están tabuladas: son lista de ítems literal (permitido; no son
  materiales). Las de proceso/campo (C.A., luz `G`, longitud del defecto, presiones,
  temperatura) se teclean (regla 14). NPS/cédula del tubo portador siguen saliendo de
  `Datos_Ref` por desplegable; el material del sleeve sigue saliendo de la cascada de la
  Sección de resolución de material.
- **Regla 13 — los motores no se cruzan.** El motor 206 lee sólo de `DB_*` (vía la Sección 7),
  nunca de un buscador ni del motor 212.
- **snake_case** en todo archivo nuevo. **Español, SI por defecto** (MPa, mm, °C).
- **Entrega parchando sobre Rev4** por el builder, no una Rev5.
- **Sistema visual Swiss Industrial Print:** sólo tokens del sistema (`PAPEL`, `TINTA`, `ROJO`,
  `MONO`, `MACRO`, `IN_FILL`, `CAJA_CAMPO`, `BAND_FILL`, etc.). `TestSistemaVisual` recorre las
  celdas con formato y falla ante cualquier color/fuente fuera de paleta. Las celdas que se
  teclean llevan `CAJA_CAMPO` (línea inferior de tinta), como el resto de la hoja hoy.
- **Sincronía Python↔VBA:** este refactor ocurre **íntegro dentro de `Collar_PCC2_Art206`**;
  **no** añade hojas navegables nuevas, así que `NAVEGABLES` / `HojasNavegables()` **no** se
  tocan. Confirmarlo (que `TestSincroniaPythonVba` siga en verde sin cambios de VBA).
- **`ws.protection`** se mantiene (`password="0000"`, `sheet=True`): las celdas de entrada
  llevan `Protection(locked=False)` vía el helper `inp`, como hoy.

---

## Estructura de archivos

- `scripts/completar_art_206_filete.py` — **Create.** Fija en la capa de texto del JSON del
  Art. 206 (los **dos** espejos) las dos ecuaciones de cateto recuperadas de las figuras:
  `w = T_s + G` (Fig. 206-3.5-1, bloque `figure` 72/su equivalente en el espejo) y
  `w_máx = 1,4·T_p + G` (Fig. 206-3.5-2). Idempotente, `--resources`; **lee los PNG
  referenciados por los propios bloques** (no un folio externo), escribe su `text` con
  `extraction_amendments` (bloque, ruta del PNG, nota «recuperada de la imagen del propio
  `resources/`»), no toca ningún otro valor. Aborta si el bloque ya no apunta a esos PNG.
- `scripts/build_db_materiales.py` — **Modify.** Reestructura `build_collar_art206()`
  (7265-7565) por pasos 1-8; añade/consume el helper de símbolos con subíndice `sym()`;
  secciones nuevas (selección guiada de tipo, C.A., cateto de filete, luz `G`, avisos de
  fatiga / soldadura en servicio / reducción de presión / cavidades / NDE / prueba). Lee de
  las bases (vía Sección 7) y del JSON de la Fase 0. La firma
  `build_collar_art206(wb, b313, iid1a, iidb, fac_info, rangos)` **no cambia**; su llamada en
  `main()` tampoco.
- `scripts/test_dashboard.py` — **Modify.** Nueva clase `TestBuildCollarArt206` con anclas por
  paso: fija por **igualdad de cadena** las fórmulas nuevas (cateto `w`, tope `G≤2,5`, `t_req`
  con C.A., tipo recomendado, rango de presión 50-80 %) contra lo que este plan especifica
  (trazado a `resources/`). Amplía la comprobación de que `Collar_PCC2_Art206` sigue visible/
  navegable y sin fórmulas de matriz dinámica.
- `scripts/verificar.py` — **Modify.** Nueva sección **§6f (recálculo del motor Art. 206 en
  Excel real)** que recalcula `w`, `w_máx`, el tope `G≤2,5`, el `t_req`+C.A. y el tipo
  recomendado del **caso semilla del 206** desde la **misma expresión que emite el motor**
  (patrón §6d/§6e: la función que genera la fórmula vive en `build_db_materiales.py` y la
  emiten el motor **y** la prueba). Registrar el caso semilla del 206 y sus valores esperados
  (que **cambian** al añadir C.A.).
- `scripts/CLAUDE.md` del proyecto — **Modify (Fase 10).** Actualizar la sección del motor 206
  (ahora 8 pasos explícitos; citar la Fase 0 y las decisiones 1-4).

> **Nota sobre el *oracle*.** El 212 tiene `parche_art212_ref.json`; el **206 no tiene
> *oracle***. Hoy se verifica por `verificar.py §1-5` (auditorías fila a fila de las `DB_*`,
> ausencia de matriz dinámica, contigüidad de cascada — book-wide) + las pruebas de navegación/
> visual, **sin** recálculo de sus fórmulas propias. Este plan **añade** ese recálculo (§6f) y
> las anclas de cadena (`TestBuildCollarArt206`); no crea un *oracle* JSON — las fórmulas
> nuevas se fijan por cadena y se recalculan en Excel, que es suficiente para una hoja 100 % en
> código.

---

## Fase 0 — Cerrar el hueco de Regla nº 1 (cateto del filete) **dentro de `resources/`** (prerrequisito)

> **`resources/` es la fuente; no se pide ningún folio externo.** El texto de 206-3.5
> (bloques 42-47) solo enuncia la regla de rama del 1,4×; las fórmulas de cateto están en las
> **figuras** de `resources/`. **Ya confirmadas leyendo ambos PNG en el diseño de este plan:**
> Fig. 206-3.5-1 imprime `T_s + G` (para `T_s ≤ 1,4 T_p`); Fig. 206-3.5-2 imprime
> `1,4·T_p + G` (para `T_s > 1,4 T_p`, con chaflán opcional). Las leyendas definen `G = gap`,
> `T_p = carrier pipe required minimum wall thickness`, `T_s = Type B sleeve nominal wall
> thickness`.

### Tarea 0.1 — Fijar las ec. de cateto `w = T_s + G` y `w_máx = 1,4·T_p + G` desde las imágenes de `resources/`

**Files:** Create `scripts/completar_art_206_filete.py`; Modify el JSON del Art. 206 (los dos
espejos: `…/art_206_full_encirclement_steel/art_206.json` y
`…/article_206_full_encirclement_steel/…json`).

**Fuente (resources/):** los bloques de figura que apuntan a `fig/fig_206_3_5_1.png` (bloque
72 en el espejo `p2_…`) y `fig/fig_206_3_5_2.png` (bloque 77) tienen `text` vacío. Leídas las
imágenes, imprimen literalmente `T_s + G` y `1,4·T_p + G` respectivamente. El dato ya está
confirmado desde `resources/`; esta tarea solo lo fija en la capa de texto para que sea
citable/trazable sin reabrir la imagen.

- [x] **Step 1: confirmar el ancla.** Bloques 72 (`fig_206_3_5_1.png`) y 77 (`fig_206_3_5_2.png`),
  `type=figure`, `text` vacío, en los dos espejos. Confirmado.
- [x] **Step 2: escribir `completar_art_206_filete.py`** — idempotente, defensivo; **leídos
  ambos PNG**: Fig. 206-3.5-1 imprime `T_s + G` (Ts≤1.4Tp), Fig. 206-3.5-2 imprime `1.4·T_p + G`
  (Ts>1.4Tp, chaflán opcional). Fija esos textos + `extraction_amendments` (bloque, ruta, SHA-256).
- [x] **Step 3: correr una vez por espejo** y `git diff`: 52 ins / 6 del, solo los bloques 72/77
  y el amendment por espejo.
- [x] **Step 4: commit** — `Art. 206: fija las ec. de cateto w=Ts+G y w_max=1.4Tp+G desde las figuras de resources/ (Regla n.1)`.

---

## Fase 1 — Paso 1: clasificación y selección guiada de tipo (206-1, 206-2)

**Files:** Modify `build_collar_art206()` — insertar en la Sección 1 (datos) las entradas de
selección y, en «Verificaciones y avisos», el dictamen de tipo recomendado y los avisos de
206-2. Modify `test_dashboard.py`.

**Fuente (resources/, Art. 206):** 206-1.1.1 (Type A: mecanismo y tasa de daño entendidos; no
fuga; no evolutivo) [bloque 5]; 206-1.1.2 (Type B: fugas o defectos que pueden fugar; refuerzo
axial) [bloque 6]; 206-2.5 (Type A no apto para defectos circunferenciales) [bloque 17];
206-2.6 (corrosión bajo-manga Type A → sellante/recubrimiento) [bloque 19]; 206-2.7
(interferencia con costura previa → esmerilar + RT/UT **o** bulge, Fig. 206-2.7-1) [bloque 21];
206-2.3 (fuga activa Type B → aislar antes de soldar) [bloque 13].

**Interfaces:** Produce las entradas del Paso 1 y una celda de **tipo recomendado** más un
**aviso de contradicción** si `D22` no coincide. El dictamen es advisory: **no** bloquea el
DICTAMEN GLOBAL (el tipo es decisión de diseño; el ingeniero puede sobreponerse).

- [x] **Step 1: test (falla).** En `TestBuildCollarArt206`, `test_paso1_seleccion_tipo`:
  construye la hoja y afirma que existen las entradas nuevas (`¿fuga o puede fugar?` con
  `dv_list '"Si,No"'`; `¿reduce resistencia axial / tasa de daño no clara?` con
  `dv_list '"Si,No"'`), que la celda de tipo recomendado devuelve `Type B` cuando cualquiera es
  `Si` y `Type A` cuando ambas son `No`, y que el aviso de contradicción se activa cuando
  `D22` (tipo elegido) ≠ recomendado. → FAIL.
- [x] **Step 2: entradas del Paso 1.** En la Sección 1: `¿Fuga o puede fugar?`
  (`dv_list '"Si,No"'`, ref 206-1.1.2/206-2.3), `¿Reduce resistencia axial o tasa de daño no
  clara?` (`dv_list '"Si,No"'`, ref 206-1.1.1/206-2.5). Mantener `¿Defecto circunferencial?`
  (ya existe, D24, 206-2.5).
- [x] **Step 3: tipo recomendado y aviso.** Celda `Tipo recomendado` =
  `IF(OR(fuga="Si", axial="Si"), TIPO_B, TIPO_A)` (citar 206-1.1.1/1.1.2). Aviso de
  contradicción = `IF($D$22<>tipo_recomendado, "AVISO: el tipo elegido difiere del recomendado
  por 206-1.1 (fuga/axial) — confirmar", "")`.
- [x] **Step 4: avisos de 206-2.** Añadir en «Verificaciones y avisos» (avisos condicionales
  `IF(cond,"…","")`, texto corto):
  - Corrosión bajo-manga: `IF($D$22=TIPO_A, "206-2.6: evaluar corrosion bajo manga; aplicar sellante/recubrimiento si aplica", "")`.
  - Interferencia con costura previa: aviso fijo `206-2.7: costura previa prominente puede impedir fit-up → esmerilar + RT/UT o manga con bulge (Fig. 206-2.7-1)`.
  - Fuga activa Type B: `IF(AND($D$22=TIPO_B,fuga="Si"), "206-2.3: aislar la fuga antes de soldar (ver 206-4.3 purga N2)", "")`.
- [x] **Step 5: test PASS + commit** — `Art. 206 Paso 1: seleccion guiada de tipo y avisos 206-2`.

---

## Fase 2 — Paso 2: espesor requerido con sobreespesor de corrosión (206-3.1/3.2/3.3)

**Files:** Modify `build_collar_art206()` (sección «Cálculo de espesor requerido», D51-D55, y
la entrada del C.A. en la Sección 1). Modify `test_dashboard.py`, `verificar.py §6f`.

**Fuente (resources/):** 206-3.1 (Type A ≥ ⅔·T_p; esfuerzos longitudinales del portador por el
código) [bloques 31,33]; 206-3.2 (Type B ≥ t para máx. presión admisible; `E = 0,80` o `1,00`
si 100 % UT; axial/flexión si junta circunferencial defectuosa) [bloques 35,36]; 206-3.3
(`t_req` por el código de construcción; material y `S` por el código; **«Corrosion allowances
applied shall be in accordance with the engineering design»**) [bloque 38].

**Interfaces:** El `t_req` Type B (D53/E53/F53) pasa a **sumar C.A.** El Type A (⅔·T_p, D54)
**no** lleva C.A. (206-3.1 no es componente a presión). `E = 0,80/1,00` (D38) y la forma por
código (D11) se conservan.

- [x] **Step 1: test (falla).** `test_paso2_treq_con_ca`: afirma que existe la entrada `C.A.`
  [mm] (tecleable), que `t_req` Type B = (forma actual por D11) **+ C.A.** en las tres columnas
  (Operación/Diseño/Envolvente), y que el Type A (⅔·T_p) **no** incluye C.A. → FAIL.
- [x] **Step 2: entrada C.A.** En la Sección 1 (datos de campo): `Sobreespesor de corrosion
  C.A.` [mm], `inp` tecleable, ref «206-3.3: por el diseño de ingenieria». Default 0 es un
  valor válido tecleado (no un default oculto).
- [x] **Step 3: `t_req` + C.A.** Cambiar D53/E53/F53 a `= (expresion actual por D11) + $CA`.
  Conservar las tres ramas (D11=1 B31.3, D11=3 esfera VIII-1, D11=2 virola VIII-1). Citar
  206-3.3.
- [x] **Step 4: gobernante.** `T_s,min gobernante` (D55) sigue = `IF(D22=TIPO_A, D54, F53)`
  (F53 = envolvente, el caso más exigente, ya con C.A.). Type A sin cambios.
- [x] **Step 5: `verificar.py §6f`** recalcula `t_req`+C.A. del caso semilla en Excel desde la
  misma expresión que emite el motor. PASS + commit — `Art. 206 Paso 2: t_req Type B con C.A. (206-3.3)`.

---

## Fase 3 — Paso 3: dimensiones del sleeve (206-3.4)

**Files:** Modify `build_collar_art206()` (verificación de longitud, D61). Modify tests.

**Fuente (resources/):** 206-3.4 — «at least 100 mm (4 in.) long and extend beyond the defect
by at least 50 mm (2 in.)» [bloque 40].

**Interfaces:** La verificación `L_s ≥ MÁX(100, defecto + 2·50)` ya existe (D61) y es correcta;
esta fase solo la deja **explícita y trazada**, y añade la lectura del cumplimiento del
sobrepaso de 50 mm a cada lado como sub-nota (no un cálculo nuevo).

- [x] **Step 1: test (falla).** `test_paso3_dimensiones`: `L_s,min = MAX(100,$D_defecto+2*50)`
  y su dictamen CUMPLE/NO CUMPLE; nota de «≥ 50 mm a cada lado». → FAIL (por la nota/anclaje).
- [x] **Step 2: confirmar/anclar** la fórmula existente y su cita 206-3.4; añadir la nota de
  sobrepaso a cada lado en `comentar`/columna de referencia. (Fase ligera: casi todo ya está.)
- [x] **Step 3: PASS + commit** — `Art. 206 Paso 3: dimensiones 206-3.4 explicitas`.

---

## Fase 4 — Paso 4: filete de extremo con cateto calculado y luz radial (206-3.5, 206-4.1)

**Files:** Modify `build_collar_art206()` (rama de filete F64; nueva sección de cateto y de luz
`G`). Modify tests, `verificar.py §6f`.

**Fuente (resources/):** 206-3.5(a) filete completo si `T_s ≤ 1,4·T_p` [bloque 44]; 206-3.5(b)
si `T_s > 1,4·T_p`, extremos as-is o achaflanados [bloque 45]; **cateto** `w = T_s + G` /
`w_máx = 1,4·T_p + G` (Figs. 206-3.5-1/2, **Fase 0**); transición suave del pie sin muesca
[bloque 46]; **luz radial** «no gap; radial gap of up to 2.5 mm (3/32 in.) maximum» (206-4.1)
[bloque 63].

**Interfaces:** El motor deja de solo **elegir la rama** (texto en F64) y pasa a **calcular el
cateto** `w`, con la verificación de la luz `G ≤ 2,5 mm`. `T_p` = espesor del tubo portador
(D21). `G` = luz radial (D31, ya existe como entrada). Solo aplica a Type B (Type A no lleva
soldadura circunferencial de extremo, 206-1.1.1).

- [x] **Step 1: test (falla).** `test_paso4_filete_cateto`: en Type B,
  `w = IF($D$29<=1.4*$D$21, $D$29+$G, 1.4*$D$21+$G)` (Ts=D29, Tp=D21, G=luz radial); rama de
  texto conservada; y verificación de luz `IF($G<=2.5,"CUMPLE","NO CUMPLE — excede 2,5 mm
  (206-4.1)")`. En Type A, `w` = «No aplica». → FAIL.
- [x] **Step 2: cateto `w`.** Nueva fila «Cateto del filete de extremo `w`» [mm]:
  `IF($D$22=TIPO_B, IF($D$29<=1.4*$D$21, $D$29+$Gluz, 1.4*$D$21+$Gluz), "No aplica - Type A
  (206-1.1.1)")`. Citar Figs. 206-3.5-1/2 (Fase 0). **Confirmar la referencia de la luz**: hoy
  D31 = «Luz radial sleeve-portador»; usarla como `G`. (Ver Riesgos: `T_p` nominal vs. required
  minimum de la leyenda.)
- [x] **Step 3: rama de texto.** Conservar F64 (filete completo 206-3.5a / as-is-chaflán
  206-3.5b) tal como está, ahora acompañada del `w` numérico.
- [x] **Step 4: verificación de luz `G ≤ 2,5 mm`.** Nueva fila en «Verificaciones y avisos»:
  requerido `2,5`, adoptado `=$D$31`, resultado `IF($D$31<=2.5,"CUMPLE","NO CUMPLE — excede
  2,5 mm (206-4.1)")`. Cablear al DICTAMEN GLOBAL (entra en el AND de las verificaciones).
- [x] **Step 5: `verificar.py §6f`** recalcula `w` (las dos ramas) y el tope `G` del caso
  semilla en Excel. PASS + commit — `Art. 206 Paso 4: cateto del filete (Fig. 206-3.5) y luz G<=2,5 mm`.

---

## Fase 5 — Paso 5: presión externa, relleno de cavidades y bulging (206-3.6/3.7/3.9/3.10)

**Files:** Modify `build_collar_art206()` (avisos + una entrada condicional). Modify tests.

**Fuente (resources/):** 206-3.6 (presión externa sobre el tubo dentro del Type B; ajuste
ceñido / rellenar anular / balancear) [bloque 48]; 206-3.7 (daño externo con Type A → relleno
endurecible de resistencia a compresión adecuada; rellenar vacíos entre Type B y portador)
[bloque 50]; 206-3.9 (restraint of bulging: (a) relleno epoxi si el defecto es externo,
(b) reducir presión de línea al instalar) [bloques 53-56]; 206-3.10 (Type A: contacto íntimo,
filler apropiado) [bloque 58].

**Interfaces:** Entrada condicional `¿Defecto externo / pérdida de pared externa?` que activa el
aviso de relleno endurecible; avisos de presión externa y de reducción de presión al instalar.
Todo son notas/avisos (no cálculo nuevo).

- [x] **Step 1: test (falla).** `test_paso5_cavidades`: entrada `¿Defecto externo?`
  (`dv_list '"Si,No"'`); aviso `IF(externo="Si","206-3.7/3.9: rellenar cavidades con material
  endurecible (epoxi) de resistencia a compresion adecuada","")`; aviso fijo de presión externa
  206-3.6. → FAIL.
- [x] **Step 2: entrada + avisos.** Implementar la entrada y los tres avisos (relleno externo,
  presión externa, reducir presión de línea al instalar 206-3.9(b) — este último se cruza con
  el rango 50-80 % de la Fase 7). Citar cada bloque.
- [x] **Step 3: PASS + commit** — `Art. 206 Paso 5: presion externa, cavidades y bulging (206-3.6/7/9/10)`.

---

## Fase 6 — Paso 6: fatiga por operación cíclica y dilatación diferencial (206-2.4/3.8/3.11)

**Files:** Modify `build_collar_art206()` (una entrada condicional + avisos; la nota 206-3.11
ya existe). Modify tests.

**Fuente (resources/):** 206-2.4 (ciclos frecuentes → evaluación de fatiga por 206-3.8; Type B
con gradientes térmicos through-wall → fatiga de los filetes) [bloque 15]; 206-3.8 (todo Type B
se evalúa para fatiga; si se requiere, por VIII-2 / API 579-1 / equivalente) [bloque 52];
206-3.11 (dilatación térmica diferencial tubo/manga, ambos tipos) [bloque 60].

**Interfaces:** Entrada `¿Servicio cíclico (presión/térmico)?` que activa el aviso de
evaluación de fatiga; la nota 206-3.11 (ya existe, `nota206b`) se conserva y se traza.

- [x] **Step 1: test (falla).** `test_paso6_fatiga`: entrada `¿Servicio ciclico?`
  (`dv_list '"Si,No"'`); aviso `IF(ciclico="Si","206-2.4/3.8: requiere evaluacion de fatiga
  (VIII-2 / API 579-1/ASME FFS-1)","")`; nota 206-3.11 presente. → FAIL.
- [x] **Step 2: entrada + aviso** de fatiga; conservar/anclar la nota 206-3.11. Citar bloques.
- [x] **Step 3: PASS + commit** — `Art. 206 Paso 6: fatiga (206-2.4/3.8) y dilatacion diferencial (206-3.11)`.

---

## Fase 7 — Paso 7: fabricación y soldadura en servicio (206-4)

**Files:** Modify `build_collar_art206()` (Sección de fabricación: avisos + rango de presión).
Modify tests.

**Fuente (resources/):** 206-4.1 (limpieza a metal blanco; relleno en indentaciones; fit-up
ceñido; luz ≤ 2,5 mm — ya en Fase 4) [bloque 63]; 206-4.2 (el relleno no debe extruir a la
soldadura) [bloque 67]; 206-4.3 (fuga → aislar + **purga N₂** en fluidos inflamables) [bloque
69]; 206-4.4 (costuras longitudinales a tope penetración completa + venteo si hay filetes de
cierre; bajo hidrógeno — ya existe `nota206`) [bloques 71,80]; 206-4.5 (**reducir presión a
50-80 %**; API RP 2201; burn-through) [bloque 82]; 206-4.6 (Art. 210: (a) H₂ en ZAC,
(b) ZAC dura, (c) burn-through) [bloques 84-87].

**Interfaces:** Añade la sección de fabricación con los avisos trazados y el **rango de presión
de instalación** calculado `[0,50·P_op ; 0,80·P_op]`. La nota 206-4.4 (ya existe) se conserva.

- [x] **Step 1: test (falla).** `test_paso7_fabricacion`: celdas `P_instal,min = 0.5*$Pop` y
  `P_instal,max = 0.8*$Pop` (206-4.5); aviso de purga N₂ `IF(fuga="Si","206-4.3: purgar el
  anular con N2/gas inerte en fluidos inflamables","")`; avisos fijos de metal blanco (206-4.1),
  relleno (206-4.2) y soldadura en servicio Art. 210 (206-4.6). → FAIL.
- [x] **Step 2: rango de presión.** `P_instal,min = 0.5*$D_Pop`, `P_instal,max = 0.8*$D_Pop`
  (usar la presión de operación en las unidades de la hoja; unidades visibles). Citar 206-4.5 +
  API RP 2201.
- [x] **Step 3: avisos de fabricación.** Purga N₂ (206-4.3, cruza con Paso 1), metal blanco
  (206-4.1), relleno sin extruir (206-4.2), soldadura en servicio Art. 210 con los tres riesgos
  (206-4.6). Conservar `nota206` (206-4.4). Textos cortos (evitar `_LONGTEXT`).
- [x] **Step 4: PASS + commit** — `Art. 206 Paso 7: fabricacion, reduccion de presion 50-80% y Art. 210`.

---

## Fase 8 — Paso 8: examen (NDE) y prueba de hermeticidad (206-5, 206-6)

> **Decisión 3: mínima y fiel al 206.** Selector + notas; **sin** ecuaciones de energía
> almacenada (el texto del 206 no las publica — Regla nº 1). No se importa el motor neumático
> del App. 501 del 212.

**Files:** Modify `build_collar_art206()` (nueva sub-sección de examen y prueba). Modify tests.

**Fuente (resources/):** 206-5.1 (VT de todos los fit-ups y soldaduras) [bloque 92]; 206-5.2
(Type A: VT de raíz + PT/MT/UT de longitudinales) [bloque 94]; 206-5.3 (Type B: UT del base del
portador; primer/último pase MT o PT; **NDE ≥ 24 h**, ≥ 48 h en servicio con alta probabilidad
de H₂) [bloque 96]; 206-6 ((a) presurizar el anular, presión tal que el tubo interno **no
colapse**; (b) prueba sensible de fugas B31.3 345.8; Art. 501 guía adicional) [bloques 102-104].

**Interfaces:** Notas de NDE por tipo + selector de prueba y su aviso. El selector no gobierna
ningún cálculo (no hay energía): solo materializa el requisito y la advertencia de presión.

- [x] **Step 1: test (falla).** `test_paso8_nde_prueba`: selector `Tipo de prueba de
  hermeticidad` (`dv_list '"Prueba del anular,Prueba sensible de fugas,No requerida"'`); aviso
  `IF(tipo="Prueba del anular","206-6(a): presion de prueba tal que el tubo interno NO colapse;
  Art. 501 guia adicional","")`; nota de NDE diferido `IF($D$22=TIPO_B,"206-5.3: NDE de
  circunferenciales >=24 h (>=48 h si servicio con H2)","")`. → FAIL.
- [x] **Step 2: notas de NDE.** VT (206-5.1); Type A PT/MT/UT longitudinal (206-5.2); Type B UT
  del portador + primer/último pase MT/PT + NDE diferido (206-5.3). Avisos condicionales por
  `D22`. Textos cortos.
- [x] **Step 3: selector de prueba.** `dv_list` con las dos vías del 206-6 + «No requerida»; el
  aviso cita la limitación de presión (que el tubo interno no colapse) y remite al Art. 501.
  **Sin** celdas de E/TNT/distancia (decisión 3).
- [x] **Step 4: PASS + commit** — `Art. 206 Paso 8: NDE por tipo y prueba de hermeticidad 206-6 (minima)`.

---

## Fase 9 — Notación de símbolos con subíndice (decisión 4)

**Files:** Modify `build_db_materiales.py` (helper `sym()` — crearlo si el plan del 212 aún no
lo introdujo; si ya existe, reutilizarlo). Aplicarlo en los rótulos de símbolo del 206
(`T_s`, `T_p`, `L_s`, `L_defecto`, `S(T)`, `E_j`, `w`, `w_máx`, `C.A.`, `P_op`, etc.). Modify
`TestSistemaVisual` si hace falta que su recorrido lea los *runs* de `CellRichText`.

**Interfaces:** `sym(base, sub)` → `CellRichText` con el subíndice en
`InlineFont(vertAlign="subscript")`, fuente `MONO`. No cambia ninguna fórmula (los símbolos son
rótulos, no referencias).

- [ ] **Step 1: test (falla).** `test_simbolos_sin_guion_bajo_206`: recorre los rótulos de
  símbolo del 206 y afirma que **ninguno** contiene `"_"` en su texto plano y que los que llevan
  subíndice son `CellRichText`. → FAIL.
- [ ] **Step 2: helper `sym()`** (crear/compartir) con `CellRichText`/`TextBlock`/
  `InlineFont(vertAlign="subscript")`.
- [ ] **Step 3: aplicar** a los símbolos de todas las secciones del 206.
- [ ] **Step 4: comprobar `TestSistemaVisual`** (que lea los *runs* del rich text; ampliarlo si
  solo mira `cell.font`). PASS + commit — `Art. 206 Paso 9: simbolos con subindice real`.

---

## Fase 10 — Anclar, recalcular, verificar y entregar

**Files:** Modify `test_dashboard.py` (`TestBuildCollarArt206` completo), `verificar.py §6f`;
Modify `CLAUDE.md` (estado del motor 206); Modify este plan (estado final).

- [x] **Step 1: build completo.** Sin abortos.
- [x] **Step 2: los gates sin recálculo.** `pytest` = **252 passed** (incl. `TestBuildCollarArt206`,
  8 anclas por paso). `TestSistemaVisual`, `TestSincroniaPythonVba` y la visibilidad/navegación
  del collar verdes. `verificar.py §1-5` sin discrepancias ni matriz dinámica.
- [x] **Step 3: entregar para F9 en Excel.** `.xlsm` entregado por `SendUserFile` (2026-09-11).
  Nota: con **C.A.=0 (default) el `t_req` no cambia de valor** y el cateto solo aplica a Type B
  (el seed es Type A → w «No aplica»); los cambios de valor los ejerce el ingeniero al teclear
  C.A. o poner Type B.
- [x] **Step 4: `verificar.py §6f`.** Añadida: recalcula en Excel real el cateto `w`
  (Type B forzado en el qa), la luz `G≤2,5` y el tipo recomendado. w=9.5, luz CUMPLE, tipo Type A.
  El `t_req`+C.A. lo fijan las anclas de cadena (`TestBuildCollarArt206`) + §1-5 book-wide.
- [x] **Step 5: `verificar.py` completo en Windows** (§1-11 + §6f) → **0 fallos**.
- [~] **Step 6: documentar.** Nota de estado añadida al `CLAUDE.md` del proyecto. **Pendiente de
  cierre: F9 del ingeniero + Fase 9 (subíndices, agrupada con el 212).** El plan queda en
  `outputs/plans/` hasta entonces.

---

## Verificación del plan contra el Art. 206 (evidencia de aprobación)

Trazado de cada cláusula/regla del Art. 206 a la tarea que la implementa, cruzado contra
`resources/` (Regla nº 1). El flujo del ingeniero es fiel al código; este plan cierra la brecha
entre el flujo y el motor.

| Cláusula / regla (resources/) | Requisito | Tarea | Estado tras el plan |
|---|---|---|---|
| 206-1.1.1 / 1.1.2 [5],[6] | Type A (no presión) / Type B (presión, fuga/axial) | Fase 1 | Selección guiada de tipo |
| 206-2.3 [13] | fuga activa Type B → aislar antes de soldar | Fase 1/7 | Aviso + purga N₂ |
| 206-2.4 [15] | ciclos → fatiga por 206-3.8 | Fase 6 | Aviso condicional |
| 206-2.5 [17] | Type A no apto para defectos circunferenciales | Fase 1 (ya existía F63) | Conservado + entra en la recomendación |
| 206-2.6 [19] | corrosión bajo-manga Type A → sellante/recubrimiento | Fase 1 | Aviso |
| 206-2.7 [21] | costura previa → esmerilar + RT/UT o bulge | Fase 1 | Aviso |
| 206-3.1 [31] | Type A ≥ ⅔·T_p | Fase 2 (ya existía D54) | Conservado, sin C.A. |
| 206-3.2 [35] | Type B ≥ t; E=0,80 / 1,00 si UT; axial si girth | Fase 2 (ya existía D38/D53) | Conservado |
| 206-3.3 [38] | `t_req` por código + **C.A. por el diseño** | Fase 2 | **C.A. añadido** (decisión 2) |
| 206-3.4 [40] | ≥ 100 mm y defecto + 2·50 mm | Fase 3 (ya existía D61) | Explícito y trazado |
| 206-3.5(a)/(b) [44],[45] | rama 1,4× | Fase 4 (ya existía F64) | Conservado |
| 206-3.5 Figs. [72],[77] imagen | **`w = T_s + G` / `w_máx = 1,4·T_p + G`** | **Fase 0 + Fase 4** | Recuperado de la imagen de `resources/`, fijado en texto, luego calculado |
| 206-3.6/3.7/3.9/3.10 [48],[50],[53-58] | presión externa; relleno epoxi; reducir presión | Fase 5 | Entrada condicional + avisos |
| 206-3.8 [52] | Type B → evaluar fatiga (VIII-2/API 579) | Fase 6 | Aviso condicional |
| 206-3.11 [60] | dilatación térmica diferencial | Fase 6 (ya existía nota206b) | Conservado |
| 206-4.1 [63] | metal blanco; **luz `G ≤ 2,5 mm`** | Fase 4/7 | **Verificación de luz añadida** |
| 206-4.2 [67] | relleno no debe extruir a la soldadura | Fase 7 | Aviso |
| 206-4.3 [69] | fuga → aislar + purga N₂ | Fase 7 | Aviso |
| 206-4.4 [71],[80] | longitudinales a tope + venteo; bajo H | Fase 7 (ya existía nota206) | Conservado |
| 206-4.5 [82] | reducir presión **50-80 %**; API RP 2201 | Fase 7 | Rango calculado |
| 206-4.6 [84-87] | Art. 210: H₂, ZAC dura, burn-through | Fase 7 | Aviso |
| 206-5.1/5.2/5.3 [92],[94],[96] | VT; PT/MT/UT; UT portador; **NDE ≥ 24/48 h** | Fase 8 | Notas por tipo |
| 206-6 [102-104] | prueba del anular / sensible de fugas; presión sin colapso; Art. 501 | Fase 8 | Selector + notas (sin energía, decisión 3) |

**Hueco de `resources/` que el plan cierra antes de codificar (Regla nº 1):** el cateto del
filete (`w = T_s + G` / `w_máx = 1,4·T_p + G`), **imagen presente** en `resources/`
(Figs. 206-3.5-1/2) → recuperado leyendo los PNG y fijado en texto en la Fase 0. **No hay
ningún vacío real** para el Art. 206. En ningún caso se pide un folio externo: `resources/` es
la fuente.

---

## Verificación — qué se puede aquí y qué no

- **Aquí (sin recálculo):** `pytest` (tres suites), build sin abortos, `verificar.py §1-5`
  (book-wide), y las anclas por paso de `TestBuildCollarArt206` (igualdad de cadena de las
  fórmulas nuevas contra lo que este plan especifica, trazado a `resources/`). Para los valores,
  extraer `w`, `t_req`+C.A., tipo recomendado con openpyxl y cotejarlos a mano contra el código.
- **Sólo en Excel de Windows (ingeniero):** `verificar.py §6f` (recálculo real de `w`, tope `G`,
  `t_req`+C.A. y tipo recomendado del caso semilla), F9 del caso precargado (Type A y Type B), y
  la revisión visual de la hoja exportada a PDF/PNG (openpyxl miente sobre bordes de rango
  fusionado). Los valores esperados del caso semilla se re-baselinan **después** de esa
  validación.

## Riesgos

- **`T_p` nominal vs. «required minimum wall thickness».** El texto 206-3.5(a) compara con el
  espesor **nominal** del portador; la **leyenda de las figuras** rotula `T_p = carrier pipe
  required minimum wall thickness`. El motor hoy usa D21 (t nominal de `Datos_Ref`) para la
  rama del 1,4×; el cateto de la Fase 4 usa el mismo `T_p`. **Confirmar con el ingeniero** si el
  cateto debe usar el nominal (D21, coherente con la rama) o el mínimo requerido; por defecto,
  el nominal (lo que compara el texto de la rama). Declararlo en la nota de la celda.
- **La referencia de la luz `G` (D31).** Hoy D31 = «Luz radial sleeve-portador», y ya alimenta
  la geometría (OD/RI del sleeve, D45/D46). Reutilizarla como `G` del cateto y del tope 2,5 mm
  es correcto (es la misma magnitud física), pero **verificar** que el default (1,5 mm) y su uso
  en D45/D46 siguen coherentes tras añadir la verificación del tope.
- **Desplazamiento de direcciones al insertar filas.** `construir_seccion7_material()` se llama
  con `modo_cell="$D$11"`, `temp_fuente_cell="$D$25"` y `destino_st="D37…"`. Insertar filas
  **por encima** de D11/D25/D37 rompería esas referencias. Mitigación: **añadir** las entradas
  nuevas dentro de las bandas existentes por debajo de esas anclas, o **actualizar los
  argumentos de la llamada** si se reubican. Anclar por cadena en el test para detectarlo.
- **El caso semilla cambia de valor.** El C.A. en el `t_req` mueve el gobernante Type B y quizá
  el dictamen. `verificar.py §6f` y los valores esperados del caso semilla se re-baselinan tras
  la validación del ingeniero — **no** antes. No cerrar el plan hasta esa confirmación.
- **`CellRichText` y `TestSistemaVisual`.** El recorrido de la prueba visual debe leer los
  *runs* del rich text; si sólo mira `cell.font`, ampliarlo, o los subíndices podrían escapar al
  control de paleta.
- **`sym()` compartido con el 212.** Si el plan del 212 aún no se ejecutó, `sym()` no existe;
  esta Fase 9 lo crea. Si se ejecutan ambos planes, evitar la doble definición (crear una vez,
  reutilizar).
- **Textos de aviso como fórmula.** Los avisos condicionales `IF(cond,"texto…","")` no deben
  superar 255 caracteres (→ `_xlfn._LONGTEXT`, `#NAME?` fuera de Excel 365). Mantener textos
  cortos; si uno crece, materializarlo como celda de texto con `IF` corto y dejar que
  `normalizar_textos_como_formula` lo cubra.

## Self-review (hecho)

- **Cobertura del flujo:** los 8 pasos tienen fase (tabla de trazado arriba). Los huecos del
  cotejo (cateto de filete, luz `G≤2,5`, C.A. en `t_req`, selección de tipo, y los avisos de
  206-2/3/4/5/6) tienen tarea.
- **Regla nº 1:** ningún número nuevo sale de memoria; el único que es imagen (el cateto) se
  fija en la Fase 0 antes de codificarse. Sin vacío real. Sin importar el App. 501 (el 206 no lo
  publica).
- **Consistencia:** `build_collar_art206(wb, b313, iid1a, iidb, fac_info, rangos)` mantiene su
  firma; la Sección de resolución de material (`construir_seccion7_material`) **no cambia** de
  comportamiento (parity con el 212); el tipo (D22) sigue siendo la decisión del ingeniero,
  ahora con recomendación; `w` (Fase 4) usa `T_s` (D29), `T_p` (D21) y `G` (D31).
- **Sin placeholders:** cada fórmula nueva se da en forma matemática con su bloque/figura de
  `resources/`; la cadena Excel exacta la produce la implementación y la fija
  `TestBuildCollarArt206` + `verificar.py §6f`.
