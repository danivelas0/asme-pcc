# Motor de búsqueda de propiedades físicas — ASME B31.3, Apéndice C

**Documento:** PLAN-BUSC-APXC-001 · Rev. 0 · 2026-09-07
**Entregable objetivo:** `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm` (38 hojas, una visible)
**Fuente única de verdad:** `resources/ASME B31/ASME B31.3/APPEX/appendix_c/`

---

## ✅ EJECUTADO — 2026-09-07

Fases 0 a 5 completas. `verificar.py` devuelve **0 fallos** en sus nueve secciones.

| Fase | Estado | Resultado |
|---|---|---|
| 0 · Cerrar el vacío normativo | ✅ (ya lo estaba) | `completar_apendice_c.py` |
| 1 · `DB_B31_C` / `DB_B31_CC` | ✅ | 201 filas por edición; las cuatro `DB_C_*` desaparecen |
| 2 · Motor `Buscar_Prop_B31_3` | ✅ | cascada de 5 niveles, conmutador SI/US, rama PUNTO, bloqueo en los dos extremos |
| 3 · Separar Apéndice B y navegación | ✅ | `Buscar_Prop_IID`, `Buscar_NoMetalicos` solo Ap. B |
| 4 · Verificación | ✅ | §1, §2/§4, §3b (4 726 valores, 0 discrepancias), §6b (10 casos recalculados en Excel) |
| 5 · `CLAUDE.md` e `Instrucciones` | ✅ | regla 10 ya estaba; se añaden §3b de Instrucciones y la sección del Apéndice C |

**Dos desviaciones del plan, ambas a mejor y documentadas en el código:**

1. **27 columnas en vez de 25.** La definición impresa del coeficiente y el texto
   impreso del factor de escala van en columna propia en lugar de componerse dentro
   de una fórmula: la ficha del motor los cita verbatim y siguen siendo auditables
   desde la hoja. El contrato que importa —las **8 primeras columnas idénticas a
   `STRESS_COLS`** y `n_pts` como última de identificación— se mantiene entero, y es
   lo que permite que `build_listas` y las §2/§4 de `verificar.py` no cambien ni una
   línea.
2. **Se corrigió un fallo real que el plan no preveía:** `INDEX` sobre una celda
   vacía devuelve **0**, no cadena vacía. Sin envolverlo, una fila de tipo PUNTO
   —que no tiene banda tabulada— daba `T1 = 0` y `ISNUMBER(valor único) = VERDADERO`
   sobre un cero inventado, y el motor mostraba `0` donde el código publica un
   intervalo de texto. Corregido en el motor y en el banco de pruebas.

Los rótulos de estado (`EST_BAJO`, `EST_ALTO`, `EST_PUNTO`…) y las funciones que
generan las fórmulas (`formula_estado_apxc`, `formula_valor_apxc`) viven en
`build_db_materiales.py` y las emiten **tanto el motor como `verificar.py` §6b**:
la prueba ejerce el original, no una copia.

---

## Contexto

El Apéndice C del B31.3 es la tabla de **propiedades físicas de materiales de tubería**:
dilatación térmica y módulo de elasticidad, para metales y para no metálicos, en SI y en
U.S. Customary. Hoy esos datos están repartidos y mezclados con material de otros códigos:

| Tabla | Dónde vive hoy | Desde dónde se consulta |
|---|---|---|
| C-1 / C-1C (dilatación, metales) | `DB_C_dilatacion` / `DB_C_dilatacionC` | bloque 2 de `Buscar_Propiedades`, junto a la II-D |
| C-3 / C-3C (módulo, metales) | `DB_C_modulo` / `DB_C_moduloC` | bloque 3 de `Buscar_Propiedades`, junto a la II-D |
| C-2 (dilatación, no metálicos) | enterrada en `DB_NoMetalicos`, como pares campo/valor | `Buscar_NoMetalicos`, junto al Apéndice B |
| C-4 (módulo corto plazo, no metálicos) | ídem | ídem |

Tres consecuencias: el ingeniero que busca «el módulo E del A106 a 350 °C» tiene que saber
de antemano en cuál de dos motores mirar; el Apéndice B —que es otra cosa, *esfuerzos de
diseño hidrostático y presión admisible*— comparte motor con propiedades físicas; y las
tablas C-2 y C-4 no tienen ni curva ni ficha, solo pares de texto.

**Objetivo:** un motor propio, `Buscar_Prop_B31_3`, que **categoriza primero por propiedad**
—ese es el nivel 0 de su cascada— y que cubre el Apéndice C entero, en SI y en US.
El Apéndice B se queda solo en su motor. Un dato, un motor.

### Decisiones ya tomadas

1. **Hoja nueva dentro del libro Rev. 3** (que pasa a Rev. 4), no un `.xlsm` aparte:
   reusa builder, Dashboard, capa VBA y `verificar.py`, y evita una segunda fuente de
   verdad de la navegación.
2. **Conmutador SI ↔ US en el propio motor**, y se eleva a regla del proyecto para todos
   (fase 5).
3. **El Apéndice C sale entero de los otros dos motores.** `Buscar_Propiedades` se renombra
   `Buscar_Prop_IID` y queda solo con II-D; `Buscar_NoMetalicos` queda solo con Apéndice B.
4. **El vacío normativo se cierra extrayendo del PDF** (fase 0, bloqueante).

---

## Fase 0 — Cerrar el vacío normativo — ✅ **EJECUTADA (2026-09-07)**

> **Estado:** hecha. `scripts/completar_apendice_c.py` corrido sobre las 6 tablas;
> `test_build_db.py::TestApendiceCCompletado` en verde (9 pruebas; 62 en total).
> Auditoría cruzada OCR ↔ `resources/`: **todos los números de C-2, C-3, C-3C y C-4
> coinciden** y **0 filas con valor o nombre alterado** en las 6 tablas.
> Lo que sigue —Fases 1 a 5— está pendiente.


El usuario aportó dos insumos:

- `APENDICE C - ASME B31.3.pdf` — 24 páginas, folios impresos **407–428**. **Es un escaneo:
  no tiene capa de texto** (`pdfplumber` devuelve 0 líneas en las 24 páginas), así que un
  extractor tipo `extraer_notas_ii_d.py` no puede leerlo directamente.
- `datalab-output-APENDICE C - ASME B31.3.pdf.json` — **la salida OCR estructurada del mismo
  PDF** (24 páginas, 0 fallidas). Esto sí es máquina-legible y **hace viable un extractor
  reproducible**: las notas vienen como tablas de 3 columnas, los encabezados de grupo como
  filas `colspan` y las definiciones de coeficientes como texto LaTeX.

Cotejo ya hecho, página impresa a página impresa: **los valores numéricos del JSON de
`resources/` son correctos** (verificadas las cabeceras de las 6 tablas y filas completas de
cada una, incluida la fila `Carbon steels ≤0.30% C` de C-3 entera). Lo que está roto es la
**capa de metadatos**, y **cinco filas llevan un grupo falso**, no solo ausente.

### C-1 / C-1C — definición de los coeficientes (folios 408 y 414)

Lo que faltaba, ahora leído del código:

| | Table C-1 (SI) | Table C-1C (US) |
|---|---|---|
| `A` | Mean Coefficient of Thermal Expansion, **10⁻⁶ mm/mm/°C** | Mean Coefficient of Thermal Expansion, **10⁻⁶ in./in./°F** |
| `B` | Linear Thermal Expansion, **mm/m** | Linear Thermal Expansion, **in./100 ft** |
| Referencia | *in Going From **20 °C** to Indicated Temperature* [Note (1)] | *in Going From **70 °F** to Indicated Temperature* [Note (1)] |

Se añaden como `coefficient_defs`. Con esto el motor ya puede rotular KPI y ejes, y se
confirma lo que hasta ahora solo se había deducido por aritmética.

### C-1 / C-1C — miembros de las Notas (2)…(6) (folios 412 y 418)

Las notas se imprimen **a tres columnas**; la extracción las leyó por filas y las dejó
intercaladas en un único string, con un bloque de composiciones huérfano al final.
Contenido real:

| Nota | Grupo | Miembros |
|---|---|---:|
| (2) | Group 1 alloys | 51 |
| (3) | Group 2 alloys | 6 |
| (4) | Group 3 alloys | 13 |
| (5) | Group 4 alloys | 14 |
| (6) | Aluminum alloys (por UNS) | 18 |

**Las dos ediciones imprimen las mismas notas, con la misma numeración y los mismos
miembros** (folio 412 frente a 418, comparados uno a uno). A diferencia de TM-1 de la II-D,
**aquí no hay errata de numeración**. Se añaden como `note_members` —mismo campo y misma
semántica que en la II-D, así que builder y `verificar.py` ya saben leerlo—. El string
`notes` original **se conserva**: `note_members` es la forma estructurada, no un reemplazo.

Esto es lo que hace usable la tabla: sin ello no se puede decir si un 1¼Cr–½Mo es Grupo 1.

### C-2 — el mapa de grupos está **mal**, no solo incompleto (folios 419–420)

En el impreso, un material de subgrupo va **indentado** y uno de nivel superior va a ras.
La extracción perdió la indentación y arrastró el subgrupo anterior:

| Corrección | Filas | Qué dice el JSON hoy | Qué imprime el código |
|---|---:|---|---|
| **Subgrupo falso → se elimina** | 3 | `Polybutylene PB 2110` y `Polyether, chlorinated` como `Chlorinated poly(vinyl chloride)`; `Polyphenylene POP 2125` como `Cross-linked polyethylene` | los tres van **a ras**, sin subgrupo, dentro de `Thermoplastics` |
| **Grupo ausente → se completa** | 9 | sin `material_group`: PVC2116, PVC2120, PVDF, PVDC, PTFE, FEP, PFA ×3 | `Thermoplastics`; el encabezado no se reimprime en la página de continuación. PVC2116 y PVC2120 van además indentados bajo `Poly(vinyl chloride)` |
| **Grupo truncado → se completa** | 5 | `"and Reinforced Plastic Mortars"` | `Reinforced Thermosetting Resins and Reinforced Plastic Mortars` (título a dos líneas) |

Reparto correcto de las 44 filas: **`Thermoplastics` 38** —subgrupos
`Acrylonitrile-butadiene-styrene` 4, `Cellulose acetate butyrate` 2,
`Chlorinated poly(vinyl chloride)` **1** (solo CPVC 4120), `Polyethylene` 7,
`Cross-linked polyethylene` **2** (solo PEX0006 y PEX0008), `Polypropylene` 5,
`Poly(vinyl chloride)` **6** (incluye PVC2116 y PVC2120), y **11 a ras**—;
**`Reinforced Thermosetting Resins and Reinforced Plastic Mortars` 5**;
**`Other Nonmetallic Materials` 1** (Borosilicate glass, 1,8 / 3,25).

Factor confirmado: *Mean Coefficients (**Divide** Table Values by 10⁶)*.

> `Acetal AP2012 = 2` (in./in./°F) `/ 3.6` (mm/mm/°C) —que parecía un valor truncado— es
> **exactamente lo que imprime el código**. Se conserva (regla 9). No era un defecto.

### C-3 / C-3C — el mismo arrastre (folio 421)

`Austenitic stainless steels:` cubre **solo los seis `Type …`** (304, 310, 316, 321, 347,
309). `Straight chromium stainless steels (12Cr, 17Cr, 27Cr)` y **`Gray iron`** van a ras:
el `material_subgroup: "Austenitic stainless steels:"` que hoy llevan es **falso**.
`Chromium steels:` cubre ½Cr–2Cr, 2¼Cr–3Cr y 5Cr–9Cr.

Los sufijos `(Cont'd)` de `Nickel Alloys`, `Aluminum and Aluminum Alloys` y
`Copper and Copper Alloys` son artefacto de corte de página —y cortan en filas distintas en
SI y en US—: se normalizan a un solo grupo.

Factores confirmados con su superíndice: **`Multiply Tabulated Values by 10³`** (MPa, C-3) y
**`10⁶`** (psi, C-3C). Verificada además la fila `Carbon steels ≤0.30% C` completa
(220 · 216 · 212 · 209 · 202 … 107, nulos desde 700 °C): coincide con el JSON.

### C-4 — correcto (folio 428)

Grupos `Thermoplastics [Note (1)]` (25 filas), `Thermosetting Resins, Axially Reinforced`
(4) y `Other` (1). Se añaden `scale_factor: 1` explícito y la temperatura de referencia
impresa (**73,4 °F / 23 °C**), que hoy solo vive dentro del nombre de la columna.

### Cómo se aplica

Primero, **el OCR entra en `resources/`**: se copia el JSON de Datalab a
`resources/ASME B31/ASME B31.3/APPEX/appendix_c/ocr_apendice_c_datalab.json` (y el PDF a
`resources/.../appendix_c/pdf/`), para que la corrección trace a una fuente versionada y no
a un archivo suelto de `Downloads`.

`scripts/completar_apendice_c.py --resources … --ocr ocr_apendice_c_datalab.json` — script
**idempotente**, que deriva del OCR casi todo mecánicamente:

| Qué | De dónde sale, en el OCR |
|---|---|
| `coefficient_defs` de C-1 / C-1C | el bloque `Table` de las páginas 2 y 8 trae `A = \text{Mean Coefficient of Thermal Expansion, } 10^{-6}\ \text{mm/mm/°C}` y su gemelo US; se normaliza el LaTeX |
| `note_members` (2)…(6) | páginas 6 y 12: cada nota es un `ListGroup` con su rótulo seguido de un `Table` de **3 columnas** — se lee **por columnas**, que es justo lo que la extracción original hizo mal |
| Jerarquía de grupos | dentro del `<table>`, un encabezado es una fila `colspan`; **grupo** si va en `<b>`, **subgrupo** si no. Así `Thermoplastics` se distingue de `Acrylonitrile-butadiene-styrene` sin heurística |
| Título truncado de C-2 | el OCR ya lo trae entero: `Reinforced Thermosetting Resins and Reinforced Plastic Mortars` |
| Las 9 filas sin grupo | se resuelven arrastrando el último grupo **a través del corte de página**, que es lo que la extracción original no hizo |
| `scale_factor` | del encabezado con su superíndice: `by 10<sup>6</sup>`, `by 10<sup>3</sup>` |

Lo único que el OCR **no** puede dar es la **indentación**: el HTML no la conserva, así que
`Polybutylene PB 2110`, `Polyether, chlorinated`, `Polyphenylene POP 2125` (C-2) y
`Straight chromium stainless steels` y `Gray iron` (C-3/C-3C) siguen apareciendo pegados al
subgrupo anterior. Son **5 excepciones**, y van en el script como **tabla literal con su
folio impreso y la verificación visual que las respalda** — no como regla. Cualquier fila
que no esté en esa lista se resuelve sola.

El script comprueba que el JSON de partida es el esperado (**aborta si ya cambió**) y deja
constancia en `extraction_amendments` dentro de cada archivo: qué se añadió, qué se
corrigió, de qué folio, con qué fuente y con qué fecha.

**El builder aborta** si faltan `coefficient_defs`, `note_members` o el mapa de grupos
corregido — igual que aborta hoy sin el `note_members` de la II-D.

**Prueba:** `test_build_db.py::TestApendiceCCompletado` — las dos ediciones traen
`coefficient_defs` con unidad; las Notas (2)–(6) tienen 51/6/13/14/18 miembros y son
idénticas entre ediciones; ninguna de las tres filas corregidas de C-2 conserva subgrupo;
las 44 filas de C-2 reparten 38/5/1; `Gray iron` no tiene subgrupo; los exponentes son
3, 6 y −6; y **cada valor numérico del OCR coincide con el que ya está en `resources/`**
(auditoría cruzada de las 4 tablas — si el OCR y la extracción original discrepan, el script
para y lo reporta en vez de elegir).

---

## Fase 1 — Bases unificadas `DB_B31_C` / `DB_B31_CC`

Refactor de `build_appendix_c` (`build_db_materiales.py:659-698`) a
`build_apendice_c(res, wb, edicion, headers)`. Absorbe las cuatro tablas y **sustituye** a
las cuatro hojas `DB_C_*`, que dejan de existir.

**201 filas por hoja**, en las dos ediciones:

| Propiedad (nivel 0) | Tabla SI | Tabla US | Filas | Naturaleza |
|---|---|---|---:|---|
| DILATACION TERMICA - METALES | C-1 | C-1C | 52 | curva vs. T (26 materiales × coef. A y B) |
| DILATACION TERMICA - NO METALICOS | C-2 | C-2 (col. `in./in.,°F`) | 44 | valor único + rango de validez |
| MODULO DE ELASTICIDAD - METALES | C-3 | C-3C | 75 | curva vs. T |
| MODULO ELASTICIDAD CORTO PLAZO - NO METALICOS | C-4 | C-4 (col. `E, ksi`) | 30 | valor único a 23 °C / 73,4 °F |

### Contrato de columnas — `APXC_COLS`

**Las ocho primeras posiciones son idénticas a `STRESS_COLS`.** No es cosmético: con eso,
`build_listas` (`:1592`), la §2 de unicidad y la §4 de contigüidad de `verificar.py`
funcionan sobre la base nueva **sin tocar una línea**.

```
 1 material_id      5 k2            9 Propiedad          13 Variante/coeficiente
 2 Tabla            6 k3           10 Grupo impreso      14 Tipo de dato
 3 k0               7 k4           11 Subgrupo impreso   15 Unidad impresa
 4 k1               8 clave_bi     12 Material           16 Factor de escala
17 Valor unico   18 Valor unico (texto)   19 Rango de validez
20 T primera tabulada   21 T ultima tabulada   22 Notas   23 Observacion
24 Linea   25 n_pts
```
`N_IDENT_C = 25`, y a continuación la **banda impresa** (unión de las rejillas de
temperatura) y la **banda compacta** de `append_packed` (`:191`), sin cambios.

### Cascada de 5 niveles

| Nivel | Contenido | Cuando el código no lo imprime |
|---|---|---|
| `k0` | **Propiedad** (4 opciones) | — |
| `k1` | Grupo impreso, de `group_headers` | `(sin grupo impreso)` |
| `k2` | Subgrupo impreso | `(sin subgrupo impreso)` |
| `k3` | Material, tal como lo nombra la tabla | — |
| `k4` | Variante / coeficiente | `(unico)` |

`k4` vale `A - coeficiente medio` / `B - expansion total` en C-1; el **rango de validez** en
las tres filas de `Poly(perfluoroalkoxy alkane)` de C-2, que repiten
`material_description` con rangos distintos (70–212 °F, 212–300 °F, 300–408 °F) y por eso
no son clave única; `(unico)` en el resto.

Orden de escritura: `k0 → k1 → k2 → k3 → k4`, con el mismo `sort` que las claves, para que
cada bloque sea contiguo (regla 5) — precondición del `MATCH`+`COUNTIF` de las listas
dependientes.

### Reglas de construcción, con su porqué

1. **Valores tal como están impresos** (regla 9): se carga `202`, no `202 000`. El factor
   va en su columna y lo aplica el motor a la vista.
2. **Reutilizar lo que ya resuelve los artefactos**, no reescribirlo:
   - `temp_to_number` (`db_lib`) ya normaliza el **U+2212** de las claves de temperatura
     (`"−200"` no es `"-200"`) y la coma de millares (`"1,000"`). Cubierto por
     `TestTemperaturas`.
   - `_dedup_frase` (`:644`) ya colapsa `"Gray iron Gray iron"` → `"Gray iron"`: 21 de los
     26 materiales de C-1 traen el nombre duplicado.
   - `txt` / `comp_key` (`db_lib`) ya pliegan la raya U+2013 y la barra U+2215 (`1∕2Mo`).
3. **La clave de coeficiente difiere entre ediciones**: C-1 usa `coefficient`, C-1C usa
   **`coeffi_cient`** (el encabezado venía partido como `"Coeffi- cient"`). El lector acepta
   las dos; un parser que asuma una sola falla en la mitad de los casos.
4. **Enlace SI ↔ US posicional.** `clave_bi = f"{prop_id}#{indice_de_fila_impresa}"`.
   No por nombre: los nombres divergen entre ediciones (`"Type 309. 23Cr–12Ni"` con punto
   en C-3 frente a coma en C-3C; `"25Cr-20Ni"` con guion ASCII en una y raya en la otra;
   el duplicado de nombre de C-1 no coincide fila a fila con C-1C).
   Contrapartida obligatoria: **el build aborta** si el número de filas por propiedad no
   coincide, o si un nombre normalizado difiere fuera de una lista blanca de artefactos
   documentados. Un enlace posicional sin esa aserción es una bomba silenciosa.
5. **Valores que el código publica como intervalo** (`"9–13"` en C-2, `"1,200–1,900"` en
   C-4) se cargan como **texto**, `Tipo de dato = PUNTO`, y el motor los muestra sin operar.
6. **`T primera` / `T ultima`** = primer y último punto **con valor** de la banda compacta.
   Sustituyen a la columna `Temp. máx.` que el Apéndice C no tiene, y son el límite de
   rango del motor.
7. **Columna `Observacion`**: declara fila a fila cada artefacto conservado o reparado y
   cita la página del PDF. Es lo que hace auditable la fase 0.
8. `record_meta` por cada una de las 6 tablas (archivo fuente, tabla, edición, unidad,
   filas), como el resto de bases.

### Cobertura que el motor debe declarar, no ocultar

Verificado sobre los JSON: **la edición métrica y la US no cubren siempre el mismo rango**.
No son redondeos de conversión, son puntos que una tabla imprime y la otra no.

| Material | SI | US |
|---|---|---|
| C70600, C97600, C71000, C71500 (cuproníqueles), C-3 | hasta **200 °C** | hasta **700 °F** (371 °C) |
| Aceros al carbono ≤0,30 % C, C-3 | hasta **650 °C** | hasta **1 100 °F** (593 °C) |
| N02200 / N02201, C-1 | hasta **825 °C** | hasta **1 400 °F** (760 °C) |
| N04400 / N04405, C-1 | hasta **800 °C** | hasta **1 500 °F** (816 °C) |

El aviso `G5:L5` del motor dice siempre qué tabla y qué edición está leyendo, y el estado
bloquea por el límite de **la edición activa**. Interpolar en SI y convertir **no** equivale
a leer la tabla US, y el motor no debe sugerir que sí.

Materiales que existen en una tabla y no en la otra (el motor lo dice, no lo inventa):
zirconio R60702/R60705 y catorce aleaciones de níquel tienen módulo en C-3 y **no** tienen
dilatación en C-1; la fundición dúctil tiene dilatación en C-1 y **no** módulo en C-3.
C-2 y C-4 tampoco listan el mismo inventario (44 frente a 30) ni lo nombran igual
(`"Glass–epoxy, centrifugally cast"` en C-2 frente a `"Epoxy–glass, centrifugally cast"`
en C-4 — términos invertidos).

---

## Fase 2 — El motor `Buscar_Prop_B31_3`

Funciones nuevas `build_buscador_prop_c` / `finish_buscador_prop_c`, junto a las existentes.

**Por qué no se reutiliza `build_buscador`:** esa pareja la comparten los cinco buscadores
de esfuerzos y sostiene 271 276 valores ya auditados. El Apéndice C necesita tres cosas que
ella no tiene —rama de dato puntual, límite de rango tomado de la propia banda en vez de
una columna `Temp. máx.`, y un conmutador que en dos tablas cambia de **columna** en lugar
de hoja—. Tocarla arriesga la regresión de todo lo verificado.

**Lo que sí se reutiliza, verbatim:** `new_sheet` (:100), `banda` (:1697), `_mrg` (:1705),
`campo` (:1714), `dv_list` (:165), `_nota` (:157), `append_packed` (:191) / `packed_refs`
(:210), `interp_value` (:228), el patrón de lista dependiente
`IF(COUNTIF(...)<i,"",INDEX(...,MATCH(...)+i-1))` de `:1861-1863`, y el bloque de gráfica de
`finish_buscador` (:2163-2205).

> `cascade_formula` (:174) está definida y **no la llama nadie**: es la versión con `OFFSET`
> de la Rev. 2, no portable como origen de validación. No usarla.

### Layout (12 columnas A..L, como el resto de motores)

```
┌────────────────────────────────────────────────────────────────────────────────┐
│ PROPIEDADES FISICAS DE MATERIALES DE TUBERIA — ASME B31.3-2024, APENDICE C      │ NAVY
│ Tablas C-1/C-1C · C-2 · C-3/C-3C · C-4 — valores tal como estan impresos        │
│ ◂ VOLVER AL DASHBOARD                                                          │ fila 3
├────────────────────────────────────────────────────────────────────────────────┤
│ 1 · QUE PROPIEDAD QUIERE CONSULTAR                                             │ banda BLUE
│   Sistema de unidades        [ SI ]                 ┌──────────────────────┐   │ D5
│   0 · Propiedad              [ .............. ]     │ Leyendo Table C-3    │   │ D6
│   1 · Grupo del codigo       [ .............. ]     │ (SI) — E en MPa      │   │ D7   G5:L5
│   2 · Subgrupo               [ .............. ]     └──────────────────────┘   │ D8
│   3 · Material               [ .............. ]                                │ D9
│   4 · Coeficiente / variante [ .............. ]                                │ D10
│   TEMPERATURA DE CONSULTA    [    350   ] °C     [ SELECCION COMPLETA ]        │ D11 lavanda · G11:I11
│   Modo de lectura            [ Interpolado ]                                   │ D12
├────────────────────────────────────────────────────────────────────────────────┤
│ 2 · RESULTADO                                                                  │
│  ┌───────────────┬───────────────┬───────────────┬───────────────┐             │
│  │   179 000     │      350      │  Interpolado  │   EN RANGO    │             │ KPI
│  │      MPa      │      °C       │ segun MODO_S  │ C-3 · Nota(1) │             │ unidad
│  └───────────────┴───────────────┴───────────────┴───────────────┘             │
├────────────────────────────────────────────────────────────────────────────────┤
│ 3 · FICHA DE LA PROPIEDAD — dos columnas, cada campo con su unidad             │
│   Tabla · Propiedad · Grupo impreso · Subgrupo impreso · Material              │
│   Coeficiente y su definicion impresa · Unidad impresa · Factor de escala      │
│   Tipo de dato · Rango de validez · T primera y T ultima tabuladas             │
│   Notas del codigo · Observacion de extraccion                                 │
├────────────────────────────────────────────────────────────────────────────────┤
│ 4 · TRAZABILIDAD — T1, T2, valores tabulados IMPRESOS, factor aplicado,        │
│     ecuacion empleada y archivo de resources/ del que sale la fila             │
├────────────────────────────────────────────────────────────────────────────────┤
│ 5 · CURVA DE LA PROPIEDAD  (vacia y rotulada cuando Tipo de dato = PUNTO)      │
└────────────────────────────────────────────────────────────────────────────────┘
```

### Piezas propias

**a) Conmutador SI ↔ US, en `D5`.** Dos mecánicas según cómo publique el código:

- **C-1/C-1C y C-3/C-3C** — una tabla por edición: el conmutador cambia de **hoja**
  (`DB_B31_C` ↔ `DB_B31_CC`), resolviendo por `clave_bi`, igual que los cinco buscadores.
- **C-2 y C-4** — el código imprime **los dos sistemas en la misma tabla**
  (`in./in.,°F` junto a `mm/mm,°C`; `E, ksi` junto a `E, MPa`): el conmutador cambia la
  **columna** leída, no la hoja. Las filas de estas dos tablas se escriben en las dos hojas
  con la misma `clave_bi` y cada hoja lee la columna de su sistema.

En ningún caso hay conversión: los dos valores están impresos (regla 9).

**b) Bloqueo fuera de rango, en los dos extremos.** El Apéndice C no publica `Temp. máx.`;
el límite es el propio primer y último punto tabulado de la fila:

```
=IF(FIL="","SIN SELECCION",
 IF(TIPO="PUNTO","VALOR UNICO — no depende de T; vea el rango de validez",
  IF(NP=0,"SIN VALOR TABULADO",
   IF($D$11<T_PRIMERA,"FUERA DE RANGO — T por debajo del primer punto tabulado",
    IF($D$11>T_ULTIMA, "FUERA DE RANGO — T por encima del ultimo punto tabulado",
     "EN RANGO")))))
```

y el KPI de valor **se bloquea** cuando el estado empieza por `FUERA DE RANGO`, en lugar de
sostener el valor del extremo. Es una diferencia deliberada frente a los cinco buscadores de
esfuerzos, donde el tope lo pone la columna `Temp. máx.` del código: aquí esa columna no
existe y sostener el último valor sería extrapolar, que es lo que prohíbe la regla 4.
El formato condicional rojo sobre el bloque KPI se conserva igual.

**c) Rama `PUNTO`.** C-2 y C-4 no dependen de la temperatura. El KPI muestra el valor
impreso con la unidad de la edición activa; la ficha muestra el **rango de validez impreso**
(`45–55 °F` / `7–13 °C` en C-2, `73,4 °F / 23 °C` en C-4). Si el código publica un intervalo
en vez de un número, la celda lo muestra como texto y la trazabilidad lo dice.
La celda de temperatura sigue siendo editable pero el estado avisa de que no interviene.

**d) Factor de escala visible.** El KPI muestra el valor **con el factor aplicado** y su
unidad final (`MPa`, `psi`, `10⁻⁶ mm/(m·°C)`); la trazabilidad muestra el valor **impreso**
y el factor tal como lo publica `value_axis`. Se cumplen a la vez la regla 9 (se carga lo
impreso) y la regla 6 (unidad visible en toda variable).

**e) Gráfica** — estilo idéntico al del resto (regla de estilo 3): `ch.scatterStyle="line"`,
serie del valor tabulado en azul `2F5597`, `marker="none"`, `smooth=False`; punto consultado
en rombo rojo `FF0000` sin línea. Ejes rotulados con la unidad de las dos ediciones.
Columnas propias en `_Curvas`: **`c0 = 400`** (los cinco buscadores usan `1 + i*6`, los de
grupo `200 + i*6`). Cuando `Tipo de dato = PUNTO` la gráfica queda vacía con el rótulo
«esta propiedad no depende de la temperatura».

**f) Semáforo `G11:I11`** con las mismas constantes que los cinco buscadores de cascada:
`SEL_OK_FILL` `C6EFCE` / `SEL_OK_FONT` `006100` con `SELECCION COMPLETA`; `SEL_BAD_FILL`
`FFEB9C` / `SEL_BAD_FONT` `9C6500` con `SELECCION INCOMPLETA`. El disparador es **`$D$9`**
(nivel 3 · Material), porque el nivel 4 solo tiene contenido real en C-1. Mayúsculas y sin
tilde, como todo el texto de celda del libro.

**g) Celda de temperatura** en `D11:E11`, con `TEMP_INPUT_FILL` lavanda `CCC0DA` y borde
medio rojo `C00000` — la única celda de escritura del motor, distinta del amarillo `IN_FILL`
de las listas (regla de estilo 1).

**h) Columnas ocultas: 36–39** para las listas materializadas de los niveles 1..4 y **41–42**
para las auxiliares (fila resuelta, `n_pts`, `p1`, `T1/S1/T2/S2`, valor, estado). Quedan por
debajo de `AUX_COL = 46`, de `CASC_COL = 50..54`, del bloque 60–65 de
`build_buscador_grupo` y, sobre todo, de `COL_CLAVE_BASE = 66`, que es lo que comprueba
`test_dashboard.py::test_las_claves_no_invaden_ningun_layout`.

---

## Fase 3 — Separar el Apéndice B y recolocar la navegación

| Qué | Cambio |
|---|---|
| `build_nometalicos` (`:701-744`) | deja de cargar C-2 y C-4. `DB_NoMetalicos` queda **solo con B-1…B-6**. El conteo esperado en `verificar.py` (lista `simples`, hoy `931`) baja al valor que reporte el builder |
| `Buscar_NoMetalicos` | se rerotula **«B31.3 · APENDICE B — esfuerzos de diseno hidrostatico y presion admisible, no metalicos»**. Sin propiedades físicas |
| `build_buscador_grupo` (`:2212`) | pierde los bloques *dilatación C-1* y *módulo C-3*; conserva **E (TM-1…5)** y **Poisson/densidad (PRD)**. La hoja se renombra `Buscar_Prop_IID` |
| `DB_C_dilatacion`, `DB_C_dilatacionC`, `DB_C_modulo`, `DB_C_moduloC` | **desaparecen**, sustituidas por `DB_B31_C` / `DB_B31_CC` |

> **Pendiente declarado, no efecto colateral:** tras el reparto, `Buscar_Prop_IID` no expone
> dilatación térmica de la II-D. `DB_TE` / `DB_TEC` (605 y 544 filas) siguen construyéndose y
> auditándose, pero ningún motor las consulta. Añadir el bloque TE-1 no es trivial —esa tabla
> se indexa por temperatura × columnas de grupo, no por material— y queda fuera de este plan.

### Las listas de navegación, todas a la vez

| Sitio | Cambio |
|---|---|
| `build_db_materiales.py:3187` `NAVEGABLES` | `Buscar_Propiedades` → `Buscar_Prop_IID`; insertar `Buscar_Prop_B31_3` |
| `scripts/vba/mod_nav.vba:63-72` `HojasNavegables()` | los mismos literales **y en la misma posición**: `TestSincroniaPythonVba` compara listas ordenadas, no conjuntos |
| `build_db_materiales.py:3365-3383` lista `buscadores` | tarjeta nueva `("B31.3 · APENDICE C", ["Propiedades fisicas: dilatacion y modulo", "Metales y no metalicos · SI y US"], "Buscar_Prop_B31_3")`. Quedan **9 tarjetas**, que es exactamente 3 × 3 |
| `build_db_materiales.py:3637-3646` `order` | insertar la hoja nueva tras `Buscar_Prop_IID` y sustituir las 4 `DB_C_*` por las 2 nuevas |
| `ANCLA_VOLVER` (`:3210`) | se cubre por comprensión: el layout deja libre la fila 3, columnas A..C |
| `make_vba_seed.py` | **reejecutar.** Lo que viaja al entregable es `vbaProject.bin`, no el `.vba`: sin resiembra, el libro se abre con la lista vieja |

Resultado: **38 hojas** — 1 visible (`Dashboard`), 10 `hidden` (el motor Art. 212, los 8
buscadores e `Instrucciones`) y 27 `veryHidden`.

`rewrite_instrucciones` documenta el motor nuevo: qué propiedad da cada tabla, qué significan
los coeficientes A y B de C-1, los factores ×10³ / ×10⁶ / ÷10⁶, por qué C-2 y C-4 no tienen
curva, y por qué el conmutador se comporta distinto en unas tablas y otras.

---

## Fase 4 — Verificación

### `verificar.py`

| § | Cambio |
|---|---|
| §1 conteo | añadir a `checks` (`:189-198`) `DB_B31_C` y `DB_B31_CC` = **201 filas** (52 + 44 + 75 + 30); retirar las 4 `DB_C_*` |
| §2 unicidad y §4 contigüidad | añadir `"DB_B31_C"` y `"DB_B31_CC"` a la lista `bases` (`:224-225`). **Sin más cambios**: es lo que compra el contrato de las 8 primeras columnas |
| §3b auditoría | nueva `audita_apendice_c()` en sustitución de las dos llamadas a `audita_grupo` para `DB_C_*`. Compara, por cada una de las 6 tablas y por cada fila del JSON, el vector completo de valores **y** el nombre del material contra la hoja |
| §5 portabilidad | automático: barre todas las hojas y todas las celdas |
| §6 interpolación | **generalizar el ternario de `verificar.py:566`**, que hoy solo contempla dos hojas: sustituirlo por un diccionario `{hoja: grid(wb[hoja])}`. La generación de fórmulas de `_QA` sí es genérica y no cambia. Extender `interp()` (`:155-171`) con la rama `PUNTO` |
| §8 navegación | ningún cambio: deriva de `B.NAVEGABLES` |

Casos nuevos de §6, todos sobre `DB_B31_C` y su gemela US (el valor esperado lo calcula
`interp()`, no se escribe a mano):

- C-3, acero al carbono ≤0,30 % C, a **25 °C** — punto tabulado exacto.
- ídem a **375 °C** — interpolación entre 350 y 400.
- ídem a **700 °C** — la tabla imprime `null` desde 700: **FUERA DE RANGO**, valor bloqueado.
  Es la prueba de que no se sostiene el último valor.
- ídem a **−300 °C** — por debajo del primer punto (−255): **FUERA DE RANGO**.
- C-1, Grupo 1, **coeficiente B**, a 400 °C — expansión total acumulada.
- C-1, Grupo 1, **coeficiente A**, a 412 °C — interpolación entre 400 y 425.
- C-4, `Acetal` — `PUNTO`: mismo valor a cualquier temperatura, estado «VALOR UNICO».
- C-2, `Glass–epoxy, filament-wound` — valor de texto (`"16–23.5"`), sin operar.
- El mismo material en SI y en US, comprobando que se lee la tabla nativa de cada edición.

### `test_build_db.py`

- `TestEncabezadosApendiceC` (fase 0).
- `TestApendiceC` — que el exponente sale de `value_axis` y **no** está codificado a mano;
  que `temp_to_number` resuelve las claves con U+2212 y coma de millares; que `_dedup_frase`
  colapsa `"Gray iron Gray iron"`; que la lista blanca de divergencias de nombre SI/US es
  cerrada y que cualquier divergencia fuera de ella aborta.
- `TestFormulas` — un método más: la fórmula de estado bloquea en los **dos** extremos y no
  usa ninguna función de matriz dinámica.

### `test_dashboard.py`

`TestSincroniaPythonVba` y `test_las_claves_no_invaden_ningun_layout` ya cubren los dos
fallos más probables de esta fase (olvidar el VBA, invadir la columna 66). No hay que
añadir nada.

### Comandos

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts

# Fase 0 — una sola vez, o al reponer la extraccion del Apendice C
python completar_apendice_c.py --resources ..\..\..\resources `
    --ocr ocr_apendice_c_datalab.json

# Obligatorio: se toco vba\mod_nav.vba
python make_vba_seed.py

python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```

`verificar.py` debe seguir devolviendo **0**. Las 271 276 comparaciones fila a fila de las
bases de esfuerzos, la contigüidad de la cascada, cero matriz dinámica, cero validaciones no
portables y la regresión del caso semilla (`Sa` 138 MPa, dictamen **APTO**) **no pueden
moverse**: si se mueven, el reparto del Apéndice C rompió algo ajeno.

### Comprobación manual, con el `.xlsm` abierto

- El Dashboard muestra **9 tarjetas** en 3 filas de 3; la nueva abre `Buscar_Prop_B31_3`.
- Seleccionar `MODULO DE ELASTICIDAD - METALES` → `(sin grupo impreso)` →
  `Carbon steels with carbon content 0.30% or less` → `(unico)`, a 350 °C: **179 000 MPa**
  (impreso 179, factor 10³). Conmutar a **US**: la misma fila en `Table C-3C`, a 662 °F.
- Poner 700 °C: KPI en rojo, «FUERA DE RANGO — T por encima del ultimo punto tabulado».
- Seleccionar `MODULO ELASTICIDAD CORTO PLAZO - NO METALICOS` → `Thermoplastics` →
  `Acetal`: 2 830 MPa / 410 ksi, estado «VALOR UNICO», gráfica vacía y rotulada.
- `Buscar_NoMetalicos` ya no ofrece dilatación ni módulo; `Buscar_Prop_IID` ya no ofrece
  Apéndice C.
- Ninguna `DB_*` aparece en el menú «Mostrar» de Excel.

---

## Fase 5 — `CLAUDE.md`

Registrar la regla que pidió el usuario, en «Reglas de diseño del libro — no romper»:

> **10. Conmutador de unidades en todo motor.** Todo motor de búsqueda y todo motor de
> cálculo del libro lleva un conmutador **SI ↔ US**. Los dos sistemas se cargan como
> **extracciones independientes** de la edición correspondiente del código, nunca por
> conversión (corolario de la regla 9). Cuando el código publica los dos sistemas en la
> **misma tabla** —C-2 y C-4 del B31.3— el conmutador cambia la **columna** leída; cuando
> publica **una tabla por edición** —A-1/A-1C, C-1/C-1C, C-3/C-3C, II-D métrica y US—
> cambia la **hoja**. Si un material no tiene homólogo en la otra edición, el motor muestra
> «sin equivalente en la edicion US», nunca un valor convertido.

Y actualizar el resto del documento: 38 hojas; entregable `Motor_de_Calculo_ASME_PCC_Rev4.xlsm`;
`extraer_apendice_c.py` en el bloque de reconstrucción; el reparto Apéndice B / Apéndice C;
y la nota de que el bloqueo por rango del Apéndice C se toma de la banda tabulada, no de una
columna `Temp. máx.`.

`LEEME_Nota_de_Version.md`: sección Rev. 4 con el motor nuevo, el reparto y el pendiente
de la dilatación II-D (TE-1).

---

## Orden de ejecución

1. Copiar el PDF y el OCR de Datalab a `resources/.../appendix_c/`; correr
   `completar_apendice_c.py` + `TestApendiceCCompletado`. **Bloqueante**: sin
   `coefficient_defs` ni `note_members` no se construye nada.
2. `build_apendice_c` → `DB_B31_C` / `DB_B31_CC`, con la aserción de paridad SI/US.
   Retirar las 4 `DB_C_*`.
3. Recortar `build_nometalicos` y `build_buscador_grupo`; renombrar `Buscar_Prop_IID`.
4. `build_buscador_prop_c` / `finish_buscador_prop_c` + entrada en `build_listas`.
5. Navegación: `NAVEGABLES`, `mod_nav.vba`, tarjeta del Dashboard, `order`; reejecutar
   `make_vba_seed.py`.
6. Ampliar `verificar.py` (§1, §2/§4, §3b, §6) y `test_build_db.py`.
7. `CLAUDE.md`, `Instrucciones` y la nota de versión.

## Riesgos

| Riesgo | Mitigación |
|---|---|
| El OCR de Datalab introduce un valor distinto del que ya está en `resources/` | El script **no elige**: audita los dos y para si discrepan. Los valores no se tocan; la fase 0 solo escribe metadatos |
| Las 5 excepciones de indentación se olvidan al reponer la extracción | Van en el script como tabla literal con folio, y `TestApendiceCCompletado` comprueba que `Gray iron` y las tres filas de C-2 no tienen subgrupo |
| El enlace posicional SI↔US se desalinea en una reextracción futura | Aserción de paridad en el build: filas por propiedad y nombre normalizado contra lista blanca. Aborta, no avisa |
| Confundir el factor 10³ con 10⁶ | El exponente se **lee** de `value_axis` y se guarda con su texto impreso; test dedicado |
| Olvidar el VBA y que el libro se abra con la lista vieja de hojas | `TestSincroniaPythonVba` falla; y `make_vba_seed.py` está en el orden de ejecución, no como nota al pie |
| Regresión en las bases de esfuerzos al tocar el builder | `verificar.py` audita las 271 276 comparaciones y el caso semilla en cada corrida |
| Que el motor sostenga el último valor tabulado fuera de rango | Bloqueo explícito en los dos extremos + 2 casos de §6 que lo comprueban recalculando en Excel |

---

*Herramienta de ingeniería de referencia. Verificar entradas y resultados antes de emitir
para construcción.*
