# Plan — Bases de datos de ASME BPVC Sección II, partes A, B y C

**Documento:** PLAN-SECII-ABC-001 · Rev. 0 · 2026-09-07
**Entregable objetivo:** `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm` (39 → **48 hojas**)
**Fuente única de verdad:** `resources/ASME_BPVC/Sec_II/bpvc_ii_{a_1,a_2,b,c}/specifications/*.json`
**Unidades:** las impresas por el código (la Sección II imprime doble, `48 000 [330]`)

---

## ⏸ PARADO EN EL PUNTO DE DECISIÓN DE LA FASE 2 — 2026-09-07

| Fase | Estado |
|---|---|
| 0 · Guardar el plan | ✅ |
| 1 · Núcleo `secii_tablas.py` y piloto | ✅ · 12 especificaciones, más `test_secii_tablas.py` (19 pruebas) |
| 2 · Barrido completo e informe | ✅ · `outputs/Base_Datos_Materiales_ASME/Revision_Tablas_SecII.md` |
| 3 · Volcado íntegro al libro | ⏸ **no ejecutado** |
| 4 · Hojas normalizadas | ⏸ **no ejecutado** |
| 5 · Cierre | ⏸ parcial: `CLAUDE.md` documenta el estado |

**Lo medido sobre las cuatro partes** (barrido completo en 13 s):

| Magnitud | Valor |
|---|---:|
| Especificaciones recorridas | 368 |
| Tablas lógicas (tras unir continuaciones) | 2 572 |
| Filas reconstruidas | 54 198 |
| Bloques `Line` consumidos | 124 115 |
| Celdas escritas | 218 202 |
| Notas al pie recogidas | 5 233 |
| **Comprobación sin pérdida** | **0 fallos** |
| Huecos declarados | 101 `Form` (parte C), 145 `Caption` vacíos, 22 `Table` sin `Line` |

| Confianza | Filas | % de todas | Filas de dato | % de las de dato |
|---|---:|---:|---:|---:|
| EXACTA | 16 451 | 30,4 % | 9 189 | 25,4 % |
| POR CONTEO | 12 935 | 23,9 % | 12 935 | 35,8 % |
| **AMBIGUA** | 24 812 | 45,8 % | **13 996** | **38,7 %** |

**Por qué se para.** El plan lo dice con todas las letras: *«Aquí se para y se
enseña: si la fracción AMBIGUA no es marginal, se corrige el algoritmo antes de
tocar el libro. Escribir 124 000 filas dudosas es peor que no escribirlas.»* Un
38,7 % no es marginal.

El algoritmo ya pasó por cuatro rondas de corrección, cada una con su motivo, y bajó
la ambigüedad de 66,8 % a 38,7 %:

1. **El agrupamiento de filas iba sobre el borde y no sobre el centro.** Entre el
   pie de una fila y la cabeza de la siguiente hay 2 pt, dentro de la tolerancia de
   3 pt que hace falta para no partir un superíndice: la tabla entera colapsaba en
   una fila.
2. **El superíndice se separaba de su valor.** `0.25<i><sup>A</sup></i>` daba
   `0.25 A` y la fila pasaba de 3 fichas de valor a 6.
3. **Las líneas anchas tapaban los huecos entre columnas.** Un rótulo que cruza
   varias columnas, dentro de la proyección en x, fusiona la tabla en una banda.
4. **Una banda se parte cuando una fila la desmiente.** Si dos `Line` de la misma
   fila caen en ella, esa banda son dos columnas —y el corte va por el hueco más
   ancho, nunca por uno menor que el que separa dos trozos de una misma celda—.

Y se detuvo y corrigió un fallo propio: el reparto por bandas **reordenaba** el
texto de 23 filas —un rótulo que cruza columnas se colocaba por su punto medio
detrás de lo que va después—. Ahora esa condición marca la fila AMBIGUA en vez de
reordenar, y la comprobación sin pérdida vuelve a 0.

**Lo que sí está garantizado, y es lo que se puede garantizar sin el PDF:** ninguna
fila pierde texto, ninguna fila AMBIGUA se rellena por interpolación sobre el
`bbox`, y los huecos del origen se cuentan y se nombran.

**La decisión que queda al ingeniero:** seguir a las Fases 3-5 tal cual, acotarlas a
las tablas que sí se reparten (las de confianza EXACTA o POR CONTEO al 100 %), o
reponer la extracción de la Sección II con una herramienta que conserve los `Span`
—que es la causa raíz: sin ellos, la posición de cada palabra dentro de un `Line`
no está en el fichero—.

---

## Contexto

El commit `f97c154` metió en `resources/` las partes A, B y C de la Sección II: **379 entradas
de índice** (368 especificaciones SA-/SB-/SFA- y 11 apéndices), 227 MB de JSON, 4 218 páginas.
**Ningún script las consume todavía.** `build_db_materiales.py` no las nombra ni una vez; la
única referencia en todo el repo está en `verificar_seccion_ii.py:47-50`, que audita paginación.

El motor Rev. 3 sabe *qué esfuerzo admite* un material (II-D y B31.3) pero no sabe **qué es** ese
material: no tiene su composición química, ni sus requisitos de tracción por grado y forma, ni sus
notas de tabla. Ese hueco ya se ha pagado a mano dos veces —la decisión del 9Cr-1Mo-V se apoyó en
la Parte A y en SFA-5.5 leídas a ojo (`CLAUDE.md:328-334`), y `prompt_notebooklm.md` pide a un
tercero confirmar 40+ identidades UNS↔composición contra la Parte A— y sigue bloqueando las **118
decisiones abiertas** de `Revision_MAP_Grupo.md`, casi todas aleaciones de níquel y austeníticos
cuya composición sí está impresa en las partes A y B ya cargadas.

**Resultado buscado:** las tablas de las 368 especificaciones, tabuladas y consultables dentro del
propio libro, sin que ningún valor pase por criterio del modelo.

### Decisiones ya tomadas por el ingeniero

| | |
|---|---|
| Alcance | Partes A (2 vol.), B y C. **Todas** las tablas de cada especificación |
| Entrega | Hojas en el `.xlsm`. **Sin capa JSON derivada** en `resources/` |
| Libro | El mismo motor → **Rev. 4** |
| Disposición | Híbrido: catálogo + normalizadas (química y tracción) + volcado íntegro por parte |
| Consulta | Solo hojas de datos. Sin buscador ni fórmulas de consulta |
| PDFs | No se usan: las extracciones ya se auditaron contra ellos |

**Consecuencia buena de no crear JSON derivado:** no existe una segunda copia de los valores del
código. Cada celda se traza a `spec_id` → `block_id` (`/page/224/Table/27`) → `bbox`, es decir al
propio fichero de `resources/`. Es la lectura más estricta posible de la Regla nº 1.

### Un límite que hay que decir en voz alta

`verificar_seccion_ii.py` auditó **anclaje, folio impreso, cobertura de rangos y figuras** — que
cada especificación empieza donde dice. **No auditó el contenido de ninguna celda.** Y el propio
`meta.json` declara: cobertura de texto **97,0 %** en A1, pero **88,1 % en sus 41 páginas
apaisadas** (`SA-240/SA-240M` al 87 %) — que son exactamente las tablas anchas de aleación y
propiedades mecánicas. Una tabla puede estar perfectamente **localizada** y aun así haber perdido
texto.

Esto no cambia la decisión (se construye desde el JSON), cambia **qué se promete**: el extractor
lleva una comprobación *sin pérdida* que garantiza que nuestra capa no añade ni quita un solo
carácter, y toda fila que no se pueda repartir sin ambigüedad **se declara**, no se rellena.

---

## Lo que hay que resolver: el JSON no trae tablas

El `html` de todo bloque `Table` es `"<p></p>"`. No hay `<tr>` ni `<td>`. Es **deliberado** —
`meta.json.conversion.why_inverted_table_config` dice que se rechazó la reconstrucción de tablas de
pdftext para conservar los bloques de texto originales con sus `bbox`. Las filas cuelgan como
`Table` → `Line`, y **los `Span` no se conservan: el bloque más fino es `Line`.**

Materia prima medida:

| Parte | Specs | Bloques `Table` | de ellos sueltos | Bloques `Line` | Cobertura |
|---|---:|---:|---:|---:|---:|
| `bpvc_ii_a_1` | 76 | 604 | 215 | 26 936 | 97,0 % |
| `bpvc_ii_a_2` | 114 | 551 | 160 | 20 223 | 96,8 % |
| `bpvc_ii_b` | 143 | 1 256 | 387 | 46 424 | 95,2 % |
| `bpvc_ii_c` | 35 | 485 | 152 | 30 532 | 95,5 % |
| | **368** | **2 896** | **914** | **124 115** | |

**Invariante verificado:** los 26 936 `Line` de A1 son, sin excepción, hijos directos de un
`Table`. Si ves un `Line`, estás dentro de una tabla. Eso acota el problema entero.

### Los ocho casos que rompen un extractor ingenuo

1. **Una fila puede ser un solo `Line`** (`"Carbon, max 0.25ᴬ 0.30ᴮ 0.35ᴮ"`) **o varios**
   (`"Tensile strength, min, psi [MPa]"` + `"48 000 [330]"` + `"60 000 [415]"` + `"70 000 [485]"`,
   cuatro `Line` con la misma `y`). **Las dos formas conviven en la misma tabla** (SA-106 Tabla 2).
2. **El espacio no siempre separa celdas:** `48 000` es separador de millares; `[330]` es la unidad
   SI que imprime el propio código.
3. **Una celda puede ser dos `Line`:** SB-111 Tabla 1 → `"99.99"` + `"minᴬ"`, contiguos en x, con la
   `y` desplazada ~1 pt porque el superíndice sube la caja.
4. **Continuaciones entre páginas, con dos convenciones:** `<b>TABLE A2.1</b> <i>Continued</i>` en
   A y B (cero `(Continued)` en A1), `Table 1 (Continued)` en la parte C (51 casos en 16 ficheros).
   **La cabecera solo está en la primera página**: hay que arrastrarla.
5. **El 32 % de los `Table` va suelto**, sin `TableGroup` ni `Caption`: el título es un
   `SectionHeader` encima (SA-6 p.71).
6. **Marker duplica bloques y solo uno lleva el texto:** en SFA-5.9 p.359 hay 8 pares
   `Text`/`Footnote` con `bbox` casi idéntico, y el contenido **alterna** entre uno y otro. Hay que
   deduplicar por solape de `bbox` quedándose con el `html` no vacío.
7. **Etiquetas pegadas:** `<b>TABLE</b><b>1</b><b>Chemical</b><b>Requirements</b>` (SB-111). Quitar
   tags a lo bruto da `TABLE1ChemicalRequirements`.
8. **Cabeceras a dos niveles y guionadas:** `Other Required Elements` cubre `Element`+`Amount` (solo
   se deduce del solape en x); `"Longitu"` + `"dinal"` son dos `Line` apilados, y
   `"Transverse Longitu"` mezcla el final de una columna con el principio de la siguiente.

Además, ya localizados y **declarables como hueco, no reparables sin el PDF**: los **101 bloques
`Form` de la parte C** (`html` vacío, `children: []`), los `Caption` vacíos de SA-6 (p. 62, 77, 100),
los marcadores de nota corruptos (SA-106 Tabla 2 → `"For longitudinal strip tests AA A"`) y el texto
pegado sin espacios de `/page/88/Caption/5`.

**No usar `section_hierarchy` como semántica:** el `Caption` de la Tabla 1 de SA-106 cuelga de
`/page/224/SectionHeader/32` = «11. Bending Requirements», y la tabla pertenece a la §7.

---

## Arquitectura

Un módulo nuevo, puro y sin Excel, más un consumidor en el builder ya existente:

```
outputs/Base_Datos_Materiales_ASME/scripts/
  secii_tablas.py      NUEVO — reconstruye las tablas desde Line+bbox. Sin PDF, sin Excel.
                       Es librería (la importa el builder) y CLI (--informe, sin tocar el libro).
  build_db_materiales.py   MODIFICADO — 9 hojas nuevas + NAVEGABLES
  verificar.py             MODIFICADO — nueva §10
  test_secii_tablas.py NUEVO — regresiones ancladas a filas concretas del código
  test_dashboard.py        MODIFICADO — sincronía Python↔VBA de las 9 hojas
  vba/mod_nav.vba          MODIFICADO — HojasNavegables()
```

### Qué se reutiliza (no reinventar)

| Pieza | Dónde | Para qué |
|---|---|---|
| `Resources` | `scripts/db_lib.py:57` | Cargar y cachear JSON. Ningún `open()` a mano |
| normalización de texto/guiones/elipsis, `build_material_id`, `disambiguate` | `scripts/db_lib.py` | Claves y limpieza, con las mismas reglas que el resto del libro |
| `_colof(x0, x1, bounds)` | `tools/b31_3_extractor/tables.py:234` | Asignar un run a su columna por el **punto medio**, con tolerancia ±3 |
| `_bands(lines, min_gap)` / `_band_of` | `tools/b31_3_extractor/special.py:14,:29` | Clustering por x **cuando no hay reglas impresas** — que es justo nuestro caso |
| `merge_wrapped(rows, key_col=0)` | `tools/b31_3_extractor/build.py:58` | Reintegrar la línea de continuación por criterio estructural |
| `_colnames` / `_key` / `_uniq` | `tools/b31_3_extractor/build.py:14,:34,:46` | Apilar cabeceras multinivel en un nombre por columna |
| `clean` / `num` / `is_ellipsis` | `tools/b31_3_extractor/b313.py:32,:45,:39` | Tipado de celda impresa: millares con espacio, menos Unicode, elipsis ASME |

Del resto de `tools/b31_3_extractor/` **no sirve nada**: asume PyMuPDF, reglas vectoriales y
caracteres con `bbox` propio, y aquí no hay ninguna de las tres cosas. Importar con
`sys.path.insert` como ya hacen los scripts entre sí; si las rutas *hardcoded* de otra máquina
(`b313.py:11`, `run.py:22`) estorban, copiar esas seis funciones puras a `secii_tablas.py` citando
su origen.

### El algoritmo, y dónde se planta

```
cargar_spec → aplanar bloques → deduplicar por IoU de bbox (quedarse con html no vacío)
  → recoger todo Table (en TableGroup y suelto)
  → titular: Caption hermano (probar posición 0 Y 1) → si no, SectionHeader/Caption encima
  → unir continuaciones (dos regex, A/B y C) arrastrando la banda de cabecera
  → agrupar Line en filas por solape en y (tol. 3 pt: el superíndice sube la caja ~1 pt)
  → derivar ejes de columna de la banda de cabecera (Line con <b> en el tercio superior)
     + histograma de x de las filas; fusionar los dos niveles por solape en x
  → repartir cada fila
  → recoger notas al pie (Text/Footnote bajo el bbox, que empiezan por marcador)
```

**El reparto de fila tiene tres resultados, y ninguno adivina:**

| Confianza | Cuándo | Cómo |
|---|---|---|
| `EXACTA` | la fila trae ≥2 `Line` | cada `Line` a su columna por `bbox.x` (`_colof`). Sin partir texto |
| `POR CONTEO` | la fila es **un** `Line` | se protegen los patrones `48 000`, `[330]`, `0.27–0.93`, `8 × C to 1.0`, `11⁄2`; se parte por espacios **desde la derecha** (las n-1 últimas fichas son valores, el resto es la etiqueta); solo vale si el conteo cuadra con el nº de columnas **y** el `bbox` del `Line` abarca esas columnas |
| `AMBIGUA` | ni lo uno ni lo otro | el texto impreso se conserva **entero en una celda**, la fila se marca y entra en el informe de revisión. **No se interpola sobre el `bbox`** |

Esa última línea es la decisión de diseño importante. Sin `Span`, la posición de cada palabra dentro
de un `Line` **no está en el fichero**: repartirla por interpolación lineal sobre el ancho del
`bbox` sería inventar estructura con una fuente proporcional. Es exactamente lo que la Regla nº 1
prohíbe. La geometría se usa para **confirmar** un reparto por conteo, nunca para producirlo.

**Comprobación sin pérdida (`verificar_sin_perdida`), obligatoria.** Para cada fila reconstruida, la
concatenación normalizada de sus celdas debe ser **carácter a carácter** igual a la concatenación
normalizada del texto de sus `Line` de origen. Garantiza que la tabulación es una *repartición* del
texto impreso y nunca una adición. Corre en el build y en las pruebas; un fallo aborta.

No demuestra que el JSON reproduzca el PDF —eso se auditó aguas arriba y con las reservas del §
anterior—, demuestra que **nuestra capa no mete nada**. Es lo que sí podemos garantizar sin el PDF.

---

## Las 9 hojas nuevas

| Hoja | Filas aprox. | Contenido |
|---|---:|---|
| `CAT_SecII` | 379 | Catálogo: parte, `folder`, spec id, título, designación ASTM/AWS equivalente y su nota, páginas PDF, folios impresos, nº de figuras, nº de tablas detectadas, fichero fuente |
| `IDX_SecII_Tablas` | ~1 400 | Una fila por **tabla lógica** (tras unir continuaciones): spec, `tabla_id`, título, páginas, bloques que la componen, filas, columnas, apaisada s/n, reparto `EXACTA`/`POR CONTEO`/`AMBIGUA` y motivo |
| `DB_SecII_A1` `DB_SecII_A2` `DB_SecII_B` `DB_SecII_C` | ~124 000 en total | **Volcado íntegro.** Una fila por fila impresa, formato ragged: `Spec. No. \| Tabla \| Fila \| Tipo (cabecera/dato/nota) \| Pagina PDF \| Folio \| Bloque \| Confianza \| C01 … Cnn` |
| `DB_SecII_Quimica` | ~12 000 | Normalizada. Fila = (spec, tabla, designación/grado/UNS); una columna por elemento (C, Mn, P, S, Si, Cr, Ni, Mo, V, Nb, N, Cu, Al, Ti, B, Co, W, Ta, Zr, Fe, Pb, Sn, Zn) + `Otros elementos` + `Notas` + trazabilidad |
| `DB_SecII_Traccion` | ~6 000 | Normalizada. Fila = (spec, tabla, grado, forma/condición, dirección); `Rm min/max`, `Re min/max`, `Alargamiento %`, `Base de medida`, `Reduccion de area %`, `Dureza` |
| `DB_SecII_Notas` | ~1 500 | Notas al pie y bloques `NOTE n—`, con su marcador (`A`, `B`, `a`, `b`…) y la tabla a la que pertenecen |

**Reglas que gobiernan estas hojas:**

- **Valor tal como está impreso** (Regla 9). La celda de química dice `0.27–0.93` o `0.25 max`, no
  un mínimo y un máximo numéricos. Parsear un rango es interpretación, no transcripción, y no entra
  en esta revisión. Lo que no encaje en una columna de elemento va **entero** a `Otros elementos`.
- **La doble unidad es del código, no nuestra.** `48 000 [330]` se parte en `Rm min (psi)` = 48 000 y
  `Rm min (MPa)` = 330. Es una **partición de la celda impresa**, jamás una conversión (Regla 10:
  cuando el código publica los dos sistemas en la misma tabla, el conmutador cambia de columna).
- **Tablas transpuestas.** SA-106 Tabla 2 imprime los **grados como columnas** y las propiedades
  como filas. `DB_SecII_Traccion` detecta ese caso (primera columna = nombres de propiedad,
  cabecera = grados) y transpone; si no lo detecta con certeza, la tabla se queda solo en el volcado
  íntegro y se lista.
- **Normalizar solo lo que casa entero.** Una tabla pasa a `Quimica`/`Traccion` únicamente si
  **todos** sus encabezados se resuelven contra el diccionario de elementos/propiedades. Las demás
  se quedan en el volcado y aparecen en el informe. Sin encaje forzado.
- **`Pagina PDF` es citable (1-based).** `pdf_pages` es **0-based** en A/B/C — al revés que la II-D.
  La hoja publica `pdf_page + 1`, más `Folio impreso`, más `Bloque` (que lleva el índice crudo).
- Cero funciones de matriz dinámica (Regla 1 del libro). Estas hojas no llevan fórmulas.

**Navegación.** «Solo hojas de datos» significa sin buscador ni fórmulas de consulta, no
inalcanzables: el libro abre con todo oculto salvo el `Dashboard`. Las 9 hojas se añaden a
`NAVEGABLES` (`build_db_materiales.py`) **y** a `HojasNavegables()` (`scripts/vba/mod_nav.vba`) —
las dos a la vez, o el libro se abre mostrando algo distinto de lo que se construyó.

---

## Fases

**Fase 0 — Guardar el plan.** Copiarlo a
`outputs/plans/Plan_Bases_Datos_Seccion_II_ABC.md`, junto a los otros cuatro planes del proyecto y
con su misma cabecera de metadatos.

**Fase 1 — Núcleo y piloto.** Escribir `secii_tablas.py` con el pipeline y el CLI
`--informe` (mide, no escribe nada). Piloto sobre 12 especificaciones elegidas por ser las que
duelen: `SA-106` (los dos modos de fila conviviendo), `SA-6` (5,5 MB, tablas sueltas, `Caption`
vacíos), `SA-240` (la peor cobertura, 87 %), `SA-182`, `SA-193`, `SA-312`, `SA-335`, `SB-111`
(celda en dos `Line`), `SB-265`, `SFA-5.9` (continuación de 4 páginas, apaisada, duplicados
`Text`/`Footnote`), `SFA-5.5`, `SA-479`. Salida: reparto por confianza y coste en filas y celdas.

**Fase 2 — Barrido completo y punto de decisión.** Las cuatro partes. Genera
`outputs/Base_Datos_Materiales_ASME/Revision_Tablas_SecII.md` con las tablas ambiguas y los huecos
declarados (los 101 `Form` de la parte C, los `Caption` vacíos, los marcadores corruptos).
**Aquí se para y se enseña**: si la fracción `AMBIGUA` no es marginal, se corrige el algoritmo antes
de tocar el libro. Escribir 124 000 filas dudosas es peor que no escribirlas.

**Fase 3 — Volcado íntegro al libro.** `CAT_SecII`, `IDX_SecII_Tablas`, los cuatro `DB_SecII_*` y
`DB_SecII_Notas` en `build_db_materiales.py`. Medir tamaño y tiempo de build.

**Fase 4 — Hojas normalizadas.** `DB_SecII_Quimica` y `DB_SecII_Traccion`, con el diccionario de
elementos y propiedades y la detección de transposición.

**Fase 5 — Cierre.** Navegación sincronizada, `verificar.py` §10, pruebas, y actualizar `CLAUDE.md`
(estructura, hojas, procedimiento de reconstrucción) y `knowledge/claude.md`, cuyo §5 todavía
enumera un `resources/` sin las partes A, B y C.

---

## Verificación

`test_secii_tablas.py` — estilo del repo: clases `TestX` sin `unittest`, `assert` planos, y la mitad
de las aserciones contra los JSON reales de `resources/`. Regresiones ancladas a filas concretas:

| Prueba | Qué detiene |
|---|---|
| SA-106 Tabla 1 → 11 filas × 4 columnas; `Manganese` = `0.27–0.93 \| 0.29–1.06 \| 0.29–1.06`; `Carbon, max` conserva los marcadores `A`, `B`, `B` | el reparto por conteo y los superíndices |
| SA-106 Tabla 2 detectada como transpuesta; `Grade B` → `Rm min` = `60 000` psi / `415` MPa | la transposición y la partición de la doble unidad |
| SB-111 Tabla 1 → la celda `99.99 minᴬ` sale de un solo trozo | la reunión de dos `Line` en una celda |
| SFA-5.9 Tabla 1 → **una** tabla lógica de las páginas 356-359, con la cabecera arrastrada | la unión de continuaciones y el `(Continued)` de la parte C |
| SFA-5.9 p.359 → el texto de las 8 notas aparece **una vez** | la deduplicación `Text`/`Footnote` |
| SA-6 p.71 Tabla 2 titulada desde su `SectionHeader` | las tablas sueltas (914 en total) |
| SB-111 `Caption` → `TABLE 1 Chemical Requirements`, no `TABLE1ChemicalRequirements` | el separador al cerrar tags inline |
| `verificar_sin_perdida` sobre ≥50 tablas de las cuatro partes | que la capa no añada ni pierda un carácter |
| El nº de `Form` de la parte C sigue siendo 101 y todos declarados como hueco | que un hueco no se rellene en silencio |

`verificar.py` §10 — audita el libro construido: conteo tabla→hoja por parte, unicidad de
`tabla_id`, la comprobación sin pérdida releída **desde la hoja**, ausencia de matriz dinámica en
las 9 hojas y presencia de las 9 en la capa de navegación. Devuelve 0 solo si todo pasa.

`test_dashboard.py::TestSincroniaPythonVba` — extendida a las 9 hojas.

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts

# Fase 1-2: mide y reporta, no escribe nada
python secii_tablas.py --resources ..\..\..\resources --informe ..\Revision_Tablas_SecII.md

# Fase 3-5
python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_secii_tablas.py test_build_db.py test_dashboard.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```

`verificar.py` necesita Excel instalado: recalcula con el motor real.

---

## Riesgos

| Riesgo | Mitigación |
|---|---|
| **El libro se vuelve inmanejable.** ~2 M celdas nuevas sobre las 271 276 actuales | Se mide en la Fase 3 antes de decidir. El formato ragged (una fila por fila impresa) cuesta ~4× menos celdas que el formato largo celda-a-celda |
| **El build se alarga:** 227 MB de JSON parseados en cada corrida | Aceptable si queda en minutos. Si duele, caché opcional en el directorio de trabajo temporal — **nunca** en `resources/` ni en el repo |
| **Fracción `AMBIGUA` alta**, sobre todo en las 247 páginas apaisadas | Punto de parada explícito en la Fase 2. Se declara, no se rellena |
| **Encabezados heterogéneos**: cada spec nombra sus elementos a su manera | Solo se normaliza la tabla cuyos encabezados casan **todos**; el resto vive en el volcado íntegro, que no pierde nada |
| **La navegación se desincroniza** entre Python y VBA | `test_dashboard.py::TestSincroniaPythonVba`, ya existente |
| **Excel COM se cuelga en un modal invisible** (`tasks/lessons.md`, 2026-09-06) | No se toca el proyecto VBA: `make_vba_seed.py` solo se recorre si cambia `mod_nav.vba`, y su `lint_vba()` corre antes de tocar COM |
| **Rev. 3 se rompe al pasar a Rev. 4** | `verificar.py` audita las 39 hojas existentes en cada build; se corre entero, no solo la §10 |
