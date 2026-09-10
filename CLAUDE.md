# Proyecto ASME PCC — Motor de Cálculo e Ingeniería de Reparación

Ingeniería de reparación de equipos a presión y tubería según la familia **ASME PCC**
(PCC-1 / PCC-2 / PCC-3), con los códigos de construcción **ASME B31.3-2024** y
**ASME BPVC Sección VIII-1 + II-D 2025** como referencia.

Responde siempre en **español**. Unidades **SI por defecto** (MPa, mm, °C).
Tono profesional y directo, sin relleno.

---

## Regla nº 1 — `resources/` es la única fuente de verdad

`resources/` contiene los códigos y normas extraídos a JSON. **Ningún valor normativo
—esfuerzos admisibles, factores, fórmulas, límites— puede salir de la memoria del
modelo.** Siempre se lee el archivo fuente antes de citar o calcular.

Si un dato no existe en `resources/`, **no se inventa ni se aproxima**: se declara
el vacío explícitamente y se pregunta cómo proceder.

Todo coeficiente o tabla que alimente un motor de cálculo debe **trazar** al archivo
concreto de `resources/` del que procede, para permitir auditoría posterior.

## Estructura

```
knowledge/     Instrucciones de cálculo ASME PCC-2 en SI. Leer antes de cualquier tarea.
resources/     Códigos y normas (JSON). Fuente única de verdad.
               ├─ ASME B31/ASME B31.3/
               │  ├─ APPEX/                    Apéndices A, B y C
               │  └─ CHAPTERS/tables/          Tablas del cuerpo normativo
               ├─ ASME PCC/pcc_2/              Artículos de PCC-2
               └─ ASME_BPVC/Sec_II/
                  ├─ bpvc_ii_a_1/, a_2/, b/, c/  Partes A, B y C: texto íntegro
                  │                              de 379 especificaciones
                  ├─ bpvc_ii_d_metric_2025/    II-D métrica (MPa, °C)
                  └─ bpvc_ii_d_customary_2025/ II-D U.S. Customary (ksi, °F)
outputs/       Entregables.
               ├─ Base_Datos_Materiales_ASME/  Los scripts del motor y sus informes.
               ├─ plans/                       Planes EN CURSO.
               │  └─ registro/                 Planes ya ejecutados; ver su LEEME.md.
               └─ <proyecto>/                  Un entregable de ingeniería por carpeta.
templates/     maestro_con_macros.xlsm — Rev. 0 + proyecto VBA, entrada del builder.
               Se genera con scripts/make_vba_seed.py; no se edita a mano.
tools/         b31_3_extractor: el extractor con el que se produjo el B31.3 de
               resources/. Solo se toca para reponer esa extracción.
tasks/         lessons.md — errores ya pagados y la regla que dejó cada uno.
Generales/     Material de referencia suelto (el PNG del layout del Dashboard).
.agents/skills/  Skills instalados + skills-lock.json. `.claude/skills/` son
                 junctions a estos y no se versionan.
```

**Nunca leas `outputs/` ni `templates/`** salvo que se te señale un archivo. La
excepción es `outputs/Base_Datos_Materiales_ASME/scripts/`, que es **código
fuente**, no entregable: ahí vive el builder y sus pruebas.
**Guarda todo entregable en `outputs/` dentro de una subcarpeta** con nombre de proyecto.

Ante una duda de alcance, pregunta antes de producir.

---

### `pdf_pages` no usa la misma base en todas las extracciones

Cuesta caro equivocarse aquí, y ya se pagó una vez: una «corrección» de 16 citas
de II-D que en realidad las rompió, revertida después.

| Fuente | Base de `pdf_pages` | Cómo se comprobó |
|---|---|---|
| BPVC **II-D** (métrica y US) | **1-based** inclusiva | El UNS `K01700` de la Tabla 1A está en la página 58, la primera que declara |
| BPVC **II A, B y C** | **0-based** inclusiva | `[56, 119]` de SA-6/SA-6M son las páginas 57 a 120; la 56 está en blanco |
| B31.3 (capítulos, apéndices, tablas) | 1-based | El rótulo de cada tabla cae en la página que declara |

**No se detecta el desplazamiento con el rótulo de la tabla**: se repite en cada
página de continuación, así que casa con las dos convenciones y no distingue
nada. Hay que usar algo que aparezca **una sola vez** — un UNS, o una fila
concreta. Ese fue exactamente el error.

`verificar_resources.py` audita PCC-2, los capítulos del B31.3, sus apéndices
D–Z y las dos ediciones de II-D. `verificar_seccion_ii.py` audita **solo** las
partes A, B y C: la II-D se le quitó para no mantener dos scripts con dos
convenciones.

**Los dos necesitan los PDF y devuelven distinto de 0 sin ellos.** No es un
fallo: los PDF no se versionan (copyright de ASME) y sin `--pdfs` cada fuente
sale como `REVISAR`, que significa «no comprobado», no «mal». Solo los puede
correr en verde quien tenga los PDF en disco. Esto **no** afecta a
`verificar.py`, que audita el libro contra los JSON y sí corre en cualquier
copia del repositorio.

### Dónde están los PDF fuente

`C:\Users\dvelasquez\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES`
tiene los códigos y normas en PDF (II-D métrica/US, Secciones II A/B/C, B31.3)
que alimentan `--pdf`/`--pdfs` en los scripts de arriba y en
`verificar_resources.py`/`verificar_seccion_ii.py`.

**Nunca se añaden PDF a este repositorio** (copyright de ASME; ver
`.gitignore`). Y el PDF **solo se usa para corregir y verificar** los JSON de
`resources/` —auditar folios, reconstruir una tabla colapsada, resolver una
ambigüedad de extracción—, **nunca para poblar un motor o un cálculo
directamente**: todo dato que entra a un motor sale de `resources/`, tal como
exige la Regla nº 1. Un PDF que aporta un dato nuevo pasa primero por un
script de extracción/corrección que lo escribe en el JSON correspondiente
(con su `extraction_amendments` o metadato de procedencia), y solo entonces
ese JSON alimenta el builder.

### Sección II, partes A, B y C — cómo está troceada

379 especificaciones de material (SA-, SB-, SFA-), una por archivo en
`specifications/`, más los apéndices en `appendices/` y las figuras en
`figures/`. `index.json` es el catálogo —designación, título, páginas PDF y
folios impresos— y `meta.json` la procedencia.

**El contenido de cada tabla NO viene como `<table>`.** Los bloques `Table`
tienen el `html` vacío y las filas cuelgan de ellos como bloques `Table` →
`Line`, cada uno con su texto y su `bbox`. Una fila llega así:

```
/page/224/Table/27  →  "Carbon, max 0.25A 0.30B 0.35B"
                       "Manganese 0.27–0.93 0.29–1.06 0.29–1.06"
```

Es decir: el dato está entero, pero **construir una base exige recomponer las
columnas desde los `Line` y sus `bbox`**, no leer una tabla ya formada. Cuando
una fila reparte sus celdas en varios `Line` (Tabla 2 de SA-106), el `bbox` es
lo único que dice a qué columna va cada una.

`verificar_seccion_ii.py` audita las seis partes contra sus PDF: que cada
especificación empiece donde dice, que el folio impreso cuadre y que no haya
solapes ni huecos. Los PDF no se versionan; se pasan con `--pdfs`.

## Motor de cálculo — estado actual

Entregable vigente: `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm`
(70 hojas, **una sola visible**, 11 MB). Se **genera por script**, nunca se edita
a mano. El build entero tarda ~45 s.

**Es un libro con macros.** Al abrirlo se ve solo el `Dashboard`; la navegación la
hace un proyecto VBA de dos componentes, y alcanza **36 hojas**: los doce motores
(Art. 212, los 10 buscadores, `Instrucciones`), las nueve hojas de datos de la
Sección II y las **quince hojas `NAV_*`** del árbol de navegación. Los estados de
visibilidad van **grabados en el archivo**, así que con las macros bloqueadas no se
expone ninguna base que alimente un motor.

**Primero el tipo de artefacto, y solo después la norma.** El `Dashboard` conserva
sus tres bandas —`1 · MOTORES DE CÁLCULO`, `2 · MOTORES DE BÚSQUEDA`, `3 · BASES DE
DATOS`—, que responden a la primera pregunta de quien abre el libro: *qué quiero
hacer*. **Dentro** de cada banda empieza la categorización, con el camino con el que
se **cita** una norma —`PUBLICANTE > DISCIPLINA > CÓDIGO DE LA DISCIPLINA > STANDARD
CONCRETO`— y cada tarjeta abre el nivel siguiente:

```
Dashboard
├ 1 · MOTORES DE CALCULO
│    ASME → REPARACIONES ───→ PCC ─→ PCC-2 ──→ Art. 212
├ 2 · MOTORES DE BUSQUEDA
│    ASME → PIPING ─────────→ B31 ─→ B31.3 ─→ 5 buscadores del B31.3
│         → PRESSURE VESSELS → BPVC → SEC. II → 5 buscadores de la Parte D
├ 3 · BASES DE DATOS
│    ASME → PRESSURE VESSELS → BPVC → SEC. II → 9 hojas de las Partes A, B y C
├ 4 · TRANSVERSAL A TODA NORMA  →  MANUAL DE USO
└ 5 · ESTADO DEL LIBRO (KPI) y pie de responsabilidad
```

**Cada banda tiene su propia rama completa, y por eso `ASME` aparece tres veces.**
No es duplicación accidental: la rama del BPVC que lleva a los cinco buscadores de
la Parte D **no** es la que lleva a las nueve hojas de datos de las Partes A, B y C.
La separación por tipo atraviesa todo el recorrido y en cada pantalla todo lo que se
ve es del mismo tipo; un motor nunca se cruza con una hoja de datos. Las hojas
llevan el prefijo de su banda: `NAV_CAL_*`, `NAV_BUS_*`, `NAV_DAT_*`. Lo que sí se
comparte es el **código** que construye la cascada repetida: `_rama_bpvc()` la emite
una sola vez para las dos bandas, porque los tres niveles de arriba son literalmente
la misma cita y escribirlos dos veces era garantizar que un día divergieran.

El árbol se declara **una sola vez**, en `ARBOL` de `build_db_materiales.py`, y de
él se derivan en preorden `HOJAS_NAV`, `DESTINOS`, `NAVEGABLES`, `PADRE`, `ROTULO`
y `ANCLA_VOLVER`. Añadir el B31.1 mañana es añadir un `Nodo`: no toca el
constructor de hojas ni el mecanismo de navegación.

Lo que la norma publica y este libro **no** carga (B31.1, B16, SEC. VIII, PCC-1,
PCC-3, el resto de artículos de PCC-2) aparece como **tarjeta marcador**: gris, sin
hipervínculo ni clave, rotulada `NO CARGADO EN ESTE LIBRO`. Es solo un rótulo de
documento —cero dato normativo, no roza la Regla nº 1—; sirve para que un nivel con
un solo hijo cargado siga explicando la taxonomía y para ver de un vistazo qué falta.

**Lo que está cargado pero en otra banda se dice con texto, no con un marcador.**
`NAV_BUS_SEC_II` advierte al pie que las Partes A, B y C no llevan buscador y que su
volcado está en la banda 3. Un marcador ahí sería **falso** —sí están cargadas— y
enlazarlas cruzaría las dos bandas, que es justo lo que esta separación evita.

Los KPI y el pie de responsabilidad se quedan en el `Dashboard`, que sigue siendo la
única portada. `Instrucciones` cuelga de la raíz **y además** tiene botón fijo
`? MANUAL DE USO` en la cabecera de toda hoja `NAV_*`: es transversal a las normas y
debe estar siempre a un clic.

**Las nueve de la Sección II sí son navegables, y es una excepción declarada.**
La regla —una base de datos que alimenta un motor es insumo auditado, no interfaz:
abrirla dejaría leer un valor sin la cascada, sin el bloqueo por rango y sin la
nota del material— no aplica ahí, porque en esas nueve el entregable **es** la
hoja de datos: no alimentan ningún motor y no llevan una sola fórmula. Ocultarlas
no protegería nada y solo las haría inútiles. `verificar.py` §8 y
`test_dashboard.py` comprueban la regla **y** su excepción, por nombre.

### Los diez buscadores

| Hoja | Qué responde | Fuente |
|---|---|---|
| `Buscar_B31_3` | S admisible | B31.3 Apéndice A, A-1/A-4 y A-1C/A-4C |
| `Buscar_BPVC_IID` | S admisible, ferrosos | II-D Tabla 1A |
| `Buscar_BPVC_IID_B` | S admisible, no ferrosos y pernería | II-D Tablas 1B y 3 |
| `Buscar_Su` · `Buscar_Sy` | Su y Sy frente a T | II-D Tablas U y Y-1 |
| `Buscar_Prop_IID` | Módulo E, dilatación (Coef. B) y Poisson/densidad **por grupo** | II-D TM-1..5, TE-1..5 y PRD |
| `Buscar_Prop_B31_3` | Dilatación y módulo, metales y no metálicos | B31.3 Apéndice C, C-1..C-4 |
| `Buscar_B31_B1` | Esfuerzo de diseño hidrostático **HDS** frente a T | B31.3 Apéndice B, B-1 y B-1C |
| `Buscar_Ec_A2` | Factor de calidad de fundición **Ec** | B31.3 Tabla A-2 + 302.3.3-1 |
| `Buscar_Ej_A3` | Factor de calidad de junta longitudinal **Ej** | B31.3 Tabla A-3 |

**Un dato, un motor.** El Apéndice C salió de `Buscar_Propiedades` (que pasó a
llamarse `Buscar_Prop_IID`) y de `Buscar_NoMetalicos`, y tiene motor propio con
conmutador SI ↔ US. Por la misma razón, en la Rev. 4d salió la Tabla B-1 —ver
abajo—, y con ella desapareció `Buscar_NoMetalicos`: lo que le quedaba (B-2 a
B-6) se retiró del libro por alcance.

**Todo buscador que reúna más de una tabla dice cuál resolvió la cascada y qué
publica esa tabla.** `Buscar_B31_3` monta A-1 (esfuerzos básicos en tracción de
los metales) junto a A-4 (esfuerzos de diseño de la **pernería**), y los dos
admisibles se leen igual sin serlo: confundirlos cambia el cálculo. La ficha
lleva la fila `Funcion de la tabla` justo debajo de `Tabla del codigo`, y el
subtítulo de la hoja lo dice antes de que se elija nada. El texto vive en
`FUNCION_TABLA` del builder —es una **descripción del título impreso**, no un
dato normativo— y `formula_funcion_tabla()` emite solo las ramas de las tablas
que esa base contiene de verdad, leídas de la propia base; una tabla nueva sin
descripción **aborta el build** en vez de imprimir un rótulo genérico. Las
ediciones US comparten descripción con su gemela métrica: A-1C no es otra tabla,
es la misma en otras unidades.

### `Buscar_B31_B1` — el HDS del Apéndice B (Rev. 4d)

**Del Apéndice B, este libro carga solo la Tabla B-1 / B-1C.** Las Tablas B-2 y
B-3 (listados de especificación de RTR y RPM) y B-4, B-5 y B-6 (presiones
admisibles de concreto, vidrio borosilicato y PEX-AL-PEX) existían como
`DB_NoMetalicos` + `Buscar_NoMetalicos` y **se retiraron en la Rev. 4d por
decisión de alcance**: son tablas que este trabajo no usa. La extracción sigue
intacta en `resources/.../appendix_b/table_b_2..b_6.json` —que es la fuente de
verdad; lo que se quitó es la carga al libro— y en el árbol quedan como **tarjeta
marcador** `NO CARGADO EN ESTE LIBRO`, igual que el B31.1 o la Sección VIII.
`verificar.py` §3b comprueba que no vuelvan por descuido y `test_dashboard.py`
que ninguna de las dos hojas exista.

La Tabla B-1 (B-1C en US) es lo **único** del Apéndice B tabulado frente a la
temperatura: las demás publican listados de especificación o una presión
admisible puntual. Vivía dentro de `Buscar_NoMetalicos`, que es una ficha
campo/valor sin temperatura de consulta y cuya clave de selección era **solo la
designación de material**. Como el código publica
`PE2708` bajo D2737, D3035 y F714 con HDS distinto, **17 de las 36 filas eran
inalcanzables** y la ficha mezclaba campos de varias especificaciones. La
cascada es ahora material → Spec. No. → designación de tubería → Cell Class →
variante, y `test_dashboard.py::TestTablaB1` fija que las 36 filas de cada
edición tengan clave propia.

Conmutador SI ↔ US de **hoja** (`DB_B31_B1` ↔ `DB_B31_B1C`), nunca conversión:
el código publica una tabla por edición. El enlace entre ellas es **posicional**
(`clave_bi = B1#fila impresa`) y `verificar_paridad_b1()` es su contrapartida
obligatoria: aborta si la identidad —designación de material + Spec. No.— no
casa fila a fila. Lo que sí diverge se declara y no se toca: F2389 imprime `PR`
como designación de tubería en B-1 e `IPS Sch. 80` en B-1C, y 110 °C frente a
210 °F de máxima (regla 9; está en `observaciones_del_codigo`).

**Las tres reglas de rango son otras, y son del Capítulo VII, no de las tablas de
metales.** Están citadas una a una en el código y en la sección 4 del motor:

- **Se interpola.** `para. A302.3.1(b)`: *«Straight-line interpolation between
  temperatures is permissible.»* Con conmutador `Interpolado` /
  `Tabulado-conservador`, igual que los cinco de esfuerzos.
- **Por debajo de la primera temperatura tabulada NO se extrapola: se sostiene.**
  Nota (3) de la propia tabla, anclada a la columna de 23 °C (73 °F): *«Use these
  hydrostatic design stress (HDS) values at all lower temperatures.»* Concuerda
  con `para. A323.2.2(b)`. Es la **única excepción declarada** a la regla 4 del
  proyecto, y la escribe el código.
- **Se bloquea por arriba en dos sitios distintos**, con aviso distinto para cada
  uno: el **límite máximo recomendado** que imprime la fila (Notas (1) y (2),
  `para. A323.2.1(a)`) y el **último punto tabulado**. El orden importa y no es
  cosmético: F441/CPVC4120-05 imprime máxima 93,3 °C y su último HDS a 82 °C, así
  que 90 °C y 95 °C tienen que decir cosas distintas —«el código no publica el
  dato» no es «el material no se recomienda ahí»—. `formula_estado_b1()` comprueba
  primero los límites recomendados y después la banda tabulada; `verificar.py`
  §6d recalcula en Excel ese par exacto y `test_dashboard.py` fija el orden.

El límite **por abajo** es el mínimo recomendado impreso, no el primer punto
tabulado: entre uno y otro la Nota (3) sigue dando un valor válido (PEX0006 tiene
mínima −50 °C y su primer HDS a 23 °C).

Una fila puede no publicar HDS a ninguna temperatura —el ABS solo trae límites de
temperatura— y el estado lo dice: `SIN HDS TABULADO`, no un cero.

El HDS es el `S` de la eq. (26a) del `para. A304.1.2`, `t = PD/(2S+P)`. El pie del
motor transcribe el aviso del `para. A302.3.1(a)`: usar el HDS para cálculos
distintos del diseño a presión **no está verificado**.

`formula_estado_b1` y `formula_valor_b1` viven en `build_db_materiales.py` y las
emiten el motor **y** `verificar.py` §6d: la prueba ejerce el original, no una
copia. Once casos recalculados en Excel real, más la conducción de la cascada
completa (23 → 13,8 MPa; 60 → 7,375 interpolado y 3,45 tabulado-conservador;
95 → BLOQUEADO por límite; 40 en PVC1120 → BLOQUEADO por banda; 0 en PEX0006 →
4,34 por Nota (3); 140 °F en US → 1,07 ksi leído de B-1C).

**`Buscar_Prop_IID` expone la dilatación de la II-D (Rev. 4b), pero solo el
Coeficiente B.** TE-1..5 publican tres coeficientes por grupo —A (instantáneo),
B (medio, de 20 °C a T) y C (expansión acumulada)— indexados por temperatura ×
columnas de grupo, no por material (`DB_TE`/`DB_TEC` conservan esa banda tal
como está impresa, para auditar). El motor solo pivota y expone **B**: es el
que se usa en cálculo de dilatación/flexibilidad, la misma magnitud que "alfa"
en la Tabla C-1 del Apéndice C del B31.3. Los coeficientes A y C **no
alimentan el buscador** —siguen impresos tal cual en `DB_TE`/`DB_TEC`—: decisión
de alcance declarada, no un hueco silencioso.

`DB_TE_G`/`DB_TE_GC` son la base pivotada (grupo en filas, temperatura en
columnas) que sostiene el bloque nuevo, construida por
`build_dilatacion_grupo()`. El rótulo de cada fila —`etiqueta_columna_b_te1()`—
tiene que ser idéntico al que `build_map_grupo()` escribe en `Grupo dilatación
(TE)` de `MAP_Grupo`/`MAP_GrupoC`, letra a letra, o la selección en el
buscador no encontraría la fila que el ingeniero fue a buscar; `_clave_fila_te1()`
traduce ese rótulo impreso a la clave snake_case real de `rows` (las filas de
TE-1..5 son dispersas: una columna sin dato en esa temperatura no trae su
clave, así que no vale alinear por posición). `verificar.py` (`audita_grupo_te`)
compara, fila a fila y usando esas mismas dos funciones —nunca una copia—, el
vector de la hoja contra el JSON de TE-1..5.

**Tabla TE-2 (aleaciones de aluminio) queda fuera, declarada — verificado en
las dos ediciones, no solo en la SI.** `table_te_2.json` de
`bpvc_ii_d_metric_2025` y de `bpvc_ii_d_customary_2025` imprimen exactamente
las mismas cuatro columnas: `Temperature` + "A", "B", "C" sueltas, sin
rótulo de grupo ni de aleación en el título de columna (a diferencia de TE-1,
TE-3, TE-4 y TE-5, que sí nombran el grupo o la designación en cada columna).
Es la única de las cinco tablas TE con un solo grupo para toda la tabla, y no
hay con qué distinguir esa fila sin inventar un nombre que el código no
imprime — exactamente la heurística que la Rev. 3 prohíbe (ver "No
reintroducir una heurística de composición" más abajo). Por eso
`etiqueta_columna_b_te1()` devuelve `None` para las tres columnas y el
builder lo declara en `ISSUES` en vez de omitirlo en silencio.
Consecuencias, explícitas: las aleaciones de aluminio (i) no tienen fila de
dilatación en `Buscar_Prop_IID` (el bloque TE del buscador no puede ofrecer
un grupo que no existe), (ii) tampoco tienen `Grupo dilatación (TE)` en
`MAP_Grupo`/`MAP_GrupoC` — mismo patrón de columna bloquea a
`columnas_nombradas_te1()` — y (iii) sí conservan el dato íntegro, tal como
está impreso, en `DB_TE`/`DB_TEC` (regla 9): no se pierde, solo no alimenta
este buscador. No hay una vía de extracción alternativa que resuelva esto:
el código mismo no imprime el dato que identificaría la fila, así que
cerrarlo exigiría una convención inventada por fuera del texto normativo —
lo que la Regla nº 1 de este proyecto prohíbe. Queda cerrado como límite de
la fuente, no como pendiente de ingeniería.

**`Buscar_Prop_IID` lleva conmutador SI ↔ US (Regla 10), cerrado tras Rev. 4b.**
Hasta esa revisión la hoja era SI-only —sin la celda "Sistema de unidades"—
y el bloque de dilatación solo consumía `DB_TE_G`; `DB_TE_GC` (edición US) se
construía y auditaba igual que `DB_EC`, por paridad, pero ningún motor la
consultaba. Los tres bloques (módulo E, dilatación, Poisson/densidad) son
una tabla por edición, no columnas de una tabla compartida, así que el
conmutador cambia de **hoja** (`DB_E`/`DB_EC`, `DB_TE_G`/`DB_TE_GC`,
`DB_PRD`/`DB_PRDC`), igual que en `Buscar_Prop_B31_3`, nunca convierte. El
rótulo de GRUPO («Material Group A», «Group 1») no es un valor físico y el
código lo imprime idéntico en las dos ediciones, así que la lista de grupos
por tabla se lee siempre de la edición SI: lo único que cambia con el
selector es de qué hoja se lee el valor para ese grupo. Un matiz que **no**
es una conversión encubierta: el origen del Coeficiente B de dilatación no es
la misma temperatura en las dos ediciones — 20 °C en la SI, **70 °F** en la
US (primer punto impreso de TE-1 US, no 20 °F) — porque son extracciones
independientes de lo que cada edición imprime (Regla 9), no una traducción de
unidades del mismo origen. Validado en Excel real (no solo por `verificar.py`,
que no inspecciona el layout de celdas de este buscador): módulo E de
`Material Group A [Note (1)]` en Tabla TM-1 da 200×10³ MPa a 25 °C en SI y
29,26×10⁶ psi a 25 °F en US (fila US interpola entre -100 °F y 70 °F, que es
la banda que esa edición imprime); Poisson/densidad de `A02040` da
2800 kg/m³ en SI y 0,101 lb/in³ en US.

### Reconstruir

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts

# Solo la primera vez, o al tocar vba/*.vba.
# Activa "Confiar en el acceso al modelo de objetos de proyectos de VBA" unos
# segundos y lo restaura, verificando el resultado. Si avisa de que no pudo
# restaurarlo, desactivelo a mano en el Centro de confianza.
#
# El script restaura el valor que ENCUENTRA, no el seguro: si una corrida
# anterior murio sin restaurarlo, lo deja encendido y lo advierte por stderr.
# Comprobarlo despues, y apagarlo si quedo a 1:
#   Get-ItemProperty HKCU:\Software\Microsoft\Office\16.0\Excel\Security AccessVBOM
#   Set-ItemProperty HKCU:\Software\Microsoft\Office\16.0\Excel\Security AccessVBOM 0
python make_vba_seed.py

# Solo si se repone la extraccion de II-D: notas de grupo de TM-1 / TE-1.
# Una corrida por edicion; nunca copiar las notas de una en la otra.
python extraer_notas_ii_d.py --edicion si --resources ..\..\..\resources `
    --pdf "C:\Users\dvelasquez\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES\SECCION II\D Metric 2025\D Metric 2025 _p1201-p1500.pdf"
python extraer_notas_ii_d.py --edicion us --resources ..\..\..\resources `
    --pdf "C:\Users\dvelasquez\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES\SECCION II\D Customary 2025\D Customary 2025 _p1201-p1500.pdf"

# Solo si se repone la extraccion de los Apendices B o C del B31.3. Los dos son
# idempotentes y NO tocan ningun valor: solo la capa de metadatos. El PDF del
# codigo y su OCR NO estan en el repo (copyright de ASME, ver .gitignore): el de
# la C hay que tenerlo en disco y pasarlo con --ocr si no esta en su sitio.
#   C: unidades de los coeficientes A y B de C-1/C-1C, miembros de sus Notas
#      (2)..(6) y grupo impreso de C-2/C-3.
#   B: celdas fusionadas de B-1, las 6 especificaciones de B-3, las columnas
#      minimo/maximo de B-4 y B-5, y el material a dos lineas de B-6.
#   A: la costura de las paginas enfrentadas — devuelve a la curva los 197
#      esfuerzos a 200 F del bloque de niquel de A-1C, separa Class/Description
#      en A-3 y Spec/Grade en A-1C, y limpia la elipsis de columna vacia.
python completar_apendice_c.py --resources ..\..\..\resources
python completar_apendice_b.py --resources ..\..\..\resources
python completar_apendice_a.py --resources ..\..\..\resources

# Solo si se repone la extraccion de las tablas del cuerpo del B31.3. Los tres
# son idempotentes y solo escriben metadatos (302_3_4 tambien columns/rows de
# table_302_3_4_1.json, pero solo una vez: la segunda corrida no toca nada).
#   302_3_3: declara el canonico del par de la Tabla 302.3.3-1 y recupera sus
#            rotulos de columna del para. 302.3.3(c).
#   302_3_4: reconstruye el CUERPO de la Tabla 302.3.4-1 (Ej), que la
#            extraccion original colapso dentro de los encabezados de columna.
#            Exige el folio impreso -PDF pagina 51 del codigo- porque no esta
#            en el repo (copyright ASME); el ingeniero lo aporta con --pdf. El
#            builder ABORTA si no se ha corrido (o si el cuerpo vuelve a
#            colapsarse): Buscar_Ej_A3 no se construye con un dato que no esta.
#   canonicas: resuelve los 32 pares de doble prefijo de CHAPTERS/tables. Si
#            algun par deja de encajar en una de las clases mecanicas o
#            declaradas EXTERNA, no escribe nada y lo dice.
python completar_tabla_302_3_3.py --resources ..\..\..\resources
python completar_tabla_302_3_4.py --resources ..\..\..\resources --pdf "<PDF con el folio impreso de la Tabla 302.3.4-1>"
python declarar_tablas_canonicas.py --resources ..\..\..\resources

python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm

# Opcional. Las tablas de la Seccion II YA entran al libro (el builder importa
# secii_tablas.py como libreria); este CLI no escribe nada: solo mide y refresca
# Revision_Tablas_SecII.md. Ver la seccion de abajo.
python secii_tablas.py --resources ..\..\..\resources `
    --informe ..\Revision_Tablas_SecII.md
```

`verificar.py` tarda varios minutos: su §10 relee las 54 198 filas de la Sección
II desde la hoja y las contrasta contra una reconstrucción independiente hecha
desde `resources/`.

La entrada del builder es el maestro sembrado en `templates/`, que a su vez sale de
`outputs/Motor_de_Calculo_ASME_PCC_Rev0_respaldo.xlsx` (3 hojas). El
`Motor_de_Calculo_ASME_PCC.xlsx` de la raíz **no está en el repo**.

Ese respaldo estuvo en `outputs/Base_Datos_Materiales_ASME/` y se movió a
`outputs/`; `make_vba_seed.py` lo busca **en los dos sitios** y solo pide `--in`
si no está en ninguno. Es un paso que se corre de tarde en tarde, y fijar una
sola ruta lo dejaba roto —con un «no existe la entrada»— cada vez que el archivo
cambiaba de carpeta.

`verificar.py` devuelve 0 solo si todo pasa. Audita **fila a fila** cada valor
tabulado contra el JSON del código (271 536 valores de esfuerzos, más 5 708 de
módulo E y dilatación de la II-D, 4 726 del Apéndice C, 720 de la Tabla B-1
—incluidos los dos límites de temperatura recomendados de cada fila, que son lo
que bloquea el resultado— y 743 de los factores de calidad), la contigüidad de la cascada,
la ausencia de fórmulas de matriz dinámica, la interpolación recalculada en hoja, el
caso semilla y la capa de navegación. **Requiere Excel instalado**: recalcula con el
motor real, no con LibreOffice. **Ejecútalo siempre después de tocar el builder.**

Sus §6b, §6c y §6d **recalculan en Excel la misma expresión que lleva el motor**, no
una copia: las funciones que la generan (`formula_estado_apxc`, `formula_valor_apxc`,
`formula_factor_aplicable`, `formula_estado_b1`, `formula_valor_b1`) viven en
`build_db_materiales.py` y las emiten los dos.
Si alguien cambia la lógica en el motor, la prueba la ejerce cambiada; si cambia solo
el texto de un estado, el literal se lee de la misma constante y no puede divergir.

### Capa de navegación — dos fuentes de verdad que no pueden divergir

La tabla de navegación vive por duplicado: en `build_db_materiales.py`, que la graba en
el archivo, y en `scripts/vba/mod_nav.vba`, que la reaplica al abrir. Si se separan, el
libro se abre mostrando algo distinto de lo que se construyó. `test_dashboard.py` lo
comprueba (`TestSincroniaPythonVba`). Al tocar una, tocar la otra:

| Concepto | Python | VBA |
|---|---|---|
| Hojas navegables (36, **en preorden**) | `NAVEGABLES` | `HojasNavegables()` |
| Columna base de claves | `COL_CLAVE_BASE = 66` | `COL_CLAVE_BASE` |
| Celda del aviso | `FILA_AVISO = 4` | `CELDA_AVISO = "A4"` |
| Texto y color del aviso | `build_dashboard` + `AVISO_ROJO_*` | `TXT_INACTIVAS` · `TXT_ACTIVAS` y sus `RGB(...)` |

El aviso entra en la tabla porque el VBA **reescribe A4 al abrir**: si su texto o
su color no son los que grabó el builder, el libro cambia de aspecto en cuanto se
abre. Los `RGB(...)` de `mod_nav.vba` son `TINTA`/`VERDE` y `PAPEL`/`ROJO`.

`HojasNavegables()` es el **único** punto del VBA que crece con el árbol.

**La clave es siempre el destino.** Antes había dos clases de clave: el nombre de la
hoja a abrir y el literal `"VOLVER"`, que el VBA resolvía siempre al Dashboard. Con
cinco niveles eso deja de servir: subir tiene que llevar al **padre**. Ahora la celda
oculta guarda **siempre el nombre de la hoja destino**, se esté bajando, subiendo o
saltando por la miga de pan, y `PADRE` —derivado del árbol— es la fuente del botón de
retorno. `CLAVE_VOLVER` y `VolverAlDashboard` desaparecieron y `AbrirHoja` pasó a ser
`IrAHoja(destino, origen)`: **una sola rama**, que ya no crece con el árbol. El
guardarraíl «el origen debe ser el Dashboard» se retiró porque era redundante
—`EsNavegable` ya impide destapar una `DB_*` o una `MAP_*`— y con el árbol el origen
legítimo dejó de ser una sola hoja.

**La tarjeta entera es clicable**, no solo la barra inferior: sus cuatro filas llevan
hipervínculo y clave propia, en filas distintas de la misma columna, así que las claves
siguen sin pisarse.

Tres trampas ya pagadas, documentadas en el código:

- **El VBA referencia hojas por `.Name`, nunca por CodeName.** openpyxl no asigna
  `codeName` a las hojas que crea; Excel se los inventa al abrir.
- **Toda declaración de módulo (`Const`, `Dim`, `Type`) precede a la primera rutina.**
  Si no, VBA reporta «Variable not defined» en cada uso, Excel abre un diálogo modal al
  compilar durante `SaveAs`, y la automatización se cuelga sin mensaje. `make_vba_seed.py`
  lo comprueba con `lint_vba()` antes de tocar COM.
- **VBA no admite más de 25 continuaciones de línea (`_`) en una línea lógica.** Con 31
  hojas, el `Array( _ … )` de `HojasNavegables()` necesitaba 30 —hoy son 36— y `AddFromString` lo
  rechazó: el módulo quedó **vacío**, y el síntoma visible no fue ese sino un «No se ha
  definido Sub o Function» al guardar —ThisWorkbook llamaba a rutinas que ya no
  existían— dentro de un diálogo modal que colgó Excel y dejó el maestro borrado. La
  lista se arma ahora **concatenando** (`s = s & "|…"` + `Split`), que no tiene tope y
  deja cada hoja en su línea; `lint_vba()` comprueba el límite antes de tocar COM.

### Reglas de diseño del libro — no romper

1. **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`,
   `XLOOKUP`, `VSTACK`, `_xlfn`. Solo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`.
   **Enmienda Rev. 3:** el entregable es `.xlsm` y ya **no abre en Google Sheets** —
   la capa de navegación es VBA. Lo que se pierde es solo la navegación. La
   restricción sobre las fórmulas **se mantiene entera**: todo el cálculo debe seguir
   siendo portable, y `verificar.py` sigue fallando si aparece una matriz dinámica.
   Corolario aprendido: **un texto no se guarda como fórmula.** Un `="texto…"` de más
   de 255 caracteres lo parte Excel en `_xlfn._LONGTEXT(...)` al reguardar, y la celda
   pasa a mostrar `#NAME?` fuera de Excel 365. El builder lo normaliza
   (`normalizar_textos_como_formula`).
2. **Validación de datos: solo rango literal o lista de ítems.** Nunca una fórmula
   (`OFFSET`, `INDIRECT`) como origen. Las listas dependientes se materializan en
   celdas de columnas ocultas y la validación apunta a ese rango.
3. **Banda compacta.** Las tablas ASME dejan huecos interiores (la A-1 no imprime
   125 °C para A106 Gr.B). Junto a la banda impresa —que se conserva para auditar—
   cada base lleva una banda con solo los puntos existentes, sin huecos, y la
   consulta trabaja sobre ella.
4. **No extrapolar.** Por encima de la última temperatura tabulada o de la Temp. máx.
   del material, dictamen "FUERA DE RANGO" y resultado bloqueado. Lo prohíbe el código.
5. **Cascada contigua.** Cada base se ordena por familia → composición → forma →
   especificación → grado, para que cada nivel sea un bloque contiguo.
6. **Unidades visibles** en toda variable, indicadores y ejes de gráficas.
7. **`material_id` único**, desambiguado en este orden: notas del código → Rm → Re →
   número de línea impreso. El código repite spec+grado con admisibles distintos.
8. La **familia de material** es una agrupación de navegación derivada del UNS y la
   composición impresa. **No es dato normativo** y no entra en ningún cálculo.
9. Los valores se cargan **tal como están impresos**; SI y US son extracciones
   independientes, nunca conversiones.
10. **Conmutador de unidades en todo motor.** Todo motor de búsqueda y todo motor
    de cálculo del libro lleva un conmutador **SI ↔ US**. Es corolario de la
    regla 9: los dos sistemas se leen de la edición correspondiente del código,
    nunca por conversión. Cuando el código publica los dos sistemas en la **misma
    tabla** —C-2 y C-4 del B31.3— el conmutador cambia la **columna** leída;
    cuando publica **una tabla por edición** —A-1/A-1C, C-1/C-1C, C-3/C-3C, II-D
    métrica y US— cambia la **hoja**. Si un material no tiene homólogo en la otra
    edición, el motor muestra «sin equivalente en la edicion US», nunca un valor
    convertido.

    **Excepción declarada (Rev. 4): los factores adimensionales.** `Ec` (Tabla A-2)
    y `Ej` (Tabla A-3) no tienen unidad y el código publica **una sola tabla** para
    los dos sistemas: no existen A-2C ni A-3C. Ahí `Buscar_Ec_A2` y `Buscar_Ej_A3`
    ponen una **celda fija y rotulada** —«FACTOR ADIMENSIONAL — identico en SI y en
    US (el codigo publica una sola tabla)»— en el sitio donde los demás ponen el
    conmutador. Se cumple el espíritu de la regla —el usuario ve siempre en qué
    sistema lee y nunca un valor convertido— sin fabricar un interruptor que no
    gobierna nada. `TXT_ADIMENSIONAL` en el builder, con su prueba.

11. **El factor publicado es un MÍNIMO, y el motor tiene que decirlo.** Las Notas
    (4) y (5) de la Tabla A-2 dicen cosas opuestas con el mismo formato: la (4)
    —*«can be enhanced by supplementary examination»*— permite subir el factor con
    la Tabla 302.3.3-1; la (5) —*«applicable only when proper supplementary
    examination has been performed»*— avisa de que el factor **ya supone** ese
    examen. Confundirlas mueve el espesor requerido. `Buscar_Ec_A2` deriva el
    estado de las notas que cita **la propia fila**, nunca de una suposición, y
    contempla los cuatro casos: solo (4), solo (5), **las dos** —A451 es la única
    fila que las cita juntas— y ninguna. Para `Ej` el mecanismo existe
    (para. 302.3.4(b) y Tabla 302.3.4-1). La extracción de esa tabla en
    `resources/` estaba inservible —el cuerpo se había colapsado dentro de los
    encabezados de columna— y `completar_tabla_302_3_4.py` la reconstruyó desde
    el folio impreso (aportado por el ingeniero fuera del repo, copyright ASME;
    su SHA-256 queda en `extraction_amendments` de `table_302_3_4_1.json` para
    auditoría). A diferencia de Ec, el B31.3 **no imprime una correspondencia
    fila a fila** entre la Tabla A-3 y la 302.3.4-1: el factor lo decide el tipo
    de junta/costura/examen, no la especificación de material, así que
    `Buscar_Ej_A3` no infiere esa correspondencia —sería exactamente la clase de
    heurística que la Rev. 3 eliminó— y en su lugar transcribe las diez filas
    íntegras de la 302.3.4-1 para que el ingeniero identifique la suya y lea su
    Ej directamente. La Nota (1) —prohíbe incrementar Ej en las juntas 1 y 2— se
    marca en cada una de esas dos filas. El par `table_302_3_4_1.json` /
    `table_table_302_3_4_1.json` queda excluido del criterio mecánico de
    `declarar_tablas_canonicas.py` (clase `RECONSTRUIDA_DEL_FOLIO_IMPRESO`): su
    canónico ya no es comparable columna a columna con el gemelo a propósito.

12. **Todo material se selecciona de una base de datos, nunca de una lista fija.**
    Cualquier campo de un motor de cálculo que pida un material se resuelve con una
    cascada de listas desplegables materializada contra la base de datos que
    corresponda según el código ASME PCC aplicable —`DB_B31_3`, `DB_BPVC_IID`/
    `DB_BPVC_IID_B`, u otra hoja de datos del libro—, igual que ya hacen los motores
    de búsqueda y la Sección 7 de resolución de material de los motores de cálculo
    (Art. 212 y Art. 206). **Nunca** una lista de texto tecleada directo en el
    `formula1` de una `DataValidation` ni en un valor por defecto: una lista así no
    se audita, no se actualiza si la base cambia, y puede ofrecer un material que la
    base ni siquiera admite. Es corolario de la regla 2 (mecanismo de validación)
    aplicado a la fuente: la validación puede ser "lista de ítems", pero esos ítems
    tienen que venir de una base de datos real, no de un literal.

    **Violación detectada y pendiente de corrección (2026-09-10):** en
    `Parche_PCC2_Art212`, D22/D23 ("Material de tubería / envolvente" y "Material
    del collar / parche") usan hoy una lista fija de 6 materiales tecleada en
    `build_db_materiales.py` (función `build_parche_art212`, bloque de
    `DataValidation` con `formula1='"A106 Gr.B,A516 Gr.70,A105,A285 Gr.C,A333
    Gr.6,A53 Gr.B"'`), declarada puramente descriptiva —el material que sí
    alimenta el cálculo lo resuelve la cascada de la Sección 7—. Ese campo
    descriptivo debe pasar también a leer de la base de datos correspondiente, no
    quedarse como excepción.

13. **Los motores de búsqueda y los motores de cálculo no se conectan entre sí.**
    Ningún motor de cálculo lee una celda de un motor de búsqueda, ni un motor de
    búsqueda lee una celda de un motor de cálculo. Los dos leen **exclusivamente**
    de las bases de datos (`DB_*`) del libro. Esto evita acoplar el resultado de un
    motor a que otro exista, esté abierto, o esté en el estado correcto, y mantiene
    cada motor auditable de forma independiente contra su propia fuente.

14. **Toda variable de entrada que exista en una base de datos se elige de una lista
    desplegable; nunca se teclea.** Diámetro nominal, cédula, espesor, material — si
    el valor está tabulado en una base del libro, el motor lo ofrece en un
    desplegable alimentado por esa base (regla 12) y el ingeniero no lo escribe a
    mano. El objetivo es eliminar el error de tecleo como clase de fallo, no solo
    documentar de dónde sale el dato.

    Lo que **sí** se teclea es lo que ninguna base publica: datos de proceso y de
    campo —temperatura y presión de operación, dimensiones del defecto, sobreespesor
    de corrosión, medidas de la reparación—. La línea es: **¿el libro tiene una tabla
    con este valor? → desplegable. ¿Es un dato del servicio o de la inspección? →
    entrada manual.**

### Sistema visual — Swiss Industrial Print (Rev. 4e)

**Las 70 hojas van en un solo sistema, y esa es toda la regla.** Antes había dos:
el azul corporativo de la Rev. 0 en las tres hojas que trae el maestro y el del
builder en las otras 67. Ahora hay uno: papel de documentación sin blanquear,
tinta carbón y **un** acento rojo.

Vive en un bloque único de tokens al principio de `build_db_materiales.py`. Nada
de color ni de fuente se escribe suelto: `verificar.py` no lo mira, pero
`test_dashboard.py::TestSistemaVisual` recorre los 2,4 millones de celdas con
formato del libro construido y falla si aparece un color fuera de la paleta o una
fuente que no sea una de las dos.

| Token | Valor | Para qué |
|---|---|---|
| `PAPEL` · `PAPEL_2` | `F4F4F0` · `EAE8E3` | sustrato · compartimento (tarjeta, KPI, campo) |
| `TINTA` · `TINTA_2` | `050505` · `111111` | bloque estructural y texto · trazo de curva |
| `ROJO` | `E61919` | **único** acento: aviso, bloqueo, dato vital |
| `GRIS` · `GRIS_2` | `8A8A85` · `C9C7C1` | trama 55 % (metadato, marcador) · 25 % (retícula) |
| `VERDE` · `AMBAR` · `AMBAR_TXT` | `4AF626` · `E6A019` · `8A5D00` | semáforo funcional |
| `MACRO` · `MONO` | Arial Black · Consolas | estructura y cifra de KPI · todo el dato |

Las dos fuentes vienen instaladas con Windows **a propósito**: una que Excel no
encuentra la sustituye en silencio y deshace la retícula, así que aquí no entra
ninguna de descarga (JetBrains Mono, Archivo Black) por bien que encaje.

El **semáforo de tres estados se conserva retonado**, y es la excepción declarada
al acento único: en un libro de cálculo el estado de la consulta es información
de seguridad y se lee sin leerse. Va como bloque macizo con tinta encima, nunca
como pastel de relleno suave, y el rojo mantiene un único significado en todo el
libro: bloqueado.

**El sustrato se graba a nivel de COLUMNA, no celda a celda** (`sustrato()` /
`aplicar_sustrato()`, último paso antes de guardar). Excel admite un estilo por
columna, así que las 54 198 filas del volcado de la Sección II heredan papel y
mono sin un solo estilo de celda: el mismo resultado visual sin multiplicar el
tamaño del `.xlsm` por 2,6 millones de celdas con formato. El archivo pasó de
11,52 a 11,56 MB. Toda celda con estilo propio lo pisa, que es lo que se quiere.

**Tres trampas ya pagadas, comprobadas exportando la hoja y mirándola:**

- **En un rango fusionado el RELLENO se hereda de la celda ancla, pero el BORDE
  no.** La banda de tinta sale entera poniéndosela solo a la ancla; la franja
  roja puesta igual sale como un muñón de una columna. Hay que recorrer el rango
  (`franja()`), y por eso esa función recorre aunque el rango esté fusionado.
- **openpyxl miente en los dos sentidos y no sirve para comprobarlo.** Al leer un
  libro reconstruye el borde de la ancla sobre todo el rango (parece que
  estuviera) y al escribir no propaga nada si el estilo se puso después de
  fusionar. La única evidencia válida es exportar la hoja a PDF/PNG desde Excel
  e ir a mirarla.
- **Un campo de formulario es una LÍNEA, no una caja.** Siete campos seguidos con
  caja completa se ven como una escalera de barrotes, y además el borde de un
  campo compite con el de su vecino en la arista que comparten — que es justo
  donde se perdía el recuadro rojo de la única celda que se teclea. Con la línea
  (`CAJA_CAMPO = Border(bottom=REGLA)`), la caja roja es el único rectángulo
  cerrado de la zona.

**Lo que trae el maestro se retona al construir** (`retonar_heredadas()`), porque
la plantilla no se edita a mano. Es una tabla de equivalencia **exacta**
(`MAPA_RELLENO` / `MAPA_FUENTE`), no una aproximación por cercanía de color: así
es idempotente, no puede tocar una celda que ya nació en el sistema nuevo, y lo
que quede sin traducir se declara en `ISSUES` en vez de quedarse con el aspecto
de la Rev. 0. `TEXTOS_HEREDADOS` cubre el caso aparte: un rótulo del maestro que
**describe** el sistema visual («celdas azules sobre fondo amarillo = editables»)
caduca con él, y retonar la celda sin reescribir el texto dejaría al libro
explicando una convención que ya no existe.

**`Parche_PCC2_Art212` dejó de ser una hoja heredada (2026-09-10).** `HOJAS_HEREDADAS`
pasó de tres hojas (`Instrucciones`, `Datos_Ref`, `Parche_PCC2_Art212`) a solo dos
(`Instrucciones`, `Datos_Ref`); `TEXTOS_HEREDADOS` quedó vacío en consecuencia. La hoja
del Art. 212 nace ahora 100 % en código con `build_parche_art212()`, hermana de
`build_collar_art206()` (mismo patrón: `new_sheet()`, helpers `lab`/`inp`/`calc`/`band`/
`header`, sin herencia del maestro) — retiró la última hoja de cálculo que todavía
dependía de `retonar_heredadas()` para su sistema visual. Plan y commits:
`outputs/plans/plan_desanclado_total_motor_art212.md`, rango `40fc235..3653ccb`.

### Estilo de diseño de los buscadores — no romper

Vive en `build_buscador` / `finish_buscador` (los 5 buscadores de cascada: B31_3,
BPVC_IID, BPVC_IID_B, Su, Sy) y se replica en `build_buscador_grupo`. Tocar el
estilo de un buscador significa tocar las constantes de módulo
(`TEMP_INPUT_FILL`/`CAJA_TECLEO`, `CAJA_CAMPO`, `SEL_OK_FILL`/`SEL_OK_FONT`,
`SEL_BAD_FILL`/`SEL_BAD_FONT`, declaradas junto a `CARD_FILL`/`KPI_FILL`) para que
cambien a la vez en todos.

1. **Celda única de escritura (temperatura de consulta): caja roja
   (`CAJA_TECLEO`, borde medio en `ROJO`)**, frente a la línea inferior de tinta
   (`CAJA_CAMPO`) de las celdas de lista desplegable. Las dos van sobre papel
   limpio: lo que las distingue no es el relleno sino que la que se teclea es el
   **único rectángulo cerrado** —y el único rojo— de la zona de selección. Su
   rótulo va también en rojo y dice `TEMPERATURA DE CONSULTA >>> SE TECLEA`.
   Aplica también a `build_buscador_grupo`.
2. **Semáforo de cascada completa/incompleta**, exclusivo de los 5 buscadores de
   cascada (`build_buscador`/`finish_buscador`): la celda de aviso junto a la
   temperatura (`G12:I12`) lleva formato condicional — bloque verde `SEL_OK_FILL`
   (`4AF626`, texto en tinta) con "SELECCION COMPLETA" cuando la cascada (pasos 0
   a 4) está resuelta; bloque ámbar `SEL_BAD_FILL` (`E6A019`, texto en tinta) con
   "SELECCION INCOMPLETA" mientras falte un paso. El texto va en mayúsculas y sin
   tilde ("SELECCION", no "SELECCIÓN"): todo el texto de celda del libro evita
   acentos (ver "SELECCION DEL MATERIAL", "Composicion nominal") para no arrastrar
   problemas de codificación fuera de Excel 365. `build_buscador_grupo` no lleva
   este semáforo: no tiene una única cascada que completar, cada bloque ya avisa
   "(elija grupo)" en su propia celda de valor.
3. **Curva del material: línea continua, sin marcadores, en tinta (`TINTA_2`)** —
   `ch.scatterStyle = "line"`, serie del valor tabulado con `marker="none"` y
   `smooth=False` (la interpolación del código es lineal — regla 3 de esta lista
   arriba — una curva suavizada la representaría mal). El punto consultado es un
   rombo en el rojo del acento (`ROJO`) sin línea, para distinguir el valor
   puntual de la curva. El resto de la gráfica lo estila `estilizar_chart()`:
   papel, marco de tinta, malla en trama de 25 % y texto mono. **No se usa
   `ch.style`**: los presets de Office traen su propia paleta de series y es justo
   lo que este sistema no admite. Aplica a los cuatro sitios que crean gráfica.

### `DB_B31_C` / `DB_B31_CC` — el Apéndice C entero en una base por edición

201 filas por edición: C-1/C-1C (52), C-2 (44), C-3/C-3C (75) y C-4 (30). El nivel
0 de la cascada es la **propiedad**, no la tabla: el ingeniero pregunta «quiero el
módulo E», no «quiero la Tabla C-3».

Tres cosas que este motor hace distinto de los cinco de esfuerzos, y por qué:

- **Bloquea en los DOS extremos.** El Apéndice C no publica columna «Temp. máx.»:
  el límite es el primer y el último punto que tabula la propia fila. Sostener el
  último valor por encima —que es lo correcto en los buscadores de esfuerzos,
  donde el tope lo pone una columna del código— aquí sería extrapolar.
- **Rama de dato puntual.** C-2 y C-4 no dependen de la temperatura: publican un
  valor único, la gráfica queda vacía **a propósito** y el estado dice VALOR ÚNICO.
- **El conmutador cambia de columna en dos tablas y de hoja en las otras dos**
  (regla 10). Nunca convierte.

El **factor de escala se lee del JSON**, jamás se codifica: confundir el ×10³ de
C-3 con el ×10⁶ de C-3C son tres órdenes de magnitud en el módulo E. Va en columna
propia junto con su texto impreso, y el KPI muestra el valor con el factor aplicado
mientras la sección 4 muestra el impreso sin él.

**El enlace SI↔US es posicional** (`clave_bi = propiedad#fila impresa`), porque los
nombres divergen entre ediciones por artefactos de impresión: «Type 309.» con punto
en C-3 y con coma en C-3C, «25Cr–20Ni» con raya en una y guion en la otra. Un enlace
posicional sin red es una bomba silenciosa, así que el build **aborta** si el número
de filas por propiedad no coincide o si un nombre normalizado difiere fuera de
`DIVERGENCIAS_NOMBRE_C`, que hoy tiene **una sola entrada**.

### Sección II partes A, B y C — las nueve hojas del libro

`secii_tablas.py` reconstruye las tablas de las 368 especificaciones desde los
bloques `Line` y sus `bbox`. Es **librería y CLI**: el builder la importa para
escribir las nueve hojas, y `--informe` mide y reporta a
`Revision_Tablas_SecII.md` sin tocar `resources/` ni el libro.

El JSON no trae tablas: el `html` de todo bloque `Table` es `<p></p>` y los `Span`
no se conservan, así que el bloque más fino es `Line`. El reparto de cada fila
tiene tres resultados y **ninguno adivina**: `EXACTA` (≥2 `Line`, cada uno a su
columna por el punto medio de su `bbox`), `POR CONTEO` (un solo `Line`, partido
desde la derecha y solo si el conteo cuadra con las columnas que la tabla ya
demostró tener) y `AMBIGUA` (el texto se conserva **entero** en una celda y se
declara). Sin `Span`, repartir por interpolación sobre el ancho del `bbox` sería
inventar estructura con una fuente proporcional: la geometría **confirma** un
reparto, nunca lo produce.

La garantía que sí se puede dar sin el PDF es la **comprobación sin pérdida**: la
concatenación de las celdas de cada fila coincide carácter a carácter con la de sus
`Line` de origen. Pasa sobre las 54 198 filas y los 124 115 `Line`, **y se corre
tres veces**: en el CLI, en el build (un fallo lo **aborta**: no se graba una fila
que no sea una repartición exacta) y en `verificar.py` §10, que la relee **desde
la hoja** y la contrasta contra una reconstrucción independiente desde
`resources/`.

**Estado: volcado entero al libro (Rev. 4c).** 2 572 tablas lógicas, 54 198 filas
y 5 233 notas al pie, en nueve hojas navegables:

| Hoja | Filas | Contenido |
|---|---:|---|
| `CAT_SecII` | 379 | Catálogo del índice de las cuatro partes: spec, título, páginas PDF y folios |
| `IDX_SecII_Tablas` | 2 572 | Una fila por tabla lógica: reparto por confianza, motivo dominante y por qué no se normaliza |
| `DB_SecII_A1` · `A2` · `B` · `C` | 54 198 | **Volcado íntegro**, formato ragged: `C01..C48` |
| `DB_SecII_Notas` | 5 233 | Notas al pie con su marcador |
| `DB_SecII_Quimica` | 220 | Normalizada: 68 tablas |
| `DB_SecII_Traccion` | 182 | Normalizada: 38 tablas |

**Las 24 812 filas AMBIGUAS entran al libro, y eso no contradice el punto de
parada de la Fase 2.** El plan avisaba de que escribir 124 000 filas dudosas es
peor que no escribirlas; lo que sería peor es escribirlas **como si estuvieran
tabuladas**. Aquí entran con su texto impreso **entero en `C01`** y marcadas
`AMBIGUA` con su motivo, `IDX_SecII_Tablas` lo cuenta tabla a tabla y el
Dashboard publica el total en ámbar. No se pierde nada y se puede consultar.

**Las normalizadas cubren muy poco, y el motivo es de la fuente.** El plan
estimaba ~12 000 filas de química; salen **220**. La regla —«solo se normaliza la
tabla cuyos encabezados se resuelven ENTEROS contra el vocabulario del código»—
rechaza el resto por dos motivos que se cuentan: *no se pudo componer un nombre
por columna desde la cabecera* (908 tablas) y *ninguna columna se resuelve como
elemento* (971). Los encabezados de la Sección II llegan casi siempre sin partir,
porque son filas de un solo `Line` **sin fichas de valor** con las que el conteo
pueda partirlas. Forzar el encaje daría una hoja que *parece* completa: el dato
sigue íntegro en el volcado y el motivo está impreso tabla a tabla.

Dos criterios que costaron y que no hay que revertir:

- **El marcador de nota va pegado al calificador.** El código imprime
  «Chromium, maxC» —el superíndice pegado, que `texto()` conserva a propósito—,
  así que `\bmax\b` **no casa**: entre «x» y «C» no hay frontera de palabra. Sin
  absorberlo, cinco de los diez elementos de la Tabla 1 de SA-106 acababan en
  «Otros elementos» teniendo columna propia.
- **Un guardia de fichas de valor separa un requisito de la prosa.** La Tabla 19
  de SB-111 («Significance of Numerical Limits») rotula filas «Tensile strength»
  y «Yield strength» cuyo contenido es una frase con números dentro. Un guardia
  que solo mirase «¿lleva dígitos?» la dejaba pasar; el que exige que **la mitad
  de las fichas sean fichas de valor** no.

**El encabezado que llega en una sola celda se parte, pero solo con dos reglas.**
«Grade A Grade B Grade C» se reparte porque el código **repite** la palabra clave
y porque el número de trozos es **exactamente** el de columnas de valor que la
tabla ya demostró tener (`partir_encabezado`). Si no cuadra, devuelve `None` y la
tabla no se normaliza. No se interpola nada, igual que en el resto de la capa.

**Un carácter que XML no admite.** El `index.json` de SA-533 trae un **U+FFFE** en
el título, donde el PDF imprime un guion. openpyxl lo escribe tal cual y produce
un `.xlsm` que Excel abre pero que **ningún parser XML lee** —rompía las 16
pruebas de `test_dashboard.py` sin que ninguna de sus aserciones fuese falsa—.
`xml_seguro()` sustituye los caracteres prohibidos por **U+FFFD**, que es lo que
Unicode reserva para «aquí había algo irrepresentable»: se ve, no se pierde la
posición y no se inventa el carácter. `verificar.py` §10 aplica la misma
sustitución al lado del origen antes de comparar, de modo que la única diferencia
admitida entre lo impreso y lo grabado queda declarada.

**Estas nueve hojas no llevan ni una fórmula.** Y el criterio de la prueba es el
**tipo** de celda, no si el texto empieza por «=»: el código imprime celdas como
«= 3.18 mm in any 1.524 m», que son texto y tienen que seguir siéndolo.
`_txt_celda` las fuerza a `s`; `test_dashboard.py` y `verificar.py` §10 lo
comprueban.

### `MAP_Grupo` — pertenencia a grupo de propiedades

El grupo que da E (TM-1) y dilatación (TE-1) **no se infiere de la composición**:
está impreso en las Notas al pie de esas dos tablas. Viven en `resources/` bajo
`note_members`, **en las dos ediciones**, y las extrae del PDF de II-D:

```powershell
python extraer_notas_ii_d.py --edicion si --resources ..\..\..\resources `
    --pdf "C:\Users\dvelasquez\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES\SECCION II\D Metric 2025\D Metric 2025 _p1201-p1500.pdf"
python extraer_notas_ii_d.py --edicion us --resources ..\..\..\resources `
    --pdf "C:\Users\dvelasquez\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES\SECCION II\D Customary 2025\D Customary 2025 _p1201-p1500.pdf"
```

Solo hay que reejecutarlo si se repone la extracción de II-D. El builder **aborta**
si `note_members` no está: sin fuente normativa no se construye nada.

**Las dos ediciones no numeran igual sus notas.** La métrica imprime el Grupo H
**duplicado**, en las Notas (8) y (9), y a partir de ahí toda su numeración va
corrida en uno respecto de la US, que lo imprime una sola vez en la Nota (8): 17
notas en TM-1 métrica contra 16 en la US. La errata es exclusiva de la métrica y
se conserva tal como está impresa (Regla 9). La **pertenencia** a grupo sí es
idéntica en ambas, verificado grupo a grupo. Nunca copiar las notas de una
edición a la otra: son extracciones independientes y `test_build_db.py::TestNotasEnLasDosEdiciones`
lo comprueba.

Hay **dos hojas de mapeo**, una por edición: `MAP_Grupo` (métrica) y `MAP_GrupoC`
(U.S. Customary). Cada una se resuelve contra las Notas de **su** edición.

Estado por hoja (3 454 filas cada una):

| Estado | Filas | Respaldo |
|---|---:|---|
| AUTO (UNS exacto) | 1 292 | UNS impreso en TM-1…TM-5 |
| AUTO (composición en Nota o columna) | 1 297 | Nota, o columna nombrada de TE-1, citada en la propia fila |
| AUTO (composición vía UNS en otra tabla) | 228 | el código imprime esa composición para el mismo UNS en otra de sus tablas, o —cuando el UNS a secas es ambiguo— para el mismo (UNS, **especificación**) o (UNS, **grado**) que la propia fila imprime; el UNS, con ese segundo dato si hizo falta, identifica el material de forma unívoca |
| VALIDADO POR INGENIERO | 46 | `decisiones_map_grupo.json`; 7 decisiones (ver abajo) |
| SIN MAPEO | 591 | II-D no publica el dato; cálculo bloqueado |

**Cobertura:** 2 490 filas con módulo E y 1 396 con dilatación, de 3 454. Una
fila puede tener uno y no el otro — TM-1 y TE-1 no enumeran los mismos
materiales — y la columna `Motivo` lo dice fila a fila.

### Cómo se identifica el material, y cómo se busca su grupo

Son dos preguntas distintas y el motor las separa. Confundirlas fue el origen
de los dos huecos que cerró la Rev. 4c.

**Identificar el material** (solo si la fila no imprime composición nominal):
tres vías, todas apoyadas en un dato que **la propia fila imprime**, probadas
en este orden y solo mientras la anterior deje más de un candidato —
`_candidatas_por_uns()`:

| Vía | Cuándo | Ejemplo |
|---|---|---|
| UNS a secas | el libro trae una sola composición para ese UNS | `K81340` → `9Ni` |
| (UNS, **especificación**) | el código reparte por spec: perno vs. tuerca, fundición vs. tubo fundido | `J91150`: SA-426 → `13Cr`, SA-217 → `12Cr` |
| (UNS, **grado**) | el código reparte por grado, dentro o a través de las specs | `G41400`: A193 Gr. B7 → `Cr-Mo`, Gr. B7M → `Cr-0.2Mo` |

El grado se compara **literal**, sin normalizar. A194 Gr. «6» y A193 Gr. «B6»
son grados distintos del mismo UNS `S41000` con composiciones distintas
(`12Cr` y `13Cr`): reducirlos a su dígito los confundiría y daría la
composición equivocada a una de las dos filas.

**Buscar su grupo**: dos mecanismos, los dos impresos en el código y con el
mismo rango — `_publicada_por_iid()`:

- la lista de miembros de una **Nota** de TM-1/TE-1;
- el **título de una columna nombrada** de TE-1 («Coefficients for 8Ni and 9Ni
  Steels»), que se autodescribe.

Hasta la Rev. 4b el segundo solo lo veían las filas que imprimen su
composición. El camino de composición prestada por UNS miraba **solo las
Notas**, y eso dejaba sin dilatación a materiales para los que II-D **sí la
publica**: el `9Ni` de `K81340` (38 filas entre las dos ediciones) salía como
«II-D no publica el dato» teniendo columna propia impresa, y las filas que la
Nota resolvía solo para el módulo E — `13Cr`, `15Cr`, `17Cr`, `13Cr-4Ni` —
perdían su alfa en silencio.

**Una columna condicionada no se aplica con composición prestada.** «9Cr-1Mo
Steels (Including Grades 9, 91, 911, and 92)» o «Condition 1075» no las
resuelve la composición: las resuelve un dato que la fila tiene que imprimir
aparte, y con la composición prestada ese dato no se ha comprobado. Ahí el
motor declara el caso —«TE-1 SÍ publica el dato, pero condicionado»— en vez de
elegir a ciegas entre columnas con valores distintos.

`verificar.py` §9 audita que toda cita a «TE-1 columna impresa» nombre una
columna que el JSON de **esa** edición imprima de verdad.

`verificar.py` §9 audita ambas hojas fila a fila: que toda cita exista realmente
en el JSON, que lo validado por una persona se declare como tal y que toda
composición prestada diga de dónde salió.

**Vía de retorno de las decisiones.** `Revision_MAP_Grupo.md` va en dos
bloques, y esa separación es lo importante: **casos ABIERTOS** (admiten
criterio de ingeniería) y **casos CERRADOS por límite de la fuente**. Hasta la
Rev. 4b todo lo no resuelto se listaba junto, bajo el rótulo «decisiones a
tomar» y con una casilla de firma al lado; eso describía mal la realidad, y
alguien podía rellenarla creyendo que decidía. La clase la deriva el motor del
mismo camino por el que llegó al hueco (`CIERRE_ABIERTA` / `CIERRE_LIMITE`),
nunca de leer el texto del motivo.

**Estado tras la Rev. 4c: 0 abiertos, 106 cerrados.** No queda ninguna decisión
pendiente de firma. Un caso se cierra cuando el material queda **identificado
sin ambigüedad** por alguna de las tres vías de arriba y aun así ni las Notas
ni las columnas de TM-1/TE-1 publican su grupo: ahí no hay nada que decidir —
firmar sería inventar el valor que el código no publica, justo lo que la Regla
n.º 1 prohíbe— y lo correcto es que la fila siga BLOQUEADA.

La plantilla lo refleja: `decisiones` trae solo los casos abiertos, y los
cerrados van aparte en `_cerrados_sin_decision`, sin casilla que rellenar.
Para que una decisión llegue al cálculo, copiar
`decisiones_map_grupo.plantilla.json` a `decisiones_map_grupo.json` y rellenarlo:
esas filas pasan al estado `VALIDADO POR INGENIERO`, **siempre separado de las
AUTO**. Una decisión **nunca sobreescribe** un grupo que el código sí asigna: el
choque se reporta y se conserva lo del código.

**Columna `Fila PRD`**: coincidencia literal del UNS contra las 99 filas de
`table_prd.json` que nombran los suyos (1 554 filas resueltas). Las 16 filas
restantes de PRD son categorías redactadas y encuadrar un material en ellas es
criterio de ingeniería.

**No reintroducir una heurística de composición.** La Rev. 2 asignaba grupo con
expresiones regulares y producía asignaciones *falsas*: `mo\b` rotulaba «acero de
baja aleación» a los austeníticos 16Cr-12Ni-2Mo, a los dúplex 22Cr-5Ni-3Mo-N y a
las aleaciones de níquel 62Ni-22Mo-15Cr; `8ni` casaba dentro de `18Ni`.
`test_build_db.py::TestMapeoDeGrupos` lo impide.

### TE-1 no reparte todo por Grupos numerados

Además de los Grupos 1 a 4 (que llegan por Nota), TE-1 publica dilatación en
**columnas que se autodescriben en su propio título**: `15Cr and 17Cr Steels`,
`12Cr, 12Cr–1Al, 13Cr, and 13Cr–4Ni Steels`, `27Cr Steels`, `8Ni and 9Ni Steels`,
`5Cr–1Mo and 29Cr–7Ni–2Mo–N Steels`, `5Ni–1/4Mo Steels`,
`9Cr–1Mo Steels (Including Grades 9, 91, 911, and 92)`. Son tan normativas como
las Notas y `columnas_nombradas_te1()` las resuelve **literalmente**, extrayendo
del título las designaciones que enumera. Cuando el título condiciona por
**grado** —el caso de los 9Cr-1Mo— la comprobación sigue siendo literal, porque
el grado está impreso en la propia fila.

### Las decisiones tomadas, y por qué

`decisiones_map_grupo.json` contiene **siete**, y son **todas** las que quedan:
tras la Rev. 4c `Revision_MAP_Grupo.md` no lista ni un caso abierto.

La primera: `9Cr-1Mo-V` (Grado 91 / P91 / F91 /
T91, 17 filas) → `Material Group E` para el **módulo E**. TM-1 **no** lo asigna
literalmente —ninguna de sus Notas nombra ese material ni el UNS K90901— y se
aplica la Nota (5), *«9Cr–Mo, including variations thereof»*. Apoyos: Parte A
(K90901 es 9Cr-1Mo con V, Nb y N), Parte C SFA-5.5 §A7.2.3.1 (describe el
electrodo EB91 como *«a 9% Cr–1% Mo electrode modified with niobium and
vanadium»*) y la propia TE-1, que agrupa el Grado 91 con los 9Cr-1Mo.
**Solo afecta a E**: la dilatación la da la columna impresa de TE-1, que nombra
el Grado 91, y por eso `grupo_te` va vacío en esta decisión.

Las otras cinco resuelven UNS que el libro asociaba con **más de una**
composición —el caso que el propio motor se niega a decidir por cuenta propia
(`_composicion_prestada`)— pero donde las candidatas no son en realidad dos
materiales distintos: son la misma composición, con una de sus apariciones
rota por la costura de páginas enfrentadas del Apéndice A-1C (la misma clase
de defecto que ya corrigió `a66f4dd` en otras filas de esa tabla). La
evidencia es interna: el mismo UNS trae, en otra fila del mismo Apéndice A,
la composición limpia.

| UNS | Candidatas del libro | Composición real | Grupo E (TM) | Grupo dilatación (TE) | Filas |
|---|---|---|---|---|---:|
| `K41545` | `5Cr-1/2Mo` · `5Cr-1/2Mo A387 Gr. 5 Cl. 1` | `5Cr-1/2Mo` (Nota (5), Grupo E) | Material Group E | Columna `5Cr-1Mo and 29Cr-7Ni-2Mo-N Steels` | 18 |
| `K90941` | `9Cr-1Mo` · `9Cr-1Mo A387 Gr. 9 Cl. 1` | `9Cr-1Mo` (Nota (5), Grupo E) | Material Group E | Columna `9Cr-1Mo Steels (Including Grades 9, 91, 911, and 92)` — todas las filas son Grado 9 | 12 |
| `K11820` | `C-1/2Mo` · `A` | `C-1/2Mo` (Nota (1), Grupo A) — «A» es el Tipo/Grado de SA-204 pegado por la costura | Material Group A | Group 1 | 8 |
| `K12020` | `C-1/2Mo` · `B` | `C-1/2Mo` — mismo defecto, «B» es Tipo/Grado de SA-204 | Material Group A | Group 1 | 8 |
| `K12320` | `C-1/2Mo` · `C-1/2Mo A204 Gr..` | `C-1/2Mo` — misma costura, grado pegado a la composición | Material Group A | Group 1 | 8 |

`K90901` (Grado 91) y `K90941` (Grado 9) son UNS distintos: esta decisión no
sustituye ni contradice la del 9Cr-1Mo-V.

**La séptima (Rev. 4c): `J82090`, el 9Cr-1Mo moldeado** (SA-217 Gr. C12 y
SA-426 Gr. CP9, 4 filas) → `Material Group E`, `grupo_te` **deliberadamente
vacío**. La fila de II-D no imprime composición, pero el Apéndice A del B31.3
sí la imprime limpia para ese mismo UNS en sus dos ediciones: `9Cr-1Mo`.
Identificado el material, el módulo E lo resuelve la Nota (5) de TM-1
—*«9Cr–Mo, including variations thereof»*—, el mismo fundamento normativo que
ya sostiene `K90901` y `K90941`; por ser regla redactada y no lista literal, va
como decisión y no como AUTO. La **dilatación se deja vacía a propósito**:
TE-1 publica el 9Cr-1Mo en una columna condicionada por grado —«Including
Grades 9, 91, 911, and 92»— y los grados impresos aquí son `C12` y `CP9`,
designaciones de pieza moldeada que esa columna no nombra. Hacer casar el «9»
de `CP9` con el «Grade 9» de la columna sería coincidencia de dígitos, no una
afirmación del código.

Los **106 casos restantes están CERRADOS por límite de la fuente**, no
pendientes: el material queda identificado sin ambigüedad por alguna de las
tres vías, y aun así su composición no figura en ninguna Nota ni columna de
TM-1/TE-1 (o el UNS no aparece con composición en ningún otro sitio del
libro). II-D no publica E ni dilatación y **lo correcto es que sigan
bloqueadas**. Verificado que ninguna coincide con una nota con los elementos
en otro orden.

**Comprobado también contra Sección II Partes A/B/C (2026-09-08).** Las 379
especificaciones ya extraídas en `resources/ASME_BPVC/Sec_II/{a_1,a_2,b,c}`
—no volcadas al libro, ver más abajo— no aportan ninguna cita nueva: sus
tablas de composición imprimen rangos numéricos (%Cr, %Ni, %Mo), y
`comp_key()` exige coincidencia literal con el vocabulario taquigráfico de
las Notas de TM-1/TE-1 ("12Cr", "9Cr-1Mo"). Son formatos incompatibles por
diseño, no una pérdida de extracción. Para las filas de mayor peso (aleaciones
de níquel tipo 353 MA, Incoloy, Hastelloy) se confirmó que TM-1/TE-1
sencillamente no las tabula, ni por UNS ni por Nota: II-D no publica el dato
para esa familia, no es un hueco de esta base.

Sí se encontró una vía real —pero en el Apéndice A del B31.3, ya indexado, no
en Sección II— para los cuatro UNS ambiguos: desambiguan limpio **por
especificación** o **por grado**, los dos datos impresos en la propia fila.
`indice_composicion_por_uns` construye los dos índices finos y
`_composicion_prestada` los consulta como *fallback*, nunca primero. Probarlos
siempre —no solo como fallback— habría cambiado la redacción del motivo en las
~200 filas que el camino simple ya resolvía bien, sin necesidad;
`verificar.py` §9 lo detectó en la primera versión del cambio (270 filas sin la
cita esperada) y obligó a corregir el orden. Ni la especificación ni el grado
son criterios elegidos aparte, así que el resultado es `AUTO (composición vía
UNS en otra tabla)`, nunca una decisión de ingeniería.

Desenlace de los cuatro, tras añadir la vía por grado (Rev. 4c):

| UNS | Desambigua por | Composición | Resultado |
|---|---|---|---|
| `S41000` | **grado** `410` (A240 Gr. 410 en el Apéndice A) | `13Cr` | **resuelto entero**: Grupo F + columna TE-1. La vía por spec no llegaba —el Apéndice A no trae fila de SA-479 para este UNS— pero su grado impreso sí está allí |
| `J91150` | especificación (SA-426 / SA-217) | `13Cr` / `12Cr` | **resuelto**: SA-426 da Grupo F y columna TE-1; SA-217 solo columna TE-1 (`12Cr` no es miembro de ninguna Nota de TM-1, así que no hay módulo E) |
| `S41003` | especificación (SA-1010) | `12Cr-1Ni` | **cerrado**: identificado, pero `12Cr-1Ni` no figura en ninguna Nota ni columna |
| `G41400` | **grado** (`B7` → `Cr-Mo`, `B7M` → `Cr-0.2Mo`) | según grado | **cerrado**: identificado, pero ni `Cr-Mo` ni `Cr-0.2Mo` figuran en ninguna Nota ni columna |

Los dos cerrados ya no son «no se elige por cuenta propia» sino «identificado,
y el código no publica el dato»: es la diferencia entre una firma pendiente y
un límite de la fuente. `test_build_db.py::TestComposicionPrestadaPorEspecificacion`
y `::TestCierreDeLaRevision` cubren las cuatro ramas, incluida la que exige que
el grado **no** se reduzca a su dígito.

**Las columnas de TE-1 partidas por tratamiento térmico sí se resuelven, contra
el dato impreso en la propia fila.** El `17Cr–4Ni–4Cu` tiene dos columnas B en
TE-1 —`Condition 1075` y `Condition 1150`, con valores distintos— y `class_condition_temper`
de `table_1a.json`/`table_3.json` (columna `Clase/Cond./Temple` del libro) sí
imprime cuál aplica (`H1075`, `H1100`, `H1150`) fila a fila. `columnas_nombradas_te1`
sigue colapsando esa condición a una sola por diseño (es la "pegajosa" que evita
perder la condición cuando el rótulo viene truncado); `condiciones_tratamiento_te1`
no colapsa nada y guarda las dos, y `build_map_grupo` las casa contra el
tratamiento impreso en la fila. De las 12 filas del `17Cr–4Ni–4Cu`: **7 (4×H1150,
3×H1075) resuelven automáticamente**, citando el dato impreso; **5 (H1100) quedan
bloqueadas porque TE-1 sencillamente no publica columna para esa condición** —no
es ambigüedad de mapeo, es que el código no tiene el dato— y el `Motivo` lo dice
así, distinto del caso en que la fila no imprime tratamiento alguno.

`validado_por` es **opcional**: se registra si está, pero no se exige. Lo que
separa una decisión de un dato del código es el **estado** de la fila
(`VALIDADO POR INGENIERO` frente a `AUTO (…)`) y su justificación, no que alguien
haya escrito un nombre en un JSON.

---

## Protocolo de verificación antes de entregar un cálculo

1. ¿Todas las entradas en SI? (psi → MPa, in → mm, °F → °C)
2. ¿Verificado el espesor remanente contra el mínimo admisible?
3. ¿Incluido el margen de corrosión para la vida remanente?
4. ¿Citado el artículo de ASME PCC-2 que soporta cada fórmula aplicada?
5. ¿Leídas las notas del material? (`Notas_Codigo`; restringen soldadura, PWHT y servicio)

## Código Python

Modular, con pruebas, documentado. Son herramientas de ingeniería críticas: cada
decisión no obvia va comentada explicando **por qué**, no qué hace.
