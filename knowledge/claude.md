# Guía de Instrucciones de Cálculo — ASME PCC-2 (Sistema Internacional)

Estas instrucciones establecen las reglas, criterios de diseño y fórmulas fundamentales del estándar **ASME PCC-2** (*Repair of Pressure Equipment and Piping*) adaptadas estrictamente al **Sistema Internacional de Unidades (SI)**. Deben utilizarse como instrucciones de sistema para garantizar que los cálculos de ingeniería de reparación se realicen de forma 100% precisa, consistente y normativa.

---

## 1. Reglas Generales de Unidades (SI por defecto, con soporte dual SI/US)

**Cálculos:** por defecto, los motores de cálculo operan en Sistema Internacional (SI). Si el usuario entrega datos en unidades inglesas (US Customary: psi, in, °F), se convierten a SI antes de aplicar las fórmulas. Las unidades internas de trabajo son:

- **Presión** ($P$, $P_D$, $P_{test}$): Megapascales (MPa). (Conversión: $1\ \text{bar} = 0.1\ \text{MPa}$; $1\ \text{psi} \approx 0.006895\ \text{MPa}$).
- **Dimensiones lineales** ($D$, $t$, $L$, $d$, $W$): Milímetros (mm). (Para módulos elásticos u otras ecuaciones de gran escala, usar metros (m) según se indique).
- **Esfuerzos y Resistencia a la Tracción** ($S$, $S_{act}$, $S_y$): Megapascales (MPa).
- **Módulo de Elasticidad** ($E$, $E_c$, $E_s$): Megapascales o Gigapascales (MPa o GPa).
- **Fuerzas lineales/axiales** ($F$, $F_{CP}$, $F_{LP}$): Newton por milímetro (N/mm) para cargas unitarias, o Newtons (N) para cargas totales.
- **Energía almacenada** ($E$): Joules (J).
- **Temperatura** ($T$, $T_d$, $T_g$): Grados Celsius (°C).

**Bases de datos y buscadores (dual SI/US):** las bases de datos de materiales almacenan y muestran los valores en **ambos sistemas de unidades**, y cada buscador (motor de consulta) incluye un **conmutador SI ↔ US**. Fuentes:
- **ASME B31.3** — tablas nativas en ambos sistemas: SI (A-1, A-4, B-1, C-1, C-3) y US Customary (sufijo C: A-1C, A-4C, B-1C, C-1C, C-3C). Se cargan las dos versiones tal como están impresas.
- **ASME BPVC II-D** — tablas nativas en **ambos sistemas**: Métrica/SI en `resources/ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/` (°C, MPa) y **U.S. Customary** en `resources/ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025/` (°F, ksi). Ambas ediciones son técnicamente equivalentes (un material = una sola curva de diseño); se cargan tal como están impresas. Nota: un material con especificación dual (p. ej. `SA-516/SA-516M`) es aplicable indistintamente en cualquiera de las dos ediciones (`SA-516 Gr.70` ↔ `SA-516M Gr.485`).
- Al leer o citar un valor tabulado, respetar la unidad y el redondeo **tal como están impresos** en la tabla del sistema elegido; no mezclar SI y US en un mismo cálculo.

---

## 2. Criterios de Cálculo por Parte del Estándar

### PARTE 2: Reparaciones Soldadas (Welded Repairs)

#### A. Reconstrucción de Pared por Soldadura (Weld Buildup — Artículo 202)

**1. Extensión mínima del sobreespesor de soldadura ($B$):**

El depósito de soldadura debe extenderse con su espesor nominal completo una distancia mínima $B$ en todas las direcciones más allá del límite del metal base degradado:

$$B = \frac{3}{4} \sqrt{R \cdot t_{nom}}$$

Donde:
- $B$: Distancia de extensión mínima del parche de soldadura (mm).
- $R$: Radio exterior del componente (mm), es decir, $D_o/2$.
- $t_{nom}$: Espesor nominal de pared del componente (mm).

**2. Espaciamiento mínimo entre depósitos adyacentes:**

Dos reconstrucciones por soldadura no deben estar separadas a una distancia menor que:

$$\text{Espaciamiento} \geq 3 \sqrt{R \cdot t_{nom}} \quad \text{(medido de pie a pie de soldadura adyacente)}$$

**3. Presión mínima de rotura para pruebas de calificación ($P$):**

$$P = \frac{2 \cdot t \cdot S_{act}}{D}$$

Donde:
- $D$: Diámetro exterior del componente (mm).
- $S_{act}$: Resistencia a la tracción real reportada para el material base de prueba (MPa).
- $t$: Espesor mínimo especificado del material base (excluyendo tolerancias de fabricación) (mm).

#### B. Parches de Plancha con Soldadura de Filete (Fillet Welded Patches — Artículos 207 y 212)

**1. Cargas de presión de diseño en el parche:**

- Fuerza circunferencial unitaria (Aro / Hoop): $$F_{CP} = \frac{P \cdot D_m}{2}$$
- Fuerza longitudinal unitaria: $$F_{LP} = \frac{P \cdot D_m}{4}$$

Donde:
- $F_{CP}, F_{LP}$: Fuerzas unitarias debido a la presión interna (N/mm).
- $P$: Presión de diseño interna (MPa).
- $D_m$: Diámetro medio del componente (mm) ($D_o - t$).

**2. Cargas de diseño totales (considerando cargas externas):**

Si existen momentos de flexión o cargas de tensión externos, deben agregarse de la siguiente manera:

$$F_C = F_{CP} + F_{CO}$$
$$F_L = F_{LP} + F_{LO}$$

Donde $F_{CO}$ y $F_{LO}$ son las fuerzas circunferenciales y longitudinales unitarias de origen externo.

**3. Distancia de setback mínima a discontinuidades ($L_{min}$):**

Para evitar la superposición de campos de esfuerzos con boquillas (nozzles) u otras discontinuidades estructurales, el parche debe colocarse a una distancia mínima de:

$$L_{min} = 2 \sqrt{R_m \cdot t}$$

Donde:
- $L_{min}$: Distancia mínima de separación (mm).
- $R_m$: Radio medio del componente (mm).
- $t$: Espesor de pared del componente (mm).

**4. Límites de deformación en frío (Cold Forming Limits):**

El alargamiento de la fibra extrema para planchas de acero al carbono o de baja aleación conformadas en frío no debe exceder el 5% sin tratamiento térmico posterior. Se calcula como:

- Doble curvatura: $$\text{Alargamiento} = \frac{75 \cdot T}{R_f} \left(1 - \frac{R_f}{R_o}\right)\%$$
- Curvatura simple: $$\text{Alargamiento} = \frac{50 \cdot T}{R_f} \left(1 - \frac{R_f}{R_o}\right)\%$$

Donde:
- $T$: Espesor del parche (mm).
- $R_f$: Radio final de la línea media del parche (mm).
- $R_o$: Radio original de la línea media (mm) ($\infty$ si originalmente es plano).

---

### PARTE 4: Reparaciones No Metálicas y Enlazadas (Nonmetallic Repairs)

#### A. Reparaciones con Materiales Compuestos (Composite Wraps — Artículos 401 y 402)

El espesor de reparación mínimo ($t_{min}$) o el número mínimo de capas ($n_H, n_A$) se rigen por la presencia o ausencia de fugas y la pérdida de espesor del sustrato metálico.

**1. Diseño Tipo A (Componentes no activos / Sin fuga — Artículo 402):**

- Número mínimo de envolturas circunferenciales ($n_H$):

$$n_H = \frac{P \cdot D}{2 \cdot S_{wh} \cdot d_f}$$

Donde:
- $P$: Presión de diseño (MPa).
- $D$: Diámetro exterior del componente (mm).
- $S_{wh}$: Resistencia circunferencial a la tracción de la envoltura por capa por milímetro de ancho (N/mm).
- $d_f$: Factor de diseño normativo, establecido por defecto en 0.2 (para Artículo 402).

- Número mínimo de capas axiales ($n_A$) — requerido solo si el sustrato se ha reducido en un 50% o más de su espesor original:

$$n_A = \frac{P \cdot D}{4 \cdot S_{wa} \cdot d_f}$$

Donde $S_{wa}$ es la resistencia axial por capa por milímetro de ancho.

**2. Longitud de traslape axial ($L_{over}$):**

El parche compuesto debe extenderse axialmente más allá de los límites del defecto una distancia mínima $L_{over}$:

$$L_{over} = 2.5 \sqrt{\frac{D \cdot t}{2}} \quad \text{con un límite mínimo absoluto de } 50\ \text{mm (2 in)}$$

Donde $t$ es el espesor nominal del tubo metálico sustrato.

**3. Diseño Tipo B (Defecto con fuga activa / Tubo pasante — Artículo 401):**

Para cálculos de deflexión y capacidad de sellado en fugas vivas, cuando se calcula el espesor de laminado mínimo $t_{min}$ (utilizando ecuaciones que involucren el factor de liberación de energía de adherencia $\gamma_{LCL}$ o $\gamma_{mean}$), se debe multiplicar obligatoriamente el valor de $\gamma_{LCL}$ por un factor de $1 \times 10^{-3}$ si se opera en el Sistema Internacional (SI) para balancear las unidades en Joules y milímetros:

$$\gamma_{LCL,SI} = \gamma_{LCL} \times 10^{-3}$$

---

### PARTE 5: Pruebas y Ensayos (Examination and Testing)

#### A. Cálculos de Energía Almacenada en Pruebas Neumáticas (Artículo 501)

Cuando se planee realizar una prueba neumática de presión, se debe calcular rigurosamente la energía almacenada para evaluar los márgenes de seguridad.

**1. Fórmula de energía almacenada en el Sistema Internacional (SI):**

$$E = 2.5 \cdot P_a \cdot V \left[ 1 - \left(\frac{P_a}{P_{at}}\right)^{0.286} \right]$$

Donde:
- $E$: Energía almacenada (Joules, J).
- $P_a$: Presión atmosférica absoluta constante, definida en $101{,}000\ \text{Pa}$ ($0.101\ \text{MPa}$).
- $P_{at}$: Presión absoluta de prueba neumática (Pa) (Presión manométrica de prueba + $101{,}000\ \text{Pa}$).
- $V$: Volumen total bajo presión en la prueba (m³). (Nota: Para sistemas de tubería extensos, se puede limitar el volumen equivalente a una longitud de 8 diámetros nominales en análisis de falla local).

**2. Conversión a Equivalencia en TNT (kg):**

$$\text{TNT (kg)} = \frac{E}{4{,}266{,}920}$$

**3. Distancia de Seguridad de Onda de Choque (Safe Blast Distance — $R$):**

- Para energía almacenada baja ($E \leq 8{,}130{,}000\ \text{J}$): $$R_{min} = 30\ \text{metros}$$
- Para energías elevadas ($E > 8{,}130{,}000\ \text{J}$): $$R = R_{scaled} \cdot (2 \cdot \text{TNT (kg)})^{1/3}$$

Donde $R_{scaled}$ es el factor de consecuencia de escala normativo.

---

## 3. Protocolo de Verificación

Antes de entregar cualquier resultado de cálculo de ASME PCC-2, se debe realizar la siguiente autocomprobación:

1. **¿Están todas las entradas en unidades del SI?** Si el usuario proporciona datos en PSI, pulgadas o Fahrenheit, convertirlos inmediatamente a MPa, mm y °C antes de aplicar cualquier fórmula.
2. **¿Se ha verificado la integridad estructural del sustrato?** Validar que el espesor remanente del tubo ($t_s$) no viole los límites normativos mínimos admisibles.
3. **¿Se ha incluido el margen de corrosión aplicable?** El diseño del espesor de reparación ($t_{repair}$ o $T$) debe sumar siempre el sobreespesor por corrosión proyectado para la vida remanente del diseño.
4. **Citar de forma estricta:** Al reportar las salidas, se debe indicar qué artículo de ASME PCC-2 soporta la fórmula aplicada (ej. Artículo 202 para Weld Buildup, Artículo 401 para Composite Wraps en aplicaciones de alto riesgo).

---

## 4. Mantenimiento de este documento

Este archivo debe actualizarse (fórmulas, coeficientes, artículos citados) contrastando su contenido contra el material de referencia de ASME PCC-2 almacenado en `resources/ASME PCC/`; ante cualquier discrepancia, ese material de `resources/` prevalece como fuente fidedigna y debe reflejarse aquí.

---

## 5. Fuente Única de Verdad — Carpeta `resources/`

La carpeta **`resources/`** es la **única fuente autorizada** de códigos, normas y estándares para este proyecto. Se debe usar exclusivamente su contenido para:

- **Citas y referencias normativas** (artículos, secciones, tablas, notas de código).
- **Textos técnicos** que describan criterios, fórmulas o requisitos de un estándar.
- **Construcción de motores de cálculo** en hojas de cálculo (Excel/Python), incluyendo coeficientes, factores de tabla y límites admisibles.

Reglas de aplicación:

1. **No usar conocimiento general ni memoria del modelo** como fuente de un valor normativo (esfuerzos admisibles, factores, fórmulas, límites) si existe un archivo correspondiente en `resources/`. Siempre leer el archivo fuente antes de citar o calcular.
2. **Estructura vigente de `resources/`** (verificar antes de cada uso, puede reorganizarse):
   - `resources/ASME PCC/` — ASME PCC-2 (manifiesto + `pcc_2/`).
   - `resources/ASME B31/ASME B31.3/APPEX/` — ASME B31.3 2024: índice (`asme_b31_3_2024_index.json`) y Apéndices A, B y C:
     - **Apéndice A** — esfuerzos admisibles básicos en tracción de metales (Tabla A-1 / A-1C), esfuerzos de diseño de **pernería** (Tabla A-4), y factores de calidad: fundición `Ec` (Tabla A-2) y junta longitudinal soldada `Ej` (Tabla A-3).
     - **Apéndice B** — esfuerzos de diseño hidrostático (HDS) y tablas de presión admisible para **tuberías NO metálicas** (termoplásticos / RTR), Tablas B-1 a B-6. (No trata factores de fatiga en soldadura.)
     - **Apéndice C** — propiedades físicas de materiales: dilatación térmica de metales (C-1) y no metales (C-2), y módulo de elasticidad de metales (C-3) y no metales (C-4).
   - `resources/ASME B31/ASME B31.3/CHAPTERS/` — cuerpo normativo del B31.3: capítulos (`chapter_01..10.json`), definiciones, figuras y `tables/` con las tablas del articulado (302.3.3-1 `Ec` incrementado, 302.3.4-1 `Ej`, 341.3.2-1, 326.1.1-1…).
     - **Cuidado con el doble prefijo:** cada tabla existe como `table_X.json` y `table_table_X.json`. **Se cita siempre la de prefijo simple**, que es la extracción recompuesta; el gemelo conserva los fragmentos crudos del corte de línea. Los 32 pares llevan la declaración dentro del propio archivo (`canonical_declaration` en el canónico, `superseded_by` en el gemelo), escrita por `declarar_tablas_canonicas.py`.
     - **Hueco declarado:** `table_302_3_4_1.json` perdió el cuerpo de la tabla — sus filas se colapsaron dentro de los encabezados de columna. Lo dice su propio bloque `extraction_gap`. **No se puede citar de ella el factor `Ej` incrementado** del para. 302.3.4(b); hay que leerlo del folio impreso.
   - `resources/ASME_BPVC/Sec_II/bpvc_ii_a_1/`, `bpvc_ii_a_2/`, `bpvc_ii_b/`, `bpvc_ii_c/` — BPVC Sección II, **Partes A, B y C**: texto íntegro de 379 especificaciones de material (SA-, SB-, SFA-) más apéndices y figuras. `index.json` es el catálogo y `meta.json` la procedencia.
     - **Las tablas NO vienen como `<table>`.** El `html` de todo bloque `Table` está vacío y las filas cuelgan como bloques `Line` con su `bbox`. Recomponer las columnas exige leer esos `Line`: lo hace `secii_tablas.py`, que el builder importa como librería para escribir las nueve hojas del libro (`CAT_SecII`, `IDX_SecII_Tablas`, los cuatro `DB_SecII_*`, `DB_SecII_Notas` y las normalizadas de química y tracción), y que como CLI mide y reporta a `Revision_Tablas_SecII.md` sin escribir nada.
     - **El 45,8 % de las filas queda `AMBIGUA`, y eso no es una pérdida.** Sin `Span`, la posición de cada palabra dentro de un `Line` no está en el fichero: repartirla por interpolación sería inventar estructura. Esas filas entran al libro con su texto impreso **entero en una celda** y marcadas; `IDX_SecII_Tablas` dice cuántas hay por tabla y por qué. Al citar de ellas, leer la celda completa.
     - **Cobertura de texto declarada por el propio drop:** 97,0 % (A-1), 96,8 % (A-2), 95,2 % (B) y 95,5 % (C), pero **las páginas apaisadas bajan hasta el 74 %** en la Parte B y el 85,7 % en A-2 — y son justo las tablas anchas de aleación y propiedades mecánicas. Para valores leídos de esas tablas, **el PDF manda**.
   - `resources/ASME_BPVC/Sec_II/bpvc_ii_d_metric_2025/` — BPVC Sección II-D **(Metric/SI)** 2025 (°C, MPa): tablas de materiales (1A, 1B, 2A, 2B, 3, 4, 5A, 5B, 6A-6D, U, Y-1, Y-2, TE, TM, PRD, TCD), notas, apéndices y Subparte 3 (presión externa).
   - `resources/ASME_BPVC/Sec_II/bpvc_ii_d_customary_2025/` — BPVC Sección II-D **(U.S. Customary)** 2025 (°F, ksi): mismas tablas y estructura que la métrica. Se usa como fuente nativa cuando el buscador está en modo US.
### Lectura de tablas de esfuerzos de la Subparte 1 (II-D, Tablas 1A/1B/2A/2B/3/5x/6x, U, Y-1)
Las extracciones JSON ya fusionan las "páginas enfrentadas" del PDF en una sola fila por material (identificación + aplicabilidad + valores + notas). Reglas al leer o calcular:
- **Localizar la fila** por: Nominal Composition, Product Form, Spec. No., Type/Grade, UNS No. y Size/Thickness; el `line_no` es el ancla de la fila.
- **Aplicabilidad y Tmáx por código** (columnas I / III / VIII-1 / VIII-2 / XII): `NP` = no permitido para esa construcción; `SPT` = solo permitido como soporte. Verificar antes de usar el material.
- **External Pressure Chart No.** (p. ej. `CS-2`, `HA-4`): remite a la Subparte 3 para cálculos de presión externa/pandeo.
- **Notas** (`G5`, `W12`, `T1`, …): leer siempre; restringen soldadura, tratamiento térmico o temperatura.
- **Tipografía del valor**: texto normal = régimen a corto plazo (tracción/fluencia); **cursiva** = valores gobernados por **creep/termofluencia** (dependiente del tiempo). Si la fuente no capturó la tipografía, señalarlo como limitación.
- **Interpolación**: lineal entre las dos temperaturas tabuladas vecinas; redondear al mismo número de decimales que el valor de la temperatura superior. Se permite usar valores listados por encima del límite de uso del código **solo** para interpolar hasta el límite.
- **Extrapolación**: **prohibida** por encima de la última temperatura tabulada, salvo autorización explícita del código de construcción.
- **Factores de escala** al leer propiedades: esfuerzos en ksi → ×1000 para psi; en MPa → ×1000 para kPa. Módulo E de TM: aplicar el factor indicado en el título de la tabla.

3. Si un dato requerido **no existe** en `resources/`, no se debe inventar ni aproximar desde memoria general: se debe indicar explícitamente el vacío y preguntar al usuario (vía AskUserQuestion) cómo proceder o si debe añadirse el documento faltante.
4. Cualquier motor de cálculo (Excel, Python, etc.) construido para este proyecto debe referenciar y trazar sus coeficientes/tablas hacia el archivo específico de `resources/` del que provienen, para permitir auditoría posterior.
5. `OUTPUTS/` y `TEMPLATES/` nunca son fuente de citas o normas — solo contienen entregables o plantillas de formato.
