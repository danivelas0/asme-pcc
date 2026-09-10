# Motor Art. 206 (collar de encierro total) y separación parche / collar

**Estado:** EN CURSO. Fase 1 validada en Excel real por el ingeniero. Fase 2
ejecutada: `construir_seccion7_material()` extraída (con test de paridad que
fija las fórmulas de Art. 212 antes/después) y `Collar_PCC2_Art206` (Type
A/B) construida 100% en código, verificada con `verificar.py` §1-10 (0
fallos, recálculo real incluido) y a mano en Excel para un caso Type A y uno
Type B. Fase 3 (nodo del árbol + `mod_nav.vba`) quedó resuelta como parte de
la Fase 2, ya que sin la sincronía de navegación la hoja nueva no pasaba
`test_dashboard.py::TestSincroniaPythonVba`. **Fase 4 (desanclar el 212 del
Rev0) sin ejecutar** — el ingeniero decidió alcance "Total", pendiente de un
plan propio dado el riesgo de reescribir ~130 filas de fórmulas ya validadas
sin poder volver a validarlas en este entorno salvo con Excel real.

Libro que cita este plan: `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm`,
**71 hojas**, dos motores de cálculo (`Parche_PCC2_Art212`,
`Collar_PCC2_Art206`).

---

## Origen — por qué este plan existe

Al abrir el motor de cálculo se detectaron dos fallos y una decisión de diseño:

1. **`#N/D` en el espesor de pared** y en cascada toda la geometría. Causa:
   el lookup de la cédula casaba `$D$19` contra la fila de encabezados guardada
   como **texto** (incluye `STD/XS/XXS`); al elegir la cédula del desplegable
   Excel la mete como **número** y el `MATCH` fallaba.
2. **`Sa_c`/`Sa_b` en `#N/D`** pese al "caso precargado": la cascada de material
   de la Sección 7 se entrega vacía, y los selectores visibles `D22/D23` ya no
   alimentan el cálculo.
3. La hoja se titula **"PARCHE / COLLAR"** pero implementa solo el **Art. 212**
   (parche). El usuario pidió una selección previa parche-vs-collar.

Sobre (3) se decidió, alineado con el principio del proyecto **"Un dato, un
motor"**: no un conmutador dentro de una hoja, sino **dos motores distintos**,
porque son **dos artículos distintos** del PCC-2 con bases de diseño distintas.

Confirmado leyendo `resources/ASME PCC/pcc_2/`:

| | **Art. 212 — Parche** (Fillet Welded Patches) | **Art. 206 — Collar** (Full Encirclement Steel Reinforcing Sleeves) |
|---|---|---|
| Ecuaciones propias | (1),(3),(4),(5),(6),(7) | **ninguna** (`counts.equations = 0`) |
| Cómo dimensiona | Esfuerzo de soldadura membrana+flexión ≤ 1,5·Sa; `w_mín=F/(E·Sa)`; `L_mín=2√(Rm·t)`; conformado ≤5% | **Remite al código de construcción** (t_req por B31.3 / VIII-1) + reglas geométricas |
| Sub-tipos | — | **Type A** (refuerzo, no contiene presión) y **Type B** (contiene presión) |

---

## Decisiones ya tomadas (con el ingeniero)

- **D-1. Dos motores**, no un conmutador. `Art. 212` = parche; `Art. 206` = collar.
- **D-2. Art. 206 con Type A y Type B** (el artículo completo).
- **D-3. Entrega parchando sobre Rev4** (generado por el builder), no una Rev5 nueva.
- **D-4. Desanclar del Rev0**: el motor nuevo nace en código; el 212 se va
  desanclando (Fase 1 ya escribe sus correcciones en el builder, no en el maestro).
- **D-5. Secuencia: Fase 1 primero** (212, ya hecha y a validar), luego el 206.

---

## Fase 1 — Motor Art. 212 (parche) · EJECUTADA, a validar

Todo en `build_db_materiales.py`, dentro de `corregir_art212_fase1()`
(llamada desde `integrate_motor`), para que la revisión posea las correcciones
y no queden en el maestro Rev0.

- **Espesor robusto:** `D21` pasa a `MATCH(""&$D$19, Datos_Ref!$C$4:$P$4, 0)`.
  El `""&` fuerza la cédula a texto: idempotente, y tolera número o `STD/XS/XXS`.
- **Caso precargado que calcula:** se siembra el material real por la celda
  *Variante* (paso 5) de la Sección 7 —`D114` = A106 Gr.B (base),
  `E114` = A516 Gr.70 (collar)— con `material_id` **verificado contra `DB_B31_3`**
  (el build aborta si no existe). El dictamen `F90` distingue
  **"ELIJA MATERIAL (Seccion 7)"** de **"FUERA DE RANGO"**.
- **Títulos de parche:** `A1`/`A2` a Art. 212; `G22`/`G23` marcan `D22/D23` como
  descriptivos (el S(T) rige en la Sección 7).

**Verificado aquí:** build sin abortos; `verificar.py §1-5` (271 536 valores,
0 discrepancias; 0 fórmulas de matriz dinámica; 0 validaciones no portables);
203 pruebas en verde (los 4 fallos de `test_dashboard.py::TestLintVba` son de
`winreg`, exclusivo de Windows).

**Pendiente (lado del ingeniero, en Excel real):** F9 y comprobar
`t`=6,35 mm (12" · SCH 20), geometría completa, `Sa_b≈138 / Sa_c≈161 MPa`.
El recálculo real de `verificar.py §6+` necesita `win32com` (Excel de Windows).

---

## Fase 2 — Motor Art. 206 (collar de encierro total) · SIN EJECUTAR

Hoja nueva `Collar_PCC2_Art206`, **generada íntegramente por el builder**
(nace desacoplada del Rev0). Selector **Type A / Type B**.

### Lo que publica el Art. 206 (leído de `resources/`, no de memoria)

El Art. 206 **no trae ecuaciones propias**: su diseño es (i) el t_req del código
de construcción y (ii) reglas geométricas con números que el propio artículo
imprime. Todo lo de abajo se cita del texto normativo:

- **206-1.1.1 Type A** — extremos **no** soldados circunferencialmente; **no**
  contiene presión, funciona como refuerzo. Costuras longitudinales soldadas.
- **206-1.1.2 Type B** — extremos soldados circunferencialmente; **sí** contiene
  presión.
- **206-2.5** — Type A puede no servir para defectos **circunferenciales** (no
  resiste cargas axiales).
- **206-3.1 Type A** — espesor **≥ ⅔ del espesor del tubo portador**.
- **206-3.2 Type B** — espesor **≥ el requerido para la máxima presión de diseño**
  admisible (o el espesor pleno del tubo si lo exige el diseño); si va sobre una
  junta circunferencial defectuosa, diseñar para cargas axiales y de flexión.
- **206-3.3 Pressure Design** — el t_req se calcula con el **código de
  construcción** (B31.3 / VIII-1). Material y esfuerzo admisible del sleeve por
  el código.
- **206-3.4 Sleeve Dimensions** — Type A y B: **≥ 100 mm (4 in) de largo** y
  **extender ≥ 50 mm (2 in) más allá del defecto**.
- **206-3.5 Type B Fillet Welds** — (a) filete **completo** si el espesor del
  sleeve **≤ 1,4× el espesor nominal del tubo**; (b) si es más grueso, extremos
  como están o achaflanados.
- **206-3.11** — considerar la **dilatación térmica diferencial** tubo/sleeve.
- **206-4.4** — si hay filetes de extremo circunferenciales, las **costuras
  longitudinales del sleeve van a tope, penetración completa**; prever venteo.

### Diseño de la hoja (reutiliza la maquinaria del 212)

- **Sección 1 — Entradas.** NPS + cédula del tubo portador → `t_portador`
  (reutiliza el lookup de `Datos_Ref` **ya robustecido con `""&`**); presión de
  diseño; longitud del defecto; espesor adoptado del sleeve `T_s`; selector
  **Type A / Type B**; selector de código (B31.3 / VIII-1) como en el 212.
- **Sección 7 — Material → S(T).** **Se reutiliza tal cual** la maquinaria de la
  Sección 7 del 212 (cascada + banda compacta + `S(T)` interpolado/tabulado,
  con conmutador de modo). El sleeve necesita su `S(T)` para el t_req.
- **Cálculo Type A:** `T_s ≥ ⅔·t_portador` (206-3.1). Dictamen CUMPLE/NO CUMPLE.
  No es componente a presión: **no** hay t_req de presión. Aviso 206-2.5
  (defecto circunferencial).
- **Cálculo Type B:** `t_req` de presión por el código activo (misma ecuación
  que `D65` del 212, con el `S(T)` y el `E_j` del sleeve) → `T_s ≥ t_req`
  (206-3.2/3.3). Regla de filete de extremo por el **1,4×** (206-3.5). Costura
  longitudinal a tope penetración completa (206-4.4).
- **Verificaciones comunes:** largo ≥ 100 mm y ≥ defecto + 2×50 mm (206-3.4);
  nota de dilatación diferencial (206-3.11).
- **Dictamen global** con la misma lógica de estados que el 212
  (ELIJA MATERIAL / FUERA DE RANGO / CUMPLE / REVISAR).

### Factorización a decidir (ver *Preguntas abiertas*)

La Sección 7 hoy vive dentro de `integrate_motor`. Para que los dos motores la
compartan sin duplicar fórmulas hay que **extraerla a una función**
`construir_seccion7_material(ws, fila_base, …)` reutilizable por ambos.

---

## Fase 3 — Dashboard: dos tarjetas bajo PCC-2 · SIN EJECUTAR

Hoy la banda 1 (`MOTORES DE CALCULO`) llega a **una** tarjeta:
`ASME → REPARACIONES → PCC → PCC-2 → Art. 212`.

- Añadir un `Nodo` en `ARBOL` de `build_db_materiales.py` para `Art. 206`,
  hermano de `Art. 212` bajo `PCC-2`. Por diseño del árbol, de ahí se derivan
  en preorden `HOJAS_NAV`, `DESTINOS`, `NAVEGABLES`, `PADRE`, `ROTULO`,
  `ANCLA_VOLVER` — no se toca el constructor de hojas ni la navegación.
- **Sincronía Python↔VBA:** `NAVEGABLES` (36 hojas) crece a 37;
  `HojasNavegables()` en `mod_nav.vba` **debe** crecer igual, o
  `test_dashboard.py::TestSincroniaPythonVba` falla. Es el único punto del VBA
  que crece con el árbol.
- Botón fijo `? MANUAL DE USO` y miga de pan salen solos del árbol.

---

## Fase 4 — Desanclar el motor 212 del Rev0 · DECISIÓN DE ALCANCE

Hoy la hoja `Parche_PCC2_Art212` (y `Datos_Ref`, `Instrucciones`) es **heredada**
del maestro (`HOJAS_HEREDADAS`), que sale del respaldo Rev0. La Fase 1 solo
escribe en el builder **las celdas que tocó**; el resto de la hoja sigue heredado.

Dos caminos, a elegir:

- **(a) Desanclado total:** mover toda la construcción de la hoja 212 al builder
  (geometría, secciones 1-6, formatos). El maestro deja de aportar la hoja.
  Máxima coherencia con D-4, pero es reescribir ~130 filas de fórmulas ya
  probadas; riesgo alto **sin poder validar en Excel aquí**.
- **(b) Desanclado incremental (recomendado):** el motor 206 nace 100% en código
  (ya cumple D-4 para lo nuevo); el 212 se mantiene heredado + *overrides* del
  builder, y se migran celdas al builder solo cuando haya que tocarlas. Riesgo
  bajo, avance seguro.

**Recomendación:** (b). El desanclado total del 212 no da beneficio funcional
inmediato y su riesgo es alto sin banco de Excel.

---

## Verificación — qué se puede aquí y qué no

- **Aquí (Linux, sin Excel):** `python -m pytest test_build_db.py test_dashboard.py
  test_secii_tablas.py`; `verificar.py §1-5` (conteos, unicidad, auditoría fila a
  fila, contigüidad de cascada, ausencia de matriz dinámica); build sin abortos;
  para el 206, extraer los **valores** (t_req, S(T), ⅔·t, 1,4×) con `uno`/openpyxl
  y compararlos a mano contra el código.
- **Solo en Excel de Windows (lado del ingeniero):** `verificar.py §6+` (recálculo
  real con `win32com`), F9 sobre el caso precargado, y la revisión visual de la
  hoja exportada a PDF/PNG (openpyxl miente sobre bordes de rango fusionado).
- **Regla dura:** ningún número normativo del 206 sale de memoria. `⅔`, `100 mm`,
  `50 mm`, `1,4×` son del texto del Art. 206; el t_req sale del código de
  construcción por la misma vía que el 212. Si un dato no está, se declara el
  vacío, no se inventa.

---

## Riesgos

- **Sin banco de Excel en este entorno.** Todo cambio de fórmula se entrega a
  validar. Mitigación: cambios mínimos, idempotentes, y validación de datos por
  `uno`/openpyxl antes de entregar.
- **Sincronía Python↔VBA del árbol.** Olvidar `HojasNavegables()` deja el libro
  abriéndose distinto de como se construyó. Lo ataja `TestSincroniaPythonVba`.
- **Límite de 25 continuaciones `_` en VBA.** Con 37 hojas navegables, revisar
  que la lista de `mod_nav.vba` siga concatenando (`s = s & "|…"`), no `Array(_…)`.
- **Reutilizar la Sección 7** exige extraerla a función sin cambiar el
  comportamiento del 212 ya probado; hacerlo con paridad de fórmulas y test.

---

## Preguntas abiertas (resolver antes de la Fase 2)

1. **Type B, ¿el `E_j` (junta longitudinal) es el del sleeve o el del tubo?**
   El sleeve tiene sus propias costuras longitudinales a tope (206-4.4); el t_req
   de presión usa el `E_j` de **esas** costuras, no las del tubo portador.
   Confirmar con el ingeniero la fuente del factor.
2. **¿Un solo NPS/cédula (tubo portador) o también dimensiones propias del
   sleeve** (OD/ID del sleeve)? El código pide `T_s` y el largo; el diámetro del
   sleeve se deriva del tubo + luz, como en el 212.
3. **Alcance del desanclado (Fase 4):** ¿(a) total o (b) incremental? (Recom: b.)
4. **¿El 206 lleva conmutador SI↔US de valores?** Los mínimos (100/50 mm, 1,4×)
   el código los imprime en los dos sistemas; el t_req y el S(T) ya conmutan por
   la Sección 7. Confirmar que basta con eso.

---

## Orden de ejecución sugerido

1. Validar la Fase 1 en Excel (ingeniero). Si algo falla, corregir antes de seguir.
2. Extraer `construir_seccion7_material()` de `integrate_motor` (con test de
   paridad de fórmulas del 212: nada cambia en el 212).
3. Construir `Collar_PCC2_Art206` (Type A/B) reutilizando esa función.
4. Añadir el `Nodo` Art. 206 al `ARBOL` + `HojasNavegables()` del VBA.
5. `pytest` + `verificar.py §1-5` + extracción de valores del 206 y cotejo manual.
6. Entregar Rev4 regenerado para validación en Excel del ingeniero.
