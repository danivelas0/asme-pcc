# Flujo de proceso — ASME PCC-2 Art. 206 (Mangas / collares de encierro total de acero)

Transcripción del diagrama de flujo aportado por el ingeniero
(`asmepcc2articulo206flujo.pdf`, 3 páginas, rotulado *ASME PCC-2*). Es el
**flujo de referencia** contra el que debe alinearse el motor de cálculo del
Art. 206 (`Collar_PCC2_Art206`). Cada paso se cruza con la cláusula del Art. 206
en `resources/asme_pcc/pcc_2/p2_welded_repairs/art_206_full_encirclement_steel/
art_206.json` (y su espejo `part_2_welded_repairs/...`), que es la fuente de
verdad (Regla nº 1); si el flujo y `resources/` discreparan, manda `resources/`.

> **Nota de Regla nº 1 sobre `resources/`.** El Art. 206 **no trae ecuaciones
> propias** (`counts.equations = 0`): su diseño es (i) el t_req del código de
> construcción (206-3.3) y (ii) reglas geométricas con números que el propio
> artículo imprime en texto (⅔, 100 mm, 50 mm, 1,4×, 0,80/1,00, 2,5 mm,
> 50-80 %). El **único** dato que no está en la capa de texto es el **cateto del
> filete de extremo**: el texto de 206-3.5 (bloques 42-47) solo enuncia la regla
> de rama del 1,4×; las fórmulas `w = T_s + G` y `w_máx = 1,4·T_p + G` viven en
> las **figuras** `fig_206_3_5_1.png` y `fig_206_3_5_2.png` (bloques 72 y 77),
> cuyas leyendas (bloques 74-79) definen `G = gap`, `T_p = carrier pipe required
> minimum wall thickness`, `T_s = Type B sleeve nominal wall thickness`. **Leídas
> ambas imágenes de `resources/`**, confirman literalmente `T_s + G` (Fig. 1,
> para `T_s ≤ 1,4 T_p`) y `1,4·T_p + G` (Fig. 2, para `T_s > 1,4 T_p`, con chaflán
> opcional). Es el caso "imagen en `resources/`", idéntico al de la ec. (2) del
> Art. 212: se recupera leyendo el PNG, **no** un folio externo. **No hay ningún
> vacío real** en `resources/` para este artículo.

---

## 1. Alcance y clasificación de mangas (206-1)

El Art. 206 aplica al diseño e instalación de mangas de acero de **encirclement
completo** para tubería y ductos, soldadas en las dos costuras longitudinales
(bloques 2-4). Dos tipos estructurales según su capacidad de contener presión y
soportar cargas axiales:

| Característica | Type A (no resistente a presión) | Type B (resistente a presión) |
|---|---|---|
| Soldadura de extremos (206-1.1.1 / 206-1.1.2) | Extremos **no** soldados circunferencialmente | Extremos con **filete circunferencial completo** |
| Contención de presión | No contiene presión; solo refuerza | Contiene presión interna si el tubo fuga |
| Cargas axiales | No transmite cargas axiales/flexión | Restaura capacidad axial y de flexión |
| Aplicación típica | Defectos sin fuga, no evolutivos, con mecanismo/tasa de daño entendidos | Defectos con fuga o que pueden fugar; refuerzo axial en junta circunferencial defectuosa |

## 2. Precauciones y limitaciones (206-2)

| Cláusula | Requisito del código | Condición / acción |
|---|---|---|
| 206-2.3 Leaking Defects | Type B con fuga activa | Aislar la fuga **antes** de soldar (cruza con 206-4.3) |
| 206-2.4 Cyclic Operation | Ciclos frecuentes de presión / gradientes térmicos through-wall | Evaluación de **fatiga** por 206-3.8 (componente y filetes de extremo del sleeve) |
| 206-2.5 Circumferential Defects | Defecto orientado circunferencialmente | Type A **puede no ser apto** (no resiste cargas axiales) → considerar Type B |
| 206-2.6 Undersleeve Corrosion | Type A: ingreso de humedad por extremos no soldados | Evaluar corrosión bajo manga; si aplica, **sellante o recubrimiento** |
| 206-2.7 Weld Reinforcement | Costura previa (girth/longitudinal) con sobre-refuerzo prominente que impide el fit-up | (a) esmerilar + **RT/UT** del cordón esmerilado, o (b) manga con **abultamiento (bulge)**, Fig. 206-2.7-1 |

**Requisito de seguridad (206-2.3 & 206-4.3):** en fuga activa con Type B, aislar
la fuga y —en fluidos inflamables— purgar la cavidad anular con **nitrógeno** u
otro gas inerte antes de soldar.

## 3. Flujo metodológico de diseño (206-3)

**PASO 1 — Clasificación y selección de tipo (206-1, 206-2)**
Árbol de decisión del flujo:
- ¿El defecto presenta fuga o puede fugar? → **SÍ: Type B.**
- Si NO: ¿el defecto reduce la resistencia axial o su tasa de daño no es clara/entendida? → **SÍ: Type B | NO: Type A** (206-1.1.1 exige mecanismo y tasa de daño entendidos para Type A).
- ¿Hay costura circunferencial (girth weld) en la zona? → RT/UT del tubo → esmerilar a nivel **o** manga con abultamiento (206-2.7).

**PASO 2 — Espesor requerido (206-3.1, 206-3.2, 206-3.3)**
- **Type A:** `T_s ≥ (2/3)·T_p` (206-3.1). No contiene presión; garantiza soporte rígido externo del defecto. Los esfuerzos longitudinales del tubo portador deben cumplir el código de construcción.
- **Type B:** espesor ≥ el requerido para la **máxima presión de diseño admisible** (o resistencia equivalente plena del tubo si lo exige el diseño). Factor de eficiencia de junta longitudinal **E = 0,80**, o **E = 1,00 si la costura se examina 100 % por UT** (206-3.2). Si refuerza axialmente una junta circunferencial defectuosa, diseñar para cargas axiales y de flexión.
- **206-3.3 Pressure Design:** el `t_req` se calcula con el **código de construcción** (B31.3 / VIII-1); material y esfuerzo admisible del sleeve por el mismo código. **El sobreespesor de corrosión (C.A.) se aplica según el diseño de ingeniería.**

**PASO 3 — Dimensiones del sleeve (206-3.4)**
- `L_s ≥ máx(100 mm, L_defecto + 2 × 50 mm)`. La manga debe medir ≥ 100 mm (4 in) y sobrepasar el defecto ≥ 50 mm (2 in) a cada lado.

**PASO 4 — Filete de extremo Type B y ajuste (206-3.5, 206-4.1)**
- **Caso 1 (`T_s ≤ 1,4·T_p`):** filete completo de cateto **`w = T_s + G`** (Fig. 206-3.5-1, imagen de `resources/`).
- **Caso 2 (`T_s > 1,4·T_p`):** cateto **`w_máx = 1,4·T_p + G`**, extremo de la manga *as-is* o **achaflanado** (Fig. 206-3.5-2).
- El pie del cordón sobre el tubo portador debe transicionar suave, sin muesca filosa ni socavado (206-3.5, último párrafo).
- **Luz radial (fit-up):** se busca "no gap"; se admite una holgura radial de hasta **`G ≤ 2,5 mm (3/32 in)`** máx (206-4.1). Con holgura excesiva pueden requerirse pases de manteca (buttering).

**PASO 5 — Presión externa, cavidades y bulging (206-3.6, 206-3.7, 206-3.9, 206-3.10)**
- Considerar la **presión externa** sobre el tubo dentro del Type B; ajustar la manga lo más ceñida posible para minimizar el anular; si no, rellenar el anular con material endurecible o balancear la presión (206-3.6).
- **Daño externo / pérdida de pared externa:** rellenar las cavidades con **material endurecible (epoxi)** de resistencia a compresión adecuada para transferir la carga a la manga (206-3.7 / 206-3.9 / 206-3.10).
- Considerar reducir la presión de línea al instalar (206-3.9).

**PASO 6 — Fatiga y dilatación diferencial (206-2.4, 206-3.8, 206-3.11)**
- Todo Type B se evalúa para determinar si requiere **análisis de fatiga**; si se requiere, por ASME BPVC VIII-2, API 579-1/ASME FFS-1 u otra metodología equivalente (206-3.8).
- Considerar la **dilatación térmica diferencial** tubo/manga en ambos tipos (206-3.11).

## 4. Fabricación, ensamblaje, NDE y prueba (206-4, 206-5, 206-6)

**PASO 7 — Fabricación y soldadura en servicio (206-4)**
- Limpieza a **metal blanco** en toda la circunferencia a cubrir; rellenar con material endurecible las indentaciones/picaduras/vacíos; fit-up ceñido con abrazaderas hidráulicas o tirantes; holgura radial máx **2,5 mm** (206-4.1).
- El relleno no debe extruir a la zona de soldadura; quemarlo compromete el cordón (206-4.2).
- **Fuga activa:** aislar el defecto antes de soldar; en fluidos inflamables, **purgar con N₂/gas inerte** (206-4.3).
- Si hay filetes de cierre circunferenciales, las costuras longitudinales del sleeve van **a tope, penetración completa**, y se prevé **venteo** en el cordón de cierre final; técnica de **bajo hidrógeno** (206-4.4).
- **Reducir la presión de operación a 50-80 %** durante la instalación, manteniendo flujo; ver API RP 2201 (206-4.5). Considerar burn-through.
- **Soldadura en servicio (Art. 210):** calificar el proceso contra (a) agrietamiento por hidrógeno en la ZAC, (b) ZAC dura por la química de base, (c) burn-through (206-4.6). Procedimientos y soldadores por Sección IX / código de post-construcción (206-4.7).

**PASO 8 — Examen (NDE) y prueba de hermeticidad (206-5, 206-6)**
- Inspeccionar **todos los fit-ups** antes de soldar; VT de todas las soldaduras (206-5.1).
- **Type A:** VT de la raíz durante la soldadura; costuras longitudinales por **PT, MT o UT** al terminar (206-5.2).
- **Type B:** UT del **material base del tubo portador** (espesor, grietas, laminaciones) en la zona de los filetes circunferenciales; costuras longitudinales inspeccionadas al terminar; primer y último pase de los circunferenciales por **MT o PT**; donde haya riesgo de agrietamiento diferido, **NDE ≥ 24 h** después (≥ 48 h para soldadura en servicio con alta probabilidad de H₂) (206-5.3).
- **Prueba de hermeticidad (206-6), solo Type B, si el propietario la exige:** (a) presurizar el **anular** entre manga y tubo, con presión de prueba tal que **el tubo interno no colapse** por presión externa; o (b) **prueba sensible de fugas** (B31.3 345.8 u otro estándar). **El Art. 501 da guía adicional.** *(El texto del 206 no publica ecuaciones de energía almacenada; el motor mantiene la prueba fiel al 206, sin importar el motor neumático del App. 501 — Regla nº 1.)*

## 5. Formulario de reglas de cálculo (Art. 206)

| Regla / módulo | Fórmula o criterio | Fuente en `resources/` |
|---|---|---|
| Espesor Type A | `T_s,min = (2/3)·T_p` | 206-3.1 (texto, bloque 31) |
| Espesor Type B | `t_req` por el código de construcción (B31.3 / VIII-1) **+ C.A.**; `E = 0,80` (o `1,00` si 100 % UT) | 206-3.2 / 206-3.3 (texto, bloques 35-38) |
| Longitud de manga | `L_s,min = MÁX(100, L_defecto + 2·50)` [mm] | 206-3.4 (texto, bloque 40) |
| Cateto filete (`T_s ≤ 1,4·T_p`) | `w = T_s + G` | **Fig. 206-3.5-1** (imagen, bloque 72) |
| Cateto filete (`T_s > 1,4·T_p`) | `w_máx = 1,4·T_p + G` (chaflán opcional) | **Fig. 206-3.5-2** (imagen, bloque 77) |
| Holgura radial (fit-up) | `G ≤ 2,5 mm (3/32 in)` | 206-4.1 (texto, bloque 63) |
| Reducción de presión al soldar | `0,50·P_op ≤ P_instal ≤ 0,80·P_op` | 206-4.5 (texto, bloque 82) |

Donde `T_s` = espesor nominal de la manga; `T_p` = espesor del tubo portador
(nominal para la rama del 1,4×, texto 206-3.5(a); la leyenda de las figuras lo
rotula "required minimum wall thickness" — matiz a confirmar, ver el plan);
`G` = luz radial manga-tubo; `L_defecto` = longitud axial del daño.
