# Flujo de proceso — ASME PCC-2 Art. 212 (Parches soldados de filete)

Transcripción del diagrama de flujo aportado por el ingeniero
(`flujo_proceso_art_212.pdf`, 3 páginas, rotulado *ASME PCC-2–2022*). Es el
**flujo de referencia** contra el que debe alinearse el motor de cálculo del
Art. 212. Cada paso se cruza con la cláusula del Art. 212 en
`resources/asme_pcc/pcc_2/.../article_212_fillet_welded_patches.json`, que es
la fuente de verdad (Regla nº 1); si el flujo y `resources/` discreparan,
manda `resources/`.

> **Nota de Regla nº 1 sobre `resources/`.** La extracción del Art. 212 dejó
> como **imagen (bloque `figure`)** las ecuaciones (2) `F_LP = P·D_m/4` y parte
> de (5); el resto —(1), (3), (4), (6), (7)— sí está en la capa de texto. El
> coeficiente `4` de `F_LP` y la forma tabulada de (5) deben confirmarse contra
> el folio impreso antes de codificarse trazablemente, o extraerse de la imagen
> de la ecuación al JSON.

---

## 1. Alcance y elegibilidad (212-1 y 212-2)

| Criterio | Requisito del código | Condición de parada |
|---|---|---|
| Mecanismo de daño | Adelgazamiento local (corrosión, erosión) o perforación traspasante | ELEGIBLE si la tasa de daño es conocida; NO USAR si el daño no se puede caracterizar |
| Geometría | Cilíndrica, esférica, plana, cónica o conformada | ELEGIBLE (evaluar mangas si domina el comportamiento axisimétrico) |
| Temperatura de servicio | Desde > temp. de transición dúctil-frágil (nil-ductility) hasta máx. **345 °C (650 °F)** | < 0 °C: evaluar tenacidad a la entalla; > 345 °C: evaluar creep y fatiga |
| Grietas / crack-like | Prohibido en general para grietas activas o no analizadas | Solo si el crecimiento se arrestó/es predecible + análisis FFS (API 579) |
| Servicio letal | Servicio letal o fluidos de extrema peligrosidad | PROHIBIDO (usar Art. 201 inserto a tope o reemplazo de sección) |

## 2. Flujo del proceso — evaluación inicial y cargas

**PASO 1 — Elegibilidad y caracterización del daño (212-1 y 212-2)**
- Verificar mecanismo de daño = adelgazamiento local / corrosión / perforación.
- Confirmar T ≤ 345 °C (650 °F) y ausencia de servicio letal.
- Determinar dimensiones del área degradada actual y su proyección a fin de vida útil.
- Traslape mínimo del parche: cubrir metal base sano ≥ **25 mm (1 pulg.)** en todo el perímetro.

**PASO 2 — Cargas de presión y externas combinadas (212-3.2)**
- Fuerza circunferencial (hoop) unitaria: **F_CP = (P · D_m) / 2**  — ec. (1)
- Fuerza longitudinal unitaria: **F_LP = (P · D_m) / 4** — ec. (2)
- Cargas adicionales (flexión, momentos, viento, sismo): F_CO y F_LO.
- Fuerzas totales: **F_C = F_CP + F_CO** y **F_L = F_LP + F_LO**.

## 3. Diseño detallado — ramas y verificaciones físicas

**PASO 3 — Proximidad a discontinuidades (212-3.3)**
- Separación mínima: **L_min = 2 · (R_m · t)^(1/2)** — ec. (3)
- **Rama 3A** (dist ≥ L_min): parche rectangular/circular libre, esquinas redondeadas (R_min = 75 mm / 3 pulg.).
- **Rama 3B** (dist < L_min a boquilla): rodear la apertura 360° como placa de refuerzo con soldadura a penetración completa, O justificar con FEA.
- **Rama 3C** (parches adyacentes): mantener L_min entre bordes de soldadura de parches adyacentes.

**PASO 4 — Soldadura perimetral de filete (212-3.4A)**
- Pierna mínima: **w_min = F_max / (E · S_a)**, con **E = 0.55** — ec. (4).
- Límites: **w_min ≤ min(T_parche, t_envolvente)** y **w_min ≤ 40 mm (1.5 pulg.)**.
- Opción de biselado: aumentar garganta efectiva sin exceder los espesores base.

**PASO 5 — Excentricidad de carga en la junta de filete (212-3.4C)**
- Excentricidad nominal: **e = (T + t) / 2**.
- Esfuerzo combinado: **S_w = (P · D_m)/(2T) + [3 · P · D_m · e] / T²** — ec. (5).
- Criterio: **¿S_w ≤ 1.5 · S_a?** Si cumple → conforme; si no → modificar T, biselar o recalcular con menor presión.

**PASO 6 — Límite de conformado en frío (212-3.5)**
- Curvatura simple (cilindros): **% Elong = (50 · T / R_f) · (1 − R_f/R_o)** — ec. (7).
- Curvatura doble (cabezales/esferas): **% Elong = (75 · T / R_f) · (1 − R_f/R_o)** — ec. (6).
- Criterio: **¿% Elong ≤ 5.0 %?** Si ≤ 5 %: conformado directo; si > 5 %: PWHT post-conformado antes de soldar.

## 4. Fabricación, ensamblaje y NDE (212-4 y 212-5/6)

**PASO 7 — Fabricación y secuencia de montaje (212-4)**
- Corte mecánico o térmico; si térmico, esmerilar mín. 1.5 mm (1/16 pulg.) de material calcinado.
- Si T > 25 mm: inspeccionar bordes con MT/PT (laminaciones).
- Preparación de superficie: remover pintura/óxido/líquidos ≥ 40 mm (1.5 pulg.) a cada lado.
- Soldaduras existentes bajo el parche: esmerilar a ras e inspeccionar MT/PT.
- Fit-up: separación máx. 5 mm (3/16 pulg.). **Si g está entre 1.5 y 5 mm → recalcular e = (T + t + g)/2 y recalcular S_w.**
- Secuencia: 1º costuras internas del parche (si es dividido); 2º soldadura perimetral de filete.
- Venteo de gas para evitar sobrepresión entre placas durante cierre / PWHT.

**PASO 8 — Examen no destructivo (NDE) y prueba de hermeticidad (212-5 y 212-6)**
- NDE soldadura perimetral: 100 % MT o PT (ASME BPVC Sec. V).
- NDE costuras del parche: RT o UT en costuras a tope.
- Puntos de fijación temporal (orejetas, cuñas): MT/PT en zonas de retiro.
- Prueba de presión/hermeticidad: hidrostática o neumática según el código de post-construcción.
- Prueba neumática: aplicar cálculos de energía almacenada y distancia segura de ráfaga (Apéndice Obligatorio 501-II/III).
- Recubrimiento final solo tras aprobar NDE y prueba de presión.

## 5. Formulario de ecuaciones (Art. 212)

| Ecuación | Fórmula | Criterio / variables |
|---|---|---|
| (1) Fuerza circunferencial | F_CP = (P · D_m)/2 | P = presión de diseño; D_m = diámetro medio |
| (2) Fuerza longitudinal | F_LP = (P · D_m)/4 | fuerza unitaria longitudinal por presión |
| (3) Distancia de separación | L_min = 2·(R_m·t)^(1/2) | separación mínima a boquillas u otros parches |
| (4) Pierna de filete | w_min = F_max/(E·S_a) | E = 0.55; w_min ≤ min(T,t) y ≤ 40 mm |
| (5) Esfuerzo excentricidad | S_w = (P·D_m)/(2T) + 3·P·D_m·e/T² | S_w ≤ 1.5·S_a; e = (T+t+g)/2 (g = separación) |
| (6) Elongación curvatura doble | %Elong = (75·T/R_f)·(1 − R_f/R_o) | ≤ 5 % (si > 5 %, PWHT) |
| (7) Elongación curvatura simple | %Elong = (50·T/R_f)·(1 − R_f/R_o) | ≤ 5 % (si > 5 %, PWHT) |

Donde R_f = radio final de la línea media del parche; R_o = radio original de
la línea media (∞ si originalmente plano); T = espesor del parche; t = espesor
de pared del componente.
