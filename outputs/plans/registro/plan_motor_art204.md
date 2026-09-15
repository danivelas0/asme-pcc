# Plan — Motor de cálculo del Art. 204 de ASME PCC-2 (Welded Leak Box Repair)

> Fecha: 2026-09-15. Fuente: `resources/ASME PCC/pcc_2/p2_welded_repairs/art_204_welded_leak_box/art_204.json`
> (ASME PCC-2 **2022**, Part 2, Art. 204, páginas 45–48; 106 bloques).

---

## Estado de ejecución final (2026-09-15) — CERRADO

Ejecutado de punta a punta por orden explícita del ingeniero, tras la parada
anterior. `MOTORES_DECLARADOS = (MOTOR_ART_204,)`: el libro carga el motor,
`Caja_PCC2_Art204` existe (rama **TE + CAPS**, alcance de esta primera
entrega), y el árbol de navegación lo expone con sus tres documentos.

**Definición de terminado (las 5 condiciones de la skill motor_pcc2):**

1. ✅ Hoja `Caja_PCC2_Art204` construida por el builder: cascada de material a
   dos columnas (D metal base, E material de la caja), cascada dimensional
   norma→NPS→cédula contra `DB_B36_10`/`DB_B36_19`, conmutador SI/US, leyenda,
   semáforo, botón de reinicio, área de impresión.
2. ✅ Sus dos pestañas: `Espec_PCC2_Art204` e `Instruc_PCC2_Art204`.
3. ✅ Nivel `NAV_CAL_ART204` en el árbol (derivado) y sus tres líneas en
   `HojasNavegables()` de `scripts/vba/mod_nav.vba` — 47 hojas navegables
   (antes 44 con el Art. 204 sin registrar). `make_vba_seed.py` corrido y
   `templates/maestro_con_macros.xlsm` regenerado.
4. ✅ `pytest test_build_db.py test_dashboard.py test_motor_declarado.py`:
   **407 pruebas, 0 fallos.**
5. ✅ `verificar.py --resources ... --wb ...` (sin `--pdfs`, no hacía falta
   para la §12): **0 fallos en las 17 secciones**, incluida la §12 nueva
   ("Motores declarados (maquinaria, no ingeniería)": 1 motor, 0 fallos).

**Extensiones de chasis que esta ejecución añadió** (no previstas en la
versión anterior del plan):

- **Etapa 1c completada**: el bucle en `build_motor_declarado()` que, tras
  `emitir_tabla`, recorre `motor.dimensionales`, llama a
  `_materializar_cascada_b36` con columnas nuevas del contador compartido
  (`siguiente_col`), y escribe en `od_te`/`esp_te` el `INDEX/MATCH` con
  `_choose_b36` y el conmutador SI/US. Probado end-to-end: el motor real
  produce un desplegable de verdad, no solo pasa los guardias.
- **`{UMBRAL:clave}` dentro de una fórmula declarada**, mecanismo gemelo del
  centinela `{S}` de la Etapa 1b: `motor_declarado.resolver_umbrales()` +
  `nombres_de_umbral()` + `_centinela_umbral()`, resuelto en
  `build_motor_declarado()` con `umbral(umbrales, clave, es_si)` (que YA
  conoce `build_db_materiales`, sin crear un import circular). Sin esto no
  se podían declarar los avisos de pared delgada del Art. 210 (204-4.4) sin
  teclear la cifra normativa a mano.
- **`_citas_del_motor()` exime de cita obligatoria a las filas `od`/`espesor`
  de una `Dimensional`**: su procedencia es la auditoría fila a fila de
  `DB_B36_10`/`DB_B36_19` (§11 de `verificar.py`), no un párrafo de texto;
  exigirles además una `Cita` a un bloque pedía una trazabilidad que no les
  corresponde.
- **Regresión real destapada y corregida**: `t_req_te` (eq. 3a) depende de
  `{S@E}` (esfuerzo admisible del material de la caja), y la columna E no
  tiene semilla (`Material.semilla` solo siembra la columna D). En el caso
  semilla sin material de caja seleccionado, `S` es `NA()` a propósito (mismo
  patrón que el `S_t` del sleeve en `Collar_PCC2_Art206`), y ese `#N/A`
  llegaba crudo a `chk_espesor_te`/`chk_refuerzo` — un error de Excel que
  ningún semáforo reconoce. `verificar.py` §12 lo encontró (2 fallos:
  `D33`/`D43` pintando gris "sin calcular" en vez del color prometido). Se
  envolvieron las dos verificaciones en `IFERROR(...)`, degradando a un texto
  `"REVISAR - seleccione el material de la caja"` legible y coherente con el
  semáforo, sin inventar ningún valor normativo.

**Huecos declarados** (los 9 de la sección 8 original, todos presentes en el
motor — comentario de celda, `Especificacion` o entrada hand-off, según el
caso; ninguno se calculó con un coeficiente inventado):

1. El Art. 204 no publica ecuaciones (límite de la fuente).
2/2bis/2ter. BPVC Sección VIII no extraída y B16.9 sin capa de texto/OCR:
   `rating_te` es entrada hand-off (fuera de la ruta crítica, `para. 303`
   autoriza calcular con `304`), el espesor del cierre (`t_h`) es entrada
   hand-off citando la Tabla 304.4.1-1 → UG-32/33/34, y no se emite la
   eq. (15) de un *blank* para una tapa soldada.
3. Resistencia del filete perimetral frente al empuje axial: entrada
   `empuje_axial` + verificación `chk_empuje_axial` (hand-off, por defecto
   REVISAR).
4. Colapso por presión externa: entrada `colapso_externo` +
   `chk_colapso_externo` (hand-off, por defecto REVISAR).
5. Longitud mínima de la caja: declarada en `Especificacion`, sin emitir
   `2√(Rt)`.
6. Figura 204-1-1: `Especificacion` declarando que es fotografía sin dato
   acotado.
7. Decisión 2 (Lista.cita): el guardia comprueba que el bloque citado existe
   y publica texto, no que cada opción aparezca literalmente (documentado en
   el comentario de `tipo_caja`/`metodo_prueba`).
8. El `4,8 mm` / `6,4 mm` del Art. 210 entran como `{UMBRAL:...}` — avisos,
   nunca compuertas.
9. Tabla 210-4.2.1-1 sin reconstruir: no se usa en este motor (gobierna
   probetas de calificación, fuera de alcance de esta entrega).

**Caso semilla**: nace en `REVISAR` (esperado — hay compuertas duras y
hand-offs sin confirmar, mismo patrón que el 212 tras su Fase 2).

**Pendiente para una entrega futura, fuera del alcance de este plan**: la
rama "codo + tapas planas" (Decisión 5 del plan original), y la Etapa 0a
(extracción de B16.9, bloqueada por falta de OCR) si algún día se resuelve.

---

## Estado de ejecución previo (2026-09-15, histórico — superado por lo de arriba)

Se ejecutó parcialmente por orden del ingeniero y se **detuvo a petición suya**.
El libro **se construye exactamente igual que antes**: `MOTORES_DECLARADOS` sigue
vacío, así que nada de lo hecho cambia el `.xlsm`. **436 pruebas en verde.**

| Etapa | Estado | Qué quedó |
|---|---|---|
| **0a** — extraer B16.9 | ❌ **BLOQUEADA** | El PDF de B16.9-2024 **no tiene capa de texto**: 46 páginas escaneadas, 0 caracteres, y no hay OCR instalado. Ver «El tope de B16.9» abajo. |
| **0b** — reparar la ec. de CE del Art. 210 | ✅ **HECHA** | `completar_art_210_ce.py`. Idempotente, los dos espejos, SHA-256 del PDF en `extraction_amendments`. |
| **0c** — reconstruir la Tabla 210-4.2.1-1 | ⏸ **no empezada** | `structure: not_reconstructed`. Es periférica: gobierna probetas de calificación. |
| **1** — chasis multi-fuente | ✅ **HECHA** | `Cita.fuente`, `bloques_citables()`, citas de `Verificacion`/`Especificacion` validadas, `Lista.cita`, `UMBRALES_PCC2` por ruta. |
| **1b** — `{S}` dentro de una ecuación | ✅ **HECHA** | Extensión nueva, no prevista en el plan original. Ver abajo. |
| **1c** — cascada dimensional en el chasis | ⚠️ **A MEDIAS (cerrada en la ejecución final, ver arriba)** | Ver «El punto exacto de reanudación». |
| **2** — declarar el motor | ⏸ **no empezada (hecha en la ejecución final, ver arriba)** | `motores/art_204.py` no existe. |
| **3** — verificar y entregar | ⏸ **no empezada (hecha en la ejecución final, ver arriba)** | No se ha regenerado el `.xlsm` ni corrido `verificar.py`. |

### Lo que la ejecución añadió y el plan original no preveía

**Etapa 1b — `{S}` de la cascada, usable dentro de una ecuación.** La plantilla de la
skill lo declaraba como límite: *«si el artículo necesita `{S}` dentro de una ecuación,
eso es una extensión del chasis que todavía no existe»*. El Art. 204 la necesita —la
eq. (3a) es `t = PD/(2(SEW + PY))`, sin `S` no hay ecuación—. Resuelto con el mismo
patrón que ya recomponía el dictamen: el bloque de material se construye **después** de
la tabla, así que la fórmula sale con un **centinela** (`__MAT@S@D__`, con `@` a
propósito: Excel no lo admite en una referencia, de modo que un centinela que
sobreviviera revienta en vez de calcular contra otra celda) y `resolver_material()` lo
cambia por la dirección real. Nombres disponibles: `{S}`, `{Tmax}`, `{material_id}`,
`{dictamen_material}`, con `{S@E}` para una segunda columna de material.

### El punto exacto de reanudación — etapa 1c

Para que el ingeniero **elija un tamaño de te**, la regla 14 exige desplegable contra la
base que lo tabula. El mecanismo existía y estaba probado (`_materializar_cascada_b36`:
norma → NPS → cédula) pero **cableado a los dos motores escritos a mano**: una sola
tanda de columnas ocultas (`COL_NORMA_B36` = 25..28) y la columna `D` fija. El Art. 204
necesita **dos** cascadas —componente portador y envolvente—, y dos llamadas se pisarían.

Hecho ya:
- `Lista.cascada` (tercera forma de `Lista`) y su guardia en `comprobar_listas()`.
- `Dimensional(norma, nps, cedula, od, espesor)` y `Motor.dimensionales`.
- `comprobar_dimensionales()`, con las dos direcciones: un nivel que nombra una fila que
  no existe aborta, y una fila marcada como nivel que **ninguna** `Dimensional` recoge
  también —se quedaría sin desplegable y sin decirlo—.
- `_materializar_cascada_b36(..., col_base=None, col_valor="D")`: los defaults dejan a
  los dos motores escritos a mano exactamente como estaban.
- El contador de columnas ocultas se comparte entre `emitir_tabla` y las cascadas
  (`_helpers_del_libro(ws, wb, siguiente_col)`), o la segunda cascada pisaría las listas
  de la primera y el desplegable ofrecería las cédulas de otro NPS.
- El helper de listas **no** materializa un nivel de cascada: lo escribe su constructor.

**Falta —y es lo único de la 1c—:** el bucle en `build_motor_declarado()` que, tras
`emitir_tabla`, recorre `motor.dimensionales`, llama a `_materializar_cascada_b36` con
cuatro columnas nuevas del contador compartido, y escribe en las filas `od`/`espesor` el
`INDEX/MATCH` con `_choose_b36(idx, ...)` y el conmutador SI/US. Hoy declarar una
`Dimensional` pasa los guardias y **no produce cascada**. Ningún motor la declara, así
que no afecta al libro; pero la funcionalidad está a medias y ese es el primer paso al
reanudar.

### El tope de B16.9

`ASME B16.9 - Factory-Made Wrought Buttwelding Fittings 2024.pdf` **existe** en la
carpeta de standards, pero es un **escaneo sin capa de texto** (46 páginas, una imagen
por página, `get_text()` devuelve 0 caracteres) y no hay `tesseract` instalado.
Transcribir a ojo decenas de filas × columnas de dimensiones para alimentar un motor de
cálculo **no es alternativa aceptable**: un dígito mal leído cambia una dimensión y nada
lo avisa.

**No bloquea el desplegable de tes**, porque una te de tope se suelda a tubería y su OD
y su espesor de pared son los del tubo de igual NPS y cédula, que sí están en
`DB_B36_10` (779 filas) y `DB_B36_19` (114), auditadas. Lo que B16.9 añade y B36 no es
la dimensión **centro-a-extremo**, que gobierna el *largo* de la caja —y el largo lo
verifica 204-3.11 contra la distancia a metal sano **medida en campo**, que es una
entrada de todos modos—.

> Matiz que se **declara y no se afirma como código**: que un accesorio B16.9 se fabrique
> al espesor de la tubería coincidente lo dice B16.9, que no está en `resources/`. Va en
> el comentario de la celda, no en la columna de referencia.

### Corrección al propio plan: el `4,8 mm` del Art. 210

**La versión anterior de este plan lo llamaba «umbral», y no lo es.** El Art. 210
bloque 3 dice:

> *«Welding onto pressure components or pipelines with thin walls [**e.g.**, 4.8 mm
> (0.188 in.) **or less**] **is possible as long as precautions are taken**.»*

Es un **ejemplo de pared delgada** que pide precauciones, no un mínimo de aceptación. El
bloque 6 trae otro (`< 6,4 mm (0.250 in.)` → puede hacer falta bajar el aporte térmico).
**No hay ningún «shall be ≥ 4,8 mm» en el Art. 210.** Las demás cifras en mm de ese
artículo son criterios de aceptación de **probetas de calificación** (socavado 0,8 mm,
poro 1,6 mm, escoria 0,8/3,2/12,7 mm) y el corte de 12,7 mm de la Tabla 210-4.2.1-1 para
el número de probetas.

En el motor la fila de burn-through es por tanto un **aviso**, nunca una compuerta.
Bloquear un diseño en 4,8 mm inventaría un requisito que el código no pone.

---

## 0. El hallazgo que gobierna todo el plan

El barrido completo del artículo (Fase 1 de la skill `motor_pcc2`) da esto:

```
counts: {blocks: 106, figures: 1, tables: 0, equations: 0}
tipos:  67 paragraph · 38 section_header · 1 figure
umbrales de doble unidad "n mm (n in.)":  NINGUNO
cifras en todo el texto:  1..14 y "204"  (numeración de párrafos, nada más)
```

**El Art. 204 no publica ni una ecuación, ni una tabla, ni una cifra.** La única
figura (bloque 11, `fig/fig_204_1_1.png`) se abrió con `Read`: es una **fotografía**
de una caja soldada sobre una te, sin acotar. No es una extracción rota — el primer
sitio donde este proyecto mira antes de declarar una rotura es la imagen, y ahí no
hay dato. Los dos espejos del árbol son idénticos salvo la ruta de la imagen.

El artículo **delega** el dimensionamiento, y lo dice dos veces:

> **204-3.5** (b39): *«Leak repair boxes and attachment welds shall be designed for
> design conditions and anticipated transient loads…, following the design
> requirements of the construction or post-construction code.»*
>
> **204-3.6** (b41): *«In cases where there are no applicable design requirements,
> the principles of the applicable construction code or post-construction code
> shall be followed.»*

**Consecuencia:** el Art. 204 es el primer artículo del libro cuyo motor es
mayoritariamente de **elegibilidad, compuertas, verificaciones y procedimiento**.
El único cálculo que el propio artículo publica es el criterio relativo de
204-5.2(a) (σ < ½·S). Todo lo demás, o sale del código de construcción —vía chasis
multi-fuente— o queda declarado como hueco.

Es además **el primer motor declarado del libro**: `motores.MOTORES_DECLARADOS`
está hoy vacío y el chasis nunca se ha ejercido contra un artículo real.

---

## 1. Evaluación del flujo propuesto por el ingeniero

| Paso propuesto | Veredicto | Qué publica de verdad `resources/` |
|---|---|---|
| 1. Elegibilidad por grietas | ✅ real, ❌ incompleto | 204-2.2 (b13–b17) da **cuatro** vías en OR; el flujo implementa solo la (c). Rechazaría casos que el código admite por (a), (b) o (d). |
| 2. Burn-through ≥ 4,8 mm | ❌ el número no es del 204 | 204-3.11 (b61) exige *«sufficient wall thickness… to avoid burn-through»*, **sin cuantificar**. `4,8 mm (0.188 in.)` y el electrodo `2,4 mm (0.094 in.)` los imprime el **Art. 210** (verificado en su JSON). |
| 3. `t_req = P·R/(S·E − 0,6·P) + CA` | ❌ no es del 204 | Es UG-27(c)(1) de **BPVC VIII-1**, que **no está en `resources/`** (solo Sec. II A/B/C/D). Del 204 sí salen E (b48) y CA (b44) — los parámetros, no la ecuación. |
| 4. `F = P·π/4·D_o²`, `w = F/(π·D_o·0,55·S_a)` | ⚠️ requisito real, coeficientes inventados | 204-3.9(a) (b53) sí exige considerar el empuje por separación circunferencial total **y permite renunciar a él** si la resistencia remanente al fin de vida basta. El **0,55 no aparece en el artículo**; el área con `D_o²` sobreestima; el tope de 40 mm del cateto es del **Art. 212**. |
| 5. Colapso por sellante / presión externa | ✅ requisito real, ❌ ecuación ajena | 204-3.12 (b63) y 204-6.3 (b105) lo exigen. UG-28 es de VIII-1: fuera de `resources/`. |
| 6. `L = L_deg + 2·(2√(R·t))` | ❌ | 204-3.11 dice *«sufficiently long to extend to a sound area»*, sin fórmula. El `2√(Rt)` es longitud de atenuación clásica, no una afirmación de este artículo. Los 100/50 mm son del Art. 206. |

**Omisiones del flujo que sí pertenecen al motor:** fluido **estancado** (b26), vida
de diseño (b28), los **cinco modos de fallo nuevos** que introduce la caja
(b31–b35), temperatura mínima y tenacidad (b37), la **prohibición dura** de soldar
al *knuckle* de cabezales formados (b42), peso de la caja **con el fluido atrapado**
(b49), expansión diferencial (b50), viento/sismo/golpe de ariete (b54), vents y
drenajes (b57–b59), *seepage* del sellante (b66), juntas de expansión/rótula (b68),
todo el 204-4, todo el 204-5 y todo el 204-6.

---

## 2. Decisiones que este plan toma, y por qué

### Decisión 1 — El chasis pasa a admitir **varias fuentes** por motor

Hoy `comprobar_procedencia()` (`motor_declarado.py:336`) valida **toda** `Cita`
contra `motor.fuente`, un solo JSON. `Cita.archivo` existe pero es **decorativo**:
solo se imprime en los mensajes de error y en la tabla de trazabilidad
(`motor_declarado.py:347,361,365`).

El cambio es hacer `Cita.archivo` **portante**: cada cita se valida contra su propio
archivo, con `motor.fuente` como valor por defecto cuando `archivo` queda vacío.
Con eso el motor del 204 puede citar el B31.3 para la cadena que el 204 delega.

**Obstáculo real:** el B31.3 **no usa la clave `blocks`**. `chapter_02.json` tiene
`preamble` + `sections`, cada `section` con `id`, `heading`, `paragraphs`
(`{kind, marker, indent, text}`) y `sections` anidadas. Hace falta un lector
normalizador (ver Fase A.2).

### Decisión 2 — Una enumeración que publica el CÓDIGO puede alimentar un desplegable

La regla 12 del libro prohíbe la lista tecleada, y `comprobar_listas()` solo admite
`Lista(opciones=…)` en `CLAVES_CON_ENUMERACION = ("unidad", "modo")`. Pero el 204
publica dos enumeraciones cerradas que **ninguna base de datos tabula y el propio
código imprime**:

- **tipo de caja**: no estructural / estructural — 204-1(f), b7;
- **método de prueba**: los cuatro de 204-6.2 — b99–b103.

Lo que la regla 12 ataca es la lista *«que no se audita, no se actualiza si la base
cambia, y puede ofrecer un material que la base ni siquiera admite»*. Una
enumeración **con `Cita` al bloque que la imprime** sí se audita. Se propone añadir
`Lista.cita` y admitir `opciones` fuera de las claves reservadas **solo** con cita
válida, que `comprobar_procedencia()` verifica como cualquier otra (existe, y no es
un `section_header`).

**Límite declarado:** el guardia comprueba que el bloque citado existe y publica
texto; **no** comprueba que cada opción aparezca literalmente en él (el libro está
en español y el código en inglés). Eso queda escrito en el comentario de la celda.

### Decisión 3 — El Art. 210 entra, y aporta más de lo previsto ✅ APROBADA Y EJECUTADA

204-4.4 (b79) exige que la calificación del procedimiento para soldadura en servicio
*«properly address preheat temperature, weld cooling rate, the risk of
burn-through…»*, y el Art. 210 es el artículo de PCC-2 sobre exactamente eso.
`UMBRALES_PCC2` acepta ya una **ruta** además de un artículo declarado, y las dos
cifras del 210 se leen en sus dos unidades desde `resources/`:

```
210_pared_delgada          SI=4.8   US=0.188
210_pared_reducir_aporte   SI=6.4   US=0.25
```

**Pero no son umbrales de aceptación** —ver «Corrección al propio plan» arriba—: entran
como **aviso de precauciones**, con el texto sin citar la cifra (regla ya establecida:
un aviso que dice «excede 40 mm» es falso en modo US).

Lo que el Art. 210 sí aporta como cálculo y como requisito:

| | Bloque | Qué |
|---|---|---|
| **Carbono equivalente** | b45, **reparado** | `CE = C + Mn/6 + (Cr + Mo + V)/5 + (Ni + Cu)/15`. Cálculo real y citable. Agrupa el procedimiento de soldadura en servicio (210-4.1.1.3). |
| **Tabla 210-4.2.1-1** | b67 | Tipo y número de probetas por espesor (corte en 12,7 mm) y tipo de soldadura. La de extremo de la caja es de **filete**, así que la fila aplica directo. Pendiente de reconstruir (etapa 0c). |
| **Variables esenciales** | 210-4.1.1 | Potencial de enfriamiento, CE, consumible, aporte térmico, *bakeout*. |

> **El Art. 208-3.4 publica OTRA definición de CE** —`C + (Mn + Si)/6 + …`, con silicio—
> para el mismo código. Son distintas a propósito y no se copia una en la otra. Es el
> mejor argumento a favor de haber derivado la del 210 de la geometría del PDF en vez de
> escribir «la fórmula de CE que uno recuerda»: la de memoria habría sido la del 210 en
> el sitio del 208, o al revés.

### Decisión 5 — La envolvente se elige entre **dos configuraciones**, y eso define el cálculo

Premisa del ingeniero (2026-09-15): la caja se fabrica normalmente con **te + caps** o
con **codo + tapas planas**. No es una restricción del código —204-1(c) admite
cilíndrica, rectangular, con cabezal plano o formado (b4), y 204-1(d) dice que se
hacen *«by welding split pipe, pipe caps, or plates»* (b5)— pero acota el motor a lo
que de verdad se construye, y con eso el cálculo deja de ser genérico.

**Orden de construcción fijado por el ingeniero (2026-09-15): primero te + caps.** Es
la configuración que más se usa en refinería, especialmente para fugas de vapor, y es
además la que deja el cálculo dentro de `resources/`.

**`303` autoriza calcular un componente listado, y eso cambia el plan.** El párrafo dice:

> *«Components manufactured in accordance with standards listed in Table 326.1.1-1 shall
> be considered suitable for use at pressure–temperature ratings… The rules in para. 304
> are intended for pressure design of components not covered in Table 326.1.1-1, **but
> may be used for a special or more-rigorous design of such components**…»*

Decisión del ingeniero: **las dos vías en paralelo en la misma hoja** —selección del
componente listado por *rating* (`303`), y las ecuaciones de `304` como **cálculo de
respaldo y verificación**—. La verificación compara una contra otra: el espesor real de
la te (de `DB_B36_10` por NPS + cédula) contra el `t_req` de la eq. (3a).

Consecuencia práctica: el *rating* de B16.9 **sale de la ruta crítica**. Era la reserva
principal, porque B16.9 no se puede extraer (PDF escaneado).

**Lo que cambia respecto de la versión anterior del plan:** la eq. (3a) de tubo recto
deja de ser la única vía. Con accesorios estándar la cadena es otra, y está mejor
publicada:

| | **Te + caps** | **Codo + tapas planas** |
|---|---|---|
| El accesorio | componente **listado**: `303` / `304.7.1` → **rating**, no fórmula | ídem |
| El cierre | cap formado: `304.4.1(b)` eq. (13) `t_m = t + c` → Tabla 304.4.1-1 → UG-32/UG-33 | tapa plana → Tabla 304.4.1-1 → **UG-34** |
| La abertura de paso | `304.4.2` + `304.3.3` — **cálculo cerrado en `resources/`** | ídem |
| Compuerta de tamaño | abertura ≤ ½ D_int del cierre (`304.4.2(a)`); mayor → reductor `304.6` | ídem; si el cierre es plano y la abertura mayor → brida `304.5` |
| Compuerta dura | **204-3.6 (b42)**: no soldar al *knuckle* del cap sin calificación **con fatiga** | aplica si la tapa es formada |

**El corte en el cap es el cálculo que el 204 nombra por su nombre.** 204-3.6 (b41)
dice *«…fabricated by machining standard fittings (such as cutting-out an opening in
standard pipe caps to make end pieces) shall be qualified by analysis or testing…,
and be reinforced if necessary»*. Su contrapartida en el B31.3 es
`304.4.2 Openings in Closures`, cuyo apartado (e) manda calcular área y zona de
refuerzo **por `304.3.3`**, tratando el cierre como el *header*. Y `304.3.3` está
entero en `resources/`:

```
(6)   A1 = t_h · d1 · (2 − sin β)        área de refuerzo REQUERIDA
(6a)  A2 + A3 + A4 ≥ A1                  criterio
(7)   A2 = (2·d2 − d1)(T_h − t_h − c)    disponible en el cierre
(8)   A3 = 2·L4·(T_b − t_b − c)/sin β    disponible en la conexión
```

`304.4.2(c)` añade el reparto: ≥ ½ del área requerida a cada lado del plano.

**Límite declarado:** `304.4.2(d)` remite además a **UG-37(b), UG-38 y UG-39** de
VIII-1 para el área requerida. Se emite la vía B31.3 —eq. (6), que es la que el
propio Código publica y la que (e) manda usar para el área disponible— y ese remite
queda escrito en el comentario de la celda, no silenciado.

**La Tabla 304.4.1-1 está extraída limpia (6 filas) y es una tabla de REFERENCIAS,
no de fórmulas**: mapea tipo de cierre → párrafo de VIII-1. Así que el espesor del
cierre queda como hand-off **con la cadena de cita completa impresa en la columna de
referencia**: `204-3.5 → B31.3 304.4.1(b) eq.(13) → Tabla 304.4.1-1 → UG-34`. El
ingeniero ve exactamente por dónde entrar; el motor no inventa el número.

### Decisión 4 — Lo que sigue sin fuente queda **hand-off declarado**

Ni con multi-fuente hay en `resources/` una ecuación publicada para:

- la **resistencia del filete perimetral** frente al empuje axial (el 0,55 no existe);
- el **colapso por presión externa** (UG-28 es de VIII-1, no extraído).

Esas dos entran como **entrada del ingeniero** (valor calculado aparte), con el texto
del 204 que las exige en su columna de referencia, el hueco en el comentario de la
celda y en `ISSUES`. Es el patrón de la Tabla TE-2 y de las Tablas B-2..B-6: el hueco
se escribe, no se calla.

---

## 3. Cobertura — los 106 bloques a su destino

Los 38 `section_header` son rótulos y no se citan (el chasis lo rechaza). Quedan
67 `paragraph` + 1 `figure`; **los 68 tienen destino**.

### 204-1 Descripción · 204-2 Limitaciones

| Bloque | Cláusula | Destino |
|---|---|---|
| 2,3,4,5,6 | 204-1(a)–(e) | Pestaña de instrucciones: qué es y cuándo se usa |
| **7** | 204-1(f) | **Fila LISTA `tipo_caja`** (no estructural / estructural) — gobierna la vía (c) de 204-2.2 |
| 10 | 204-2.1 | Especificación: la Parte 1 aplica en conjunto |
| 11 | Fig. 204-1-1 | Instrucciones: ilustración. **Hueco declarado:** fotografía, sin dato acotado |
| **13** | 204-2.2 | **Compuerta `elegibilidad_grieta`** (regla base) |
| **14,15,16,17** | 204-2.2(a)(b)(c)(d) | Cuatro entradas Sí/No en **OR** que abren la compuerta |
| 19 | 204-2.3 | Especificación + Paso 1: personal calificado en condiciones representativas |
| 21,22,23 | 204-2.4 | Paso 1: revisión de peligros; avisos |

### 204-3 Diseño

| Bloque | Cláusula | Destino |
|---|---|---|
| 26 | 204-3.1 | Verificación `material_estancado` + aviso de modo de deterioro |
| 28 | 204-3.2 | Entrada `vida_diseno` + aviso |
| 30–35 | 204-3.3 | Cinco avisos `modo_fallo_*`: pernos encapsulados, condensación, punto de rocío, fluencia de pernería, expansión constreñida |
| 37 | 204-3.4 | Verificación `temp_toughness` |
| **39** | 204-3.5 | Ancla del **hand-off**: cita del selector de código de construcción |
| 41 | 204-3.6 | Aviso `qualif_fittings` (piezas mecanizadas de accesorios estándar) |
| **42** | 204-3.6 | **Compuerta `knuckle`**: no soldar al nudillo de cabezales formados sin calificación |
| 44 | 204-3.7 | Entrada `CA` |
| 46,47 | 204-3.8(a) | Entradas P y T + aviso de presiones/temperaturas coincidentes máx. y mín. |
| **48** | 204-3.8 | Entrada `E` + verificación `E_vs_NDE` |
| 49 | 204-3.8(b) | Entrada `W_caja` (peso con fluido atrapado y material anular) — **hand-off**, sin fórmula publicada |
| 50 | 204-3.8(c) | Aviso de expansión diferencial |
| **52,53** | 204-3.9(a) | Verificación `empuje_axial` + entrada `waiver_empuje` (la renuncia justificada que el código admite) — **hand-off** |
| 54,55 | 204-3.9(b)(c) | Avisos: viento, sismo, transitorios de fluido, otras condiciones |
| 57,58,59 | 204-3.10 | Paso 5 + especificación: vents, drenajes, tapón/brida/válvula |
| **61** | 204-3.11 | Entradas `L_sano` y `t_actual`; verificaciones de metal sano y de burn-through (Decisión 3) |
| **63** | 204-3.12 | Verificación `colapso_sellante` — **hand-off** |
| 64 | 204-3.12 | Aviso de *off-gassing* del sellante al curar |
| 66 | 204-3.13 | Aviso de *seepage* al componente dañado |
| 68 | 204-3.14 | Compuerta/aviso `juntas_especiales` (expansión, deslizantes, rótula) |

### 204-4 Fabricación · 204-5 Examen · 204-6 Prueba

| Bloque | Cláusula | Destino |
|---|---|---|
| 71 | 204-4.1 | Paso 6 + especificación: superficie libre de depósitos, pintura, mastics |
| 73,74,75 | 204-4.2 | Paso 6: izado, bisel, montaje lateral y deslizamiento |
| 77 | 204-4.3 | Paso 7 + especificación: WPS y soldador calificados |
| **79** | 204-4.4 | Paso 7 + hand-off Art. 210 (preheat, velocidad de enfriamiento, burn-through) |
| 81 | 204-4.5 | Paso 7: detener la fuga antes de soldar |
| 83 | 204-4.6 | Paso 7 + especificación: PWHT |
| 86 | 204-5.1 | Especificación + verificación: NDE consistente con E |
| **88,89,90** | 204-5.2 | **Único cálculo propio del artículo**: `sigma_vs_medio_admisible` (σ < ½·S) + riesgo de corrosión en resquicio |
| 92 | 204-5.3 | Especificación + verificación: PT/MT cuando la geometría impide RT/UT |
| 94 | 204-5.4 | Especificación: evaluación según el código |
| 97 | 204-6.1 | Entrada `P_prueba` + Paso 8: tipo de prueba por riesgo |
| **99–103** | 204-6.2 | **Fila LISTA `metodo_prueba`** con los cuatro métodos (Decisión 2) |
| **105** | 204-6.3 | Verificación `colapso_prueba` — **hand-off** |

---

## 4. Fase A — Chasis multi-fuente  ✅ EJECUTADA

Todo lo de esta fase está hecho y con pruebas. Lo que de verdad se escribió difiere en
dos puntos de lo planeado, y los dos son mejoras:

**A.1 — `Cita.fuente`, NO `Cita.archivo` portante.** El plan decía sobrecargar
`archivo` como ruta. **Era un error:** `archivo` es una ETIQUETA legible ("art_204.json",
"B31.3 cap. II") y los motores y las pruebas existentes la usan como nombre corto, no
como ruta —`test_citar_el_parrafo_que_si_publica_el_dato_no_aborta` lo destapó de
inmediato—. Se añadió un campo **nuevo** y explícito, `Cita.fuente`, que vale "" para el
caso normal (la fuente del motor). Compatible hacia atrás y más claro.
`Verificacion.cita` y `Especificacion.cita` entran ya al mismo recorrido: hasta ahora
**nadie las comprobaba** y una cita rota en una especificación pasaba en verde.

**A.2 — `bloques_citables(doc)`, y normaliza TRES formas, no dos.**

- PCC-2 / marker: `blocks` tal cual.
- Capítulo del B31.3: `preamble` + árbol `sections` en **preorden**; cada sección emite
  `section_header` con `id + heading`, luego sus `paragraphs` (`type = kind`), luego sus
  subsecciones.
- **Tabla suelta del B31.3** (no prevista): el rótulo como `section_header` —que el
  guardia rechaza, porque un rótulo no publica un valor— y una entrada `table_row` por
  fila impresa. Hace falta para citar la Tabla 304.4.1-1 y la 304.1.1-1.

El guardia de estabilidad está puesto y fija el archivo real:
`chapter_02.json` → **1106 bloques**, `295` = rótulo de `304.1.2`, `297` = eq. (3a),
`401/407/409` = eqs. (6)/(7)/(8) del refuerzo por área.

**A.3 — `Lista.cita`** ✅ y `comprobar_listas()` admite `opciones` fuera de
`CLAVES_CON_ENUMERACION` solo con cita. Añadido además el reverso: una enumeración del
**marco** (`unidad`, `modo`) **con** cita también aborta —afirmaría una procedencia que
no existe—.

**A.4 — `UMBRALES_PCC2` acepta ruta** ✅. Registradas `210_pared_delgada` y
`210_pared_reducir_aporte`, leídas en sus dos unidades. **Renombradas** respecto del
plan: no se llaman `204_burn_through` ni `204_electrodo` porque no son umbrales del 204
ni criterios de aceptación (ver la corrección arriba).

**A.5 — `verificar.py` §12** no cambia ✅: llama a la misma `comprobar_procedencia()`.

**A.6 — Pruebas** ✅. 19 pruebas nuevas en `test_motor_declarado.py` (48 → 67), y el
total del repositorio pasa de 429 a **436, todas en verde**.

---

## 4 bis. Fase A′ — lo que la ejecución obligó a añadir

**A′.1 — `{S}` de la cascada, dentro de una ecuación.** ✅ HECHA. Ver «Estado de
ejecución» arriba. Sin esto la eq. (3a) no se puede escribir.

**A′.2 — Cascada dimensional en el chasis declarado.** ⚠️ A MEDIAS. Ver «El punto
exacto de reanudación» arriba. Es el primer paso al continuar.

---

## 5. Fase B — Declarar el motor

`motores/art_204.py`, registrado en `motores/__init__.py`.

```
hoja:      Caja_PCC2_Art204
pestañas:  Espec_PCC2_Art204 · Instruc_PCC2_Art204
nodo:      NAV_CAL_ART204   (PCC-2 → Art. 204, con sus tres documentos)
casos:     ("Diseno",)      — UNA columna de valor (el chasis aborta con dos)
```

**Alcance de la primera entrega: la rama TE + CAPS.** Orden del ingeniero. La rama
«codo + tapas planas» se añade después, y su único añadido real es el hand-off a UG-34.

Secciones, en el orden que impone el marco:

1. **Datos de Entrada** — componente portador y defecto
2. **Resolución de Material** — cascada de cinco niveles contra `DB_B31_3` /
   `DB_BPVC_IID`, **dos columnas**: metal base (D) y material de la caja (E)
3. **Elegibilidad** — las cuatro vías de 204-2.2 y las compuertas duras
4. **Envolvente: selección del accesorio** — **cascada dimensional** norma → NPS →
   cédula contra `DB_B36_10`/`DB_B36_19`, de donde salen OD y espesor de la te. Más el
   *rating* del componente listado (`303`) como entrada declarada.
5. **Parámetros de Cálculo** — P, T, E, CA, `{S}`, Y, W
6. **Cálculo por `para. 304` — respaldo y verificación** — eq. (3a) para el cuerpo de
   la te, eqs. (6)/(6a)/(7)/(8) para la abertura; empuje y colapso como hand-off
7. **Verificaciones y avisos** — espesor real vs. `t_req`, área disponible vs.
   requerida, los cinco modos de fallo de 204-3.3, el aviso de pared delgada del 210
8. **Dictamen** — compuertas → material → AND (lo compone `formula_dictamen()`)
9. **Pasos del flujo (1–8)** — anexo

**B.3 — El cálculo de respaldo de `para. 304`, y lo que cuesta de verdad.**

Bloques del B31.3 ya localizados en el aplanado de `chapter_02.json`:

| Parte de la envolvente | Párrafo | Ecuación | Bloque |
|---|---|---|---|
| Cuerpo de la te (corrida) | `304.1.2(a)` | (3a) `t = PD/(2(SEW + PY))` | **297** |
| Abertura: área requerida | `304.3.3(b)` | (6) `A₁ = t_h·d₁·(2 − sin β)` | **401** |
| Abertura: área disponible | `304.3.3(c)` | (7) `A₂`, (8) `A₃` | **407**, **409** |
| Abertura en el cap | `304.4.2(e)` → `304.3.3` | las mismas, cierre como *header* | 497 |
| Compuerta de tamaño | `304.4.2(a)` | abertura ≤ ½ D_int; mayor → 304.6 / 304.5 | 493 |
| Espesor del cap | `304.4.1(b)` | (13) `t_m = t + c` → Tabla 304.4.1-1 | 478, **479** |
| Vía de calificación | `304.7.2` | sin fórmula: 5 vías, (a) = experiencia de servicio | 554+ |

La eq. (3a) arrastra dos dependencias:

- **Y** de la Tabla 304.1.1-1 — ya extraída limpia en
  `resources/ASME B31/ASME B31.3/CHAPTERS/tables/table_304_1_1_1.json`
  (5 filas × 8 columnas de temperatura, `null` donde el código imprime elipsis).
  Exige una base nueva `DB_B31_Y` con banda compacta e interpolación
  (*«The value of Y may be interpolated for intermediate temperatures»*), igual que
  las bandas de esfuerzos. **No es trabajo trivial: es una base más del libro.**
  **Alternativa acotada, si se quiere entregar antes:** `Y` como entrada con cita a la
  tabla, y la base en una fase posterior — declarado como hueco, no en silencio.
- **W** (factor de reducción de junta soldada, `para. 302.3.5(e)`): entrada con cita.
- **`S`** ya no es un problema: lo resuelve `{S}` (etapa 1b), leído de la cascada de
  material de la **columna E** (el material de la caja), no del metal base.

La rama **BPVC VIII-1** del selector de código queda **bloqueada y declarada**: la
Sección VIII no está en `resources/`. El motor lo dice en la celda, no lo calcula.

**B.4 — Pasos del flujo (8).** 1 Peligros y elegibilidad (204-2.2, 204-2.4) ·
2 Caracterización del componente y del defecto (204-3.11) · 3 Material de la caja
(204-3.1, 204-3.4) · 4 Diseño: cargas y espesor (204-3.5 a 204-3.9) · 5 Vents,
drenajes y sellante (204-3.10, 204-3.12, 204-3.13) · 6 Preparación e instalación
(204-4.1, 204-4.2) · 7 Soldadura, en servicio y PWHT (204-4.3 a 204-4.6) ·
8 Examen y prueba (204-5, 204-6).

---

## 6. Fase C — Navegación

`_nodo_de_articulo()` deriva el nodo solo. Lo que **no** se deriva, a propósito, son
las **tres líneas de `HojasNavegables()`** en `scripts/vba/mod_nav.vba` — ese es el
guardia que hace fallar `TestSincroniaPythonVba`. Tras tocarlo: `make_vba_seed.py` y
reconstruir. Las hojas navegables pasan de 43 a **47** (3 del artículo + `NAV_CAL_ART204`).

---

## 7. Fase D — Verificar y entregar

```
python build_db_materiales.py --resources ..\..\..\resources ^
    --in ..\..\..\templates\maestro_con_macros.xlsm ^
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_motor_declarado.py -q
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```

La §12 recalcula el motor en Excel real, exige cero errores de Excel, repite la
procedencia, comprueba que todo veredicto pinta y que el dictamen no se contradice.

---

## 7 bis. Dónde entra la BPVC aunque se trabaje en B31.3

Premisa del ingeniero (2026-09-15): la envolvente se fabrica con **material de
tubería existente**, así que el código de construcción natural es el B31.3. Eso
**no** saca a la BPVC del recorrido. Hay dos sentidos distintos y no se mezclan:

**(1) La BPVC como código de construcción del componente REPARADO.** El 204 delega
al código aplicable **al equipo encerrado**, no al material de la caja. Tubería de
proceso → B31.3; recipiente, intercambiador o **boquilla de recipiente** → Sec. VIII.
204-1(c) (b4) admite encerrar *«flanges and valves or fittings, branches, nozzles, or
vents and drains»*, y una boquilla puede ser de recipiente. Lo gobierna el selector
de código (`modo`); la rama VIII queda bloqueada y declarada.

**(2) La BPVC a la que el propio B31.3 REMITE.** Barrido completo del Capítulo II:

| Párrafo B31.3 | Remite a | Relevancia para la caja | ¿En `resources/`? |
|---|---|---|---|
| 302.2.4(b) | II-D Tabla Y-1 (S_y) | límite de esfuerzo circunferencial | **sí** |
| 302.3.1 / 302.3.6 | Sec. II, bases de admisibles | criterio de admisibles | **sí** |
| 304.1.3 | VIII-1 **UG-28..UG-30** | colapso por presión externa (sellante, prueba) | no |
| 304.4.1 | Tabla 304.4.1-1 → **UG-32/33/34** | espesor del cierre | no |
| 304.4.2 | **UG-36, UG-37(b), UG-38, UG-39** | tamaño límite y área requerida de la abertura | no |
| 304.5.1 / 304.5.2 | VIII-1 Ap. 2, Ap. Y, UG-34 | bridas y bridas ciegas | no |
| 304.7.2 | VIII-2 Anexo 5-F, VIII-1 UG-101, VIII-2 Parte 5 | vía de calificación (experimental, proof test, FEA) | no |
| 306.1.3 | B16.9 / SP-97 / UG-101 | accesorios de derivación *proprietary* | no |
| 322.6.3 | UG-125..UG-136 | dispositivos de alivio | no |

**La línea divisoria es limpia: de la BPVC, la Sección II está extraída y la Sección
VIII no.** Todo lo que el B31.3 manda a Sec. II lo resuelve el motor solo; todo lo que
manda a Sec. VIII es hand-off con la cadena de cita impresa.

**Consecuencia sobre las dos ramas de la Decisión 5:**

- **Te + caps** es la que mejor se cierra dentro de `resources/`: el cap sin cortar va
  por rating (`303`) y el corte por `304.4.2` + `304.3.3`, que está entero. Solo se
  escapa el área requerida de UG-37(b), y ahí el propio B31.3 publica su eq. (6).
- **Codo + tapas planas** hace hand-off a **UG-34 siempre**. El B31.3 no publica forma
  cerrada para un cierre plano soldado: lo manda a UG-34 por dos caminos
  independientes (`304.4.1` vía Tabla 304.4.1-1, fila *«Flat (pressure on either
  side)»*, y `304.5.2`). Lo único plano con fórmula propia es `304.5.3 Blanks`
  eq. (15) `t_m = d_g·√(3P/(16·SEW)) + c`, **y no se emite**: un *blank* es una placa
  entre bridas con junta (`d_g` = diámetro interior de la junta), no una tapa soldada;
  aplicarla sería una analogía que el código no hace.

## 8. Huecos declarados (van a `ISSUES` y al comentario de su celda)

1. **El Art. 204 no publica ecuaciones.** Todo el dimensionamiento se delega —lo dice
   el propio 204-3.5/3.6—. Límite de la fuente, no pendiente de ingeniería.
2. **BPVC Sección VIII no está en `resources/`** (sí la Sección II, entera). Quedan
   fuera de cita, cada uno con la cadena completa impresa en su columna de referencia:
   **UG-28..UG-30** (colapso externo, vía 304.1.3), **UG-32/UG-33/UG-34** (espesor del
   cierre, vía 304.4.1 + Tabla 304.4.1-1), **UG-36/UG-37(b)/UG-38/UG-39** (tamaño y
   área requerida de la abertura, vía 304.4.2), **UG-101** y **VIII-2 Parte 5 /
   Anexo 5-F** (vías de calificación de 304.7.2). La rama VIII del selector de código
   queda bloqueada.
2 bis. **ASME B16.9 no está en `resources/` y NO SE PUEDE EXTRAER hoy**: el PDF es un
   escaneo sin capa de texto y no hay OCR instalado. Dos consecuencias, las dos
   acotadas: (i) el *rating* presión-temperatura de la te y los caps entra como entrada
   declarada —pero `303` autoriza calcular con `304`, así que **no está en la ruta
   crítica**—; (ii) la dimensión **centro-a-extremo** del accesorio también, y esa sí no
   tiene sustituto —gobierna el largo de la caja, que de todos modos se verifica contra
   la distancia a metal sano medida en campo (204-3.11)—. El OD y el espesor de pared
   **sí** salen de `DB_B36_10`/`DB_B36_19`, por NPS y cédula.
2 quater. **Que un accesorio B16.9 se fabrique al espesor de la tubería coincidente lo
   dice B16.9, no `resources/`.** Va en el comentario de la celda como nota, nunca en la
   columna de referencia como cita normativa.
2 ter. **No se emite `304.5.3` eq. (15) para una tapa soldada.** Es la fórmula de un
   *blank* entre bridas con junta; usarla como cierre plano soldado sería una analogía
   que el código no hace.
3. **Resistencia del filete perimetral frente al empuje axial**: sin ecuación
   publicada en `resources/`. Entrada del ingeniero, no un 0,55 tecleado.
4. **Colapso por presión externa** (204-3.12, 204-6.3): íd.
5. **Longitud mínima de la caja**: 204-3.11 exige metal sano *sin cuantificar*. Se
   teclea la distancia medida en campo; no se emite `2√(Rt)`.
6. **Figura 204-1-1**: fotografía, sin dato acotado.
7. **Decisión 2**: el guardia comprueba que el bloque citado de una enumeración
   existe, no que cada opción aparezca literalmente en él (el libro está en español y
   el código en inglés).
8. **El `4,8 mm` del Art. 210 no es un criterio de aceptación.** Es un ejemplo de pared
   delgada que pide precauciones. Entra como aviso; no bloquea. Ver la corrección al
   principio de este plan.
9. **Tabla 210-4.2.1-1 sin reconstruir** (`structure: not_reconstructed`). El texto
   está entero en un solo bloque; la estructura de columnas, no. Etapa 0c, pendiente.

---

## 9. Riesgos

- **El primer motor declarado del libro.** `MOTORES_DECLARADOS` está vacío: el
  chasis, `verificar.py` §12 y `TestMotoresDeclarados` nunca se han ejercido contra
  un artículo real. Esperar defectos latentes del chasis, no del artículo.
- **El aplanado del B31.3 es una nueva superficie de error silencioso** (A.2).
  ✅ Mitigado: `test_el_aplanado_del_b31_3_es_estable` fija 1106 bloques y cuatro
  índices concretos contra el archivo real.
- **`DB_B31_Y` es alcance nuevo** (B.3), con banda compacta, interpolación y auditoría
  fila a fila en `verificar.py`. Si se quiere acotar el plan, `Y` puede empezar como
  entrada con cita a la tabla y la base dejarse para una fase posterior — se declara
  como hueco en vez de hacerlo en silencio.
- **Decisión 3 dejó de ser la de más criterio**, porque resultó que la cifra del 210 no
  es un criterio de aceptación: entra como aviso y la discusión se disuelve.
- **La cascada dimensional está a medias** (etapa 1c). Hoy no afecta al libro —ningún
  motor declara una `Dimensional`— pero declarar una pasaría los guardias sin producir
  desplegable. Cerrarlo es el primer paso al reanudar.
- **El caso semilla del 204 va a nacer en REVISAR, y es correcto.** El artículo tiene
  compuertas duras (grietas, *knuckle*, juntas de expansión) y cadenas en hand-off: un
  APTO de salida significaría que algo no está mordiendo. Mismo patrón que el 212 tras
  la Fase 2.
