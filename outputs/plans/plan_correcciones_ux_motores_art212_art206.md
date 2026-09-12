# Rediseño UX/visual de los motores Art. 212 y Art. 206 — Plan de implementación

> **Para ejecutores:** usar `superpowers:subagent-driven-development` o
> `superpowers:executing-plans`. Casillas `- [ ]` para seguimiento tarea por tarea.
> Rama de trabajo: `worktree-motores-art-212-206` (ahí vive el código más reciente;
> `main` no tiene aún los pasos 1-8). Todo esto se hace **sobre esa rama**, no sobre `main`.
>
> **EN EJECUCIÓN desde 2026-09-12** por orden explícita del ingeniero, fase a fase
> con checkpoint (pytest + `verificar.py` + commit aislado antes de seguir).

## Refinamientos acordados al arrancar la ejecución (2026-09-12)

Tras explorar el código se cambiaron dos cosas del plan escrito:

1. **La Fase 2 (presión) se ejecuta ANTES que la Fase 3 (reubicación).** Colapsar
   3→2 columnas elimina filas; es más limpio hacerlo antes del remapeo que después.
2. **La Fase 3 no se hace editando direcciones a mano.** El Art. 212 escribe ~1 400
   líneas de direcciones literales (`calc("D61", …)`, `ws["D39"]`), `comentar_art212_base()`
   indexa sus cuatro diccionarios **por número de fila**, `verificar.py` §6e lee
   `D52`/`D144`/`D42`/`D183..D191` literales y `test_dashboard.py` ancla sección a
   sección. En su lugar:
   - `MAPA_FILAS_212` / `MAPA_FILAS_206`: dict `fila_vieja → fila_nueva`, **única**
     fuente de verdad del desplazamiento.
   - `remapear_filas(ws, mapa)`: pase final que mueve valores, comentarios,
     `merged_cells`, `data_validations`, `row_dimensions`, formato condicional y
     celdas de clave, y reescribe toda referencia `$?[A-G]$?\d+` de las fórmulas.
   - `remapear_oracle(mapa)`: aplica **el mismo mapa** a `parche_art212_ref.json`.
     Decisión del ingeniero: el oracle **no se vuelca desde la hoja reconstruida**
     (eso lo convertiría en un espejo del builder y perdería su valor probatorio);
     se deriva del oracle viejo por traslación de filas, y el diff tiene que mostrar
     **solo cambios de dirección, ninguna fórmula alterada**.
   - `verificar.py` §6e importa el mapa del builder en vez de repetir direcciones.
   - **Punto de parada obligatorio:** el `MAPA_FILAS_212` y el diff del oracle se
     muestran al ingeniero antes de aplicarlos.

## Contexto

El ingeniero validó los motores Art. 212 y Art. 206 en Excel (F9) y encontró 14
correcciones de UX/ingeniería antes de dar el sign-off final. Se investigó a fondo
el código actual (`build_db_materiales.py`, rama `worktree-motores-art-212-206`,
~10 534 líneas) con 3 agentes de exploración en paralelo, y se resolvieron 4
ambigüedades de diseño con el ingeniero (ver «Decisiones tomadas» abajo). Este plan
cierra las 14 correcciones **en ambos motores**, reutilizando en lo posible
`construir_seccion7_material()` (función compartida) para no duplicar trabajo.

**Archivo único a modificar:** `outputs/Base_Datos_Materiales_ASME/scripts/build_db_materiales.py`
— funciones clave: `build_parche_art212()` (L6310-7704), `build_collar_art206()`
(L7931-8551), `construir_seccion7_material()` (L5855-6169, compartida), `comentar_art212_base()`
(L7707-7922), `rewrite_instrucciones()` (L8722-8752), `mod_nav.vba`/`this_workbook.vba`
(navegación y macros), `test_dashboard.py` (paridad/oracle/sistema visual).

## Decisiones tomadas con el ingeniero

1. **Paleta de colores:** se agrega **AMARILLO** como token nuevo del sistema único
   (documentado en CLAUDE.md, sumado a la lista blanca de `test_dashboard.py::TestSistemaVisual`).
   **Blanco** y **gris** ya existen en la paleta (`PAPEL`/`GRIS_2`) — solo cambia
   *dónde y cómo* se aplican, no se crean tokens nuevos para ellos. Alcance de la
   aplicación: los dos motores de cálculo (Art. 212 y Art. 206); el resto del libro
   no se toca en este plan.
2. **Reordenar la Sección de material justo después de la Sección 1:** SÍ, con
   reordenamiento físico de filas y **re-baseline completo del oracle** del Art. 212
   (`parche_art212_ref.json`). Es el ítem de mayor riesgo/esfuerzo del plan.
3. **Presión evaluada — verificado contra el código fuente (Regla nº1, sin inventar):**
   - **Art. 212** (`212-3.2`, `resources/.../art_212.json`): el código define una
     única `P = internal design pressure` para las ecs. (1)/(2) de carga. No exige
     un caso "envolvente" separado — es un margen de conservadurismo que agregó el
     motor, no el código.
   - **Art. 206** (`206-3.3`, `resources/.../art_206.json`): el texto es explícito —
     *"Type B pressure containing sleeves shall have a wall thickness equal to or
     greater than the wall thickness required for **the maximum allowable design
     pressure**..."* — es decir, el código **sí exige** dimensionar contra la presión
     de diseño **máxima admisible** (lo que el motor llama hoy "Envolvente / rating").
   - **Resolución:** se colapsan las 3 columnas actuales (Operación / Diseño típico /
     Envolvente) en **2**: *Presión de operación* y *Presión de diseño (máxima
     admisible / rating)* — esta última fusiona lo que hoy son dos campos separados
     ("Diseño típico" y "Envolvente") en uno solo, porque el código (206-3.3) exige
     literalmente la máxima admisible, no un valor "típico" intermedio menos
     conservador. Aplica igual en los dos motores para mantener el mismo modelo de
     presión. El campo "Diseño típico" **se elimina** (no se conserva como
     informativo): mantenerlo confundiría cuál presión gobierna el t_req.
4. **Filas de diagnóstico de la Sección de material:** se colapsan en un **grupo de
   filas plegable de Excel (outline)**, visible por defecto colapsado. No se borra
   nada (Regla nº1 de trazabilidad): el ingeniero puede expandir con el "+" para
   auditar `material_id resuelto`, `Base de datos activa`, `Índice de base`, `Fila
   localizada`, `n_pts`, `T1`, `T2`, `S en T1`, `S en T2`, `Temp. máx. admisible`,
   `Dictamen de rango`.

## Diseño final — numeración de secciones (ambos motores)

Los números de sección que el ingeniero usó en su pedido ("sección 3,4,5", "sección
7") se referían al esquema **actual**. Tras mover el material justo después de la
Sección 1, la numeración se corre. Mapa final:

**Art. 212** (`Parche_PCC2_Art212`):

| Nº nuevo | Sección | Nº actual | Notas |
|---|---|---|---|
| — | Título / Identificación / Aplicación y código | — | sin numerar, sin cambios de posición |
| — | **Leyenda de colores** (nueva) | — | banda nueva, antes de Sección 1 |
| 1 | Datos de Entrada | 1 | sin parentético; 2 columnas de presión (no 3) |
| 2 | **Resolución de Material** | 7 | **reubicada aquí** — filas físicamente movidas |
| 3 | Parámetros de Cálculo | (sin número) | ahora numerada |
| 4 | Geometría y Propiedades Derivadas | 2 | |
| 5 | Cálculo de Cargas y Soldadura | 3 | 2 columnas de presión; comentarios en col. parámetro + valor |
| 6 | Resultados del Diseño | 4 | comentarios en col. parámetro + valor |
| 7 | Verificaciones | 5 | comentarios en col. parámetro + valor + verificación; semáforo VERDE/ROJO; **Dictamen Global** como bloque propio, resaltado |
| — | Especificaciones Técnicas | 6 | **sale del sheet**, pasa a pestaña propia navegable |
| Anexo | Pasos 1-8 del flujo | (igual) | contenido sin cambios; solo se actualizan citas a "Sección 7" por su nuevo número 2 |

**Art. 206** (`Collar_PCC2_Art206`) — mapeo análogo, hoy sin numeración salvo 1 y 7:

| Nº nuevo | Sección | Hoy | Notas |
|---|---|---|---|
| — | Identificación / Aplicación y código | sin número | sin cambios |
| — | **Leyenda de colores** (nueva) | — | nueva |
| 1 | Datos de Entrada | "1. DATOS DE ENTRADA" | sin parentético |
| 2 | **Resolución de Material** | "7. RESOLUCION..." | **reubicada aquí** |
| 3 | Parámetros de Cálculo | "PARAMETROS DE CALCULO" | ahora numerada |
| 4 | Geometría del Sleeve | "GEOMETRIA DEL SLEEVE" | ahora numerada |
| 5 | Cálculo de Espesor Requerido | "CALCULO DE ESPESOR REQUERIDO" | sin parentético; 2 columnas de presión; comentarios param+valor |
| 6 | Verificaciones y Avisos | "VERIFICACIONES Y AVISOS" | comentarios param+valor+verificación; semáforo; **Dictamen Global** resaltado |
| — | Especificaciones Técnicas | *no existe hoy* | **nueva**, contenido generado desde cero, pestaña propia |
| Anexo | Pasos 1-8 del flujo | (igual) | sin cambios de contenido |

## Tareas

### Fase 0 — Preparación ✅
- [x] Baseline en limpio: `pytest` = **254**. El oracle queda congelado por git
      (`fba1417`), sin copia suelta.

### Fase 1 — Paleta de color y leyenda (ambos motores) ✅
- [x] Token `AMARILLO = FAEFC0` + `MOTOR_IN_FILL` / `MOTOR_CALC_FILL` / `MOTOR_LBL_FILL`,
      declarados juntos para que la leyenda y lo que se pinta no puedan divergir.
- [x] Documentado en CLAUDE.md (excepción declarada y **acotada por nombre de hoja**,
      mismo patrón que el semáforo).
- [x] `test_dashboard.py::TestSistemaVisual`: `AMARILLO` entra en `PALETA`, y **tres
      pruebas nuevas** fijan el alcance — solo las dos hojas, leyenda presente con el
      relleno que describe, y **toda** celda editable en amarillo (el reverso, que es
      lo que hace verdad la leyenda).
- [x] `inp()` de los dos motores y de `construir_seccion7_material()` → `MOTOR_IN_FILL`.
- [x] **En vez de etiquetar `calc()` celda a celda** (~200 sitios, imposible de
      mantener sin huecos): `aplicar_leyenda_motor()` **deriva** el color del estado
      real de cada celda —`locked=False` → amarillo, valor que empieza por `=` → gris,
      resto sin relleno propio → papel— como último paso de cada motor. Solo toca
      relleno y fuente, así que el *oracle* (que compara valores) no se ve afectado.
      Excluye por hipervínculo los botones, que también van desbloqueados.
- [x] Leyenda impresa en **D3/E3/F3 con la nota en G3**, sin fusionar ninguna celda:
      `A3:C3` ya lo ocupa el botón de volver y un merge nuevo por debajo de
      `FILA_ANEXO_FLUJO_212` rompería la paridad de fusionados contra el oracle.
- [x] Hallazgo del pase: había celdas **con texto y sin fuente propia** en los dos
      motores; se veían bien (heredaban la mono del sustrato de columna) pero se
      colaban por la auditoría de fuentes. El mismo pase les fija `DATA_F`.
- [x] **Checkpoint:** `pytest` = **257**, `verificar.py` = **0 fallos**.

### Fase 2 — Modelo de presión (2 casos, ambos motores) ✅
*Ejecutada antes que la Fase 3 (ver refinamientos): colapsar columnas elimina celdas y
es más limpio hacerlo antes del remapeo de filas.*

- [x] Art. 212 Sección 1: retirada la fila 27 ("Presión de diseño (típica)"); la fila 28
      pasa a **"Presión de diseño (máxima admisible / rating)"**, símbolo `P_dis`, con la
      cita a 212-3.2 y 206-3.3 en el comentario de celda.
- [x] Art. 212 "Cálculo de cargas y soldadura" (60-69) y Paso 2 del anexo (143-150):
      de 3 columnas a **2** (D/E = Operación/Diseño). La columna F se retira entera; E
      pasa a leer `$D$28`. Las filas 62-69 de la columna E no cambian de forma —
      encadenan desde E61—, así que siguen ancladas al oracle.
- [x] Verificaciones: `D84 = MAX(D64:E64)`, `D87 = E65`, `E89 = $D$28`, y la presión de
      prueba hidrostática (`B99`, `D192`) sobre `$D$28`.
- [x] Art. 206: mismo colapso (51-53); `T_s,mín gobernante` pasa a `$E$53`.
- [x] Comentarios (`_nota`) y `comentar_art212_base()` actualizados — su bucle `trip`
      recorre ahora `(4, 5)`, no `(4, 5, 6)`.
- [x] `verificar.py` §6e: la terna "Envolvente" sale; quedan dos casos.
- [x] ~25 divergencias contra el oracle declaradas una a una (retiradas vs. reemplazadas),
      fijadas por `TestModeloDePresionDosCasos`.
- [x] **Checkpoint:** `pytest` = **262**, `verificar.py` = **0 fallos**.

#### Hallazgo de ingeniería — el caso semilla pasó de APTO a REVISAR

No es un fallo de código. Con una única presión de diseño —la máxima admisible— el
parche de 8 mm del caso semilla **no cumple** el límite `1,5·Sa` de la ec. (5) del
212-3.4c a su rating: `S_w = 248,3 MPa` contra `1,5·Sa = 207 MPa` (verificación F85).
Hasta ahora la Sección 5 juzgaba ese esfuerzo contra los 10 kg/cm² del "diseño típico",
mientras el rating de 20 kg/cm² —que la hoja **ya traía** en D28— solo se miraba de lado
y no entraba al dictamen global.

Para volver a APTO hay que cambiar el **diseño** (espesor del parche, cateto, material)
o la presión de diseño de entrada, **no el motor**. Queda a decisión del ingeniero.

Como consecuencia, `verificar.py` §7 dejó de contrastar el dictamen contra el literal
`"APTO"` y pasa a contrastarlo contra **lo que implican los criterios que su propio AND
consulta**. Es un guardia más fuerte —comprueba que el dictamen no contradiga a sus
propias verificaciones— y no obliga a reescribir un literal cada vez que una decisión de
ingeniería mueve el resultado, que es justo cuando hay que mirar y no silenciar.

### Fase 3 — Reubicar la Sección de Material después de la Sección 1 ✅
*Aprobada por el ingeniero con el mapa a la vista (2026-09-12), renumeración incluida
en la misma fase.*

- [x] `MAPA_FILAS_212` / `MAPA_FILAS_206` + `remapear_filas()` + `remapear_referencias()`:
      la hoja se construye donde siempre y se mueve entera después, con un solo mapa.
      **El anexo de Pasos 1-8 no se mueve** — deja intactas las direcciones de
      `verificar.py` §6e y las anclas por-paso.
- [x] `remapear_oracle_212.py`: el oracle se **traslada** con el mismo mapa, no se
      vuelve a volcar. 333 de 469 celdas cambian de dirección, **0 fórmulas cambian de
      estructura**.
- [x] **Prueba del movimiento por recálculo en Excel real:** los 968 valores de A..G de
      los dos motores son idénticos bajo el mapa, 0 diferencias. Fusionados (49 y 13),
      validaciones (25 y 19) y comentarios verificados uno a uno.
- [x] Renumeración de bandas (1 Datos · **2 Material** · 3 Parámetros · 4 Geometría ·
      5 Cargas · 6 Resultados · 7 Verificaciones · 8 Especificaciones) y banda del
      material con `banda_literal()` en el 212 / `rotulo()` en el 206, vía los
      parámetros nuevos `titulo_banda` / `banda_rotulo`.
- [x] Citas de fila legibles: `mapa_citas` + `cita()` en `construir_seccion7_material()`
      (un `f"(fila {F + 15})"` apuntaba a la fila de **antes** de la mudanza) y repunte
      de las citas literales de los comentarios con patrones que no pueden confundirse
      con una designación de material.
- [x] Claves de `DIVERGENCIAS_*`, anclas de `test_dashboard.py` (281 literales) y
      direcciones de `verificar.py` §7/§6e (22) movidas con el mismo mapa.
- [x] 11 divergencias de texto por la renumeración, declaradas una a una;
      `_ancla()` hace que las listas de anclas respeten la tabla de divergencias en vez
      de perder cobertura, y `TestNumeracionDeSecciones` fija el **orden** de las bandas.
- [x] **Checkpoint:** `pytest` = **266**, `verificar.py` = **0 fallos**.

<details><summary>Plan original de la fase (para referencia)</summary>
Aplica a los dos motores; hacer Art. 212 primero (tiene oracle), luego Art. 206
(sin oracle, más simple) reusando el patrón validado.

- [ ] Mover físicamente el bloque de filas de `construir_seccion7_material()` (hoy
      escrito en la fila base 105 para Art.212 / 75 para Art.206) a una fila base
      inmediatamente posterior al fin de la Sección 1 (fin de Datos de Entrada).
- [ ] Renumerar/reindexar **todas** las referencias absolutas (`$D$nn`, `$E$nn`) de
      las secciones que quedan debajo (Parámetros de Cálculo, Geometría, Cálculo de
      Cargas, Resultados, Verificaciones) y del anexo de Pasos 1-8, sumando el
      desplazamiento de filas introducido por el bloque de material. Hacerlo de forma
      sistemática (un solo passthrough con mapeo fila-vieja→fila-nueva), no a mano
      celda por celda, para evitar desalinear una referencia.
- [ ] Actualizar el "cableado Sección 7 → Secciones 1/3" (hoy L6432-6484 en Art.212:
      escribe directo en D39/D40/D44) a las nuevas filas de destino.
- [ ] Actualizar `FILA_ANEXO_FLUJO_212` (y crear `FILA_ANEXO_FLUJO_206`, que hoy no
      existe) con el nuevo valor tras el desplazamiento.
- [ ] Re-generar el banner de la sección de material con `banda_literal()` (estilo
      consistente con las demás bandas — mayúscula/minúscula mixta, sin corchetes,
      **no** `rotulo()`) para que no desentone visualmente al quedar junto a la
      Sección 1 (hoy usa `rotulo()`, que la deja en `[ MAYÚSCULAS // CON BARRAS ]`,
      inconsistente con el resto — ver Fase 5, se resuelve junto con la limpieza de
      paréntesis).
- [ ] Re-baselinar `parche_art212_ref.json` desde la hoja reconstruida (una vez
      validada visualmente en Excel), como ya estaba anotado como pendiente en
      CLAUDE.md (Fase 10, "el re-baseline del oracle del 212... agrupado post-F9").
- [ ] Actualizar `test_dashboard.py` (`TestBuildParcheContraOracle` y afines) para
      leer contra el oracle re-baselinado; confirmar que sigue excluyendo solo el
      anexo real (rows ≥ `FILA_ANEXO_FLUJO_212` nuevo).
- [ ] Art. 206: mismo movimiento de filas, sin oracle que re-baselinar (usa anclas de
      cadena + `verificar.py §6f`); crear `FILA_ANEXO_FLUJO_206` como constante nueva
      (hoy no existe) y usarla igual que en 212 para excluir el anexo de cualquier
      test de paridad futuro.

</details>

*Nota de ejecucion: `FILA_ANEXO_FLUJO_206` no hizo falta. Al no mover el anexo en
ninguno de los dos motores, no hay frontera de paridad que acotar ahi.*


### Fase 4 — Simplificar la Sección de Material ✅

- [x] Las once filas de trazabilidad (`material_id resuelto` … `Temperatura máxima`)
      van en un **outline de Excel plegado por defecto**; el "+" del margen las abre.
      `remapear_filas()` se extendió para arrastrar `outline_level`/`hidden` junto con
      el alto: si no, el plegado se habría quedado en las filas de antes de la Fase 3.
- [x] **Desviación deliberada de la decisión 4 del plan:** `Dictamen de rango` **no**
      se pliega, aunque la decisión lo listaba. Es lo que **bloquea** el cálculo, y una
      condición de bloqueo escondida detrás de un "+" es una condición que nadie ve.
      Se pliega `F+10..F+19` y quedan siempre visibles `Dictamen de rango` y
      `S(T) resuelto`, que es lo que la propia tarea de la Fase 4 pedía dejar a la vista.
- [x] La columna "Referencia / Notas" decía `Lista desplegable en cascada` en los cinco
      niveles; ahora lo dice una vez, en el nivel 0.
- [x] Etiquetas sin jerga: `Indice de base` → `Base ASME aplicada`, `n_pts (puntos
      tabulados) / p1` → `Puntos tabulados de la fila`, `Temp. max. admisible / limite
      VIII-1` → `Temperatura maxima admisible`, `Fila localizada` → `… en la base`.
- [x] 10 divergencias de texto declaradas; `TestDiagnosticoPlegable` fija el grupo, que
      el dictamen y el resultado **no** se plieguen, y que **nada del rastro se pierda**
      (las once filas conservan rótulo y valor: plegar no es borrar — Regla nº 1).
- [x] **Checkpoint:** `pytest` = **273**, `verificar.py` = **0 fallos**.

<details><summary>Plan original de la fase</summary>

- [ ] En `construir_seccion7_material()`: agrupar (Excel `ws.row_dimensions[r].outline_level`
      + `outlinePr`/`sheet_view.showOutlineSymbols`) las filas de diagnóstico —
      `material_id resuelto`, `Base de datos activa`, `Índice de base`, `Fila
      localizada`, `n_pts`, `T1`, `T2`, `S en T1`, `S en T2`, `Temp. máx. admisible`,
      `Dictamen de rango` — bajo un grupo colapsado por defecto
      (`ws.sheet_properties.outlinePr.summaryBelow = False` si el resumen queda arriba,
      revisar orientación). Dejar visibles siempre: la cascada de selección (Familia →
      Composición → Forma → Especificación → Tipo/Grado → Variante), el modo de
      lectura S(T), la temperatura de evaluación, y los dos resultados finales:
      **S(T) resuelto** y el dictamen de rango resumido.
- [ ] Condensar la columna "Referencia / Notas" — hoy repite "Lista desplegable en
      cascada" 5 veces (una por nivel) — a una sola nota fija arriba de la cascada
      ("Los 5 campos de abajo son una cascada de listas desplegables dependientes")
      en vez de repetirla en cada fila.
- [ ] Simplificar visualmente las etiquetas de columna A (hoy con jerga tipo
      "0 · Familia de material", "n_pts (puntos tabulados) / p1") a redacción más
      llana donde no perjudique la trazabilidad (ej. mantener el número de nivel de
      cascada pero simplificar paréntesis técnicos).
- [ ] Verificar que agrupar filas no rompe ninguna referencia absoluta ni el
      `merge_cells` existente (el agrupamiento es solo metadato de fila, no mueve
      celdas — bajo riesgo, pero confirmar con `verificar.py` después).

</details>


### Fase 5 — Encabezados sin paréntesis + comentarios por sección ✅

- [x] **Paréntesis explicativos fuera** de las 22 bandas de los dos motores. La **cita
      al código se conserva** —es normativa— pero sale del paréntesis, que en el resto
      de la hoja significa «aclaración prescindible»: `(ASME PCC-2 Art. 212-3.2)` pasa a
      `— ASME PCC-2 Art. 212-3.2`. El **título** de la hoja (A1) conserva el suyo: ahí
      el paréntesis es la cita, no una aclaración.
- [x] Reglas de comentario **declaradas por sección** (`REGLAS_COMENTARIO_212/206`) y
      repartidas por un pase (`aplicar_reglas_de_comentario`), no editando ~200 llamadas:
      por omisión solo la columna de valor; `+A` en Cálculo de Cargas/Espesor y en
      Resultados (ahí el rótulo es un símbolo y es lo que se consulta); `+A +Resultado`
      en Verificaciones.
- [x] **No se inventó texto:** cuando una columna obligada no traía comentario propio se
      le pone el de la fila, tomado de la columna que la sección declara como fuente. En
      Verificaciones la fuente es **Resultado**, no Requerido: el texto que describe la
      fila entera es «CUMPLE si …», no «valor mínimo exigido».
- [x] `TestReglasDeComentario` fija las dos direcciones —que ninguna columna lleve de
      más y que las obligadas lo lleven— y que ninguna banda conserve paréntesis.
- [x] **Checkpoint:** `pytest` = **279**, `verificar.py` = **0 fallos**.

**Corrección de dos aserciones mías, que eran más estrictas que la realidad:** (i) dos
de las seis verificaciones del 212 no resuelven en CUMPLE/NO CUMPLE sino en una **ruta**
(«Refuerzo 360°» / «Parche local»; «OK — parche» / «Migrar (Art. 206)»), así que exigir
la palabra «CUMPLE» era falso; se exige en su lugar que el globo del Resultado **no sea
el mismo** que el del Requerido. (ii) El título de la hoja no es una banda de sección.

<details><summary>Plan original de la fase</summary>

- [ ] Quitar el texto entre paréntesis de **todas** las bandas de sección en ambos
      motores (lista completa en el reporte de exploración: "1. DATOS DE ENTRADA
      (campo con linea inferior = editable)", "PARÁMETROS DE CÁLCULO (constantes —
      editables)", "3. CÁLCULO DE CARGAS Y SOLDADURA (ASME PCC-2, Art. 212)", "5.
      VERIFICACIONES (criterios de aceptación)", "6. ESPECIFICACIONES TÉCNICAS
      (generadas automáticamente...)", "APLICACIÓN Y CÓDIGO DE CONSTRUCCIÓN
      (selector...)", "7. RESOLUCIÓN DE MATERIAL... (cascada de listas
      desplegables)", y los 8 títulos "PASO N · ... (ASME PCC-2 Art. ...)" — en estos
      últimos, conservar la cita del artículo fuera del paréntesis o como parte del
      título sin paréntesis, ya que es información normativa, no explicativa; solo se
      quita lo puramente aclaratorio/explicativo, no la cita al código). Aplicar
      igual en Art. 206.
- [ ] Reglas de comentarios (`_nota`) — auditar y corregir columna por columna:
  - **Regla general:** comentario solo en la(s) columna(s) de **valor** (D, y E/F
    cuando la sección tenga varios casos en paralelo).
  - **Excepción — Sección "Cálculo de Cargas y Soldadura"/"Cálculo de Espesor
    Requerido"**: agregar también comentario en columna A (parámetro).
  - **Excepción — Sección "Resultados del Diseño"**: agregar también comentario en
    columna A (parámetro).
  - **Excepción — Sección "Verificaciones"**: agregar comentario en columna A
    (parámetro) **y** en la columna de Resultado (verificación).
  - Todas las demás secciones (Datos de Entrada, Geometría, Resolución de Material,
    Parámetros de Cálculo): comentario **solo** en columna de valor; quitar los que
    hoy existan en columna A si no corresponde.
  - Revisar `comentar_art212_base()` (el pase idempotente final) — hoy aplica
    comentarios en A/D-E-F/verif/B según diccionarios `simples`/`trip`/`verif`/`textos`;
    ajustar esos diccionarios a la regla nueva en vez de aplicar comentario de columna A
    a todo.

</details>


### Fase 6 — Semáforo VERDE/ROJO + Dictamen Global resaltado ✅

- [x] Tokens `CUMPLE_OK_*` / `CUMPLE_BAD_*` y `DICTAMEN_*`. El **ROJO aparece como
      relleno por primera vez fuera del aviso de macros**, con el mismo significado que
      tiene en todo el libro: bloqueado. Va en formato condicional (dxf), no como estilo
      de celda, así que `test_el_aviso_de_macros_es_el_unico_relleno_rojo` sigue valiendo
      sin tocarlo.
- [x] Semáforo en las 8 celdas de Resultado del 212 (incluidos los dos topes de filete
      del anexo) y las 3 del 206. Los textos favorables se declaran **fila a fila**
      (`SEMAFORO_212/206`): cuatro verificaciones resuelven en CUMPLE/NO CUMPLE, pero dos
      resuelven en una **ruta** de reparación y ahí lo verde es la rama que deja seguir
      con el parche.
- [x] **El rojo se pinta por complemento**, no por lista de textos malos: así un `#N/A`
      o un texto que nadie previó sale bloqueado, que es el lado seguro. Enumerar lo malo
      dejaría lo imprevisto en blanco, indistinguible de «aún no calculado».
- [x] **Dictamen Global como bloque propio:** banda de tinta a todo el ancho, `MACRO` 16,
      fila más alta y **franja roja arriba** separándolo de la tabla. CF sobre la fila
      entera: APTO → verde, `ELIJA MATERIAL…` → **ámbar** (falta una entrada, no falla
      una verificación: pintarlo de rojo confundiría «aún no has elegido» con «no
      cumple»), todo lo demás → rojo.
- [x] Se ejecuta **después** del pase de leyenda, a propósito: la leyenda pinta de gris
      toda celda de fórmula y el dictamen lo es, pero no es un dato más de la tabla.
- [x] `TestSemaforoDeAceptacion` (8 pruebas), incluida la que ataja el **fallo silencioso
      más probable de la fase**: que el texto considerado favorable no sea letra por
      letra el que escribe la fórmula — una raya larga distinta bastaría para que el
      semáforo no encendiera nunca y la celda saliera roja siempre, que parece un
      resultado.
- [x] **Checkpoint:** `pytest` = **289**, `verificar.py` = **0 fallos**.

*Desviación: el dictamen no se movió a una fila nueva con separador en blanco (habría
exigido otro remapeo completo). Se le dio separación con la franja roja superior y el
alto de fila, que consigue el mismo aislamiento visual sin volver a mover la hoja.*

<details><summary>Plan original de la fase</summary>

- [ ] Agregar tokens nuevos `CUMPLE_OK_FILL`/`CUMPLE_OK_FONT` (reusar `VERDE`+`TINTA`,
      igual patrón que `SEL_OK_FILL/FONT`) y `CUMPLE_BAD_FILL`/`CUMPLE_BAD_FONT`
      (`ROJO` de fill + texto en `PAPEL` para contraste, ya que `ROJO` como fill sólido
      es nuevo uso pero semánticamente coherente con "bloqueado/no cumple").
- [ ] Sobre cada celda de Resultado de la Sección Verificaciones (F84-F89 hoy en
      212, F60/F61/F123 hoy en 206, y las de topes de filete F161/F162 del anexo):
      `ws.conditional_formatting.add(rango, FormulaRule(formula=['&lt;celda&gt;="CUMPLE"'],
      fill=CUMPLE_OK_FILL, font=CUMPLE_OK_FONT))` + regla espejo para "NO CUMPLE".
      Para las filas con textos de ruta distintos de CUMPLE/NO CUMPLE (F88/F89 en 212:
      "Refuerzo 360°"/"Parche local", "OK — parche"/"Migrar (Art.206)"), usar el mismo
      patrón adaptando el texto esperado a VERDE/ROJO según cuál sea la rama favorable.
- [ ] **Dictamen Global** (F90 en 212, F69 en 206): sacarlo de la tabla de
      verificación a su **propia fila**, con una fila en blanco antes como separador.
      Aumentar fuente a `MACRO` tamaño ≥ 16-18 (vs. 11 hoy), fill de banda completa
      (ancho A:G) con `BAND_FILL`/`TINTA` de fondo y texto en `PAPEL`, salvo que el
      resultado sea APTO/REVISAR/PROHIBIDO/ELIJA MATERIAL — en ese caso, aplicar
      `conditional_formatting` sobre la fila entera: APTO→verde, REVISAR/PROHIBIDO/NO
      ELEGIBLE/FUERA DE ALCANCE→rojo, ELIJA MATERIAL→ámbar (reusando `AMBAR`).

</details>


### Fase 7 — Conmutador SI/US (Sección de Material + resto del motor)
- [ ] Agregar celda selector "Sistema de unidades" con `dv_list(ws, celda, '"SI,US"')`
      en la banda "Aplicación y código de construcción" de cada motor (mismo patrón
      que los buscadores).
- [ ] Extender `construir_seccion7_material()` para aceptar también las bases US
      (`b313c`, `iid1ac`, `iidbc`) como parámetros opcionales, y envolver cada
      `CHOOSE`/`IF({M}=...)` existente (que hoy solo conmuta B31.3↔BPVC según el modo
      de componente) en un segundo nivel: `IF($unidad="SI", &lt;rama SI ya existente&gt;,
      &lt;misma rama contra la base US&gt;)`. Usar `CHOOSE` con índice combinado si el anidado
      de `IF` exige entrada matricial al alimentar un `MATCH` (mismo criterio que ya
      usa `_choose_b36`).
- [ ] `main()`: pasar `b313c, iidc, iidbc` (ya construidos hoy solo para los
      buscadores) también a `build_parche_art212()` y `build_collar_art206()`.
- [ ] Aplicar el mismo patrón ya usado en los buscadores para el resto del motor:
      el selector SI/US solo cambia **de qué hoja/base se lee** (materiales, y las
      etiquetas de unidad que dependan de tablas del código); los valores que el
      ingeniero ya tecleó (dimensiones, presiones, temperaturas) **no se convierten
      numéricamente** al cambiar de sistema — el ingeniero debe volver a teclearlos en
      el nuevo sistema, igual que ya ocurre en los 5 buscadores de cascada (Regla nº9:
      no se convierte, se relee de la edición correspondiente).
- [ ] Actualizar las etiquetas de unidad de los campos de entrada (kg/cm², mm, °C)
      para que muestren la unidad correspondiente al sistema elegido (ej.
      `=IF($unidad="SI","kg/cm²","psi")`), sin tocar el valor numérico subyacente.
- [ ] `verificar.py`: agregar/extender el bloque que recalcula en Excel real la
      Sección de Material en modo US para al menos un material de control (paralelo a
      §6e/§6f existentes).

### Fase 8 — Botón de reinicio (limpiar entradas)
- [ ] Al construir cada motor, recolectar en una lista Python la dirección de cada
      celda `inp()` (ya se llama una vez por celda editable — agregar `.append()` a la
      lista en el propio helper `inp()` de cada motor).
- [ ] Escribir esa lista como manifiesto en un rango de celdas ocultas de la propia
      hoja (mismo patrón que las listas de cascada materializadas en columnas
      ocultas) — así VBA no necesita una copia hardcodeada de direcciones (evita el
      riesgo de "dos fuentes de verdad que divergen" ya documentado para
      `HojasNavegables()`).
- [ ] En `mod_nav.vba`: nueva `Public Sub LimpiarEntradas(ByVal hoja As String)` que
      lee el manifiesto de esa hoja y pone en blanco cada celda listada (`.ClearContents`,
      no `.Clear`, para no tocar formato/validación).
- [ ] En `this_workbook.vba` (`Workbook_SheetFollowHyperlink`): distinguir claves de
      reinicio (prefijo `"RESET:"`) de claves de navegación, despachando a
      `LimpiarEntradas` en vez de `IrAHoja` cuando corresponda.
- [ ] En Python: agregar un botón (`_boton()`) rotulado "⟲ REINICIAR ENTRADAS" cerca
      del título de cada motor, con clave `"RESET:Parche_PCC2_Art212"` /
      `"RESET:Collar_PCC2_Art206"`.
- [ ] `test_dashboard.py`: agregar prueba de que el manifiesto de celdas a limpiar
      coincide exactamente con las celdas que de verdad tienen `Protection(locked=False)`
      en cada hoja (ninguna celda editable queda fuera del reinicio, y ninguna celda
      de botón/navegación queda incluida por error).

### Fase 9 — Pestaña "Especificaciones Técnicas" por motor
- [ ] Crear `build_especificaciones_art212(wb)` y `build_especificaciones_art206(wb)`:
      hojas nuevas navegables, mismo sistema visual del libro (no heredado).
- [ ] Art. 212: migrar el contenido actual de la Sección 6 (método, juntas de
      cierre, filetes, soldadura en servicio Art.210, END, recubrimiento, prueba de
      hermeticidad) y **expandirlo** citando explícitamente las cláusulas de
      `resources/.../art_212.json` y `App. 501` que sustentan cada especificación
      (hoy el texto es generado pero no siempre cita el párrafo exacto).
  - Quitar del sheet principal la Sección 6 (ya no vive ahí — Fase 3/5 no debe
    reservarle numeral).
- [ ] Art. 206: construir desde cero un contenido equivalente (no existe hoy) —
      método (Type A/B), preparación de bordes, soldadura en servicio Art.210,
      END, prueba de hermeticidad del anular — citando `resources/.../art_206.json`
      artículo por artículo.
- [ ] Agregar ambas hojas al árbol `ARBOL` como hijas del nodo de cada motor
      (`_hoja_final(...)`), y sincronizar `HojasNavegables()` en `mod_nav.vba` en la
      misma posición de preorden (`test_dashboard.py::TestSincroniaPythonVba`).
- [ ] Botón de acceso desde cada motor ("📋 ESPECIFICACIONES TÉCNICAS →") y botón de
      retorno (`ANCLA_VOLVER` por defecto `(3,1,3)` debería servir; confirmar layout).

### Fase 10 — Pestaña "Instrucciones" por motor (paso a paso)
- [ ] Crear `build_instrucciones_art212(wb)` y `build_instrucciones_art206(wb)`
      (hojas nuevas, navegables, mismo patrón de `rewrite_instrucciones` pero con
      contenido propio — **no** se mezcla con la `Instrucciones` global, que es sobre
      buscadores/bases).
- [ ] Contenido por cada motor, estructurado por sección (siguiendo la numeración
      final de la Fase 3):
  1. Qué hace el motor y para qué artículo del código sirve.
  2. Por cada sección: qué se hace ahí, en lenguaje llano.
  3. Por cada fila/celda **no bloqueada o de lista desplegable**: qué debe insertar
     el ingeniero, con ejemplo.
  4. Por cada celda de **fórmula**: qué calcula, y de qué cláusula del código sale
     (cita explícita al artículo/ecuación de `resources/`, no una paráfrasis de
     memoria — Regla nº1).
  5. Explicar la leyenda de colores (Fase 1) y el conmutador SI/US (Fase 7).
  6. Explicar el botón de reinicio (Fase 8) y el semáforo de verificaciones (Fase 6).
- [ ] Botón de acceso ("❔ INSTRUCCIONES DE ESTE MOTOR →") desde cada motor, además
      del botón global "MANUAL DE USO" ya existente (no se reemplaza, se suma).
- [ ] Sincronizar `ARBOL`/`HojasNavegables()` igual que en la Fase 9.

### Fase 11 — Verificación
- [ ] `pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q`
- [ ] `verificar.py --resources ... --wb ...` (0 fallos, incluidas §6e/§6f
      ampliadas de la Fase 7).
- [ ] Exportar ambas hojas a PDF/PNG desde Excel real y revisar visualmente: leyenda
      de colores, semáforo, Dictamen Global, agrupamiento plegable, ausencia de
      paréntesis en bandas, dos columnas de presión.
- [ ] Entregar el `.xlsm` reconstruido para el F9 final del ingeniero.
- [ ] Actualizar CLAUDE.md con el nuevo estado (numeración, paleta, conmutador
      SI/US en estos motores, botón de reinicio) y mover este plan a
      `outputs/plans/registro/` con su fila en `LEEME.md` una vez cerrado.

## Notas de riesgo

- La Fase 3 (reordenar Sección de Material) es, con diferencia, la de mayor riesgo:
  toca todas las referencias absolutas del motor y obliga a re-baselinar el oracle.
  Hacerla en un commit aislado, verificable independientemente del resto de fases,
  antes de continuar con las fases visuales/cosméticas.
- La Fase 7 (SI/US) es la segunda de mayor tamaño: extiende una función compartida
  por los dos motores con un segundo eje de conmutación. Probar primero sobre Art.212
  y confirmar en Excel real antes de replicar a Art.206.
- Fases 1, 4, 5, 6, 8, 9, 10 son de riesgo bajo/medio y no tocan referencias
  absolutas de cálculo (solo estilo, agrupamiento, contenido nuevo en pestañas
  aparte, y una macro nueva aislada).
