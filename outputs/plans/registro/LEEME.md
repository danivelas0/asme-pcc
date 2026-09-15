# Registro de planes ejecutados

Lo que hay aquí **no describe el libro vigente**. Son los planes de trabajo tal
como se redactaron antes de ejecutarlos: conservan el diagnóstico, las
alternativas descartadas y el motivo de cada decisión, que es lo único que no
se puede reconstruir leyendo el código.

Cada uno cita el estado del libro **en su momento** —conteos de hojas, nombres
de hoja retirados después, rutas que luego cambiaron—. Para el estado vigente,
la fuente es `CLAUDE.md` en la raíz y, para el detalle por revisión,
`outputs/Base_Datos_Materiales_ASME/LEEME_Nota_de_Version.md`.

`outputs/plans/` (el nivel de arriba) queda para los planes **en curso**.

| Plan | Fecha | Qué resolvió | Estado del libro que cita |
|---|---|---|---|
| `Plan_Actualizacion_Base_Materiales_Motor_ASME_PCC.md` | 2026-09-06 | PLAN-DB-MAT-001: las bases de materiales del motor, desde el libro sin macros | `Motor_de_Calculo_ASME_PCC.xlsx` (ya no está en el repo) |
| `plan_dashboard_navegacion_xlsm_rev3.md` | 2026-09-06 | El paso a `.xlsm`: Dashboard único, capa VBA, hojas ocultas | Rev. 2 → Rev. 3, 37 → 38 hojas |
| `plan_motor_busqueda_apendice_c_b31_3.md` | 2026-09-07 | `Buscar_Prop_B31_3` y la separación del Apéndice C de los otros dos buscadores | Rev. 3 → Rev. 4, 38 hojas |
| `plan_motores_busqueda_A2_A3_factores_calidad.md` | 2026-09-07 | `Buscar_Ec_A2` y `Buscar_Ej_A3`, los factores de calidad | 39 → 43 hojas |
| `Plan_Bases_Datos_Seccion_II_ABC.md` | 2026-09-07 | Las nueve hojas de las Partes A, B y C desde los bloques `Line` | 39 → 48 hojas |
| `Plan_Dashboard_Jerarquico.md` | 2026-09-08 | El árbol de navegación de cinco niveles y las quince hojas `NAV_*` | 69 hojas |
| `plan_motor_art206_collar_y_separacion_parche_collar.md` | 2026-09-10 | `Collar_PCC2_Art206` (Type A/B), la Sección 7 de material compartida con el 212, y la decisión de desanclado total del 212 (delegada al plan siguiente) | 71 hojas |
| `plan_desanclado_total_motor_art212.md` | 2026-09-10 | `Parche_PCC2_Art212` nace 100 % en código; `HOJAS_HEREDADAS` queda en `("Instrucciones", "Datos_Ref")` | 71 hojas |
| `plan_entradas_de_motor_desde_base_de_datos.md` | 2026-09-11 | Toda entrada tabulada sale de una base por desplegable bloqueante (reglas 12/14); `DB_B36_10`/`DB_B36_19` para dimensiones; se retira `Datos_Ref` y `HOJAS_HEREDADAS` queda en `("Instrucciones",)`. Su `spec_*` acompaña. | 72 hojas |
| `plan_rediseno_motor_art212_flujo_completo.md` | 2026-09-11 | Los ocho pasos del flujo aprobado dentro del motor del Art. 212: compuerta de elegibilidad, cargas externas combinadas, proximidad a discontinuidades, los dos topes de filete de la NOTA de 212-3.4, excentricidad con separación, conformado en frío y energía neumática del App. 501. Cerrado: su único pendiente era la revisión en Excel, y este libro no lleva sign-off. | 71 hojas |
| `plan_rediseno_motor_art206_flujo_completo.md` | 2026-09-11 | Lo mismo para el collar de encierro total: selección guiada de tipo, C.A. en el `t_req` Type B, cateto del filete de las Figs. 206-3.5, luz radial de 206-4.1 y los avisos de 206-2/3/4/5/6. Cerrado por la misma razón. | 71 hojas |
| `plan_correcciones_ux_motores_art212_art206.md` | 2026-09-12 | Las 14 correcciones de UX/ingeniería de los dos motores tras el F9 del ingeniero, en once fases: leyenda de color (token `AMARILLO`, acotado por hoja), **una sola presión de diseño** (la máxima admisible, 212-3.2 / 206-3.3), la Sección de Material a la posición 2 con remapeo de filas y traslado del *oracle*, rastro de material plegable, bandas sin paréntesis y comentarios por sección, semáforo de aceptación y dictamen en bloque propio, **conmutador SI↔US de todo el motor**, botón de reinicio con manifiesto publicado en la hoja, y dos pestañas nuevas por artículo (especificaciones técnicas citando el párrafo de PCC-2, e instrucciones derivadas del propio motor). | 72 → **78 hojas**, 36 → **43 navegables** |
| `plan_flujo_github_ramas.md` | 2026-09-11 | GitHub Flow: `CONTRIBUTING.md`, plantilla de PR, CI parcial sin Excel (Actions, verde al primer run), tag `rev4e`, y el primer PR real (#1, merge `--no-ff`). **Fase 3 (ruleset) bloqueada:** los rulesets exigen GitHub Pro en repo privado. | No toca el libro; marca `rev4e` sobre el estado vigente |
| `plan_motor_art204.md` | 2026-09-15 | Primer motor DECLARADO del libro (chasis `motor_declarado.py`, skill `motor_pcc2`): `Caja_PCC2_Art204`, rama Te+caps, multi-fuente al B31.3 (delega el dimensionamiento, 204-3.5/204-3.6), cascada dimensional B36 en el chasis (etapa 1c), `{UMBRAL:...}` como extensión nueva del chasis, y el fix de `IFERROR` en las verificaciones que dependen de `{S@E}` sin material sembrado en esa columna. `verificar.py` §12 nueva: 0 fallos. | 82 hojas, 47 navegables |

Referencias que ya no valen, y que se dejan tal como se escribieron:

- Los conteos de hojas de todos ellos. El libro vigente tiene **78**.
- `Buscar_NoMetalicos` y `DB_NoMetalicos`, retiradas en la Rev. 4d por alcance.
- `Buscar_Propiedades`, renombrada `Buscar_Prop_IID`.
- `CLAVE_VOLVER`, `AbrirHoja` y `VolverAlDashboard`, sustituidos por
  `IrAHoja(destino, origen)` y `PADRE`.
- Los estados `PROPUESTA (VALIDAR)` y `REVISAR (regla textual)` de `MAP_Grupo`,
  que ya no existen.
- `prompt_notebooklm.md`, `propuesta_map_grupo.json` y `proponer_decisiones.py`,
  retirados al cerrarse `MAP_Grupo` con 0 casos abiertos: no quedaba nada que
  proponer ni que validar. Se citan en `Plan_Bases_Datos_Seccion_II_ABC.md`.
