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

Referencias que ya no valen, y que se dejan tal como se escribieron:

- Los conteos de hojas de todos ellos. El libro vigente tiene **72**.
- `Buscar_NoMetalicos` y `DB_NoMetalicos`, retiradas en la Rev. 4d por alcance.
- `Buscar_Propiedades`, renombrada `Buscar_Prop_IID`.
- `CLAVE_VOLVER`, `AbrirHoja` y `VolverAlDashboard`, sustituidos por
  `IrAHoja(destino, origen)` y `PADRE`.
- Los estados `PROPUESTA (VALIDAR)` y `REVISAR (regla textual)` de `MAP_Grupo`,
  que ya no existen.
- `prompt_notebooklm.md`, `propuesta_map_grupo.json` y `proponer_decisiones.py`,
  retirados al cerrarse `MAP_Grupo` con 0 casos abiertos: no quedaba nada que
  proponer ni que validar. Se citan en `Plan_Bases_Datos_Seccion_II_ABC.md`.
