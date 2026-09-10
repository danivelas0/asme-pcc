# Spec — Toda entrada de motor que exista en una base se elige de ella

**Fecha:** 2026-09-10
**Estado:** diseño aprobado por el ingeniero; pendiente de plan de implementación.
**Origen:** revisión de `Parche_PCC2_Art212`, donde el campo «Material de tubería /
envolvente» resultó ser una lista de seis materiales tecleada en el builder.

---

## 1. Objetivo

Que **ninguna variable de entrada de un motor de cálculo que exista en una base de
datos del libro se teclee a mano ni salga de una lista fija**. Si el libro tabula el
valor, el motor lo ofrece desde esa tabla, en un desplegable que además **rechaza** lo
que no esté en ella. El fin es eliminar el error de tecleo y la divergencia
lista-vs-tabla como clases de fallo, no solo documentar de dónde sale el dato.

Gobiernan las tres reglas registradas en `CLAUDE.md` → «Reglas de diseño del libro»:

- **Regla 12** — todo material se selecciona de una base de datos, nunca de una lista fija.
- **Regla 13** — motores de búsqueda y de cálculo no se conectan entre sí; ambos leen solo de las bases.
- **Regla 14** — toda variable de entrada que exista en una base se elige de un desplegable; nunca se teclea. La línea: ¿el libro tabula este valor? → desplegable. ¿Es dato de servicio o de inspección? → entrada manual.

---

## 2. Hallazgos verificados

Todo lo de esta sección se comprobó leyendo el código y el `.xlsm` construido, no de
memoria. Las líneas citadas son del `build_db_materiales.py` en el commit `9c67baf`.

### 2.1 Inventario de listas fijas

Motores de cálculo del libro: **dos** — `build_parche_art212` (6165) y
`build_collar_art206` (7265).

| Motor | Celda | Selecciona | Dónde vive el dato de verdad | Veredicto |
|---|---|---|---|---|
| Art. 212 | D22/D23 | Material (6 ítems, `DataValidation` directa en 6478-6483) | `DB_B31_3` / `DB_BPVC_IID*` vía cascada de la Sección 7 | **viola regla 12** |
| Art. 212 | D18 | NPS (33 ítems) | `Datos_Ref!$A$5:$A$37` (33 filas) | **viola regla 14** |
| Art. 212 | D19 | Cédula (14 ítems, línea 6432) | `Datos_Ref!$C$4:$P$4` (14 columnas) | **viola regla 14** |
| Art. 206 | D18 | NPS (33 ítems) | `Datos_Ref!$A$5:$A$37` | **viola regla 14** |
| Art. 206 | D19 | Cédula (14 ítems, línea 7378) | `Datos_Ref!$C$4:$P$4` | **viola regla 14** |

Listas fijas **legítimas**, que este trabajo no toca: son modos y booleanos del propio
motor, no datos tabulados de un código.

| Lista | Dónde | Por qué se queda fija |
|---|---|---|
| `SI,US` | buscadores (3865, 4321, 4675, 5191) | conmutador de unidades (regla 10) |
| `Interpolado,Tabulado-conservador` | Sección 7 (5927) y buscadores | modo de lectura, regla del código |
| `Si,No` | Art. 206 D23/D24 (7408, 7415) | booleanos de examen UT y defecto circunferencial |
| `Type A,Type B` | Art. 206 D22 | los dos tipos que define 206-1.1.1/206-1.1.2 |

### 2.2 Los desplegables actuales no restringen nada

`dv_list()` (364-366) construye **toda** validación del libro con
`showErrorMessage=False`. Excel muestra la lista, pero acepta en silencio cualquier
valor tecleado fuera de ella. Hoy el desplegable es una sugerencia, no una
restricción — lo contrario del objetivo de esta spec.

### 2.3 `Datos_Ref` no es todavía una base auditada

Es una de las **dos** hojas que quedan en `HOJAS_HEREDADAS` (8697:
`("Instrucciones", "Datos_Ref")`), es decir, viene del maestro Rev0 y **no se
construye desde `resources/`**. Ni B36.10M ni B36.19M están extraídos ahí (`resources/`
solo tiene `ASME B31`, `ASME PCC` y `ASME_BPVC`). Forma real de la hoja en el libro
construido: `A1:P49`; fila 4 = `NPS (in)`, `OD (mm)` y las 14 cédulas
(`5,10,20,30,40,60,80,100,120,140,160,STD,XS,XXS`); `A5:A37` = 33 NPS de 0,5" a 48";
`C5:P37` = la matriz de espesores.

Además, sus **filas 40-47** contienen un bloque de esfuerzos admisibles de seis
materiales rotulado por la propia hoja como
`[OBSOLETO — ver DB_B31_3 / DB_BPVC_IID] … Se conserva como respaldo historico de las
memorias ya emitidas; el motor ya NO lee de aqui.` Es exactamente la «lista corta» que
la lista fija de D22/D23 espejaba: la lista fija duplicaba un bloque que la hoja
declara muerto.

### 2.4 El `""&` de la cédula es un síntoma, no una peculiaridad

La fórmula de espesor lleva `MATCH(""&$D$19, Datos_Ref!$C$4:$P$4, 0)` porque la lista
fija entrega la cédula como número mientras la fila de encabezados la guarda como
texto. Es divergencia lista-vs-tabla en estado puro; desaparece sola cuando el
desplegable lee de la propia fila de encabezados.

### 2.5 Dónde están las normas dimensionales

Accesibles desde esta máquina en el perfil `User` (el `CLAUDE.md` documenta la ruta
bajo el perfil `dvelasquez`, que **no** resuelve desde esta sesión):

```
C:\Users\User\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES\
  0-STANDARDS\ASME\B36-DIMENSION OF  WROUGHT STEEL PIPES\
    ASME B36.10M - Welded and Seamless Wroth Steel Pipe 2022.pdf
    ASME B36.10M - Welded and Seamless Wroth Steel Pipe 2022-2.xlsx
    ASME B36.19M - Stainless Steel Pipe 2022.pdf
```

---

## 3. Alcance

**Entra:** las cinco celdas del inventario 2.1; el bloqueo de los desplegables en los
**12 motores** del libro (2 de cálculo + 10 de búsqueda); el guardia que impide la
reincidencia; la conversión de `Datos_Ref` en base auditada; la entrada de B36.19M.

**«Hoja de motor», definido** — porque el guardia de la Fase 1 depende de esta lista y
no puede quedar a interpretación. Son las **doce**: `Parche_PCC2_Art212`,
`Collar_PCC2_Art206`, `Buscar_B31_3`, `Buscar_BPVC_IID`, `Buscar_BPVC_IID_B`,
`Buscar_Su`, `Buscar_Sy`, `Buscar_Prop_IID`, `Buscar_Prop_B31_3`, `Buscar_B31_B1`,
`Buscar_Ec_A2` y `Buscar_Ej_A3`. Las hojas `DB_*`, `MAP_*`, `NAV_*`, `Dashboard`,
`Instrucciones` y `Datos_Ref` **no** son hojas de motor: no llevan entradas del
ingeniero y el guardia no las mira.

**Un plan por tanda, no un plan para las cuatro fases.** El plan de implementación que
sigue a esta spec cubre **solo las Fases 1 y 2**, que no dependen de nada externo. Las
Fases 3 y 4 tendrán su propio plan cada una, cuando llegue el JSON de su norma: escribir
hoy tareas detalladas contra un archivo que no existe sería inventar su forma.

**No entra:** las listas legítimas de 2.1 (modos y booleanos); las entradas que ninguna
base publica y que por tanto se siguen tecleando (temperatura y presión de operación,
dimensiones del defecto, sobreespesor de corrosión, medidas de la reparación); el
bloque obsoleto de `Datos_Ref` filas 40-47, que se conserva tal cual como respaldo
histórico de memorias ya emitidas.

---

## 4. Diseño

Cuatro fases. Las dos primeras no dependen de nada externo; la tercera y la cuarta
esperan a que el ingeniero entregue los JSON de las normas dimensionales.

### Fase 1 — Cierre del hueco actual y el guardia

1. **Eliminar las filas D22/D23 del Art. 212.** Son descriptivas, la Sección 7 ya
   resuelve esos mismos dos materiales con la cascada auditada, y la lista que las
   alimenta espeja un bloque que `Datos_Ref` declara obsoleto. Decisión del ingeniero:
   eliminarlas, no convertirlas.
2. **Desplegables bloqueantes en todo el libro.** `dv_list()` pasa a
   `showErrorMessage=True, errorStyle="stop"`. Antes de activarlo hay que comprobar,
   celda a celda, que **ningún valor sembrado hoy caiga fuera de su propia lista**: con
   `errorStyle="stop"` una semilla inválida deja el libro abriéndose con una celda que
   Excel rechaza.
3. **El guardia.** `verificar.py` §5 ya recorre todas las validaciones del libro, pero
   con el criterio inverso al que necesitamos: hoy acepta una lista literal como
   «portable» (740-763). Se extiende para que, **en hojas de motor**, falle si una
   validación de lista (a) no es bloqueante, o (b) tiene origen literal sin estar en una
   lista blanca declarada en el builder — la de 2.1, modos y booleanos. Fuera de las
   hojas de motor el criterio actual de portabilidad se conserva intacto.

**Consecuencia sobre el *oracle* del Art. 212.** Borrar D22/D23 rompe la paridad con
`parche_art212_ref.json`, que declara esas celdas y que `TestParidadHojaParche` exige.
**No se regenera el oracle desde el build nuevo**: eso lo convertiría en un volcado de
lo que construimos y dejaría de ser un control independiente. En su lugar,
`DIVERGENCIAS_DECLARADAS`: una lista explícita de celdas que el build ya no reproduce a
propósito, cada una con su motivo, que el test consulta antes de fallar. Es el patrón
que el libro ya usa en `DIVERGENCIAS_NOMBRE_C` (1317) para el Apéndice C.

### Fase 2 — NPS y cédula dejan de ser literales

D18/D19 de los dos motores pasan a leer de la hoja dimensional con el mecanismo ya
probado en `construir_seccion7_material`: la columna de la base se **materializa en una
columna oculta** del propio motor y la validación apunta a ese rango literal. Lo exige
la regla 2 (una fórmula no puede ser origen de validación) y es exactamente lo que la
Sección 7 hace hoy con la familia de material (`FAM_COL = 11`, 5962-5971).

Al leer la cédula de la propia fila de encabezados desaparece la causa del `""&`
(2.4). **La fórmula no se toca en esta fase** — quitar el `""&` es un cambio de
comportamiento que se evalúa aparte, con el test de paridad delante.

### Fase 3 — `Datos_Ref` se convierte en `DB_B36_10` (espera JSON de B36.10M)

El ingeniero entrega el JSON de B36.10M; se verifica contra el PDF de 2.5 siguiendo el
protocolo del proyecto (**el PDF solo corrige y audita el JSON, nunca alimenta un motor
directamente**), se deposita en `resources/ASME B36/`, y el builder construye la hoja
desde ahí. La hoja se renombra a `DB_B36_10` y **sale de `HOJAS_HEREDADAS`**, que queda
solo con `Instrucciones`: ninguna hoja con datos de ingeniería vendrá ya del maestro
Rev0.

Toca todas las referencias `Datos_Ref!` de los dos motores (fórmulas de OD y espesor,
más los rangos de la Fase 2) y el árbol de navegación. El bloque obsoleto de las filas
40-47 no migra: se queda donde está, en la hoja heredada, o se retira — decisión a
tomar en el plan, no aquí.

### Fase 4 — B36.19M, acero inoxidable (espera JSON de B36.19M)

Segunda tabla dimensional (`DB_B36_19`). Aquí la cédula **deja de ser una lista plana**:
el inoxidable trae su propia serie (5S/10S/40S/80S), así que la lista de cédula pasa a
depender de qué norma dimensional aplica, y esa a su vez de la familia del material.
Es una cascada dependiente, con el mismo mecanismo que los pasos 1-4 de la Sección 7
(`cascade_formula`, 373-377). El plan de esta fase tendrá que decidir **cómo se elige
la norma dimensional**: explícita (el ingeniero la selecciona) o derivada de la familia
de material ya resuelta en la Sección 7 — con el matiz de que derivarla acopla la
Sección 1 a la Sección 7 dentro del mismo motor, lo que la regla 13 no prohíbe (habla
de motores entre sí) pero merece decidirse a la vista.

---

## 5. Criterios de éxito

1. Ninguna hoja de motor tiene una validación de lista con origen literal fuera de la lista blanca declarada.
2. Toda validación de lista del libro es bloqueante, y ningún valor sembrado cae fuera de su propia lista.
3. `Parche_PCC2_Art212` no tiene filas D22/D23, y la única divergencia con el *oracle* es la declarada en `DIVERGENCIAS_DECLARADAS`, con su motivo.
4. D18/D19 de los dos motores leen de la tabla dimensional, no de un literal.
5. `HOJAS_HEREDADAS` queda en `("Instrucciones",)` al cerrar la Fase 3.
6. Las tres suites siguen verdes y `verificar.py` vuelve 0 en cada fase, no solo al final.

## 6. Riesgos

- **Semilla fuera de lista al activar el bloqueo.** Es el riesgo real de la Fase 1: hay que barrer las 12 hojas de motor antes de activar `errorStyle="stop"`, no después.
- **El bloqueo de Excel no es total.** Actúa al teclear; no al pegar ni al escribir por macro. Reduce el error de dedo, no lo vuelve imposible. Decirlo en la hoja, no prometer de más.
- **Renombrar `Datos_Ref` tiene radio amplio** (Fase 3): fórmulas de los dos motores, `HOJAS_HEREDADAS`, `retonar_heredadas`, árbol de navegación y sus pruebas de sincronía Python/VBA.
- **La cascada de cédula por norma (Fase 4) puede acoplar Sección 1 con Sección 7** dentro del mismo motor. No lo prohíbe la regla 13, pero es una decisión de diseño que el plan debe tomar explícitamente.
- **El *oracle* deja de ser «la hoja heredada» y pasa a ser «la hoja heredada más divergencias declaradas».** Es correcto, pero solo mientras cada divergencia lleve motivo escrito y revisión; una lista que crece sin criterio vacía el control.

## 7. Dependencias externas

| Fase | Depende de | Estado |
|---|---|---|
| 1, 2 | nada | listas para planificar |
| 3 | JSON de B36.10M (lo entrega el ingeniero; PDF ya localizado para verificar) | pendiente |
| 4 | JSON de B36.19M (idem) | pendiente |
