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

### 2.6 Esquema real de los JSON de B36.10M y B36.19M (entregados 2026-09-10)

El ingeniero entregó las dos extracciones. **A diferencia de la Sección II, aquí los
bloques `Table` sí traen `<table>` HTML real**: no hay que recomponer columnas desde
`Line` y `bbox`.

| Norma | Tablas de dimensiones | Filas | Columnas |
|---|---|---:|---|
| B36.10M | 19 (págs. 13-31) | ~800 | `NPS (DN)`, `Identification [Note (1)]`, `Schedule No.`, `Outside Diameter, in. (mm)`, `Wall Thickness, in. (mm)`, `Plain End Weight (Mass), lb/ft (kg/m)` |
| B36.19M | 3 (págs. 11-13) | ~117 | igual, **sin** `Identification` (el inoxidable designa por 5S/10S/40S/80S en `Schedule No.`) |

Filas reales de muestra (pág. 13):

```
['1/8 (6)', '...', '10',  '0.405 (10.29)', '0.049 (1.24)', '0.19 (0.28)']
['1/8 (6)', 'STD', '40',  '0.405 (10.29)', '0.068 (1.73)', '0.24 (0.37)']
['1/8 (6)', 'XXS', '...', '0.405 (10.29)', '0.190 (4.83)', '0.44 (0.65)']
```

Cuatro consecuencias de diseño, todas para el plan:

1. **Una fila por (NPS, cédula)**, no la matriz 33×14 de `Datos_Ref`. Es la forma que quiere una cascada y la que usan las demás `DB_*`.
2. **`'...'` es el marcador de «no aplica» impreso por el código.** XXS no tiene número de cédula; la Sch 10 no tiene identificación. Se conserva tal cual (regla 9); convertirlo a vacío en silencio sería inventar.
3. **Los dos sistemas de unidades vienen en la MISMA celda** (`0.405 (10.29)`). Por la regla 10, el conmutador cambia de **columna**, no de hoja: no hay edición US aparte que cargar. Hay que partir cada celda en sus dos magnitudes al construir la base.
4. **Doble designación.** Una misma tubería se alcanza por `Schedule No.` o por `Identification`, a veces por las dos y a veces por una sola. La cascada tiene que dejar elegir por cualquiera de las dos sin duplicar filas ni obligar al ingeniero a saber de antemano cuál publica su tubería.

Además, el `NPS (DN)` llega como `1/8 (6)` —fraccionario más DN— mientras `Datos_Ref`
usa decimal (`0.5`, `12`, `48`). El caso semilla del Art. 212 es NPS 12: la conversión
de representación tiene que dejarlo resolviendo igual.

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

**Un plan, tres fases.** Con los JSON ya entregados (§2.6) nada queda bloqueado, así
que el plan de implementación cubre las tres fases seguidas. Cada fase cierra con sus
gates en verde (las tres suites y `verificar.py`), de modo que sigue siendo posible
detenerse entre fases con el libro en un estado entregable.

**Dónde están los JSON entregados** (rutas de la sesión del 2026-09-10; la primera
tarea del plan es depositarlos en `resources/ASME B36/`, porque el directorio de
subidas es de sesión y no es un sitio del que dependa el proyecto):

```
C:\Users\User\.claude\uploads\b554f20f-43a0-4c53-b8e7-767a3321bf31\
  974d90fb-datalaboutputASME_B36.10M__Welded_and_Seamless_Wroth_Steel_Pipe_2022.pdf.json
  ecc8b41b-datalaboutputASME_B36.19M__Stainless_Steel_Pipe_2022.pdf.json
```

**No entra:** las listas legítimas de 2.1 (modos y booleanos); las entradas que ninguna
base publica y que por tanto se siguen tecleando (temperatura y presión de operación,
dimensiones del defecto, sobreespesor de corrosión, medidas de la reparación); el
bloque obsoleto de `Datos_Ref` filas 40-47, que se conserva tal cual como respaldo
histórico de memorias ya emitidas.

---

## 4. Diseño

> **Revisión del 2026-09-10, tras recibir los JSON.** El diseño original tenía cuatro
> fases, con la 3 y la 4 esperando norma. Con las dos extracciones ya entregadas
> (§2.6) y la decisión del ingeniero de **retirar `Datos_Ref`**, las fases 2 y 3
> originales se colapsan: no tiene sentido apuntar D18/D19 a `Datos_Ref` para
> repuntarlas a `DB_B36_10` inmediatamente después. Quedan **tres fases**, ninguna
> bloqueada, y B36.19M entra junto con B36.10M en la misma fase de extracción en vez
> de en una fase propia — son la misma clase de trabajo sobre el mismo esquema.

Tres fases. Ninguna depende ya de una entrega externa.

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

### Fase 2 — Extracción de B36.10M y B36.19M, y sus dos bases

Las dos normas entran juntas: mismo esquema, misma clase de trabajo (§2.6).

1. **Depositar las extracciones en `resources/ASME B36/`**, con la misma disciplina de
   procedencia que el resto del árbol (`meta.json` con origen y fecha).
2. **Verificarlas contra los PDF** de §2.5, con el protocolo del proyecto: el PDF
   **solo corrige y audita el JSON, nunca alimenta un motor directamente**. La
   verificación tiene que apoyarse en algo que aparezca **una sola vez** —una fila
   concreta de una tabla—, nunca en el rótulo de la tabla, que se repite en cada
   página de continuación y casa con cualquier convención de folio (la trampa ya
   pagada, documentada en `CLAUDE.md`).
3. **Construir `DB_B36_10` y `DB_B36_19`** desde esos JSON, una fila por
   (NPS, cédula), partiendo cada celda de doble unidad en sus dos magnitudes y
   conservando `'...'` tal como lo imprime el código.

Nada de esta fase toca todavía a los motores: al terminar, las bases existen y están
auditadas, y los motores siguen leyendo de `Datos_Ref`.

### Fase 3 — Los motores leen de las bases nuevas y `Datos_Ref` se retira

1. **D18/D19 de los dos motores pasan a leer de las bases**, con el mecanismo ya
   probado en `construir_seccion7_material`: la columna de la base se **materializa en
   una columna oculta** del propio motor y la validación apunta a ese rango literal
   (la regla 2 prohíbe una fórmula como origen; es lo que la Sección 7 hace hoy con la
   familia de material, `FAM_COL = 11`, 5962-5971).
2. **La cédula pasa a ser lista dependiente**, no plana: depende de la norma
   dimensional y del NPS elegido — `cascade_formula` (373-377), el mismo mecanismo de
   los pasos 1-4 de la Sección 7.
3. **Las fórmulas de OD y espesor** dejan de hacer `INDEX/MATCH` contra la matriz
   33×14 y pasan a resolver contra la fila de la base. Aquí desaparece la causa del
   `""&` (§2.4): es un cambio de comportamiento y se hace con el test de paridad
   delante, no de paso.
4. **`Datos_Ref` se retira del libro** y `HOJAS_HEREDADAS` queda en
   `("Instrucciones",)`. Ninguna hoja con datos de ingeniería vendrá ya del maestro
   Rev0. Decisión del ingeniero (2026-09-10): retirarla, no renombrarla.

**Tres decisiones que esta fase tiene que tomar explícitamente:**

- **Cómo se elige la norma dimensional.** Explícita (el ingeniero selecciona B36.10M o B36.19M) o derivada de la familia de material ya resuelta en la Sección 7. Derivarla acopla la Sección 1 con la Sección 7 dentro del mismo motor — la regla 13 no lo prohíbe (habla de motores entre sí), pero es acoplamiento y merece decidirse a la vista.
- **Cómo se designa la cédula con doble designación** (§2.6.4): una tubería se alcanza por `Schedule No.`, por `Identification`, por ambas o por una sola. Hay que dejar elegir por cualquiera de las dos sin duplicar filas ni exigir que el ingeniero sepa de antemano cuál publica su tubería.
- **Qué representación toma el NPS.** La base trae `1/8 (6)`; `Datos_Ref` usa decimal. El caso semilla es NPS 12 y tiene que seguir resolviendo con el mismo dictamen.

**El bloque obsoleto de `Datos_Ref` filas 40-47** —esfuerzos admisibles de seis
materiales, rotulado por la propia hoja `[OBSOLETO — ver DB_B31_3 / DB_BPVC_IID] … se
conserva como respaldo historico de las memorias ya emitidas`— **se retira con la
hoja**: el dato vivo está en las bases auditadas y el libro anterior queda en el
historial de git. Si el ingeniero prefiere conservarlo, se mueve a una hoja de
respaldo antes de borrar, no se deja en medio.

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

Ninguna pendiente. El ingeniero entregó los dos JSON el 2026-09-10 (§2.6) y los PDF
para verificarlos están localizados (§2.5). Las tres fases se pueden planificar y
ejecutar seguidas.
