# Reporte de verificacion — PLAN-DB-MAT-001

Libro verificado: `Motor_de_Calculo_ASME_PCC_Rev4.xlsm`  ·  73 hojas

## 1. Conteo de filas (JSON fuente -> hoja)

| Fuente | Filas JSON | Filas en la hoja | Estado |
|---|---|---|---|
| A-1 + A-4 -> DB_B31_3 | 1288 | 1288 | OK |
| 1A -> DB_BPVC_IID | 1808 | 1799 | OK — 9 filas sin identificacion ni valores (ruido de extraccion) |
| 1B + 3 -> DB_BPVC_IID_B | 1662 | 1655 | OK — 7 filas sin identificacion ni valores (ruido de extraccion) |
| U -> DB_Su | 2484 | 2473 | OK — 11 filas sin identificacion ni valores (ruido de extraccion) |
| Y-1 -> DB_Sy | 2475 | 2462 | OK — 13 filas sin identificacion ni valores (ruido de extraccion) |
| C-1 + C-2 + C-3 + C-4 -> DB_B31_C | 201 | 201 | OK |
| C-1C + C-2 + C-3C + C-4 -> DB_B31_CC | 201 | 201 | OK |
| B-1 -> DB_B31_B1 | 36 | 36 | OK |
| B-1C -> DB_B31_B1C | 36 | 36 | OK |
| A-2 -> DB_A2_Ec | 24 | 24 | OK |
| A-3 -> DB_A3_Ej | 127 | 127 | OK |
| 302.3.3-1 -> DB_Ec_Incremento | 6 | 6 | OK |

## 2. Unicidad de material_id

| Hoja | Filas | Claves unicas | Estado |
|---|---|---|---|
| DB_B31_3 | 1288 | 1288 | OK |
| DB_B31_3C | 1282 | 1282 | OK |
| DB_BPVC_IID | 1799 | 1799 | OK |
| DB_BPVC_IIDC | 1799 | 1799 | OK |
| DB_BPVC_IID_B | 1655 | 1655 | OK |
| DB_BPVC_IID_BC | 1655 | 1655 | OK |
| DB_Su | 2473 | 2473 | OK |
| DB_Sy | 2462 | 2462 | OK |
| DB_B31_C | 201 | 201 | OK |
| DB_B31_CC | 201 | 201 | OK |
| DB_B31_B1 | 36 | 36 | OK |
| DB_B31_B1C | 36 | 36 | OK |
| DB_A2_Ec | 24 | 24 | OK |
| DB_A3_Ej | 127 | 127 | OK |

## 3. Auditoria fila a fila contra el JSON fuente

No es un muestreo: de CADA fila del JSON del codigo se toma su vector completo de valores tabulados junto con su especificacion y grado, y se comprueba que exista exactamente la misma fila en la hoja. Se comparan asi todos los numeros cargados, no una seleccion.

| Hoja | Filas auditadas | Valores comparados | Filas sin correspondencia |
|---|---|---|---|
| DB_B31_3 | 1271 | 21029 | 0 |
| DB_BPVC_IID | 1799 | 31113 | 0 |
| DB_BPVC_IID_B | 1655 | 25596 | 0 |
| DB_Su | 2473 | 30312 | 0 |
| DB_Sy | 2462 | 42543 | 0 |
| DB_B31_3C | 1265 | 15231 | 0 |
| DB_BPVC_IIDC | 1799 | 24810 | 0 |
| DB_BPVC_IID_BC | 1655 | 19349 | 0 |
| DB_SuC | 2426 | 27114 | 0 |
| DB_SyC | 2452 | 34439 | 0 |

**Total: 271536 valores tabulados auditados, 0 filas sin correspondencia exacta en la hoja.**

## 3b. Auditoria de las bases por familia, auxiliares y no metalicos

| Hoja | Filas JSON | Filas en la hoja | Valores comparados | Discrepancias |
|---|---|---|---|---|
| DB_E | 135 | 135 | 1456 | 0 |
| DB_EC | 135 | 135 | 1629 | 0 |
| DB_TE_G | 53 | 53 | 1364 | 0 |
| DB_TE_GC | 53 | 53 | 1259 | 0 |
| DB_B31_C | 201 | 201 | 2795 | 0 |
| DB_B31_CC | 201 | 201 | 1931 | 0 |
| DB_B31_B1 | 36 | 36 | 360 | 0 |
| DB_B31_B1C | 36 | 36 | 360 | 0 |
| DB_A2_Ec | 24 | 24 | 96 | 0 |
| DB_A3_Ej | 127 | 127 | 635 | 0 |
| DB_Ec_Incremento | 6 | 6 | 12 | 0 |

| Hoja | Filas JSON | Filas en la hoja | Estado |
|---|---|---|---|
| DB_PRD | 115 | 115 | OK |
| DB_PRDC | 115 | 115 | OK |
| DB_TE | 605 | 641 | OK |
| DB_TEC | 544 | 580 | OK |
| MAP_Factores | 151 | 151 | OK |
| Notas_Codigo | 325 | 325 | OK |
| Apendice B (B-2..B-6) | no cargadas (marcador) | 0 | OK |

## 4. Contiguidad de los bloques de la cascada

Las listas dependientes se resuelven con OFFSET/MATCH/COUNTIF, que exige que todas las filas de una misma clave sean consecutivas.

| Hoja | k0 | k1 | k2 | k3 | k4 | Estado |
|---|---|---|---|---|---|---|
| DB_B31_3 | 0 | 0 | 0 | 0 | 0 | OK |
| DB_B31_3C | 0 | 0 | 0 | 0 | 0 | OK |
| DB_BPVC_IID | 0 | 0 | 0 | 0 | 0 | OK |
| DB_BPVC_IIDC | 0 | 0 | 0 | 0 | 0 | OK |
| DB_BPVC_IID_B | 0 | 0 | 0 | 0 | 0 | OK |
| DB_BPVC_IID_BC | 0 | 0 | 0 | 0 | 0 | OK |
| DB_Su | 0 | 0 | 0 | 0 | 0 | OK |
| DB_Sy | 0 | 0 | 0 | 0 | 0 | OK |
| DB_B31_C | 0 | 0 | 0 | 0 | 0 | OK |
| DB_B31_CC | 0 | 0 | 0 | 0 | 0 | OK |
| DB_B31_B1 | 0 | 0 | 0 | 0 | 0 | OK |
| DB_B31_B1C | 0 | 0 | 0 | 0 | 0 | OK |
| DB_A2_Ec | 0 | 0 | 0 | 0 | 0 | OK |
| DB_A3_Ej | 0 | 0 | 0 | 0 | 0 | OK |

(0 = ningun bloque fragmentado)

## 5. Compatibilidad — ausencia de funciones de matriz dinamica

Formulas con funciones de matriz dinamica encontradas: **0**

Validacion de datos: Google Sheets solo admite un RANGO literal o una lista de items como origen; una formula (OFFSET/INDIRECT) se pierde al importar.

Validaciones de lista revisadas en todo el libro; con origen NO portable: **0**

Regla 12/14: en una hoja de motor, una validacion de lista debe (a) bloquear y (b) tener origen de RANGO, no una lista literal — salvo los modos y booleanos declarados en LITERALES_PERMITIDOS.

Validaciones infractoras en hojas de motor: **0**

## 6. Interpolacion con huecos, modo tabulado y bordes (recalculo en hoja)

| Base | material_id | T | Modo | T1 | T2 | S(T) hoja | S(T) referencia | Estado |
|---|---|---|---|---|---|---|---|---|
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 25 | Interpolado | 40 | 65 | 138 | 138 | OK |
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 110 | Interpolado | 100 | 150 | 138 | 138.0 | OK |
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 125 | Interpolado | 100 | 150 | 138 | 138.0 | OK |
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 212 | Interpolado | 200 | 250 | 136.56 | 136.56 | OK |
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 263 | Interpolado | 250 | 300 | 130.44 | 130.44 | OK |
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 337 | Interpolado | 325 | 350 | 120.08 | 120.08 | OK |
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 1200 | Interpolado | 600 | None | 6.89 | 6.89 | OK |
| DB_B31_3 | `A-1 | A106 | B | Pipe & tube |` | 263 | Tabulado-conservador | 250 | 300 | 126 | 126 | OK |
| DB_B31_3 | `A-1 | A516 | 70 | Plate, bar, ` | 25 | Interpolado | 40 | 65 | 161 | 161 | OK |
| DB_B31_3 | `A-1 | A516 | 70 | Plate, bar, ` | 110 | Interpolado | 100 | 150 | 158 | 158.0 | OK |
| DB_B31_3 | `A-1 | A516 | 70 | Plate, bar, ` | 263 | Interpolado | 250 | 300 | 141.92 | 141.92 | OK |
| DB_B31_3 | `A-1 | A516 | 70 | Plate, bar, ` | 337 | Interpolado | 325 | 350 | 130.08 | 130.08 | OK |
| DB_B31_3 | `A-1 | A516 | 70 | Plate, bar, ` | 263 | Tabulado-conservador | 250 | 300 | 136 | 136 | OK |
| DB_B31_3 | `A-1 | A312 | TP321 | Smls. pip` | 25 | Interpolado | 40 | 65 | 115 | 115 | OK |
| DB_B31_3 | `A-1 | A312 | TP321 | Smls. pip` | 263 | Interpolado | 250 | 275 | 110.44 | 110.44 | OK |
| DB_B31_3 | `A-1 | A312 | TP321 | Smls. pip` | 337 | Interpolado | 325 | 350 | 103.04 | 103.04 | OK |
| DB_B31_3 | `A-1 | A312 | TP321 | Smls. pip` | 263 | Tabulado-conservador | 250 | 275 | 109 | 109 | OK |
| DB_BPVC_IID | `1A | SA-516 | 70 | Plate | K02` | 25 | Interpolado | 40 | 65 | 138 | 138 | OK |
| DB_BPVC_IID | `1A | SA-516 | 70 | Plate | K02` | 180 | Interpolado | 150 | 200 | 138 | 138.0 | OK |
| DB_BPVC_IID | `1A | SA-516 | 70 | Plate | K02` | 263 | Interpolado | 250 | 300 | 137.48 | 137.48 | OK |
| DB_BPVC_IID | `1A | SA-516 | 70 | Plate | K02` | 263 | Tabulado-conservador | 250 | 300 | 136 | 136 | OK |
| DB_BPVC_IID | `1A | SA-516 | 70 | Plate | K02` | 900 | Interpolado | 550 | None | 12.9 | 12.9 | OK |

**22 casos, 0 fallos.** El caso T = 125 C de A106 Gr.B verifica el salto de huecos interiores: la Tabla A-1 no imprime ese punto para ese material y la hoja interpola entre 100 y 150 C, no entre celdas vacias.

### 6b. Apendice C — bloqueo en los dos extremos y rama de dato puntual

El Apendice C no publica columna «Temp. max.»: el limite es el primer y el ultimo punto que tabula la propia fila. Estos casos comprueban, recalculando en Excel la misma expresion que lleva el motor, que por encima y por debajo de esa banda el resultado queda BLOQUEADO en vez de sostener el valor del extremo, y que C-2 y C-4 se resuelven por su valor unico impreso.

| Hoja | material_id | T | Tipo | Estado hoja | Estado referencia | Valor hoja | Valor referencia | Estado |
|---|---|---|---|---|---|---|---|---|
| DB_B31_C | `C-3 | Carbon steels with carbon co` | 25 | CURVA | EN RANGO | EN RANGO | 202000 | 202000.0 | OK |
| DB_B31_C | `C-3 | Carbon steels with carbon co` | 375 | CURVA | EN RANGO | EN RANGO | 175000 | 175000.0 | OK |
| DB_B31_C | `C-3 | Carbon steels with carbon co` | 700 | CURVA | FUERA DE RANGO — T por encima del  | FUERA DE RANGO — T por encima del  | BLOQUEADO | BLOQUEADO | OK |
| DB_B31_C | `C-3 | Carbon steels with carbon co` | -300 | CURVA | FUERA DE RANGO — T por debajo del  | FUERA DE RANGO — T por debajo del  | BLOQUEADO | BLOQUEADO | OK |
| DB_B31_C | `C-1 | Group 1 carbon and low alloy` | 400 | CURVA | EN RANGO | EN RANGO | 13.8 | 13.8 | OK |
| DB_B31_C | `C-1 | Group 1 carbon and low alloy` | 412 | CURVA | EN RANGO | EN RANGO | 13.896 | 13.896 | OK |
| DB_B31_C | `C-4 | Acetal | (unico)` | 25 | PUNTO | VALOR UNICO — no depende de T; vea | VALOR UNICO — no depende de T; vea | 2830 | 2830 | OK |
| DB_B31_C | `C-4 | Acetal | (unico)` | 300 | PUNTO | VALOR UNICO — no depende de T; vea | VALOR UNICO — no depende de T; vea | 2830 | 2830 | OK |
| DB_B31_C | `C-2 | Glass-epoxy, filament-wound ` | 25 | PUNTO | VALOR UNICO — no depende de T; vea | VALOR UNICO — no depende de T; vea | 16–23.5 | 16–23.5 | OK |
| DB_B31_CC | `C-3C | Carbon steels with carbon c` | 662 | CURVA | EN RANGO | EN RANGO | 25880000 | 25880000.0 | OK |

**10 casos del Apendice C, 0 fallos.**

### 6d. Tabla B-1 — orden de los bloqueos y excepcion de la Nota (3)

La Tabla B-1 tiene tres reglas de rango que no son las de los metales, y las tres se recalculan aqui con la misma expresion que lleva el motor. Se INTERPOLA linealmente (para. A302.3.1(b)). Por debajo de la primera temperatura tabulada NO se extrapola: se sostiene ese HDS, porque lo manda la Nota (3) de la propia tabla. Y se bloquea por arriba en DOS sitios distintos —el limite maximo recomendado de las Notas (1) y (2) y el ultimo punto tabulado (para. A323.2.1(a))— con aviso distinto para cada uno; los casos de F441/CPVC4120-05 a 90 y a 95 °C son justo el par que separa los dos.

| Hoja | material_id | T | Modo | Estado hoja | Estado referencia | HDS hoja | HDS referencia | Estado |
|---|---|---|---|---|---|---|---|---|
| DB_B31_B1 | `B-1 | PVC1120 | D1785 | Sch. 40, 80, 1` | 23 | Interpolado | EN RANGO | EN RANGO | 13.8 | 13.8 | OK |
| DB_B31_B1 | `B-1 | PVC1120 | D1785 | Sch. 40, 80, 1` | 30 | Interpolado | FUERA DE RANGO — T por encima del ulti | FUERA DE RANGO — T por encima del ulti | BLOQUEADO | BLOQUEADO | OK |
| DB_B31_B1 | `B-1 | PVC1120 | D1785 | Sch. 40, 80, 1` | 10 | Interpolado | FUERA DE RANGO — T por debajo del limi | FUERA DE RANGO — T por debajo del limi | BLOQUEADO | BLOQUEADO | OK |
| DB_B31_B1 | `B-1 | CPVC4120-05 | F441 | Sch. 40, 80` | 23 | Interpolado | EN RANGO | EN RANGO | 13.8 | 13.8 | OK |
| DB_B31_B1 | `B-1 | CPVC4120-05 | F441 | Sch. 40, 80` | 60 | Interpolado | EN RANGO | EN RANGO | 7.375 | 7.375 | OK |
| DB_B31_B1 | `B-1 | CPVC4120-05 | F441 | Sch. 40, 80` | 60 | Tabulado-con | EN RANGO | EN RANGO | 3.45 | 3.45 | OK |
| DB_B31_B1 | `B-1 | CPVC4120-05 | F441 | Sch. 40, 80` | 90 | Interpolado | FUERA DE RANGO — T por encima del ulti | FUERA DE RANGO — T por encima del ulti | BLOQUEADO | BLOQUEADO | OK |
| DB_B31_B1 | `B-1 | CPVC4120-05 | F441 | Sch. 40, 80` | 95 | Interpolado | FUERA DE RANGO — T por encima del limi | FUERA DE RANGO — T por encima del limi | BLOQUEADO | BLOQUEADO | OK |
| DB_B31_B1 | `B-1 | ABS | (sin Spec. No. impreso) | ` | 20 | Interpolado | SIN HDS TABULADO — la fila solo public | SIN HDS TABULADO — la fila solo public | SIN HDS TABULADO — la fila solo public | SIN HDS TABULADO — la fila solo public | OK |
| DB_B31_B1 | `B-1 | PEX0006 | F2788/F2788M | SDR/DR-` | 0 | Interpolado | EN RANGO — Nota (3): se sostiene el HD | EN RANGO — Nota (3): se sostiene el HD | 4.34 | 4.34 | OK |
| DB_B31_B1C | `B-1C | CPVC4120-05 | F441 | Sch. 40, 8` | 140 | Interpolado | EN RANGO | EN RANGO | 1.0699999999999998 | 1.0699999999999998 | OK |

**11 casos de la Tabla B-1, 0 fallos.**

### 6c. Factores de calidad — el motor muestra el factor de la fila, sin redondeo

Ec y Ej son escalares: no hay interpolacion que recalcular. Lo que se comprueba, conduciendo la cascada y recalculando en Excel, es que el KPI es EXACTAMENTE el valor impreso y que el incremento de la Tabla 302.3.3-1 solo se aplica donde la fila lo admite.

| Base | Fila | Examen | Factor hoja | Factor codigo | Admite | Aplicable hoja | Aplicable referencia | Estado |
|---|---|---|---|---|---|---|---|---|
| DB_A2_Ec | Spec. No.=A395 · Descripcion=Ductile and ferri | (1) and (3)(a) or (3)(b) | 0.8 | 0.8 | SI — admite incremen | 1 | 1 | OK |
| DB_A2_Ec | Spec. No.=A451 · Grupo impreso=Stainless Steel | (1) | 0.9 | 0.9 | SI — admite incremen | 0.9 | 0.9 | OK |
| DB_A2_Ec | Spec. No.=A451 · Grupo impreso=Stainless Steel | (1) and (3)(a) or (3)(b) | 0.9 | 0.9 | SI — admite incremen | 1 | 1 | OK |
| DB_A2_Ec | Spec. No.=A426 | (1) and (3)(a) or (3)(b) | 1 | 1 | NO — el factor ya su | 1 | 1 | OK |
| DB_A2_Ec | Spec. No.=A47 | (1) and (3)(a) or (3)(b) | 1 | 1 | El codigo no se pron | 1 | 1 | OK |
| DB_A3_Ej | Spec. No.=API 5L · Descripcion=Continuous weld | — | 0.6 | 0.6 | VER para. 302.3.4(b) | 0.6 | 0.6 | OK |
| DB_A3_Ej | Spec. No.=API 5L · Descripcion=Seamless pipe | — | 1 | 1 | VER para. 302.3.4(b) | 1 | 1 | OK |
| DB_A3_Ej | Spec. No.=A312 · Descripcion=Electric fusion w | — | 0.8 | 0.8 | VER para. 302.3.4(b) | 0.8 | 0.8 | OK |
| DB_A3_Ej | Spec. No.=A312 · Descripcion=Electric fusion w | — | 0.85 | 0.85 | VER para. 302.3.4(b) | 0.85 | 0.85 | OK |

Y el mismo calculo conducido por la cascada real de cada motor, que es lo que comprueba el cableado de las listas desplegables:

| Motor | Fila | Factor hoja | Factor codigo | Admite hoja | Aplicable hoja | Aplicable referencia | Estado |
|---|---|---|---|---|---|---|---|
| Buscar_Ec_A2 | Spec. No.=A395 · Descripcion=Ductile and ferri | 0.8 | 0.8 | SI — admite incremen | 1 | 1 | OK |
| Buscar_Ej_A3 | Spec. No.=API 5L · Descripcion=Continuous weld | 0.6 | 0.6 | VER para. 302.3.4(b) | None | None | OK |

**11 casos de factores, 0 fallos.**

## 7. Regresion del caso semilla (collar 12"-CWS-46-032-B1)

Temperatura de evaluacion: **25 °C** · metal base `A-1 | A106 | B | Pipe & tube | K03006 | ` · collar `A-1 | A516 | 70 | Plate, bar, shps., she`

| Magnitud | Referencia Python (MPa) | Hoja recalculada (MPa) | Dictamen | Estado |
|---|---|---|---|---|
| Sa collar (A516 Gr.70) | 161 | 161 | OK | OK |
| Sa metal base (A106 Gr.B) | 138 | 138 | OK | OK |
| Sa gobernante | 138 | 138 | — | OK |

Dictamen global del modulo: **APTO** (OK).

## 8. Capa de navegacion (Dashboard y proyecto VBA)

| Comprobacion | Detalle | Estado |
|---|---|---|
| Unica hoja visible es el Dashboard | Dashboard | OK |
| Las 37 hojas navegables estan hidden | 37 hojas | OK |
| El resto esta veryHidden | 35 hojas | OK |
| Ninguna base que alimente un motor es alcanzable desde la UI |  | OK |
| El paquete conserva xl/vbaProject.bin | .xlsm | OK |
| Los botones del arbol cubren las 37 hojas navegables | 38 destinos distintos en 16 hojas | OK |
| Cada hoja navegable vuelve a SU PADRE, no a la raiz |  | OK |
| Toda hoja navegable se alcanza desde el Dashboard | huerfanas: ninguna | OK |

La visibilidad esta grabada en el archivo, no la impone la macro: con las macros bloqueadas el usuario sigue sin ver ninguna base de datos.

## 9. Mapeo de grupos de propiedades (MAP_Grupo y MAP_GrupoC -> Notas de TM-1 / TE-1)

Cada hoja se audita contra las Notas de SU edicion: no numeran igual, asi que cruzarlas ocultaria una cita mal puesta.

| Comprobacion | Detalle | Estado |
|---|---|---|
| Todo grupo asignado cita su fuente | 0 filas con grupo y sin fuente | OK |
| La Nota citada lista esa composicion | 0 citas que el JSON del codigo no respalda | OK |
| No sobrevive ningun estado de conjetura | sin filas 'PROPUESTA' | OK |
| Las dos ediciones traen sus Notas de grupo | 4/4 archivos con note_members | OK |
| Lo decidido por una persona se declara como tal | 0 filas VALIDADO sin marca en la fuente | OK |
| Toda composicion tomada por UNS cita su tabla de origen | 0 filas sin citar la tabla de origen | OK |
| La columna de TE-1 citada existe en esa edicion | 0 citas a una columna que el JSON no imprime | OK |

| Estado del mapeo | Filas |
|---|---:|
| AUTO (composicion en Nota) | 2594 |
| AUTO (UNS exacto) | 2584 |
| SIN MAPEO | 1182 |
| AUTO (composicion via UNS en otra tabla) | 456 |
| VALIDADO POR INGENIERO | 92 |

Las filas SIN MAPEO no son un defecto de la extraccion: son materiales para los que II-D no publica modulo ni dilatacion. En ellas el calculo queda bloqueado, que es lo que exige el codigo.

## 10. Seccion II, partes A, B y C (volcado integro y normalizadas)

La comprobacion sin perdida se relee DESDE LA HOJA: la concatenacion de las celdas C01..Cnn de cada fila grabada tiene que coincidir, caracter a caracter, con la del texto de sus bloques `Line` de origen. Es lo unico que se puede garantizar sin el PDF, y garantiza que esta capa reparte el texto impreso sin anadir ni quitar nada.

| Comprobacion | Detalle | Estado |
|---|---|---|
| Las 9 hojas de la Seccion II estan en el libro | faltan: ninguna | OK |
| Las 9 estan en la capa de navegacion | CAT_/IDX_/DB_SecII_* en NAVEGABLES | OK |
| Ninguna de las 9 hojas lleva formulas | 0 celdas empiezan por '=' | OK |
| Cada fila impresa aparece una sola vez | 0 claves (spec, tabla, fila) repetidas | OK |
| La hoja conserva el texto impreso, caracter a caracter | 0 de 54198 filas en las que la concatenacion de las celdas grabadas no coincide con la de sus `Line` | OK |
| IDX_SecII_Tablas cuadra con el volcado | 0 tablas cuyo recuento de filas no coincide | OK |

Una fila AMBIGUA no es un fallo: es una fila que no se pudo repartir en columnas sin adivinar, y cuyo texto impreso se conserva ENTERO en C01. Cuantas hay, y por que, esta en `IDX_SecII_Tablas`.

## 11. Bases dimensionales B36 (hoja contra JSON de resources)

- B36.10M: 779 filas x 14 columnas auditadas.
- B36.19M: 114 filas x 14 columnas auditadas.
Discrepancias: **0**

## Resultado

| Seccion | Fallos |
|---|---|
| 1. Conteos | 0 |
| 2. Unicidad | 0 |
| 3. Auditoria fila a fila | 0 |
| 4. Contiguidad de la cascada | 0 |
| 5. Portabilidad de formulas | 0 |
| 5b. Guardia listas fijas en motor (regla 12/14) | 0 |
| 6. Interpolacion recalculada | 0 |
| 7. Caso semilla | 0 |
| 8. Capa de navegacion | 0 |
| 9. Mapeo de grupos | 0 |
| 10. Seccion II A/B/C | 0 |
| 11. Bases dimensionales B36 | 0 |

**Total de fallos: 0.**