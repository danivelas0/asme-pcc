# Reporte de verificacion — PLAN-DB-MAT-001 Rev. 2

Libro verificado: `Motor_v2.xlsx`  ·  37 hojas

## 1. Conteo de filas (JSON fuente -> hoja)

| Fuente | Filas JSON | Filas en la hoja | Estado |
|---|---|---|---|
| A-1 + A-4 -> DB_B31_3 | 1288 | 1288 | OK |
| 1A -> DB_BPVC_IID | 1808 | 1799 | OK — 9 filas sin identificacion ni valores descartadas (ruido de extraccion) |
| 1B + 3 -> DB_BPVC_IID_B | 1662 | 1655 | OK — 7 filas sin identificacion ni valores descartadas (ruido de extraccion) |
| U -> DB_Su | 2484 | 2473 | OK — 11 filas sin identificacion ni valores descartadas (ruido de extraccion) |
| Y-1 -> DB_Sy | 2475 | 2462 | OK — 13 filas sin identificacion ni valores descartadas (ruido de extraccion) |

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

## 3. Auditoria fila a fila contra el JSON fuente

No es un muestreo: de CADA fila del JSON del codigo se toma su vector completo de valores tabulados junto con su especificacion y grado, y se comprueba que exista exactamente la misma fila en la hoja. Se comparan asi todos los numeros cargados, no una seleccion.

| Hoja | Filas auditadas | Valores comparados | Filas sin correspondencia |
|---|---|---|---|
| DB_B31_3 | 1271 | 21029 | 0 |
| DB_BPVC_IID | 1799 | 31113 | 0 |
| DB_BPVC_IID_B | 1655 | 25596 | 0 |
| DB_Su | 2473 | 30312 | 0 |
| DB_Sy | 2462 | 42543 | 0 |
| DB_B31_3C | 1265 | 14971 | 0 |
| DB_BPVC_IIDC | 1799 | 24810 | 0 |
| DB_BPVC_IID_BC | 1655 | 19349 | 0 |
| DB_SuC | 2426 | 27114 | 0 |
| DB_SyC | 2452 | 34439 | 0 |

**Total: 271276 valores tabulados auditados, 0 filas sin correspondencia exacta en la hoja.**

## 3b. Auditoria de las bases por familia, auxiliares y no metalicos

| Hoja | Filas JSON | Filas en la hoja | Valores comparados | Discrepancias |
|---|---|---|---|---|
| DB_E | 135 | 135 | 1456 | 0 |
| DB_EC | 135 | 135 | 1629 | 0 |
| DB_C_dilatacion | 52 | 52 | 1568 | 0 |
| DB_C_dilatacionC | 52 | 52 | 777 | 0 |
| DB_C_modulo | 75 | 75 | 1109 | 0 |
| DB_C_moduloC | 75 | 75 | 1036 | 0 |

| Hoja | Filas JSON | Filas en la hoja | Estado |
|---|---|---|---|
| DB_PRD | 115 | 115 | OK |
| DB_PRDC | 115 | 115 | OK |
| DB_TE | 605 | 641 | OK |
| DB_TEC | 544 | 580 | OK |
| MAP_Factores | 151 | 151 | OK |
| Notas_Codigo | 325 | 325 | OK |
| DB_NoMetalicos (pares campo/valor) | 931 | 931 | OK |

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

(0 = ningun bloque fragmentado)

## 5. Compatibilidad — ausencia de funciones de matriz dinamica

Formulas con funciones de matriz dinamica encontradas: **0**

Validacion de datos: Google Sheets solo admite un RANGO literal o una lista de items como origen; una formula (OFFSET/INDIRECT) se pierde al importar.

Validaciones de lista revisadas en todo el libro; con origen NO portable: **0**

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

## 7. Regresion del caso semilla (collar 12"-CWS-46-032-B1)

| Magnitud | Antes (Datos_Ref) | Ahora (DB ASME) | Dictamen |
|---|---|---|---|
| Sa collar (A516 Gr.70) | 160,7 | 161 | OK |
| Sa metal base (A106 Gr.B) | 137,9 | 138 | OK |
| Sa gobernante | 137,9 | 138 | — |

Dictamen global del modulo: **APTO**.