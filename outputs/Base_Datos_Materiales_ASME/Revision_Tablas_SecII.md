# Revision de las tablas de ASME BPVC Seccion II, partes A, B y C

Generado por `secii_tablas.py`. Mide la reconstruccion de las tablas desde los bloques `Line` y sus `bbox`; no escribe nada en `resources/` ni en el libro.

## Resumen

| Magnitud | Valor |
|---|---:|
| Especificaciones recorridas | 368 |
| Tablas logicas (tras unir continuaciones) | 2572 |
| de ellas, continuadas de pagina | 219 |
| de ellas, sueltas (sin TableGroup) | 732 |
| de ellas, sin titulo localizable | 143 |
| Filas reconstruidas | 54198 |
| Bloques `Line` consumidos | 124115 |
| Celdas escritas | 218202 |
| Notas al pie recogidas | 5233 |

## Reparto por confianza

Se dan dos columnas de porcentaje. La segunda es la que importa: una banda o un encabezado que no se deja partir es un rotulo, y cuesta mucho menos que una fila de DATOS que no se deja partir, que son numeros del codigo. Una fila cuenta como dato si su texto trae alguna ficha de valor.

| Confianza | Filas | % de todas | Filas de dato | % de las de dato |
|---|---:|---:|---:|---:|
| EXACTA | 16451 | 30.4 % | 9189 | 25.4 % |
| POR CONTEO | 12935 | 23.9 % | 12935 | 35.8 % |
| AMBIGUA | 24812 | 45.8 % | 13996 | 38.7 % |

**Comprobacion sin perdida: 0 filas en las que la concatenacion de celdas no coincide caracter a caracter con la de sus `Line` de origen.** Cero es la unica cifra aceptable: significa que esta capa reparte el texto impreso y no anade ni quita nada.

## Huecos declarados (no reparables sin el PDF)

| Hueco | Bloques |
|---|---:|
| Caption vacio | 145 |
| Form sin contenido (parte C) | 101 |
| Table sin ningun Line | 22 |

## Tablas con filas AMBIGUAS

Una fila AMBIGUA conserva su texto impreso ENTERO en una celda: no se pierde nada, pero tampoco queda tabulada. Se listan las 60 tablas con mas filas ambiguas.

| Parte | Spec | Tabla | Titulo | Pags. | Cols. | Filas | AMBIGUA | Motivo dominante |
|---|---|---:|---|---|---:|---:|---:|---|
| bpvc_ii_a_1 | SA-6/SA-6M | 69 | TABLE A2.1 | 102-103 | 17 | 291 | 261 | linea unica que no cuadra con 17 columnas |
| bpvc_ii_a_1 | SA-53/SA-53M | 7 | (sin titulo) | 204-205 | 6 | 260 | 151 | linea unica que no cuadra con 6 columnas |
| bpvc_ii_a_1 | SA-182/SA-182M | 4 | TABLE 2 | 271-272 | 9 | 166 | 122 | linea unica que no cuadra con 9 columnas |
| bpvc_ii_a_1 | SA-240/SA-240M | 4 | TABLE 2 | 401-402 | 11 | 200 | 108 | dos Line caen en la misma banda de columna |
| bpvc_ii_c | SFA-5.5/SFA-5.5M | 12 | Table 7 Preheat, Interpass, and Postweld Heat Treatm | 242-243 | 3 | 166 | 108 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_b | SB-221M | 3 | TABLE 2 | 412-413 | 15 | 197 | 103 | linea unica que no cuadra con 15 columnas |
| bpvc_ii_a_2 | SA-564/SA-564M | 4 | TABLE 4 MECHANICAL TEST REQUIREMENTS AFTER AGE HARDE | 277-278 | 3 | 204 | 98 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_a_2 | SA-705/SA-705M | 3 | TABLE 3 MECHANICAL TEST REQUIREMENTS AFTER AGE HARDE | 475-476 | 26 | 141 | 91 | dos Line caen en la misma banda de columna |
| bpvc_ii_c | SFA-5.28/SFA-5.28M | 2 | Table 2 Tension Test Requirements | 918-919 | 7 | 113 | 80 | linea unica que no cuadra con 7 columnas |
| bpvc_ii_c | SFA-5.29/SFA-5.29M | 1 | Table 1 A5.29 [A5.29M] Mechanical Property Requireme | 961-962 | 6 | 94 | 79 | el reparto por bandas alteraria el orden de lectura  |
| bpvc_ii_a_1 | SA-29/SA-29M | 3 | TABLE 1 Grade Designations and Chemical Compositions | 158 | 6 | 74 | 74 | linea unica que no cuadra con 6 columnas |
| bpvc_ii_a_1 | SA-276 | 4 | TABLE 2 | 466-467 | 13 | 122 | 73 | linea unica que no cuadra con 13 columnas |
| bpvc_ii_a_1 | SA-370 | 13 | TABLE 6 Brinell Hardness NumbersA (Ball 10 mm in Dia | 653 | 16 | 76 | 70 | linea unica que no cuadra con 16 columnas |
| bpvc_ii_b | SB-168 | 3 | TABLE 3 Mechanical Properties for Plate, Sheet, and  | 286 | 5 | 75 | 69 | linea unica que no cuadra con 5 columnas |
| bpvc_ii_b | SB-221 | 3 | TABLE 2 | 395-396 | 11 | 101 | 67 | linea unica que no cuadra con 11 columnas |
| bpvc_ii_a_2 | SA-524/SA-524M | 6 | SA-524/SA-524M ASME BPVC.II.A-2025 | 197 | 7 | 67 | 66 | linea unica que no cuadra con 7 columnas |
| bpvc_ii_c | SFA-5.5/SFA-5.5M | 18 | Table A1 Comparison of Classifications | 259-260 | 1 | 66 | 66 | fila entera en un solo Line, sin geometria ni fichas |
| bpvc_ii_a_1 | SA-182/SA-182M | 3 | TABLE 2 Chemical RequirementsA | 270 | 7 | 78 | 65 | linea unica que no cuadra con 7 columnas |
| bpvc_ii_a_2 | SA-666 | 3 | TABLE 2 TENSILE PROPERTY REQUIREMENTS (A) | 390-391 | 5 | 130 | 65 | linea unica que no cuadra con 5 columnas |
| bpvc_ii_a_1 | SA-213/SA-213M | 8 | (sin titulo) | 342-343 | 5 | 113 | 63 | linea unica que no cuadra con 5 columnas |
| bpvc_ii_c | SFA-5.23/SFA-5.23M | 5 | Table 4 Chemical Composition Requirements for Solid  | 812-813 | 9 | 77 | 63 | dos Line caen en la misma banda de columna |
| bpvc_ii_a_1 | SA-182/SA-182M | 1 | TABLE 1 Heat Treating Requirements | 268 | 6 | 76 | 61 | linea unica que no cuadra con 6 columnas |
| bpvc_ii_a_1 | SA-234/SA-234M | 1 | TABLE 1 Chemical Requirements | 383 | 3 | 70 | 60 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_a_1 | SA-6/SA-6M | 3 | (sin titulo) | 62 | 3 | 64 | 59 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_b | SB-163 | 3 | TABLE 3 Mechanical Properties of Tubes | 231 | 4 | 68 | 59 | linea unica que no cuadra con 4 columnas |
| bpvc_ii_c | SFA-5.14/SFA-5.14M | 2 | Table A.1 Comparison of Classificationsa | 536-537 | 4 | 64 | 59 | linea unica que no cuadra con 4 columnas |
| bpvc_ii_c | SFA-5.8/SFA-5.8M | 10 | Table 9 Standard Forms and Sizes of Brazing Filler M | 323-324 | 2 | 74 | 59 | linea unica que no cuadra con 2 columnas |
| bpvc_ii_a_1 | SA-240/SA-240M | 2 | TABLE 1 | 396-397 | 15 | 206 | 58 | linea unica que no cuadra con 15 columnas |
| bpvc_ii_a_1 | SA-53/SA-53M | 8 | TABLE X2.3 Dimensions, Weights (Masses) per Unit Len | 208 | 9 | 61 | 57 | dos Line caen en la misma banda de columna |
| bpvc_ii_a_2 | SA-479/SA-479M | 3 | TABLE 2 Mechanical Property Requirements | 85 | 8 | 77 | 57 | linea unica que no cuadra con 8 columnas |
| bpvc_ii_a_2 | SA-693 | 6 | TABLE 5 | 448-449 | 3 | 113 | 57 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_a_1 | SA-182/SA-182M | 7 | TABLE 4 Repair Welding Requirements | 278 | 3 | 76 | 56 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_a_1 | SA-370 | 14 | TABLE 6 | 654 | 16 | 61 | 55 | linea unica que no cuadra con 16 columnas |
| bpvc_ii_c | SFA-5.5/SFA-5.5M | 3 | Table 3 Tension Test Requirementsa,b | 225-226 | 2 | 98 | 55 | linea unica que no cuadra con 2 columnas |
| bpvc_ii_b | SB-622 | 3 | (sin titulo) | 980 | 4 | 62 | 54 | linea unica que no cuadra con 4 columnas |
| bpvc_ii_a_2 | SA-988/SA-988M | 3 | TABLE 2 Heat Treating Requirements | 767 | 4 | 58 | 53 | linea unica que no cuadra con 4 columnas |
| bpvc_ii_a_1 | SA-182/SA-182M | 2 | TABLE 1 | 269 | 4 | 58 | 52 | linea unica que no cuadra con 4 columnas |
| bpvc_ii_a_2 | SA-1058 | 3 | (sin titulo) | 851 | 9 | 54 | 52 | linea unica que no cuadra con 9 columnas |
| bpvc_ii_b | SB-108/SB-108M | 2 | TABLE 2 Tensile RequirementsA (Inch-Pound Units) | 115 | 6 | 54 | 52 | linea unica que no cuadra con 6 columnas |
| bpvc_ii_c | SFA-5.1/SFA-5.1M | 9 | Table 8 Requirements for Preparation of Fillet Weld  | 103-104 | 8 | 81 | 52 | linea unica que no cuadra con 8 columnas |
| bpvc_ii_c | SFA-5.4/SFA-5.4M | 9 | Table 6 All-Weld-Metal Mechanical Property Requireme | 182 | 1 | 52 | 52 | fila entera en un solo Line, sin geometria ni fichas |
| bpvc_ii_a_1 | SA-358/SA-358M | 1 | TABLE 1 Plate and Filler Metal Specifications | 620 | 16 | 79 | 51 | linea unica que no cuadra con 16 columnas |
| bpvc_ii_a_2 | SA-524/SA-524M | 5 | TABLE X1.2 Dimensions, Weights and Test Pressures fo | 196 | 7 | 74 | 51 | linea unica que no cuadra con 7 columnas |
| bpvc_ii_a_1 | SA-276 | 2 | TABLE 1 | 463 | 12 | 64 | 50 | linea unica que no cuadra con 12 columnas |
| bpvc_ii_b | SB-168 | 4 | TABLE 3 | 287 | 5 | 53 | 50 | linea unica que no cuadra con 5 columnas |
| bpvc_ii_a_1 | SA-193/SA-193M | 9 | TABLE 5 Marking of Austenitic Steels | 297 | 2 | 54 | 49 | linea unica que no cuadra con 2 columnas |
| bpvc_ii_a_1 | SA-20/SA-20M | 38 | TABLE X4.1 Group Designations for Cold Bending | 153 | 2 | 49 | 49 | linea unica que no cuadra con 2 columnas |
| bpvc_ii_a_1 | SA-370 | 8 | TABLE 2 Approximate Hardness Conversion Numbers for  | 649 | 9 | 58 | 49 | linea unica que no cuadra con 9 columnas |
| bpvc_ii_a_2 | SA-1058 | 4 | TABLE 4 Approximate Hardness Conversion Numbers for  | 852 | 9 | 83 | 49 | linea unica que no cuadra con 9 columnas |
| bpvc_ii_a_2 | SA-814/SA-814M | 1 | TABLE 1 Pipe DimensionsA | 630 | 3 | 53 | 49 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_b | SB-249/SB-249M | 17 | TABLE X1.1 Densities of Coppers and Copper Alloys | 502 | 3 | 69 | 49 | linea unica que no cuadra con 3 columnas |
| bpvc_ii_a_2 | SA-487/SA-487M | 4 | TABLE 3 Required Mechanical Properties | 137 | 7 | 48 | 48 | linea unica que no cuadra con 7 columnas |
| bpvc_ii_a_2 | SA-988/SA-988M | 2 | TABLE 1 | 764-765 | 16 | 73 | 48 | linea unica que no cuadra con 16 columnas |
| bpvc_ii_b | SB-42 | 6 | TABLE 3 Standard Dimensions, Weights, and Tolerances | 70 | 6 | 55 | 48 | linea unica que no cuadra con 6 columnas |
| bpvc_ii_a_2 | SA-1008/SA-1008M | 7 | TABLE X1.1 Suggested Minimum Inside Radius for Cold  | 812 | 2 | 49 | 47 | linea unica que no cuadra con 2 columnas |
| bpvc_ii_a_2 | SA-484/SA-484M | 1 | TABLE 1 Product Analysis Tolerances | 122 | 5 | 51 | 47 | linea unica que no cuadra con 5 columnas |
| bpvc_ii_b | SB-221 | 4 | (sin titulo) | 398-399 | 12 | 87 | 47 | linea unica que no cuadra con 12 columnas |
| bpvc_ii_a_1 | SA-358/SA-358M | 2 | TABLE 1 | 621 | 16 | 71 | 46 | linea unica que no cuadra con 16 columnas |
| bpvc_ii_a_2 | SA-479/SA-479M | 1 | TABLE 1 Chemical Requirements | 83 | 11 | 75 | 46 | linea unica que no cuadra con 11 columnas |
| bpvc_ii_a_2 | SA-988/SA-988M | 1 | TABLE 1 Chemical Requirements | 763 | 14 | 59 | 46 | dos Line caen en la misma banda de columna |

## Motivos de ambiguedad, agregados

| Motivo | Filas |
|---|---:|
| continuacion del rotulo de la fila anterior; se deja en la columna 0, sin fusionar, para no reordenar el texto impreso | 4434 |
| linea unica que no cuadra con 2 columnas | 4116 |
| fila entera en un solo Line, sin geometria ni fichas de valor que la separen | 3522 |
| linea unica que no cuadra con 3 columnas | 2915 |
| dos Line caen en la misma banda de columna | 2364 |
| linea unica que no cuadra con 4 columnas | 2305 |
| linea unica que no cuadra con 5 columnas | 1892 |
| linea unica que no cuadra con 6 columnas | 1518 |
| varios Line en la fila y ninguna banda de columna que los separe | 1188 |
| linea unica que no cuadra con 7 columnas | 961 |
| linea unica que no cuadra con 9 columnas | 693 |
| linea unica que no cuadra con 8 columnas | 485 |
| linea unica que no cuadra con 12 columnas | 423 |
| linea unica que no cuadra con 17 columnas | 371 |
| tabla de una sola columna | 331 |
| linea unica que no cuadra con 15 columnas | 293 |
| linea unica que no cuadra con 11 columnas | 290 |
| linea unica que no cuadra con 16 columnas | 285 |
| linea unica que no cuadra con 10 columnas | 235 |
| linea unica que no cuadra con 13 columnas | 177 |
| linea unica que no cuadra con 20 columnas | 162 |
| linea unica que no cuadra con 14 columnas | 126 |
| linea unica que no cuadra con 18 columnas | 118 |
| el reparto por bandas alteraria el orden de lectura de la fila | 92 |
| linea unica que no cuadra con 21 columnas | 60 |
| linea unica que no cuadra con 19 columnas | 55 |
| linea unica que no cuadra con 25 columnas | 48 |
| linea unica que no cuadra con 23 columnas | 23 |
| el bbox de la linea no abarca las columnas que el conteo le asigna | 23 |
| linea unica que no cuadra con 28 columnas | 17 |
| linea unica que no cuadra con 24 columnas | 12 |
| linea unica que no cuadra con 48 columnas | 11 |
| linea unica que no cuadra con 22 columnas | 10 |
| linea unica que no cuadra con 32 columnas | 9 |
| linea unica que no cuadra con 43 columnas | 8 |
| linea unica que no cuadra con 26 columnas | 5 |

## Punto de decision — Fase 2 del PLAN-SECII-ABC-001

El plan fija aqui una parada: *«si la fraccion AMBIGUA no es marginal, se corrige el algoritmo antes de tocar el libro. Escribir 124 000 filas dudosas es peor que no escribirlas.»* La fraccion medida es **38.7 % de las filas de dato**.

Lo que SI esta garantizado, y es lo que se puede garantizar sin el PDF:

- Ninguna fila pierde texto. La comprobacion sin perdida pasa sobre las 54198 filas y los 124115 bloques `Line`.
- Ninguna fila AMBIGUA se rellena por interpolacion sobre el `bbox`: se conserva entera en una celda y se declara. Sin `Span`, la posicion de cada palabra dentro de un `Line` no esta en el fichero, y repartirla con una fuente proporcional seria inventar estructura.
- Los huecos del origen se cuentan y se nombran; ninguno se rellena.

Lo que NO conviene hacer todavia: volcar las nueve hojas al libro. Una fila AMBIGUA en la hoja se lee como una fila de tabla que no esta tabulada, y a esa escala el volcado seria mas ruido que dato. La decision de seguir a las Fases 3 a 5 —o de acotarlas a las tablas que si se reparten— es del ingeniero, y este informe es el insumo para tomarla.
