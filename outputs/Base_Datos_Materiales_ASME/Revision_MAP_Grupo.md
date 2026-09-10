# Revision de MAP_Grupo — grupos de propiedades sin resolver por el codigo

Generado por `build_db_materiales.py` el 2026-09-10.
Fuente: Notas de las Tablas TM-1 y TE-1 de ASME BPVC II-D (Metrica) 2025,
extraidas a `resources/` por `extraer_notas_ii_d.py`.

## Que hay que decidir aqui

Las filas que **no** aparecen en este documento ya estan resueltas contra el
codigo y citan la nota que las sostiene: no requieren criterio de ingenieria.
Lo que sigue es lo que el codigo no resuelve por si solo.

| Estado | Filas de material |
|---|---:|
| AUTO (UNS exacto) | 1292 |
| AUTO (composicion en Nota) | 1297 |
| VALIDADO POR INGENIERO | 46 |
| REVISAR (regla textual del codigo) | 0 |
| AUTO (composicion via UNS en otra tabla) | 228 |
| SIN MAPEO | 591 |

Casos distintos sin resolver por el codigo: **106** (sobre 1026 filas de material).

- **0 ABIERTOS** — admiten criterio de ingenieria: el codigo dice algo sobre ese material (una regla redactada, una columna condicionada, o dos composiciones reales entre las que elegir) y una persona puede resolverlo con el codigo delante.
- **106 CERRADOS** — limite de la fuente: el material esta identificado sin ambiguedad y II-D sencillamente no tabula ni modulo E ni dilatacion para el. **No hay nada que firmar aqui.** Rellenar una casilla seria inventar un valor que el codigo no publica, que es justo lo que la Regla n.º 1 del proyecto prohibe. Se listan para que conste que se miraron y por que vias.

Aparte, **156 filas no imprimen composicion nominal** y su UNS no figura en TM-1..TM-5.
No entran en esta revision porque no hay dato de entrada que juzgar: para
asignarles grupo habria que identificar el material por otra via (la
especificacion y el grado en la tabla de origen).

Las filas se cuentan sobre las DOS ediciones (metrica y U.S. Customary):
un mismo caso afecta al material en ambas, porque la pertenencia a grupo no
depende del sistema de unidades.

## 1. Casos ABIERTOS — admiten criterio de ingenieria

*Ninguno.* Todo lo que el codigo permitia decidir esta decidido y registrado en `decisiones_map_grupo.json`; el resto es limite de la fuente y esta cerrado en el bloque siguiente.

## 2. Casos CERRADOS — limite de la fuente, no requieren firma

Para cada uno se agotaron las dos vias literales que el codigo publica —la
lista de miembros de una Nota de TM-1/TE-1 y el titulo de una columna
nombrada de TE-1— en LAS DOS ediciones, y la identificacion del material se
intento ademas por UNS y por (UNS, especificacion impresa en la propia fila)
contra las cuatro tablas indexadas del libro (II-D 1A y 1B/3, y Apendice A
del B31.3, cada una en sus dos ediciones). El resultado es el correcto: la
fila queda BLOQUEADA para modulo E y dilatacion.

| # | Material | Filas | Estado | Por que esta cerrado |
|---:|---|---:|---|---|
| 1 | `23Cr-25Ni-5.5Mo-N` | 42 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-403, SA-182, +8 mas |
| 2 | `19Cr-15Ni-4Mo` | 38 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +6 mas |
| 3 | `25Cr-22Ni-2Mo-N` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-312, +2 mas |
| 4 | `42Fe-33Ni-21Cr` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-276, SA-479, SA-182, +6 mas |
| 5 | `60Ni-25Cr-9.5Fe-2.1Al` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-462, SB-564, +6 mas |
| 6 | `23Cr-12Ni-Cb` | 32 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |
| 7 | `25Cr-20Ni-Cb` | 32 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |
| 8 | `37Ni-33Fe-25Cr` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-163, SB-564, +5 mas |
| 9 | `47Ni-22Cr-20Fe-7Mo` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-581, SB-622, +3 mas |
| 10 | `58Ni-33Cr-8Mo` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |
| 11 | `59Ni-23Cr-16Mo-1.6Cu` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |
| 12 | `62Ni-22Mo-15Cr` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |
| 13 | `60Ni-19Cr-19Mo-1.8Ta` | 26 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-564, SB-575, SB-574, +4 mas |
| 14 | `21Cr-5Mn-1.5Ni-Cu-N` | 24 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240, +2 mas |
| 15 | `33Ni-42Fe-21Cr` | 24 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-564, SB-409, +3 mas |
| 16 | `21Ni-30Fe-22Cr-18Co-3Mo-3W` | 22 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-572, SB-435, SB-622, +2 mas |
| 17 | `26Ni-43Fe-22Cr-5Mo` | 22 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-621, SB-620, SB-622, +2 mas |
| 18 | `18Cr-3Ni-12Mn` | 20 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-240, SA-312, +2 mas |
| 19 | `49Ni-25Cr-18Fe-6Mo` | 18 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-622, SB-619, +2 mas |
| 20 | `C-Mn-Si-V-Cb` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-656 |
| 21 | `G41400` (UNS) | 16 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «G41400» trae varias composiciones en el libro, pero el grado que la propia fila imprime («B7») deja una sola: «Cr-Mo». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-193, SA-320, SA-574 |
| 22 | `60Ni-23Cr-Fe` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-168, SB-167, +1 mas |
| 23 | `20Cr-3Ni-1.5Mo-N` | 14 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-789, SA-240, SA-790 |
| 24 | `17Cr-4Ni-6Mn` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-666 |
| 25 | `27Ni-22Cr-7Mo-Mn-Cu-N` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-249 |
| 26 | `18Cr-12Ni-2Mo` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351, SA-451 |
| 27 | `K14073` (UNS) | 12 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-540 |
| 28 | `1Cr-1Mn-1/4Mo` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |
| 29 | `21/4Cr-1Mo-V` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-182, SA-336, SA-541, +2 mas |
| 30 | `3Ni-13/4Cr-1/2Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543, SA-372 |
| 31 | `26Cr-3Ni-3Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240, SA-803 |
| 32 | `26Cr-4Ni-Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |
| 33 | `26Cr-4Ni-Mo-N` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |
| 34 | `27Cr-1Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-182, +1 mas |
| 35 | `K21903` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K21903» aparece como «21∕4Ni». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-333, SA-334 |
| 36 | `S43035` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-268, SA-479, SA-803 |
| 37 | `29Cr-4Mo` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |
| 38 | `29Cr-4Mo-2Ni` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |
| 39 | `35Ni-30Fe-24Cr-6Mo-Cu` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |
| 40 | `37Ni-33Fe-23Cr-4Mo-Cu` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |
| 41 | `K13047` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |
| 42 | `G41350` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |
| 43 | `G41370` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «G41370» aparece como «Cr-Mo». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-574, SA-372 |
| 44 | `K13548` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |
| 45 | `K12521` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-533 |
| 46 | `K12023` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-250, SA-209 |
| 47 | `K11422` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-250, SA-209 |
| 48 | `23/4Ni-11/2Cr-1/2Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543 |
| 49 | `18Cr-2Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |
| 50 | `27Cr-1Mo-Ti` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |
| 51 | `J13002` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-487 |
| 52 | `K14072` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K14072» aparece como «Cr-Mo-V». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-193 |
| 53 | `G40370` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «G40370» aparece como «Cr-Mo». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-574, SA-320 |
| 54 | `K13050` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-350 |
| 55 | `K61365` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-553 |
| 56 | `19Cr-9Ni-2Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |
| 57 | `S40300` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-479 |
| 58 | `S41003` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S41003» trae varias composiciones en el libro, pero la especificacion que la propia fila imprime («1010») deja una sola: «12Cr-1Ni». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-1010 |
| 59 | `S30100` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |
| 60 | `S40800` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-268 |
| 61 | `S43036` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S43036» aparece como «18Cr-Ti». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-268 |
| 62 | `16Cr-12Ni-2Mo-Cb` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |
| 63 | `16Cr-4Ni-6Mn` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |
| 64 | `16Cr-9Mn-2Ni-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |
| 65 | `18Cr-9Ni-3Cu-Cb-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |
| 66 | `20.5Cr-8.8Ni-Mo-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |
| 67 | `25Cr-20Ni-Cb-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |
| 68 | `25Cr-4Ni-4Mo-Ti` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-268 |
| 69 | `29Cr-4Mo-Ti` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268 |
| 70 | `19Cr-10Ni-3Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |
| 71 | `19Cr-9Ni-1/2Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |
| 72 | `Mn-1/4Mo-V` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-487 |
| 73 | `G40420` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |
| 74 | `G41420` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |
| 75 | `G41450` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |
| 76 | `Co-26Cr-9Ni-5Mo-3Fe-2W` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-818, SB-815 |
| 77 | `S30300` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S30300» aparece como «18Cr-9Ni». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-320 |
| 78 | `12Cr-1Mo-V-W` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-437 |
| 79 | `67Ni-28Cu-3Al` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SF-468 |
| 80 | `A02040` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SB-108, SB-26 |
| 81 | `K12520` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-336 |
| 82 | `K22036` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-350 |
| 83 | `K14508` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |
| 84 | `K32026` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-765 |
| 85 | `K21703` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K21703» aparece como «21∕4Ni». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-203 |
| 86 | `K22103` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K22103» aparece como «21∕4Ni». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-203 |
| 87 | `K11224` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-562 |
| 88 | `K12047` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-213 |
| 89 | `11/2Si-1/2Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-335 |
| 90 | `13/4Ni-3/4Cr-Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-372 |
| 91 | `S41500` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-182 |
| 92 | `S40910` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |
| 93 | `S40920` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |
| 94 | `S40930` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |
| 95 | `S43932` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |
| 96 | `20Cr-10Ni` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479 |
| 97 | `J31545` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «J31545» aparece como «3Cr-Mo». El material queda identificado; lo que falta es el dato: esa composicion no figura en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para ella. <br>Especificaciones: SA-426 |
| 98 | `17Cr-4Ni-3Cu` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-747 |
| 99 | `K50100` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-193 |
| 100 | `2Ni-3/4Cr-1/3Mo-V` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |
| 101 | `32Ni-45Fe-20Cr-Cb` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |
| 102 | `53Ni-17Mo-16Cr-6Fe-5W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |
| 103 | `59Ni-22Cr-14Mo-4Fe-3W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |
| 104 | `62Ni-28Mo-5Fe` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |
| 105 | `R61702` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SB-752 |
| 106 | `95.5Zr + 2.5Nb` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1 ni en ninguna columna nombrada de TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-752 |

## Como usar este documento

1. El bloque 2 no se rellena. Esta cerrado: la unica forma de darle grupo a
   esos materiales seria que ASME publicase el dato, que hoy no publica.
2. En el bloque 1, decida el grupo de TM-1 (modulo E) y/o TE-1 (dilatacion)
   que corresponde, anotelo y firme.
3. Mientras una fila siga sin grupo, el motor deja el calculo BLOQUEADO para
   esos materiales. Es el comportamiento correcto: el codigo prohibe
   extrapolar y prohibe inventar la pertenencia a un grupo.


## Vuelta al motor

Para que estas decisiones lleguen al calculo, copie
`decisiones_map_grupo.plantilla.json` a `decisiones_map_grupo.json`,
rellene `grupo_tm` / `grupo_te` con el rotulo tal como lo imprime el codigo
(«Material Group E», «Group 1») y firme cada entrada. El builder las lee en
la siguiente corrida y esas filas pasan al estado VALIDADO POR INGENIERO,
**siempre separado de las filas AUTO**: quien audite el libro tiene que poder
distinguir lo que dice el codigo de lo que decidio una persona.

Una decision no puede pisar al codigo: si contradice un grupo que el codigo
si asigna, no se aplica y el choque se reporta en las limitaciones del libro.
