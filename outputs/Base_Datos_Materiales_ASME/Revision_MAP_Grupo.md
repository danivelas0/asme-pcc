# Revision de MAP_Grupo — grupos de propiedades sin resolver por el codigo

Generado por `build_db_materiales.py` el 2026-09-08.
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
| VALIDADO POR INGENIERO | 44 |
| REVISAR (regla textual del codigo) | 0 |
| AUTO (composicion via UNS en otra tabla) | 193 |
| SIN MAPEO | 628 |

Decisiones distintas a tomar: **113** (sobre 1100 filas de material).

Aparte, **156 filas no imprimen composicion nominal** y su UNS no figura en TM-1..TM-5.
No entran en esta revision porque no hay dato de entrada que juzgar: para
asignarles grupo habria que identificar el material por otra via (la
especificacion y el grado en la tabla de origen).

## Decisiones

Las filas se cuentan sobre las DOS ediciones (metrica y U.S. Customary):
una misma decision desbloquea el material en ambas, porque la pertenencia a
grupo no depende del sistema de unidades.

| # | Se decide sobre | Filas | Estado | Que hay que decidir | Grupo asignado | Firma / fecha |
|---:|---|---:|---|---|---|---|
| 1 | `23Cr-25Ni-5.5Mo-N` | 42 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-403, SA-182, +8 mas |  |  |
| 2 | `K81340` (UNS) | 38 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K81340» aparece como «9Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-333, SA-334, SA-353, +3 mas |  |  |
| 3 | `19Cr-15Ni-4Mo` | 38 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +6 mas |  |  |
| 4 | `25Cr-22Ni-2Mo-N` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-312, +2 mas |  |  |
| 5 | `42Fe-33Ni-21Cr` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-276, SA-479, SA-182, +6 mas |  |  |
| 6 | `60Ni-25Cr-9.5Fe-2.1Al` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-462, SB-564, +6 mas |  |  |
| 7 | `23Cr-12Ni-Cb` | 32 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |  |  |
| 8 | `25Cr-20Ni-Cb` | 32 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |  |  |
| 9 | `37Ni-33Fe-25Cr` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-163, SB-564, +5 mas |  |  |
| 10 | `47Ni-22Cr-20Fe-7Mo` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-581, SB-622, +3 mas |  |  |
| 11 | `58Ni-33Cr-8Mo` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 12 | `59Ni-23Cr-16Mo-1.6Cu` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 13 | `62Ni-22Mo-15Cr` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 14 | `60Ni-19Cr-19Mo-1.8Ta` | 26 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-564, SB-575, SB-574, +4 mas |  |  |
| 15 | `21Cr-5Mn-1.5Ni-Cu-N` | 24 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240, +2 mas |  |  |
| 16 | `33Ni-42Fe-21Cr` | 24 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-564, SB-409, +3 mas |  |  |
| 17 | `21Ni-30Fe-22Cr-18Co-3Mo-3W` | 22 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-572, SB-435, SB-622, +2 mas |  |  |
| 18 | `26Ni-43Fe-22Cr-5Mo` | 22 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-621, SB-620, SB-622, +2 mas |  |  |
| 19 | `18Cr-3Ni-12Mn` | 20 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-240, SA-312, +2 mas |  |  |
| 20 | `S41000` (UNS) | 18 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «S41000» mas de una: 12Cr; 13Cr. No se elige por cuenta propia. <br>Especificaciones: SA-479, SA-182, SA-268, +2 mas |  |  |
| 21 | `49Ni-25Cr-18Fe-6Mo` | 18 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-622, SB-619, +2 mas |  |  |
| 22 | `C-Mn-Si-V-Cb` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-656 |  |  |
| 23 | `G41400` (UNS) | 16 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «G41400» mas de una: Cr-0.2Mo; Cr-Mo; Cr-1∕5Mo; Cr-Mo Stainless Steel. No se elige por cuenta propia. <br>Especificaciones: SA-193, SA-320, SA-574 |  |  |
| 24 | `60Ni-23Cr-Fe` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-168, SB-167, +1 mas |  |  |
| 25 | `20Cr-3Ni-1.5Mo-N` | 14 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-789, SA-240, SA-790 |  |  |
| 26 | `17Cr-4Ni-6Mn` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-666 |  |  |
| 27 | `27Ni-22Cr-7Mo-Mn-Cu-N` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-249 |  |  |
| 28 | `18Cr-12Ni-2Mo` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351, SA-451 |  |  |
| 29 | `K14073` (UNS) | 12 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-540 |  |  |
| 30 | `1Cr-1Mn-1/4Mo` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |  |  |
| 31 | `21/4Cr-1Mo-V` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-182, SA-336, SA-541, +2 mas |  |  |
| 32 | `3Ni-13/4Cr-1/2Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543, SA-372 |  |  |
| 33 | `26Cr-3Ni-3Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240, SA-803 |  |  |
| 34 | `26Cr-4Ni-Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |  |  |
| 35 | `26Cr-4Ni-Mo-N` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |  |  |
| 36 | `27Cr-1Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-182, +1 mas |  |  |
| 37 | `K21903` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K21903» aparece como «21∕4Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-333, SA-334 |  |  |
| 38 | `S43035` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-268, SA-479, SA-803 |  |  |
| 39 | `29Cr-4Mo` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |  |  |
| 40 | `29Cr-4Mo-2Ni` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |  |  |
| 41 | `35Ni-30Fe-24Cr-6Mo-Cu` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |  |  |
| 42 | `37Ni-33Fe-23Cr-4Mo-Cu` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |  |  |
| 43 | `K13047` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 44 | `G41350` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 45 | `G41370` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «G41370» aparece como «Cr-Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-574, SA-372 |  |  |
| 46 | `K13548` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 47 | `K71340` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K71340» aparece como «8Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-553, SA-522 |  |  |
| 48 | `K12521` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-533 |  |  |
| 49 | `K12023` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-250, SA-209 |  |  |
| 50 | `K11422` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-250, SA-209 |  |  |
| 51 | `23/4Ni-11/2Cr-1/2Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543 |  |  |
| 52 | `18Cr-2Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |  |  |
| 53 | `27Cr-1Mo-Ti` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |  |  |
| 54 | `J13002` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-487 |  |  |
| 55 | `K14072` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K14072» aparece como «Cr-Mo-V», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-193 |  |  |
| 56 | `G40370` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «G40370» aparece como «Cr-Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-574, SA-320 |  |  |
| 57 | `K13050` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-350 |  |  |
| 58 | `K61365` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-553 |  |  |
| 59 | `19Cr-9Ni-2Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 60 | `S40300` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-479 |  |  |
| 61 | `S41003` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «S41003» mas de una: 12Cr; 12Cr-1Ni. No se elige por cuenta propia. <br>Especificaciones: SA-1010 |  |  |
| 62 | `S30100` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 63 | `S40800` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-268 |  |  |
| 64 | `S43036` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S43036» aparece como «18Cr-Ti», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-268 |  |  |
| 65 | `16Cr-12Ni-2Mo-Cb` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 66 | `16Cr-4Ni-6Mn` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 67 | `16Cr-9Mn-2Ni-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 68 | `18Cr-9Ni-3Cu-Cb-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |  |  |
| 69 | `20.5Cr-8.8Ni-Mo-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 70 | `25Cr-20Ni-Cb-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |  |  |
| 71 | `25Cr-4Ni-4Mo-Ti` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-268 |  |  |
| 72 | `29Cr-4Mo-Ti` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268 |  |  |
| 73 | `J82090` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «J82090» aparece como «9Cr-1Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-426, SA-217 |  |  |
| 74 | `J91150` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «J91150» mas de una: 12Cr; 13Cr. No se elige por cuenta propia. <br>Especificaciones: SA-426, SA-217 |  |  |
| 75 | `19Cr-10Ni-3Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 76 | `19Cr-9Ni-1/2Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 77 | `Mn-1/4Mo-V` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-487 |  |  |
| 78 | `G40420` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |  |  |
| 79 | `G41420` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |  |  |
| 80 | `G41450` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |  |  |
| 81 | `Co-26Cr-9Ni-5Mo-3Fe-2W` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-818, SB-815 |  |  |
| 82 | `S30300` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S30300» aparece como «18Cr-9Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-320 |  |  |
| 83 | `12Cr-1Mo-V-W` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-437 |  |  |
| 84 | `67Ni-28Cu-3Al` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SF-468 |  |  |
| 85 | `A02040` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SB-108, SB-26 |  |  |
| 86 | `K12520` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-336 |  |  |
| 87 | `K22036` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-350 |  |  |
| 88 | `K14508` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 89 | `K32026` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-765 |  |  |
| 90 | `K21703` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K21703» aparece como «21∕4Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-203 |  |  |
| 91 | `K22103` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K22103» aparece como «21∕4Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-203 |  |  |
| 92 | `K41583` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K41583» aparece como «5Ni-1∕4Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-645 |  |  |
| 93 | `K11224` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-562 |  |  |
| 94 | `K12047` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-213 |  |  |
| 95 | `11/2Si-1/2Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-335 |  |  |
| 96 | `13/4Ni-3/4Cr-Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-372 |  |  |
| 97 | `S41500` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-182 |  |  |
| 98 | `S40910` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 99 | `S40920` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 100 | `S40930` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 101 | `S43932` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 102 | `S44600` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S44600» aparece como «27Cr», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-268 |  |  |
| 103 | `20Cr-10Ni` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479 |  |  |
| 104 | `J31545` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «J31545» aparece como «3Cr-Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-426 |  |  |
| 105 | `17Cr-4Ni-3Cu` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-747 |  |  |
| 106 | `K50100` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-193 |  |  |
| 107 | `2Ni-3/4Cr-1/3Mo-V` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |  |  |
| 108 | `32Ni-45Fe-20Cr-Cb` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 109 | `53Ni-17Mo-16Cr-6Fe-5W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 110 | `59Ni-22Cr-14Mo-4Fe-3W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 111 | `62Ni-28Mo-5Fe` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 112 | `R61702` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SB-752 |  |  |
| 113 | `95.5Zr + 2.5Nb` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-752 |  |  |

## Como usar este documento

1. Para cada fila, decida el grupo de TM-1 (modulo E) y/o TE-1 (dilatacion)
   que corresponde, o confirme que II-D no publica el dato para ese material.
2. Anote el grupo en la columna correspondiente y firme.
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
