# Revision de MAP_Grupo — grupos de propiedades sin resolver por el codigo

Generado por `build_db_materiales.py` el 2026-09-06.
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
| REVISAR (regla textual del codigo) | 17 |
| SIN MAPEO | 848 |

Decisiones distintas a tomar: **65** (sobre 438 filas de material).

Aparte, **427 filas no imprimen composicion nominal** y su UNS no figura en TM-1..TM-5.
No entran en esta revision porque no hay dato de entrada que juzgar: para
asignarles grupo habria que identificar el material por otra via (la
especificacion y el grado en la tabla de origen).

## Decisiones

| # | Composicion nominal | Filas | Estado | Que hay que decidir | Grupo asignado | Firma / fecha |
|---:|---|---:|---|---|---|---|
| 1 | `9Cr-1Mo-V` | 17 | REVISAR (regla textual del codigo) | TM-1 Nota (5) lista «9Cr–Mo, including variations thereof». Aplicar la regla y confirmar si este material queda dentro de Material Group E. <br>Especificaciones: SA-182, SA-387, SA-335, +4 mas |  |  |
| 2 | `23Cr-25Ni-5.5Mo-N` | 21 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-403, SA-182, +8 mas |  |  |
| 3 | `19Cr-15Ni-4Mo` | 19 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +6 mas |  |  |
| 4 | `25Cr-22Ni-2Mo-N` | 18 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-312, +2 mas |  |  |
| 5 | `42Fe-33Ni-21Cr` | 18 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-276, SA-479, SA-182, +6 mas |  |  |
| 6 | `60Ni-25Cr-9.5Fe-2.1Al` | 18 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-462, SB-564, +6 mas |  |  |
| 7 | `23Cr-12Ni-Cb` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |  |  |
| 8 | `25Cr-20Ni-Cb` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |  |  |
| 9 | `37Ni-33Fe-25Cr` | 15 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-163, SB-564, +5 mas |  |  |
| 10 | `47Ni-22Cr-20Fe-7Mo` | 15 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-581, SB-622, +3 mas |  |  |
| 11 | `58Ni-33Cr-8Mo` | 15 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 12 | `59Ni-23Cr-16Mo-1.6Cu` | 15 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 13 | `62Ni-22Mo-15Cr` | 15 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 14 | `60Ni-19Cr-19Mo-1.8Ta` | 13 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-564, SB-575, SB-574, +4 mas |  |  |
| 15 | `21Cr-5Mn-1.5Ni-Cu-N` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240, +2 mas |  |  |
| 16 | `33Ni-42Fe-21Cr` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-564, SB-409, +3 mas |  |  |
| 17 | `21Ni-30Fe-22Cr-18Co-3Mo-3W` | 11 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-572, SB-435, SB-622, +2 mas |  |  |
| 18 | `26Ni-43Fe-22Cr-5Mo` | 11 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-621, SB-620, SB-622, +2 mas |  |  |
| 19 | `18Cr-3Ni-12Mn` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-240, SA-312, +2 mas |  |  |
| 20 | `49Ni-25Cr-18Fe-6Mo` | 9 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-622, SB-619, +2 mas |  |  |
| 21 | `C-Mn-Si-V-Cb` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-656 |  |  |
| 22 | `60Ni-23Cr-Fe` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-168, SB-167, +1 mas |  |  |
| 23 | `20Cr-3Ni-1.5Mo-N` | 7 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-789, SA-240, SA-790 |  |  |
| 24 | `17Cr-4Ni-6Mn` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-666 |  |  |
| 25 | `27Ni-22Cr-7Mo-Mn-Cu-N` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-249 |  |  |
| 26 | `18Cr-12Ni-2Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351, SA-451 |  |  |
| 27 | `1Cr-1Mn-1/4Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |  |  |
| 28 | `21/4Cr-1Mo-V` | 5 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-182, SA-336, SA-541, +2 mas |  |  |
| 29 | `3Ni-13/4Cr-1/2Mo` | 5 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543, SA-372 |  |  |
| 30 | `26Cr-3Ni-3Mo` | 5 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240, SA-803 |  |  |
| 31 | `26Cr-4Ni-Mo` | 5 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |  |  |
| 32 | `26Cr-4Ni-Mo-N` | 5 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |  |  |
| 33 | `27Cr-1Mo` | 5 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-182, +1 mas |  |  |
| 34 | `29Cr-4Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |  |  |
| 35 | `29Cr-4Mo-2Ni` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |  |  |
| 36 | `35Ni-30Fe-24Cr-6Mo-Cu` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |  |  |
| 37 | `37Ni-33Fe-23Cr-4Mo-Cu` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |  |  |
| 38 | `23/4Ni-11/2Cr-1/2Mo` | 3 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543 |  |  |
| 39 | `18Cr-2Mo` | 3 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |  |  |
| 40 | `27Cr-1Mo-Ti` | 3 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |  |  |
| 41 | `19Cr-9Ni-2Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 42 | `16Cr-12Ni-2Mo-Cb` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 43 | `16Cr-4Ni-6Mn` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 44 | `16Cr-9Mn-2Ni-N` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 45 | `18Cr-9Ni-3Cu-Cb-N` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |  |  |
| 46 | `20.5Cr-8.8Ni-Mo-N` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 47 | `25Cr-20Ni-Cb-N` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |  |  |
| 48 | `25Cr-4Ni-4Mo-Ti` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-268 |  |  |
| 49 | `29Cr-4Mo-Ti` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268 |  |  |
| 50 | `19Cr-10Ni-3Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 51 | `19Cr-9Ni-1/2Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 52 | `Mn-1/4Mo-V` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-487 |  |  |
| 53 | `Co-26Cr-9Ni-5Mo-3Fe-2W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-818, SB-815 |  |  |
| 54 | `12Cr-1Mo-V-W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-437 |  |  |
| 55 | `67Ni-28Cu-3Al` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SF-468 |  |  |
| 56 | `11/2Si-1/2Mo` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-335 |  |  |
| 57 | `13/4Ni-3/4Cr-Mo` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-372 |  |  |
| 58 | `20Cr-10Ni` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479 |  |  |
| 59 | `17Cr-4Ni-3Cu` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-747 |  |  |
| 60 | `2Ni-3/4Cr-1/3Mo-V` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |  |  |
| 61 | `32Ni-45Fe-20Cr-Cb` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 62 | `53Ni-17Mo-16Cr-6Fe-5W` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 63 | `59Ni-22Cr-14Mo-4Fe-3W` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 64 | `62Ni-28Mo-5Fe` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 65 | `95.5Zr + 2.5Nb` | 1 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-752 |  |  |

## Como usar este documento

1. Para cada fila, decida el grupo de TM-1 (modulo E) y/o TE-1 (dilatacion)
   que corresponde, o confirme que II-D no publica el dato para ese material.
2. Anote el grupo en la columna correspondiente y firme.
3. Mientras una fila siga sin grupo, el motor deja el calculo BLOQUEADO para
   esos materiales. Es el comportamiento correcto: el codigo prohibe
   extrapolar y prohibe inventar la pertenencia a un grupo.
