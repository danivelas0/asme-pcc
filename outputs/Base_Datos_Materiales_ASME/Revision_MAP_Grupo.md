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
| VALIDADO POR INGENIERO | 0 |
| REVISAR (regla textual del codigo) | 17 |
| REVISAR (composicion de otra tabla) | 193 |
| SIN MAPEO | 655 |

Decisiones distintas a tomar: **157** (sobre 1574 filas de material).

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
| 1 | `9Cr-1Mo-V` | 34 | REVISAR (regla textual del codigo) | TM-1 Nota (5) lista «9Cr–Mo, including variations thereof». Aplicar la regla y confirmar si este material queda dentro de Material Group E. <br>Especificaciones: SA-182, SA-387, SA-335, +4 mas |  |  |
| 2 | `S30400` (UNS) | 86 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S30400» aparece como «18Cr-8Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-320, SA-193, SA-182, +13 mas |  |  |
| 3 | `S30403` (UNS) | 54 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S30403» aparece como «18Cr-8Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-182, SA-312, SA-249, +10 mas |  |  |
| 4 | `S30409` (UNS) | 52 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S30409» aparece como «18Cr-8Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-182, SA-312, SA-249, +9 mas |  |  |
| 5 | `K31545` (UNS) | 16 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K31545» aparece como «3Cr-1Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-336, SA-387, SA-369, +3 mas |  |  |
| 6 | `K22035` (UNS) | 14 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K22035» aparece como «2Ni-1Cu» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-333, SA-234, SA-182, +2 mas |  |  |
| 7 | `J92600` (UNS) | 14 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J92600» aparece como «18Cr-8Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-351, SA-451 |  |  |
| 8 | `K11562` (UNS) | 12 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K11562» aparece como «1Cr-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-250, SA-369, SA-182, +2 mas |  |  |
| 9 | `J92500` (UNS) | 12 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J92500» aparece como «18Cr-8Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-351, SA-451 |  |  |
| 10 | `K11522` (UNS) | 10 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K11522» aparece como «C-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-250, SA-369, SA-335, +1 mas |  |  |
| 11 | `S40500` (UNS) | 10 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S40500» aparece como «12Cr-Al» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-240, SA-268, SA-479 |  |  |
| 12 | `S30500` (UNS) | 10 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S30500» aparece como «18Cr-11Ni» en DB_B31_3, DB_B31_3C, DB_BPVC_IID, DB_BPVC_IIDC, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-193 |  |  |
| 13 | `K31918` (UNS) | 8 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K31918» aparece como «31∕2Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-333, SA-334 |  |  |
| 14 | `K11757` (UNS) | 8 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K11757» aparece como «1Cr-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-387, SA-691 |  |  |
| 15 | `S30200` (UNS) | 8 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S30200» aparece como «18Cr-8Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-479, SA-240 |  |  |
| 16 | `S43000` (UNS) | 8 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S43000» aparece como «17Cr» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-268, SA-479, SA-240 |  |  |
| 17 | `S42900` (UNS) | 6 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S42900» aparece como «15Cr» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-268, SA-240 |  |  |
| 18 | `K11564` (UNS) | 4 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K11564» aparece como «1Cr-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-182, SA-336 |  |  |
| 19 | `K42544` (UNS) | 4 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K42544» aparece como «5Cr-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-182, SA-336 |  |  |
| 20 | `K32025` (UNS) | 4 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K32025» aparece como «31∕2Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-350 |  |  |
| 21 | `K32018` (UNS) | 4 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K32018» aparece como «31∕2Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-203 |  |  |
| 22 | `K12021` (UNS) | 4 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K12021» aparece como «Mn-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-302, SA-672 |  |  |
| 23 | `S40900` (UNS) | 4 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S40900» aparece como «11Cr-Ti» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-268 |  |  |
| 24 | `J42045` (UNS) | 4 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J42045» aparece como «5Cr-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-426, SA-217 |  |  |
| 25 | `K12822` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K12822» aparece como «C-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-182 |  |  |
| 26 | `K31718` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K31718» aparece como «31∕2Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-203 |  |  |
| 27 | `K12022` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K12022» aparece como «Mn-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-302 |  |  |
| 28 | `K12821` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K12821» aparece como «C-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-234 |  |  |
| 29 | `K12062` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «K12062» aparece como «1Cr-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-234 |  |  |
| 30 | `S41008` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S41008» aparece como «13Cr» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-240 |  |  |
| 31 | `S30453` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S30453» aparece como «18Cr-8Ni-N» en DB_BPVC_IID, DB_BPVC_IIDC, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-358 |  |  |
| 32 | `J12521` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J12521» aparece como «C-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-426 |  |  |
| 33 | `J11562` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J11562» aparece como «1Cr-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-426 |  |  |
| 34 | `J12524` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J12524» aparece como «C-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-217 |  |  |
| 35 | `J12522` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J12522» aparece como «C-1∕2Mo» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-352 |  |  |
| 36 | `J22500` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J22500» aparece como «21∕2Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-352 |  |  |
| 37 | `J31550` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J31550» aparece como «31∕2Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-352 |  |  |
| 38 | `J91540` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «J91540» aparece como «13Cr-4Ni» en DB_B31_3, DB_B31_3C, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-487 |  |  |
| 39 | `S30451` (UNS) | 2 | REVISAR (composicion de otra tabla) | La fila no imprime composicion nominal. Su UNS «S30451» aparece como «18Cr-8Ni-N» en DB_BPVC_IID, DB_BPVC_IIDC, y esa composicion si figura en Nota. Confirmar que es el mismo material antes de usar E o dilatacion. <br>Especificaciones: SA-193 |  |  |
| 40 | `23Cr-25Ni-5.5Mo-N` | 42 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-403, SA-182, +8 mas |  |  |
| 41 | `K81340` (UNS) | 38 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K81340» aparece como «9Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-333, SA-334, SA-353, +3 mas |  |  |
| 42 | `19Cr-15Ni-4Mo` | 38 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +6 mas |  |  |
| 43 | `25Cr-22Ni-2Mo-N` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-312, +2 mas |  |  |
| 44 | `42Fe-33Ni-21Cr` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-276, SA-479, SA-182, +6 mas |  |  |
| 45 | `60Ni-25Cr-9.5Fe-2.1Al` | 36 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-462, SB-564, +6 mas |  |  |
| 46 | `23Cr-12Ni-Cb` | 32 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |  |  |
| 47 | `25Cr-20Ni-Cb` | 32 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-312, SA-479, SA-240, +4 mas |  |  |
| 48 | `37Ni-33Fe-25Cr` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-163, SB-564, +5 mas |  |  |
| 49 | `47Ni-22Cr-20Fe-7Mo` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-581, SB-622, +3 mas |  |  |
| 50 | `58Ni-33Cr-8Mo` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 51 | `59Ni-23Cr-16Mo-1.6Cu` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 52 | `62Ni-22Mo-15Cr` | 30 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-462, SB-564, SB-575, +5 mas |  |  |
| 53 | `60Ni-19Cr-19Mo-1.8Ta` | 26 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-564, SB-575, SB-574, +4 mas |  |  |
| 54 | `21Cr-5Mn-1.5Ni-Cu-N` | 24 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240, +2 mas |  |  |
| 55 | `33Ni-42Fe-21Cr` | 24 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-408, SB-564, SB-409, +3 mas |  |  |
| 56 | `21Ni-30Fe-22Cr-18Co-3Mo-3W` | 22 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-572, SB-435, SB-622, +2 mas |  |  |
| 57 | `26Ni-43Fe-22Cr-5Mo` | 22 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-621, SB-620, SB-622, +2 mas |  |  |
| 58 | `18Cr-3Ni-12Mn` | 20 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479, SA-240, SA-312, +2 mas |  |  |
| 59 | `K41545` (UNS) | 18 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «K41545» mas de una: 5Cr-1∕2Mo; 5Cr-1∕2Mo A387 Gr. 5 Cl. 1. No se elige por cuenta propia. <br>Especificaciones: SA-387, SA-369, SA-182, +5 mas |  |  |
| 60 | `S41000` (UNS) | 18 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «S41000» mas de una: 12Cr; 13Cr. No se elige por cuenta propia. <br>Especificaciones: SA-479, SA-182, SA-268, +2 mas |  |  |
| 61 | `49Ni-25Cr-18Fe-6Mo` | 18 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-582, SB-622, SB-619, +2 mas |  |  |
| 62 | `C-Mn-Si-V-Cb` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-656 |  |  |
| 63 | `G41400` (UNS) | 16 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «G41400» mas de una: Cr-0.2Mo; Cr-Mo; Cr-1∕5Mo; Cr-Mo Stainless Steel. No se elige por cuenta propia. <br>Especificaciones: SA-193, SA-320, SA-574 |  |  |
| 64 | `60Ni-23Cr-Fe` | 16 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-166, SB-168, SB-167, +1 mas |  |  |
| 65 | `20Cr-3Ni-1.5Mo-N` | 14 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-789, SA-240, SA-790 |  |  |
| 66 | `K90941` (UNS) | 12 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «K90941» mas de una: 9Cr-1Mo; 9Cr-1Mo A387 Gr. 9 Cl. 1. No se elige por cuenta propia. <br>Especificaciones: SA-234, SA-369, SA-182, +3 mas |  |  |
| 67 | `17Cr-4Ni-6Mn` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-666 |  |  |
| 68 | `27Ni-22Cr-7Mo-Mn-Cu-N` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-213, SA-249 |  |  |
| 69 | `18Cr-12Ni-2Mo` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351, SA-451 |  |  |
| 70 | `K14073` (UNS) | 12 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-540 |  |  |
| 71 | `1Cr-1Mn-1/4Mo` | 12 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |  |  |
| 72 | `21/4Cr-1Mo-V` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-182, SA-336, SA-541, +2 mas |  |  |
| 73 | `3Ni-13/4Cr-1/2Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543, SA-372 |  |  |
| 74 | `26Cr-3Ni-3Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240, SA-803 |  |  |
| 75 | `26Cr-4Ni-Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |  |  |
| 76 | `26Cr-4Ni-Mo-N` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-790, SA-789, SA-240 |  |  |
| 77 | `27Cr-1Mo` | 10 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-182, +1 mas |  |  |
| 78 | `K21903` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K21903» aparece como «21∕4Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-333, SA-334 |  |  |
| 79 | `K11820` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «K11820» mas de una: C-1∕2Mo; A. No se elige por cuenta propia. <br>Especificaciones: SA-691, SA-204, SA-672 |  |  |
| 80 | `K12020` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «K12020» mas de una: C-1∕2Mo; B. No se elige por cuenta propia. <br>Especificaciones: SA-691, SA-204, SA-672 |  |  |
| 81 | `K12320` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «K12320» mas de una: C-1∕2Mo; C-1∕2Mo A204 Gr.. No se elige por cuenta propia. <br>Especificaciones: SA-691, SA-204, SA-672 |  |  |
| 82 | `S43035` (UNS) | 8 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-268, SA-479, SA-803 |  |  |
| 83 | `29Cr-4Mo` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |  |  |
| 84 | `29Cr-4Mo-2Ni` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-479, SA-240 |  |  |
| 85 | `35Ni-30Fe-24Cr-6Mo-Cu` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |  |  |
| 86 | `37Ni-33Fe-23Cr-4Mo-Cu` | 8 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-464, SB-468 |  |  |
| 87 | `K13047` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 88 | `G41350` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 89 | `G41370` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «G41370» aparece como «Cr-Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-574, SA-372 |  |  |
| 90 | `K13548` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 91 | `K71340` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K71340» aparece como «8Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-553, SA-522 |  |  |
| 92 | `K12521` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-533 |  |  |
| 93 | `K12023` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-250, SA-209 |  |  |
| 94 | `K11422` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-250, SA-209 |  |  |
| 95 | `23/4Ni-11/2Cr-1/2Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-543 |  |  |
| 96 | `18Cr-2Mo` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |  |  |
| 97 | `27Cr-1Mo-Ti` | 6 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268, SA-240 |  |  |
| 98 | `J13002` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-487 |  |  |
| 99 | `K14072` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K14072» aparece como «Cr-Mo-V», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-193 |  |  |
| 100 | `G40370` (UNS) | 6 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «G40370» aparece como «Cr-Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-574, SA-320 |  |  |
| 101 | `K13050` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-350 |  |  |
| 102 | `K61365` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-553 |  |  |
| 103 | `19Cr-9Ni-2Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 104 | `S40300` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-479 |  |  |
| 105 | `S41003` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «S41003» mas de una: 12Cr; 12Cr-1Ni. No se elige por cuenta propia. <br>Especificaciones: SA-1010 |  |  |
| 106 | `S30100` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 107 | `S40800` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-268 |  |  |
| 108 | `S43036` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S43036» aparece como «18Cr-Ti», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-268 |  |  |
| 109 | `16Cr-12Ni-2Mo-Cb` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 110 | `16Cr-4Ni-6Mn` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 111 | `16Cr-9Mn-2Ni-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 112 | `18Cr-9Ni-3Cu-Cb-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |  |  |
| 113 | `20.5Cr-8.8Ni-Mo-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240 |  |  |
| 114 | `25Cr-20Ni-Cb-N` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-213 |  |  |
| 115 | `25Cr-4Ni-4Mo-Ti` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-240, SA-268 |  |  |
| 116 | `29Cr-4Mo-Ti` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-268 |  |  |
| 117 | `J82090` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «J82090» aparece como «9Cr-1Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-426, SA-217 |  |  |
| 118 | `J91150` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y el libro asocia a su UNS «J91150» mas de una: 12Cr; 13Cr. No se elige por cuenta propia. <br>Especificaciones: SA-426, SA-217 |  |  |
| 119 | `19Cr-10Ni-3Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 120 | `19Cr-9Ni-1/2Mo` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 121 | `Mn-1/4Mo-V` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-487 |  |  |
| 122 | `G40420` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |  |  |
| 123 | `G41420` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |  |  |
| 124 | `G41450` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-574 |  |  |
| 125 | `Co-26Cr-9Ni-5Mo-3Fe-2W` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-818, SB-815 |  |  |
| 126 | `S30300` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S30300» aparece como «18Cr-9Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-320 |  |  |
| 127 | `12Cr-1Mo-V-W` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-437 |  |  |
| 128 | `67Ni-28Cu-3Al` | 4 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SF-468 |  |  |
| 129 | `A02040` (UNS) | 4 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SB-108, SB-26 |  |  |
| 130 | `K12520` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-336 |  |  |
| 131 | `K22036` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-350 |  |  |
| 132 | `K14508` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-372 |  |  |
| 133 | `K32026` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-765 |  |  |
| 134 | `K21703` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K21703» aparece como «21∕4Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-203 |  |  |
| 135 | `K22103` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K22103» aparece como «21∕4Ni», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-203 |  |  |
| 136 | `K41583` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «K41583» aparece como «5Ni-1∕4Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-645 |  |  |
| 137 | `K11224` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-562 |  |  |
| 138 | `K12047` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-213 |  |  |
| 139 | `11/2Si-1/2Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-335 |  |  |
| 140 | `13/4Ni-3/4Cr-Mo` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-372 |  |  |
| 141 | `S41500` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-182 |  |  |
| 142 | `S40910` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 143 | `S40920` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 144 | `S40930` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 145 | `S43932` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-240 |  |  |
| 146 | `S44600` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «S44600» aparece como «27Cr», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-268 |  |  |
| 147 | `20Cr-10Ni` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-479 |  |  |
| 148 | `J31545` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal. Su UNS «J31545» aparece como «3Cr-Mo», pero esa composicion no figura en ninguna Nota de TM-1 ni TE-1. <br>Especificaciones: SA-426 |  |  |
| 149 | `17Cr-4Ni-3Cu` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-747 |  |  |
| 150 | `K50100` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SA-193 |  |  |
| 151 | `2Ni-3/4Cr-1/3Mo-V` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-540 |  |  |
| 152 | `32Ni-45Fe-20Cr-Cb` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-351 |  |  |
| 153 | `53Ni-17Mo-16Cr-6Fe-5W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 154 | `59Ni-22Cr-14Mo-4Fe-3W` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 155 | `62Ni-28Mo-5Fe` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SA-494 |  |  |
| 156 | `R61702` (UNS) | 2 | SIN MAPEO | La fila no imprime composicion nominal y su UNS no aparece con composicion en ninguna tabla del libro: no hay dato con el que buscar el grupo. <br>Especificaciones: SB-752 |  |  |
| 157 | `95.5Zr + 2.5Nb` | 2 | SIN MAPEO | El UNS no figura en TM-1..TM-5 y la composicion nominal no esta listada en ninguna Nota de TM-1 ni TE-1. II-D no publica E ni dilatacion para este material. <br>Especificaciones: SB-752 |  |  |

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
