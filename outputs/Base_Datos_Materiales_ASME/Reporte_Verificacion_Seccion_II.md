# Verificacion de ASME BPVC Seccion II, partes A, B y C

Cada especificacion troceada se contrasta con su PDF de origen: que el corte empiece donde dice, que el folio impreso cuadre y que no se pierdan paginas entre una especificacion y la siguiente.

| Parte | Comprobacion | Detalle | Estado |
|---|---|---|---|
| bpvc_ii_a_1 | paginas declaradas frente a las que suman las entradas | 746 / 746 (el PDF tiene 803) | OK |
| bpvc_ii_a_1 | entradas con su archivo en disco | 76 de 76 | OK |
| bpvc_ii_a_1 | archivos sin entrada en el indice | 0 de 76 | OK |
|   | ruta del indice | `file` apunta a `SA-6-SA-6M/SA-6-SA-6M.json`, que no existe: en disco estan aplanados en `specifications/` | revisar |
| bpvc_ii_a_1 | la designacion aparece en la 1a pagina declarada | 76 de 76 | OK |
| bpvc_ii_a_1 | el folio impreso coincide | 76 de 76 | OK |
| bpvc_ii_a_1 | rangos que se solapan | 0 | OK |
| bpvc_ii_a_1 | paginas cubiertas por las entradas | 746 de 803 (56 preliminares, 1 finales, 0 en huecos interiores) | OK |
| bpvc_ii_a_1 | figuras PNG | 48 declaradas / 48 en disco | OK |
| bpvc_ii_a_2 | paginas declaradas frente a las que suman las entradas | 888 / 888 (el PDF tiene 943) | OK |
| bpvc_ii_a_2 | entradas con su archivo en disco | 119 de 119 | OK |
| bpvc_ii_a_2 | archivos sin entrada en el indice | 0 de 119 | OK |
|   | ruta del indice | `file` apunta a `SA-451-SA-451M/SA-451-SA-451M.json`, que no existe: en disco estan aplanados en `specifications/` | revisar |
| bpvc_ii_a_2 | la designacion aparece en la 1a pagina declarada | 119 de 119 | OK |
| bpvc_ii_a_2 | el folio impreso coincide | 119 de 119 | OK |
| bpvc_ii_a_2 | rangos que se solapan | 0 | OK |
| bpvc_ii_a_2 | paginas cubiertas por las entradas | 888 de 943 (54 preliminares, 1 finales, 0 en huecos interiores) | OK |
| bpvc_ii_a_2 | figuras PNG | 55 declaradas / 55 en disco | OK |
| bpvc_ii_b | paginas declaradas frente a las que suman las entradas | 1266 / 1266 (el PDF tiene 1319) | OK |
| bpvc_ii_b | entradas con su archivo en disco | 148 de 148 | OK |
| bpvc_ii_b | archivos sin entrada en el indice | 0 de 148 | OK |
|   | ruta del indice | `file` apunta a `SB-26-SB-26M/SB-26-SB-26M.json`, que no existe: en disco estan aplanados en `specifications/` | revisar |
| bpvc_ii_b | la designacion aparece en la 1a pagina declarada | 148 de 148 | OK |
| bpvc_ii_b | el folio impreso coincide | 148 de 148 | OK |
| bpvc_ii_b | rangos que se solapan | 0 | OK |
| bpvc_ii_b | paginas cubiertas por las entradas | 1266 de 1319 (52 preliminares, 1 finales, 0 en huecos interiores) | OK |
| bpvc_ii_b | figuras PNG | 65 declaradas / 65 en disco | OK |
| bpvc_ii_c | paginas declaradas frente a las que suman las entradas | 1110 / 1110 (el PDF tiene 1153) | OK |
| bpvc_ii_c | entradas con su archivo en disco | 36 de 36 | OK |
| bpvc_ii_c | archivos sin entrada en el indice | 0 de 36 | OK |
|   | ruta del indice | `file` apunta a `SFA-5.01M-SFA-5.01/SFA-5.01M-SFA-5.01.json`, que no existe: en disco estan aplanados en `specifications/` | revisar |
| bpvc_ii_c | la designacion aparece en la 1a pagina declarada | 36 de 36 | OK |
| bpvc_ii_c | el folio impreso coincide | 36 de 36 | OK |
| bpvc_ii_c | rangos que se solapan | 0 | OK |
| bpvc_ii_c | paginas cubiertas por las entradas | 1110 de 1153 (42 preliminares, 1 finales, 0 en huecos interiores) | OK |
| bpvc_ii_c | figuras PNG | 150 declaradas / 150 en disco | OK |
| bpvc_ii_d_metric_2025 | convencion de `pdf_pages` | desplazamiento +0 (8 de 8 tablas de muestra) | OK |
| bpvc_ii_d_metric_2025 | tablas cuyo `table_id` aparece en su primera pagina | 27 de 27 | OK |
| bpvc_ii_d_metric_2025 | rangos dentro del PDF | 27 de 27 (el PDF tiene 1537 paginas) | OK |
| bpvc_ii_d_metric_2025 | `row_count` coincide con las filas cargadas | 27 de 27 | OK |
| bpvc_ii_d_customary_2025 | convencion de `pdf_pages` | desplazamiento +0 (8 de 8 tablas de muestra) | OK |
| bpvc_ii_d_customary_2025 | tablas cuyo `table_id` aparece en su primera pagina | 27 de 27 | OK |
| bpvc_ii_d_customary_2025 | rangos dentro del PDF | 27 de 27 (el PDF tiene 1533 paginas) | OK |
| bpvc_ii_d_customary_2025 | `row_count` coincide con las filas cargadas | 27 de 27 | OK |

**Total de fallos: 0.**
