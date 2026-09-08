# Verificacion de resources/ contra los PDF de origen

De cada bloque se toma una muestra del centro de su texto y se busca en la pagina del PDF que el propio bloque declara. Si no aparece alli, se exploran las paginas contiguas para separar 'no esta' de 'esta desplazado'.

| Fuente | Comprobacion | Detalle | Estado |
|---|---|---|---|
| ASME PCC-2 | PDF de origen | ASME PCC-2 REPAIR OF PRESSURE EQUIPMENT AND PIPING.pdf, 326 paginas | OK |
| ASME PCC-2 | bloques localizados en el PDF | 2472 de 2477 (99.8%) | REVISAR |
| ASME PCC-2 | bloques con pagina propia desplazados | 0 | OK |
| ASME PCC-2 | bloques que no aparecen | 5 | REVISAR |
| ASME PCC-2 | bloques demasiado cortos para comprobar | 1757 de 4234 | informativo |
| ASME PCC-2 | de los no localizados: contenido presente, orden de lectura entrelazado | 4 | informativo |
| ASME PCC-2 | de los no localizados: sin palabras largas (ecuacion o simbolos) | 1 | informativo |
| B31.3 capitulos | PDF de origen | ASME B31.3 2024 Process Piping.pdf, 594 paginas | OK |
| B31.3 capitulos | bloques localizados en el PDF | 1705 de 1722 (99.0%) | REVISAR |
| B31.3 capitulos | parrafos hallados mas adelante en su seccion | 173 (hasta +5 paginas) | informativo |
| B31.3 capitulos | bloques con pagina propia desplazados | 0 | OK |
| B31.3 capitulos | bloques que no aparecen | 17 | REVISAR |
| B31.3 capitulos | bloques demasiado cortos para comprobar | 2511 de 4233 | informativo |
| B31.3 capitulos | de los no localizados: contenido presente, orden de lectura entrelazado | 12 | informativo |
| B31.3 capitulos | de los no localizados: sin palabras largas (ecuacion o simbolos) | 3 | informativo |
| B31.3 capitulos | de los no localizados: parcialmente presente | 2 | REVISAR |
|   | texto ausente | chapter_03.json, pagina 85 (paragraph), 64% de sus palabras: 'Temperature Reduction, 8C (8F) GENERAL NOTES: (a) See Table ' | REVISAR |
|   | texto ausente | chapter_05.json, pagina 103 (paragraph), 62% de sus palabras: 'T2 = minimum thickness of fabricated lap fillet weld shall f' | REVISAR |
| B31.3 apendices D-Z | bloques localizados en el PDF | 486 de 496 (98.0%) | REVISAR |
| B31.3 apendices D-Z | parrafos hallados mas adelante en su seccion | 82 (hasta +4 paginas) | informativo |
| B31.3 apendices D-Z | bloques con pagina propia desplazados | 0 | OK |
| B31.3 apendices D-Z | bloques que no aparecen | 10 | REVISAR |
| B31.3 apendices D-Z | bloques demasiado cortos para comprobar | 910 de 1406 | informativo |
| B31.3 apendices D-Z | de los no localizados: contenido presente, orden de lectura entrelazado | 8 | informativo |
| B31.3 apendices D-Z | de los no localizados: sin palabras largas (ecuacion o simbolos) | 1 | informativo |
| B31.3 apendices D-Z | de los no localizados: parcialmente presente | 1 | REVISAR |
|   | texto ausente | appendix_s\appendix_s_text.json, pagina 548 (paragraph), 93% de sus palabras: 'para. 320.2, and are applied to both the stresses due to sus' | REVISAR |
| B31.3 tablas | bloques localizados en el PDF | 114 de 114 (100.0%) | OK |
| B31.3 tablas | bloques con pagina propia desplazados | 0 | OK |
| B31.3 tablas | bloques que no aparecen | 0 | OK |
| B31.3 tablas | bloques demasiado cortos para comprobar | 0 de 114 | informativo |
| II-D metrica | PDF de origen | D Metric 2025 .pdf, 1537 paginas | OK |
| II-D metrica | filas cuyo UNS aparece en el rango de su tabla | 10356 de 10356 (100.0%) | OK |
| II-D metrica | filas cuya especificacion aparece en el rango | 11124 de 11124 (100.0%) | OK |
| II-D metrica | filas sin especificacion en el JSON (continuacion) | 13 de 11137 | informativo |
| II-D U.S. Customary | PDF de origen | D Customary 2025 .pdf, 1533 paginas | OK |
| II-D U.S. Customary | filas cuyo UNS aparece en el rango de su tabla | 10329 de 10329 (100.0%) | OK |
| II-D U.S. Customary | filas cuya especificacion aparece en el rango | 11066 de 11066 (100.0%) | OK |
| II-D U.S. Customary | filas sin especificacion en el JSON (continuacion) | 15 de 11081 | informativo |

**Total de bloques mal ubicados: 3.**
