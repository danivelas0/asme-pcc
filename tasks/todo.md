# Estado abierto del proyecto

**Actualizado:** 2026-09-08 · Libro vigente: `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm`

Lo que fue este archivo —el registro de las tres tandas de validación de
`MAP_Grupo`, con sus casillas ya marcadas— se retiró al cerrarse ese trabajo. El
relato por revisión vive en `outputs/Base_Datos_Materiales_ASME/LEEME_Nota_de_Version.md`
y el porqué de cada decisión, en `CLAUDE.md` y en los comentarios del código que
la implementa. Duplicarlo aquí solo garantizaba que una de las dos copias
envejeciera.

## No hay nada pendiente de firma

`MAP_Grupo` cerró con **0 casos abiertos**. Las **7 decisiones** de
`decisiones_map_grupo.json` son todas las que hay, y los **106 casos restantes
están CERRADOS por límite de la fuente**: el material queda identificado sin
ambigüedad y aun así II-D no publica su grupo. Firmar ahí sería inventar el
valor que el código no imprime. Lo correcto es que esas filas sigan bloqueadas.

`Revision_MAP_Grupo.md` separa los dos bloques y solo el de ABIERTOS lleva
casilla. Si un día vuelve a haber una, aparecerá ahí sola.

## Límites declarados, que no son deuda

Ninguno de estos es un hueco por cerrar: es la fuente la que no publica el dato,
o es una decisión de alcance ya tomada. Se listan para que nadie los reabra
creyendo que se olvidaron.

| Límite | Clase | Dónde está documentado |
|---|---|---|
| Tabla TE-2 (aleaciones de aluminio) sin fila de dilatación | límite de la fuente: el código no imprime con qué distinguir la columna | `CLAUDE.md`, `etiqueta_columna_b_te1()` |
| Coeficientes A y C de TE-1..5 fuera del buscador | alcance: solo el B alimenta cálculo de dilatación | `CLAUDE.md`; siguen impresos en `DB_TE`/`DB_TEC` |
| Tablas B-2 a B-6 del Apéndice B no cargadas | alcance (Rev. 4d) | extracción intacta en `resources/`; tarjeta marcador en el árbol |
| 16 filas de la Tabla PRD sin resolver | criterio de ingeniería: son categorías redactadas, no listas de UNS | `CLAUDE.md`, columna `Fila PRD` |
| 24 812 filas AMBIGUAS de la Sección II (45,8 %) | límite de la fuente: sin `Span` no hay posición de palabra | `IDX_SecII_Tablas`, tabla a tabla |
| Las normalizadas de química y tracción cubren 220 y 182 filas | límite de la fuente: los encabezados llegan sin partir | `Revision_Tablas_SecII.md`, con los dos motivos contados |
| `Ej` sin correspondencia fila a fila entre A-3 y 302.3.4-1 | el B31.3 no la imprime; se transcriben las 10 filas | `CLAUDE.md`, regla 11 |
| B31.1, B16, SEC. VIII, PCC-1, PCC-3 y el resto de PCC-2 | alcance: no cargados | tarjeta marcador `NO CARGADO EN ESTE LIBRO` |

## Lo único que exige una máquina concreta

`verificar_resources.py` y `verificar_seccion_ii.py` **necesitan los PDF de ASME**,
que no se versionan. Sin `--pdfs` cada fuente sale como `REVISAR`, que significa
«no comprobado», no «mal». Solo los corre en verde quien tenga los PDF en disco
(ver la ruta en `CLAUDE.md`). No afecta a `verificar.py`, que audita el libro
contra los JSON y corre en cualquier copia del repositorio.

---

Los errores ya pagados —y la regla que dejó cada uno— están en `lessons.md`.
