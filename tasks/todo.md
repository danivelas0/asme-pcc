# Validación de MAP_Grupo — cierre del vacío normativo en `resources/`

**Fecha:** 2026-09-06
**Origen:** pendiente declarado en `CLAUDE.md` — «1 518 PROPUESTA que requieren
validación del ingeniero antes de usar E o dilatación».

## Diagnóstico

1. Las 1 518 filas PROPUESTA se reducen a **124 composiciones nominales distintas**.
2. `FERROUS_RULES` (regex sobre la composición impresa) produce asignaciones
   **falsas**, no solo inciertas: `mo\b` captura austeníticos, dúplex y aleaciones
   de níquel y los rotula «acero de baja aleación»; `8ni` casa dentro de `18Ni`.
3. El mapeo composición → grupo **sí es normativo**: son las Notas de las Tablas
   TM-1 y TE-1 de ASME BPVC II-D. La extracción en `resources/` estaba incompleta:

   | Archivo | Notas presentes | Notas faltantes |
   |---|---|---|
   | `table_tm_1.json` | (1), (2) → Grupos A, B | (3)–(17) → Grupos C…J + 6 UNS PH |
   | `table_te_1.json` | GENERAL + (1) → Grupo 1 | (2), (3), (4) → Grupos 2, 3, 4 |

4. Contraste literal de las 124 composiciones contra las listas impresas:
   **72 composiciones / 1 174 filas (77 %) resuelven 1:1 y auditables.**
5. Las 344 filas restantes son en su mayoría **aleaciones de níquel** (`N06985`,
   `N06022`, …) que TM-1 no cubre por ser tabla de materiales ferrosos, y cuyo UNS
   tampoco figura en TM-2…TM-5. **II-D no publica E para ellas**: corresponde
   bloquear el cálculo, no remapear.

## Tareas

- [x] 1. `scripts/extraer_notas_ii_d.py` — extrae del PDF de II-D métrica las Notas
      de TM-1 (1)–(17) y TE-1 (1)–(4). Verbatim + `note_members` segmentado, con
      folio impreso de origen para auditoría.
- [x] 2. Completar `resources/.../table_tm_1.json` y `table_te_1.json`.
      Conservar la duplicación real de las Notas (8) y (9) tal como está impresa
      (Regla 9), con la observación anotada.
- [x] 3. Reescribir `build_map_grupo` en `build_db_materiales.py`; eliminar
      `FERROUS_RULES`. Resolución por coincidencia literal contra `note_members`.
      Cada fila cita la nota que la sostiene.
- [x] 4. Estados nuevos; se elimina `PROPUESTA (VALIDAR)`.
- [x] 5. Hoja de revisión con el residuo no resuelto.
- [x] 6. Cobertura en `test_build_db.py` y auditoría fila a fila en `verificar.py`.
- [x] 7. Reconstruir el libro; `pytest` verde y `verificar.py` devolviendo 0.

## Revisión

**Resultado.** El pendiente pasó de «1 518 filas que requieren validación» a
**65 decisiones distintas**, y ninguna cifra del motor descansa ya en una regex.

| Estado | Filas | Respaldo |
|---|---:|---|
| AUTO (UNS exacto) | 1 292 | UNS impreso en TM-1…TM-5 |
| AUTO (composición en Nota) | 1 297 | Nota de TM-1 / TE-1, citada en la fila |
| REVISAR (regla textual) | 17 | TM-1 Nota (5), «9Cr–Mo, including variations thereof» |
| SIN MAPEO | 848 | II-D no publica el dato; cálculo bloqueado |

De las 848 SIN MAPEO, **427 no imprimen composición nominal** en la tabla de
origen: no hay dato de entrada que juzgar, y se cuentan aparte en la hoja de
revisión para no inflar la lista de decisiones.

**Hallazgos de fidelidad al impreso.** Dos defectos de la extracción original de
II-D salieron a la luz y quedaron corregidos:

- La Nota (1) de TM-1 se imprime a dos columnas y el JSON conservaba solo una:
  guardaba 4 de los 8 miembros del Grupo A.
- La Nota (2) continúa del folio 1188 al 1189; faltaban 7 de sus 21 miembros.

Y uno que **no** es defecto y se conserva tal cual (Regla 9): ASME imprime las
Notas (8) y (9) duplicadas, ambas definiendo el Grupo H con idéntica lista.
Verificado por coordenadas en el folio 1190. No se deduplica; queda anotado en
`notes_extraction.observaciones`.

**Verificación.** `pytest test_build_db.py test_dashboard.py` → 53 pasan.
`verificar.py` → 0 fallos en las 9 secciones, código de salida 0. La sección 9
es nueva: comprueba fila a fila que la nota citada liste realmente esa
composición, y falla si reaparece un estado de conjetura.

**Cambio de alcance frente al plan.** Se descartó remapear las aleaciones de
níquel contra TM-3/TM-4: se verificó que TM-3 es cobre y TM-4 alto níquel
indexada por UNS, y que sus UNS (`N06985`, `N06022`, …) no figuran en ninguna
tabla TM. II-D sencillamente no publica E para ellas, así que bloquear es lo
correcto y no había nada que remapear.

**Queda pendiente:** la columna `Grupo PRD` sigue con la coincidencia de
subcadena heredada de la Rev. 2. No alimenta ningún cálculo. Sin tocar.
