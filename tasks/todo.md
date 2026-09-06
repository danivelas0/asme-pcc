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

---

## Segunda tanda — cierre del vacío en la edición U.S. Customary

`bpvc_ii_d_customary_2025` tenía el mismo vacío que la métrica: TM-1 y TE-1 sin
`note_members`. El libro construye bandas de propiedades en ambas ediciones
(`DB_E`/`DB_EC`, `DB_TE`/`DB_TEC`), así que media tabla quedaba sin vía trazable
para saber a qué grupo pertenece un material.

- [x] `extraer_notas_ii_d.py` acepta `--edicion si|us`; cada edición sale de su
      propio PDF. Nunca copiar una en la otra.
- [x] Notas de TM-1 (1)–(16) y TE-1 (1)–(4) de la edición US en `resources/`.
      Cambio puramente aditivo; la métrica quedó intacta (verificado).
- [x] La observación de duplicado se detecta **por los datos**, no por número de
      nota. Antes estaba codificada a la Nota (8), que en la US es legítima.
- [x] `verificar.py` §9 audita que las 4 combinaciones edición × tabla traigan
      `note_members`, para que el vacío no pueda reabrirse en silencio.
- [x] `TestNotasEnLasDosEdiciones` en `test_build_db.py`.

**Hallazgo.** Las dos ediciones **no numeran igual sus notas**. La métrica
imprime el Grupo H duplicado —Notas (8) y (9)— y desde ahí toda su numeración va
corrida en uno respecto de la US, que lo imprime una sola vez en la Nota (8):
**17 notas en TM-1 métrica contra 16 en la US**. Esto confirma que la errata es
exclusiva de la métrica y que conservarla verbatim fue lo correcto.

La **pertenencia** a grupo sí es idéntica en las dos ediciones, verificada grupo
a grupo. La única diferencia de contenido es el orden de dos miembros del
`Group 2` de TE-1 (`25Cr–7Ni–4Mo–N` y `25Cr–6Ni–Mo–N` intercambiados), artefacto
de la maquetación a columnas; cada edición se carga en el orden que imprime.

**Verificación:** 57 pruebas pasan; `verificar.py` 0 fallos, salida 0, con
`4/4 archivos con note_members`.

---

## Tercera tanda — el resto de lo abierto

- [x] **Columna PRD.** Se eliminó la coincidencia de subcadena. Ahora es
      coincidencia literal del UNS contra las **99 de 115 filas de PRD que
      nombran los suyos** (142 UNS): 1 554 filas resueltas por hoja. Las 16
      restantes son categorías redactadas y quedan al criterio del ingeniero.
- [x] **Las filas sin composición nominal impresa.** 386 de las 427 sí tienen
      UNS, y ese UNS aparece con composición en otra tabla del libro. Se
      recupera por ahí — pero **nunca como AUTO**: estado propio
      `REVISAR (composicion de otra tabla)`, citando la tabla de origen. 193
      filas por hoja salen del bloqueo con propuesta trazable. Si el libro
      asocia dos composiciones distintas al mismo UNS, no se elige.
- [x] **Mapeo de la edición US.** `MAP_GrupoC`, resuelta contra las Notas de la
      edición US. El libro pasa a 39 hojas; la nueva es `veryHidden`.
- [x] **Vía de retorno de las decisiones.** `decisiones_map_grupo.json` +
      plantilla autogenerada. Estado `VALIDADO POR INGENIERO`, separado de las
      AUTO, con firma y fecha en la fuente. Una decisión no puede pisar al
      código: el choque se reporta y se conserva lo normativo.

**Dos defectos encontrados al hacerlo:**

1. **Barras de fracción.** II-D imprime `C-1/2Mo` y el Apéndice A del B31.3
   `C-1∕2Mo` (U+2215). No se pliegan por NFKC, así que el contraste entre
   tablas fallaba **en silencio**: no daba error, simplemente no encontraba
   nada. Corregido en `comp_key`.
2. **Cita de la nota huérfana.** La métrica citaba `TM-1 Nota (8)` para el
   Grupo H, pero su tabla referencia la **(9)** — la (8) es la duplicada. Ahora
   se lee del rótulo impreso qué nota referencia cada tabla y se cita esa.
   Métrica → (9), US → (8).

**Verificación:** 69 pruebas pasan; `verificar.py` 0 fallos, salida 0. La
sección 9 audita **las dos hojas**, cada una contra las Notas de su edición.

**Queda solo lo que es tuyo:** las 157 decisiones de `Revision_MAP_Grupo.md`.
