# Proyecto ASME PCC — Motor de Cálculo e Ingeniería de Reparación

Ingeniería de reparación de equipos a presión y tubería según la familia **ASME PCC**
(PCC-1 / PCC-2 / PCC-3), con los códigos de construcción **ASME B31.3-2024** y
**ASME BPVC Sección VIII-1 + II-D 2025** como referencia.

Responde siempre en **español**. Unidades **SI por defecto** (MPa, mm, °C).
Tono profesional y directo, sin relleno.

---

## Regla nº 1 — `resources/` es la única fuente de verdad

`resources/` contiene los códigos y normas extraídos a JSON. **Ningún valor normativo
—esfuerzos admisibles, factores, fórmulas, límites— puede salir de la memoria del
modelo.** Siempre se lee el archivo fuente antes de citar o calcular.

Si un dato no existe en `resources/`, **no se inventa ni se aproxima**: se declara
el vacío explícitamente y se pregunta cómo proceder.

Todo coeficiente o tabla que alimente un motor de cálculo debe **trazar** al archivo
concreto de `resources/` del que procede, para permitir auditoría posterior.

## Estructura

```
knowledge/     Instrucciones de cálculo ASME PCC-2 en SI. Leer antes de cualquier tarea.
resources/     Códigos y normas (JSON). Fuente única de verdad.
               ├─ ASME B31/ASME B31.3/APPEX/   Apéndices A, B y C
               ├─ ASME PCC/pcc_2/              Artículos de PCC-2
               ├─ bpvc_ii_d_metric_2025/       II-D métrica (MPa, °C)
               └─ bpvc_ii_d_customary_2025/    II-D U.S. Customary (ksi, °F)
outputs/       Entregables.
templates/     maestro_con_macros.xlsm — Rev. 0 + proyecto VBA, entrada del builder.
               Se genera con scripts/make_vba_seed.py; no se edita a mano.
```

**Nunca leas `outputs/` ni `templates/`** salvo que se te señale un archivo.
**Guarda todo entregable en `outputs/` dentro de una subcarpeta** con nombre de proyecto.

Ante una duda de alcance, pregunta antes de producir.

---

## Motor de cálculo — estado actual

Entregable vigente: `outputs/Motor_de_Calculo_ASME_PCC_Rev3.xlsm`
(39 hojas, **una sola visible**). Se **genera por script**, nunca se edita a mano.

**Es un libro con macros.** Al abrirlo se ve solo el `Dashboard`; la navegación a los
nueve motores (Art. 212, los 7 buscadores, `Instrucciones`) la hace un proyecto VBA de
dos componentes. Los estados de visibilidad van **grabados en el archivo**, así que con
las macros bloqueadas no se expone ninguna base de datos.

### Reconstruir

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts

# Solo la primera vez, o al tocar vba/*.vba.
# Activa "Confiar en el acceso al modelo de objetos de proyectos de VBA" unos
# segundos y lo restaura, verificando el resultado. Si avisa de que no pudo
# restaurarlo, desactivelo a mano en el Centro de confianza.
python make_vba_seed.py

# Solo si se repone la extraccion de II-D: notas de grupo de TM-1 / TE-1.
python extraer_notas_ii_d.py --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Metric 2025\D Metric 2025 _p1201-p1500.pdf"

python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm
python -m pytest test_build_db.py test_dashboard.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm
```

La entrada del builder es el maestro sembrado en `templates/`, que a su vez sale de
`outputs/Base_Datos_Materiales_ASME/Motor_de_Calculo_ASME_PCC_Rev0_respaldo.xlsx`
(3 hojas). El `Motor_de_Calculo_ASME_PCC.xlsx` de la raíz **no está en el repo**.

`verificar.py` devuelve 0 solo si todo pasa. Audita **fila a fila** cada valor
tabulado contra el JSON del código (271 276 valores), la contigüidad de la cascada,
la ausencia de fórmulas de matriz dinámica, la interpolación recalculada en hoja, el
caso semilla y la capa de navegación. **Requiere Excel instalado**: recalcula con el
motor real, no con LibreOffice. **Ejecútalo siempre después de tocar el builder.**

### Capa de navegación — dos fuentes de verdad que no pueden divergir

La tabla de navegación vive por duplicado: en `build_db_materiales.py`, que la graba en
el archivo, y en `scripts/vba/mod_nav.vba`, que la reaplica al abrir. Si se separan, el
libro se abre mostrando algo distinto de lo que se construyó. `test_dashboard.py` lo
comprueba (`TestSincroniaPythonVba`). Al tocar una, tocar la otra:

| Concepto | Python | VBA |
|---|---|---|
| Hojas navegables | `NAVEGABLES` | `HojasNavegables()` |
| Columna base de claves | `COL_CLAVE_BASE = 66` | `COL_CLAVE_BASE` |
| Celda del aviso | `FILA_AVISO = 4` | `CELDA_AVISO = "A4"` |

Dos trampas ya pagadas, documentadas en el código:

- **El VBA referencia hojas por `.Name`, nunca por CodeName.** openpyxl no asigna
  `codeName` a las 35 hojas que crea; Excel se los inventa al abrir.
- **Toda declaración de módulo (`Const`, `Dim`, `Type`) precede a la primera rutina.**
  Si no, VBA reporta «Variable not defined» en cada uso, Excel abre un diálogo modal al
  compilar durante `SaveAs`, y la automatización se cuelga sin mensaje. `make_vba_seed.py`
  lo comprueba con `lint_vba()` antes de tocar COM.

### Reglas de diseño del libro — no romper

1. **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`,
   `XLOOKUP`, `VSTACK`, `_xlfn`. Solo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`.
   **Enmienda Rev. 3:** el entregable es `.xlsm` y ya **no abre en Google Sheets** —
   la capa de navegación es VBA. Lo que se pierde es solo la navegación. La
   restricción sobre las fórmulas **se mantiene entera**: todo el cálculo debe seguir
   siendo portable, y `verificar.py` sigue fallando si aparece una matriz dinámica.
   Corolario aprendido: **un texto no se guarda como fórmula.** Un `="texto…"` de más
   de 255 caracteres lo parte Excel en `_xlfn._LONGTEXT(...)` al reguardar, y la celda
   pasa a mostrar `#NAME?` fuera de Excel 365. El builder lo normaliza
   (`normalizar_textos_como_formula`).
2. **Validación de datos: solo rango literal o lista de ítems.** Nunca una fórmula
   (`OFFSET`, `INDIRECT`) como origen. Las listas dependientes se materializan en
   celdas de columnas ocultas y la validación apunta a ese rango.
3. **Banda compacta.** Las tablas ASME dejan huecos interiores (la A-1 no imprime
   125 °C para A106 Gr.B). Junto a la banda impresa —que se conserva para auditar—
   cada base lleva una banda con solo los puntos existentes, sin huecos, y la
   consulta trabaja sobre ella.
4. **No extrapolar.** Por encima de la última temperatura tabulada o de la Temp. máx.
   del material, dictamen "FUERA DE RANGO" y resultado bloqueado. Lo prohíbe el código.
5. **Cascada contigua.** Cada base se ordena por familia → composición → forma →
   especificación → grado, para que cada nivel sea un bloque contiguo.
6. **Unidades visibles** en toda variable, indicadores y ejes de gráficas.
7. **`material_id` único**, desambiguado en este orden: notas del código → Rm → Re →
   número de línea impreso. El código repite spec+grado con admisibles distintos.
8. La **familia de material** es una agrupación de navegación derivada del UNS y la
   composición impresa. **No es dato normativo** y no entra en ningún cálculo.
9. Los valores se cargan **tal como están impresos**; SI y US son extracciones
   independientes, nunca conversiones.

### `MAP_Grupo` — pertenencia a grupo de propiedades

El grupo que da E (TM-1) y dilatación (TE-1) **no se infiere de la composición**:
está impreso en las Notas al pie de esas dos tablas. Viven en `resources/` bajo
`note_members`, **en las dos ediciones**, y las extrae del PDF de II-D:

```powershell
python extraer_notas_ii_d.py --edicion si --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Metric 2025\D Metric 2025 _p1201-p1500.pdf"
python extraer_notas_ii_d.py --edicion us --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Customary 2025\D Customary 2025 _p1201-p1500.pdf"
```

Solo hay que reejecutarlo si se repone la extracción de II-D. El builder **aborta**
si `note_members` no está: sin fuente normativa no se construye nada.

**Las dos ediciones no numeran igual sus notas.** La métrica imprime el Grupo H
**duplicado**, en las Notas (8) y (9), y a partir de ahí toda su numeración va
corrida en uno respecto de la US, que lo imprime una sola vez en la Nota (8): 17
notas en TM-1 métrica contra 16 en la US. La errata es exclusiva de la métrica y
se conserva tal como está impresa (Regla 9). La **pertenencia** a grupo sí es
idéntica en ambas, verificado grupo a grupo. Nunca copiar las notas de una
edición a la otra: son extracciones independientes y `test_build_db.py::TestNotasEnLasDosEdiciones`
lo comprueba.

Hay **dos hojas de mapeo**, una por edición: `MAP_Grupo` (métrica) y `MAP_GrupoC`
(U.S. Customary). Cada una se resuelve contra las Notas de **su** edición.

Estado por hoja (3 454 filas cada una):

| Estado | Filas | Respaldo |
|---|---:|---|
| AUTO (UNS exacto) | 1 292 | UNS impreso en TM-1…TM-5 |
| AUTO (composición en Nota o columna) | 1 297 | Nota, o columna nombrada de TE-1, citada en la propia fila |
| AUTO (composición vía UNS en otra tabla) | 193 | el código imprime esa composición para el mismo UNS en otra de sus tablas; el UNS identifica el material de forma unívoca |
| VALIDADO POR INGENIERO | 17 | `decisiones_map_grupo.json`; hoy solo el 9Cr-1Mo-V |
| SIN MAPEO | 655 | II-D no publica el dato; cálculo bloqueado |

**Cobertura:** 2 451 filas con módulo E y 1 318 con dilatación, de 3 454. Una
fila puede tener uno y no el otro — TM-1 y TE-1 no enumeran los mismos
materiales — y la columna `Motivo` lo dice fila a fila.

`verificar.py` §9 audita ambas hojas fila a fila: que toda cita exista realmente
en el JSON, que lo validado por una persona se declare como tal y que toda
composición prestada diga de dónde salió.

**Vía de retorno de las decisiones.** `Revision_MAP_Grupo.md` lista las
**157 decisiones distintas** (por composición, o por UNS cuando la fila no
imprime composición). Para que lleguen al cálculo, copiar
`decisiones_map_grupo.plantilla.json` a `decisiones_map_grupo.json` y rellenarlo:
esas filas pasan al estado `VALIDADO POR INGENIERO`, **siempre separado de las
AUTO**. Una decisión **nunca sobreescribe** un grupo que el código sí asigna: el
choque se reporta y se conserva lo del código.

**Columna `Fila PRD`**: coincidencia literal del UNS contra las 99 filas de
`table_prd.json` que nombran los suyos (1 554 filas resueltas). Las 16 filas
restantes de PRD son categorías redactadas y encuadrar un material en ellas es
criterio de ingeniería.

**No reintroducir una heurística de composición.** La Rev. 2 asignaba grupo con
expresiones regulares y producía asignaciones *falsas*: `mo\b` rotulaba «acero de
baja aleación» a los austeníticos 16Cr-12Ni-2Mo, a los dúplex 22Cr-5Ni-3Mo-N y a
las aleaciones de níquel 62Ni-22Mo-15Cr; `8ni` casaba dentro de `18Ni`.
`test_build_db.py::TestMapeoDeGrupos` lo impide.

### TE-1 no reparte todo por Grupos numerados

Además de los Grupos 1 a 4 (que llegan por Nota), TE-1 publica dilatación en
**columnas que se autodescriben en su propio título**: `15Cr and 17Cr Steels`,
`12Cr, 12Cr–1Al, 13Cr, and 13Cr–4Ni Steels`, `27Cr Steels`, `8Ni and 9Ni Steels`,
`5Cr–1Mo and 29Cr–7Ni–2Mo–N Steels`, `5Ni–1/4Mo Steels`,
`9Cr–1Mo Steels (Including Grades 9, 91, 911, and 92)`. Son tan normativas como
las Notas y `columnas_nombradas_te1()` las resuelve **literalmente**, extrayendo
del título las designaciones que enumera. Cuando el título condiciona por
**grado** —el caso de los 9Cr-1Mo— la comprobación sigue siendo literal, porque
el grado está impreso en la propia fila.

### La única decisión tomada, y por qué

`decisiones_map_grupo.json` contiene **una**: `9Cr-1Mo-V` (Grado 91 / P91 / F91 /
T91, 17 filas) → `Material Group E` para el **módulo E**. TM-1 **no** lo asigna
literalmente —ninguna de sus Notas nombra ese material ni el UNS K90901— y se
aplica la Nota (5), *«9Cr–Mo, including variations thereof»*. Apoyos: Parte A
(K90901 es 9Cr-1Mo con V, Nb y N), Parte C SFA-5.5 §A7.2.3.1 (describe el
electrodo EB91 como *«a 9% Cr–1% Mo electrode modified with niobium and
vanadium»*) y la propia TE-1, que agrupa el Grado 91 con los 9Cr-1Mo.

**Solo afecta a E.** La dilatación la da la columna impresa de TE-1, que nombra
el Grado 91, y por eso `grupo_te` va vacío en la decisión.

Las otras **118** no admiten propuesta: su composición no figura en ninguna Nota
ni columna, así que II-D no publica E ni dilatación y **lo correcto es que sigan
bloqueadas**. Verificado que ninguna coincide con una nota con los elementos en
otro orden.

**Dos columnas de TE-1 no se resuelven solas y no deben forzarse:** las del
`17Cr–4Ni–4Cu`, partidas en `Condition 1075` y `Condition 1150` con valores
distintos. La tabla de materiales no imprime el tratamiento, así que esas 12
filas quedan sin dilatación y el `Motivo` explica dónde leerla a mano.

`validado_por` es **opcional**: se registra si está, pero no se exige. Lo que
separa una decisión de un dato del código es el **estado** de la fila
(`VALIDADO POR INGENIERO` frente a `AUTO (…)`) y su justificación, no que alguien
haya escrito un nombre en un JSON.

---

## Protocolo de verificación antes de entregar un cálculo

1. ¿Todas las entradas en SI? (psi → MPa, in → mm, °F → °C)
2. ¿Verificado el espesor remanente contra el mínimo admisible?
3. ¿Incluido el margen de corrosión para la vida remanente?
4. ¿Citado el artículo de ASME PCC-2 que soporta cada fórmula aplicada?
5. ¿Leídas las notas del material? (`Notas_Codigo`; restringen soldadura, PWHT y servicio)

## Código Python

Modular, con pruebas, documentado. Son herramientas de ingeniería críticas: cada
decisión no obvia va comentada explicando **por qué**, no qué hace.
