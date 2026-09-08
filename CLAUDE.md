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
               │  └─ CHAPTERS/tables/          Tablas del cuerpo normativo
               ├─ ASME PCC/pcc_2/              Artículos de PCC-2
               └─ ASME_BPVC/Sec_II/
                  ├─ bpvc_ii_a_1/, a_2/, b/, c/  Partes A, B y C: texto íntegro
                  │                              de 379 especificaciones
                  ├─ bpvc_ii_d_metric_2025/    II-D métrica (MPa, °C)
                  └─ bpvc_ii_d_customary_2025/ II-D U.S. Customary (ksi, °F)
outputs/       Entregables.
templates/     maestro_con_macros.xlsm — Rev. 0 + proyecto VBA, entrada del builder.
               Se genera con scripts/make_vba_seed.py; no se edita a mano.
```

**Nunca leas `outputs/` ni `templates/`** salvo que se te señale un archivo.
**Guarda todo entregable en `outputs/` dentro de una subcarpeta** con nombre de proyecto.

Ante una duda de alcance, pregunta antes de producir.

---

### `pdf_pages` no usa la misma base en todas las extracciones

Cuesta caro equivocarse aquí, y ya se pagó una vez: una «corrección» de 16 citas
de II-D que en realidad las rompió, revertida después.

| Fuente | Base de `pdf_pages` | Cómo se comprobó |
|---|---|---|
| BPVC **II-D** (métrica y US) | **1-based** inclusiva | El UNS `K01700` de la Tabla 1A está en la página 58, la primera que declara |
| BPVC **II A, B y C** | **0-based** inclusiva | `[56, 119]` de SA-6/SA-6M son las páginas 57 a 120; la 56 está en blanco |
| B31.3 (capítulos, apéndices, tablas) | 1-based | El rótulo de cada tabla cae en la página que declara |

**No se detecta el desplazamiento con el rótulo de la tabla**: se repite en cada
página de continuación, así que casa con las dos convenciones y no distingue
nada. Hay que usar algo que aparezca **una sola vez** — un UNS, o una fila
concreta. Ese fue exactamente el error.

`verificar_resources.py` audita PCC-2, los capítulos del B31.3, sus apéndices
D–Z y las dos ediciones de II-D. `verificar_seccion_ii.py` audita **solo** las
partes A, B y C: la II-D se le quitó para no mantener dos scripts con dos
convenciones.

### Sección II, partes A, B y C — cómo está troceada

379 especificaciones de material (SA-, SB-, SFA-), una por archivo en
`specifications/`, más los apéndices en `appendices/` y las figuras en
`figures/`. `index.json` es el catálogo —designación, título, páginas PDF y
folios impresos— y `meta.json` la procedencia.

**El contenido de cada tabla NO viene como `<table>`.** Los bloques `Table`
tienen el `html` vacío y las filas cuelgan de ellos como bloques `Table` →
`Line`, cada uno con su texto y su `bbox`. Una fila llega así:

```
/page/224/Table/27  →  "Carbon, max 0.25A 0.30B 0.35B"
                       "Manganese 0.27–0.93 0.29–1.06 0.29–1.06"
```

Es decir: el dato está entero, pero **construir una base exige recomponer las
columnas desde los `Line` y sus `bbox`**, no leer una tabla ya formada. Cuando
una fila reparte sus celdas en varios `Line` (Tabla 2 de SA-106), el `bbox` es
lo único que dice a qué columna va cada una.

`verificar_seccion_ii.py` audita las seis partes contra sus PDF: que cada
especificación empiece donde dice, que el folio impreso cuadre y que no haya
solapes ni huecos. Los PDF no se versionan; se pasan con `--pdfs`.

## Motor de cálculo — estado actual

Entregable vigente: `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm`
(43 hojas, **una sola visible**). Se **genera por script**, nunca se edita a mano.

**Es un libro con macros.** Al abrirlo se ve solo el `Dashboard`; la navegación a los
doce motores (Art. 212, los 10 buscadores, `Instrucciones`) la hace un proyecto VBA de
dos componentes. Los estados de visibilidad van **grabados en el archivo**, así que con
las macros bloqueadas no se expone ninguna base de datos.

### Los diez buscadores

| Hoja | Qué responde | Fuente |
|---|---|---|
| `Buscar_B31_3` | S admisible | B31.3 Apéndice A, A-1/A-4 y A-1C/A-4C |
| `Buscar_BPVC_IID` | S admisible, ferrosos | II-D Tabla 1A |
| `Buscar_BPVC_IID_B` | S admisible, no ferrosos y pernería | II-D Tablas 1B y 3 |
| `Buscar_Su` · `Buscar_Sy` | Su y Sy frente a T | II-D Tablas U y Y-1 |
| `Buscar_Prop_IID` | Módulo E, Poisson y densidad **por grupo** | II-D TM-1..5 y PRD |
| `Buscar_Prop_B31_3` | Dilatación y módulo, metales y no metálicos | B31.3 Apéndice C, C-1..C-4 |
| `Buscar_NoMetalicos` | Esfuerzo de diseño hidrostático y presión admisible | B31.3 Apéndice B |
| `Buscar_Ec_A2` | Factor de calidad de fundición **Ec** | B31.3 Tabla A-2 + 302.3.3-1 |
| `Buscar_Ej_A3` | Factor de calidad de junta longitudinal **Ej** | B31.3 Tabla A-3 |

**Un dato, un motor.** El Apéndice C salió de `Buscar_Propiedades` (que pasó a
llamarse `Buscar_Prop_IID`) y de `Buscar_NoMetalicos`, y tiene motor propio con
conmutador SI ↔ US. `Buscar_Prop_IID` **no expone la dilatación de la II-D**:
`DB_TE`/`DB_TEC` se siguen construyendo y auditando, pero TE-1 se indexa por
temperatura × columnas de grupo, no por material, y ningún motor la consulta
todavía. Es un pendiente declarado, no un olvido.

### Reconstruir

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts

# Solo la primera vez, o al tocar vba/*.vba.
# Activa "Confiar en el acceso al modelo de objetos de proyectos de VBA" unos
# segundos y lo restaura, verificando el resultado. Si avisa de que no pudo
# restaurarlo, desactivelo a mano en el Centro de confianza.
python make_vba_seed.py

# Solo si se repone la extraccion de II-D: notas de grupo de TM-1 / TE-1.
# Una corrida por edicion; nunca copiar las notas de una en la otra.
python extraer_notas_ii_d.py --edicion si --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Metric 2025\D Metric 2025 _p1201-p1500.pdf"
python extraer_notas_ii_d.py --edicion us --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Customary 2025\D Customary 2025 _p1201-p1500.pdf"

# Solo si se repone la extraccion de los Apendices B o C del B31.3. Los dos son
# idempotentes y NO tocan ningun valor: solo la capa de metadatos. El PDF del
# codigo y su OCR NO estan en el repo (copyright de ASME, ver .gitignore): el de
# la C hay que tenerlo en disco y pasarlo con --ocr si no esta en su sitio.
#   C: unidades de los coeficientes A y B de C-1/C-1C, miembros de sus Notas
#      (2)..(6) y grupo impreso de C-2/C-3.
#   B: celdas fusionadas de B-1, las 6 especificaciones de B-3, las columnas
#      minimo/maximo de B-4 y B-5, y el material a dos lineas de B-6.
#   A: la costura de las paginas enfrentadas — devuelve a la curva los 197
#      esfuerzos a 200 F del bloque de niquel de A-1C, separa Class/Description
#      en A-3 y Spec/Grade en A-1C, y limpia la elipsis de columna vacia.
python completar_apendice_c.py --resources ..\..\..\resources
python completar_apendice_b.py --resources ..\..\..\resources
python completar_apendice_a.py --resources ..\..\..\resources

# Solo si se repone la extraccion de las tablas del cuerpo del B31.3.
# Declara cual de los dos archivos de la Tabla 302.3.3-1 es el canonico y
# recupera sus rotulos de columna del para. 302.3.3(c). Idempotente, solo
# metadatos. El builder ABORTA si no se ha corrido.
python completar_tabla_302_3_3.py --resources ..\..\..\resources

python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm

# No entra en el libro: mide la reconstruccion de las tablas de la Seccion II
# partes A/B/C y escribe Revision_Tablas_SecII.md. Ver la seccion de abajo.
python secii_tablas.py --resources ..\..\..\resources `
    --informe ..\Revision_Tablas_SecII.md
```

La entrada del builder es el maestro sembrado en `templates/`, que a su vez sale de
`outputs/Base_Datos_Materiales_ASME/Motor_de_Calculo_ASME_PCC_Rev0_respaldo.xlsx`
(3 hojas). El `Motor_de_Calculo_ASME_PCC.xlsx` de la raíz **no está en el repo**.

`verificar.py` devuelve 0 solo si todo pasa. Audita **fila a fila** cada valor
tabulado contra el JSON del código (271 276 valores de esfuerzos, más 4 726 del
Apéndice C y 743 de los factores de calidad), la contigüidad de la cascada,
la ausencia de fórmulas de matriz dinámica, la interpolación recalculada en hoja, el
caso semilla y la capa de navegación. **Requiere Excel instalado**: recalcula con el
motor real, no con LibreOffice. **Ejecútalo siempre después de tocar el builder.**

Sus §6b y §6c **recalculan en Excel la misma expresión que lleva el motor**, no una
copia: las funciones que la generan (`formula_estado_apxc`, `formula_valor_apxc`,
`formula_factor_aplicable`) viven en `build_db_materiales.py` y las emiten los dos.
Si alguien cambia la lógica en el motor, la prueba la ejerce cambiada; si cambia solo
el texto de un estado, el literal se lee de la misma constante y no puede divergir.

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
10. **Conmutador de unidades en todo motor.** Todo motor de búsqueda y todo motor
    de cálculo del libro lleva un conmutador **SI ↔ US**. Es corolario de la
    regla 9: los dos sistemas se leen de la edición correspondiente del código,
    nunca por conversión. Cuando el código publica los dos sistemas en la **misma
    tabla** —C-2 y C-4 del B31.3— el conmutador cambia la **columna** leída;
    cuando publica **una tabla por edición** —A-1/A-1C, C-1/C-1C, C-3/C-3C, II-D
    métrica y US— cambia la **hoja**. Si un material no tiene homólogo en la otra
    edición, el motor muestra «sin equivalente en la edicion US», nunca un valor
    convertido.

    **Excepción declarada (Rev. 4): los factores adimensionales.** `Ec` (Tabla A-2)
    y `Ej` (Tabla A-3) no tienen unidad y el código publica **una sola tabla** para
    los dos sistemas: no existen A-2C ni A-3C. Ahí `Buscar_Ec_A2` y `Buscar_Ej_A3`
    ponen una **celda fija y rotulada** —«FACTOR ADIMENSIONAL — identico en SI y en
    US (el codigo publica una sola tabla)»— en el sitio donde los demás ponen el
    conmutador. Se cumple el espíritu de la regla —el usuario ve siempre en qué
    sistema lee y nunca un valor convertido— sin fabricar un interruptor que no
    gobierna nada. `TXT_ADIMENSIONAL` en el builder, con su prueba.

11. **El factor publicado es un MÍNIMO, y el motor tiene que decirlo.** Las Notas
    (4) y (5) de la Tabla A-2 dicen cosas opuestas con el mismo formato: la (4)
    —*«can be enhanced by supplementary examination»*— permite subir el factor con
    la Tabla 302.3.3-1; la (5) —*«applicable only when proper supplementary
    examination has been performed»*— avisa de que el factor **ya supone** ese
    examen. Confundirlas mueve el espesor requerido. `Buscar_Ec_A2` deriva el
    estado de las notas que cita **la propia fila**, nunca de una suposición, y
    contempla los cuatro casos: solo (4), solo (5), **las dos** —A451 es la única
    fila que las cita juntas— y ninguna. Para `Ej` el mecanismo existe
    (para. 302.3.4(b) y Tabla 302.3.4-1) pero la extracción de esa tabla está
    inservible, así que el motor **declara el hueco y no ofrece número**.

### Estilo de diseño de los buscadores — no romper

Vive en `build_buscador` / `finish_buscador` (los 5 buscadores de cascada: B31_3,
BPVC_IID, BPVC_IID_B, Su, Sy) y se replica en `build_buscador_grupo`. Tocar el
estilo de un buscador significa tocar las tres constantes de módulo
(`TEMP_INPUT_FILL`, `SEL_OK_FILL`/`SEL_OK_FONT`, `SEL_BAD_FILL`/`SEL_BAD_FONT`,
declaradas junto a `CARD_FILL`/`KPI_FILL`) para que cambien a la vez en todos.

1. **Celda única de escritura (temperatura de consulta): relleno propio
   `TEMP_INPUT_FILL` (lavanda, `CCC0DA`)**, distinto del amarillo `IN_FILL` de las
   celdas de lista desplegable. Es la única celda que se teclea en todo el
   buscador; su color debe diferenciarla a simple vista de las que solo aceptan
   una lista. Aplica también a `build_buscador_grupo`.
2. **Semáforo de cascada completa/incompleta**, exclusivo de los 5 buscadores de
   cascada (`build_buscador`/`finish_buscador`): la celda de aviso junto a la
   temperatura (`G12:I12`) lleva formato condicional — verde `SEL_OK_FILL`
   (`C6EFCE`/`006100`) con "SELECCION COMPLETA" cuando la cascada (pasos 0 a 4)
   está resuelta; amarillo `SEL_BAD_FILL` (`FFEB9C`/`9C6500`) con "SELECCION
   INCOMPLETA" mientras falte un paso. El texto va en mayúsculas y sin tilde
   ("SELECCION", no "SELECCIÓN"): todo el texto de celda del libro evita acentos
   (ver "SELECCION DEL MATERIAL", "Composicion nominal") para no arrastrar
   problemas de codificación fuera de Excel 365. `build_buscador_grupo` no lleva
   este semáforo: no tiene una única cascada que completar, cada bloque ya avisa
   "(elija grupo)" en su propia celda de valor.
3. **Curva del material: línea continua, sin marcadores, del mismo azul de la
   banda del buscador (`BLUE = 2F5597`)** — `ch.scatterStyle = "line"`, serie del
   valor tabulado con `marker="none"` y `smooth=False` (la interpolación del
   código es lineal — regla 3 de esta lista arriba — una curva suavizada la
   representaría mal). El punto consultado sigue siendo un rombo rojo (`FF0000`)
   sin línea, para distinguir el valor puntual de la curva. Aplica a
   `finish_buscador` y a `build_buscador_grupo`.

### `DB_B31_C` / `DB_B31_CC` — el Apéndice C entero en una base por edición

201 filas por edición: C-1/C-1C (52), C-2 (44), C-3/C-3C (75) y C-4 (30). El nivel
0 de la cascada es la **propiedad**, no la tabla: el ingeniero pregunta «quiero el
módulo E», no «quiero la Tabla C-3».

Tres cosas que este motor hace distinto de los cinco de esfuerzos, y por qué:

- **Bloquea en los DOS extremos.** El Apéndice C no publica columna «Temp. máx.»:
  el límite es el primer y el último punto que tabula la propia fila. Sostener el
  último valor por encima —que es lo correcto en los buscadores de esfuerzos,
  donde el tope lo pone una columna del código— aquí sería extrapolar.
- **Rama de dato puntual.** C-2 y C-4 no dependen de la temperatura: publican un
  valor único, la gráfica queda vacía **a propósito** y el estado dice VALOR ÚNICO.
- **El conmutador cambia de columna en dos tablas y de hoja en las otras dos**
  (regla 10). Nunca convierte.

El **factor de escala se lee del JSON**, jamás se codifica: confundir el ×10³ de
C-3 con el ×10⁶ de C-3C son tres órdenes de magnitud en el módulo E. Va en columna
propia junto con su texto impreso, y el KPI muestra el valor con el factor aplicado
mientras la sección 4 muestra el impreso sin él.

**El enlace SI↔US es posicional** (`clave_bi = propiedad#fila impresa`), porque los
nombres divergen entre ediciones por artefactos de impresión: «Type 309.» con punto
en C-3 y con coma en C-3C, «25Cr–20Ni» con raya en una y guion en la otra. Un enlace
posicional sin red es una bomba silenciosa, así que el build **aborta** si el número
de filas por propiedad no coincide o si un nombre normalizado difiere fuera de
`DIVERGENCIAS_NOMBRE_C`, que hoy tiene **una sola entrada**.

### Sección II partes A, B y C — medida, todavía no volcada al libro

`secii_tablas.py` reconstruye las tablas de las 368 especificaciones desde los
bloques `Line` y sus `bbox`. **No escribe nada** en `resources/` ni en el libro:
mide y reporta a `Revision_Tablas_SecII.md`.

El JSON no trae tablas: el `html` de todo bloque `Table` es `<p></p>` y los `Span`
no se conservan, así que el bloque más fino es `Line`. El reparto de cada fila
tiene tres resultados y **ninguno adivina**: `EXACTA` (≥2 `Line`, cada uno a su
columna por el punto medio de su `bbox`), `POR CONTEO` (un solo `Line`, partido
desde la derecha y solo si el conteo cuadra con las columnas que la tabla ya
demostró tener) y `AMBIGUA` (el texto se conserva **entero** en una celda y se
declara). Sin `Span`, repartir por interpolación sobre el ancho del `bbox` sería
inventar estructura con una fuente proporcional: la geometría **confirma** un
reparto, nunca lo produce.

La garantía que sí se puede dar sin el PDF es la **comprobación sin pérdida**: la
concatenación de las celdas de cada fila coincide carácter a carácter con la de sus
`Line` de origen. Pasa sobre las 54 198 filas y los 124 115 `Line`.

**Estado: parado en el punto de decisión de la Fase 2 del plan.** Medido sobre las
cuatro partes: 2 572 tablas lógicas, 54 198 filas, 5 233 notas al pie, y un
**38,7 % de las filas de dato en AMBIGUA**. El plan dice que a esa escala escribir
las filas dudosas es peor que no escribirlas, así que las nueve hojas
(`CAT_SecII`, `IDX_SecII_Tablas`, los cuatro `DB_SecII_*`, `DB_SecII_Notas` y las
normalizadas de química y tracción) **no están en el libro**. Seguir a las Fases 3-5
—o acotarlas a las tablas que sí se reparten— es decisión del ingeniero.

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
