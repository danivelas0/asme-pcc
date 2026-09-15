---
name: motor_pcc2
description: Úsala para construir el motor de cálculo de un artículo de ASME PCC-2 en este repositorio — barre el artículo entero en resources/, declara el motor con el chasis, y genera la hoja, sus dos pestañas, la navegación y las pruebas. Es autónoma de extremo a extremo. Se dispara con "motor del articulo", "nuevo motor PCC-2", "/motor_pcc2 <n>".
---

# Skill `motor_pcc2`

Construye el motor de cálculo de un artículo de ASME PCC-2 que **todavía no está
cargado** en `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm`. El Art. 212 y el Art. 206
ya existen, están escritos a mano y son lo único del libro verificado celda a celda
contra un *oracle*: **esta skill no los toca** — los usa como referencia del marco.

**El flujo es de un solo paso: el ingeniero nombra el artículo y se entrega el motor
funcionando y dando resultados coherentes.** No hay paradas intermedias, no se pide
que nadie marque qué entra, y no se le pregunta nada al ingeniero. Todo el contenido
—instrucciones de cálculo, variables, condiciones, umbrales, especificaciones— está
en el JSON del artículo, dentro de `resources/`, y de ahí sale.

## El contrato: qué es MARCO y qué es CONTENIDO

La skill es una **plantilla**. Fija el marco y el procedimiento para que todos los
motores tengan el mismo flujo; el contenido de cada artículo es suyo. Esta frontera es
lo que impide que dentro de tres artículos alguien intente generalizar una ecuación.

### Lo que la plantilla garantiza igual en TODOS los motores

El chasis lo impone, y **aborta el build si falta**:

- **El orden de lectura de la hoja**: entradas → material → cálculo → verificaciones →
  dictamen → pasos del flujo. La hoja se lee en el orden en que se rellena.
- **El sistema visual** (Swiss Industrial Print) y la **leyenda derivada del estado
  real de la celda** — amarillo = lo rellena el ingeniero, gris = lo calcula el libro,
  papel = rótulo. No se etiqueta celda a celda: se deriva, que es la única forma de
  que la leyenda impresa no pueda mentir.
- **El semáforo escrito en `fgColor` Y `bgColor`** — en un formato diferencial Excel
  pinta con `bgColor`, y escribir solo `fgColor` deja la regla disparando sin pintar.
- **La trazabilidad**: toda celda de cálculo cita archivo, bloque y cláusula, y el
  build aborta si falta, si el bloque no existe, o si la cita apunta a un
  `section_header` (un rótulo no publica un valor).
- **El material siempre de una base de datos**, por cascada de cinco niveles y
  desplegable. Nunca tecleado, nunca de una lista fija.
- **Las unidades visibles**, el **conmutador SI↔US** de todo el motor, los **dos
  decimales** en toda celda de valor con magnitud, y el **comentario solo en la
  columna de valor**.
- **El botón de reinicio y la cascada derivados de la propia hoja**, sin una sola
  dirección de celda escrita en el VBA.
- **La navegación**: el artículo es un nivel del árbol con sus tres documentos.
- **El área de impresión** declarada y la **fila de título repetida** en cada página.

### Lo que es libre en cada artículo, y el chasis no toca ni adivina

- Las **ecuaciones** y de qué bloque salen.
- Las **variables** y sus nombres.
- **Qué se teclea y qué se calcula.**
- Las **condiciones y los límites**.
- Los **umbrales** y su ancla en el texto.
- Las **verificaciones** y su criterio.
- **Qué entra al dictamen** (la *forma* del dictamen es marco; qué entra, no).
- Los **pasos del flujo**.
- Las **especificaciones técnicas**.
- **Uno o varios casos de presión**: el chasis emite **una** columna de valor y aborta
  si se declaran dos. El reparto por casos no es genérico y se resuelve artículo por
  artículo, con el artículo delante — nunca por adelantado.

**Regla de oro:** la plantilla fija el MARCO y el PROCEDIMIENTO; el contenido de cada
artículo lo pone su propia declaración. **Donde el marco no dé para lo que un artículo
necesita, el chasis para y lo dice**, en vez de producir una hoja plausible pero
incompleta.

## Lo que esta skill NO hace

- **No inventa un valor normativo.** Ni uno (Regla nº 1 del proyecto). Todo sale de
  `resources/`; nunca de la memoria del modelo y nunca de un PDF externo.
- **No toca `build_parche_art212` ni `build_collar_art206`.** Son lo único verificado
  celda a celda contra un *oracle* y contra `verificar.py` §6e/§6f.
- **No edita el `.xlsm` a mano.** El libro se genera por script y sigue siendo
  reproducible: el motor nuevo entra por el builder, no por una inyección posterior.
- **No prueba que "el número es el que pide el código".** Ver «El límite de esta
  skill» más abajo.

## Cómo está montado el mecanismo (léelo antes de tocar nada)

Esto ya existe y está verificado — la skill lo **usa**, no lo reinventa:

- `outputs/Base_Datos_Materiales_ASME/scripts/motor_declarado.py` — el chasis. Tipos
  `Cita`, `Fila`, `Lista`, `Seccion`, `Paso`, `Especificacion`, `Material`,
  `Verificacion`, `Dictamen`, `Motor`, `Helpers`; funciones `resolver_direcciones()`,
  `sustituir_nombres()`, `comprobar_procedencia()`, `comprobar_nombres_declarativos()`,
  `comprobar_casos()`, `comprobar_listas()`, `comprobar_dictamen()`,
  `formula_dictamen()`, `emitir_tabla()`.
- `build_db_materiales.py` — `build_motor_declarado()`, `build_documentos_declarados()`,
  `direccion_de()`, `_helpers_del_libro()`, `hojas_de_motor()`, `UMBRALES_PCC2` /
  `leer_umbrales_pcc2()` / `umbral()`, y los bucles `for motor in MOTORES_DECLARADOS`
  que llaman a las dos primeras y a `_nodo_de_articulo(m)` dentro del `ARBOL`.
- `motores/__init__.py` — el **registro**. Declarar un artículo es escribir
  `motores/art_XXX.py` con su `Motor(...)` y añadirlo a esa tupla.
- `verificar.py` §12 y `test_dashboard.py::TestMotoresDeclarados` /
  `TestNodoDeArticuloDerivado` — **genéricos**: recorren `MOTORES_DECLARADOS` sin que
  nadie tenga que escribir una clase nueva por artículo.

**Consecuencia que importa:** una vez que un `Motor` está en `MOTORES_DECLARADOS`, las
fases 3 a 7 de abajo **no piden código nuevo** — ya corren solas la próxima vez que se
regenera el libro y se ejecutan las pruebas. Verifica esto releyendo el código citado
arriba antes de decir "voy a escribir la fase 4": probablemente ya está escrita y solo
hay que declarar el motor.

## Las ocho fases (0 a 7), todas automáticas

| Fase | Qué hace |
|---|---|
| **0. Localizar** | A partir del número de artículo, encuentra `art_XXX.json` en los **dos espejos** del árbol de `resources/asme_pcc/pcc_2/` (`part_N_.../article_XXX_.../` y `pN_.../art_XXX_.../`). Aborta si no está en ninguno de los dos. |
| **1. Barrer el artículo entero** | Lee el JSON **completo** —todos los `blocks`— y levanta el inventario de todo lo que el artículo publica, con el número de bloque al lado de cada hallazgo. Ver abajo: la cobertura es criterio de aceptación. |
| **2. Declarar** | Escribe `motores/art_XXX.py`: un `Motor(...)` con sus `Seccion`, `Fila`, `Lista`, `Verificacion`, `Dictamen`, `Paso`, `Especificacion`, cada elemento de cálculo con su `Cita`. Registra los umbrales del artículo en `UMBRALES_PCC2`. Añade el motor a `MOTORES_DECLARADOS`. |
| **3. Chasis** | `build_motor_declarado()` construye la hoja: material, listas, unidades, dictamen, semáforo, reinicio, cascada, leyenda, decimales, impresión. **Ya está escrito y se llama solo** para todo motor del registro. |
| **4. Navegar** | `_nodo_de_articulo()` deriva el nodo del `ARBOL` (artículo = nivel con motor + especificaciones + instrucciones). **Ya está escrito.** Lo que **no** se deriva son las tres líneas de `HojasNavegables()` en el VBA — ver «Definición de terminado». |
| **5. Pestañas** | `build_documentos_declarados()` arma las especificaciones técnicas (si `motor.especificaciones` no está vacío) y las instrucciones de uso (siempre), derivadas de la declaración. **Ya está escrito.** |
| **6. Pruebas** | `test_dashboard.py::TestMotoresDeclarados` y `TestNodoDeArticuloDerivado` ya recorren `MOTORES_DECLARADOS` de forma genérica. **No se escribe una clase nueva por artículo.** |
| **7. Verificar y entregar** | Corre `pytest` y `verificar.py --resources ... --wb ...`. La §12 recalcula el motor nuevo en Excel real, exige cero errores de Excel, repite la comprobación de procedencia, comprueba que todo veredicto pinta y que el dictamen no contradice a sus verificaciones. |

## La fase 1 es un barrido exhaustivo, no una selección

**No se escogen ecuaciones.** Lo que se está construyendo es la **automatización de
cálculo del artículo, completa**: todo lo que el artículo publica y pertenece a un
motor de cálculo entra, y acaba **en algún sitio** — en una fila del motor, en una
verificación, en un paso del flujo o en la pestaña de especificaciones técnicas.

Qué clase de cosa se busca en el JSON de **cualquier** artículo, y cómo se localiza:

| Qué se busca | Cómo se localiza en el JSON |
|---|---|
| **Ecuaciones** | Bloques de tipo `equation`. Si el `text` llega vacío o inservible, el bloque (o un `figure` contiguo) trae `image: fig/….png`: esa imagen **es parte de `resources/`** y se lee con la herramienta `Read` sobre el PNG. De ahí sale la fórmula. |
| **Figuras con dato** | Bloques `figure` cuya imagen publica un valor, una geometría acotada o una condición — no solo ilustración. Se leen igual. |
| **Umbrales de doble unidad** | Texto con el patrón `<n> mm (<n> in.)`. Las dos cifras se guardan **tal cual**; nunca se convierte una en la otra. Se registran en `UMBRALES_PCC2` (ver abajo). |
| **Límites y prohibiciones** | Párrafos con `shall not`, `maximum`, `minimum`, `not permitted`, rangos de temperatura, de espesor o de servicio. Acaban en compuertas del dictamen o en verificaciones. |
| **Requisitos de fabricación, examen y prueba** | Las secciones de procedimiento del artículo. Acaban en pasos del flujo o en la pestaña de especificaciones técnicas. |
| **Tablas** | Bloques `Table` y sus filas. Se declara qué publica cada una y a qué fila del motor alimenta. |

### La cobertura es criterio de aceptación

Cada bloque relevante del artículo termina en un sitio concreto del motor, y ese mapeo
se puede recorrer. **Lo que no se pueda implementar no se omite en silencio: se
declara explícitamente, con su cláusula y su número de bloque.** Es el patrón que este
proyecto ya usa para sus huecos —la Tabla TE-2 fuera por límite de la fuente, las
Tablas B-2 a B-6 retiradas por alcance—: el hueco se escribe, no se calla. Un hueco
declarado va en el comentario de la celda que le corresponda, en la pestaña de
especificaciones, o en `ISSUES` del build; nunca solo en la cabeza de quien declaró.

### Dos casos que se confunden fácil, y se resuelven al revés

Cuando un dato no aparece donde debería, hay **dos** situaciones distintas. Confundirlas
es el error caro: una se repara y la otra se declara.

#### Caso A — extracción defectuosa: el código SÍ publica el dato, el JSON lo trae roto

Una tabla comprimida dentro de un encabezado, una ecuación con símbolos revueltos por
una costura de página, un `text` vacío. **Siempre se repara. No se para y no se
pregunta.** En este orden:

1. **Antes de llamarlo rotura, mirar si el dato está como imagen.** Un bloque
   `figure`/`equation` con `image: fig/….png` **es parte de `resources/`**: se lee
   con la herramienta `Read` sobre el PNG y de ahí sale la fórmula. Eso no es una
   rotura — es cómo llegó ese dato.
2. **Si de verdad falta, se abre el PDF del código** en
   `C:\Users\dvelasquez\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES`.
   Los PDF **nunca se versionan** (copyright de ASME; está en `.gitignore`).
3. **El PDF nunca alimenta un motor directamente.** El dato pasa por un **script de
   extracción/corrección** que lo escribe en el JSON de `resources/` con su
   `extraction_amendments` o metadato de procedencia (script, fecha, SHA-256 de la
   imagen o del folio fuente), y **solo entonces** ese JSON alimenta el builder. Es la
   Regla nº 1, y hay patrón de sobra que seguir: `completar_apendice_a.py`,
   `completar_apendice_b.py`, `completar_apendice_c.py`, `completar_tabla_302_3_4.py`,
   `completar_app_501_energia.py`, `completar_art_212_eq2.py`,
   `completar_art_206_filete.py`.
4. **Y se sigue con el motor.**

Lo que la Regla nº 1 prohíbe es **inventar** el dato desde la memoria del modelo o
desde un documento derivado. Leerlo del PDF del código y escribirlo en `resources/` por
script es exactamente lo contrario, y este proyecto ya lo ha hecho siete veces.

#### Caso B — el código NO publica el dato

Ningún PDF lo va a traer: no existe. Se **declara** el hueco con su cláusula y su
bloque, y lo que depende de él queda bloqueado — que es el resultado correcto, no un
pendiente. El precedente del proyecto es la Tabla TE-2, que no imprime rótulo de grupo:
cerrarla exigiría una convención inventada fuera del texto normativo, así que queda
cerrada como límite de la fuente.

## Los umbrales: un paso obligatorio por artículo

Un umbral normativo es un valor del código, y **no se teclea en la declaración**. Un
`formula="=100"` lo saca de la memoria de quien declaró: cumple la Regla nº 1 en la
cita y la incumple en el número. Peor: en modo US esa cifra métrica no se convierte en
nada, y **un umbral de aceptación en la unidad equivocada convierte un "NO CUMPLE" en
un "CUMPLE"**.

El mecanismo existe y es por artículo. Registrarlo es **obligatorio en cada artículo
que se declare**:

1. Localizar en el JSON del artículo los umbrales que el código imprime en **dos
   unidades**, y anotar el fragmento de texto que ancla a cada uno de forma unívoca.
2. Registrar cada umbral en `UMBRALES_PCC2` (`build_db_materiales.py`) como
   `clave: (artículo, ancla)`. `leer_umbrales_pcc2()` lee las **dos cifras impresas**
   —nunca convierte una en la otra— y el build **aborta** si el ancla deja de
   aparecer. La ruta del JSON la resuelve `_json_de_articulo()` desde `Motor.fuente`,
   así que un artículo declarado no pide tocar nada más.
3. En la hoja, el umbral entra con `umbral(umbrales, clave, es_si)`, que emite
   `IF(es_SI, métrico, US)`.

El diccionario ya trae entradas de los dos motores escritos a mano: se miran como
ejemplo de **la forma de una entrada**, nunca como el procedimiento.

**Los textos de aviso no citan la cifra.** Un aviso que dice "excede 40 mm" es falso
en modo US, y un aviso que miente sobre su propio umbral es peor que ninguno.

## Trazabilidad: la Regla nº 1, operacionalizada

Que la skill "lea `resources/`" no es una promesa, es un mecanismo que ya existe en el
chasis y que no se puede sortear:

1. **Toda `Fila`, `Verificacion` y `Especificacion` de cálculo lleva su `Cita`**
   (`archivo`, `bloque`, `clausula`). Una `Fila` con `tipo=FORMULA` y `cita=None` sí se
   puede construir como objeto; lo que no se puede es **construir el libro** con ella:
   `comprobar_procedencia()` aborta en tiempo de build, citando la fila y la hoja.
2. Esa procedencia se escribe **tres veces y en tres formas**, y las tres las escribe
   el chasis, no la declaración:
   - la **columna de referencia** que ve el ingeniero (`f.cita.clausula`, columna G);
   - el **comentario de la celda** de valor (solo en columna de valor, nunca en la de
     rótulo);
   - la **tabla de trazabilidad** que devuelve `comprobar_procedencia()`.
3. **El build aborta** si una celda de cálculo llega sin procedencia
   (`comprobar_procedencia` es la primera línea de `build_motor_declarado`, antes de
   escribir una sola celda), si una cita apunta a un bloque que no existe, y si apunta
   a un bloque de tipo `section_header` — un rótulo no publica un valor, y el bloque
   existe, así que sin ese guardia la cita parecía válida apuntando a donde el dato no
   está.
4. `verificar.py` §12 repite la comprobación contra el libro ya construido —el mismo
   guardia, no una copia— y añade lo que solo Excel puede decir.

**Nota sobre `Cita.archivo`:** es texto para que el ingeniero vea de un vistazo de qué
documento sale un dato, pero la ruta que de verdad valida `comprobar_procedencia()` es
siempre `motor.fuente` (un solo JSON por motor). Todas las `Cita` de un mismo `Motor`
describen bloques de ese mismo archivo.

## El dictamen lo compone el chasis

La **forma** del dictamen es idéntica en todo motor, así que es marco: las compuertas
bloquean primero y su propio texto pasa a ser el dictamen; después el material —sin
material resuelto no hay APTO—; al final el AND de las verificaciones.

Lo que se declara es **qué entra**: `Dictamen.compuertas` y `Dictamen.verificaciones`.
La fila reservada `dictamen` va **sin fórmula** —si trae una, el chasis aborta—, toda
clave citada tiene que estar declarada como fila **y** como `Verificacion` (de ahí
salen los textos que no bloquean), y las que entran al AND publican `"CUMPLE"` y solo
eso: es el literal contra el que `verificar.py` §12 juzga si el dictamen se contradice.

## Definición de terminado, por artículo

Un motor está terminado cuando existen las cinco cosas:

1. **La hoja del motor construida por el builder**, con cascada de material contra las
   bases, conmutador SI/US, leyenda, semáforo, botón de reinicio y área de impresión.
2. **Sus dos pestañas**: especificaciones técnicas e instrucciones de uso.
3. **Su nivel en el árbol de navegación** y **sus tres líneas en `HojasNavegables()`
   del VBA** (`scripts/vba/mod_nav.vba`). Este paso es **obligatorio y explícito**:
   esa lista es la única del libro que **no se deriva, a propósito** — es el guardia
   que hace fallar `TestSincroniaPythonVba` si alguien olvida registrar una hoja. Tras
   tocarla hay que volver a correr `make_vba_seed.py` y reconstruir.
4. **`pytest` en verde.**
5. **`verificar.py` en verde**, con la §12 mordiendo sobre ese motor.

## Qué prueba qué — y el límite declarado de la skill

| Nivel | Qué garantiza |
|---|---|
| `pytest` sobre el chasis (`test_motor_declarado.py`) | que las funciones del chasis hacen lo que dicen, con motores sintéticos, y que sus guardias abortan |
| `pytest` sobre el libro (`test_dashboard.py::TestMotoresDeclarados`, `TestNodoDeArticuloDerivado`, `TestChasisDeclarado`) | que la hoja de CADA motor declarado tiene lo que su declaración promete |
| `verificar.py` §12 | que el motor **recalcula sin errores de Excel**, que **toda cita existe** en el JSON, que **todo veredicto pinta** el color que promete, y que el **dictamen no contradice** a sus propias verificaciones |
| `verificar.py` **por artículo** (el equivalente a §6e/§6f del 212/206) | que **el número que da la hoja es el que pide el código**. **NO lo genera esta skill** |

Esa última fila es el límite declarado, y hay que decirlo sin adornos: **un chasis
genérico puede probar la maquinaria, no la ingeniería.** Que el semáforo pinte verde,
que el dictamen no se contradiga y que la cita exista de verdad en el JSON no demuestra
que la fórmula citada sea la correcta para ese caso de carga. Escribir esa red sería el
modelo comprobándose a sí mismo con su propia lectura del código.

## Referencias

- `referencias/reglas_del_libro.md` — el catálogo de reglas de diseño del libro que la
  declaración tiene que respetar: las 14 reglas, la Regla nº 1, las tres reglas del
  flujo de la cascada, el sistema visual, el semáforo, los comentarios, el ancla del
  VML, los dos decimales, y las trampas de openpyxl ya pagadas.
- `referencias/plantilla_declaracion.py` — la **forma** de una declaración `Motor(...)`,
  comentada y ejecutable. Sus valores son de un artículo real solo porque las `Cita` se
  validan contra un JSON de verdad: son ilustración, no procedimiento. Se copia la
  forma, nunca los valores.
