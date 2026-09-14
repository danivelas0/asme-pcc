---
name: motor_pcc2
description: Use when building a calculation engine for an ASME PCC-2 article in this repo — reads the article from resources/, stops for the engineer to decide which equations govern, then generates the declaration, the sheet, the tabs, the tests and the verification. Triggers on "motor del articulo", "nuevo motor PCC-2", "/motor_pcc2 <n>".
---

# Skill `motor_pcc2`

Construye el motor de cálculo de un artículo de ASME PCC-2 que **todavía no está
cargado** en `outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm`. El Art. 212 y el Art. 206
ya existen, están escritos a mano y son lo único del libro verificado celda a celda
contra un *oracle*: **esta skill no los toca**.

Autoridad de esta skill: `outputs/plans/diseno_skill_motor_pcc2.md`. Si algo de este
archivo y ese diseño llegan a discrepar, manda el diseño — actualiza este archivo, no
al revés.

## Lo que esta skill NO hace

- **No decide la ingeniería.** Qué ecuación gobierna el espesor, qué se verifica y qué
  se teclea es criterio del ingeniero. La skill lee `resources/`, propone y **se
  detiene**.
- **No inventa un valor normativo.** Ni uno (Regla nº 1 del proyecto).
- **No toca `build_parche_art212` ni `build_collar_art206`.** Son lo único verificado
  celda a celda contra un *oracle* y contra `verificar.py` §6e/§6f.
- **No edita el `.xlsm` a mano.** El libro se genera por script y sigue siendo
  reproducible: el motor nuevo entra por el builder, no por una inyección posterior.
- **No prueba que "el número es el que pide el código".** Ver «El límite de esta
  skill» más abajo.

## Cómo está montado el mecanismo (léelo antes de tocar nada)

Esto ya existe y está verificado — la skill lo **usa**, no lo reinventa:

- `outputs/Base_Datos_Materiales_ASME/scripts/motor_declarado.py` — el chasis. Tipos
  `Cita`, `Fila`, `Seccion`, `Paso`, `Especificacion`, `Material`, `Verificacion`,
  `Dictamen`, `Motor`, `Helpers`; funciones `resolver_direcciones()`,
  `sustituir_nombres()`, `comprobar_procedencia()`, `emitir_tabla()`.
- `build_db_materiales.py` — `build_motor_declarado()`, `build_documentos_declarados()`,
  `direccion_de()`, y los bucles `for motor in MOTORES_DECLARADOS` que llaman a las dos
  primeras y a `_nodo_de_articulo(m)` dentro del `ARBOL` de navegación.
- `motores/__init__.py` — el **registro**: hoy `MOTORES_DECLARADOS = ()`. Declarar un
  artículo es escribir `motores/art_XXX.py` con su `Motor(...)` y añadirlo a esa tupla.
- `verificar.py` §12 y `test_dashboard.py::TestMotoresDeclarados` /
  `TestNodoDeArticuloDerivado` — **genéricos**: recorren `MOTORES_DECLARADOS` sin que
  nadie tenga que escribir una clase nueva por artículo.

**Consecuencia que importa:** una vez que un `Motor` está en `MOTORES_DECLARADOS`, las
Fases 3 a 7 de abajo **no piden código nuevo** — ya corren solas la próxima vez que se
regenera el libro y se ejecutan las pruebas, porque los bucles y las clases genéricas ya
están escritos y esperando. Verifica esto releyendo el código citado arriba antes de
decir "voy a escribir la Fase 4" — probablemente ya está escrita y solo hay que
declarar el motor.

## Las nueve fases (0 a 8)

| Fase | Qué hace | ¿Para o sigue? |
|---|---|---|
| **0. Localizar** | A partir del número de artículo, encuentra `art_XXX.json` en los **dos espejos** del árbol de `resources/ASME PCC/pcc_2/` (`part_N_.../article_XXX_.../` y `pN_.../art_XXX_.../`). Aborta si no está en ninguno de los dos. | Sigue |
| **1. Leer y proponer** | Lee el JSON **entero** (todos los `blocks`) y arma el inventario (ver formato abajo): ecuaciones con su número de bloque, figuras con imagen, umbrales de doble unidad, límites y prohibiciones, requisitos de fabricación/examen/prueba, tablas. Propone entradas, cálculos, verificaciones y pasos. **Presenta el inventario y la propuesta al ingeniero.** | **PARA — SIEMPRE** |
| **2. Declarar** | Con lo que el ingeniero marcó (y solo eso), escribe `motores/art_XXX.py`: un `Motor(...)` con sus `Seccion`, `Fila`, `Verificacion`, `Dictamen`, `Paso`, `Especificacion`, cada `Fila`/`Verificacion`/`Especificacion` de cálculo con su `Cita`. Añade el motor a `MOTORES_DECLARADOS` en `motores/__init__.py`. | Sigue |
| **3. Chasis** | `build_motor_declarado()` construye la hoja: material, unidades, semáforo, reinicio, cascada, leyenda, decimales, impresión. **Ya está escrito y se llama solo** para todo motor en el registro — no hay nada que programar aquí. | Sigue |
| **4. Navegar** | `_nodo_de_articulo()` deriva el nodo del `ARBOL` (artículo = nivel con motor + especificaciones + instrucciones) para todo motor del registro. **Ya está escrito.** | Sigue |
| **5. Pestañas** | `build_documentos_declarados()` arma las especificaciones técnicas (si `motor.especificaciones` no está vacío) y las instrucciones de uso (siempre), derivadas de la declaración. **Ya está escrito.** | Sigue |
| **6. Pruebas** | `test_dashboard.py::TestMotoresDeclarados` y `TestNodoDeArticuloDerivado` ya recorren `MOTORES_DECLARADOS` de forma genérica: hoja en el libro y en la navegación, cada entrada editable y en el manifiesto de reinicio, comentario en columna de valor. **No se escribe una clase nueva por artículo** — eso sería repetir lo que la clase genérica ya cubre. | Sigue |
| **7. Verificar** | `verificar.py` §12 ya recorre `MOTORES_DECLARADOS`: recalcula el caso semilla en Excel real y exige cero errores (`#REF!`, `#NAME?`, `#VALUE!`...), llama a `comprobar_procedencia()` (aborta si una cita no existe), comprueba que todo veredicto pinta el color que promete (`DisplayFormat`, no openpyxl) y que el dictamen no contradice a sus propias verificaciones. **No se escribe código nuevo para esto tampoco.** | Sigue |
| **8. Entregar** | Corre `pytest`, corre `verificar.py --resources ... --wb ...`, exporta la hoja a PDF y mírala (las trampas de openpyxl con bordes y comentarios solo se detectan mirando el PDF, no leyendo el `.xlsm`). | Fin |

**La Fase 1 se detiene siempre.** No es una parada condicional ni una que se salta si
el artículo "parece simple". Es el criterio de aceptación nº 5 del diseño: *"un motor
construido sin que nadie mirara la propuesta es un motor que nadie firmó"*. La skill
no tiene autoridad para decidir por su cuenta qué ecuación de un artículo gobierna un
resultado — eso es exactamente lo que la Fase 1 existe para no dejar pasar. Si en algún
momento resulta tentador "adelantar" la Fase 2 sin que el ingeniero haya contestado al
inventario de la Fase 1, eso es un error de la skill, no una optimización.

## El formato del inventario de la Fase 1

Lo que se presenta al ingeniero no es un resumen: es **lo que el artículo imprime**,
con el número de bloque al lado para que se pueda ir a mirar el JSON directamente.
Estructura sugerida:

```
ARTÍCULO XXX — <título tal como lo imprime el bloque section_header nivel 1>
Fuente: resources/ASME PCC/pcc_2/<ruta>/art_XXX.json (encontrado en los dos espejos:
sí/no — decir cuál falta si solo está en uno)

## Ecuaciones impresas
- (1) <clausula, p.ej. "XXX-3.2"> · bloque <n> · texto tal como aparece
- (2) <...> · bloque <n> · EQUATION VACÍA, imagen en figure bloque <m>: fig/xxx.png
      -> leída con Read; texto recuperado: "<...>"   [léela de verdad, no la asumas]

## Figuras con dato (no solo ilustrativas)
- Figura XXX-Y-Z · bloque <n> · fig/xxx.png · qué dato trae si se lee

## Umbrales de doble unidad
- "<clausula>": <valor SI> (<valor US>) · bloque <n>

## Límites y prohibiciones
- <clausula> · bloque <n> · texto

## Fabricación / examen / prueba
- <clausula> · bloque <n> · texto

## Tablas
- <clausula o número> · bloque <n> · qué publica

## Huecos declarados (no se completan, se dicen)
- <qué falta> · qué archivo/bloque re-extraer · por qué no se puede resolver desde
  resources/ tal como está

## Propuesta (para que el ingeniero marque qué entra)
- Entradas: <lista con magnitud y de dónde sale su rango, si el código lo publica>
- Cálculos: <lista de Filas propuestas con su fórmula por nombre y su cita>
- Verificaciones: <requerido / adoptado / criterio, con cita>
- Pasos del flujo: <uno por paso del procedimiento del artículo>
- Especificaciones técnicas: <fabricación, examen, prueba — solo si el artículo las
  imprime como requisito, no como narrativa>
```

Ningún ítem de "Propuesta" entra a la Fase 2 sin que el ingeniero lo confirme,
corrija o descarte.

## Los tres casos de `resources/` que ya costaron caro

Ya se pagaron una vez en el 212 y en el 206 (ver el `CLAUDE.md` del repo, sección
"El F9 del ingeniero" y "Cada artículo es un nivel del árbol"). La skill los reconoce
en la Fase 1 y actúa así:

1. **Ecuación vacía con imagen.** El bloque `equation` (o un `figure` que hace de
   ecuación) llega con `text` vacío o inservible, pero trae `image: fig/....png`. Esa
   imagen **es parte de `resources/`** — no un documento externo — así que se lee con
   la herramienta `Read` sobre el PNG, y de ahí sale la fórmula. Ejemplo real ya hecho:
   el cateto de filete del Art. 206 (`206-3.5`, bloques 72 y 77 de
   `art_206.json`) llegó con el `text` vacío y la ecuación solo en
   `fig/fig_206_3_5_1.png` / `fig/fig_206_3_5_2.png`; se leyeron las imágenes y el
   texto recuperado («w = Ts + G…», «w_max = 1.4 Tp + G…») quedó grabado en el JSON
   con su `extraction_amendments` (script, fecha, SHA-256 de la imagen fuente, y la
   procedencia explícita: "Recuperadas leyendo los PNG del propio resources/... NO se
   usó ningún folio/PDF externo ni memoria del modelo"). Ese es el patrón a seguir.
2. **Umbral de doble unidad.** El código imprime "40 mm (1.5 in.)": las dos cifras se
   guardan **tal cual**, nunca se convierte una en la otra (1,5 in son 38,1 mm, y esa
   diferencia es del código, no un error de redondeo a corregir). Mismo mecanismo que
   `leer_umbrales_pcc2()` en `build_db_materiales.py`.
3. **Extracción colapsada.** Si un bloque llegó roto —una tabla que se comprimió
   dentro de un encabezado, una ecuación con símbolos revueltos por una costura de
   página (el patrón `i k jjjjj y { zzzzz ...` que aparece en algunas ecuaciones del
   212)— la skill **no lo rellena ni lo adivina**: declara el vacío en el inventario
   de la Fase 1, dice qué archivo re-extraer y **se detiene** ahí. Completar esa
   extracción es un trabajo de re-extracción sobre `resources/` (como
   `completar_tabla_302_3_4.py` o `completar_art_206_filete.py` ya hicieron para otros
   casos), no algo que la skill resuelva por su cuenta ni con el PDF externo — el PDF
   solo sirve para corregir y verificar los JSON, nunca como entrada directa de un
   motor (Regla nº 1 del `CLAUDE.md` del repo).

## Trazabilidad: la Regla nº 1, operacionalizada

Que la skill "lea `resources/`" no es una promesa, es un mecanismo que ya existe en el
chasis y que la skill tiene que respetar sin sortearlo:

1. **Toda `Fila`, `Verificacion` y `Especificacion` de tipo cálculo lleva su `Cita`**
   (`archivo`, `bloque`, `clausula`). Una `motor_declarado.Fila` con `tipo=FORMULA` y
   `cita=None` **sí** se puede construir —es un `NamedTuple` y `cita` vale `None` por
   defecto—; lo que no se puede es **construir el libro** con ella:
   `comprobar_procedencia()` aborta en tiempo de build con `SystemExit`, citando la fila
   y la hoja (corregido 2026-09-13: aquí se afirmaba que no se podía construir el objeto).
2. Esa procedencia se escribe **tres veces y en tres formas**, y las tres ya las
   escribe el chasis, no la declaración:
   - la **columna de referencia** que ve el ingeniero (`f.cita.clausula`, columna G,
     vía `helpers.rotulo(..., referencia=...)`);
   - el **comentario de la celda** de valor (`f.comentario`, vía
     `helpers.rotulo(..., comentario=...)` — solo en columna de valor, nunca en la de
     rótulo, regla del F9 2026-09-13);
   - la **tabla de trazabilidad** del motor, que devuelve `comprobar_procedencia()`
     (`(clave, clausula, archivo, bloque)` por fila de cálculo).
3. **El build aborta** si una celda de cálculo llega sin procedencia
   (`comprobar_procedencia` es la primera línea de `build_motor_declarado`, antes de
   escribir una sola celda), si una cita apunta a un bloque que no existe en el JSON
   (`if not 0 <= f.cita.bloque < len(bloques): raise SystemExit(...)`) y si apunta a un
   bloque de tipo `section_header` — un **rótulo no publica un valor**, y el bloque
   existe, así que sin este guardia la cita parecía válida apuntando a donde el dato no
   está (2026-09-13).
4. `verificar.py` §12 repite la comprobación de procedencia contra el libro ya
   construido — el mismo guardia, no una copia — y añade lo que solo Excel puede decir:
   que el caso semilla recalcula sin error, que todo veredicto pinta y que el dictamen
   no contradice a sus verificaciones.

**Nota sobre `Cita.archivo`:** es texto para que el ingeniero vea de un vistazo de
qué documento sale un dato, pero la ruta que de verdad valida `comprobar_procedencia()`
es siempre `motor.fuente` (un solo JSON por motor — el artículo entero vive en un solo
archivo de `resources/`). Todas las `Cita` de un mismo `Motor` describen bloques de ese
mismo `motor.fuente`.

## Qué prueba qué — y el límite declarado de la skill

| Nivel | Qué garantiza | Ya existe hoy |
|---|---|---|
| `pytest` sobre el chasis (`test_motor_declarado.py`) | que `resolver_direcciones`/`emitir_tabla`/`sustituir_nombres`/`comprobar_procedencia` hacen lo que dicen, con motores de prueba sintéticos | sí |
| `pytest` sobre el libro (`test_dashboard.py::TestMotoresDeclarados`, `TestNodoDeArticuloDerivado`) | que la hoja de CADA motor declarado tiene lo que su declaración promete: entrada editable y en el manifiesto de reinicio, comentario en columna de valor, nodo de navegación con sus tres hijos en orden | sí, genérico |
| `verificar.py` §12 | que el caso semilla del motor nuevo **recalcula sin errores de Excel**, que **toda cita existe** en el JSON, que **todo veredicto pinta** el color que promete, y que el **dictamen no contradice** a sus propias verificaciones | sí, genérico |
| `verificar.py` **por artículo** (el equivalente a §6e/§6f del 212/206) | que **el número que da la hoja es el que pide el código** para el caso semilla | **NO lo genera esta skill.** Lo escribe el ingeniero cuando quiere esa red, exactamente como ya hizo con §6e (Art. 212) y §6f (Art. 206) |

Esa última fila es el límite declarado de esta skill, y hay que decirlo sin adornos:
**un chasis genérico puede probar la maquinaria, no la ingeniería.** Que el semáforo
pinte verde, que el dictamen no se contradiga a sí mismo y que la cita exista de
verdad en el JSON no demuestra que la fórmula que se citó sea la fórmula correcta para
ese caso de carga — eso solo lo puede juzgar quien conoce el artículo, recalculando a
mano el caso semilla y comparándolo con lo que la hoja da. Escribir esa red sería el
modelo comprobándose a sí mismo con su propia lectura del código; por eso queda fuera
del alcance de la skill y dentro del trabajo del ingeniero.

## Referencias

- `referencias/reglas_del_libro.md` — el catálogo de reglas de diseño del libro que la
  Fase 2 (Declarar) y la Fase 3 (Chasis, ya escrita) tienen que respetar: las 14 reglas
  de diseño, la Regla nº 1, las tres reglas del flujo de la cascada, el sistema visual,
  el semáforo, los comentarios, el ancla del VML, los dos decimales, y las trampas de
  openpyxl ya pagadas.
- `referencias/plantilla_declaracion.py` — una declaración `Motor(...)` comentada, con
  ejemplos tomados del Art. 206 (sin migrar el motor real: el 206 escrito a mano sigue
  siendo la fuente de verdad de ese artículo). Úsala como plantilla para la Fase 2, no
  la copies literalmente.
