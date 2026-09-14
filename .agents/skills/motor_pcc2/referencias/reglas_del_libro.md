# Catálogo de reglas del libro

Esto es un **catálogo**, no una paráfrasis: cada regla viene con el sitio de donde
sale, para poder ir a comprobarla contra el original en vez de fiarse de este
resumen. Las fuentes son el `CLAUDE.md` del repo (raíz de `asme-pcc`) y el código de
`outputs/Base_Datos_Materiales_ASME/scripts/`. Los números de línea de código pueden
moverse con el tiempo; el nombre de la función y el archivo no.

Un motor DECLARADO (skill `motor_pcc2`) no está exento de ninguna de estas reglas
solo porque nazca de una declaración: el chasis (`motor_declarado.py` +
`build_motor_declarado`) ya las aplica por dentro, pero la Fase 2 (Declarar) tiene que
saber por qué el chasis hace lo que hace, para no proponer algo que el chasis no puede
cumplir de forma honesta (p. ej. una fila de cálculo sin `Cita`, o una unidad sin su
par SI/US).

---

## Regla nº 1 — `resources/` es la única fuente de verdad

*Fuente: `CLAUDE.md` del repo, sección "Regla nº 1 — `resources/` es la única fuente
de verdad".*

Ningún valor normativo —esfuerzos admisibles, factores, fórmulas, límites— puede salir
de la memoria del modelo. Siempre se lee el archivo fuente antes de citar o calcular.
Si un dato no existe en `resources/`, no se inventa ni se aproxima: se declara el
vacío explícitamente y se pregunta cómo proceder.

Un valor puede vivir en `resources/` como **imagen de ecuación**: el bloque
`figure`/`equation` que trae `image: fig/….png` es parte de `resources/`, así que se
lee esa imagen (con la herramienta `Read`) y de ahí se recupera la fórmula — no se pide
el PDF del código. Solo si el dato no está en `resources/` ni como texto ni como
imagen (bloque vacío y sin `image`, comprobado en **los dos espejos** del árbol de
PCC-2 — ver Fase 0 de `SKILL.md`) se trata como vacío real: se declara y se repara
re-extrayendo ese archivo.

El PDF de ASME solo sirve para **corregir y verificar** los JSON de `resources/`,
nunca como entrada directa de un motor.

**Mecanismo en el chasis** (no solo declaración de intención):
`motor_declarado.comprobar_procedencia(motor, resources)` — aborta con `SystemExit`
si una `Fila` de tipo `FORMULA` no tiene `Cita`, y aborta también si `f.cita.bloque`
cae fuera del rango de `blocks` del JSON en `motor.fuente`. Se llama como primera línea
de `build_motor_declarado()`, antes de escribir una sola celda, y de nuevo desde
`verificar.py` §12 contra el libro ya construido.

---

## Las 14 reglas de diseño del libro

*Fuente: `CLAUDE.md` del repo, sección "Reglas de diseño del libro — no romper".*

1. **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`,
   `XLOOKUP`, `VSTACK`, `_xlfn`. Solo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`.
   **Enmienda Rev. 3:** el entregable es `.xlsm` y ya **no abre en Google Sheets** —
   la capa de navegación es VBA. Lo que se pierde es solo la navegación; la
   restricción sobre las fórmulas se mantiene entera: todo el cálculo sigue
   teniendo que ser portable, y `verificar.py` sigue fallando si aparece una matriz
   dinámica. Corolario: un texto no se guarda como fórmula: un `="texto…"` de más
   de 255 caracteres lo parte Excel en `_xlfn._LONGTEXT(...)` al reguardar
   (`normalizar_textos_como_formula`).
2. **Validación de datos: solo rango literal o lista de ítems.** Nunca una fórmula
   (`OFFSET`, `INDIRECT`) como origen. Las listas dependientes se materializan en
   columnas ocultas.
3. **Banda compacta.** Junto a la banda impresa (que se conserva para auditar), cada
   base lleva una banda sin huecos para la consulta.
4. **No extrapolar.** Por encima de la última temperatura tabulada o de la Temp. máx.
   del material, dictamen "FUERA DE RANGO" y resultado bloqueado.
5. **Cascada contigua.** Cada base se ordena por familia → composición → forma →
   especificación → grado.
6. **Unidades visibles** en toda variable, indicador y eje de gráfica.
7. **`material_id` único**, desambiguado en este orden: notas del código → Rm → Re →
   número de línea impreso.
8. La **familia de material** es agrupación de navegación derivada del UNS y la
   composición impresa — no es dato normativo y no entra en ningún cálculo.
9. Los valores se cargan **tal como están impresos**; SI y US son extracciones
   independientes, nunca conversiones.
10. **Conmutador de unidades en todo motor** (SI ↔ US). Corolario de la regla 9:
    cuando el código publica los dos sistemas en la misma tabla, el conmutador cambia
    la **columna**; cuando publica una tabla por edición, cambia la **hoja**. Un
    material sin homólogo en la otra edición muestra "sin equivalente en la edición
    US", nunca un valor convertido. Excepción declarada: los factores adimensionales
    (Ec, Ej) que el código publica en una sola tabla para los dos sistemas llevan una
    celda fija rotulada en vez de un conmutador que no gobernaría nada
    (`TXT_ADIMENSIONAL`).
11. **El factor publicado es un mínimo, y el motor tiene que decirlo** cuando una nota
    del código lo condiciona (p. ej. Notas (4)/(5) de la Tabla A-2 del B31.3, que dicen
    cosas opuestas con el mismo formato).
12. **Todo material se selecciona de una base de datos, nunca de una lista fija.**
    Cualquier campo que pida un material se resuelve con una cascada de listas
    desplegables materializada contra la base que corresponda. Nunca una lista de
    texto tecleada directo en el `formula1` de una `DataValidation`.
13. **Los motores de búsqueda y los motores de cálculo no se conectan entre sí.**
    Ningún motor de cálculo lee una celda de un motor de búsqueda ni al revés: los dos
    leen exclusivamente de las bases `DB_*`.
14. **Toda variable de entrada que exista en una base de datos se elige de una lista
    desplegable; nunca se teclea.** Lo que sí se teclea es lo que ninguna base
    publica: datos de proceso y de campo (temperatura y presión de operación,
    dimensiones del defecto, sobreespesor de corrosión, medidas de la reparación).

---

## Las tres reglas del flujo de la cascada de material

*Fuente: `CLAUDE.md` del repo, sección "La cascada de material fluye de arriba hacia
abajo, y solo hacia abajo".*

Ninguna fórmula de Excel puede cumplirlas porque ninguna fórmula puede **vaciar** una
celda de entrada; van en `Workbook_SheetChange` → `ProcesarCascada` /
`ReiniciarPorCambioDeUnidades` del VBA:

1. **Un nivel está bloqueado mientras el de arriba esté vacío.** No se elige
   Tipo/Grado sin haber elegido antes familia → composición → forma → especificación.
2. **Al cambiar un nivel, todo lo que cuelga de él se reinicia** — los niveles de
   abajo, la variante, y con ellos el S(T) resuelto.
3. **Cambiar el sistema de unidades reinicia la hoja entera**, con confirmación
   previa; si se dice que no, el selector vuelve al sistema anterior.

El VBA no lleva ninguna dirección de celda para esto: cada motor publica su propia
**cadena de cascada** en una columna oculta (la 101), y toda rutina que apague los
eventos (`EnableEvents = False`) vuelve por una etiqueta que los enciende — si una
muriese apagada, las tres reglas dejarían de aplicarse el resto de la sesión y nada lo
avisaría.

---

## Sistema visual — Swiss Industrial Print

*Fuente: `CLAUDE.md` del repo, sección "Sistema visual — Swiss Industrial Print
(Rev. 4e)"; tokens en el bloque de constantes al principio de `build_db_materiales.py`.*

Un solo sistema para las 72 hojas: papel de documentación sin blanquear (`PAPEL`
`F4F4F0`), tinta carbón (`TINTA` `050505`) y **un** acento rojo (`ROJO` `E61919`).
`AMARILLO` (`FAEFC0`) es la única excepción declarada al acento único, y está acotada
por nombre de hoja: solo en los motores de cálculo, nunca en un buscador ni en una
hoja de datos. `test_dashboard.py::TestSistemaVisual` recorre las celdas con formato
del libro construido y falla si aparece un color fuera de la paleta o una fuente que
no sea Arial Black (`MACRO`) o Consolas (`MONO`).

Un motor DECLARADO usa los **mismos** `Helpers` de estilo que el 212 y el 206
(`_helpers_del_libro` en `build_db_materiales.py`), así que hereda el sistema visual
automáticamente — no hay nada que declarar sobre color en `motor_declarado.py`.

---

## El semáforo: `bgColor` **y** `fgColor`

*Fuente: `CLAUDE.md` del repo, sección "El F9 del ingeniero (2026-09-13)"; función
`relleno_dxf()` en `build_db_materiales.py`.*

En un formato condicional (`dxf`), Excel pinta el relleno con **`bgColor`**, no con
`fgColor`. `PatternFill("solid", fgColor=...)` a secas deja la regla disparando —el
color de la fuente del `dxf` sí cambia— pero el relleno se queda como estuviera la
celda: las verificaciones salían grises en vez de verdes o rojas, y el dictamen del
206 quedaba con letra clara sobre fondo claro. `relleno_dxf()` escribe los **dos**
colores iguales a propósito, así que da igual cuál interprete Excel.

**Ninguna prueba de openpyxl puede verlo**: la regla *está* escrita y eso es todo lo
que openpyxl sabe leer. Solo Excel real, vía `DisplayFormat`, dice qué color pinta de
verdad — es lo que hace `verificar.py` §6h (y §12 para los motores declarados).

---

## Los comentarios viven solo en la columna de VALOR

*Fuente: `CLAUDE.md` del repo, sección "El F9 del ingeniero", punto 3; función
`aplicar_reglas_de_comentario()` y constante `COLS_VALOR_MOTOR = (4, 5, 6)` en
`build_db_materiales.py`.*

El comentario de una fila va **solo** en su columna de VALOR (D, E, F — F publica el
veredicto y cuenta como valor), nunca repetido también en la columna de rótulo. Los
botones (H..J) conservan el suyo aparte. En el chasis declarativo, esto lo garantiza
`_helpers_del_libro.rotulo()` en la ruta de `build_motor_declarado`: escribe el
comentario en `COLS_VALOR_MOTOR[0]` (columna D) y en ningún otro sitio.
`build_motor_declarado()` **sí** llama después a `aplicar_reglas_de_comentario`
(con `_reglas_de_comentario_de(motor)`; busque la llamada por nombre, los números de
línea de este archivo se mueven con cada parche) — pero para
un motor DECLARADO esa pasada es, en la práctica, un no-op de confirmación: solo
existe una columna de valor por fila (ver el cuarto límite, más abajo:
`motor.casos` no genera una columna por caso), así que no hay una segunda columna
donde redistribuir el comentario. La garantía real de esta regla, hoy, sigue siendo
la línea de `rotulo()` — no el barrido posterior.

El comentario se **dimensiona con su texto y se ancla a su celda** (ver siguiente
regla); 90 px fijos cortaban 102 de 811 comentarios reales.

---

## El ancla del VML

*Fuente: `CLAUDE.md` del repo, sección "El F9 del ingeniero", punto 4; función
`anclar_comentarios()` en `build_db_materiales.py`.*

openpyxl **no escribe** el `<x:Anchor>` de la forma VML de un comentario: sin él,
Excel mide el `margin-left` desde la esquina de la **hoja**, no de la celda, así que
editar la nota de una fila lejana abre el cuadro arriba del todo. `anclar_comentarios()`
lo inyecta **después de guardar**, reescribiendo el VML del paquete `.xlsm` con la
geometría real de la hoja (anchos de columna, altos de fila) — algo que openpyxl no
puede calcular desde la forma en memoria.

---

## Los dos decimales

*Fuente: `CLAUDE.md` del repo, sección "El F9 del ingeniero", contexto de
`aplicar_dos_decimales()` en `build_db_materiales.py`.*

Dos decimales en toda celda de valor cuya fila declare una **magnitud** (mm, MPa,
kg/cm², N/mm, °C…): antes de esto el motor mostraba `77.82802606` junto a un dato de
entrada de dos cifras, que es leer ruido. Se fija el **formato**, nunca el valor — el
número de abajo sigue entero y los cálculos aguas abajo no cambian ni un dígito. La
regla es la unidad de la **fila** (su `magnitud`), no el tipo de la celda: una fila que
publica "—", un modo o un factor adimensional se queda como está.

En un motor DECLARADO todo esto arranca en `Fila.magnitud`, y la cadena tiene **tres
eslabones, no dos** (corregido 2026-09-13: esta sección decía que
`aplicar_dos_decimales` se alimentaba de `_unidades_de` «vía el mismo `_semaforo_de`»,
y es falso):

1. `_helpers_del_libro.rotulo()` escribe el rótulo SI de la magnitud en la **columna C**
   (`COL_UNIDAD_MOTOR`) de la fila. Hasta el 2026-09-13 no lo escribía en ningún sitio, y
   eso dejaba muertos los dos eslabones siguientes.
2. `aplicar_unidades_motor(ws, _unidades_de(motor), sel)` sustituye esa celda por la
   fórmula del conmutador SI/US. Solo pisa una celda que **ya** tenga rótulo: con la C
   vacía se limitaba a emitir un `ISSUE` por fila.
3. `aplicar_dos_decimales(ws, semaforo)` **deriva de la columna C leída de la hoja**, no
   de `_unidades_de` ni de `_semaforo_de`. Su argumento `semaforo` sirve para otra cosa:
   marcar las filas de verificación, que no declaran unidad y donde el número es magnitud
   igual.

`_unidades_de` y `_reglas_de_comentario_de` recorren `MD.bloques(motor)` —Secciones **y**
Pasos—: una fila de cálculo dentro de un `Paso` del anexo del flujo recibe unidad,
decimales y regla de comentario como cualquier otra. Una `Fila` cuyo valor es discreto
(una lista, un modo) declara `magnitud=""` (el valor por defecto) para quedar fuera de
esta regla, y su columna C se queda vacía.

---

## Trampas de openpyxl ya pagadas

*Fuente: `CLAUDE.md` del repo, secciones "Reglas de diseño del libro" (nº1, corolario),
"Sección II partes A, B y C", y "El F9 del ingeniero"; funciones citadas de
`build_db_materiales.py`.*

1. **El relleno de un rango fusionado se hereda de la celda ancla, pero el borde NO.**
   `franja()` pone la banda de tinta y la franja roja recorriendo celda a celda del
   rango, no solo la ancla, precisamente porque el borde de una celda fusionada que no
   sea la ancla no se pinta si se aplica solo ahí — sale como un muñón de una columna.
2. **openpyxl miente en los dos sentidos al leer y escribir bordes.** Al leer un libro
   reconstruye el borde de la ancla sobre todo el rango fusionado (parece que
   estuviera, aunque no lo esté); al escribir, no propaga nada si el estilo se puso
   después de fusionar. La única evidencia válida es **exportar la hoja a PDF/PNG
   desde Excel** e ir a mirarla — no comprobarlo con openpyxl.
3. **El tamaño de un comentario se pierde al clonarlo.** `Comment(texto, autor)` a
   secas devuelve la caja al 144×79 por defecto de openpyxl, deshaciendo el alto que
   `_nota()` había calculado a partir del texto real. `_clonar_nota()` copia también
   `width`/`height` del comentario original — necesario porque los comentarios se
   clonan más de una vez en el pipeline del libro (remapeo de filas, reparto por
   sección) y cada clonación sin este cuidado vuelve a perder el tamaño.
4. **Un texto que empieza por `=` o `+` se convierte en fórmula si no se le dice lo
   contrario.** El código de la Sección II imprime celdas como
   `"= 3.18 mm in any 1.524 m"`, que son texto, no fórmula. `_txt_celda()` fuerza
   `cell.data_type = "s"` cuando el primer carácter es `=`, `+`, `-` o `@`; sin esto,
   Excel evalúa la celda como fórmula rota y muestra `#NAME?`.
5. **openpyxl no asigna `codeName` a las hojas que crea.** El VBA tiene que referenciar
   hojas siempre por `.Name`, nunca por CodeName — Excel se los inventa recién al
   abrir el libro.
6. **Toda declaración de módulo VBA (`Const`, `Dim`, `Type`) precede a la primera
   rutina.** Si no, VBA reporta "Variable not defined" en cada uso, y Excel abre un
   diálogo modal al compilar durante `SaveAs` que cuelga la automatización sin
   mensaje. `make_vba_seed.py` lo comprueba con `lint_vba()` antes de tocar COM.
7. **VBA no admite más de 25 continuaciones de línea (`_`) en una línea lógica.** Una
   lista larga arma la cadena por concatenación (`s = s & "|…"` + `Split`), que no
   tiene tope, en vez de un `Array( _ … )` con una hoja por línea.

Ninguna de estas siete se detecta leyendo el `.xlsm` con openpyxl: las que tocan
apariencia (1, 2, 3) solo se ven exportando la hoja real a PDF/PNG y mirándola; las de
VBA (5, 6, 7) solo se ven abriendo el libro en Excel con las macros vivas.
