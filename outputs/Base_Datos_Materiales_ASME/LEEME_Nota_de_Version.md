# Nota de versión — Motor de Cálculo ASME PCC, Rev. 3
## Base de datos de materiales (PLAN-DB-MAT-001)

**Fecha:** 2026-09-06 · **Fuente única de verdad:** `resources/` · **Entregable:** `Motor_de_Calculo_ASME_PCC_Rev3.xlsm`

---

## Apéndice A del B31.3 — la costura de las páginas enfrentadas (2026-09-07)

El Apéndice A se imprime **a doble página**: la izquierda lleva la identificación del
material y la derecha el número de línea con la rejilla de esfuerzos admisibles. La
extracción tiene que fusionarlas por `Line No.`, y en esa costura se colaron tres defectos.

### Se recuperan 197 esfuerzos admisibles que estaban fuera de la curva

En el bloque de **aleaciones de níquel de la Tabla A-1C** (folios 334–347) la página
izquierda termina con las columnas `Min. Temp. to 100 | 200 | 300`. El extractor pegó la
palabra **`Metal`** del título de banda —*Basic Allowable Stress, S, ksi, at Metal
Temperature, °F*— al encabezado `200`, inventó un campo de identificación `metal_200` y
metió ahí el esfuerzo. Consecuencia: **la curva de 197 filas empezaba en 300 °F** y el
punto de 200 °F, que sí está impreso, no se podía consultar. En modo US, `B161 / N02201`
devolvía 6,3 ksi a 200 °F cuando el código imprime **6,4**.

Los 197 valores vuelven a su punto de la curva. Auditoría de balance: A-1C pasa de 12 488
a **12 685** valores no nulos, `+197` exactos, y **ninguna otra tabla se mueve**.

### El enlace SI ↔ US del B31.3 sube de 1 065 a 1 208 materiales

En 142 filas de A-1C la especificación y el grado quedaron pegados en un solo campo
(`"A139 A"`, `"A524 II"`). Se separan **solo** cuando el valor coincide **exactamente** con
`spec + " " + grado` de la misma línea en A-1: eso demuestra por sí solo que son dos celdas
pegadas. No se usa A-1 como referencia general —las dos ediciones nombran distinto el
grado, la SI imprime `API 5L L245` donde la US imprime `API 5L B`— sino solo como prueba
del pegado. Al recomponer la clave bilingüe, **143 materiales más** encuentran su homólogo
en la otra edición.

### Y lo demás

- **A-3 fusionaba `Class (or Type)` con `Description` en 38 filas** (`"… Seamless pipe"`,
  `"Type S Seamless pipe"`, `"All Welded pipe"`). Separadas; las tres que no eran
  mecánicas van verificadas contra los folios 366 y 368.
- **220 celdas** de A-1C, A-4 y A-4C arrastraban la **elipsis de una columna vacía
  contigua** (`spec_no = "A179 …"`, `p_no = "1 …"`). En el código `…` significa «sin
  valor»: se retira.
- **Un título de sección absorbido**: A-1C bloque 3 línea 86 guardaba
  `"A353 … Low and Intermediate Alloy Steel — Forgings and Fittings"` en `spec_no`.
  A-1 confirma que esa línea es `A353`, forma `Plate`, mismo UNS K81340.

### La Tabla A-2 estaba impecable

Cotejada fila a fila contra el folio 364: 24 filas, 8 grupos, factores `Ec` y Notas (1)–(5)
completas. **0 defectos.** `TestApendiceACorregido::test_a2_sigue_intacta` la fija para que
una reextracción futura no la estropee.

### Lo que NO se ha tocado

Los nombres de columna que el propio código escribe de dos maneras —`Min. Tensile
Strength, ksi` en unos bloques y `Minimum Tensile Strength, ksi` en otros, `Class/
Condition/ Temper` y `Class/ Condi- tion/ Temper`— **se conservan como claves distintas**:
es lo que está impreso, y el builder ya las resuelve con sus alias. Normalizarlas sería
reescribir el código, no corregir la extracción.

### 21 páginas derechas huérfanas, recuperadas del folio

La revisión de las propias correcciones encontró un segundo fallo de la misma costura: en
**21 líneas del bloque 6 de A-1C** la fusión perdió la página **izquierda entera**. Esas
filas no tenían identificación alguna —ni spec, ni UNS, ni composición— y su curva
**empezaba en 400 °F**, sin los puntos impresos de `Min. Temp. to 100`, `200` y `300` °F.

Se han recuperado leyendo los folios **336, 338, 340, 342, 344 y 346** ampliados, y
contrastando cada fila con la edición SI por composición, especificación, UNS, P-No. y
resistencias convertidas. Son B366, B435, B564, B572, B619, B622, B625, B626, B649, B675,
B688, B690, B804 y una fundición A494 CX2MW. **+63 datos**: 42 esfuerzos y 21 valores de
`Min. Temp. to 100`.

No se rellenaron desde A-1 —habría sido convertir unidades e inferir identidad—, y la
lectura directa demostró por qué: **las dos ediciones no numeran igual sus líneas**. La
línea 203 de A-1C es `B649 / N08031`; la 203 de A-1 es `B164 / N04400`. Copiar por posición
habría metido el material equivocado.

Queda declarada, sin armonizar (regla 9), una discrepancia del propio código: para el
`46Fe–24Ni–21Cr–6Mo–Cu–N` (N08367), **A-1 publica 427 °C** de temperatura máxima —unos
800 °F— mientras que **A-1C publica 900 °F** en las líneas 61 a 65, 125 y 126, y 800 °F en
la 66.

### Y el UNS dejaba de arrastrar la clase

Al transcribir apareció un tercer defecto del mismo bloque: **186 filas con el `UNS No.` y
el `Class/Condition/Temper` pegados** (`"N08031 Annealed"`), que dejaban la columna de clase
vacía en 208 de las 218 filas del bloque. El folio 336 las imprime en columnas separadas y
un UNS es una letra con cinco dígitos, así que el corte es determinista. Dos casos llevaban
el nombre de la clase partido en dos líneas (`"N08810 Sol. tr. or"` + `"ann."`) y se
recomponen uniéndolos. Un UNS traía además un carácter espurio de otra columna
(`">F33100"` → `F33100`). Las filas con clase declarada pasan de 322 a **526**.

**Alcance de la verificación:** A-2 y A-3 se cotejaron **completas** contra el impreso, y las
correcciones de A-1C se verificaron contra los folios 286 y 334–347. A-1, A-1C, A-4 y A-4C
—2 570 filas sobre 233 folios— se auditaron por **clases de defecto** y por **contraste entre
las dos ediciones**, no fila a fila. Lo corregido está verificado; no se afirma que no quede
nada por encontrar.

Se corrige con `scripts/completar_apendice_a.py`. En el libro: `DB_B31_3C` pasa de 14 971 a
**15 231** valores auditados y el total de 271 276 a **271 536**. Es la única línea que se
mueve en el protocolo de aceptación. `TestApendiceACorregido`, 8 pruebas.

**El enlace SI ↔ US del B31.3 acaba en 1 228** de 1 288 materiales, desde los 1 065 de
partida: +143 por separar spec y grado, +20 por recuperar las huérfanas.

---

## Apéndice B del B31.3 — identificación y estructura corregidas (2026-09-07)

Cotejadas las 7 tablas y las 28 entradas del índice contra los folios impresos 399–405:
**ni un solo valor mal**. Lo que fallaba era cómo quedaron identificadas y estructuradas
las filas. Seis defectos, uno de ellos peligroso:

- **B-5 tenía la cabecera desalineada.** El código imprime `Mín °C · Mín °F · Máx °C ·
  Máx °F`; el JSON declaraba `°C · °F · Maximum °F`, y **los 232 °C —que son el máximo
  del vidrio borosilicato— quedaban en un campo llamado `c`**. Quien lo leyera como
  temperatura mínima se llevaba 232 °C de mínimo. Corregido a `minimum_c/f` y
  `maximum_c/f`. Mismo reparto aplicado a **B-4**.
- **B-1 fusionaba dos columnas.** La fila del ABS guardaba `"… PR"` en `astm_spec_no`
  con la designación de tubería vacía, y la F2389 guardaba `"F2389 PR"`. B-1C separaba
  bien esas mismas filas, lo que confirmaba el defecto.
- **`F2788/F2788M` llevaba un espacio espurio** en las dos ediciones, del salto de línea
  del impreso; así no emparejaba con el índice de especificaciones.
- **B-3 tenía tres columnas colapsadas en una celda.** El impreso es una rejilla 3 × 2
  con seis especificaciones; ahora son **6 filas** —`D2517, D2996, D2997, D3517, D3754,
  AWWA C950`—, leídas por columnas, que es como quedan en orden ascendente.
- **B-6 partía un nombre de material a dos líneas entre dos filas**: la F1974 quedaba con
  un `"Metal insert fittings for"` truncado y un `"PE-AL-PE systems"` suelto como si
  fuera otro material. Y su columna `°F` pasa a `maximum_f`: la tabla solo publica
  límites máximos.
- **El índice no traía sus notas.** Se añaden las del folio 399, incluida la **Nota (2)**,
  que es normativa: *el término «fiberglass RTR» sustituye a la designación ASTM
  «fiberglass»*.

### Tres rarezas del propio código, conservadas y declaradas

Van en `observaciones_del_codigo` de cada archivo, no se corrigen (regla 9):

1. **B-1 imprime `…` como mínimo** para D2846 / CPVC4120; **B-1C imprime 73 °F** para esa
   misma fila.
2. **F2389**: máximo **110 °C** en B-1 frente a **210 °F** (= 98,9 °C) en B-1C, y
   designación de tubería `PR` en SI frente a `IPS Sch. 80` en US.
3. **B-6, F1282 a 862 kPa**: imprime **100 psi** donde sus filas gemelas imprimen 125 psi
   (862 kPa = 125 psi).

Se corrige con `scripts/completar_apendice_b.py`. El PDF queda versionado en
`resources/.../appendix_b/fuente/`; es un escaneo sin capa de texto y sin OCR
estructurado, así que las correcciones van en el script como tabla literal, cada una
citando su folio. **Ningún valor numérico se ha tocado**: 0 filas alteradas en las 5
tablas con números. `DB_NoMetalicos` pasa de 939 a **944** pares campo/valor.
`test_build_db.py::TestApendiceBCorregido` lo protege con 9 pruebas.

---

## Apéndice C del B31.3 — capa de metadatos completada (2026-09-07)

Los **valores** del Apéndice C siempre estuvieron bien. Lo que faltaba era todo lo que
rodea al número, y sin eso no se puede montar un motor de propiedades físicas:

- **C-1 y C-1C no declaraban la unidad de sus coeficientes.** Ahora sí, leída del folio
  impreso: `A = Mean Coefficient of Thermal Expansion, 10⁻⁶ mm/mm/°C` y
  `B = Linear Thermal Expansion, mm/m` desde 20 °C; en la edición US, `10⁻⁶ in./in./°F`
  y `in./100 ft` desde 70 °F.
- **Las Notas (2)–(6) de C-1 se imprimen a tres columnas** y se habían leído por filas,
  dejando los miembros de los Grupos 1 a 4 intercalados. Ahora están desintercaladas en
  `note_members` —**51, 6, 13, 14 y 18** miembros—, con la misma estructura y la misma
  convención de composición que las Notas de TM-1 / TE-1 de la II-D. Sin esto no se
  puede decir si un 1¼Cr–½Mo pertenece al Grupo 1, que es para lo que sirve la tabla.
- **C-2 y C-3 habían perdido la indentación del impreso.** Cinco filas colgaban de un
  subgrupo que no es el suyo —`Gray iron` figuraba como **acero inoxidable
  austenítico**—, nueve filas de C-2 se quedaron sin grupo al cruzar el corte de página
  y un título llegó partido por la mitad (`"and Reinforced Plastic Mortars"`). C-2
  reparte ahora sus 44 filas como el código: **38 / 5 / 1**.
- **Factores de escala explícitos**, con el texto impreso al lado: ×10³ MPa (C-3),
  ×10⁶ psi (C-3C), ÷10⁶ (C-2). El superíndice se había perdido y confundir 10³ con 10⁶
  son tres órdenes de magnitud en el módulo E.

Se corrige con `scripts/completar_apendice_c.py`, a partir del PDF del Apéndice C y su
OCR estructurado, ambos versionados en `resources/.../appendix_c/`. El PDF es un
escaneo sin capa de texto, así que lo que se parsea es el OCR.

**Ningún valor ni ningún nombre de material se ha tocado** (regla 9): comprobado fila a
fila contra la versión anterior, 0 alteraciones en las 6 tablas. El script audita además
cada número del OCR contra `resources/` y **para si discrepan** en vez de elegir. Caso
real: la C-3 imprime `Type 309.` con **punto** donde sus cinco hermanas llevan coma —una
errata del código, verificada ampliando el escaneo—; el OCR la "corrige" y `resources/`
la conserva.

En el libro solo cambia `DB_NoMetalicos`, que pasa de **931 a 939** pares campo/valor.
Todo lo demás del protocolo de aceptación queda idéntico: 271 276 valores auditados, 22
casos de interpolación, caso semilla `Sa` 138 MPa **APTO**, 0 fallos.
`test_build_db.py::TestApendiceCCompletado` lo protege con 9 pruebas.

---

## Rev. 3 — Dashboard único de navegación

El libro pasa de **37 pestañas planas y todas visibles** a **38 hojas con una sola a la
vista**. Al abrirlo aparece el `Dashboard`: nueve tarjetas con un botón `▸ ABRIR` que
llevan al motor de cálculo del Art. 212, a los siete buscadores o al manual. Cada motor
lleva en su esquina superior un enlace `◂ VOLVER AL DASHBOARD`.

Las 26 bases de datos de materiales —hasta 3 457 filas— dejan de estar a la vista. Son
insumo auditado, no interfaz.

### Estados de visibilidad

| Estado | N.º | Hojas |
|---|---|---|
| `visible` | 1 | `Dashboard` |
| `hidden` | 9 | `Parche_PCC2_Art212`, los 7 `Buscar_*`, `Instrucciones` |
| `veryHidden` | 28 | `Datos_Ref`, las 21 `DB_*`, `MAP_Factores`, `MAP_Grupo`, `Notas_Codigo`, `DB_Listas`, `_meta`, `_Curvas` |

**El estado va grabado en el archivo, no lo impone la macro.** Si abre el libro con las
macros bloqueadas verá solo el `Dashboard`, con un aviso en rojo, y ninguna base de datos
queda expuesta. Las `veryHidden` no aparecen siquiera en el menú «Mostrar» de Excel.

### Lo que cuesta: el libro ya no abre en Google Sheets

La navegación es un proyecto VBA, así que el entregable es `.xlsm` y **hay que habilitar
las macros**. Es el precio de conmutar la visibilidad sin rehacer los ocho motores ya
verificados.

Lo que **no** cambia: las fórmulas siguen sin una sola función de matriz dinámica
(`FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`, `VSTACK`) y las validaciones de datos siguen
apuntando a rangos literales. La restricción de portabilidad se mantiene en el cálculo;
lo que se pierde es solo la capa de navegación.

### Aviso de macros

La celda `A4` del `Dashboard` dice en rojo «MACROS DESHABILITADAS». Si las macros corren,
pasa a verde y dice «MACROS ACTIVAS». Es el indicador de que la navegación está viva.

### Ninguna hoja de cálculo cambió

El motor Art. 212 y los siete buscadores conservan su lógica intacta. Solo se les añadió
el enlace de retorno y se les cambió el estado de visibilidad.

---

## Cascada de selección — 5 niveles + variante

Igual en los siete buscadores y en el motor:

0. **Familia de material** — Acero al carbono, Acero de baja aleación, Acero al níquel (criogénico), Acero inoxidable, Fundición, Aleación de níquel, Aluminio, Cobre, Titanio, Circonio, Otros metales
1. **Composición nominal** — ya filtrada por la familia
2. **Forma de producto**
3. **Especificación (Spec. No.)**
4. **Tipo / Grado** — con esto queda determinada la fila
5. **Variante** (clase / tamaño) — opcional, solo si el código repite ese grado

La familia recorta la lista de composiciones: en la II-D pasa de ~190 entradas a un máximo de 64 (inoxidables) y menos de 40 en casi todas las familias.

> La familia **no es un dato normativo**: es una agrupación de navegación derivada del prefijo UNS —que asigna SAE/ASTM— y, en su defecto, de la composición nominal impresa. No interviene en ningún cálculo; el material sigue identificándose por especificación, grado, forma, UNS, clase y tamaño tal como los imprime la tabla.

**La única celda donde se escribe es la temperatura de consulta**, con su unidad (°C / °F) al lado.

## Corregido: números en la lista de formas de producto

No, no era correcto. Los 138, 145, 142… que aparecían en el desplegable de «Forma de producto» eran los valores de *S admisible* de la tabla de resultados: las columnas auxiliares de esa tabla se solapaban con las columnas de la lista. Al eliminar la tabla de resultados y separar los bloques auxiliares, el desplegable ya solo contiene formas de producto.

## Resultado y ficha: tarjeta, no tabla

Cada buscador es ahora un panel:

- **2 · RESULTADO** — el material que resuelve el filtro, con cuatro indicadores grandes: valor a la temperatura de consulta (con unidad), temperatura consultada, modo de lectura y estado del rango. El estado se pinta en rojo cuando queda fuera de rango.
- **3 · FICHA TÉCNICA** — tarjeta a dos columnas, cada campo con su etiqueta, su valor y **su unidad**: MPa/ksi para resistencias, °C/°F para temperaturas, mm/in para espesores, «adimensional» para P-No. y Poisson.
- **4 · TRAZABILIDAD** — T1, T2 y los valores tabulados usados, más la ecuación aplicada.
- **5 · CURVA** — gráfica del material con **ejes rotulados con unidades**.

No queda ninguna tabla de valores a la vista: la curva sale de una hoja auxiliar oculta.

## Vacío de datos corregido: notas del código

Las notas se cargaban solo desde la clave `notes`. El B31.3 guarda además `general_notes`, y la II-D usa una estructura distinta (`sections` → `items`): **se estaban perdiendo 241 de 325 notas**, incluidas todas las de la II-D. Ahora `Notas_Codigo` carga las 325, con la sección a la que pertenece cada una.

## Verificación completa de las bases

Todas las bases se auditan contra los JSON de `resources/`:

- **Esfuerzos y propiedades por material** (B31.3 A-1/A-4 SI y US, II-D 1A/1B/3 SI y US, U, Y-1 SI y US): de **cada** fila del código se compara su vector completo de valores. **0 filas sin correspondencia.**
- **Bases por familia** (TM-1..5 SI y US, C-1/C-1C, C-3/C-3C): mismo criterio, **0 discrepancias**.
- **Auxiliares**: PRD, TE, MAP_Factores, Notas_Codigo y los 931 pares campo/valor de los no metálicos, conteo exacto.
- **0 funciones de matriz dinámica** y **0 validaciones de lista con origen no portable** en las 37 hojas.
- Bloques de la cascada contiguos en los 5 niveles, en las 8 bases.
- Interpolación, huecos interiores, modo tabulado y bordes recalculados en la hoja y contrastados contra un motor de referencia en Python — 0 fallos.
- 18 pruebas unitarias en verde.
- Regresión del caso semilla: `Sa` 137,9 → **138 MPa** (valor impreso en B31.3-2024); dictamen **APTO**.

## Huecos interiores en las tablas del código

Las tablas ASME no imprimen todos los puntos de la rejilla para cada material: la A-1 **no lista 125 °C para A106 Gr.B**, aunque sí 100 y 150 (160 filas así en la A-1, 68 en la 1A). Junto a la banda impresa —que se conserva para auditar contra el PDF— cada base lleva una **banda compacta** con solo los puntos existentes, y la consulta trabaja sobre ella. A 125 °C interpola entre 100 y 150 °C.

## Pendiente de su validación

`MAP_Grupo` ya no propone nada por su cuenta. El grupo que da el módulo `E` (TM-1) y la dilatación (TE-1) está **impreso en las Notas al pie de esas tablas**, y ahora se lee de ahí: cada fila con grupo **cita la nota que lo sostiene**, y `verificar.py` §9 comprueba fila a fila que esa nota liste realmente esa composición.

| Estado | Filas | Qué significa |
|---|---:|---|
| AUTO (UNS exacto) | 1 292 | El UNS figura literalmente en TM-1…TM-5 |
| AUTO (composición en Nota) | 1 297 | La composición nominal está en la lista de miembros de una Nota |
| REVISAR (regla textual) | 17 | TM-1 Nota (5) dice «9Cr–Mo, *including variations thereof*»: la regla la aplica usted |
| SIN MAPEO | 848 | II-D no publica `E` ni dilatación para ese material; el cálculo queda bloqueado |

**Lo que requiere su firma son 65 decisiones**, no 1 518 filas: están agrupadas por composición nominal en `Revision_MAP_Grupo.md`, con el recuento de materiales que arrastra cada una. De las 848 sin mapeo, 427 ni siquiera imprimen composición nominal en la tabla de origen.

> En la Rev. 2 estas filas se rotulaban con expresiones regulares sobre la composición. Producían asignaciones **falsas**, no solo inciertas: los austeníticos 16Cr-12Ni-2Mo, los dúplex 22Cr-5Ni-3Mo-N y las aleaciones de níquel 62Ni-22Mo-15Cr aparecían como «acero de baja aleación». Esa heurística se eliminó entera.

### Una errata del código, conservada tal como está impresa

Las Notas de TM-1 y TE-1 están ahora en `resources/` en **las dos ediciones**. Al extraer la US apareció que **no numeran igual**: la métrica imprime el Grupo H **dos veces**, en las Notas (8) y (9), y desde ahí toda su numeración va corrida en uno frente a la US, que lo imprime una sola vez en la Nota (8) — 17 notas contra 16.

La errata es exclusiva de la edición métrica y **no se corrige**: los valores se cargan tal como están impresos, y queda anotada en `notes_extraction.observaciones`. La **pertenencia** a grupo sí es idéntica en ambas ediciones, verificada grupo a grupo.

**Enlace SI ↔ US:** la selección se hace sobre la edición métrica. El B31.3 enlaza 1 065 de 1 288 materiales con su homólogo A-1C; la II-D, el 100 % de la 1A. Los no enlazados muestran «sin equivalente en la edición US» en lugar de un valor equivocado.

## Reproducir

Desde `outputs/Base_Datos_Materiales_ASME/scripts`:

```powershell
# 1. Maestro con macros. Solo la primera vez, o al tocar vba/*.vba.
#    Activa momentaneamente "Confiar en el acceso al modelo de objetos de
#    proyectos de VBA" y lo restaura al terminar.
python make_vba_seed.py

# 2. Notas de grupo de TM-1 / TE-1 en resources/, una edicion por corrida.
#    Solo si se repone la extraccion de II-D; el builder aborta si faltan.
python extraer_notas_ii_d.py --edicion si --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Metric 2025\D Metric 2025 _p1201-p1500.pdf"
python extraer_notas_ii_d.py --edicion us --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Customary 2025\D Customary 2025 _p1201-p1500.pdf"

# 3. Entregable
python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm

# 4. Pruebas y protocolo de aceptacion
python -m pytest test_build_db.py test_dashboard.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm
```

`verificar.py` requiere Excel instalado: recalcula el libro con el motor real para
auditar lo que la hoja calcula, no lo que se supone que calcula.

---

*Herramienta de ingeniería de referencia. Verificar entradas y resultados antes de emitir para construcción.*
