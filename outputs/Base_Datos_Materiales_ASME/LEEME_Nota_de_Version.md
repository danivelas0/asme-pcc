# Nota de versión — Motor de Cálculo ASME PCC, Rev. 3
## Base de datos de materiales (PLAN-DB-MAT-001)

**Fecha:** 2026-09-06 · **Fuente única de verdad:** `resources/` · **Entregable:** `Motor_de_Calculo_ASME_PCC_Rev3.xlsm`

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

**Enlace SI ↔ US:** la selección se hace sobre la edición métrica. El B31.3 enlaza 1 065 de 1 288 materiales con su homólogo A-1C; la II-D, el 100 % de la 1A. Los no enlazados muestran «sin equivalente en la edición US» en lugar de un valor equivocado.

## Reproducir

Desde `outputs/Base_Datos_Materiales_ASME/scripts`:

```powershell
# 1. Maestro con macros. Solo la primera vez, o al tocar vba/*.vba.
#    Activa momentaneamente "Confiar en el acceso al modelo de objetos de
#    proyectos de VBA" y lo restaura al terminar.
python make_vba_seed.py

# 2. Notas de grupo de TM-1 / TE-1 en resources/. Solo si se repone la
#    extraccion de II-D; el builder aborta si faltan.
python extraer_notas_ii_d.py --resources ..\..\..\resources `
    --pdf "<...>\SECCION II\D Metric 2025\D Metric 2025 _p1201-p1500.pdf"

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
