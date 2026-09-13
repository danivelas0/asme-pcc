# Skill `motor_pcc2` — un motor de cálculo para cualquier artículo de ASME PCC-2

> **Diseño, no plan de implementación.** De aquí sale después el plan por fases.
> **NO EJECUTAR sin orden explícita del ingeniero** (regla general de `outputs/plans/`).
>
> Fecha: 2026-09-13 · Estado: diseño aprobado en conversación, pendiente de revisión escrita.

## El problema

Construir el motor del Art. 212 y el del Art. 206 costó tres revisiones, dos planes
completos y una lista larga de errores que solo se ven mirando la hoja en Excel. Casi
todo lo que costó **no era del artículo**: era el andamiaje —cascada de material,
conmutador SI/US, semáforo, dictamen, reinicio, plegado, navegación, impresión,
comentarios, decimales— y las trampas de openpyxl.

PCC-2 tiene **38 artículos**. Repetir ese trabajo a mano 36 veces no es viable, y
copiar-pegar el 206 produciría 36 copias de 1 400 líneas con direcciones literales: un
generador de deuda.

La skill existe para que **lo que costó se pague una vez**, y para que el trabajo de cada
artículo nuevo sea el único que no se puede automatizar: decidir qué dice el código.

## Lo que NO resuelve

- **No decide la ingeniería.** Qué ecuación gobierna el espesor, qué se verifica y qué se
  teclea es criterio del ingeniero. La skill lee, propone y **se detiene**.
- **No inventa un valor normativo.** Ni uno. Ver «Trazabilidad».
- **No toca `build_parche_art212` ni `build_collar_art206`.** Ver «Por qué los dos
  motores viejos se quedan como están».
- **No edita el `.xlsm` a mano.** El libro se genera por script y sigue siendo
  reproducible.

## Decisiones tomadas

| Decisión | Elegido | Por qué |
|---|---|---|
| Por dónde entra el motor nuevo | **Por el builder** | El libro sigue reproducible: se regenera desde cero, `verificar.py` lo audita entero y el motor nuevo hereda las 16 secciones de verificación. Inyectar la hoja en el `.xlsm` la dejaría fuera de la cadena y el siguiente build la borraría. |
| Autonomía | **Guiada por fases**, con una sola parada | Decidir qué ecuación gobierna es criterio de ingeniería, y un motor nuevo **no tiene oracle**: nada contradiría al modelo si se equivoca. Lo mecánico va entero automático. |
| Alcance | **Los 38 artículos, una sola forma de motor** | En un artículo sin ecuaciones (214, 310) la sección de cálculo queda pequeña y el peso se va a los pasos y a las verificaciones, que siguen teniendo semáforo y dictamen. Un solo producto que mantener. |
| Fuente | **`resources/` directamente** | Regla nº 1. Ni `knowledge/`, ni el PDF, ni la memoria del modelo. |

## Arquitectura

### El motor nuevo nace declarativo

La skill no escribe 1 400 líneas de `calc("D45", "=...")`. Escribe una **declaración** del
artículo, y un **chasis** la convierte en hoja:

```
resources/ASME PCC/pcc_2/.../art_211.json
        │  (Fase 1: lectura + parada de ingeniería)
        ▼
outputs/Base_Datos_Materiales_ASME/scripts/motores/art_211.py     ← la DECLARACIÓN
        │  (Fase 2-6: chasis)
        ▼
build_motor_declarado(wb, MOTOR_211)  ──►  hoja del libro
        │
        ├─ construir_seccion7_material()   ─┐
        ├─ aplicar_unidades_motor()          │  pases transversales que YA existen
        ├─ aplicar_reglas_de_comentario()    │  y no se reescriben
        ├─ aplicar_semaforo_motor()          │
        ├─ aplicar_dos_decimales()           │
        ├─ build_reinicio_motor()            │
        ├─ build_manifiesto_cascada()        │
        ├─ aplicar_leyenda_motor()           │
        └─ preparar_impresion()            ─┘
```

### Las fórmulas se escriben por NOMBRE, no por dirección

Es el mecanismo que hace posible todo lo demás. En la declaración, una fila se refiere a
otra por su clave:

```python
Fila("t_req", "Espesor requerido", magnitud="len",
     formula="={P}*{OD}/(2*({S}*{E}+{P}*{Y}))",
     cita="211-3.2 ec. (1)", bloque=34),
```

El chasis **asigna las filas** y sustituye `{P}` por `$D$29`, `{OD}` por `$D$21`, etc. Tres
consecuencias que importan:

1. **No hay direcciones literales que mantener.** La clase de error que obligó al mapa de
   filas de la Fase 3 —y que dejó tres defectos latentes— deja de existir.
2. **Añadir una fila no rompe nada.** El chasis reenumera y repunta.
3. **Una referencia a una fila que no existe aborta el build** en vez de producir un
   `#REF!` que nadie mira.

### La declaración, en concreto

```python
MOTOR_211 = Motor(
    articulo="211",
    hoja="Recargue_PCC2_Art211",
    titulo="MOTOR DE CALCULO — RECARGUE POR SOLDADURA (ASME PCC-2 Art. 211)",
    fuente="ASME PCC/pcc_2/p2_welded_repairs/art_211_weld_buildup_weld/art_211.json",
    aplicacion=SELECTOR_GEOMETRIA,        # o un código fijo, si el artículo no varía
    casos=("Operacion", "Diseno"),        # 1 o 2 columnas de caso
    material=Material(columnas=(("D", "Metal base"),), incluir_ej_ec=False),
    secciones=(...),                      # bandas y filas, en orden de lectura
    verificaciones=(...),                 # requerido / adoptado / favorables / criterio
    dictamen=Dictamen(compuertas=(...), verificaciones_and=(...)),
    pasos=(...),                          # el anexo del flujo, un bloque por paso
    especificaciones=(...),               # concepto / texto o fórmula / cláusula
)
```

Cada `Fila`, `Verificacion` y `Especificacion` lleva **`cita` y `bloque`**: la cláusula
impresa y el número de bloque del JSON del que salió.

### Por qué los dos motores viejos se quedan como están

`build_parche_art212` y `build_collar_art206` **no se migran al chasis**. Son lo único
verificado del libro: el 212 contra su *oracle* celda a celda y los dos contra §6e/§6f, que
recalculan sus ecuaciones en Excel real. Meterlos en un chasis nuevo sería cambiar lo
verificado por lo elegante, y este proyecto ya pagó ese tipo de cambio tres veces.

El chasis **nace probado por los motores nuevos**. Si algún día conviene, se migran de uno
en uno, con su §6 delante.

## Las nueve fases (0 a 8)

| Fase | Entrada | Salida | Automática |
|---|---|---|---|
| **0. Localizar** | nº de artículo | ruta del `art_XXX.json` en los dos espejos; aborta si no está | sí |
| **1. Leer y proponer** | el JSON entero | inventario: ecuaciones con su número, figuras con imagen, umbrales de doble unidad, límites y prohibiciones, requisitos de fabricación/examen/prueba, tablas. Y una propuesta de entradas, cálculos, verificaciones y pasos | **no: aquí para** |
| **2. Declarar** | lo que el ingeniero marcó | `motores/art_XXX.py` con la declaración y su procedencia | sí |
| **3. Chasis** | la declaración | la hoja, con material, unidades, semáforo, reinicio, cascada, leyenda, decimales e impresión | sí |
| **4. Navegar** | la declaración | nodo en `ARBOL` (artículo = nivel con motor + especificaciones + instrucciones) y línea en `HojasNavegables()` | sí |
| **5. Pestañas** | la declaración | especificaciones técnicas citando párrafo a párrafo, e instrucciones derivadas del motor | sí |
| **6. Pruebas** | la declaración | clase propia en `test_dashboard.py`, por cadena y sin oracle (patrón del 206) | sí |
| **7. Verificar** | el libro | sección propia en `verificar.py`; ver «Qué prueba qué» | sí |
| **8. Entregar** | el libro | `pytest`, `verificar.py`, exportar a PDF y mirarlo | sí |

La Fase 1 es la única parada, y es deliberada. Lo que presenta no es un resumen: es lo que
el artículo **imprime**, con el número de bloque al lado para que se pueda ir a mirar.

## Trazabilidad: la Regla nº 1, operacionalizada

Que la skill «lea `resources/`» no puede ser una promesa. Se convierte en mecanismo:

1. **Toda celda de cálculo, verificación o especificación nace con su procedencia**
   (archivo, bloque, cláusula).
2. Esa procedencia se escribe **tres veces y en tres formas**: la columna de referencia
   que ve el ingeniero, el comentario de la celda, y una tabla de trazabilidad del motor.
3. **El build aborta** si una celda de cálculo llega sin procedencia. Un número que nadie
   puede rastrear no se construye.
4. Una sección de `verificar.py` comprueba que **cada cita existe de verdad** en el JSON
   que dice: el mismo guardia que la §9 hace con `MAP_Grupo`.

### Los tres casos que ya costaron caro

- **Ecuación vacía con imagen.** El bloque `equation` llega sin texto y el `figure` trae
  `image: figures/….png`. Esa imagen **es parte de `resources/`**: se lee con `Read` sobre
  el PNG. Es lo que hubo que hacer con el cateto del 206 y la ec. (2) del 212.
- **Umbral de doble unidad.** «40 mm (1.5 in.)» se guarda como **dos cifras**, nunca
  convirtiendo una en la otra (1,5 in son 38,1 mm, y esa diferencia es del código). Mismo
  mecanismo que `leer_umbrales_pcc2()`.
- **Extracción colapsada.** Si el bloque llegó roto —como el App. 501 o la Tabla
  302.3.4-1—, la skill **no lo rellena**: declara el vacío, dice qué archivo re-extraer y
  se detiene.

## Qué prueba qué

Hay que ser honesto sobre el límite: un chasis genérico puede probar la **maquinaria**, no
la **ingeniería**.

| Nivel | Qué garantiza | Cómo |
|---|---|---|
| `pytest` sobre la declaración | que la hoja tiene lo que la declaración dice: secciones en orden, cada entrada editable, cada verificación con semáforo, cada cita con bloque | clase por motor, por cadena (patrón del 206) |
| `pytest` sobre el libro | lo transversal que ya se comprueba hoy: paleta, fuentes, amarillo acotado, manifiestos, comentarios en columna de valor, dos decimales, navegación | las clases que ya existen, extendidas al motor nuevo |
| `verificar.py` genérico | que el caso semilla **recalcula sin errores** (`#REF!`, `#NAME?`, `#VALUE!`), que **todo veredicto pinta** (DisplayFormat), que el dictamen **no contradice** a sus propias verificaciones, y que **toda cita existe** en el JSON | sección nueva, genérica para todo motor declarado |
| `verificar.py` por artículo | que **el número es el que pide el código** | lo escribe el ingeniero cuando quiere esa red, como §6e/§6f. **La skill no lo genera**: sería el modelo comprobándose a sí mismo |

Esa última fila es el límite declarado de la skill, y conviene que esté escrito.

## La skill

**Alcance del proyecto, no personal.** Vive en `.agents/skills/motor_pcc2/`, dentro del
repositorio y versionada con él —no en `~/.claude/skills/`—: la skill solo tiene sentido
con este `resources/`, este builder y estas reglas, y quien clone el repositorio se la
lleva. `.claude/skills/` son junctions y no se versionan (ver `.gitignore`).

Contenido:

- `SKILL.md` — las nueve fases y la parada.
- `referencias/reglas_del_libro.md` — el catálogo que el libro ya pagó: las 14 reglas de
  diseño, la Regla nº 1, las tres del flujo de la cascada, el semáforo con `bgColor`, los
  comentarios solo en columna de valor, el ancla del VML, los dos decimales.
- `referencias/plantilla_declaracion.py` — la declaración comentada, con el 206 escrito en
  ella como ejemplo de referencia (**sin migrar el motor real**).

## Riesgos

| Riesgo | Mitigación |
|---|---|
| El chasis no cubre un artículo con una forma que no previmos | La declaración admite una sección de «filas libres» donde el chasis emite lo que se le diga. Si un artículo necesita más, se escribe a mano: la skill no es una camisa de fuerza |
| El modelo propone una ecuación que no gobierna | La parada de la Fase 1, y que todo lleve su bloque para poder ir a mirarlo |
| El chasis se rompe y afecta a los motores nuevos ya entregados | Cada motor declarado tiene su clase de pruebas; el chasis no puede cambiar sin que fallen |
| Deriva entre el chasis y los dos motores a mano | Los pases transversales son **los mismos objetos**, no copias. Si cambia el semáforo, cambia para los cinco |

## Criterios de aceptación

1. `/motor_pcc2 211` produce un motor completo en el libro, con `pytest` y `verificar.py`
   en 0 fallos, sin que nadie escriba una dirección de celda.
2. Toda celda de cálculo del motor nuevo cita su cláusula y su bloque, y `verificar.py`
   comprueba que esa cita existe en el JSON.
3. El 212 y el 206 salen del proceso **byte a byte iguales**: su *oracle* y sus §6e/§6f
   siguen en verde.
4. Un artículo sin ecuaciones (p. ej. 214) produce un motor con pasos y verificaciones,
   sin sección de cálculo vacía de sentido.
5. La Fase 1 se detiene **siempre**. Un motor construido sin que nadie mirara la propuesta
   es un motor que nadie firmó.
