# Dashboard jerárquico: navegación real por niveles

> Plan de trabajo. Redactado el 2026-09-08. **EJECUTADO el 2026-09-08.**
> Origen: `Generales/NUEVO LAYOUT DASHBOARD.png`.
>
> Resultado: 69 hojas (1 visible / 36 hidden / 32 veryHidden), `verificar.py` a 0
> con §8 en verde y §1-§7 y §9-§10 **idénticos** al reporte anterior, 190 pruebas
> en verde y la navegación ejercida en Excel real sobre las tres ramas (clic en
> el cuerpo de la tarjeta, retorno al padre, miga de pan, manual y marcador
> inerte).
>
> **Corrección de fondo sobre lo planificado, pedida por el ingeniero:** el plan
> eliminaba las tres bandas por tipo de artefacto y ponía el árbol en la raíz.
> No era eso. El `Dashboard` **conserva** `1 · MOTORES DE CÁLCULO`,
> `2 · MOTORES DE BÚSQUEDA` y `3 · BASES DE DATOS` —esa es la primera pregunta
> de quien abre el libro— y la categorización empieza **dentro** de cada banda:
>
> ```
> 1 · MOTORES DE CALCULO   ASME → REPARACIONES ───→ PCC ─→ PCC-2 → Art. 212
> 2 · MOTORES DE BUSQUEDA  ASME → PIPING ─────────→ B31 ─→ B31.3 → 5 buscadores
>                               → PRESSURE VESSELS → BPVC → SEC. II → 5 buscadores
> 3 · BASES DE DATOS       ASME → PRESSURE VESSELS → BPVC → SEC. II → 9 hojas
> ```
>
> Cada banda tiene su rama completa (prefijos `NAV_CAL_*`, `NAV_BUS_*`,
> `NAV_DAT_*`), así que `ASME` aparece tres veces y `SEC. II` dos: la rama que
> lleva a los buscadores de la Parte D no es la que lleva a las hojas de datos
> de las Partes A, B y C. La cascada repetida se emite desde una sola función
> (`_rama_bpvc`) para que no pueda divergir. Son **15 hojas `NAV_*`**, no 10.
>
> **Dos desviaciones más, forzadas por el entorno:**
>
> 1. Las hojas de la Sección II se llaman **`NAV_BUS_SEC_II` / `NAV_DAT_SEC_II`**,
>    no `NAV_SECII`: ese nombre habría coincidido, letra a letra, con la constante
>    `NAV_SECII` del builder, que es la lista de las **nueve hojas de datos** y
>    significa otra cosa.
> 2. **`HojasNavegables()` ya no usa `Array( _ … )`.** VBA no admite más de 25
>    continuaciones de línea y hacían falta 30: `AddFromString` rechazó el módulo,
>    lo dejó **vacío**, y el síntoma visible fue un «No se ha definido Sub o
>    Function» al guardar, en un diálogo modal que colgó Excel y dejó el maestro
>    borrado. La lista se arma concatenando (`s = s & "|…"` + `Split`), que no
>    tiene tope, y `lint_vba()` comprueba ahora ese límite antes de tocar COM —
>    que es el sitio donde debía estar.

## Contexto

Hoy el Dashboard organiza las 21 hojas alcanzables por **tipo de artefacto**:
`1 · MOTORES DE CALCULO`, `2 · MOTORES DE BUSQUEDA`, `3 · BASES DE DATOS`. Es una
clasificación de implementación, no de ingeniería: no dice de qué código sale cada
motor, mezcla B31.3 con II-D en la misma banda, y no tiene dónde crecer. El libro ya
cubre tres códigos y va a cubrir más (B31.1, B16.5, B16.9, SEC. VIII, PCC-1/PCC-3).

`Generales/NUEVO LAYOUT DASHBOARD.png` propone reorganizarlo por el árbol con el que
se cita una norma, y el usuario fijó la profundidad:

> **PUBLICANTE > DISCIPLINA > CÓDIGO APLICABLE A ESA DISCIPLINA (SI APLICA) > STANDARD O CÓDIGO ESPECÍFICO**

con navegación real: cada tarjeta es un botón que abre el nivel siguiente.

Resultado buscado: el Dashboard deja de ser un índice plano y pasa a ser la raíz de un
árbol de navegación. La estructura queda declarada en **un solo sitio** y admite
ramas nuevas sin tocar el motor de navegación.

---

## El árbol

```
Dashboard ─ PUBLICANTE
├── ASME ......................................... NAV_ASME
│   ├── PIPING ................................... NAV_PIPING
│   │   ├── B31 · Codigo de tuberia a presion ..... NAV_B31
│   │   │   ├── B31.3 · Process Piping ............ NAV_B31_3
│   │   │   │   ├── Buscar_B31_3         Tablas A-1 y A-4 · S admisible
│   │   │   │   ├── Buscar_Prop_B31_3    Apendice C · dilatacion y modulo
│   │   │   │   ├── Buscar_NoMetalicos   Apendice B · no metalicos
│   │   │   │   ├── Buscar_Ec_A2         Tabla A-2 · factor Ec
│   │   │   │   └── Buscar_Ej_A3         Tabla A-3 · factor Ej
│   │   │   └── ▨ B31.1 · B31.4 · B31.5 · B31.8 · B31.9 · B31.12
│   │   └── ▨ B16 · Bridas, accesorios y valvulas
│   ├── PRESSURE VESSELS ......................... NAV_PVESSELS
│   │   └── BPVC ................................. NAV_BPVC
│   │       ├── SEC. II · Materiales ............. NAV_SECII
│   │       │   ├─ «PARTE D · propiedades de diseno»
│   │       │   │   Buscar_BPVC_IID · Buscar_BPVC_IID_B · Buscar_Su
│   │       │   │   Buscar_Sy · Buscar_Prop_IID
│   │       │   ├─ «PARTES A, B y C · especificaciones»
│   │       │   │   DB_SecII_A1 · DB_SecII_A2 · DB_SecII_B · DB_SecII_C
│   │       │   └─ «TRANSVERSAL A LAS PARTES»
│   │       │       CAT_SecII · IDX_SecII_Tablas · DB_SecII_Notas
│   │       │       DB_SecII_Quimica · DB_SecII_Traccion
│   │       └── ▨ SEC. VIII Div. 1
│   └── REPARACIONES ............................. NAV_REPARACION
│       └── PCC .................................. NAV_PCC
│           ├── PCC-2 · Reparacion de equipos a presion ... NAV_PCC2
│           │   └── Parche_PCC2_Art212   Art. 212 · parche de plancha
│           └── ▨ PCC-1 · PCC-3
└── Instrucciones ................................ MANUAL DE USO
```

`▨` = **tarjeta marcador**: gris, sin hipervínculo ni clave, rotulada
`NO CARGADO EN ESTE LIBRO`. Solo un rótulo de documento — cero dato normativo, así
que no roza la Regla nº 1. Su función es que un nivel con un solo hijo cargado
(`NAV_PIPING`, `NAV_BPVC`, `NAV_PCC`) siga explicando la taxonomía en vez de parecer
una pantalla vacía, y que se vea de un vistazo qué falta por cargar.

**10 hojas de navegación nuevas** (`NAV_*`) sobre las 54 actuales → 64 hojas.
Alcanzables (`NAVEGABLES`): 10 NAV + los 21 destinos de hoy = **31**.

`Instrucciones` cuelga de la raíz **y además** tiene un botón fijo `? MANUAL DE USO`
en la cabecera de toda hoja `NAV_*`: es transversal a las normas y debe estar
siempre a un clic.

Los KPI (`4 · ESTADO DEL LIBRO`) y el pie de responsabilidad se quedan en el
Dashboard, que sigue siendo la única portada.

---

## Cambio de mecanismo: la clave es siempre el destino

Hoy hay **dos** clases de clave: el nombre de la hoja a abrir (botones del Dashboard)
y el literal `"VOLVER"` (retorno de cada motor), que el VBA resuelve siempre al
Dashboard. Con un árbol de cinco niveles eso ya no sirve: `VOLVER` tiene que llevar al
**padre**, no a la raíz.

La solución que menos maquinaria añade es unificar: **la celda oculta guarda siempre
el nombre de la hoja destino**, se esté bajando o subiendo. `CLAVE_VOLVER` desaparece
y su papel lo ocupa `PADRE`, un mapa derivado del árbol en Python. El VBA se
**simplifica** —una sola rama— y, lo que más importa, **deja de crecer con el árbol**:
añadir B31.1 mañana no toca ni una línea de VBA.

| | Hoy | Después |
|---|---|---|
| Clave de una tarjeta | hoja destino | hoja destino (igual) |
| Clave del retorno | `"VOLVER"` | hoja **padre**, según `PADRE` |
| Rama en el VBA | `If clave = VOLVER … ElseIf Sh = Dashboard …` | `IrAHoja clave, Sh` |
| Ocultar al navegar | siempre el `Dashboard` | la hoja de origen (`Sh`) |
| Cortafuegos | `EsNavegable(destino)` + origen es Dashboard | `EsNavegable(destino)` o destino es Dashboard |

El guardarraíl de origen (`Sh.Name = HOJA_INICIO`) se retira **documentando por qué**:
era redundante: `EsNavegable` ya impide que una clave corrompida destape una `DB_*` o
una `MAP_*`, y con el árbol el origen legítimo deja de ser una sola hoja.

**Migas de pan clicables**: cada segmento de la ruta (`ASME › PIPING › B31`) es una
celda con hipervínculo y clave del ancestro correspondiente. Sale gratis con el mismo
mecanismo y evita que subir cuatro niveles cueste cuatro clics.

---

## Archivos y cambios

### 1. `outputs/Base_Datos_Materiales_ASME/scripts/build_db_materiales.py`

Bloque a reescribir: **líneas 5568-5918** (`DASH` … `aplicar_visibilidad`), más
`TARJETAS_SECII` (5717-5740) y la llamada de `main()` (6171-6199).

**Declaración única del árbol.** Sustituye a `TARJETAS_SECII` y a la lista local
`buscadores` (5790-5818):

```python
class Nodo(NamedTuple):
    hoja: str | None        # hoja NAV_* propia; None si es una tarjeta destino
    titulo: str
    lineas: tuple[str, ...] # dos lineas de detalle
    destino: str | None     # hoja final del libro
    grupo: str | None       # rotulo de banda dentro del padre (SEC. II)
    hijos: tuple = ()
    cargado: bool = True    # False -> tarjeta marcador, sin clave ni enlace
```

Derivados, calculados una vez en preorden — nada escrito a mano dos veces:

- `HOJAS_NAV` — `Dashboard` + las 10 `NAV_*`.
- `NAVEGABLES` — las 10 `NAV_*` + los 21 destinos, **en orden de preorden**.
  Sigue siendo la lista que `test_dashboard.py::TestSincroniaPythonVba` compara,
  carácter a carácter y en orden, contra `HojasNavegables()` del VBA.
- `PADRE` — `{hoja: hoja del padre}`, con `PADRE[NAV_ASME] = "Dashboard"`.
  Sustituye a `CLAVE_VOLVER` como fuente del botón de retorno.
- `ANCLA_VOLVER` — se deriva: `(3, 1, 3)` para todo, salvo la excepción ya existente
  de `Instrucciones`, `(1, 2, 3)`.

**Tarjeta clicable entera.** `_tarjeta` (5650) hoy pone hipervínculo y clave solo en
la 4.ª fila (`▸ ABRIR`). Pasa a ponerlos en **las cuatro filas**: cada fila es una
celda fusionada `c1..c2`, y la clave va en `(fila+i, COL_CLAVE_BASE + c1)`, distinta
para cada `i`, así que sigue cumpliendo `test_las_claves_no_se_pisan`. El cursor pasa
a mano sobre toda la tarjeta y el clic funciona en cualquier punto. La barra inferior
se conserva como señal visual, con el texto según el nivel: `▸ ENTRAR` en un nodo
interior, `▸ ABRIR` en un destino.

Reutilizar `_boton` (5622) partiéndolo en dos: `_enlace(ws, fila, c1, c2, clave)`
—hipervínculo señuelo a `'Dashboard'!A1` + clave, sin tocar el estilo— y el pintado
de la barra azul aparte. El orden **hipervínculo → estilo** de la línea 5632 es
obligatorio y su comentario debe sobrevivir: Excel pinta la celda con el estilo
`Hyperlink` al asignarlo y taparía el botón.

**Constructor genérico.** `build_dashboard` (5743) se reemplaza por:

- `build_nav(wb, nodo, kpis, fecha)` — construye la hoja de **un** nodo interior y se
  llama recursivamente. Layout común, reutilizando `banda()` (3224) y `_mrg()` (3232):

  | Fila | Contenido |
  |---|---|
  | 1 | Título del nodo, banda `TITLE_FILL` |
  | 2 | Subtítulo / qué contiene, misma banda |
  | 3 | Barra de acciones: `◂ VOLVER A <padre>` en A3:C3 · `? MANUAL DE USO` en J3:L3 |
  | **4** | **Dashboard**: aviso de macros (`FILA_AVISO`, acoplado a `CELDA_AVISO="A4"`). **NAV_\***: miga de pan clicable |
  | 6+ | Bandas de grupo (si las hay) y tarjetas, 3 por fila, columnas ancla 1/5/9 |

  La fila 4 del Dashboard **no se mueve**: `CELDA_AVISO = "A4"` del VBA depende de
  ella y `test_misma_celda_de_aviso` lo comprueba.
- `build_arbol(wb, kpis, fecha)` — recorre el árbol, crea el Dashboard y las 10
  `NAV_*`, y devuelve el orden de hojas.

**`link_volver` (5884)** deja de escribir `CLAVE_VOLVER` y escribe `PADRE[nombre]`.
Sigue desbloqueando la celda (`Protection(locked=False)`) en las dos hojas protegidas.

**`aplicar_visibilidad` (5902)** no cambia: `Dashboard` visible, `NAVEGABLES` (que
ahora incluye las `NAV_*`) `hidden`, el resto `veryHidden`.

**`main()` (6185-6197)**: `order` empieza por `DASH` + `HOJAS_NAV[1:]` y sigue igual.
`Dashboard` debe quedar en el índice 0 (`test_dashboard_es_la_primera_hoja`).

### 2. `outputs/Base_Datos_Materiales_ASME/scripts/vba/mod_nav.vba`

- `HojasNavegables()` (70-93): pasa de 21 a **31** entradas, en el orden de preorden de
  `NAVEGABLES`. Es el único punto que crece con el árbol.
- `AbrirHoja` (135-165) → `IrAHoja(destino, origen)`: acepta `HOJA_INICIO` como
  destino válido y oculta `origen` en vez del Dashboard siempre. Se conserva el orden
  mostrar-antes-de-ocultar (Excel se niega a ocultar la última hoja visible).
- `VolverAlDashboard` (168-182) se elimina: `IrAHoja` cubre subir y bajar.
  `RestaurarEstadoLimpio` (208) no la usa.
- `CLAVE_VOLVER` (17) se elimina.
- Regla del proyecto que hay que respetar: **toda `Const` va antes de la primera
  rutina**, o `lint_vba()` falla y Excel abre un diálogo modal al guardar.

### 3. `outputs/Base_Datos_Materiales_ASME/scripts/vba/this_workbook.vba`

`Workbook_SheetFollowHyperlink` (38-53) se reduce a leer la clave y llamar a
`IrAHoja clave, Sh`. Los otros tres eventos no cambian.

### 4. `outputs/Base_Datos_Materiales_ASME/scripts/test_dashboard.py`

- `TestBotones::test_el_dashboard_enlaza_las_nueve_hojas_navegables` (115) → pasa a
  recorrer **todas** las hojas de navegación: la unión de sus claves debe ser
  exactamente `set(NAVEGABLES)`. Comparación por **conjunto**, no por lista ordenada:
  la tarjeta clicable entera produce cuatro claves iguales por tarjeta.
- `test_ninguna_clave_es_huerfana` (119) → mismo recorrido, sobre todo el árbol.
- `test_el_hipervinculo_apunta_a_destino_inocuo` (128) → mismo recorrido; el literal
  `'Dashboard'!A1` no cambia.
- `test_cada_navegable_tiene_enlace_de_retorno` (134) → `B.PADRE[nombre] in claves`.
- `test_el_retorno_esta_desbloqueado_en_las_hojas_protegidas` (139) → el botón se
  identifica por `k == B.PADRE[nombre]` en vez de por `CLAVE_VOLVER`.
- `TestSincroniaPythonVba::test_misma_hoja_de_inicio_y_clave_de_retorno` (214) →
  se queda solo con `HOJA_INICIO`.
- **Clase nueva `TestArbolDeNavegacion`**, que es donde está el valor real:
  1. cada `NAV_*` enlaza **exactamente** a sus hijos, a su padre y al manual;
  2. todo destino de `NAVEGABLES` se alcanza desde el `Dashboard` recorriendo claves
     (el árbol es conexo, ningún motor queda huérfano);
  3. `PADRE` no tiene ciclos y toda cadena termina en `Dashboard`;
  4. las tarjetas marcador **no** llevan hipervínculo ni clave.

### 5. `outputs/Base_Datos_Materiales_ASME/scripts/verificar.py`

- **§8** (1156-1208): la fila «Los botones cubren las N hojas navegables» (1190) pasa
  a recorrer el árbol; la de «Cada hoja navegable tiene enlace de retorno» (1193)
  comprueba contra `B.PADRE`. Se añade la comprobación de conexidad —todo destino
  alcanzable desde el Dashboard—, que es la que de verdad protege el rediseño.
- **§10** (1378-1387): la lista `esperadas_secii` está escrita a mano y duplica
  `B.NAV_SECII`. Sustituirla por `B.NAV_SECII` de paso, ya que se toca la zona.

### 6. `CLAUDE.md`

La sección **«Capa de navegación — dos fuentes de verdad que no pueden divergir»**
describe el mecanismo que cambia. Actualizar la tabla (desaparece `CLAVE_VOLVER`,
entra `PADRE`; `COL_CLAVE_BASE` y `FILA_AVISO` siguen igual) y la descripción del
Dashboard en «Motor de cálculo — estado actual» (21 → 31 hojas alcanzables,
54 → 64 hojas, y el árbol como criterio de organización).

---

## Fuera de alcance, declarado

- **Ninguna banda de atajos.** Llegar a un motor cuesta cuatro clics; añadir una
  banda con los 21 destinos en el Dashboard reintroduciría el índice plano que este
  cambio elimina. Si molesta en uso diario, se decide después, con el libro delante.
- **Ningún dato normativo se toca.** Esto es capa de presentación: ni un valor, ni una
  fórmula, ni un JSON de `resources/`. `verificar.py` §1-§7 y §9-§10 deben salir
  idénticos.
- **Las tarjetas marcador no cargan nada.** Son rótulos; B31.1 y compañía siguen sin
  estar en el libro.

---

## Verificación

Desde `outputs\Base_Datos_Materiales_ASME\scripts`, **en este orden**:

```powershell
# 1. Resembrar el maestro: cambian los dos .vba. Activa AccessVBOM unos segundos
#    y lo restaura; si avisa por stderr de que no pudo, apagarlo a mano despues.
python make_vba_seed.py

# 2. Reconstruir (~45 s)
python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm

# 3. Pruebas
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q

# 4. Auditoria completa (varios minutos; exige Excel instalado)
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```

`verificar.py` devuelve 0 solo si todo pasa. Criterios concretos:

- **§8** en verde con 31 navegables, y la nueva fila de conexidad del árbol.
- **§1-§7 y §9-§10 idénticos** al reporte anterior: si un valor tabulado cambia,
  algo se ha tocado que no debía tocarse.
- Conteo de hojas: **64**, con 1 visible / 31 hidden / 32 veryHidden.

**Comprobación en Excel real** (ni pytest ni `verificar.py` ejercen el clic; el
mecanismo del hipervínculo señuelo solo se ejerce de verdad en Excel):

1. Abrir el `.xlsm` con las macros **bloqueadas**: solo el `Dashboard`, aviso en rojo,
   ninguna `NAV_*` ni ninguna `DB_*` en el menú Mostrar.
2. Habilitar macros: el aviso pasa a verde.
3. Bajar `ASME › PIPING › B31 › B31.3 › Tabla A-1` haciendo clic **en el cuerpo de la
   tarjeta**, no solo en la barra `▸ ABRIR`. En cada paso debe haber una sola pestaña
   a la vista.
4. Subir con `◂ VOLVER`: tiene que devolver al **padre**, no al Dashboard.
5. Desde `NAV_B31_3`, clic en el segmento `PIPING` de la miga de pan: salta dos
   niveles de golpe.
6. `? MANUAL DE USO` desde una `NAV_*` cualquiera abre `Instrucciones`; su `VOLVER`
   devuelve al Dashboard.
7. Una tarjeta marcador (`B31.1`) **no hace nada** al pulsarla.
8. Cerrar sin guardar y reabrir: el libro vuelve a mostrar solo el `Dashboard`
   (`Workbook_BeforeClose` → `RestaurarEstadoLimpio`).
