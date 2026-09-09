# Dashboard único con navegación por macros — Motor de Cálculo ASME PCC Rev. 3

## Contexto

El libro `outputs/Motor_de_Calculo_ASME_PCC_Rev2.xlsx` tiene **37 hojas planas**, sin
portada ni índice: al abrirlo el usuario cae directamente sobre una fila de pestañas
donde conviven un motor de cálculo, siete buscadores y veintiséis bases de datos de
materiales con hasta 3 457 filas. Las bases nunca deben estar a la vista —son insumo
auditado, no interfaz— y hoy lo están.

El objetivo es que el usuario vea **una sola hoja, el Dashboard**, desde la que
selecciona qué motor quiere usar: el motor de cálculo ASME PCC-2 o uno de los siete
motores de búsqueda. El resto de pestañas permanece oculto.

Decisión tomada (con su consecuencia asumida): **libro con macros, `.xlsm`**. Es la
única forma de conmutar la visibilidad de hojas sin refactorizar los ocho motores ya
verificados. **Rompe la regla nº 1 de `CLAUDE.md`: el entregable deja de abrirse en
Google Sheets** y exige habilitar macros. Las fórmulas siguen siendo portables (cero
matriz dinámica, validaciones por rango literal); lo que se pierde es solo la capa de
navegación. Alcance: se listan **únicamente los motores que existen hoy** (Art. 212 y
los 7 buscadores); no se prometen los artículos 202 / 401 / 402 / 501.

Entrada del builder: `outputs/Base_Datos_Materiales_ASME/Motor_de_Calculo_ASME_PCC_Rev0_respaldo.xlsx`
(el maestro de la raíz no está en el repo).

---

## Arquitectura

```
templates/maestro_con_macros.xlsm      ← paso único, se genera con Excel COM y se versiona
        │  (= Rev0_respaldo + proyecto VBA: ThisWorkbook + modNav)
        ▼
build_db_materiales.py  --in templates/maestro_con_macros.xlsm
                        --out outputs/Motor_de_Calculo_ASME_PCC_Rev3.xlsm
        │  load_workbook(..., keep_vba=True)  → openpyxl conserva vbaProject.bin
        │  + build_dashboard()  + estados de visibilidad
        ▼
outputs/Motor_de_Calculo_ASME_PCC_Rev3.xlsm   (38 hojas, 1 visible)
```

Se descarta la cirugía sobre el ZIP: llevar el VBA en el propio maestro y abrirlo con
`keep_vba=True` deja que openpyxl emita los content-types y la relación de
`vbaProject.bin` por sí solo.

### Estados de visibilidad (se graban en el archivo, no dependen de la macro)

| Estado | Hojas |
|---|---|
| `visible` | `Dashboard` |
| `hidden` (la macro las muestra bajo demanda) | `Parche_PCC2_Art212`, los 7 `Buscar_*`, `Instrucciones` |
| `veryHidden` (nunca accesibles desde la UI) | las 29 restantes: `DB_*`, `MAP_*`, `Notas_Codigo`, `DB_Listas`, `Datos_Ref`, `_meta`, `_Curvas` |

Grabar el estado en el archivo es lo que hace seguro el modo sin macros: si el usuario
las bloquea, no ve ninguna base de datos — solo el Dashboard con un aviso en rojo.

---

## Diseño del Dashboard

Hoja `Dashboard`, primera del libro, sin líneas de cuadrícula
(`ws.sheet_view.showGridLines = False`), pestaña en NAVY, `freeze_panes` desactivado.
Reutiliza la paleta y los helpers que ya existen en `build_db_materiales.py`:
`NAVY 1F3864` · `BLUE 2F5597` · `GREY F2F2F2` · `YELL FFF2CC` · `CARD_FILL EEF3FA` ·
`KPI_FILL DCE6F5`, más `banda()` (:1007), `_mrg()` (:1015) y `campo()` (:1024).

Rejilla de 12 columnas: `A` y `L` como márgenes de 2; `B..K` de contenido; ancho 18
salvo separadores de 3. Columna `N` oculta: guarda la clave de destino de cada tarjeta.

```
┌──────────────────────────────────────────────────────────────────────────┐
│  MOTOR DE CÁLCULO ASME PCC                                       Rev. 3  │  banda NAVY, alto 34
│  Ingeniería de reparación · PCC-2 · B31.3-2024 · BPVC II-D 2025 · SI     │
├──────────────────────────────────────────────────────────────────────────┤
│  ⚠ MACROS DESHABILITADAS — habilítelas para navegar                      │  fila 5, roja; la macro la reescribe
├──────────────────────────────────────────────────────────────────────────┤
│  1 · MOTORES DE CÁLCULO — ASME PCC-2                                     │  banda()
│  ┌──────────────────────────────┐                                        │
│  │ ART. 212                     │                                        │  tarjeta 4 col × 4 filas
│  │ Parche de plancha con        │                                        │  CARD_FILL + CARD_BORDER
│  │ soldadura de filete          │                                        │
│  │ Tubería B31.3 · virola VIII-1│                                        │
│  │ ▸ ABRIR                      │                                        │  celda-botón BLUE, hipervínculo
│  └──────────────────────────────┘                                        │
├──────────────────────────────────────────────────────────────────────────┤
│  2 · MOTORES DE BÚSQUEDA — bases normativas                              │
│  ┌───────────────┐ ┌───────────────┐ ┌───────────────┐                   │  3 tarjetas por fila
│  │ B31.3 · A-1   │ │ II-D · 1A     │ │ II-D · 1B y 3 │                   │
│  │ S admisible   │ │ S admisible   │ │ S admisible   │                   │
│  │ MPa / ksi     │ │ MPa / ksi     │ │ MPa / ksi     │                   │
│  │ ▸ ABRIR       │ │ ▸ ABRIR       │ │ ▸ ABRIR       │                   │
│  └───────────────┘ └───────────────┘ └───────────────┘                   │
│  ┌───────────────┐ ┌───────────────┐ ┌───────────────┐                   │
│  │ II-D · Tabla U│ │ II-D · Y-1    │ │ TM / TE / PRD │                   │
│  │ Su tracción   │ │ Sy fluencia   │ │ E, α, ν, ρ    │                   │
│  │ ▸ ABRIR       │ │ ▸ ABRIR       │ │ ▸ ABRIR       │                   │
│  └───────────────┘ └───────────────┘ └───────────────┘                   │
│  ┌───────────────┐ ┌───────────────┐                                     │
│  │ B31.3 · B y C │ │ MANUAL        │                                     │
│  │ No metálicos  │ │ Instrucciones │                                     │
│  │ ▸ ABRIR       │ │ ▸ ABRIR       │                                     │
│  └───────────────┘ └───────────────┘                                     │
├──────────────────────────────────────────────────────────────────────────┤
│  3 · ESTADO DEL LIBRO                                                    │
│  ┌────────┐ ┌────────┐ ┌────────┐ ┌────────┐                             │  KPI_FILL, patrón de finish_buscador:1216
│  │ 1 291  │ │ 1 802  │ │ 2 476  │ │ 1 518  │                             │
│  │ mater. │ │ mater. │ │ mater. │ │ filas  │                             │
│  │ B31.3  │ │ II-D 1A│ │ Tabla U│ │ por    │                             │
│  │        │ │        │ │        │ │validar │  ← MAP_Grupo, en ámbar      │
│  └────────┘ └────────┘ └────────┘ └────────┘                             │
│  Compilado 2026-09-06 · fuente única: resources/ · valores tal como impresos│
├──────────────────────────────────────────────────────────────────────────┤
│  Herramienta de ingeniería de referencia. Verificar entradas y resultados │  pie GREY
│  antes de emitir para construcción.                                       │
└──────────────────────────────────────────────────────────────────────────┘
```

Reglas de estilo que se respetan: **unidad visible en todo indicador** (regla 6 de
`CLAUDE.md`), conteos escritos por `record_meta`/`build_meta` y no a mano, y ninguna
promesa sobre `MAP_Grupo` — se rotula como pendiente de validación del ingeniero.

Cada tarjeta lleva en la celda `▸ ABRIR` un hipervínculo interno a `Dashboard!A1`
(destino inocuo) y, en la columna oculta `N` de esa misma fila, el nombre de la hoja
destino. El evento de libro intercepta el clic y hace el trabajo; así no hacen falta
formas ni controles, que openpyxl no sabe crear con `OnAction`.

---

## Proyecto VBA

Tres componentes, todos de código (sin formularios), para que sobrevivan al
round-trip de openpyxl.

**`ThisWorkbook`**
- `Workbook_Open` → llama `AplicarVisibilidad`, escribe "MACROS ACTIVAS" en la celda de
  aviso (verde), activa `Dashboard`.
- `Workbook_SheetFollowHyperlink(Sh, Target)` → si `Sh.Name = "Dashboard"`, lee la clave
  de `Cells(fila, "N")` y llama `AbrirHoja`; si la clave es `VOLVER`, llama
  `VolverAlDashboard`.
- `Workbook_BeforeClose` / `Workbook_BeforeSave` → restaura el estado limpio (solo
  Dashboard visible, aviso reseteado a rojo) para que el archivo guardado nunca
  conserve una hoja abierta.

**`modNav`**
- `AplicarVisibilidad()` — aplica la tabla de estados de arriba; es la única fuente de
  verdad de qué se ve.
- `AbrirHoja(nombre)` — oculta la hoja activa anterior, pone `xlSheetVisible` en la
  destino, la activa. Si la hoja no existe, mensaje y retorno.
- `VolverAlDashboard()` — oculta la activa y vuelve.

**Enlace de retorno en cada motor**: el builder escribe `◂ VOLVER AL DASHBOARD` en la
esquina superior de las 9 hojas navegables, con hipervínculo y la clave `VOLVER`. En
`Parche_PCC2_Art212` e `Instrucciones` **hay que escribirlo antes de la protección**
(`build_db_materiales.py:1762` y `:1896`), o la celda queda bloqueada.

---

## Archivos a tocar

| Archivo | Cambio |
|---|---|
| `outputs/Base_Datos_Materiales_ASME/scripts/make_vba_seed.py` | **Nuevo.** Paso único con Excel COM: abre `Rev0_respaldo.xlsx`, inyecta los dos componentes VBA desde `scripts/vba/*.bas`, guarda `templates/maestro_con_macros.xlsm`. Requiere «Confiar en el acceso al modelo de objetos de proyectos de VBA» en el Centro de confianza. |
| `outputs/Base_Datos_Materiales_ASME/scripts/vba/ThisWorkbook.cls`, `modNav.bas` | **Nuevos.** Código fuente VBA versionado en texto, no enterrado en el binario. |
| `outputs/Base_Datos_Materiales_ASME/scripts/build_db_materiales.py` | `load_workbook(a.inp, keep_vba=True)` (:1947); nueva función `build_dashboard(wb, meta)` junto a los buscadores; `link_volver(ws)` aplicado a las 9 hojas navegables **antes** de proteger; `aplicar_visibilidad(wb)` al final; `"Dashboard"` como **primer elemento** de la lista `order` (:2072); llamada a `build_dashboard` antes de `build_meta`. |
| `outputs/Base_Datos_Materiales_ASME/scripts/test_build_db.py` | Tests: existe `Dashboard` y es la hoja 1; exactamente 1 hoja `visible`; 29 `veryHidden`; ninguna DB en estado `hidden`; el Dashboard no usa matriz dinámica ni validación con origen no portable; las 8 claves de la columna `N` corresponden a hojas reales. |
| `outputs/Base_Datos_Materiales_ASME/scripts/verificar.py` | Reparametrizar: `argparse` para `--resources` y `--wb` (hoy `:26-27` apuntan a `/mnt/user-data/uploads/...` y `Motor_v2.xlsx`); sustituir el recálculo con `soffice` (`:437`, **no instalado**) por Excel COM; el caso semilla (`:464`) lee el libro recalculado en vez de `lo2/test_v2.xlsx`. Añadir **sección 8**: auditoría de visibilidad, presencia de `xl/vbaProject.bin` en el ZIP y round-trip de macro. |
| `CLAUDE.md` | 38 hojas; entregable `Motor_de_Calculo_ASME_PCC_Rev3.xlsm`; comandos de reconstrucción nuevos; **enmendar reglas 1 y 2**: la compatibilidad con Google Sheets se pierde en el entregable por la capa VBA, pero la prohibición de matriz dinámica y de validaciones no portables **se mantiene** en todas las fórmulas. |
| `outputs/Base_Datos_Materiales_ASME/LEEME_Nota_de_Version.md` | Sección Rev. 3: el Dashboard, la tabla de visibilidad, y el aviso de macros. |

No se toca ninguna hoja DB, buscador ni el motor Art. 212 en su lógica: solo se les
añade el enlace de retorno y se les cambia el estado de visibilidad.

---

## Orden de ejecución

1. Escribir `vba/ThisWorkbook.cls` y `vba/modNav.bas`.
2. Escribir y correr `make_vba_seed.py` → `templates/maestro_con_macros.xlsm`.
   Verificar con COM que el proyecto tiene 2 componentes y que las 3 hojas del maestro
   siguen intactas.
3. `build_dashboard` + `link_volver` + `aplicar_visibilidad` en el builder; `order`.
4. Reconstruir a `.xlsm` y comprobar con `openpyxl` + `zipfile` que `vbaProject.bin`
   sobrevivió al guardado.
5. Ampliar `test_build_db.py`; reparametrizar `verificar.py`.
6. Actualizar `CLAUDE.md` y la nota de versión.

## Verificación

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts
python make_vba_seed.py                     # solo la primera vez
python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm
python -m pytest test_build_db.py -q
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev3.xlsm
```

`verificar.py` debe seguir devolviendo 0: las 271 276 comparaciones fila a fila contra
los JSON de `resources/`, la contigüidad de la cascada, cero matriz dinámica, cero
validaciones no portables, la interpolación recalculada en hoja (ahora vía Excel) y la
regresión del caso semilla — `Sa` 138 MPa, dictamen APTO — no pueden moverse. Si alguna
se mueve, el Dashboard rompió algo.

Comprobación manual, con el `.xlsm` abierto y macros habilitadas:
- Al abrir solo se ve la pestaña `Dashboard` y el aviso está en verde.
- Los 9 botones `▸ ABRIR` llevan a su hoja; `◂ VOLVER` la oculta de nuevo.
- Ninguna hoja `DB_*` / `MAP_*` aparece en el menú «Mostrar» de Excel (son `veryHidden`).
- Con macros deshabilitadas: solo `Dashboard`, aviso en rojo, ninguna base expuesta.
