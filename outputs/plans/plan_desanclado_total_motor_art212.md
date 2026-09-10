# Desanclado total del motor Art. 212 del Rev0 — Plan de implementación

> **Para ejecutores:** SUB-SKILL REQUERIDA: usar `superpowers:subagent-driven-development`
> (recomendado) o `superpowers:executing-plans` para ejecutar tarea por tarea. Los pasos
> usan casillas (`- [ ]`) para el seguimiento.

**Objetivo:** Que la hoja `Parche_PCC2_Art212` nazca 100 % en código (como ya hace
`Collar_PCC2_Art206`) y el maestro Rev0 deje de aportarla, sin cambiar una sola fórmula
ni valor de la hoja ya validada en Excel.

**Arquitectura:** Se captura la hoja 212 actual del `.xlsm` real a un *oracle* JSON
versionado (fórmulas y valores, celda a celda). Se escribe `build_parche_art212()` que
reconstruye las secciones 1-6 en código, reutiliza `construir_seccion7_material()` (ya
compartida con el 206) para la Sección 7 y transcribe el cableado D39/D40/D44/F90, y que
el `main()` llama **en vez de** `integrate_motor()`. Un test de caracterización de hoja
entera, que compara lo construido contra el *oracle*, es la red que garantiza la
equivalencia **en este entorno, sin Excel** — el mismo mecanismo que hoy ya protege la
Sección 7 (`TestParidadSeccion7Art212`), extendido a las filas 5-101.

**Stack:** Python 3 + openpyxl; pytest; el builder `build_db_materiales.py`; Excel de
Windows **solo** del lado del ingeniero para la validación final (§6+ de `verificar.py`,
F9 del caso semilla).

**Spec:** `outputs/plans/plan_motor_art206_collar_y_separacion_parche_collar.md`, Fase 4,
opción (a) «desanclado total». Decisión de alcance del ingeniero: **Total** ("Project
scope").

## Global Constraints

Copiadas literalmente de `asme-pcc/CLAUDE.md` → «Reglas de diseño del libro». Todas las
tareas las incluyen implícitamente:

- **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`,
  `VSTACK`, `_xlfn`. Solo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`. `verificar.py` falla si
  aparece una.
- **Un texto no se guarda como fórmula.** Un `="texto…"` > 255 caracteres lo parte Excel
  en `_xlfn._LONGTEXT(...)`. El builder ya lo normaliza; no introducir literales `="..."`
  largos nuevos.
- **Validación de datos: solo rango literal o lista de ítems**, nunca una fórmula como
  origen. Se usa `dv_list(ws, cell, '"item1,item2,…"', com)`.
- **Regla nº 1 — ningún valor normativo sale de memoria.** Las fórmulas y valores de las
  secciones 1-6 del 212 **se transcriben del *oracle* capturado del `.xlsm` real** (Tarea
  1), nunca se reescriben de memoria ni se «mejoran». El *oracle* es la fuente; el test de
  caracterización es la prueba de que la transcripción es fiel.
- **snake_case** en todo archivo nuevo.
- **Entrega parchando sobre Rev4** por el builder (D-3 del spec), no una Rev5.
- **No hay Excel en este entorno.** Ninguna fórmula se puede recalcular aquí. La
  equivalencia se garantiza por igualdad de cadena contra el *oracle*; el recálculo real
  (`verificar.py §6+`) y el F9 del caso semilla los corre el ingeniero en Windows.
- **Español, SI por defecto** (MPa, mm, °C). Tono directo, sin relleno.

---

## Estructura de archivos

- `scripts/build_db_materiales.py` — **Modify.**
  - Nueva función `build_parche_art212(wb, b313, iid1a, iidb, fac_info, rangos)` (junto a
    `build_collar_art206`, ~línea 6510).
  - `integrate_motor()` (6156) y `corregir_art212_fase1()` (6224) se **absorben** en la
    nueva función y se **eliminan**. `comentar_art212_base()` (6287) se **conserva** y se
    llama desde la nueva función (sus comentarios ya no dependen de la herencia).
  - `HOJAS_HEREDADAS` (7949): se quita `"Parche_PCC2_Art212"` → queda
    `("Instrucciones", "Datos_Ref")`.
  - `TEXTOS_HEREDADOS` (8019): se elimina la entrada `("Parche_PCC2_Art212", "A16")`
    (el texto nuevo se escribe directo en el builder).
  - `main()` (8340): `integrate_motor(...)` → `build_parche_art212(...)`.
- `scripts/parche_art212_ref.json` — **Create.** *Oracle* versionado: por cada celda no
  vacía de `Parche_PCC2_Art212` (filas 1-129, columnas A-G), su `value` (fórmula o
  literal); más la lista de rangos de validación de datos con su `formula1`, los rangos
  fusionados y los anchos de columna. Es la única fuente de las fórmulas de las secciones
  1-6; se genera una vez desde el Rev4 committeado y no se edita a mano.
- `scripts/_dump_parche_ref.py` — **Create.** Script de un solo uso que lee el `.xlsm`
  real y escribe `parche_art212_ref.json`. Se versiona para poder regenerar el *oracle*
  si algún día se repone desde un Rev validado.
- `scripts/test_dashboard.py` — **Modify.** `TestParidadSeccion7Art212` (520) se sustituye
  por `TestParidadHojaParche`, que carga el *oracle* y compara la hoja entera. Se añade
  `TestBuildParcheContraOracle` (unidad: construye en un wb de usar y tirar y compara
  contra el *oracle*, sin depender de regenerar el entregable).

Nota sobre el sistema visual: la hoja reconstruida usa **tokens del sistema** (`IN_F`,
`IN_FILL`, `CAJA_CAMPO`, `BAND_FILL`, `HDR_F`, `HDR_FILL`, `BOX_FRANJA`, `MONO`, `MACRO`,
etc.), exactamente como `build_collar_art206`. Por eso el *oracle* **solo fija
fórmulas/valores/validaciones** (el comportamiento), no fuentes ni rellenos heredados del
Rev0: la conformidad visual la sigue garantizando `TestSistemaVisual`, que recorre los 2,4
M de celdas con formato y falla ante cualquier color o fuente fuera de la paleta. Separar
así las dos capas evita acoplar la hoja nueva al resultado del retonado del Rev0.

---

## Tarea 1: *Oracle* del 212 y test de caracterización de hoja entera

**Files:**
- Create: `scripts/_dump_parche_ref.py`
- Create: `scripts/parche_art212_ref.json`
- Modify: `scripts/test_dashboard.py` (sustituir `TestParidadSeccion7Art212`, ~520-564)

**Interfaces:**
- Produce: `parche_art212_ref.json` con esquema
  `{"formulas": {"<celda>": <valor|null>, …}, "validaciones": [{"sqref": "...",
  "formula1": "..."}, …], "fusionados": ["A16:G16", …], "anchos": {"A": 44.0, …}}`.
  `<valor>` es la cadena de fórmula (con `=`) o el literal (texto/número) tal como
  `openpyxl` lo lee de la celda; `null` si la celda existe en el mapa pero está vacía.
- Produce: `TestParidadHojaParche` y el helper de carga `cargar_oracle_parche()` que las
  tareas siguientes reutilizan.

- [ ] **Step 1: Escribir el script de volcado del *oracle***

En `scripts/_dump_parche_ref.py`:

```python
"""Un solo uso: vuelca Parche_PCC2_Art212 del .xlsm real a parche_art212_ref.json.
El JSON es el *oracle* del desanclado total — la fuente de las formulas de las
secciones 1-6, que la Regla n.1 prohibe reescribir de memoria. Regenerarlo solo
si se repone la hoja desde un Rev validado en Excel."""
import json, os
from pathlib import Path
import openpyxl

RUTA = Path(os.environ.get(
    "MOTOR_XLSM",
    Path(__file__).resolve().parent.parent / "Motor_de_Calculo_ASME_PCC_Rev4.xlsm"))
SALIDA = Path(__file__).resolve().parent / "parche_art212_ref.json"
HOJA = "Parche_PCC2_Art212"


def main():
    wb = openpyxl.load_workbook(RUTA, keep_vba=True)
    ws = wb[HOJA]
    formulas = {}
    for fila in ws.iter_rows(min_row=1, max_row=129, min_col=1, max_col=7):
        for c in fila:
            if c.value is not None:
                formulas[c.coordinate] = c.value
    validaciones = sorted(
        ({"sqref": str(dv.sqref), "formula1": dv.formula1} for dv in ws.data_validations.dataValidation),
        key=lambda d: d["sqref"])
    fusionados = sorted(str(r) for r in ws.merged_cells.ranges)
    anchos = {col: dim.width for col, dim in ws.column_dimensions.items() if dim.width}
    SALIDA.write_text(json.dumps(
        {"formulas": formulas, "validaciones": validaciones,
         "fusionados": fusionados, "anchos": anchos},
        ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    print(f"{len(formulas)} celdas, {len(validaciones)} validaciones -> {SALIDA}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Generar el *oracle* desde el Rev4 committeado**

Run (PowerShell, desde `scripts/`):
```powershell
python _dump_parche_ref.py
```
Espera: imprime algo como `~190 celdas, ~15 validaciones -> parche_art212_ref.json`.
Comprobar que el JSON contiene las claves `D115`, `D126`, `F90`, `D21`, `D114`, `E114`,
`A1`, `A2`, `A16` (las anclas conocidas del test existente y de la Fase 1).

- [ ] **Step 3: Reemplazar el test de paridad por el de hoja entera**

En `scripts/test_dashboard.py`, sustituir la clase `TestParidadSeccion7Art212`
(líneas ~512-564) por:

```python
# ---------------------------------------------------------------------------
# El desanclado total no puede cambiar el Art. 212
# ---------------------------------------------------------------------------
# parche_art212_ref.json es el *oracle*: la hoja 212 tal como la entregaba el
# maestro Rev0, congelada desde el ultimo build valido. Reconstruirla en codigo
# (build_parche_art212) debe reproducir cada formula, literal y validacion al
# caracter. Sin Excel aqui, esta es la unica red: si una transcripcion difiere,
# falla ahora y no en la mesa del ingeniero.
def cargar_oracle_parche():
    p = Path(__file__).resolve().parent / "parche_art212_ref.json"
    return json.loads(p.read_text(encoding="utf-8"))


class TestParidadHojaParche:
    HOJA = "Parche_PCC2_Art212"

    def test_todas_las_formulas_y_literales(self, wb):
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        for celda, esperado in oracle["formulas"].items():
            assert ws[celda].value == esperado, celda

    def test_validaciones_de_datos(self, wb):
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        reales = sorted(
            ({"sqref": str(dv.sqref), "formula1": dv.formula1}
             for dv in ws.data_validations.dataValidation),
            key=lambda d: d["sqref"])
        assert reales == oracle["validaciones"]

    def test_rangos_fusionados(self, wb):
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        assert sorted(str(r) for r in ws.merged_cells.ranges) == oracle["fusionados"]
```

Asegurarse de que `import json` y `from pathlib import Path` ya están en la cabecera del
archivo (lo están; `RUTA` ya los usa).

- [ ] **Step 4: Correr el test contra el Rev4 actual (debe pasar: aún es la hoja heredada)**

Run:
```powershell
python -m pytest test_dashboard.py::TestParidadHojaParche -q
```
Espera: PASS. En este punto el `.xlsm` todavía trae la hoja heredada, así que es idéntica
al *oracle* (se está comparando el archivo con un volcado de sí mismo). El test cobra
valor como red de regresión cuando el archivo se regenere desde código (Tarea 9).

- [ ] **Step 5: Commit**

```powershell
git add scripts/_dump_parche_ref.py scripts/parche_art212_ref.json scripts/test_dashboard.py
git commit -m "Oracle del Art. 212 y test de paridad de hoja entera (Fase 4)"
```

---

## Tarea 2: Esqueleto de `build_parche_art212()` — identificación, Sección 7 y cableado

**Files:**
- Modify: `scripts/build_db_materiales.py` (nueva función tras `corregir_art212_fase1`,
  ~línea 6286; aún no se toca `main`)
- Modify: `scripts/test_dashboard.py` (añadir `TestBuildParcheContraOracle`)

**Interfaces:**
- Consume: `construir_seccion7_material(ws, fila_base=105, …, columnas=(("D","Metal base"),
  ("E","Collar / parche")), modo_cell="$D$11", temp_fuente_cell="$D$25",
  incluir_ej_ec=True, destino_st=…, nota_extra_cascada=…)`; helpers `new_sheet`,
  `autosize`, `dv_list`, `franja`, `_nota`; tokens `IN_F`, `IN_FILL`, `CAJA_CAMPO`,
  `BAND_FILL`, `HDR_F`, `HDR_FILL`, `BOX_FRANJA`, `MONO`, `MACRO`, `PAPEL`, `TINTA`, `GRIS`;
  constantes `MOTOR`, `SEED_ART212_BASE`, `SEED_ART212_COLLAR`.
- Produce: `build_parche_art212(wb, b313, iid1a, iidb, fac_info, rangos)` que crea la hoja
  `MOTOR` con `new_sheet`, la identificación (filas 5-6), la Sección 7 (desde la fila 105)
  y el cableado D39/D40/D44/F90; protege la hoja con `0000`. Las secciones 1-6 las añaden
  las Tareas 3-8. Reproduce el `del wb[MOTOR]` defensivo antes de crear (el maestro aún
  trae la hoja; sin esto `create_sheet` duplicaría el nombre).

- [ ] **Step 1: Escribir el test de unidad contra el *oracle* (falla primero)**

En `scripts/test_dashboard.py`, al final del archivo:

```python
class TestBuildParcheContraOracle:
    """Construye Parche_PCC2_Art212 en un wb de usar y tirar y lo compara contra
    el *oracle*, sin regenerar el entregable. Conforme se anaden secciones
    (Tareas 3-8), mas celdas del oracle quedan cubiertas; COBERTURA_PARCIAL
    enumera las que cada tarea ya debe reproducir."""

    def _construir(self):
        import build_db_materiales as B
        wb = openpyxl.Workbook()
        # Bases minimas que la Seccion 7 referencia por nombre (Excel resuelve al
        # abrir; aqui solo deben existir para que el seed de la Fase 1 valide).
        for n in ("DB_B31_3", "Datos_Ref", "MAP_Factores", "DB_BPVC_IID",
                  "DB_BPVC_IID_B"):
            wb.create_sheet(n)
        # Sembrar en DB_B31_3 los dos material_id del caso precargado para que
        # el guardia de corregir_art212_fase1 (ahora en build_parche_art212) pase.
        db = wb["DB_B31_3"]
        db.cell(B.R_DATA, 1, B.SEED_ART212_BASE)
        db.cell(B.R_DATA + 1, 1, B.SEED_ART212_COLLAR)
        B.build_parche_art212(wb, b313_stub(), iid1a_stub(), iidb_stub(),
                              fac_stub(), rangos_stub())
        return wb["Parche_PCC2_Art212"]
```

> Los `*_stub()` devuelven los dicts mínimos que `construir_seccion7_material` consume
> (`{"sheet": "...", "last_row": N}` para las bases, `packed_refs`-compatibles, y
> `rangos={"B313":…,"IID1A":…}`). Copiar su forma exacta de cómo `main()` arma `b313`,
> `iid`, `iidb`, `fac`, `rangos` (líneas ~8150-8340 de `build_db_materiales.py`) — **no
> inventarla**: leer esos dicts reales y reproducir solo los campos que la Sección 7 usa
> (`sheet`, `last_row`, y lo que `packed_refs` lee). Si montar los stubs resulta frágil,
> alternativa válida: el test construye el libro completo con el `main` real sobre un
> `--out` temporal y abre esa hoja (más lento, cero stubs).

Añadir la aserción de cobertura incremental:

```python
    ANCLAS_T2 = ("A1", "A2", "D5", "D6", "D115", "E115", "D126", "E126",
                 "D39", "D40", "D44", "F90", "D114", "E114", "D128")

    def test_anclas_de_la_tarea_2(self):
        oracle = cargar_oracle_parche()
        ws = self._construir()
        for celda in self.ANCLAS_T2:
            assert ws[celda].value == oracle["formulas"][celda], celda
```

Run:
```powershell
python -m pytest test_dashboard.py::TestBuildParcheContraOracle::test_anclas_de_la_tarea_2 -q
```
Espera: FAIL con `AttributeError: module 'build_db_materiales' has no attribute
'build_parche_art212'` (o `KeyError: 'Parche_PCC2_Art212'`).

- [ ] **Step 2: Escribir el esqueleto de `build_parche_art212()`**

En `scripts/build_db_materiales.py`, tras `corregir_art212_fase1` (~6286), añadir. Copiar
los helpers internos `lab`/`inp`/`calc`/`band`/`header` **tal cual** de
`build_collar_art206` (6536-6572) — mismo estilo, mismos tokens:

```python
def build_parche_art212(wb, b313, iid1a, iidb, fac_info, rangos):
    """Motor Art. 212 (parche soldado), 100% en codigo — desanclado del maestro
    Rev0 (Fase 4). Antes vivia heredado + corregido por integrate_motor/
    corregir_art212_fase1; ahora nace con new_sheet como Collar_PCC2_Art206.
    Reutiliza construir_seccion7_material (Seccion 7, compartida con el 206) y el
    lookup de Datos_Ref robustecido con ""& contra la cedula. Las formulas de las
    secciones 1-6 se transcriben del oracle parche_art212_ref.json (Regla n.1:
    no se reescriben de memoria); TestParidadHojaParche las fija al caracter."""
    if MOTOR in wb.sheetnames:        # el maestro aun trae la hoja heredada
        del wb[MOTOR]
    ws = new_sheet(
        wb, MOTOR,
        "MOTOR DE CALCULO — PARCHE SOLDADO (ASME PCC-2 Art. 212)",
        'ASME PCC-2 Art. 212 (Fillet Welded Patches). El collar de encierro '
        'total (Art. 206) es un motor aparte. Caso precargado: Linea '
        '12"-CWS-46-032-B1 (U46).')
    autosize(ws, {"A": 44, "B": 8, "C": 14, "D": 16, "E": 16, "F": 16, "G": 46})

    # [helpers lab/inp/calc/band/header — copiar de build_collar_art206 6536-6572]

    # --- Seccion 7: cascada de material base + collar/parche ----------------
    construir_seccion7_material(
        ws, 105, b313, iid1a, iidb, fac_info, rangos,
        columnas=(("D", "Metal base"), ("E", "Collar / parche")),
        modo_cell="$D$11", temp_fuente_cell="$D$25", incluir_ej_ec=True,
        destino_st="D39/D40 de la seccion 3",
        nota_extra_cascada="La lista corta de Datos_Ref queda como respaldo "
        "historico y ya no alimenta el calculo.")

    # --- Identificacion (filas 5-6) -----------------------------------------
    # [transcribir A5/A6 y D5/D6 del oracle: rotulos + inp()]

    # --- Cableado Seccion 7 -> Secciones 1/3 (D39/D40/D44/F90) --------------
    # Transcrito literal del oracle (coincide con integrate_motor 6171-6205):
    ws["D39"] = '=IF($E$125="OK",$E$126,NA())'
    ws["D40"] = '=IF($D$125="OK",$D$126,NA())'
    ws["D44"] = "=$E$128"
    ws["D44"].font = Font(name=MONO, size=10, color=TINTA)
    ws["D44"].fill = PAPEL_FILL
    ws["D44"].protection = Protection(locked=True)
    ws["F90"] = ('=IF(OR($D$125="SIN MATERIAL SELECCIONADO",'
                 '$E$125="SIN MATERIAL SELECCIONADO"),"ELIJA MATERIAL (Seccion 7)",'
                 'IF(OR($D$125<>"OK",$E$125<>"OK"),"REVISAR — MATERIAL FUERA DE RANGO",'
                 'IF(AND(F84="CUMPLE",F85="CUMPLE",F86="CUMPLE",F87="CUMPLE"),'
                 '"APTO","REVISAR")))')
    ws["G39"] = "S(T) del collar — base de datos ASME (seccion 7)"
    ws["G40"] = "S(T) del metal base — base de datos ASME (seccion 7)"
    ws["G44"] = "Ej por lookup (B31.3 Tabla A-3) — seccion 7"
    # [notas _nota() de D39/D40/D44/F90 — copiar de integrate_motor 6173-6205]

    # --- Caso precargado (seed por la celda Variante, con guardia Regla n.1)--
    db = wb["DB_B31_3"]
    ids = {db.cell(r, 1).value for r in range(R_DATA, db.max_row + 1)}
    for etiqueta, mid in (("base", SEED_ART212_BASE), ("collar", SEED_ART212_COLLAR)):
        if mid not in ids:
            raise SystemExit(
                f"Art.212: el material_id sembrado ({etiqueta}) no existe en "
                f"DB_B31_3: {mid!r} (Regla n.1).")
    ws["D114"] = SEED_ART212_BASE
    ws["E114"] = SEED_ART212_COLLAR
    # [notas de D114/E114 — copiar de corregir_art212_fase1 6268-6274]

    # --- Secciones 1-6: las anaden las Tareas 3-8 ---------------------------

    ws.protection.password = "0000"
    ws.protection.sheet = True
```

> Nota: `D128` ('A106 | Seamless pipe') y las fórmulas D115/E115/D126/… las escribe
> `construir_seccion7_material`, no esta función — por eso la ancla `D128` de la Tarea 2
> ya queda cubierta al llamarla. Verificar contra el *oracle* que `construir_seccion7_material`
> con `fila_base=105` y estas columnas reproduce D107-D129 idénticos (el test de hoja
> entera lo confirma en la Tarea 9; el `TestParidadSeccion7Art212` viejo ya lo probaba).

- [ ] **Step 3: Correr el test de anclas de la Tarea 2 (debe pasar)**

Run:
```powershell
python -m pytest test_dashboard.py::TestBuildParcheContraOracle::test_anclas_de_la_tarea_2 -q
```
Espera: PASS. Si falla en `D5`/`D6`, revisar la transcripción de identificación contra
`oracle["formulas"]["D5"]` / `["D6"]`.

- [ ] **Step 4: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/test_dashboard.py
git commit -m "build_parche_art212: esqueleto, Seccion 7 y cableado D39/D40/D44/F90"
```

---

## Tareas 3-8: Transcribir las secciones 1-6 desde el *oracle*

**Patrón común a las seis** (una tarea por sección; cada una es un bloque contiguo de
filas que un revisor puede aceptar o rechazar por separado). Para cada sección:

1. Abrir `parche_art212_ref.json` y localizar las celdas de su rango de filas.
2. En `build_parche_art212`, antes de `ws.protection…`, escribir cada celda **copiando el
   valor del *oracle* al carácter** con el helper que corresponda: `band()` para las
   bandas de sección, `header()` para las filas de encabezado, `lab()` para los rótulos de
   columna A + unidad en C + referencia en G, `inp()` para las celdas de entrada (con
   `dv_list` donde el *oracle* liste una validación sobre esa celda) y `calc()` para las
   fórmulas.
3. Añadir las anclas de la sección a un nuevo método `test_seccion_N` dentro de
   `TestBuildParcheContraOracle`, o ampliar `ANCLAS_*`, con **todas** las celdas del rango
   que el *oracle* declara (no una muestra: la sección entera).
4. `python -m pytest test_dashboard.py::TestBuildParcheContraOracle -q` → PASS.
5. Llamar a `comentar_art212_base(ws)` queda para la Tarea 8 (necesita las celdas ya
   escritas); aquí solo se transcriben las notas `_nota()` propias si el *oracle* no las
   cubre (el *oracle* no guarda comentarios — los comentarios los pone
   `comentar_art212_base`, que se conserva).
6. Commit con mensaje `Art. 212 desanclado: seccion <N> en codigo`.

> **Regla dura de estas seis tareas:** no se teclea ninguna fórmula «de cabeza». Si una
> celda del rango no está en el *oracle*, **estaba vacía** y se deja vacía. Si el *oracle*
> trae un literal numérico (p. ej. el factor de conversión `0.0980665` en D47, o `1.5`),
> se escribe ese literal, no una expresión equivalente. Para evitar error de tecleo en las
> fórmulas largas, se permite generar los literales Python desde el *oracle* con un helper
> de un solo uso (`for k,v in oracle["formulas"].items(): print(...)`) y pegar el
> resultado revisado; el test de hoja entera es, de todos modos, el árbitro.

- [ ] **Tarea 3 — Aplicación y código de construcción (filas 10-14).** D10 (selector
  `dv_list` Tubería/Virola/Cabezal), D11 (MODO), D12 (código), D13 (tabla), D14 (kf).
  Anclas: `D10,D11,D12,D13,D14` + rótulos A10-A14.
- [ ] **Tarea 4 — Sección 1, datos de entrada (filas 16-35).** Banda A16 (texto nuevo:
  `"1.  DATOS DE ENTRADA  (campo con linea inferior = editable)"`, que era el destino de
  `TEXTOS_HEREDADOS`), encabezado fila 17, NPS/cédula (D18/D19 con `dv_list`), OD (D20),
  **t con el `""&` robusto** (D21 = `=INDEX(Datos_Ref!$C$5:$P$37,MATCH($D$18,Datos_Ref!$A$5:$A$37,0),MATCH(""&$D$19,Datos_Ref!$C$4:$P$4,0))`),
  materiales descriptivos D22/D23 (con G22/G23 = `"Descriptivo · el S(T) rige en la
  Seccion 7"`), fluido D24, T D25, presiones D26-D28, espesor/altura/luz/filete/solape/
  defecto/distancia D29-D35. Anclas: todo D18-D35 presente en el *oracle* + A16 + G22/G23.
- [ ] **Tarea 5 — Sección 2, esfuerzos admisibles y factores (filas 39-48).** D39/D40/D44
  **ya** los puso la Tarea 2 (cableado); aquí van D41 (gobernante), D42 (Ej adoptado), D43
  (Y), D45 (densidad), D46 (factor prueba), D47 (conversión kg/cm²→MPa), D48 (1,5·Sa).
  Anclas: D41,D42,D43,D45,D46,D47,D48 + rótulos.
- [ ] **Tarea 6 — Geometría (filas 52-57).** D52 (Dm), D53 (Rm), D54 (Ri), D55 (e), D56
  (Rf), D57 (C_sw). Anclas: D52-D57.
- [ ] **Tarea 7 — Sección 3 (casos de presión, filas 61-69) + resultados (73-80).** Filas
  triples D/E/F 61-69 (presión, F_m, filete, t_req, S_w membrana/flexión/total, CUMPLE) y
  73-80 (L_min, P_max MPa/kg, margen, conformado, desarrollo, corte, peso). Anclas: todas
  las celdas D/E/F presentes en 61-69 y D73-D80.
- [ ] **Tarea 8 — Sección 5 verificaciones (84-90) + Sección 6 especificaciones (93-99) +
  aviso (101) + comentarios.** Filas 84-89 (triple, con F90 ya puesto en Tarea 2), textos
  B93-B99, aviso A101. **Al final de la sección**, una sola llamada
  `comentar_art212_base(ws)` (se conserva intacta; ahora todas sus celdas existen, así que
  su guardia `if c.value is not None` se cumple). Anclas: D84-F89, B93-B99, A101.

---

## Tarea 9: Desanclar en `main()` y limpiar la herencia

**Files:**
- Modify: `scripts/build_db_materiales.py` (`HOJAS_HEREDADAS` 7949, `TEXTOS_HEREDADOS`
  8019-8023, `main` 8340; eliminar `integrate_motor` 6156-6212 y `corregir_art212_fase1`
  6224-6284)

**Interfaces:**
- Consume: `build_parche_art212` (Tareas 2-8).
- Produce: un `.xlsm` en el que `Parche_PCC2_Art212` nace 100 % en código.

- [ ] **Step 1: Quitar la hoja de la herencia**

`HOJAS_HEREDADAS = ("Instrucciones", "Datos_Ref")` (quitar `"Parche_PCC2_Art212"`).
Eliminar de `TEXTOS_HEREDADOS` la entrada `("Parche_PCC2_Art212", "A16"): (…)` — el texto
nuevo ya lo escribe la Tarea 4. Si `TEXTOS_HEREDADOS` queda vacío, conservar el dict vacío
(`retonar_heredadas` lo itera sin problema).

- [ ] **Step 2: Cambiar la llamada en `main()`**

Línea 8340: `integrate_motor(wb, b313, iid, iidb, fac, rangos)` →
`build_parche_art212(wb, b313, iid, iidb, fac, rangos)`. Dejar `build_collar_art206(...)`
justo después, sin cambios.

- [ ] **Step 3: Eliminar el código muerto**

Borrar `integrate_motor` (6156-6212) y `corregir_art212_fase1` (6224-6284): su contenido
vive ahora en `build_parche_art212`. **Conservar** `comentar_art212_base` (la llama la
Tarea 8), `SEED_ART212_BASE`/`SEED_ART212_COLLAR` y el comentario de cabecera del seed.
Comprobar que ningún otro sitio llama a las funciones borradas:
```powershell
python -c "import re,io; s=open('build_db_materiales.py',encoding='utf-8').read(); print('integrate_motor:',s.count('integrate_motor')); print('corregir_art212_fase1:',s.count('corregir_art212_fase1'))"
```
Espera: `integrate_motor: 0` y `corregir_art212_fase1: 0` (ya borradas las definiciones y
la llamada). Buscar también en los tests:
```powershell
```
(usar Grep sobre `scripts/` por `integrate_motor|corregir_art212_fase1`; no debe quedar
ninguna referencia.)

- [ ] **Step 4: Build completo**

Run (desde `scripts/`; si es la primera corrida de la sesión o se tocó `vba/*.vba`, correr
antes `python make_vba_seed.py` — aquí **no** se tocó VBA, así que normalmente no hace
falta):
```powershell
python build_db_materiales.py --resources ..\..\..\resources --in ..\..\..\templates\maestro_con_macros.xlsm --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```
Espera: termina sin abortos. Si aborta en el guardia del seed, es que `DB_B31_3` se
construye después de `build_parche_art212`; comprobar que la llamada queda en el mismo
punto del pipeline que tenía `integrate_motor` (después de construir las bases), no antes.

- [ ] **Step 5: Los tres gates que corren sin Excel**

Run:
```powershell
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
```
Espera: verde salvo los fallos conocidos de `TestLintVba` por `winreg` en no-Windows (en
Windows deben pasar también). `TestParidadHojaParche` **ahora sí** es una red real: el
archivo se regeneró desde código y debe seguir igual al *oracle*. `TestSistemaVisual` debe
pasar (la hoja nueva usa solo tokens del sistema). `TestSincroniaPythonVba` no se toca
(no cambió el árbol ni `NAVEGABLES`).

```powershell
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```
Espera: §1-5 en verde (0 discrepancias en los 271 536 valores, 0 fórmulas de matriz
dinámica, 0 validaciones no portables, cascada contigua). Requiere Excel instalado para
§6+; si el ingeniero lo corre en Windows, debe volver 0.

- [ ] **Step 6: Commit**

```powershell
git add scripts/build_db_materiales.py
git commit -m "Desancla Parche_PCC2_Art212 del Rev0: nace 100% en codigo (Fase 4)"
```

---

## Tarea 10: Regenerar el entregable, documentar y entregar al ingeniero

**Files:**
- Modify: `C:\Users\User\Documents\development\asme-pcc\CLAUDE.md` (sección «Motor de
  cálculo — estado actual»: `HOJAS_HEREDADAS` pasa a dos hojas; el 212 ya no es heredado)
- Modify: `outputs/plans/plan_motor_art206_collar_y_separacion_parche_collar.md` (Fase 4 →
  EJECUTADA)
- Modify: `outputs/plans/plan_desanclado_total_motor_art212.md` (este plan → estado final)

- [ ] **Step 1: Confirmar el entregable regenerado**

El `.xlsm` ya se regeneró en la Tarea 9 Step 4. Verificar que `Parche_PCC2_Art212` sigue
en el `order` de `main()` (8402, por nombre — no cambia) y que abre mostrando solo el
`Dashboard` (la visibilidad la graba `aplicar_visibilidad`; el 212 sigue navegable).

- [ ] **Step 2: Actualizar la documentación**

En `CLAUDE.md`, donde describe las hojas heredadas y el motor 212, anotar que la hoja nace
en código (ya no hereda del Rev0) y que `build_parche_art212` es su constructor, hermano
de `build_collar_art206`. En el plan del 206, marcar la Fase 4 como **EJECUTADA** con la
fecha (2026-09-10). En este plan, añadir una sección «Estado final» con el resultado de
los gates.

- [ ] **Step 3: Entregar para validación en Excel**

Enviar el `.xlsm` regenerado al ingeniero con `SendUserFile` y pedir la validación que
este entorno no puede dar: F9 sobre el caso precargado (Línea 12"-CWS-46-032-B1), que
`t`=6,35 mm (12" · SCH 20), la geometría completa, `Sa_b≈138 / Sa_c≈161 MPa`, y que
`verificar.py §6+` vuelve 0 en Windows. **No** cerrar la Fase 4 hasta que el ingeniero
confirme en Excel real, igual que se hizo con la Fase 1.

- [ ] **Step 4: Commit**

```powershell
git add CLAUDE.md outputs/plans/plan_motor_art206_collar_y_separacion_parche_collar.md outputs/plans/plan_desanclado_total_motor_art212.md
git commit -m "Documenta el desanclado total del Art. 212 (Fase 4 ejecutada)"
```

---

## Verificación — qué se puede aquí y qué no

- **Aquí (sin Excel):** `pytest` (las tres suites), `verificar.py §1-5`, build sin abortos,
  y sobre todo `TestParidadHojaParche` + `TestBuildParcheContraOracle` — la igualdad
  carácter a carácter contra el *oracle* es lo que sustituye a la validación en Excel para
  las fórmulas.
- **Solo en Excel de Windows (ingeniero):** `verificar.py §6+` (recálculo real con
  `win32com`), F9 del caso precargado, y la revisión visual de la hoja exportada a PDF/PNG
  (openpyxl miente sobre bordes de rango fusionado).
- **Regla dura:** ninguna fórmula del 212 sale de memoria. Todas vienen del *oracle*
  capturado del `.xlsm` real; si el *oracle* no tiene una celda, estaba vacía.

## Riesgos

- **Error de transcripción en una fórmula de las secciones 1-6.** Mitigación: el *oracle*
  + `TestParidadHojaParche` lo detectan aquí, no en Excel; se permite generar los literales
  desde el *oracle* en vez de teclearlos.
- **Nombre de hoja duplicado.** El maestro aún trae la hoja heredada; `build_parche_art212`
  hace `del wb[MOTOR]` antes de `new_sheet`. Sin eso, `create_sheet` duplicaría el nombre.
- **Orden en el pipeline.** `build_parche_art212` debe llamarse después de que existan
  `DB_B31_3`/`Datos_Ref`/`MAP_Factores` (igual que `integrate_motor` hoy). El guardia del
  seed aborta si se adelanta.
- **El retonado ya no toca el 212.** Al salir de `HOJAS_HEREDADAS`, `retonar_heredadas` no
  lo procesa; la conformidad visual pasa a depender solo de que el builder use tokens del
  sistema. `TestSistemaVisual` lo verifica.
- **Stubs del test de unidad frágiles.** Si armar `b313_stub()`/etc. resulta quebradizo,
  usar la alternativa declarada: construir el libro completo con `main` sobre un `--out`
  temporal y abrir la hoja de ahí.

## Self-review (hecho)

- **Cobertura del spec (Fase 4 opción a):** toda la construcción de la hoja 212 se mueve
  al builder (Tareas 2-8), el maestro deja de aportarla (Tarea 9), se entrega sobre Rev4
  (Tarea 10). Cubierto.
- **Sin placeholders de fórmula inventada:** deliberado y alineado con la Regla nº 1 — las
  fórmulas de las secciones 1-6 se transcriben del *oracle*, no se escriben en el plan de
  memoria. Las fórmulas que **sí** aparecen literales en el plan (D39/D40/D44/F90, D21)
  están verificadas: son las ya fijadas por `TestParidadSeccion7Art212` y por
  `corregir_art212_fase1` en el código actual.
- **Consistencia de tipos/nombres:** `build_parche_art212(wb, b313, iid1a, iidb, fac_info,
  rangos)` usa la misma firma que `build_collar_art206`; la llamada en `main` pasa
  `(wb, b313, iid, iidb, fac, rangos)`, los mismos argumentos que recibía `integrate_motor`.
