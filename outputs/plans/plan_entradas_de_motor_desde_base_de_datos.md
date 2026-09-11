# Entradas de motor desde base de datos — Plan de implementación

> **Para ejecutores:** SUB-SKILL REQUERIDA: usar `superpowers:subagent-driven-development`
> (recomendado) o `superpowers:executing-plans` para ejecutar tarea por tarea. Los pasos
> usan casillas (`- [ ]`) para el seguimiento.

**Objetivo:** Que ninguna entrada de un motor de cálculo que exista en una base de datos
se teclee a mano ni salga de una lista fija: se elige de un desplegable alimentado por la
base, que además rechaza lo que no esté en ella.

**Arquitectura:** Tres fases. La 1 cierra el hueco actual (elimina las dos filas de
material del Art. 212, hace bloqueantes los desplegables de los 12 motores y planta el
guardia que impide la reincidencia). La 2 extrae B36.10M y B36.19M a `resources/` y
construye `DB_B36_10` y `DB_B36_19`. La 3 repunta NPS y cédula de los dos motores contra
esas bases y retira `Datos_Ref`, con lo que `HOJAS_HEREDADAS` queda en `("Instrucciones",)`.

**Stack:** Python 3 + openpyxl; pytest; el builder `build_db_materiales.py`; Excel de
Windows para `verificar.py` §6+ (disponible en esta máquina, comprobado el 2026-09-10).

**Spec:** `outputs/plans/spec_entradas_de_motor_desde_base_de_datos.md` — léelo antes de
la Tarea 1; este plan argumenta desde él y no lo repite.

## Global Constraints

Copiadas de `asme-pcc/CLAUDE.md` → «Reglas de diseño del libro». Aplican a todas las
tareas, aunque no se repitan en cada una:

- **Regla nº 1 — ningún valor normativo sale de la memoria del modelo.** Todo dato que
  entre a una base sale del JSON de `resources/`. El PDF **solo corrige y audita** el
  JSON; nunca alimenta un motor ni una hoja directamente.
- **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`,
  `VSTACK`, `_xlfn`. Solo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`.
- **Regla 2 — validación de datos: solo rango literal o lista de ítems**, nunca una
  fórmula como origen. Las listas dependientes se materializan en columnas ocultas.
- **Regla 5 — cascada contigua.** Cada base se ordena de modo que cada nivel sea un
  bloque contiguo; el mecanismo `COUNTIF`/`INDEX`/`MATCH` de las listas dependientes lo
  exige.
- **Regla 9 — los valores se cargan tal como están impresos.** El marcador `'...'` de
  B36.10M/B36.19M es «no aplica» impreso por el código: se conserva, no se convierte a
  vacío.
- **Regla 10 — conmutador de unidades.** B36.10M y B36.19M publican las dos unidades en
  la **misma celda**, así que el conmutador cambia de **columna**, nunca convierte.
- **Reglas 12, 13 y 14** — las tres que motivan este plan; están en `CLAUDE.md`.
- **snake_case** en todo archivo nuevo. **Español, SI por defecto.**
- **Entrega parchando sobre Rev4** por el builder, no una Rev5.

## Decisiones ya tomadas (no reabrir sin motivo nuevo)

1. **La norma dimensional se elige con un selector explícito**, no se deriva de la
   familia del material: una tubería inoxidable puede fabricarse a cédulas B36.10M, así
   que derivarla daría el espesor equivocado en ese caso.
2. **La cédula se ofrece con un designador combinado** — `40 (STD)`, `80 (XS)`, `XXS`,
   `10` — construido de las columnas `Schedule No.` e `Identification`. Una sola cosa que
   elegir, las dos designaciones a la vista, sin filas duplicadas.
3. **El NPS se muestra como lo imprime el código** (`12 (300)`), con el decimal parseado
   en columna oculta para los lookups y para que el caso semilla (NPS 12) siga
   resolviendo.
4. **`Datos_Ref` se retira entera**, incluido su bloque obsoleto de las filas 40-47.

---

## Estructura de archivos

- `resources/ASME B36/b36_10m_2022/table_dimensiones.json` — **Create.** La extracción
  entregada, depositada con su `meta.json` de procedencia.
- `resources/ASME B36/b36_19m_2022/table_dimensiones.json` — **Create.** Ídem.
- `scripts/b36_dimensiones.py` — **Create.** Librería + CLI que convierte los bloques
  `<table>` de esas extracciones en filas normalizadas. Mismo patrón que
  `secii_tablas.py`: el builder la importa como librería, y `--informe` mide y reporta
  sin escribir nada.
- `scripts/test_b36_dimensiones.py` — **Create.** Sus pruebas.
- `scripts/build_db_materiales.py` — **Modify.** `dv_list` bloqueante;
  `LITERALES_PERMITIDOS`; `DIVERGENCIAS_DECLARADAS`; `build_db_b36()`; la cascada
  dimensional de los dos motores; borrado de D22/D23; `HOJAS_HEREDADAS`.
- `scripts/verificar.py` — **Modify.** §5 pasa a exigir bloqueo y origen no literal en
  hojas de motor; §11 nueva, auditoría fila a fila de las dos bases B36.
- `scripts/test_dashboard.py` — **Modify.** Consume `DIVERGENCIAS_DECLARADAS`; prueba el
  guardia y la cascada dimensional.

---

## Tarea 1: `DIVERGENCIAS_DECLARADAS` y borrado de D22/D23

**Files:**
- Modify: `scripts/build_db_materiales.py` (`build_parche_art212`, bloque de D22/D23 y su `DataValidation`)
- Modify: `scripts/test_dashboard.py` (`TestParidadHojaParche`)

**Interfaces:**
- Produce: `DIVERGENCIAS_DECLARADAS: dict[tuple[str, str], str]` en `build_db_materiales.py`,
  con clave `(hoja, celda)` y valor el motivo escrito. Lo consume `TestParidadHojaParche`
  y lo consumirán las tareas posteriores que retiren más celdas heredadas.

- [x] **Step 1: Escribir la prueba que falla**

En `scripts/test_dashboard.py`, dentro de `TestParidadHojaParche`, sustituir el método
`test_todas_las_formulas_y_literales` por una versión que consulte las divergencias:

```python
    def test_todas_las_formulas_y_literales(self, wb):
        import build_db_materiales as B
        oracle = cargar_oracle_parche()
        ws = wb[self.HOJA]
        for celda, esperado in oracle["formulas"].items():
            motivo = B.DIVERGENCIAS_DECLARADAS.get((self.HOJA, celda))
            if motivo:
                # Celda que el build ya no reproduce a proposito. Se exige que
                # este VACIA: una divergencia declarada que resulta traer otro
                # valor es un error, no una divergencia.
                assert ws[celda].value is None, f"{celda}: {motivo}"
                continue
            assert ws[celda].value == esperado, celda
```

Y añadir la prueba de que la lista no crece sin motivo:

```python
    def test_toda_divergencia_declara_motivo(self):
        import build_db_materiales as B
        for (hoja, celda), motivo in B.DIVERGENCIAS_DECLARADAS.items():
            assert isinstance(motivo, str) and len(motivo) > 20, (hoja, celda)
```

- [x] **Step 2: Correr las pruebas para verlas fallar**

Run (desde `outputs/Base_Datos_Materiales_ASME/scripts`):
```powershell
python -m pytest test_dashboard.py::TestParidadHojaParche -q
```
Espera: FAIL con `AttributeError: module 'build_db_materiales' has no attribute 'DIVERGENCIAS_DECLARADAS'`.

- [x] **Step 3: Declarar las divergencias en el builder**

En `scripts/build_db_materiales.py`, junto a las demás constantes de módulo (misma zona
que `HOJAS_HEREDADAS`/`TEXTOS_HEREDADOS`):

```python
# Celdas que el *oracle* del Art. 212 declara pero que el build ya NO reproduce a
# proposito. El oracle sigue siendo la hoja heredada tal como se capturo; esta lista
# es la unica forma declarada de apartarse de ella, y cada entrada lleva su motivo.
# Sin esto, la alternativa era regenerar el oracle desde el build nuevo — con lo que
# dejaria de ser un control independiente y pasaria a ser un volcado de si mismo.
DIVERGENCIAS_DECLARADAS = {
    ("Parche_PCC2_Art212", "A22"): (
        "Fila retirada: el campo descriptivo de material salia de una lista fija de "
        "seis items, prohibida por la regla 12. La Seccion 7 ya resuelve ese mismo "
        "material con la cascada auditada."),
    ("Parche_PCC2_Art212", "B22"): "Idem A22: la fila entera se retira.",
    ("Parche_PCC2_Art212", "C22"): "Idem A22: la fila entera se retira.",
    ("Parche_PCC2_Art212", "D22"): "Idem A22: la fila entera se retira.",
    ("Parche_PCC2_Art212", "G22"): "Idem A22: la fila entera se retira.",
    ("Parche_PCC2_Art212", "A23"): (
        "Fila retirada por el mismo motivo que A22: el material del parche lo "
        "resuelve la cascada de la Seccion 7, no una lista fija."),
    ("Parche_PCC2_Art212", "B23"): "Idem A23: la fila entera se retira.",
    ("Parche_PCC2_Art212", "C23"): "Idem A23: la fila entera se retira.",
    ("Parche_PCC2_Art212", "D23"): "Idem A23: la fila entera se retira.",
    ("Parche_PCC2_Art212", "G23"): "Idem A23: la fila entera se retira.",
}
```

> Antes de escribirlo, **comprobar contra el *oracle*** qué celdas de las filas 22 y 23
> declara de verdad, y ajustar la lista a esas y solo esas:
> ```powershell
> python -c "import json; o=json.load(open('parche_art212_ref.json',encoding='utf-8')); print(sorted(k for k in o['formulas'] if k[1:] in ('22','23')))"
> ```
> Si aparece alguna celda que la lista de arriba no contempla, añadirla con su motivo;
> si alguna de las de arriba no existe en el *oracle*, quitarla. La lista tiene que
> casar exactamente con lo que el *oracle* declara, o `test_todas_las_formulas_y_literales`
> seguirá comparando una celda que ya nadie escribe.

- [x] **Step 4: Borrar las dos filas del motor**

En `build_parche_art212`, eliminar el bloque completo que hoy escribe las filas 22 y 23:
los dos `lab(...)`, los dos `ws.cell(.., 2, "—")`, los dos `inp(...)`, sus `com22`/`com23`,
la constante `ref_mat`, y **la `DataValidation` de `D22:D23` entera** (la construcción
directa con `formula1='"A106 Gr.B,A516 Gr.70,A105,A285 Gr.C,A333 Gr.6,A53 Gr.B"'` y su
`dv_mat.sqref = "D22:D23"`).

Dejar en su lugar un comentario que explique el hueco, para que nadie lo rellene por
reflejo:

```python
    # Filas 22-23 (material descriptivo de tuberia y de parche) retiradas: salian de
    # una lista fija, que la regla 12 prohibe. El material que rige el calculo lo
    # resuelve la cascada de la Seccion 7 contra DB_B31_3 / DB_BPVC_IID. Las filas se
    # dejan VACIAS a proposito, declaradas en DIVERGENCIAS_DECLARADAS; no se reutilizan
    # para otra cosa, o el *oracle* dejaria de cuadrar sin que nadie se entere.
```

- [x] **Step 5: Correr las pruebas**

Run:
```powershell
python -m pytest test_dashboard.py::TestParidadHojaParche test_dashboard.py::TestBuildParcheContraOracle -q
```
Espera: PASS. Si falla en una celda de la fila 22 o 23, la lista del Step 3 no casa con
lo que el *oracle* declara — corregirla, no relajar la prueba.

- [x] **Step 6: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/test_dashboard.py
git commit -m "Retira el material de lista fija del Art. 212 y declara la divergencia"
```

---

## Tarea 2: Desplegables bloqueantes en los doce motores

**Files:**
- Modify: `scripts/build_db_materiales.py` (`dv_list`, ~364-370)
- Modify: `scripts/test_dashboard.py` (prueba nueva)

**Interfaces:**
- Consume: `dv_list(ws, cell, formula, comentario=None)` — la firma **no cambia**.
- Produce: toda `DataValidation` de tipo lista del libro sale con
  `showErrorMessage=True, errorStyle="stop"`.

- [ ] **Step 1: Barrer las semillas antes de tocar nada**

Con `errorStyle="stop"`, un valor sembrado que no esté en su propia lista deja la celda
en un estado que Excel rechaza. Hay que saberlo **antes**, no después. Script de un solo
uso (escribirlo en el scratchpad, no versionarlo):

```python
import openpyxl, re
wb = openpyxl.load_workbook(r"..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm", keep_vba=True)
malas = []
for ws in wb.worksheets:
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list":
            continue
        f1 = str(dv.formula1 or "")
        if not f1.startswith('"'):
            continue          # origen de rango: se comprueba en Excel, no aqui
        items = [x.strip() for x in f1.strip('"').split(",")]
        for rng in dv.sqref.ranges:
            for fila in ws[str(rng)]:
                for c in (fila if isinstance(fila, tuple) else (fila,)):
                    v = c.value
                    if v is None or (isinstance(v, str) and v.startswith("=")):
                        continue
                    if str(v).strip() not in items:
                        malas.append(f"{ws.title}!{c.coordinate} = {v!r} no esta en {items}")
print(len(malas), "semillas fuera de lista")
print("\n".join(malas))
```

Run:
```powershell
python <script>
```
Espera: `0 semillas fuera de lista`. **Si sale distinto de 0, parar y reportarlo**: cada
caso hay que resolverlo a mano (o el valor sembrado está mal, o la lista está incompleta)
y decidirlo es del ingeniero, no de esta tarea.

- [ ] **Step 2: Escribir la prueba que falla**

En `scripts/test_dashboard.py`, al final:

```python
class TestValidacionesBloqueantes:
    """Un desplegable que no rechaza lo que no esta en su lista es una sugerencia,
    no una restriccion — y el error de tecleo sigue siendo posible. Con
    showErrorMessage=False (como estaba hasta esta tarea) Excel aceptaba en
    silencio cualquier valor escrito a mano."""

    def test_toda_validacion_de_lista_bloquea(self, wb):
        flojas = []
        for ws in wb.worksheets:
            for dv in ws.data_validations.dataValidation:
                if dv.type != "list":
                    continue
                if not dv.showErrorMessage or dv.errorStyle != "stop":
                    flojas.append(f"{ws.title}!{sorted(str(r) for r in dv.sqref.ranges)[0]}")
        assert flojas == [], flojas
```

- [ ] **Step 3: Correr la prueba para verla fallar**

Run:
```powershell
python -m pytest test_dashboard.py::TestValidacionesBloqueantes -q
```
Espera: FAIL, con una lista larga (hoy **ninguna** validación del libro bloquea).

- [ ] **Step 4: Hacer bloqueante `dv_list`**

En `scripts/build_db_materiales.py`, sustituir el cuerpo de `dv_list` (364-370):

```python
def dv_list(ws, cell, formula, comentario=None):
    # errorStyle="stop": Excel RECHAZA lo que no este en la lista, en vez de
    # aceptarlo en silencio. Es la mitad que le faltaba a la regla 14 — apuntar el
    # desplegable a la base no sirve de nada si el usuario puede teclear al lado.
    # Aviso honesto y deliberado: el bloqueo actua al TECLEAR; no al pegar ni al
    # escribir por macro. Reduce el error de dedo, no lo vuelve imposible.
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showErrorMessage=True, errorStyle="stop",
                        errorTitle="Valor fuera de lista",
                        error="Elija uno de los valores de la lista desplegable.")
    ws.add_data_validation(dv)
    dv.add(ws[cell])
    if comentario:
        _nota(ws[cell], comentario)
```

- [ ] **Step 5: Reconstruir y correr las pruebas**

Run:
```powershell
python build_db_materiales.py --resources ..\..\..\resources --in ..\..\..\templates\maestro_con_macros.xlsm --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
```
Espera: verde, incluida `TestValidacionesBloqueantes`.

- [ ] **Step 6: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/test_dashboard.py outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm
git commit -m "Los desplegables del libro rechazan lo que no esta en su lista"
```

> El `.xlsm` entra en el commit: es la convención del repositorio (las dos fases
> anteriores lo commitearon con el código que lo construye) y la revisión final del plan
> anterior marcó su omisión como defecto Critical.

---

## Tarea 3: El guardia — ninguna lista fija nueva en una hoja de motor

**Files:**
- Modify: `scripts/build_db_materiales.py` (`LITERALES_PERMITIDOS`, `HOJAS_DE_MOTOR`)
- Modify: `scripts/verificar.py` (§5, ~740-763)

**Interfaces:**
- Produce: `HOJAS_DE_MOTOR: tuple[str, ...]` y
  `LITERALES_PERMITIDOS: dict[str, str]` (lista literal → motivo) en el builder;
  `verificar.py` los importa.

- [ ] **Step 1: Declarar las dos listas en el builder**

En `scripts/build_db_materiales.py`, junto a `HOJAS_HEREDADAS`:

```python
# Las doce hojas donde el ingeniero introduce datos. El guardia de verificar.py §5
# solo mira estas: una DB_*, MAP_* o NAV_* no lleva entradas y no le aplica.
HOJAS_DE_MOTOR = (MOTOR, COLLAR_MOTOR,
                  "Buscar_B31_3", "Buscar_BPVC_IID", "Buscar_BPVC_IID_B",
                  "Buscar_Su", "Buscar_Sy", "Buscar_Prop_IID",
                  "Buscar_Prop_B31_3", "Buscar_B31_B1",
                  "Buscar_Ec_A2", "Buscar_Ej_A3")

# Listas literales que SI pueden quedarse: son modos y booleanos del propio motor,
# no datos tabulados por un codigo. Todo lo demas que sea literal en una hoja de
# motor es una lista fija espejando una tabla — exactamente lo que la regla 12
# prohibe— y el guardia lo rechaza.
LITERALES_PERMITIDOS = {
    '"SI,US"': "Conmutador de unidades (regla 10). No es un dato tabulado.",
    '"Interpolado,Tabulado-conservador"': (
        "Modo de lectura de S(T): la interpolacion lineal la autoriza el propio "
        "codigo (p.ej. para. A302.3.1(b)); no es una tabla de valores."),
    '"Si,No"': "Booleano de examen UT / defecto circunferencial del Art. 206.",
}
```

> `Type A`/`Type B` del Art. 206 se construye con f-string a partir de `TIPO_A`/`TIPO_B`.
> Localizar su literal real con
> `python -c "import build_db_materiales as B; print(f'\"{B.TIPO_A},{B.TIPO_B}\"')"`
> y añadirlo a `LITERALES_PERMITIDOS` con el motivo
> «Los dos tipos de sleeve que define el propio Art. 206 (206-1.1.1 / 206-1.1.2).»

- [ ] **Step 2: Escribir el guardia en `verificar.py`**

En `scripts/verificar.py` §5, **después** del bloque que hoy cuenta las validaciones no
portables (no sustituirlo: ese criterio de portabilidad sigue valiendo para el resto del
libro), añadir:

```python
    log("Regla 12/14: en una hoja de motor, una validacion de lista debe (a) bloquear "
        "y (b) tener origen de RANGO, no una lista literal — salvo los modos y "
        "booleanos declarados en LITERALES_PERMITIDOS.")
    log("")
    infractoras = []
    for w in wb.worksheets:
        if w.title not in B.HOJAS_DE_MOTOR:
            continue
        for dv in w.data_validations.dataValidation:
            if dv.type != "list":
                continue
            donde = f"{w.title}!{sorted(str(x) for x in dv.sqref.ranges)[0]}"
            f1 = str(dv.formula1 or "")
            if not dv.showErrorMessage or dv.errorStyle != "stop":
                infractoras.append(f"{donde} (no bloquea)")
            if f1.startswith('"') and f1 not in B.LITERALES_PERMITIDOS:
                infractoras.append(f"{donde} (lista fija: {f1[:60]})")
    log(f"Validaciones infractoras en hojas de motor: **{len(infractoras)}**"
        + ("" if not infractoras else "  → " + ", ".join(infractoras[:10])))
    fallos += len(infractoras)
    log("")
```

> Comprobar cómo se llama en `verificar.py` el acumulador de fallos de la sección y el
> alias del builder importado; arriba se asumen `fallos` y `B`. Si el archivo usa otros
> nombres, adaptarlos — **no** introducir un contador nuevo en paralelo.

- [ ] **Step 3: Correr `verificar.py`**

Run:
```powershell
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```
Espera: §5 con `Validaciones infractoras en hojas de motor: **0**` y el script
devolviendo 0. Las cinco celdas del inventario (D18/D19 de los dos motores) **todavía
son literales** en este punto — así que este paso confirmará que el guardia las detecta:
espera **4 infractoras** y `verificar.py` devolviendo distinto de 0.

- [ ] **Step 4: Declarar la deuda conocida y volver a verde**

Las cuatro infractoras son reales y las arreglan las Tareas 7-9; el guardia no puede
quedarse rojo tres tareas. Añadir en el builder, junto a las listas anteriores:

```python
# Deuda declarada, con fecha de vencimiento: las Tareas 7-9 de este plan repuntan
# NPS y cedula de los dos motores contra DB_B36_10/DB_B36_19. Hasta entonces el
# guardia las tolera nombrandolas una a una — nunca por patron, para que una lista
# fija NUEVA en la misma celda no entre por el mismo hueco.
DEUDA_LISTA_FIJA = {
    (MOTOR, "D18"): "NPS: pendiente de DB_B36_10 (Tarea 7).",
    (MOTOR, "D19"): "Cedula: pendiente de DB_B36_10 (Tarea 7).",
    (COLLAR_MOTOR, "D18"): "NPS: pendiente de DB_B36_10 (Tarea 9).",
    (COLLAR_MOTOR, "D19"): "Cedula: pendiente de DB_B36_10 (Tarea 9).",
}
```

Y en el guardia, antes de acusar una lista fija:

```python
            celda = sorted(str(x) for x in dv.sqref.ranges)[0].split(":")[0]
            if f1.startswith('"') and f1 not in B.LITERALES_PERMITIDOS \
                    and (w.title, celda) not in B.DEUDA_LISTA_FIJA:
                infractoras.append(f"{donde} (lista fija: {f1[:60]})")
```

La Tarea 10 borra `DEUDA_LISTA_FIJA` entera y comprueba que el guardia sigue en 0.

- [ ] **Step 5: Correr `verificar.py` y las pruebas**

Run:
```powershell
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
```
Espera: `verificar.py` devuelve 0; las tres suites verdes.

- [ ] **Step 6: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/verificar.py
git commit -m "Guardia: ninguna lista fija nueva en una hoja de motor"
```

---

## Tarea 4: Depositar las dos extracciones en `resources/`

**Files:**
- Create: `resources/ASME B36/b36_10m_2022/table_dimensiones.json`
- Create: `resources/ASME B36/b36_10m_2022/meta.json`
- Create: `resources/ASME B36/b36_19m_2022/table_dimensiones.json`
- Create: `resources/ASME B36/b36_19m_2022/meta.json`

**Interfaces:**
- Produce: las dos rutas de `resources/` que consumen las Tareas 5 y 6.

- [ ] **Step 1: Copiar los JSON entregados**

Origen (directorio de subidas de la sesión del 2026-09-10 — **es de sesión, no un sitio
del que el proyecto pueda depender**, por eso esta tarea existe):

```powershell
$u = "C:\Users\User\.claude\uploads\b554f20f-43a0-4c53-b8e7-767a3321bf31"
New-Item -ItemType Directory -Force "..\..\..\resources\ASME B36\b36_10m_2022"
New-Item -ItemType Directory -Force "..\..\..\resources\ASME B36\b36_19m_2022"
Copy-Item "$u\974d90fb-datalaboutputASME_B36.10M__Welded_and_Seamless_Wroth_Steel_Pipe_2022.pdf.json" `
          "..\..\..\resources\ASME B36\b36_10m_2022\table_dimensiones.json"
Copy-Item "$u\ecc8b41b-datalaboutputASME_B36.19M__Stainless_Steel_Pipe_2022.pdf.json" `
          "..\..\..\resources\ASME B36\b36_19m_2022\table_dimensiones.json"
```

Si el directorio de subidas ya no existe, **parar y pedir los archivos al ingeniero**;
no buscar sustitutos ni reextraer del PDF (Regla nº 1: el PDF no es fuente de datos).

- [ ] **Step 2: Escribir los `meta.json` de procedencia**

`resources/ASME B36/b36_10m_2022/meta.json`:

```json
{
  "norma": "ASME B36.10M-2022",
  "titulo": "Welded and Seamless Wrought Steel Pipe",
  "origen": "Extraccion datalab del PDF, entregada por el ingeniero el 2026-09-10",
  "pdf_referencia": "ASME B36.10M - Welded and Seamless Wroth Steel Pipe 2022.pdf",
  "pdf_no_versionado": true,
  "pdf_pages_base": "0-based (declarado por la clave 'page' de cada children)",
  "tablas_de_dimensiones": "paginas 13 a 31",
  "verificado_contra_pdf": false
}
```

`resources/ASME B36/b36_19m_2022/meta.json`: igual, con
`"norma": "ASME B36.19M-2022"`, `"titulo": "Stainless Steel Pipe"`,
`"pdf_referencia": "ASME B36.19M - Stainless Steel Pipe 2022.pdf"` y
`"tablas_de_dimensiones": "paginas 11 a 13"`.

> `"pdf_pages_base"` va declarado desde el principio **a propósito**: equivocarse en la
> base de folios ya costó caro una vez en este proyecto (ver `CLAUDE.md`). Confirmarlo
> antes de escribirlo, con una fila que aparezca **una sola vez** — nunca con el rótulo
> de la tabla, que se repite en cada página de continuación y casa con las dos
> convenciones.

- [ ] **Step 3: Confirmar la base de folios y marcar el meta**

Abrir el PDF de B36.10M en la ruta de normas
(`C:\Users\User\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES\0-STANDARDS\ASME\B36-DIMENSION OF  WROUGHT STEEL PIPES\`)
y comprobar que la fila `NPS 1/8 (6) · STD · Sch 40 · OD 0.405 (10.29) · espesor
0.068 (1.73)` está en la página que el JSON declara para su tabla. Es un dato que
aparece una sola vez.

Con el resultado, poner `"verificado_contra_pdf": true` y añadir
`"verificado_con": "NPS 1/8 Sch 40 STD, espesor 0.068 in (1.73 mm)"`. Repetir con
B36.19M usando una fila igual de única de su tabla de la página 11.

- [ ] **Step 4: Commit**

```powershell
git add "..\..\..\resources\ASME B36"
git commit -m "Deposita las extracciones de B36.10M y B36.19M en resources"
```

---

## Tarea 5: `b36_dimensiones.py` — de `<table>` a filas normalizadas

**Files:**
- Create: `scripts/b36_dimensiones.py`
- Create: `scripts/test_b36_dimensiones.py`

**Interfaces:**
- Consume: los dos `table_dimensiones.json` de la Tarea 4.
- Produce:
  - `cargar(ruta: Path) -> list[dict]` — una fila por (NPS, cédula) con las claves
    `nps_impreso`, `nps_in`, `dn_mm`, `identificacion`, `cedula`, `designador`,
    `od_in`, `od_mm`, `t_in`, `t_mm`, `peso_lb_ft`, `peso_kg_m`, `pagina`.
  - `partir_doble_unidad(txt: str) -> tuple[float | None, float | None]`
  - `partir_nps(txt: str) -> tuple[str, float | None, float | None]`
  - `designador(identificacion: str, cedula: str) -> str`

- [ ] **Step 1: Escribir las pruebas primero**

`scripts/test_b36_dimensiones.py`:

```python
"""Pruebas del parser de B36.10M / B36.19M. Las filas de muestra estan copiadas
LITERALES de la extraccion (pagina 13 de B36.10M): no se inventan valores."""
import b36_dimensiones as B


def test_parte_celda_de_doble_unidad():
    assert B.partir_doble_unidad("0.405 (10.29)") == (0.405, 10.29)
    assert B.partir_doble_unidad("0.068 (1.73)") == (0.068, 1.73)


def test_doble_unidad_ausente_no_inventa():
    assert B.partir_doble_unidad("...") == (None, None)
    assert B.partir_doble_unidad("") == (None, None)


def test_parte_el_nps_fraccionario():
    assert B.partir_nps("1/8 (6)") == ("1/8 (6)", 0.125, 6.0)
    assert B.partir_nps("12 (300)") == ("12 (300)", 12.0, 300.0)
    assert B.partir_nps("2 1/2 (65)") == ("2 1/2 (65)", 2.5, 65.0)


def test_designador_combina_las_dos_designaciones():
    # El codigo publica una tubería por numero de cedula, por identificacion, por
    # las dos, o por una sola. El designador las junta sin perder ninguna.
    assert B.designador("STD", "40") == "40 (STD)"
    assert B.designador("XS", "80") == "80 (XS)"
    assert B.designador("...", "10") == "10"
    assert B.designador("XXS", "...") == "XXS"


def test_no_convierte_el_marcador_del_codigo_a_vacio():
    # '...' es 'no aplica' IMPRESO por la norma (regla 9). La fila lo conserva.
    filas = B.cargar(B.RUTA_B3610)
    xxs = [f for f in filas if f["nps_impreso"] == "1/8 (6)" and f["identificacion"] == "XXS"]
    assert len(xxs) == 1
    assert xxs[0]["cedula"] == "..."
    assert xxs[0]["designador"] == "XXS"


def test_caso_semilla_del_art_212():
    # NPS 12, Sch 20 -> espesor 6.35 mm. Es el caso precargado del motor Art. 212:
    # si esta fila no resuelve, el motor deja de reproducir su propio caso.
    filas = B.cargar(B.RUTA_B3610)
    f = [x for x in filas if x["nps_in"] == 12.0 and x["cedula"] == "20"]
    assert len(f) == 1
    assert f[0]["t_mm"] == 6.35
    assert f[0]["od_mm"] == 323.8


def test_b3619_no_trae_columna_de_identificacion():
    filas = B.cargar(B.RUTA_B3619)
    assert filas, "la extraccion de B36.19M no produjo filas"
    assert all(f["identificacion"] == "" for f in filas)
    assert any(f["cedula"].endswith("S") for f in filas), "faltan las cedulas 5S/10S/40S/80S"
```

- [ ] **Step 2: Correr las pruebas para verlas fallar**

Run:
```powershell
python -m pytest test_b36_dimensiones.py -q
```
Espera: FAIL con `ModuleNotFoundError: No module named 'b36_dimensiones'`.

- [ ] **Step 3: Escribir el parser**

`scripts/b36_dimensiones.py`:

```python
"""Convierte las extracciones de ASME B36.10M / B36.19M en filas normalizadas,
una por (NPS, cedula).

A diferencia de la Seccion II, aqui los bloques `Table` SI traen <table> HTML
real, asi que no hay que recomponer columnas desde `Line` y `bbox`: se parsea el
HTML y se parten las celdas de doble unidad.

Es libreria y CLI: el builder la importa para escribir DB_B36_10 / DB_B36_19, y
`--informe` mide y reporta sin escribir nada.
"""
import json
import re
from fractions import Fraction
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3] / "resources" / "ASME B36"
RUTA_B3610 = RAIZ / "b36_10m_2022" / "table_dimensiones.json"
RUTA_B3619 = RAIZ / "b36_19m_2022" / "table_dimensiones.json"

# El codigo imprime '...' donde una designacion no aplica (una XXS no lleva numero
# de cedula; una Sch 10 no lleva identificacion). Se conserva tal cual (regla 9).
NO_APLICA = "..."


def _texto(html_celda):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html_celda)).strip()


def partir_doble_unidad(txt):
    """'0.405 (10.29)' -> (0.405, 10.29). Sin par de numeros -> (None, None)."""
    m = re.match(r"^\s*([\d.]+)\s*\(([\d.]+)\)\s*$", txt or "")
    if not m:
        return (None, None)
    return (float(m.group(1)), float(m.group(2)))


def partir_nps(txt):
    """'2 1/2 (65)' -> ('2 1/2 (65)', 2.5, 65.0).

    El decimal no se imprime en la norma: se deriva para poder ordenar la base y
    para que el caso semilla del Art. 212 (NPS 12) siga casando. El texto impreso
    se conserva intacto como `nps_impreso` (regla 9) y es lo que ve el ingeniero.
    """
    impreso = (txt or "").strip()
    m = re.match(r"^\s*([\d\s/]+?)\s*\((\d+(?:\.\d+)?)\)\s*$", impreso)
    if not m:
        return (impreso, None, None)
    crudo, dn = m.group(1).strip(), float(m.group(2))
    total = sum(Fraction(p) for p in crudo.split())
    return (impreso, float(total), dn)


def designador(identificacion, cedula):
    """Une las dos designaciones que publica el codigo en una sola etiqueta.

    El ingeniero no tiene por que saber de antemano si SU tuberia se publica por
    numero de cedula, por identificacion o por las dos: elige una cosa y ve ambas.
    """
    ident = (identificacion or "").strip()
    ced = (cedula or "").strip()
    tiene_ident = ident and ident != NO_APLICA
    tiene_ced = ced and ced != NO_APLICA
    if tiene_ced and tiene_ident:
        return f"{ced} ({ident})"
    if tiene_ced:
        return ced
    if tiene_ident:
        return ident
    return ""


def _tablas_de_dimensiones(doc):
    """Bloques Table cuya cabecera nombra NPS y Wall Thickness. Descarta la tabla
    de cambios de edicion ('Page / Location / Change') sin depender de su pagina."""
    for pag in doc["children"]:
        for b in (pag.get("children") or []):
            if b.get("block_type") != "Table":
                continue
            html = b.get("html") or ""
            cab = " ".join(_texto(c) for c in re.findall(r"<th>(.*?)</th>", html, re.S))
            if "NPS" in cab and "Wall" in cab:
                yield pag.get("page"), html


def cargar(ruta):
    doc = json.loads(Path(ruta).read_text(encoding="utf-8"))
    filas = []
    for pagina, html in _tablas_de_dimensiones(doc):
        cabeceras = [_texto(c) for c in re.findall(r"<th>(.*?)</th>", html, re.S)]
        idx = {h: i for i, h in enumerate(cabeceras)}
        col_ident = next((i for h, i in idx.items() if h.startswith("Identification")), None)
        col_nps = next(i for h, i in idx.items() if h.startswith("NPS"))
        col_ced = next(i for h, i in idx.items() if h.startswith("Schedule"))
        col_od = next(i for h, i in idx.items() if h.startswith("Outside"))
        col_t = next(i for h, i in idx.items() if h.startswith("Wall"))
        col_w = next((i for h, i in idx.items() if h.startswith("Plain")), None)
        for tr in re.findall(r"<tr>(.*?)</tr>", html, re.S):
            celdas = [_texto(c) for c in re.findall(r"<t[hd]>(.*?)</t[hd]>", tr, re.S)]
            if len(celdas) < len(cabeceras) or celdas[col_nps].startswith("NPS"):
                continue
            impreso, nps_in, dn = partir_nps(celdas[col_nps])
            ident = celdas[col_ident] if col_ident is not None else ""
            od_in, od_mm = partir_doble_unidad(celdas[col_od])
            t_in, t_mm = partir_doble_unidad(celdas[col_t])
            w_lb, w_kg = partir_doble_unidad(celdas[col_w]) if col_w is not None else (None, None)
            filas.append(dict(
                nps_impreso=impreso, nps_in=nps_in, dn_mm=dn,
                identificacion="" if ident == NO_APLICA and col_ident is None else ident,
                cedula=celdas[col_ced],
                designador=designador(ident, celdas[col_ced]),
                od_in=od_in, od_mm=od_mm, t_in=t_in, t_mm=t_mm,
                peso_lb_ft=w_lb, peso_kg_m=w_kg, pagina=pagina))
    return filas


def main():
    import argparse
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--informe", type=Path, help="ruta del informe a refrescar")
    a = p.parse_args()
    lineas = []
    for nombre, ruta in (("B36.10M", RUTA_B3610), ("B36.19M", RUTA_B3619)):
        filas = cargar(ruta)
        nps = sorted({f["nps_impreso"] for f in filas})
        sin_t = [f for f in filas if f["t_mm"] is None]
        lineas.append(f"- **{nombre}**: {len(filas)} filas, {len(nps)} NPS distintos, "
                      f"{len(sin_t)} sin espesor legible.")
        print(lineas[-1])
    if a.informe:
        a.informe.write_text("# Revision de las bases dimensionales B36\n\n"
                             + "\n".join(lineas) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Correr las pruebas**

Run:
```powershell
python -m pytest test_b36_dimensiones.py -q
```
Espera: PASS las 7. Si `test_caso_semilla_del_art_212` falla, **no tocar la prueba**:
NPS 12 Sch 20 con espesor 6,35 mm es el caso que el motor ya reproduce hoy contra
`Datos_Ref`, y si la base nueva no lo da, el parser está mal o la extracción tiene un
hueco — cualquiera de las dos cosas hay que resolverla antes de seguir.

- [ ] **Step 5: Correr el CLI y mirar los totales**

Run:
```powershell
python b36_dimensiones.py --informe ..\Revision_Dimensiones_B36.md
```
Espera: ~800 filas para B36.10M y ~117 para B36.19M, y **0 sin espesor legible**.
Un número apreciable de filas sin espesor significa celdas que el parser no supo partir:
investigarlas antes de construir una base con huecos.

- [ ] **Step 6: Commit**

```powershell
git add scripts/b36_dimensiones.py scripts/test_b36_dimensiones.py ..\Revision_Dimensiones_B36.md
git commit -m "Parser de las dimensiones de B36.10M y B36.19M"
```

---

## Tarea 6: `DB_B36_10` y `DB_B36_19`, y su auditoría

**Files:**
- Modify: `scripts/build_db_materiales.py` (`build_db_b36`, llamada en `main`)
- Modify: `scripts/verificar.py` (§11 nueva)
- Modify: `scripts/test_dashboard.py` (pruebas de las dos hojas)

**Interfaces:**
- Consume: `b36_dimensiones.cargar`.
- Produce: `build_db_b36(wb, norma) -> dict` con `{"sheet": str, "last_row": int}`,
  el mismo contrato que devuelven las demás bases y que consumen las Tareas 7 y 9.

- [ ] **Step 1: Escribir las pruebas de la hoja**

En `scripts/test_dashboard.py`:

```python
class TestBasesDimensionalesB36:
    HOJAS = ("DB_B36_10", "DB_B36_19")

    def test_existen_y_no_llevan_formulas(self, wb):
        for h in self.HOJAS:
            ws = wb[h]
            for fila in ws.iter_rows(min_row=1):
                for c in fila:
                    assert not (isinstance(c.value, str) and c.value.startswith("=")), \
                        f"{h}!{c.coordinate}: una base no lleva formulas"

    def test_orden_contiguo_por_nps(self, wb):
        # Regla 5: la cascada dependiente usa COUNTIF/INDEX/MATCH sobre bloques
        # contiguos. Si un NPS aparece en dos tramos separados, la lista de cedulas
        # de ese NPS sale truncada y nadie se entera.
        import build_db_materiales as B
        ws = wb["DB_B36_10"]
        col = B.COL_B36["nps_impreso"]
        vistos, ultimo = set(), None
        for r in range(B.R_DATA, ws.max_row + 1):
            v = ws.cell(r, col).value
            if v != ultimo:
                assert v not in vistos, f"NPS {v!r} reaparece fuera de su bloque"
                vistos.add(v)
                ultimo = v

    def test_caso_semilla_presente(self, wb):
        import build_db_materiales as B
        ws = wb["DB_B36_10"]
        cn, cc, ct = (B.COL_B36["nps_in"], B.COL_B36["cedula"], B.COL_B36["t_mm"])
        hit = [r for r in range(B.R_DATA, ws.max_row + 1)
               if ws.cell(r, cn).value == 12.0 and str(ws.cell(r, cc).value) == "20"]
        assert len(hit) == 1
        assert ws.cell(hit[0], ct).value == 6.35
```

- [ ] **Step 2: Correr para verlas fallar**

Run:
```powershell
python -m pytest test_dashboard.py::TestBasesDimensionalesB36 -q
```
Espera: FAIL con `KeyError: 'DB_B36_10'`.

- [ ] **Step 3: Construir las dos hojas**

En `scripts/build_db_materiales.py`, junto a las demás funciones `build_*` de bases:

```python
import b36_dimensiones

# Layout de las dos bases dimensionales. El indice de columna se declara una sola
# vez y lo consumen el builder, las pruebas y verificar.py: si alguien inserta una
# columna, se entera todo el mundo a la vez.
COL_B36 = {"nps_impreso": 1, "nps_in": 2, "dn_mm": 3, "designador": 4,
           "cedula": 5, "identificacion": 6, "od_in": 7, "od_mm": 8,
           "t_in": 9, "t_mm": 10, "peso_lb_ft": 11, "peso_kg_m": 12}

CABECERA_B36 = ["NPS impreso", "NPS (in)", "DN (mm)", "Designador",
                "Schedule No.", "Identification", "OD (in)", "OD (mm)",
                "Espesor (in)", "Espesor (mm)", "Peso (lb/ft)", "Peso (kg/m)"]


def build_db_b36(wb, norma):
    """Base dimensional de una de las dos normas de tuberia.

    Una fila por (NPS, cedula), ordenada por NPS y luego por espesor para que cada
    NPS sea un bloque contiguo (regla 5) — lo exige la lista dependiente de cedula
    de los motores, que resuelve con COUNTIF/INDEX/MATCH y no con matrices.

    Las dos unidades vienen en columnas separadas porque la norma las publica en la
    MISMA celda: el conmutador de los motores cambia de columna, nunca convierte
    (regla 10).
    """
    cfg = {"B36.10M": (b36_dimensiones.RUTA_B3610, "DB_B36_10",
                       "ASME B36.10M-2022 — Welded and Seamless Wrought Steel Pipe"),
           "B36.19M": (b36_dimensiones.RUTA_B3619, "DB_B36_19",
                       "ASME B36.19M-2022 — Stainless Steel Pipe")}[norma]
    ruta, nombre, titulo = cfg
    filas = b36_dimensiones.cargar(ruta)
    # Orden: NPS ascendente y, dentro de cada NPS, espesor ascendente. El espesor
    # ordena mejor que la cedula porque la cedula es texto y mezcla numeros con
    # STD/XS/XXS y con 5S/10S: ordenarla como texto pondria "10" antes que "5".
    filas.sort(key=lambda f: (f["nps_in"] if f["nps_in"] is not None else 1e9,
                              f["t_mm"] if f["t_mm"] is not None else 1e9))
    ws = new_sheet(wb, nombre, titulo,
                   f"Fuente: resources/ASME B36/{ruta.parent.name}/{ruta.name}")
    for j, h in enumerate(CABECERA_B36, start=1):
        c = ws.cell(R_HDR, j, h)
        c.font, c.fill, c.border = HDR_F, HDR_FILL, BOX_FRANJA
    for i, f in enumerate(filas):
        r = R_DATA + i
        for clave, j in COL_B36.items():
            ws.cell(r, j, f[clave]).font = DATA_F
    autosize(ws, {"A": 14, "B": 10, "C": 10, "D": 14, "E": 13, "F": 13,
                  "G": 11, "H": 11, "I": 13, "J": 13, "K": 13, "L": 13})
    return {"sheet": nombre, "last_row": R_DATA + len(filas) - 1}
```

En `main()`, llamarlas junto a las demás bases, **antes** de los motores (que las
consumirán a partir de la Tarea 7):

```python
    b3610 = build_db_b36(wb, "B36.10M")
    b3619 = build_db_b36(wb, "B36.19M")
```

- [ ] **Step 4: Reconstruir y correr las pruebas**

Run:
```powershell
python build_db_materiales.py --resources ..\..\..\resources --in ..\..\..\templates\maestro_con_macros.xlsm --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_dashboard.py::TestBasesDimensionalesB36 -q
```
Espera: PASS las 3.

- [ ] **Step 5: Auditar la hoja contra el JSON en `verificar.py`**

En `scripts/verificar.py`, sección nueva al final (numerarla como la siguiente
disponible; hoy la última es §10):

```python
    log("## 11. Bases dimensionales B36 (hoja contra JSON de resources)")
    log("")
    import b36_dimensiones
    disc = 0
    for norma, ruta, hoja in (("B36.10M", b36_dimensiones.RUTA_B3610, "DB_B36_10"),
                              ("B36.19M", b36_dimensiones.RUTA_B3619, "DB_B36_19")):
        esperadas = b36_dimensiones.cargar(ruta)
        esperadas.sort(key=lambda f: (f["nps_in"] if f["nps_in"] is not None else 1e9,
                                      f["t_mm"] if f["t_mm"] is not None else 1e9))
        ws = wb[hoja]
        leidas = ws.max_row - B.R_DATA + 1
        if leidas != len(esperadas):
            log(f"- {norma}: la hoja tiene {leidas} filas y el JSON {len(esperadas)}")
            disc += 1
            continue
        for i, esp in enumerate(esperadas):
            r = B.R_DATA + i
            for clave, j in B.COL_B36.items():
                if ws.cell(r, j).value != esp[clave]:
                    disc += 1
                    if disc <= 10:
                        log(f"- {hoja}!{ws.cell(r, j).coordinate}: hoja="
                            f"{ws.cell(r, j).value!r} json={esp[clave]!r}")
        log(f"- {norma}: {len(esperadas)} filas x {len(B.COL_B36)} columnas auditadas.")
    log(f"Discrepancias: **{disc}**")
    fallos += disc
    log("")
```

- [ ] **Step 6: Correr `verificar.py` y las tres suites**

Run:
```powershell
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
```
Espera: §11 con `Discrepancias: **0**`; `verificar.py` devuelve 0; suites verdes.

- [ ] **Step 7: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/verificar.py scripts/test_dashboard.py outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm
git commit -m "DB_B36_10 y DB_B36_19: las dos bases dimensionales, auditadas"
```

---

## Tarea 7: Cascada dimensional en el Art. 212

> **EJECUTADA junto con la Tarea 8 (2026-09-11), con tres desvíos del diseño de
> abajo, todos deliberados:**
> 1. **Norma en la fila 22, sin correr el bloque.** El diseño ponía la norma en
>    `D17`, pero `D17` es el encabezado, y correr NPS/cédula/OD/espesor habría
>    reapuntado ~12 fórmulas de cálculo aguas abajo (D52, D54, D55, D56, D65,
>    D73, D78, E65, F65, E87) que referencian OD(D20)/espesor(D21). En su lugar
>    NPS@18, cédula@19, OD@20, espesor@21 se quedan; la norma ocupa la fila 22
>    (antes material descriptivo, ya retirado). Ninguna fórmula de cálculo cambia.
> 2. **Se fundió con la Tarea 8** (OD/espesor contra `DB_B36`): como el NPS pasa
>    a guardarse con formato `nps_impreso` (`'12 (300)'`), dejar OD/espesor en
>    `Datos_Ref` daba #N/D. Separarlas dejaba la hoja rota en el interin.
> 3. **Paridad contra el oracle Rev0:** se añadió `DIVERGENCIAS_REEMPLAZADAS`
>    (celda con contenido nuevo a propósito, exigida no-vacía, verificada por
>    `TestCascadaDimensionalArt212`), distinta de `DIVERGENCIAS_DECLARADAS`
>    (celda retirada, exigida vacía). También se añadió la columna auxiliar
>    `clave_ced` a `DB_B36` para listar cédulas sin blancos, y `n_nps`/`max_ced`
>    salen de `build_db_b36` (no números mágicos: NPS reales=45, cédulas máx=33).
> Verificado: `test_build_db + test_dashboard + test_secii_tablas` → 229 pasan.
> `verificar.py` (recálculo en Excel) queda para la corrida local del ingeniero.

**Files:**
- Modify: `scripts/build_db_materiales.py` (`build_parche_art212`, filas 18-19 y columnas ocultas)
- Modify: `scripts/test_dashboard.py`

**Interfaces:**
- Consume: `build_db_b36(...)` → `{"sheet", "last_row"}`; `COL_B36`.
- Produce: en la hoja del motor, `D17` = norma dimensional, `D18` = NPS impreso,
  `D19` = designador de cédula. Las tres con validación de **rango**, no literal.

- [ ] **Step 1: Escribir la prueba**

```python
class TestCascadaDimensionalArt212:
    HOJA = "Parche_PCC2_Art212"

    def test_las_tres_celdas_tienen_validacion_de_rango(self, wb):
        ws = wb[self.HOJA]
        origen = {}
        for dv in ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                origen[str(rng)] = str(dv.formula1 or "")
        for celda in ("D17", "D18", "D19"):
            f1 = origen.get(celda, "")
            assert f1.startswith("="), f"{celda}: origen {f1!r}, se esperaba un rango"
            assert not f1.startswith('"'), f"{celda}: sigue siendo lista fija"

    def test_la_norma_ofrece_las_dos_y_solo_las_dos(self, wb):
        import build_db_materiales as B
        ws = wb[self.HOJA]
        col = B.COL_NORMA_B36
        vals = [ws.cell(r, col).value for r in range(B.R_DATA, B.R_DATA + 2)]
        assert vals == ["B36.10M", "B36.19M"]
```

- [ ] **Step 2: Correr para verla fallar**

Run:
```powershell
python -m pytest test_dashboard.py::TestCascadaDimensionalArt212 -q
```
Espera: FAIL — `D17` no existe y `D18`/`D19` siguen con origen literal.

- [ ] **Step 3: Construir la cascada**

En `build_parche_art212`, sustituir el bloque de las filas 18-19 (los dos `inp` con sus
`dv_list` literales) por:

```python
    # --- Cascada dimensional: norma -> NPS -> cedula -------------------------
    # La norma se ELIGE, no se deriva del material: una tuberia inoxidable puede
    # fabricarse a cedulas B36.10M, asi que derivarla de la familia daria el
    # espesor equivocado en ese caso (y ademas acoplaria la Seccion 1 con la 7).
    com17 = ("Entrada: norma dimensional de la tuberia. B36.10M para acero al "
             "carbono y de baja aleacion; B36.19M para inoxidable con cedulas de "
             "la serie S. Gobierna las listas de NPS y cedula de abajo.")
    lab(17, "Norma dimensional", unidad="—", ref="ASME B36.10M / B36.19M", com=com17)
    inp("D17", "B36.10M", com17)

    ancho_b36 = _materializar_cascada_b36(ws, b3610, b3619, fila_norma=17,
                                          fila_nps=18, fila_ced=19)
```

Y añadir, junto a `construir_seccion7_material` (comparten patrón y conviene que se lean
seguidas):

```python
def _materializar_cascada_b36(ws, b3610, b3619, fila_norma, fila_nps, fila_ced):
    """Listas de norma -> NPS -> cedula, materializadas en columnas ocultas.

    La regla 2 prohibe una formula como origen de validacion, asi que cada lista se
    calcula en una columna oculta de la propia hoja y la validacion apunta a ese
    rango literal. Es el mismo mecanismo que la cascada de material de la Seccion 7
    (construir_seccion7_material), y por el mismo motivo.

    La lista de cedula depende de la norma Y del NPS: se resuelve con
    COUNTIF/INDEX/MATCH sobre el bloque contiguo de ese NPS en la base, que es lo
    que la regla 5 garantiza al ordenarla.
    """
    hl = get_column_letter
    col = COL_NORMA_B36
    # Nivel 0: las dos normas. Literal en la hoja, no en la validacion: sigue
    # siendo un rango como origen, y anadir una tercera norma manana es anadir una
    # fila aqui, no tocar el motor.
    ws.cell(R_HDR, col, "lista norma B36").font = SRC_F
    for k, n in enumerate(("B36.10M", "B36.19M")):
        ws.cell(R_DATA + k, col, n)
    ws.column_dimensions[hl(col)].hidden = True
    dv_list(ws, f"D{fila_norma}", f"=${hl(col)}${R_DATA}:${hl(col)}${R_DATA + 1}")

    def rango(info, clave):
        L = hl(COL_B36[clave])
        return f"{info['sheet']}!${L}${R_DATA}:${L}${info['last_row']}"

    sel = f"$D${fila_norma}"
    # Nivel 1: NPS distintos de la norma elegida. Se listan sin repetir tomando la
    # primera aparicion de cada bloque contiguo (COUNTIF sobre lo ya emitido seria
    # cuadratico y no hace falta: la base ya viene ordenada).
    col += 1
    L_nps = hl(col)
    ws.cell(R_HDR, col, "lista NPS B36").font = SRC_F
    n_nps = 40
    for k in range(1, n_nps + 1):
        ws.cell(R_DATA + k - 1, col).value = (
            f'=IFERROR(INDEX(IF({sel}="B36.10M",{rango(b3610, "nps_impreso")},'
            f'{rango(b3619, "nps_impreso")}),'
            f'MATCH({k},IF({sel}="B36.10M",{rango(b3610, "nps_orden")},'
            f'{rango(b3619, "nps_orden")}),0)),"")')
    ws.column_dimensions[L_nps].hidden = True
    dv_list(ws, f"D{fila_nps}",
            f"=${L_nps}${R_DATA}:${L_nps}${R_DATA + n_nps - 1}")

    # Nivel 2: cedulas del NPS elegido, en la norma elegida.
    col += 1
    L_ced = hl(col)
    ws.cell(R_HDR, col, "lista cedula B36").font = SRC_F
    n_ced = 20
    clave = f"$D${fila_nps}"
    for k in range(1, n_ced + 1):
        ws.cell(R_DATA + k - 1, col).value = (
            f'=IF(COUNTIF(IF({sel}="B36.10M",{rango(b3610, "nps_impreso")},'
            f'{rango(b3619, "nps_impreso")}),{clave})<{k},"",'
            f'INDEX(IF({sel}="B36.10M",{rango(b3610, "designador")},'
            f'{rango(b3619, "designador")}),'
            f'MATCH({clave},IF({sel}="B36.10M",{rango(b3610, "nps_impreso")},'
            f'{rango(b3619, "nps_impreso")}),0)+{k}-1))')
    ws.column_dimensions[L_ced].hidden = True
    dv_list(ws, f"D{fila_ced}",
            f"=${L_ced}${R_DATA}:${L_ced}${R_DATA + n_ced - 1}")
    return col
```

> **`nps_orden` no existe todavía en `COL_B36`.** Es una columna nueva que la base
> necesita para poder listar los NPS sin repetirlos: vale `1, 2, 3…` en la **primera**
> fila de cada bloque de NPS y queda vacía en las demás. Añadirla a `COL_B36`,
> a `CABECERA_B36` y a `build_db_b36` (Tarea 6) antes de escribir esto, y extender
> `TestBasesDimensionalesB36` y la auditoría §11 para que la cubran — si se añade una
> columna sin tocar esos tres sitios, la auditoría deja de cuadrar.
>
> **Comprobar `n_nps = 40` y `n_ced = 20` contra los datos reales** antes de fijarlos:
> ```powershell
> python -c "import b36_dimensiones as b; from collections import Counter; f=b.cargar(b.RUTA_B3610); print('NPS distintos', len({x['nps_impreso'] for x in f})); print('max cedulas por NPS', max(Counter(x['nps_impreso'] for x in f).values()))"
> ```
> Si algún NPS tiene más cédulas que `n_ced`, la lista sale truncada **en silencio**:
> subir la constante, no recortar la lista.

- [ ] **Step 4: Correr las pruebas**

Run:
```powershell
python build_db_materiales.py --resources ..\..\..\resources --in ..\..\..\templates\maestro_con_macros.xlsm --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_dashboard.py::TestCascadaDimensionalArt212 -q
```
Espera: PASS.

- [ ] **Step 5: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/test_dashboard.py outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm
git commit -m "Art. 212: NPS y cedula salen de DB_B36_10 / DB_B36_19"
```

---

## Tarea 8: OD y espesor del Art. 212 contra la base

> **EJECUTADA junto con la Tarea 7 (2026-09-11)** — ver la nota al inicio de la
> Tarea 7. OD(`D20`)/espesor(`D21`) leen de `DB_B36_10`/`DB_B36_19` por la clave
> `NPS|cédula` (`clave`), con el conmutador de edición vía `CHOOSE($AB$2, …)`.

**Files:**
- Modify: `scripts/build_db_materiales.py` (`build_parche_art212`, D20 y D21)
- Modify: `scripts/test_dashboard.py`

**Interfaces:**
- Consume: la cascada de la Tarea 7 (`D17`, `D18`, `D19`).
- Produce: `D20` (OD, mm) y `D21` (espesor, mm) resueltos contra la base dimensional.

- [ ] **Step 1: Escribir la prueba**

```python
    def test_od_y_espesor_no_miran_a_datos_ref(self, wb):
        ws = wb["Parche_PCC2_Art212"]
        for celda in ("D20", "D21"):
            f = str(ws[celda].value or "")
            assert "Datos_Ref" not in f, f"{celda} sigue leyendo de Datos_Ref: {f}"
            assert "DB_B36" in f, f"{celda} no lee de la base dimensional: {f}"

    def test_el_parche_del_tipo_de_cedula_desaparecio(self, wb):
        # El ""& existia porque la lista fija entregaba la cedula como numero y la
        # fila de encabezados de Datos_Ref la guardaba como texto. Con la cedula
        # saliendo de la propia base, la divergencia de tipo ya no puede ocurrir.
        ws = wb["Parche_PCC2_Art212"]
        assert '""&' not in str(ws["D21"].value or "")
```

- [ ] **Step 2: Correr para verla fallar**

Run:
```powershell
python -m pytest test_dashboard.py::TestCascadaDimensionalArt212 -q
```
Espera: FAIL — `D20`/`D21` siguen apuntando a `Datos_Ref`.

- [ ] **Step 3: Reescribir las dos fórmulas**

En `build_parche_art212`, sustituir los `calc("D20", ...)` y `calc("D21", ...)` actuales:

```python
    # OD y espesor se resuelven contra la fila (NPS, cedula) de la base dimensional.
    # La clave es NPS + designador de cedula, que es unica por construccion en la
    # base. Ya no hace falta el ""& de la version anterior: la cedula sale de la
    # propia base, asi que su tipo no puede divergir del de la columna que la indexa.
    fila_b36 = (f'MATCH($D$18&"|"&$D$19,'
                f'IF($D$17="B36.10M",{rango(b3610, "clave")},{rango(b3619, "clave")}),0)')
    com20 = ("Calculo: diametro exterior leido de la base dimensional (DB_B36_10 o "
             "DB_B36_19 segun la norma elegida en D17), por NPS y cedula.")
    lab(20, "Diámetro exterior", unidad="mm", ref="ASME B36 (auto)", com=com20)
    ws.cell(20, 2, "OD").font = Font(name=MONO, size=10, color=TINTA)
    calc("D20", f'=IFERROR(INDEX(IF($D$17="B36.10M",{rango(b3610, "od_mm")},'
                f'{rango(b3619, "od_mm")}),{fila_b36}),"")', com20)
    com21 = ("Calculo: espesor de pared leido de la misma fila de la base "
             "dimensional que el OD. Es el espesor NOMINAL de la norma; el "
             "remanente medido en campo va en la seccion de datos de inspeccion.")
    lab(21, "Espesor de pared", unidad="mm", ref="ASME B36 (auto)", com=com21)
    ws.cell(21, 2, "t").font = Font(name=MONO, size=10, color=TINTA)
    calc("D21", f'=IFERROR(INDEX(IF($D$17="B36.10M",{rango(b3610, "t_mm")},'
                f'{rango(b3619, "t_mm")}),{fila_b36}),"")', com21)
```

> `clave` es otra columna nueva de la base: `nps_impreso & "|" & designador`,
> precalculada en `build_db_b36` **como texto, no como fórmula** (regla: una base no
> lleva fórmulas, y `TestBasesDimensionalesB36.test_existen_y_no_llevan_formulas` lo
> comprueba). Añadirla a `COL_B36`, `CABECERA_B36`, `build_db_b36`, la auditoría §11 y
> las pruebas de la Tarea 6, igual que `nps_orden`.
>
> Hacer accesible `rango(...)` fuera de `_materializar_cascada_b36` — subirla a función
> de módulo (`_rango_b36(info, clave)`) y usarla desde los dos sitios, en vez de
> duplicar la expresión.

- [ ] **Step 4: Reconstruir, correr pruebas y verificar en Excel**

Run:
```powershell
python build_db_materiales.py --resources ..\..\..\resources --in ..\..\..\templates\maestro_con_macros.xlsm --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```
Espera: suites verdes y `verificar.py` en 0. **§7 (regresión del caso semilla) es el
gate real de esta tarea**: la línea 12"-CWS-46-032-B1 tiene que seguir dando
`t = 6,35 mm` y dictamen APTO leyendo ahora de `DB_B36_10`. Si §7 cambia de dictamen,
parar: la cascada resuelve otra fila.

- [ ] **Step 5: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/test_dashboard.py outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm
git commit -m "Art. 212: OD y espesor salen de la base dimensional"
```

---

## Tarea 9: La misma cascada en el Art. 206

> **EJECUTADA (2026-09-11), mirando la implementación REAL del Art. 212 (Tareas
> 7-8), no el borrador D17 de abajo —que el 212 abandonó—. Dos desvíos
> deliberados:**
> 1. **Norma en la fila 33, no 17.** `D17` es el encabezado de la sección 1, y
>    las filas 18-32 están todas ocupadas (Tipo de sleeve@22, UT@23, defecto
>    circ.@24, presiones, geometría…); correrlas reapuntaría D45/D46 (usan OD@20)
>    y D54/D64 (usan t@21). La norma ocupa la fila 33, libre al final de la
>    sección 1 — mismo criterio que la fila 22 del Art. 212.
> 2. **OD/espesor con `_choose_b36(idx,…)` + `MATCH(clave)`**, tal como quedó el
>    212 (no el `rango(...)` literal del borrador). El Art. 206 no tiene oracle
>    Rev0, así que no necesitó `DIVERGENCIAS_*`: su paridad la fija
>    `TestCascadaDimensionalArt206` contra DB_B36. `DEUDA_LISTA_FIJA` quedó vacía.
> Verificado: `test_build_db + test_dashboard + test_secii_tablas` → 232 pasan;
> `verificar.py` → 0 fallos (§5b guardia=0, §7 caso semilla=0, §11 B36=0).
> Commit `7d78af9`.

**Files:**
- Modify: `scripts/build_db_materiales.py` (`build_collar_art206`, filas 18-21)
- Modify: `scripts/test_dashboard.py`

**Interfaces:**
- Consume: `_materializar_cascada_b36`, `_rango_b36`, `COL_B36` (Tareas 6-8).
- Produce: en `Collar_PCC2_Art206`, la misma estructura `D17`/`D18`/`D19` + `D20`/`D21`.

- [x] **Step 1: Escribir la prueba**

```python
class TestCascadaDimensionalArt206:
    HOJA = "Collar_PCC2_Art206"

    def test_las_tres_celdas_tienen_validacion_de_rango(self, wb):
        ws = wb[self.HOJA]
        origen = {}
        for dv in ws.data_validations.dataValidation:
            for rng in dv.sqref.ranges:
                origen[str(rng)] = str(dv.formula1 or "")
        for celda in ("D17", "D18", "D19"):
            f1 = origen.get(celda, "")
            assert f1.startswith("="), f"{celda}: origen {f1!r}, se esperaba un rango"

    def test_od_y_espesor_leen_de_la_base(self, wb):
        ws = wb[self.HOJA]
        for celda in ("D20", "D21"):
            f = str(ws[celda].value or "")
            assert "Datos_Ref" not in f and "DB_B36" in f, f"{celda}: {f}"
```

- [x] **Step 2: Correr para verla fallar**

Run:
```powershell
python -m pytest test_dashboard.py::TestCascadaDimensionalArt206 -q
```
Espera: FAIL.

- [x] **Step 3: Aplicar el mismo tratamiento**

En `build_collar_art206`, el Art. 206 numera sus filas igual que el 212 en esta zona
(18 = NPS, 19 = cédula, 20 = OD, 21 = t), pero **hay que confirmarlo leyendo la función**
antes de tocarla: si el 206 usa otras filas, se adaptan los índices, no se mueve el
layout del motor.

Insertar la fila 17 de norma dimensional con su `inp` (mismo texto de nota que en el
212, adaptado al portador), llamar a `_materializar_cascada_b36(ws, b3610, b3619,
fila_norma=17, fila_nps=18, fila_ced=19)`, y sustituir `D20`/`D21` por las dos fórmulas
de la Tarea 8 con `rango(...)` → `_rango_b36(...)`.

Eliminar los dos `dv_list` literales de NPS y cédula, y la firma de
`build_collar_art206` pasa a recibir `b3610, b3619` igual que el 212 — actualizar su
llamada en `main()`.

- [x] **Step 4: Reconstruir, pruebas y verificar**

Run:
```powershell
python build_db_materiales.py --resources ..\..\..\resources --in ..\..\..\templates\maestro_con_macros.xlsm --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```
Espera: verde y 0.

- [x] **Step 5: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/test_dashboard.py outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm
git commit -m "Art. 206: NPS, cedula, OD y espesor salen de la base dimensional"
```

---

## Tarea 10: Retirar `Datos_Ref` y cerrar el guardia

> **EJECUTADA (2026-09-11).** Gates finales: 236 pruebas verdes; `verificar.py`
> en Excel real → total de fallos 0 (§5b guardia=0 sin deuda, §7 caso semilla
> APTO, §11 bases B36=0). Libro resultante: 72 hojas. Dos desvíos, ambos
> deliberados:
> 1. **`deprecate_datos_ref` → `retirar_datos_ref`.** El diseño solo pedía borrar
>    la hoja; la función previa la conservaba con un bloque «OBSOLETO». Se
>    reemplazó por el borrado defensivo (`if "Datos_Ref" in wb.sheetnames: del`).
> 2. **Dos referencias vivas extra a `Datos_Ref`**, no previstas por el Step 1,
>    detectadas por `test_ninguna_formula_la_menciona` y por Grep: (a) el valor de
>    celda `A131` del Art. 212 (venía de `nota_extra_cascada`) y (b) los
>    comentarios de `comentar_art212_base` para las filas 18-22, que las Tareas 7-8
>    dejaron citando `Datos_Ref` y describiendo mal D22 (hoy «Norma dimensional»).
>    Se corrigieron para reflejar `DB_B36` / la cascada; la fila 23 salió del dict
>    de comentarios (celda retirada). `DEUDA_LISTA_FIJA` ya estaba `{}` desde las
>    Tareas 7-9; se conserva vacía (la consultan el guardia §5 y el test).

**Files:**
- Modify: `scripts/build_db_materiales.py` (`HOJAS_HEREDADAS`, `DEUDA_LISTA_FIJA`, `main`, árbol de navegación)
- Modify: `scripts/test_dashboard.py`
- Modify: `CLAUDE.md`
- Modify: `outputs/plans/spec_entradas_de_motor_desde_base_de_datos.md` (estado final)

**Interfaces:**
- Consume: todo lo anterior.
- Produce: un libro sin `Datos_Ref`, con `HOJAS_HEREDADAS = ("Instrucciones",)`.

- [x] **Step 1: Comprobar que nadie más la referencia**

Run:
```powershell
python -c "import re; s=open('build_db_materiales.py',encoding='utf-8').read(); print('Datos_Ref:', s.count('Datos_Ref'))"
```
Cada aparición restante hay que mirarla: las que queden deben ser **texto histórico en
comentarios**, nunca una fórmula ni una referencia de hoja. Buscar también con Grep en
todo `scripts/` (incluidas las pruebas) y en `vba/`.

- [x] **Step 2: Escribir la prueba de cierre**

```python
class TestDatosRefRetirada:
    def test_la_hoja_no_existe(self, wb):
        assert "Datos_Ref" not in wb.sheetnames

    def test_ninguna_formula_la_menciona(self, wb):
        restos = []
        for ws in wb.worksheets:
            for fila in ws.iter_rows():
                for c in fila:
                    if isinstance(c.value, str) and "Datos_Ref" in c.value:
                        restos.append(f"{ws.title}!{c.coordinate}")
        assert restos == [], restos

    def test_solo_queda_una_hoja_heredada(self):
        import build_db_materiales as B
        assert B.HOJAS_HEREDADAS == ("Instrucciones",)

    def test_la_deuda_de_listas_fijas_esta_saldada(self):
        import build_db_materiales as B
        assert B.DEUDA_LISTA_FIJA == {}
```

- [x] **Step 3: Retirar la hoja**

- `HOJAS_HEREDADAS = ("Instrucciones",)`.
- Vaciar `DEUDA_LISTA_FIJA = {}` (las cuatro celdas ya leen de rango desde las Tareas 7-9).
- Quitar `Datos_Ref` del árbol de navegación (`ARBOL`) y de `order` en `main()` si aparece.
- Si el maestro sigue trayendo la hoja, borrarla explícitamente al construir, con el
  mismo patrón defensivo que usa `build_parche_art212` para su propia hoja:
  ```python
  if "Datos_Ref" in wb.sheetnames:   # el maestro Rev0 todavia la trae
      del wb["Datos_Ref"]
  ```

**El bloque obsoleto de las filas 40-47** (esfuerzos admisibles de seis materiales,
rotulado `[OBSOLETO — ver DB_B31_3 / DB_BPVC_IID]`) se va con la hoja: el dato vivo está
en las bases auditadas y el libro anterior queda en el historial de git. Decisión del
ingeniero del 2026-09-10.

- [x] **Step 4: Reconstruir y correr los tres gates**

Run:
```powershell
python build_db_materiales.py --resources ..\..\..\resources --in ..\..\..\templates\maestro_con_macros.xlsm --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
python verificar.py --resources ..\..\..\resources --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```
Espera: suites verdes; `verificar.py` en 0 con §5 mostrando
`Validaciones infractoras en hojas de motor: **0**` **sin** deuda declarada, y §7
manteniendo el dictamen APTO del caso semilla.

- [x] **Step 5: Documentar**

En `CLAUDE.md`: `HOJAS_HEREDADAS` queda en una sola hoja; aparecen `DB_B36_10` y
`DB_B36_19` en la estructura de bases y en el árbol del Dashboard; las entradas
dimensionales de los dos motores salen de esas bases. En el spec, sección «Estado
final» con el resultado de los gates.

- [x] **Step 6: Commit**

```powershell
git add scripts/build_db_materiales.py scripts/test_dashboard.py CLAUDE.md outputs/plans/spec_entradas_de_motor_desde_base_de_datos.md outputs/Motor_de_Calculo_ASME_PCC_Rev4.xlsm
git commit -m "Retira Datos_Ref: ninguna hoja de datos viene ya del maestro Rev0"
```

---

## Verificación — qué se puede aquí y qué no

- **Aquí:** las tres suites de pytest, `verificar.py` completo (esta máquina **sí** tiene
  Excel y `win32com`, comprobado el 2026-09-10, así que §6+ corre de verdad), el build
  sin abortos, y sobre todo §7 — la regresión del caso semilla es lo que dice si la
  cascada nueva resuelve la misma tubería que la vieja.
- **Solo el ingeniero:** abrir el libro y teclear un valor inválido en un desplegable
  para ver el rechazo (el bloqueo es comportamiento de la UI de Excel, no algo que
  openpyxl pueda comprobar), y la revisión visual de las dos hojas nuevas exportadas a
  PDF/PNG.

## Riesgos

- **Una semilla fuera de lista al activar el bloqueo** (Tarea 2). Es el motivo del Step 1
  de esa tarea: barrer antes, no después.
- **`n_ced` corto trunca la lista de cédulas en silencio** (Tarea 7). Por eso el plan
  obliga a medir el máximo real antes de fijar la constante.
- **El orden de la base y la cascada están acoplados** (regla 5): si alguien reordena
  `build_db_b36` sin mirar, las listas dependientes salen truncadas y los tests de
  contigüidad son lo único que lo atrapa. Está cubierto por
  `test_orden_contiguo_por_nps`.
- **El caso semilla es el canario.** Si §7 cambia de dictamen en las Tareas 8 o 9, la
  cascada está resolviendo otra fila: parar y diagnosticar, no ajustar el caso.
- **Retirar `Datos_Ref` tiene radio amplio**: fórmulas de dos motores, navegación,
  `retonar_heredadas`, `TEXTOS_HEREDADOS` y la sincronía Python/VBA.

## Self-review (hecho)

- **Cobertura del spec:** Fase 1 → Tareas 1-3; Fase 2 → Tareas 4-6; Fase 3 → Tareas 7-10.
  Los seis criterios de éxito del spec tienen tarea: (1) y (2) en la Tarea 3 y la 10;
  (3) en la 1; (4) en las 7-9; (5) en la 10; (6) en el paso de gates de cada tarea.
- **Sin placeholders:** las tres constantes que el plan no puede fijar de antemano
  (`n_nps`, `n_ced`, el literal de `Type A/Type B`) llevan el comando exacto para
  medirlas, no un «ajustar según convenga».
- **Consistencia de nombres:** `COL_B36` se declara en la Tarea 6 y lo consumen las 7, 8
  y 9; las columnas `nps_orden` y `clave` que introducen las Tareas 7 y 8 llevan aviso
  explícito de que hay que darlas de alta en los cinco sitios de la Tarea 6, porque son
  el punto donde este plan es más fácil de romper a medias.
