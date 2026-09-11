# Plan: Flujo de trabajo GitHub y árbol de ramas (GitHub Flow)

> **NO EJECUTAR sin orden explícita del ingeniero.** Este plan cambia el flujo de
> trabajo del repositorio y toca la configuración remota de GitHub (ruleset,
> Actions). Confirmar antes de correr cualquier paso.

---

## Contexto

Hoy el repo `danivelas0/asme-pcc` (privado, un solo autor) trabaja **todo sobre
`main`**: historia lineal, cero ramas de feature, cero tags, sin `.github/`, sin CI,
y con 8 commits locales sin subir a `origin`. Daniel quiere dejar de anclar todo a
`main` y adoptar buenas prácticas de GitHub calibradas para un repo de un autor.

Decisiones tomadas (2026-09-11):
- **Alcance:** solo el **árbol de ramas** de git. No se reorganizan directorios ni
  se reescribe la historia existente (no destructivo).
- **Modelo:** **GitHub Flow** — `main` siempre estable; una rama corta por unidad de
  trabajo; PR contra sí mismo; merge y borrado de la rama.
- **Protección:** **anti-accidente** — bloquear force-push y borrado de `main`, pero
  permitir push directo (no se exige PR obligatorio).
- **CI:** **parcial** — en la nube corre solo lo que no necesita Excel; el build y
  `verificar.py` quedan como gate manual local documentado.

Resultado esperado: cada plan/tarea vive en su propia rama, `main` protegido contra
catástrofes, PRs como superficie de revisión y changelog, tags por revisión del
libro, y una comprobación automática de las pruebas que sí corren sin Excel.

---

## Restricción técnica que condiciona el CI

`verificar.py`, `make_vba_seed.py` y el build (`build_db_materiales.py` sobre el
maestro sembrado) dependen de **Excel + COM de Windows**. Los runners de GitHub son
Linux: **no** pueden sembrar VBA, construir el `.xlsm` ni recalcular fórmulas. El CI
en la nube se limita a las pruebas de **Python puro / openpyxl** que leen `resources/`
y el `.xlsm` ya commiteado. El gate completo (build + `verificar.py`) sigue siendo
**manual y local**, y así se documenta.

---

## Fase 0 — Sincronizar antes de tocar nada

- Subir los 8 commits locales: `git push origin main`. La base remota tiene que estar
  al día **antes** de crear el ruleset (si no, un push posterior podría chocar con la
  regla recién creada). Verificar `git status -sb` → `main...origin/main` sin ahead.

---

## Fase 1 — Convención de ramas y merge (documentar, no ejecutar código)

Crear `CONTRIBUTING.md` (raíz, en español) — GitHub lo muestra en la UI de PRs — con:

- **Nomenclatura de ramas:**
  - `plan/<slug>` — trabajo dirigido por un plan de `outputs/plans/` (varias Tareas).
  - `fix/<slug>`, `chore/<slug>`, `feat/<slug>` — cambios de una sola pieza.
- **Ciclo GitHub Flow:** ramificar desde `main` → commits (mismo estilo de mensaje
  actual) → `push -u` → abrir PR → merge → borrar rama.
- **Estilo de merge:** **merge commit `--no-ff`** para ramas de plan (deja un punto
  visible de "aquí aterrizó el plan X", que casa con los tags); rebase para one-offs
  triviales. No squash por defecto: los commits por Tarea ya tienen mensajes con
  valor y se quieren conservar.
- **Tags por revisión del libro:** tag anotado `rev4x`, `rev5`, … cuando un merge
  cambia la revisión del entregable. Mapear los tags a las "Rev" que ya usa el libro.
- **Relación con `plans/ ↔ registro/`:** una rama de plan se abre al sacar el plan a
  `outputs/plans/`; el PR se mergea cuando el plan se cierra y se mueve a `registro/`.

**Archivos:** `CONTRIBUTING.md` (nuevo).

---

## Fase 2 — Plantilla de PR (auto-revisión)

Crear `.github/PULL_REQUEST_TEMPLATE.md` (español) con una checklist que reutiliza el
**Protocolo de verificación** ya definido en el CLAUDE.md del proyecto:

- [ ] Entradas en SI; espesor remanente contra mínimo admisible; margen de corrosión.
- [ ] Artículo ASME PCC-2 citado para cada fórmula; notas del material leídas.
- [ ] `resources/` como única fuente (Regla nº 1): ningún valor desde memoria.
- [ ] Gate local corrido: `build_db_materiales.py` + `pytest` + `verificar.py` en verde.
- [ ] Plan de `outputs/plans/` actualizado (casillas) si aplica.

**Archivos:** `.github/PULL_REQUEST_TEMPLATE.md` (nuevo).

---

## Fase 3 — Ruleset anti-accidente sobre `main`

Crear un **repository ruleset** (gratis en repos privados; la protección de rama
"clásica" exige plan de pago) que:

- **Bloquea el borrado** de `main`.
- **Bloquea el push no-fast-forward** (force-push).
- **Permite** el push directo normal (no se exige PR ni revisores — decisión
  "anti-accidente").

Vía preferida `gh` (confirmar `gh auth status` primero):
```bash
gh api -X POST repos/danivelas0/asme-pcc/rulesets \
  -f name='proteger-main' -f target='branch' -f enforcement='active' \
  -f 'conditions[ref_name][include][]=refs/heads/main' \
  -f 'rules[][type]=non_fast_forward' \
  -f 'rules[][type]=deletion'
```
Alternativa: GitHub → Settings → Rules → Rulesets → New branch ruleset (UI), mismos
dos toggles. Verificar con `gh api repos/danivelas0/asme-pcc/rulesets`.

> Nota: este ruleset **no** rompe el push directo actual; solo impide force-push y
> borrado. Si más adelante se quiere endurecer a "PR obligatorio", se añade la regla
> `pull_request` al mismo ruleset.

---

## Fase 4 — CI parcial en GitHub Actions

Crear `.github/workflows/pruebas.yml`:

- **Disparo:** `push` (todas las ramas) y `pull_request` contra `main`.
- **Runner:** `ubuntu-latest`.
- **Pasos:** checkout → `actions/setup-python` (fijar la versión que usa Daniel en
  local) → `pip install openpyxl pytest` (+ dependencias que revelen las pruebas) →
  correr **solo el subconjunto sin Excel** desde `outputs/Base_Datos_Materiales_ASME/scripts`.
- **Qué corre:** las pruebas que leen `resources/` (JSON rastreado) y el `.xlsm` ya
  commiteado vía openpyxl — candidatas: `test_secii_tablas.py`, la parte de
  `test_build_db.py` que no invoca COM, y las lecturas de `test_dashboard.py` sobre el
  libro construido.
- **Qué NO corre (documentado en el YAML):** `make_vba_seed.py`, el build, y
  `verificar.py` — todos requieren Excel/COM.

> **Se espera iterar en el primer run.** La primera corrida en la nube dirá qué
> pruebas fallan por falta de Excel o por un `content.json` gitignored (los 269 MB de
> II A/B/C están excluidos por `.gitignore`). El subconjunto exacto se fija tras ese
> primer run — narrando por archivo o con un marcador `requires_excel`; **nunca
> relajando una prueba**, solo excluyendo la que dependa de Excel.

**Archivos:** `.github/workflows/pruebas.yml` (nuevo).

---

## Fase 5 — Tag retroactivo del estado actual (no destructivo)

- Tag anotado sobre el `main` ya sincronizado marcando la Rev4 vigente:
  `git tag -a rev4x -m "Motor de Cálculo ASME PCC Rev4 (estado al adoptar GitHub Flow)"`
  y `git push origin rev4x`. No reescribe nada; solo pone una etiqueta sobre HEAD.
  (Confirmar con Daniel el sufijo exacto de la Rev — 4d/4e — antes de fijar el nombre.)

---

## Fase 6 — Estrenar el flujo con la Tarea 7

Primer uso real del flujo nuevo, que además retoma el trabajo pendiente:
- Crear la rama `plan/entradas-motor-desde-bd` desde `main`.
- Ejecutar en esa rama la **Tarea 7** (y siguientes 8–10) del plan de entradas
  (`plan_entradas_de_motor_desde_base_de_datos.md`).
- Al cerrar el plan: PR con la plantilla, merge `--no-ff`, borrar rama, mover el plan
  a `registro/`, y tag si cambia la Rev del libro.

> Este plan **no** ejecuta la Tarea 7; solo deja listo el carril. La Tarea 7 arranca
> cuando Daniel lo ordene.

---

## Fuera de alcance (no en este plan; ofrecer aparte si interesa)

- `LICENSE` (decisión legal de Daniel; ojo con datos derivados de ASME ya gitignored).
- Plantillas de issues, Dependabot, `SECURITY.md`.
- Reorganización de directorios y reescritura/limpieza de la historia de `main`.
- Endurecer el ruleset a PR obligatorio y CI como status check requerido.

---

## Verificación

1. **Fase 0:** `git status -sb` muestra `main...origin/main` sin `ahead`.
2. **Fase 3:** `gh api repos/danivelas0/asme-pcc/rulesets` lista `proteger-main`
   activo; un `git push --force-with-lease origin main` de prueba es **rechazado**.
3. **Fase 4:** abrir una rama de prueba y un PR dispara el workflow; el run queda
   verde con el subconjunto sin Excel (tras la iteración del primer run).
4. **Fase 5:** `git tag` lista `rev4x`; visible en GitHub → Releases/Tags.
5. **Integridad local intacta:** el gate completo local
   (`build_db_materiales.py` → `pytest` → `verificar.py`) sigue corriendo en verde
   igual que antes; nada de este plan lo altera.

---

## Riesgos

- **Ruleset creado antes de sincronizar** → un push legítimo podría chocar. Mitiga la
  Fase 0 (push primero, ruleset después).
- **CI rojo permanente** si se intenta correr pruebas que necesitan Excel. Mitiga la
  iteración declarada en la Fase 4: excluir lo que dependa de COM, nunca debilitar la
  prueba.
- **Falso sentido de cobertura:** el CI en la nube NO valida el `.xlsm` recalculado.
  El YAML y el `CONTRIBUTING.md` lo dicen explícitamente para que nadie confunda
  "CI verde" con "entregable verificado".
