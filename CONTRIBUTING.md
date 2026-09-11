# Cómo se trabaja en este repositorio

Repo de un solo autor, privado. Se trabaja con **GitHub Flow**: `main` siempre
estable, una rama corta por unidad de trabajo, PR contra el propio repo, merge y
borrado de la rama. Todo en español, unidades SI por defecto (ver `CLAUDE.md`).

## Nomenclatura de ramas

| Prefijo | Para qué |
|---|---|
| `plan/<slug>` | Trabajo dirigido por un plan de `outputs/plans/` (varias Tareas). |
| `feat/<slug>` | Una funcionalidad nueva de una sola pieza. |
| `fix/<slug>` | Una corrección puntual. |
| `chore/<slug>` | Mantenimiento, infraestructura, andamiaje. |

El `<slug>` va en `kebab-case` y describe el trabajo, no la fecha ni el número.

## Ciclo

1. Ramificar desde `main` al día: `git switch main && git pull && git switch -c plan/<slug>`.
2. Commits con el estilo del repo: una línea, español, imperativo, describe el
   **qué** y a veces el **por qué** (ver `git log`). Cero relleno.
3. Publicar la rama: `git push -u origin plan/<slug>`.
4. Abrir el PR (`gh pr create`); rellenar la plantilla de auto-revisión.
5. Mergear y borrar la rama.

## Estilo de merge

- **Ramas de plan → merge commit `--no-ff`.** Deja un punto visible de «aquí
  aterrizó el plan X», que casa con los tags de revisión. Los commits por Tarea
  ya tienen mensajes con valor y se conservan (no squash).
- **One-offs triviales (`fix/`, `chore/`) → rebase** para no ensuciar la historia
  con merges de un commit.

## Relación con el ciclo `plans/ ↔ registro/`

Una rama `plan/<slug>` se abre al sacar el plan a `outputs/plans/`. El PR se
mergea cuando el plan se cierra y el plan se mueve a `outputs/plans/registro/`
(ver `outputs/plans/registro/LEEME.md`). Un plan marcado
**«NO EJECUTAR sin orden explícita del ingeniero»** no se corre por estar en
`outputs/plans/`: esa marca pide confirmación aparte.

## Tags por revisión del libro

Tag anotado (`rev4e`, `rev5`, …) cuando un merge **cambia la revisión del
entregable** (`outputs/Motor_de_Calculo_ASME_PCC_Rev*.xlsm`). El nombre mapea a
la «Rev» que ya usa el libro. Un merge que no toca el libro no lleva tag.

## Protección de `main`

**Server-side (GitHub):** no hay. Los *rulesets* y la protección de rama clásica
exigen **GitHub Pro** en un repo privado (`403 "Upgrade to GitHub Pro…"`). Si
algún día se pasa a Pro o el repo se hace público, se añade el ruleset con las
reglas `non_fast_forward` y `deletion` (ver `outputs/plans/registro/plan_flujo_github_ramas.md`).

**Local (esta máquina):** un hook `pre-push` versionado en `.githooks/pre-push`
rechaza, antes de que el push salga: **borrar `main`** y cualquier **push
non-fast-forward (force) a `main`**. El push normal (fast-forward) a `main` y
todo push a otras ramas quedan intactos. No se exige PR obligatorio: el PR es la
superficie de revisión y changelog por convención, no por obligación técnica.

El hook solo protege desde el clon que lo tenga activado. **En un clon nuevo,
activarlo una vez:**

```sh
git config core.hooksPath .githooks
```

Para saltarlo deliberadamente (bajo tu responsabilidad): `git push --no-verify`.

## El gate de calidad es local — el CI de la nube NO lo sustituye

El build del `.xlsm`, `make_vba_seed.py` y `verificar.py` dependen de **Excel +
COM de Windows**. Los runners de GitHub son Linux y **no** los pueden correr. El
CI en la nube (`.github/workflows/pruebas.yml`) solo ejerce las pruebas de Python
puro / openpyxl que leen `resources/` y el `.xlsm` ya commiteado.

**«CI verde» no es «entregable verificado».** Antes de mergear trabajo que toque
el motor, correr el gate completo en local:

```powershell
cd outputs\Base_Datos_Materiales_ASME\scripts
python build_db_materiales.py --resources ..\..\..\resources `
    --in ..\..\..\templates\maestro_con_macros.xlsm `
    --out ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
python -m pytest test_build_db.py test_dashboard.py test_secii_tablas.py -q
python verificar.py --resources ..\..\..\resources `
    --wb ..\..\Motor_de_Calculo_ASME_PCC_Rev4.xlsm
```
