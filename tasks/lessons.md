# Lecciones — Proyecto ASME PCC

## 2026-09-06 · Excel COM se cuelga en un diálogo modal que no puedo ver

**Corrección del ingeniero:** «it shows a warning. Compile error: Variable not defined».

Estaba sembrando el proyecto VBA con Excel COM. `wb.SaveAs(...)` se quedó bloqueado
cinco minutos y devolvió «The remote procedure call failed». Desde mi lado solo veía un
timeout; en la pantalla del ingeniero había un cuadro de diálogo de error de compilación
esperando un clic.

**Causa:** en `vba/mod_nav.vba` metí una función en medio del bloque de constantes, así
que dos `Private Const` quedaron declaradas **después** de la primera rutina. VBA exige
que toda declaración de módulo preceda a cualquier `Sub`/`Function`. No falla en la línea
de la declaración: falla con «Variable not defined» en cada uso. Y Excel compila al
guardar, así que el error aparece en `SaveAs`.

**Reglas para mí:**

1. **`DisplayAlerts = False` no suprime los diálogos del editor VBA.** Cualquier
   automatización de Excel puede quedar bloqueada en un modal invisible para mí. Si una
   llamada COM tarda mucho más de lo razonable, la hipótesis por defecto es un diálogo
   modal, no lentitud.
2. **Verificar en Python lo que Excel solo delata al compilar.** Añadido `lint_vba()` en
   `make_vba_seed.py`: comprueba que no haya `Const`/`Dim`/`Type`/`Declare`/`Enum` después
   de la primera rutina. Corre antes de tocar COM.
3. **Matar el proceso por PID, no confiar en `Quit`.** Excel sobrevive a `Quit` cuando se
   ha tocado el `VBProject`, bloquea el archivo y arruina la siguiente ejecución.
4. **Un `finally` que llama a `Close`/`Quit` sin protección tapa el error real.** Cada
   paso del cierre va en su propio `try`; si no, la excepción de limpieza sustituye a la
   que explica qué pasó.
5. **Después de editar la fuente, resembrar antes de probar.** Perdí una vuelta entera
   depurando un `.xlsm` construido con la semilla anterior.

## 2026-09-06 · Un ajuste de seguridad del usuario se quedó encendido

`make_vba_seed.py` activa `AccessVBOM` («Confiar en el acceso al modelo de objetos de
proyectos de VBA») y lo restaura en un `finally`. Aun así quedó en `1` dos veces.

**Causa:** Excel mantiene los ajustes del Centro de confianza en memoria y los vuelca al
registro *mientras se cierra*. Ese volcado asíncrono reescribía el `1` después de que yo
lo hubiera borrado.

**Reglas para mí:**

1. **Al tocar un ajuste de seguridad de la máquina, la restauración se verifica, no se
   asume.** `_restaurar_vbom()` escribe, relee, espera y reintenta.
2. **Comprobar el estado de la máquina al terminar**, no solo el entregable. Un
   protocolo en verde con un ajuste de seguridad encendido no es un trabajo terminado.
3. **Si mato un proceso a la fuerza, el `finally` no corre.** Restaurar a mano.

## 2026-09-06 · Verificar la premisa antes de escribirla en el plan

El plan afirmaba que el motor se entregaba con el caso semilla precargado, porque vi
`D22 = 'A106 Gr.B'` y `D23 = 'A516 Gr.70'` entre las celdas de entrada. Son campos de
texto libre descriptivos. La selección real del material va por la cascada
`D109..D114`, que se entrega **vacía**, así que el motor arrancaba en «SIN MATERIAL
SELECCIONADO» y la regresión no podía funcionar.

**Regla:** una celda con el nombre de un material no significa que el motor lo tenga
seleccionado. Antes de apoyar un plan en el estado de un libro, seguir la cadena de
fórmulas hasta la celda que de verdad manda.

## 2026-09-06 · Un texto no se guarda como fórmula

El maestro traía notas como `="texto de 265 caracteres…"`. Excel solo admite 255
caracteres en un literal de cadena dentro de una fórmula, así que al reguardar el archivo
lo partió en `_xlfn._LONGTEXT("trozo1","trozo2")`, que muestra `#NAME?` fuera de
Excel 365 y disparó la sección 5 del protocolo.

**Regla:** el builder normaliza estas celdas a texto plano
(`normalizar_textos_como_formula`). Un round-trip por Excel puede reescribir fórmulas
que openpyxl había leído sin quejarse: verificar el entregable después del round-trip,
no solo después del build.

## 2026-09-10 · `resources/` es la fuente para armar hojas — no se pide un folio externo

Al planear el rediseño del motor Art. 212, dos coeficientes no estaban como texto en el
JSON (ec. 2 del Art. 212 y ec. II-1 del App. 501-II). Propuse una Fase 0 que pedía al
ingeniero el **folio impreso** con `--pdf`. El ingeniero corrigió: `resources/` es la
fuente principal para construir las hojas; la información se extrae de sus JSON. Además,
mi propuesta contradecía la regla ya escrita en `CLAUDE.md` («el PDF solo se usa para
corregir y verificar los JSON… todo dato que entra a un motor sale de `resources/`»).

**Regla:** cuando falte un dato para una hoja, buscarlo en `resources/` — incluida la
**imagen de la ecuación**, que vive ahí (`bloque figure/equation` con `image: …png`): se
lee ese PNG con Read y de ahí se recupera la fórmula. Solo si no está ni como texto ni
como imagen en los dos espejos es un **vacío real**, que se declara y se repara
re-extrayendo, nunca se toma de memoria ni de `knowledge/claude.md` (derivado). Registrado
en `CLAUDE.md` bajo la Regla nº 1.

## 2026-09-11 · Los rulesets de GitHub NO son gratis en repo privado

**Contexto:** ejecutando `plan_flujo_github_ramas.md`. Su Fase 3 daba por hecho que
un *repository ruleset* era «gratis en repos privados; la protección clásica exige
plan de pago». Al crearlo, `POST`/`GET .../rulesets` devolvieron
`403 "Upgrade to GitHub Pro or make this repository public to enable this feature."`.

**Causa:** en un repo **privado** en plan gratuito, **tanto** los rulesets **como** la
protección de rama clásica exigen GitHub Pro. Solo son gratis en repos **públicos**.

**Reglas para mí:**

1. **Un plan puede traer una premisa de plataforma equivocada.** Antes de dar una fase
   por imposible o por hecha, verificar contra la API/CLI real, no contra lo que el
   plan afirma. Aquí la comprobación fue el propio `403`.
2. **Ante un bloqueo de plataforma (tier/permiso), no inventar un rodeo.** Se reportó el
   `403` tal cual, se dejó `main` sin protección server-side y se elevó la decisión
   (Pro / público / asumir riesgo) al ingeniero. No se hizo la prueba de force-push
   rechazado porque no había regla que lo rechazara: verificar lo que existe, no
   teatralizar una verificación vacía.

## Un mensaje de commit con comillas invertidas pierde palabras (2026-09-13)

`git commit -m "... \`Motor\` ..."` en bash ejecuta lo que va entre comillas
invertidas y deja el hueco: el mensaje del commit `3ea8125` perdio cinco
identificadores y nadie lo vio hasta releerlo. Un commit ya empujado a `main` no
se arregla —reescribir la historia compartida no se hace, y el hook de pre-push
lo impide—, asi que la informacion hubo que recuperarla en el documento.

**Regla:** todo mensaje de commit con codigo dentro va por heredoc citado
(`git commit -F - <<'EOF'`), nunca con `-m` y comillas dobles. El apostrofe del
`<<'EOF'` es lo que apaga la sustitucion.
