# Nota de versión — Motor de Cálculo ASME PCC, Rev. 2
## Base de datos de materiales (PLAN-DB-MAT-001)

**Fecha:** 2026-09-06 · **Fuente única de verdad:** `resources/` · **Compatible con Excel y Google Sheets**

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

`MAP_Grupo`: 1 292 filas **AUTO** (UNS exacto, auditables 1:1), **1 518 PROPUESTA — requieren que usted confirme la familia** antes de usar `E` o dilatación, 644 sin mapeo.

**Enlace SI ↔ US:** la selección se hace sobre la edición métrica. El B31.3 enlaza 1 065 de 1 288 materiales con su homólogo A-1C; la II-D, el 100 % de la 1A. Los no enlazados muestran «sin equivalente en la edición US» en lugar de un valor equivocado.

## Reproducir

```bash
python scripts/build_db_materiales.py --resources <ruta>/resources \
    --in Motor_de_Calculo_ASME_PCC.xlsx --out Motor_de_Calculo_ASME_PCC_Rev2.xlsx
python -m pytest scripts/test_build_db.py -q
python scripts/verificar.py
```

---

*Herramienta de ingeniería de referencia. Verificar entradas y resultados antes de emitir para construcción.*
