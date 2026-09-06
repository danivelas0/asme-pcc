# Proyecto ASME PCC — Motor de Cálculo e Ingeniería de Reparación

Ingeniería de reparación de equipos a presión y tubería según la familia **ASME PCC**
(PCC-1 / PCC-2 / PCC-3), con los códigos de construcción **ASME B31.3-2024** y
**ASME BPVC Sección VIII-1 + II-D 2025** como referencia.

Responde siempre en **español**. Unidades **SI por defecto** (MPa, mm, °C).
Tono profesional y directo, sin relleno.

---

## Regla nº 1 — `resources/` es la única fuente de verdad

`resources/` contiene los códigos y normas extraídos a JSON. **Ningún valor normativo
—esfuerzos admisibles, factores, fórmulas, límites— puede salir de la memoria del
modelo.** Siempre se lee el archivo fuente antes de citar o calcular.

Si un dato no existe en `resources/`, **no se inventa ni se aproxima**: se declara
el vacío explícitamente y se pregunta cómo proceder.

Todo coeficiente o tabla que alimente un motor de cálculo debe **trazar** al archivo
concreto de `resources/` del que procede, para permitir auditoría posterior.

## Estructura

```
knowledge/     Instrucciones de cálculo ASME PCC-2 en SI. Leer antes de cualquier tarea.
resources/     Códigos y normas (JSON). Fuente única de verdad.
               ├─ ASME B31/ASME B31.3/APPEX/   Apéndices A, B y C
               ├─ ASME PCC/pcc_2/              Artículos de PCC-2
               ├─ bpvc_ii_d_metric_2025/       II-D métrica (MPa, °C)
               └─ bpvc_ii_d_customary_2025/    II-D U.S. Customary (ksi, °F)
outputs/       Entregables. NO leer sin que se te indique un archivo concreto.
templates/     Plantillas de formato. NO leer sin orden.
Motor_de_Calculo_ASME_PCC.xlsx   Libro maestro (Rev. 0, intacto)
```

**Nunca leas `outputs/` ni `templates/`** salvo que se te señale un archivo.
**Guarda todo entregable en `outputs/` dentro de una subcarpeta** con nombre de proyecto.

Ante una duda de alcance, pregunta antes de producir.

---

## Motor de cálculo — estado actual

Entregable vigente: `outputs/Base_Datos_Materiales_ASME/Motor_de_Calculo_ASME_PCC_Rev2.xlsx`
(37 hojas). Se **genera por script**, nunca se edita a mano.

### Reconstruir

```bash
cd outputs/Base_Datos_Materiales_ASME/scripts
python build_db_materiales.py --resources ../../../resources \
    --in ../../../Motor_de_Calculo_ASME_PCC.xlsx \
    --out ../Motor_de_Calculo_ASME_PCC_Rev2.xlsx
python -m pytest test_build_db.py -q      # pruebas unitarias
python verificar.py                       # protocolo de aceptación completo
```

`verificar.py` devuelve 0 solo si todo pasa. Audita **fila a fila** cada valor
tabulado contra el JSON del código (271 276 valores), la contigüidad de la cascada,
la ausencia de fórmulas de matriz dinámica y la interpolación recalculada en hoja.
**Ejecútalo siempre después de tocar el builder.**

### Reglas de diseño del libro — no romper

1. **Cero funciones de matriz dinámica.** Nada de `FILTER`, `SORT`, `UNIQUE`,
   `XLOOKUP`, `VSTACK`. Solo `INDEX`, `MATCH`, `OFFSET`, `COUNTIF`. El libro debe
   abrirse igual en Excel de escritorio y en Google Sheets.
2. **Validación de datos: solo rango literal o lista de ítems.** Google Sheets
   descarta cualquier fórmula (`OFFSET`, `INDIRECT`) como origen de validación. Las
   listas dependientes se materializan en celdas de columnas ocultas y la validación
   apunta a ese rango.
3. **Banda compacta.** Las tablas ASME dejan huecos interiores (la A-1 no imprime
   125 °C para A106 Gr.B). Junto a la banda impresa —que se conserva para auditar—
   cada base lleva una banda con solo los puntos existentes, sin huecos, y la
   consulta trabaja sobre ella.
4. **No extrapolar.** Por encima de la última temperatura tabulada o de la Temp. máx.
   del material, dictamen "FUERA DE RANGO" y resultado bloqueado. Lo prohíbe el código.
5. **Cascada contigua.** Cada base se ordena por familia → composición → forma →
   especificación → grado, para que cada nivel sea un bloque contiguo.
6. **Unidades visibles** en toda variable, indicadores y ejes de gráficas.
7. **`material_id` único**, desambiguado en este orden: notas del código → Rm → Re →
   número de línea impreso. El código repite spec+grado con admisibles distintos.
8. La **familia de material** es una agrupación de navegación derivada del UNS y la
   composición impresa. **No es dato normativo** y no entra en ningún cálculo.
9. Los valores se cargan **tal como están impresos**; SI y US son extracciones
   independientes, nunca conversiones.

### Pendiente

`MAP_Grupo`: 1 292 filas AUTO (UNS exacto, auditables), **1 518 PROPUESTA que
requieren validación del ingeniero** antes de usar E o dilatación, 644 sin mapeo.

---

## Protocolo de verificación antes de entregar un cálculo

1. ¿Todas las entradas en SI? (psi → MPa, in → mm, °F → °C)
2. ¿Verificado el espesor remanente contra el mínimo admisible?
3. ¿Incluido el margen de corrosión para la vida remanente?
4. ¿Citado el artículo de ASME PCC-2 que soporta cada fórmula aplicada?
5. ¿Leídas las notas del material? (`Notas_Codigo`; restringen soldadura, PWHT y servicio)

## Código Python

Modular, con pruebas, documentado. Son herramientas de ingeniería críticas: cada
decisión no obvia va comentada explicando **por qué**, no qué hace.
