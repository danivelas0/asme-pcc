# Prompt para NotebookLM — verificación de identidad UNS ↔ composición nominal ASME

## Fuentes que debe llevar el notebook

1. **ASME BPVC Sección II, Parte A** — `A1-2025.pdf` y `A2-2025.pdf`
   (requisitos de composición química por especificación SA y grado).
2. **ASME BPVC Sección II, Parte D (Métrica)** — `D Metric 2025 .pdf`
   (Tabla 1A con su columna *Nominal Composition*, y Tablas TM-1 y TE-1 con sus
   Notas al pie).

---

## Prompt (pegar tal cual)

Actúa como ingeniero de materiales revisando datos para un motor de cálculo
basado en ASME BPVC Sección II Parte D. Responde **solo** con lo que esté en las
fuentes cargadas y **cita** especificación, tabla y página en cada respuesta. Si
algo no está en las fuentes, dilo explícitamente en vez de completarlo.

**Contexto.** Las Tablas TM-1 (módulo de elasticidad E) y TE-1 (dilatación
térmica) de la Parte D no se indexan por especificación, sino por *composición
nominal*, y sus Notas al pie enumeran qué composiciones integran cada grupo.
Algunas filas de la Tabla 1A no imprimen la composición nominal, solo el número
UNS. Necesito confirmar la identidad de esos UNS.

**Tarea 1 — identidad UNS ↔ composición nominal.** Para cada bloque, confirma o
refuta que **todos** los UNS listados corresponden al material que la Parte D
designa con esa composición nominal. Apóyate en los requisitos de composición
química de la Parte A y en la columna *Nominal Composition* de la Tabla 1A.
Señala cualquier UNS del bloque que **no** encaje y explica por qué.

1. **«18Cr-8Ni»** — ¿lo son J92500, J92600, S30200, S30400, S30403, S30409?  (propuesto: Material Group G / Group 3)
2. **«1Cr-1∕2Mo»** — ¿lo son J11562, K11562, K11564, K11757, K12062?  (propuesto: Material Group C / Group 1)
3. **«C-1∕2Mo»** — ¿lo son J12521, J12522, J12524, K11522, K12821, K12822?  (propuesto: Material Group A / Group 1)
4. **«31∕2Ni»** — ¿lo son J31550, K31718, K31918, K32018, K32025?  (propuesto: Material Group B / Group 1)
5. **«3Cr-1Mo»** — ¿lo son K31545?  (propuesto: Material Group D / Group 1)
6. **«2Ni-1Cu»** — ¿lo son K22035?  (propuesto: Material Group B / —)
7. **«12Cr-Al»** — ¿lo son S40500?  (propuesto: Material Group F / —)
8. **«18Cr-11Ni»** — ¿lo son S30500?  (propuesto: — / Group 3)
9. **«17Cr»** — ¿lo son S43000?  (propuesto: Material Group F / —)
10. **«5Cr-1∕2Mo»** — ¿lo son J42045, K42544?  (propuesto: Material Group E / —)
11. **«15Cr»** — ¿lo son S42900?  (propuesto: Material Group F / —)
12. **«Mn-1∕2Mo»** — ¿lo son K12021, K12022?  (propuesto: Material Group A / Group 1)
13. **«11Cr-Ti»** — ¿lo son S40900?  (propuesto: Material Group F / —)
14. **«18Cr-8Ni-N»** — ¿lo son S30451, S30453?  (propuesto: Material Group G / Group 3)
15. **«13Cr»** — ¿lo son S41008?  (propuesto: Material Group F / —)
16. **«21∕2Ni»** — ¿lo son J22500?  (propuesto: Material Group B / Group 1)
17. **«13Cr-4Ni»** — ¿lo son J91540?  (propuesto: Material Group F / —)

**Tarea 2 — dilatación de los aceros 9Cr-1Mo-V.** La Tabla TE-1 tiene una columna
titulada *"Coefficients for 9Cr–1Mo Steels (Including Grades 9, 91, 911, and
92)"*. Confirma qué grados y qué UNS cubre esa columna, y si los materiales
9Cr-1Mo-V (Grado 91 / P91 / F91 / T91) quedan dentro. Cita la tabla.

**Tarea 3 — módulo E de los aceros 9Cr-1Mo-V.** La Tabla TM-1 **no** tiene fila
propia para 9Cr-1Mo-V. Su Nota (5) define el *Material Group E* como los aceros
de 5Cr a 9Cr e incluye la entrada *"9Cr–Mo, including variations thereof"*.
¿Hay en las fuentes algo que indique si el 9Cr-1Mo-V (Grado 91) queda
comprendido en esa redacción? Si no lo hay, dilo: es criterio del ingeniero y
no quiero una respuesta inventada.

**Tarea 4 — columnas nombradas de TE-1.** Además de los Grupos 1 a 4, TE-1 tiene
columnas propias (27Cr; 5Cr-1Mo; 9Cr-1Mo; 5Ni-¼Mo; 7% Ni; 8Ni y 9Ni; 12Cr,
12Cr-1Al, 13Cr y 13Cr-4Ni; 15Cr y 17Cr; fundición dúctil; 17Cr-4Ni-4Cu). Lista
esas columnas con el criterio exacto de pertenencia que imprime la tabla, tal
como está redactado.

**Formato de salida.** Una tabla: bloque | veredicto (CONFIRMADO / REFUTADO /
NO ESTÁ EN LAS FUENTES) | evidencia con cita | observaciones.
