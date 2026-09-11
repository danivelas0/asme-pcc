<!--
Plantilla de auto-revisión. Reutiliza el Protocolo de verificación del CLAUDE.md
del proyecto. Marcar solo lo que aplique al cambio; tachar (~~...~~) lo que no.
-->

## Qué cambia y por qué

<!-- Una o dos frases. Si hay un plan de outputs/plans/, enlazarlo. -->

## Checklist de ingeniería

- [ ] Entradas en SI; espesor remanente contra el mínimo admisible; margen de corrosión incluido.
- [ ] Artículo ASME PCC-2 (o código de construcción) citado para cada fórmula; notas del material leídas.
- [ ] `resources/` como única fuente (Regla nº 1): ningún valor normativo desde memoria.
- [ ] Gate local en verde: `build_db_materiales.py` + `pytest` + `verificar.py` (requiere Excel; el CI de la nube NO lo cubre).
- [ ] Plan de `outputs/plans/` actualizado (casillas) y movido a `registro/` si se cierra.
- [ ] `snake_case` en todo archivo nuevo o renombrado.

## Notas de revisión

<!-- Desvíos respecto al plan, decisiones de alcance, o lo que el revisor debe mirar con cuidado. -->
