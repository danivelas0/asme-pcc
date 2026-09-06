# -*- coding: utf-8 -*-
"""
proponer_decisiones — prepara, razonada y con evidencia, la revision de MAP_Grupo.

QUE HACE Y QUE NO
-----------------
Genera `propuesta_map_grupo.json`: una recomendacion por cada caso que el codigo
no resuelve solo, con la cadena de evidencia que la sostiene. **No firma nada.**
El campo `validado_por` queda vacio a proposito y el builder ignora toda entrada
sin firmar, de modo que este archivo no puede entrar al calculo por accidente.

El nombre del archivo tampoco es el que el builder lee (`decisiones_map_grupo.json`):
hace falta un acto deliberado del ingeniero —revisar, firmar y renombrar— para
que estas propuestas surtan efecto.

POR QUE HAY TRES CLASES DE CASO, Y SOLO DOS ADMITEN PROPUESTA
-------------------------------------------------------------
1. SIN MAPEO. La composicion no figura en ninguna Nota de TM-1 ni TE-1. Aqui no
   hay nada que proponer: II-D no publica modulo ni dilatacion para ese
   material y lo correcto es que siga bloqueado. Proponer un grupo seria
   sustituir al codigo por una conjetura, que es justo lo que la Rev. 3 elimino.
   Se comprobo ademas que ninguna de estas composiciones coincide con una nota
   con los elementos en otro orden: no son coincidencias perdidas.

2. REVISAR (composicion de otra tabla). La fila no imprime composicion, pero su
   UNS aparece con una sola composicion en otra tabla del propio libro, y esa
   composicion si figura en una Nota. La cadena es verificable de punta a punta
   y se propone aceptarla, mostrando cada eslabon.

3. REVISAR (regla textual). El codigo redacta un criterio de inclusion en vez de
   listar. Es interpretacion de codigo: se expone el criterio y lo que hay que
   ponderar, sin dar la respuesta por hecha.

USO
---
    python proponer_decisiones.py --wb ..\\..\\Motor_de_Calculo_ASME_PCC_Rev3.xlsm \\
        --resources ..\\..\\..\\resources
"""
from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

import build_db_materiales as B
from db_lib import Resources

R_HDR, R_DATA = 3, 4


def leer_map(ruta: Path, hoja="MAP_Grupo"):
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    filas = list(wb[hoja].iter_rows(values_only=True))
    wb.close()
    ix = {h: i for i, h in enumerate(filas[R_HDR - 1])}
    return [r for r in filas[R_DATA - 1:] if r and r[0]], ix


def agrupar(filas, ix):
    """Agrupa por la unidad en que se decide: composicion, o UNS si no la hay."""
    grupos = defaultdict(lambda: {"n": 0, "specs": Counter(), "motivo": "",
                                  "estado": "", "grp_e": None, "grp_te": None,
                                  "fte_e": "", "fte_te": ""})
    S = lambda v: str(v).strip() if v is not None else ""
    for r in filas:
        estado = S(r[ix["Estado"]])
        if estado.startswith("AUTO") or estado.startswith("VALIDADO"):
            continue
        comp = S(r[ix["Composicion nominal"]])
        uns = S(r[ix["UNS / Alloy"]])
        clave = ("composicion", comp) if comp else ("uns", uns.upper())
        g = grupos[(clave, estado)]
        g["n"] += 1
        g["estado"] = estado
        g["motivo"] = g["motivo"] or S(r[ix["Motivo"]])
        g["grp_e"] = g["grp_e"] or S(r[ix["Grupo E (TM)"]]) or None
        g["grp_te"] = g["grp_te"] or S(r[ix["Grupo dilatacion (TE)"]]) or None
        g["fte_e"] = g["fte_e"] or S(r[ix["Fuente E"]])
        g["fte_te"] = g["fte_te"] or S(r[ix["Fuente alfa"]])
        if S(r[ix["Spec. No."]]):
            g["specs"][S(r[ix["Spec. No."]])] += 1
    return grupos


def proponer(grupos):
    salida, sin_propuesta = [], []
    for ((tipo, valor), estado), g in sorted(
            grupos.items(), key=lambda kv: (kv[0][1], -kv[1]["n"])):
        specs = ", ".join(s for s, _ in g["specs"].most_common(4))

        if "otra tabla" in estado:
            # Cadena verificable: UNS -> composicion impresa por ASME para ese
            # mismo UNS -> Nota que lista esa composicion -> grupo.
            entrada = {
                ("uns" if tipo == "uns" else "composicion"): valor,
                "grupo_tm": g["grp_e"] or "",
                "grupo_te": g["grp_te"] or "",
                "justificacion": (
                    f"El UNS designa una composicion asignada por SAE/ASTM. "
                    f"{g['motivo']} Grupo propuesto por esa via: "
                    f"{g['grp_e'] or '-'} (E) / {g['grp_te'] or '-'} (dilatacion)."),
                "validado_por": "",
                "fecha": "",
                "_recomendacion": "ACEPTAR si confirma que el UNS es ese material",
                "_confianza": "alta — cadena verificable en tablas del propio ASME",
                "_evidencia": {"fuente_E": g["fte_e"], "fuente_alfa": g["fte_te"]},
                "_filas_afectadas": g["n"],
                "_especificaciones": specs,
            }
            salida.append(entrada)

        elif "regla textual" in estado:
            entrada = {
                ("uns" if tipo == "uns" else "composicion"): valor,
                "grupo_tm": "",
                "grupo_te": "",
                "justificacion": "",
                "validado_por": "",
                "fecha": "",
                "_recomendacion": ("DECIDIR USTED. El codigo no lista este material: "
                                   "redacta un criterio de inclusion."),
                "_confianza": "no aplica — es interpretacion de codigo",
                "_a_ponderar": g["motivo"],
                "_filas_afectadas": g["n"],
                "_especificaciones": specs,
            }
            salida.append(entrada)

        else:                                    # SIN MAPEO
            sin_propuesta.append({
                ("uns" if tipo == "uns" else "composicion"): valor,
                "_filas_afectadas": g["n"],
                "_especificaciones": specs,
                "_por_que_no_hay_propuesta": g["motivo"],
            })
    return salida, sin_propuesta


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--wb", required=True, type=Path)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--out", type=Path, default=None)
    a = ap.parse_args(argv)

    filas, ix = leer_map(a.wb)
    salida, sin_propuesta = proponer(agrupar(filas, ix))

    doc = {
        "_que_es_esto": (
            "PROPUESTA, no decision. Preparada por script a partir de las tablas "
            "del codigo; `validado_por` esta vacio a proposito y el builder "
            "ignora toda entrada sin firmar."),
        "_como_usarla": [
            "1. Revise cada entrada con su `_recomendacion` y su `_evidencia`.",
            "2. Si la acepta, escriba su nombre en `validado_por` y la fecha.",
            "3. Si no, corrija `grupo_tm` / `grupo_te` o deje la entrada sin firmar.",
            "4. Guarde el archivo como `decisiones_map_grupo.json` (nombre distinto,",
            "   a proposito: renombrarlo es el acto deliberado que las activa).",
            "5. Reconstruya el libro y ejecute verificar.py.",
        ],
        "_sin_propuesta": (
            f"{len(sin_propuesta)} casos no llevan propuesta: su composicion no "
            f"figura en ninguna Nota de TM-1 ni TE-1, de modo que II-D no publica "
            f"modulo ni dilatacion para ese material. Lo correcto es que sigan "
            f"bloqueados. Se listan al final solo para que conste que se miraron."),
        "decisiones": salida,
        "_bloqueados_sin_propuesta": sin_propuesta,
    }
    destino = a.out or (a.wb.parent / "Base_Datos_Materiales_ASME" /
                        "propuesta_map_grupo.json")
    destino.write_text(json.dumps(doc, ensure_ascii=False, indent=2),
                       encoding="utf-8")
    print(f"{len(salida)} propuestas · {len(sin_propuesta)} bloqueados sin propuesta")
    print(f"escrito {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
