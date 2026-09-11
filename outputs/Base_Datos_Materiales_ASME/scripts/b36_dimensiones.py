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
    """'0.405 (10.29)' -> (0.405, 10.29). Sin par de numeros -> (None, None).

    B36.19M remite buena parte de sus espesores a una nota al pie -- 'Note (1)',
    'Notes (1), (2)' -- pegada tras el paréntesis (p.ej. '0.049 (1.24) [Note (1)]').
    La referencia se descarta para el valor numerico: no hay campo en esta base
    para el texto de la nota (a diferencia de Seccion II, aqui no se reconstruye
    una tabla de notas al pie), y el valor SIGUE siendo el que imprime el codigo
    -- omitirlo entero por la nota perdia 67 de 114 filas de B36.19M sin motivo.
    """
    m = re.match(r"^\s*([\d.]+)\s*\(([\d.]+)\)\s*(?:\[Notes?[^\]]*\])?\s*$", txt or "")
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
