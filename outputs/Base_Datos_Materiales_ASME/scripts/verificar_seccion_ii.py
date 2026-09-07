# -*- coding: utf-8 -*-
"""
verificar_seccion_ii - audita las partes A, B y C de ASME BPVC Seccion II
cargadas en `resources/ASME_BPVC/Sec_II/` contra sus PDF de origen.

QUE COMPRUEBA, Y POR QUE
------------------------
Estas cuatro partes no son tablas de valores como la II-D: son el texto integro
de 379 especificaciones de material troceado por especificacion. Lo que puede
salir mal no es un numero, es el TROCEADO: que una especificacion empiece donde
no debe, que se pierdan paginas entre dos, o que el indice prometa un archivo
que no existe. Eso es lo que se audita:

  1. Paginas          el PDF tiene las que declara meta.json / index.json.
  2. Archivos         cada entrada del indice apunta a un archivo que existe.
  3. Anclaje          en la primera pagina PDF de cada entrada aparece
                      realmente la designacion de esa especificacion. Es la
                      comprobacion fuerte: demuestra que el corte esta alineado
                      con el documento real, no con un offset heredado.
  4. Folio impreso    el numero de pagina impreso en esa pagina coincide con el
                      `printed_pages` declarado.
  5. Cobertura        los rangos no se solapan y lo que queda fuera es solo el
                      material preliminar y final que el indice declara excluido.
  6. Figuras          las PNG en disco son las que declara `totals.figures`.

Los PDF NO estan en el repositorio (copyright de ASME): se pasan por ruta.

USO
---
    python verificar_seccion_ii.py --resources ..\\..\\..\\resources ^
        --pdfs "<...>\\BPVC\\SECCION II"

Devuelve 0 solo si no hay fallos.
"""

import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path

SEC_II = Path("ASME_BPVC") / "Sec_II"

# carpeta en resources  ->  nombre del PDF
PARTES = {
    "bpvc_ii_a_1": "A1-2025.pdf",
    "bpvc_ii_a_2": "A2-2025.pdf",
    "bpvc_ii_b": "B-2025.pdf",
    "bpvc_ii_c": "C-2025.pdf",
}


def plano(s):
    """Compara ignorando lo que el PDF pierde: espacios, guiones y acentos.

    El texto extraido llega sin espacios entre palabras ("SpecificationforAluminum")
    y con guiones tipograficos, asi que la unica comparacion fiable es sobre la
    cadena sin separadores.
    """
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    for ch in "‐‑‒–—−⁄∕":
        s = s.replace(ch, "-")
    return re.sub(r"[^A-Za-z0-9./-]+", "", s).upper()


def designaciones(entrada):
    """Las formas en que la primera pagina puede nombrar la especificacion."""
    fuera = []
    for k in ("specification", "astm_designation", "equivalent_designation"):
        v = entrada.get(k)
        if isinstance(v, str) and v.strip():
            fuera.append(v.strip())
    t = entrada.get("title") or ""
    if " - " in t:
        fuera.append(t.split(" - ", 1)[0].strip())
    # "SFA-5.01M/SFA-5.01" tambien vale por cualquiera de sus mitades
    extra = []
    for d in fuera:
        extra += [x for x in re.split(r"[/,]", d) if len(x.strip()) > 3]
    return [plano(x) for x in fuera + extra if plano(x)]


def folio_impreso(texto):
    """Numeros sueltos al principio o al final de la pagina: el folio va ahi."""
    lineas = [l.strip() for l in (texto or "").splitlines() if l.strip()]
    fuera = set()
    for l in lineas[:2] + lineas[-2:]:
        for m in re.findall(r"\b(\d{1,4})\b", l):
            fuera.add(int(m))
    return fuera


def prefijo_archivo(entrada):
    """`SA-6/SA-6M` -> `sa_6_sa_6m`, que es como empieza el archivo en disco.

    No todas las entradas son especificaciones: cada parte incluye ademas sus
    apendices, que no llevan designacion y se guardan en `appendices/`. Para
    esos el prefijo se saca del titulo.
    """
    d = entrada.get("specification") or entrada.get("title") or ""
    d = d.split(" - ", 1)[0] if entrada.get("specification") else d
    return re.sub(r"[^a-z0-9]+", "_", d.lower()).strip("_")


def auditar(parte, carpeta, pdf_path, log):
    import pdfplumber

    idx = json.loads((carpeta / "index.json").read_text(encoding="utf-8"))
    entradas = idx.get("entries") or []
    totales = idx.get("totals") or {}
    fallos = 0

    with pdfplumber.open(pdf_path) as pdf:
        npag = len(pdf.pages)

        # --- 1. paginas -----------------------------------------------------
        # OJO: `totals.pdf_pages` NO es la longitud del documento, sino las
        # paginas que cubren las especificaciones. El resto es material
        # preliminar y final, que el indice declara excluido a proposito.
        cubierto_decl = totales.get("pdf_pages")
        cubierto_real = sum(e["pdf_pages"][1] - e["pdf_pages"][0] + 1
                            for e in entradas
                            if len(e.get("pdf_pages") or []) == 2)
        ok = cubierto_decl is None or cubierto_decl == cubierto_real
        log("| %s | paginas declaradas frente a las que suman las entradas | "
            "%s / %d (el PDF tiene %d) | %s |"
            % (parte, cubierto_decl, cubierto_real, npag, "OK" if ok else "REVISAR"))
        fallos += 0 if ok else 1

        # --- 2. archivos ----------------------------------------------------
        en_disco = []
        for sub in ("specifications", "appendices"):
            d = carpeta / sub
            if d.is_dir():
                en_disco += sorted(p.name for p in d.glob("*.json"))
        # El indice promete una carpeta por especificacion; en disco estan
        # aplanadas en specifications/. Se comprueba que cada entrada tiene su
        # archivo, resolviendo por el prefijo de la designacion.
        sin_archivo, usados = [], set()
        stems = {n: re.sub(r"(_source)?\.json$", "", n) for n in en_disco}
        # De mas especifico a menos: si no, `sa_31` se queda con el archivo de
        # `sa_311_sa_311m` y deja a SA-311/SA-311M sin emparejar.
        for e in sorted(entradas, key=lambda x: -len(prefijo_archivo(x))):
            pre = prefijo_archivo(e)
            cand = None
            for n, stem in stems.items():
                if n in usados or not pre:
                    continue
                # Frontera de guion bajo: `sa_31` no puede casar con `sa_311...`.
                # El sentido inverso cubre el nombre de archivo truncado a unos
                # 80 caracteres, que solo pasa con titulos largos de apendice.
                if stem == pre or stem.startswith(pre + "_") or \
                        (len(stem) >= 40 and pre.startswith(stem)):
                    cand = n
                    break
            if cand is None:
                sin_archivo.append(e.get("specification")
                                   or (e.get("title") or "?")[:50])
            else:
                usados.add(cand)
        huerfanos = [n for n in en_disco if n not in usados]
        log("| %s | entradas con su archivo en disco | %d de %d | %s |"
            % (parte, len(entradas) - len(sin_archivo), len(entradas),
               "OK" if not sin_archivo else "REVISAR"))
        fallos += 0 if not sin_archivo else 1
        log("| %s | archivos sin entrada en el indice | %d de %d | %s |"
            % (parte, len(huerfanos), len(en_disco),
               "OK" if not huerfanos else "revisar"))
        ruta_idx = entradas[0].get("file") if entradas else None
        if ruta_idx and not (carpeta / ruta_idx).is_file():
            log("|   | ruta del indice | `file` apunta a `%s`, que no existe: en "
                "disco estan aplanados en `specifications/` | revisar |" % ruta_idx)
        for s in sin_archivo[:6]:
            log("|   | sin archivo | %s | REVISAR |" % s)

        # --- 3 y 4. anclaje y folio impreso ---------------------------------
        sin_ancla, folio_mal, sin_rango = [], [], 0
        for e in entradas:
            rango = e.get("pdf_pages") or []
            if len(rango) != 2:
                sin_rango += 1
                continue
            # `pdf_pages` es 0-BASED e inclusivo: [56, 119] son las paginas 57
            # a 120 del PDF. Verificado en A-1: la 56 esta en blanco y SA-6/SA-6M
            # empieza en la 57, con el folio impreso 3 que declara la entrada.
            p0 = rango[0]
            if not (0 <= p0 < npag):
                sin_ancla.append((e.get("specification"), p0, "fuera del PDF"))
                continue
            texto = pdf.pages[p0].extract_text() or ""
            t = plano(texto)
            if not any(d in t for d in designaciones(e)):
                sin_ancla.append((e.get("specification") or e.get("title", "")[:40],
                                  p0, "designacion ausente"))
            pr = e.get("printed_pages") or []
            if len(pr) == 2 and pr[0] not in folio_impreso(texto):
                folio_mal.append((e.get("specification"), p0, pr[0]))

        log("| %s | la designacion aparece en la 1a pagina declarada | %d de %d | %s |"
            % (parte, len(entradas) - len(sin_ancla) - sin_rango, len(entradas),
               "OK" if not sin_ancla else "REVISAR"))
        fallos += 0 if not sin_ancla else 1
        log("| %s | el folio impreso coincide | %d de %d | %s |"
            % (parte, len(entradas) - len(folio_mal) - sin_rango, len(entradas),
               "OK" if not folio_mal else "REVISAR"))
        if folio_mal:
            fallos += 1

        # --- 5. cobertura ---------------------------------------------------
        rangos = sorted((e["pdf_pages"][0], e["pdf_pages"][1], e.get("specification"))
                        for e in entradas if len(e.get("pdf_pages") or []) == 2)
        solapes, huecos = [], []
        for (a0, a1, an), (b0, b1, bn) in zip(rangos, rangos[1:]):
            if b0 <= a1:
                solapes.append((an, a1, bn, b0))
            elif b0 > a1 + 1:
                huecos.append((a1 + 1, b0 - 1, an, bn))
        cubierto = sum(b - a + 1 for a, b, _ in rangos)
        preliminar = rangos[0][0] if rangos else 0          # 0-based: es el conteo
        final = npag - (rangos[-1][1] + 1) if rangos else 0
        log("| %s | rangos que se solapan | %d | %s |"
            % (parte, len(solapes), "OK" if not solapes else "REVISAR"))
        fallos += 0 if not solapes else 1
        log("| %s | paginas cubiertas por las entradas | %d de %d (%d preliminares, "
            "%d finales, %d en huecos interiores) | %s |"
            % (parte, cubierto, npag, preliminar, final,
               npag - cubierto - preliminar - final,
               "OK" if not huecos else "revisar huecos"))
        for h in huecos[:6]:
            log("|   | hueco interior | paginas %d-%d, entre %s y %s | revisar |" % h)

        # --- 6. figuras -----------------------------------------------------
        figs = carpeta / "figures"
        n_png = len(list(figs.glob("*.png"))) if figs.is_dir() else 0
        decl = totales.get("figures")
        ok = decl is None or decl == n_png
        log("| %s | figuras PNG | %s declaradas / %d en disco | %s |"
            % (parte, decl, n_png, "OK" if ok else "REVISAR"))
        fallos += 0 if ok else 1

    for s in sin_ancla[:8]:
        log("|   | sin anclaje | %s en pagina %s: %s | REVISAR |" % s)
    for s in folio_mal[:6]:
        log("|   | folio | %s pagina PDF %s: se esperaba el folio %s | REVISAR |" % s)
    return fallos


IID = {
    "bpvc_ii_d_metric_2025": "D Metric 2025 .pdf",
    "bpvc_ii_d_customary_2025": "D Customary 2025 .pdf",
}


def auditar_iid(parte, carpeta, pdf_path, log):
    """La II-D no es texto troceado sino tablas de valores: lo que se audita es
    que cada tabla empiece donde dice y que las filas cargadas sean las que
    declara. El desplazamiento de pagina se DETECTA, no se supone: las dos
    extracciones del proyecto no usan la misma convencion."""
    import pdfplumber

    tablas = []
    for f in sorted(carpeta.glob("table_*.json")):
        j = json.loads(f.read_text(encoding="utf-8"))
        if isinstance(j, dict) and j.get("pdf_pages") and j.get("table_id"):
            tablas.append((f.name, j))
    fallos = 0

    with pdfplumber.open(pdf_path) as pdf:
        npag = len(pdf.pages)
        cache = {}

        def texto(i):
            if i not in cache:
                cache[i] = plano(pdf.pages[i].extract_text() or "") \
                    if 0 <= i < npag else ""
            return cache[i]

        # 1. deteccion del desplazamiento sobre las 8 primeras tablas
        votos = {}
        for _, j in tablas[:8]:
            # `table_id` ya viene como "Table 1A": no se le antepone nada.
            ancla = plano(str(j["table_id"]))
            p = j["pdf_pages"][0]
            for off in (0, -1, 1):
                if ancla and ancla in texto(p + off):
                    votos[off] = votos.get(off, 0) + 1
        off = max(votos, key=votos.get) if votos else 0
        log("| %s | convencion de `pdf_pages` | desplazamiento %+d (%d de %d "
            "tablas de muestra) | %s |"
            % (parte, off, votos.get(off, 0), min(8, len(tablas)),
               "OK" if votos else "REVISAR"))
        fallos += 0 if votos else 1

        # 2. anclaje de cada tabla y coherencia de row_count
        sin_ancla, mal_filas, fuera = [], [], []
        for nombre, j in tablas:
            p0, p1 = j["pdf_pages"][0] + off, j["pdf_pages"][1] + off
            if not (0 <= p0 < npag and 0 <= p1 < npag):
                fuera.append((j["table_id"], j["pdf_pages"]))
                continue
            # `table_id` ya viene como "Table 1A": no se le antepone nada.
            ancla = plano(str(j["table_id"]))
            # La cabecera se repite en cada pagina de la tabla: basta con que
            # aparezca en alguna de las tres primeras del rango.
            if not any(ancla in texto(p) for p in range(p0, min(p0 + 3, p1 + 1))):
                sin_ancla.append((j["table_id"], p0 + 1))
            rc, real = j.get("row_count"), len(j.get("rows") or [])
            if rc is not None and rc != real:
                mal_filas.append((j["table_id"], rc, real))

        log("| %s | tablas cuyo `table_id` aparece en su primera pagina | %d de %d | %s |"
            % (parte, len(tablas) - len(sin_ancla) - len(fuera), len(tablas),
               "OK" if not sin_ancla else "REVISAR"))
        fallos += 0 if not sin_ancla else 1
        log("| %s | rangos dentro del PDF | %d de %d (el PDF tiene %d paginas) | %s |"
            % (parte, len(tablas) - len(fuera), len(tablas), npag,
               "OK" if not fuera else "REVISAR"))
        fallos += 0 if not fuera else 1
        log("| %s | `row_count` coincide con las filas cargadas | %d de %d | %s |"
            % (parte, len(tablas) - len(mal_filas), len(tablas),
               "OK" if not mal_filas else "REVISAR"))
        fallos += 0 if not mal_filas else 1
        for s in sin_ancla[:8]:
            log("|   | sin anclaje | Tabla %s, pagina PDF %d | REVISAR |" % s)
        for s in mal_filas[:8]:
            log("|   | filas | Tabla %s: declara %s, trae %d | REVISAR |" % s)
        for s in fuera[:4]:
            log("|   | rango | Tabla %s: %s queda fuera del PDF | REVISAR |" % s)
    return fallos


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--pdfs", required=True, type=Path)
    ap.add_argument("--report", type=Path, default=None)
    a = ap.parse_args()

    out = []

    def log(s):
        print(s)
        out.append(s)

    log("# Verificacion de ASME BPVC Seccion II, partes A, B y C")
    log("")
    log("Cada especificacion troceada se contrasta con su PDF de origen: que el "
        "corte empiece donde dice, que el folio impreso cuadre y que no se pierdan "
        "paginas entre una especificacion y la siguiente.")
    log("")
    log("| Parte | Comprobacion | Detalle | Estado |")
    log("|---|---|---|---|")

    total = 0
    for parte, pdf in PARTES.items():
        carpeta = a.resources / SEC_II / parte
        ruta_pdf = a.pdfs / pdf
        if not carpeta.is_dir():
            log("| %s | carpeta | no existe en resources/ | REVISAR |" % parte)
            total += 1
            continue
        if not ruta_pdf.is_file():
            log("| %s | PDF | no encontrado en %s | REVISAR |" % (parte, ruta_pdf))
            total += 1
            continue
        total += auditar(parte, carpeta, ruta_pdf, log)

    for parte, pdf in IID.items():
        carpeta = a.resources / SEC_II / parte
        ruta_pdf = a.pdfs / pdf
        if not carpeta.is_dir():
            log("| %s | carpeta | no existe en resources/ | REVISAR |" % parte)
            total += 1
        elif not ruta_pdf.is_file():
            log("| %s | PDF | no encontrado en %s | REVISAR |" % (parte, ruta_pdf))
            total += 1
        else:
            total += auditar_iid(parte, carpeta, ruta_pdf, log)

    log("")
    log("**Total de fallos: %d.**" % total)
    if a.report:
        a.report.write_text("\n".join(out) + "\n", encoding="utf-8")
        print("\nReporte escrito en %s" % a.report)
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
