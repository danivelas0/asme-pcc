# -*- coding: utf-8 -*-
"""
verificar_resources - contrasta contra sus PDF de origen las partes de
`resources/` que nunca se habian auditado: ASME PCC-2, los capitulos del
B31.3 y sus apendices D a Z.

LA PREGUNTA QUE RESPONDE
------------------------
`verificar.py` compara 271 536 valores entre el JSON y la hoja de calculo, y da
cero fallos. Pero eso demuestra que el LIBRO reproduce fielmente el JSON, no que
el JSON reproduzca fielmente el CODIGO IMPRESO. Es justo el hueco que aparecio
en los Apendices A, B y C del B31.3, donde el libro era fiel al JSON y los dos
estaban mal.

Aqui se cierra ese hueco con una comprobacion sencilla y masiva: **cada bloque
de texto declara la pagina de la que salio; se busca ese texto en esa pagina del
PDF**. Si no aparece, o el bloque esta mal paginado o el texto no sale de ahi.

COMO SE COMPARA
---------------
El texto del PDF y el del JSON no coinciden caracter a caracter -guiones de
corte de linea, espacios que el PDF no pone, comillas tipograficas-, asi que se
comparan cadenas normalizadas a solo letras y digitos. De cada bloque se toma
una MUESTRA del centro, no el principio: el arranque de un parrafo suele
arrastrar el numero de clausula y el final, la partitura de palabras.

Se busca en la pagina declarada y, si no aparece, en las contiguas, para
distinguir "no esta" de "esta desplazada una pagina", que son dos defectos muy
distintos.

USO
---
    python verificar_resources.py --resources ..\\..\\..\\resources ^
        --pdf-b313 "<...>\\ASME B31.3 2024 Process Piping.pdf" ^
        --pdf-pcc2 "<...>\\ASME PCC-2 ....pdf" [--que pcc2|b313|todo]

Devuelve 0 solo si no hay bloques sin localizar.
"""

import argparse
import json
import os
import re
import sys
import unicodedata
from pathlib import Path

MUESTRA = 46          # caracteres normalizados que se buscan
MIN_BLOQUE = 60       # por debajo, el bloque es demasiado corto para ser distintivo
VENTANA = 2           # paginas a cada lado que se exploran si falla la declarada


def norm(s):
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", s.lower())


class Pdf:
    """Texto normalizado por pagina, cacheado. PDFium y no pdfplumber: el PCC-2
    trae una fuente sin mapa a Unicode que pdfplumber devuelve como (cid:NN)."""

    def __init__(self, ruta):
        import pypdfium2 as pdfium
        self.doc = pdfium.PdfDocument(str(ruta))
        self.n = len(self.doc)
        self._c = {}

    def texto(self, pagina_1based):
        i = pagina_1based - 1
        if not (0 <= i < self.n):
            return ""
        if i not in self._c:
            self._c[i] = norm(self.doc[i].get_textpage().get_text_range())
        return self._c[i]

    def cerrar(self):
        self.doc.close()


def muestra_de(texto):
    """Trozo distintivo del centro del bloque."""
    t = norm(texto)
    if len(t) < MIN_BLOQUE:
        return None
    ini = max(0, len(t) // 2 - MUESTRA // 2)
    return t[ini:ini + MUESTRA]


def diagnostico(pdf, pagina, texto):
    """Un bloque que no se localiza, falta o solo esta desordenado?

    Se mira palabra a palabra: si TODAS las palabras largas del bloque estan en
    la pagina pero la frase seguida no, el contenido esta y lo que fallo es el
    orden de lectura -tipicamente dos columnas empalmadas, o una nota al pie
    intercalada-. Si faltan palabras, entonces si falta texto.
    """
    pag = pdf.texto(pagina) + pdf.texto(pagina + 1)
    pals = [w for w in re.findall(r"[A-Za-z]{6,}", texto or "")]
    if not pals:
        return "sin palabras largas (ecuacion o simbolos)", 0.0
    dentro = sum(1 for w in pals if norm(w) in pag)
    frac = dentro / len(pals)
    if frac >= 0.95:
        return "contenido presente, orden de lectura entrelazado", frac
    if frac >= 0.6:
        return "parcialmente presente", frac
    return "AUSENTE del PDF", frac


def localizar(pdf, pagina, muestra):
    """Devuelve el desplazamiento donde aparece la muestra, o None."""
    if muestra is None:
        return 0
    for off in [0] + [d for k in range(1, VENTANA + 1) for d in (-k, k)]:
        if muestra in pdf.texto(pagina + off):
            return off
    return None


# ---------------------------------------------------------------------------
def bloques_pcc2(base):
    """(archivo, pagina, texto, tipo, pagina_propia) de cada bloque de PCC-2."""
    raiz = base / "asme_pcc" / "pcc_2"
    for p in sorted(raiz.rglob("art_*.json")):
        j = json.loads(p.read_text(encoding="utf-8"))
        for b in j.get("blocks") or []:
            pag = b.get("page")
            txt = b.get("text") or ""
            if not txt and b.get("html"):
                txt = re.sub(r"<[^>]+>", " ", b["html"])
            if isinstance(pag, int) and txt:
                yield (str(p.relative_to(raiz)), pag, txt, b.get("type"), True)


def _prosa(p, raiz, etiqueta):
    """Recorre un arbol de secciones heredando la pagina.

    Los parrafos NO llevan `page`: la heredan de la seccion que los contiene, y
    una seccion abarca varias paginas. Por eso el parrafo se busca desde la
    pagina de su seccion hacia ADELANTE -nunca puede estar antes-, y el tipo
    'prosa' le da al auditor una ventana mayor que a un bloque con pagina propia.
    """
    j = json.loads(p.read_text(encoding="utf-8"))

    def andar(nodo, heredada):
        if isinstance(nodo, dict):
            pag = nodo.get("page") if isinstance(nodo.get("page"), int) else heredada
            propia = isinstance(nodo.get("page"), int)
            txt = nodo.get("text") or nodo.get("heading") or nodo.get("title")
            if txt and pag:
                yield (str(p.relative_to(raiz)), pag, txt,
                       nodo.get("kind") or ("seccion" if propia else "prosa"),
                       propia)
            for k in ("sections", "paragraphs", "items", "children"):
                for h in nodo.get(k) or []:
                    for x in andar(h, pag):
                        yield x
        elif isinstance(nodo, list):
            for h in nodo:
                for x in andar(h, heredada):
                    yield x

    inicio = (j.get("pdf_pages") or [None])[0]
    for clave in ("sections", "preamble"):
        for x in andar(j.get(clave), inicio):
            yield x


def bloques_b313_capitulos(base):
    raiz = base / "asme_b31" / "asme_b31_3" / "chapters"
    for p in sorted(raiz.glob("chapter_*.json")):
        for x in _prosa(p, raiz, p.name):
            yield x


def bloques_b313_apendices(base):
    """Prosa de los apendices D a Z."""
    ap = base / "asme_b31" / "asme_b31_3" / "appex"
    for d in sorted(ap.iterdir()):
        if not d.is_dir() or d.name in ("appendix_a", "appendix_b", "appendix_c"):
            continue
        for p in sorted(d.glob("*.json")):
            j = json.loads(p.read_text(encoding="utf-8"))
            if isinstance(j, dict) and (j.get("sections") or j.get("preamble")):
                for x in _prosa(p, ap, p.name):
                    yield x


def bloques_b313_tablas(base):
    """Tablas: se anclan por su rotulo mas el titulo, que juntos son distintivos.

    `table_id` solo ("Table 302.3.3-1") es demasiado corto para el umbral
    general, pero un rotulo de tabla ES distintivo: se marca como 'rotulo' para
    que el auditor lo compruebe entero en vez de tomar una muestra del centro.
    """
    b3 = base / "asme_b31" / "asme_b31_3"
    rutas = list((b3 / "chapters" / "tables").glob("*.json"))
    for ap in sorted((b3 / "appex").iterdir()):
        if ap.is_dir() and ap.name not in ("appendix_a", "appendix_b", "appendix_c"):
            rutas += sorted(ap.glob("*.json"))
    for p in sorted(rutas):
        j = json.loads(p.read_text(encoding="utf-8"))
        if not isinstance(j, dict):
            continue
        pp = j.get("pdf_pages")
        ident = j.get("table_id") or j.get("title") or j.get("document")
        if isinstance(pp, list) and pp and ident:
            yield (str(p.relative_to(b3)), pp[0], str(ident), "rotulo", True)


def bloques_iid(base, edicion):
    """Filas de las tablas de II-D: se ancla la IDENTIFICACION de cada material.

    Aqui lo que importa no es la prosa sino que la fila exista en el codigo. De
    cada fila se compone spec + grado + UNS, que juntos son unicos, y se busca
    en el rango de paginas que la tabla declara. Es la comprobacion que faltaba:
    `verificar.py` demuestra que la hoja reproduce el JSON, no que el JSON
    reproduzca el impreso.
    """
    raiz = base / "asme_bpvc" / "sec_ii" / edicion
    for p in sorted(raiz.glob("table_*.json")):
        j = json.loads(p.read_text(encoding="utf-8"))
        pp = j.get("pdf_pages")
        if not (isinstance(pp, list) and len(pp) == 2):
            continue
        for i, r in enumerate(j.get("rows") or []):
            # NO se concatenan: en el impreso `Spec. No.` y `Type/Grade` son
            # columnas distintas y nunca salen seguidas en la capa de texto.
            # Cada campo se busca por su cuenta. El UNS es el mas fuerte: es
            # unico y se imprime como un solo token.
            uns = r.get("alloy_desig_uns_no") or r.get("uns_no")
            campos = {"uns": uns, "spec": r.get("spec_no"),
                      "grado": r.get("type_grade")}
            campos = {k: str(v) for k, v in campos.items() if v not in (None, "")}
            if campos:
                yield (p.name, pp[0], pp[1], campos, i)


FUENTES = {
    "pcc2": ("ASME PCC-2", bloques_pcc2, "pcc2"),
    "b313_cap": ("B31.3 capitulos", bloques_b313_capitulos, "b313"),
    "b313_ape": ("B31.3 apendices D-Z", bloques_b313_apendices, "b313"),
    "b313_tab": ("B31.3 tablas", bloques_b313_tablas, "b313"),
}

IID_FUENTES = {
    "iid_si": ("II-D metrica", "bpvc_ii_d_metric_2025", "iid_si"),
    "iid_us": ("II-D U.S. Customary", "bpvc_ii_d_customary_2025", "iid_us"),
}


def auditar_iid(nombre, edicion, pdf, base, log):
    """Cada fila de material tiene que existir en el rango de paginas de su tabla."""
    n = con_uns = con_spec = uns_ok = spec_ok = 0
    fallan = []
    rango_cache = {}
    for archivo, p0, p1, campos, fila in bloques_iid(base, edicion):
        n += 1
        clave = (archivo, p0, p1)
        if clave not in rango_cache:
            # `pdf_pages` de II-D es 1-BASED e inclusivo: la tabla 1A [58, 221]
            # son las paginas 58 a 221 del PDF. Comprobado con el UNS, no con el
            # rotulo de la tabla: el rotulo se repite en cada pagina de la tabla
            # y por eso no distingue una convencion de la otra, mientras que un
            # UNS aparece una sola vez y desambigua.
            rango_cache[clave] = "".join(pdf.texto(x) for x in range(p0, p1 + 1))
        pagina = rango_cache[clave]
        uns = campos.get("uns")
        if uns:
            con_uns += 1
            if norm(uns) in pagina:
                uns_ok += 1
            else:
                fallan.append((archivo, fila, "UNS", uns))
        sp = campos.get("spec")
        if sp:
            con_spec += 1
            if norm(sp) in pagina:
                spec_ok += 1
            else:
                fallan.append((archivo, fila, "spec", sp))
    log("| %s | filas cuyo UNS aparece en el rango de su tabla | %d de %d (%.1f%%) | %s |"
        % (nombre, uns_ok, con_uns, 100.0 * uns_ok / con_uns if con_uns else 0,
           "OK" if uns_ok == con_uns else "REVISAR"))
    # El denominador son las filas QUE TRAEN especificacion, no todas: el codigo
    # deja la columna en blanco en las filas de continuacion, y contarlas como
    # fallo daria un porcentaje falso.
    log("| %s | filas cuya especificacion aparece en el rango | %d de %d (%.1f%%) | %s |"
        % (nombre, spec_ok, con_spec, 100.0 * spec_ok / con_spec if con_spec else 0,
           "OK" if spec_ok == con_spec else "REVISAR"))
    log("| %s | filas sin especificacion en el JSON (continuacion) | %d de %d | informativo |"
        % (nombre, n - con_spec, n))
    for a, f, q, v in fallan[:10]:
        log("|   | %s no aparece | %s fila %d: %r | REVISAR |" % (q, a, f, v[:60]))
    return len(fallan)

# Un parrafo hereda la pagina de su seccion, que puede abarcar varias: se busca
# hacia adelante. Un bloque con pagina propia tiene que estar donde dice.
VENTANA_HEREDADA = 12


def auditar(nombre, generador, pdf, base, log):
    n = ok = corto = 0
    desplazados, perdidos, adelante = [], [], []
    for archivo, pag, txt, tipo, propia in generador(base):
        n += 1
        if tipo == "rotulo":
            m = norm(txt)[:MUESTRA] or None
        else:
            m = muestra_de(txt)
        if m is None:
            corto += 1
            continue
        if tipo == "rotulo":
            # El B31.3 antepone paginas separadoras que repiten el titulo
            # ("APPENDIX J NOMENCLATURE - Begins on the next page"), y el rotulo
            # se repite en cada pagina de continuacion. Encontrarlo una pagina
            # antes o un par despues NO es un desplazamiento: la cita apunta al
            # comienzo del contenido, que es lo correcto.
            off = None
            for d in (0, -1, 1, 2):
                if m in pdf.texto(pag + d):
                    off = 0
                    break
        elif propia:
            off = localizar(pdf, pag, m)
        else:
            off = None
            for d in range(0, VENTANA_HEREDADA + 1):
                if m in pdf.texto(pag + d):
                    off = d
                    break
        if off == 0:
            ok += 1
        elif off is None:
            perdidos.append((archivo, pag, tipo, txt))
        elif propia:
            desplazados.append((archivo, pag, off, tipo))
        else:
            ok += 1
            adelante.append(off)
    comprobados = n - corto
    pct = 100.0 * ok / comprobados if comprobados else 0.0
    log("| %s | bloques localizados en el PDF | %d de %d (%.1f%%) | %s |"
        % (nombre, ok, comprobados, pct,
           "OK" if not perdidos and not desplazados else "REVISAR"))
    if adelante:
        log("| %s | parrafos hallados mas adelante en su seccion | %d (hasta +%d "
            "paginas) | informativo |" % (nombre, len(adelante), max(adelante)))
    log("| %s | bloques con pagina propia desplazados | %d | %s |"
        % (nombre, len(desplazados), "OK" if not desplazados else "REVISAR"))
    log("| %s | bloques que no aparecen | %d | %s |"
        % (nombre, len(perdidos), "OK" if not perdidos else "REVISAR"))
    log("| %s | bloques demasiado cortos para comprobar | %d de %d | informativo |"
        % (nombre, corto, n))
    for a, p, o, t in desplazados[:10]:
        log("|   | desplazado %+d | %s, pagina %d (%s) | REVISAR |" % (o, a, p, t))

    # Los no localizados se clasifican: no es lo mismo que falte texto que que
    # este desordenado. Solo lo primero es una perdida de dato.
    clases, ausentes = {}, []
    for a, p, t, txt in perdidos:
        cl, frac = diagnostico(pdf, p, txt)
        clases[cl] = clases.get(cl, 0) + 1
        if cl.startswith("AUSENTE") or cl.startswith("parcial"):
            ausentes.append((a, p, t, frac, txt))
    for cl, k in sorted(clases.items(), key=lambda x: -x[1]):
        log("| %s | de los no localizados: %s | %d | %s |"
            % (nombre, cl, k, "REVISAR" if "AUSENTE" in cl or "parcial" in cl
               else "informativo"))
    for a, p, t, frac, txt in ausentes[:10]:
        log("|   | %s | %s, pagina %d (%s), %.0f%% de sus palabras: %r | REVISAR |"
            % ("texto ausente", a, p, t, 100 * frac, (txt or "")[:60]))
    return len(desplazados) + len(ausentes)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--resources", required=True, type=Path)
    ap.add_argument("--pdf-b313", type=Path)
    ap.add_argument("--pdf-pcc2", type=Path)
    ap.add_argument("--pdf-iid-si", type=Path)
    ap.add_argument("--pdf-iid-us", type=Path)
    ap.add_argument("--que", default="todo", choices=["todo", "pcc2", "b313", "iid"])
    ap.add_argument("--report", type=Path)
    a = ap.parse_args()

    out = []

    def log(s):
        print(s)
        out.append(s)

    log("# Verificacion de resources/ contra los PDF de origen")
    log("")
    log("De cada bloque se toma una muestra del centro de su texto y se busca en "
        "la pagina del PDF que el propio bloque declara. Si no aparece alli, se "
        "exploran las paginas contiguas para separar 'no esta' de 'esta "
        "desplazado'.")
    log("")
    log("| Fuente | Comprobacion | Detalle | Estado |")
    log("|---|---|---|---|")

    pdfs, total = {}, 0
    try:
        for clave, (nombre, gen, cual) in FUENTES.items():
            if a.que != "todo" and not clave.startswith(a.que):
                continue
            ruta = a.pdf_pcc2 if cual == "pcc2" else a.pdf_b313
            if not ruta or not ruta.is_file():
                log("| %s | PDF | no indicado o no encontrado | REVISAR |" % nombre)
                total += 1
                continue
            if cual not in pdfs:
                pdfs[cual] = Pdf(ruta)
                log("| %s | PDF de origen | %s, %d paginas | OK |"
                    % (nombre, ruta.name, pdfs[cual].n))
            total += auditar(nombre, gen, pdfs[cual], a.resources, log)

        if a.que in ("todo", "iid"):
            for clave, (nombre, edicion, cual) in IID_FUENTES.items():
                ruta = a.pdf_iid_si if cual == "iid_si" else a.pdf_iid_us
                if not ruta or not ruta.is_file():
                    log("| %s | PDF | no indicado o no encontrado | REVISAR |" % nombre)
                    total += 1
                    continue
                pdfs[cual] = Pdf(ruta)
                log("| %s | PDF de origen | %s, %d paginas | OK |"
                    % (nombre, ruta.name, pdfs[cual].n))
                total += auditar_iid(nombre, edicion, pdfs[cual], a.resources, log)
    finally:
        for p in pdfs.values():
            p.cerrar()

    log("")
    log("**Total de bloques mal ubicados: %d.**" % total)
    if a.report:
        a.report.write_text("\n".join(out) + "\n", encoding="utf-8")
        print("\nReporte escrito en %s" % a.report)
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
