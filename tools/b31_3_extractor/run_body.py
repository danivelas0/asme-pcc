# -*- coding: utf-8 -*-
"""Extract the body of ASME B31.3-2024 - Chapters I through X - to JSON, with the figures
rendered to PNG.

Scope is PDF idx 33-196: Chapter I opens at 33 and the appendices at 197. Everything before
Chapter I (Contents, Foreword, the committee roster, Correspondence, the Introduction, the
Summary of Changes and the redesignation list) is front matter and is not extracted, and
neither is the INDEX. The appendices already exist in APPEX/ and are not touched.
"""
import io
import json
import os
import re
import sys

import pymupdf

import b313
import body as BD
import prose as P
from b313 import SOURCE_BLOCK, num

DEST = (r"C:\Users\User\Google Drive Streaming\My Drive\DANIEL\Cloude\ASME PCC"
        r"\resources\ASME B31\ASME B31.3\CHAPTERS")

FIRST, LAST = 33, 196

# roman, arabic, title, first idx, last idx (0-based), taken from the bookmark tree
CHAPTERS = [
    ("I",    1, "Scope and Definitions",                               33,  42),
    ("II",   2, "Design",                                              43,  81),
    ("III",  3, "Materials",                                           82,  93),
    ("IV",   4, "Standards for Piping Components",                     94,  97),
    ("V",    5, "Fabrication, Assembly, and Erection",                 98, 116),
    ("VI",   6, "Inspection, Examination, and Testing",               117, 130),
    ("VII",  7, "Nonmetallic Piping and Piping Lined With Nonmetals", 131, 152),
    ("VIII", 8, "Piping for Category M Fluid Service",                153, 160),
    ("IX",   9, "High Pressure Piping",                               161, 188),
    ("X",   10, "High Purity Piping",                                 189, 196),
]

GREEK = {"\u0393": "gamma", "\u03b1": "alpha", "\u03b2": "beta", "\u03b3": "gamma",
         "\u0394": "delta", "\u03b4": "delta", "\u03b5": "epsilon", "\u03b8": "theta",
         "\u03bb": "lambda", "\u03bc": "mu", "\u03bd": "nu", "\u03c1": "rho",
         "\u03c3": "sigma", "\u03a3": "sigma", "\u03c4": "tau", "\u03c6": "phi",
         "\u03c9": "omega", "\u2211": "sum"}


def slug(s):
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


def key_of(header, banners, i):
    """A stable field name. The spanning banners lead, so two columns printed with the same
    label under different banners ("Fillet Weld" under three fluid services) stay distinct.
    """
    parts = list(banners) + ([header] if header else [])
    k = " ".join(parts).lower()
    for g, w in GREEK.items():
        k = k.replace(g, w).replace(g.lower(), w)
    k = re.sub(r"\[note[^\]]*\]", "", k)
    k = k.replace("\u00b0", "deg_").replace("%", "pct").replace("\u2264", "le_")
    k = re.sub(r"[^a-z0-9]+", "_", k).strip("_")
    return k or "column_%d" % (i + 1)


def uniq(keys):
    out, seen = [], {}
    for k in keys:
        if k in seen:
            seen[k] += 1
            k = "%s_%d" % (k, seen[k])
        else:
            seen[k] = 1
        out.append(k)
    return out


# Para. 300.2 is a seven-page glossary. Read as running prose it comes out as fifty-odd
# paragraphs of terms welded together, with the italic terms mistaken for list markers - a
# worse copy of what definitions.json already holds properly structured. So the glossary
# itself is kept out of the chapter prose and the section points at that file instead. Its
# lead-in paragraph and the AWS A3.0 footnote are left in place; only the entries are cut.
# Measured: the entries start at y=600 in the right column of idx 35, pages 36-41 are
# nothing but entries, and 300.3 opens idx 42.
DEFINITIONS_BAND = {35: [{"x0": 300.0, "x1": 1e9, "y_top": 592.0, "y_bot": 694.0}]}
for _i in range(36, 42):
    DEFINITIONS_BAND[_i] = [{"x0": -1e9, "x1": 1e9, "y_top": -1e9, "y_bot": 1e9}]


def _link_definitions(tree):
    """Point para. 300.2 at the glossary file rather than repeating it badly."""
    for node in _walk(tree):
        if node.get("id") == "300.2":
            node["definitions_file"] = "definitions.json"
            node["definition_note"] = (
                "The defined terms of para. 300.2 are extracted separately, one entry per "
                "term, in definitions.json.")


def _walk(nodes):
    for n in nodes:
        yield n
        for m in _walk(n["sections"]):
            yield m


def write(path, obj):
    with io.open(path, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=1)
    return os.path.basename(path)


def table_json(e, chapter, dehyph):
    heads = [dehyph(h) if h else None for h in e["head"]]
    spans = [[dehyph(b) for b in sp] for sp in e["spans"]]
    keys = uniq([key_of(h, sp, i) for i, (h, sp) in enumerate(zip(heads, spans))])
    rows = BD.merge_rows(e["body"])
    has_groups = any(r["group"] for r in rows)
    data = []
    for row in rows:
        cells, grp = row["cells"], row["group"]
        rec = {}
        for k, v in zip(keys, cells):
            val = num(v)
            rec[k] = dehyph(val) if isinstance(val, str) else val
        if not any(v is not None for v in rec.values()):
            continue                      # a rule artefact, not a printed row
        if has_groups:
            rec["group"] = grp
        data.append(rec)
    return {
        "source": SOURCE_BLOCK,
        "chapter": chapter,
        "table_id": e["label"],
        "title": dehyph(e["title"]) if e["title"] else None,
        "pdf_pages": [min(e["pages"]), max(e["pages"])],
        "columns": [{"key": k, "header": h, "banners": sp}
                    for k, h, sp in zip(keys, heads, spans)],
        "row_count": len(data),
        "notes": e["notes"] or [],
        "rows": data,
    }


def main():
    doc = pymupdf.open(b313.SRC)
    fig_dir = os.path.join(DEST, "figures")
    tab_dir = os.path.join(DEST, "tables")
    for d in (DEST, fig_dir, tab_dir):
        os.makedirs(d, exist_ok=True)

    dehyph = P.Dehyphenator(doc, range(0, doc.page_count))
    manifest = []

    for roman, arabic, title, first, last in CHAPTERS:
        pages = list(range(first, last + 1))
        entry = {"chapter": roman, "chapter_number": arabic, "title": title,
                 "pdf_pages": [first + 1, last + 1],
                 "file": "chapter_%02d.json" % arabic,
                 "tables": [], "figures": []}

        figs = BD.collect_figures(doc, pages, fig_dir, dehyph=dehyph)
        tabs = BD.collect_tables(doc, pages, dehyph=dehyph)
        exclude = BD.excluded_regions(doc, pages, figs)
        for pg_idx, regs in DEFINITIONS_BAND.items():
            if pg_idx in exclude:
                exclude[pg_idx].extend(regs)
        pre, secs = BD.collect_prose(doc, pages, dehyph, exclude)
        tree = BD.build_tree(secs)
        _link_definitions(tree)
        nsec, npar, neq, nwords = BD.count_tree(tree)

        for e in tabs:
            js = table_json(e, roman, dehyph)
            stem = slug(e["label"])
            if not stem.startswith("table_"):
                stem = "table_" + stem
            fname = write(os.path.join(tab_dir, "%s.json" % stem), js)
            entry["tables"].append({"file": "tables/" + fname, "table_id": e["label"],
                                    "title": js["title"], "rows": js["row_count"],
                                    "pdf_pages": js["pdf_pages"]})
        for f in figs:
            rec = {"figure_id": f["figure_id"], "title": f["title"],
                   "figure_image": "figures/" + f["figure_image"],
                   "pdf_page": f["pdf_page"], "landscape": f["landscape"],
                   "sheets": f["sheets"], "notes": f["notes"]}
            if f["sheets"] > 1:
                rec["figure_images"] = ["figures/" + s["file"] for s in f["_sheets"]]
            entry["figures"].append(rec)

        write(os.path.join(DEST, entry["file"]),
              {"source": SOURCE_BLOCK, "chapter": roman, "chapter_number": arabic,
               "title": title, "pdf_pages": [first + 1, last + 1],
               "section_count": nsec, "paragraph_count": npar,
               "equation_count": neq, "word_count": nwords,
               "tables": [t["table_id"] for t in entry["tables"]],
               "figures": [f["figure_id"] for f in entry["figures"]],
               "preamble": pre, "sections": tree})
        entry.update(sections=nsec, paragraphs=npar, equations=neq, words=nwords)
        manifest.append(entry)

    defs = BD.definitions(doc, dehyph=dehyph)
    write(os.path.join(DEST, "definitions.json"),
          {"source": SOURCE_BLOCK, "chapter": "I", "paragraph": "300.2",
           "document": "Definitions", "pdf_pages": [36, 43],
           "entry_count": len(defs), "entries": defs})

    write(os.path.join(DEST, "chapters_index.json"), {
        "source": SOURCE_BLOCK,
        "document": "ASME B31.3-2024 body of the Code, Chapters I-X",
        "pdf_pages": [FIRST + 1, LAST + 1],
        "scope_note": ("Chapters I-X only. Publisher front matter (Contents, Foreword, "
                       "committee roster, Correspondence, Introduction, Summary of Changes, "
                       "the redesignation list) and the INDEX are not included; neither are "
                       "running heads, printed folios or edition change markers. The "
                       "appendices are extracted separately, in APPEX/."),
        "definitions": {"file": "definitions.json", "entries": len(defs)},
        "chapter_count": len(manifest),
        "table_count": sum(len(c["tables"]) for c in manifest),
        "figure_count": sum(len(c["figures"]) for c in manifest),
        "chapters": manifest,
    })
    return manifest


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    m = main()
    for e in m:
        print("%-5s %-52s sec=%3d par=%4d eq=%3d tab=%2d fig=%2d"
              % (e["chapter"], e["title"][:52], e["sections"], e["paragraphs"],
                 e["equations"], len(e["tables"]), len(e["figures"])))
