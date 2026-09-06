# -*- coding: utf-8 -*-
"""Post-extraction checks for the body of ASME B31.3-2024.

The strongest check available is not internal consistency but the PDF's own bookmark tree:
it names every paragraph of the Code and the page it starts on, and it was not used to build
the section trees, so it is independent evidence. The Code also prints its own List of
Figures and List of Tables, which fix how many of each there should be.
"""
import io
import json
import os
import re
import sys

import pymupdf

import b313
import run_body as R

DEST = R.DEST


def load(name):
    with io.open(os.path.join(DEST, name), encoding="utf-8") as fh:
        return json.load(fh)


def walk(nodes):
    for n in nodes:
        if n["kind"] == "section":
            yield n
        for m in walk(n["sections"]):
            yield m


def toc_lists(doc):
    """(paragraph ids, figure ids, table captions) as the book itself declares them."""
    paras, figs, tabs, mode = {}, [], [], None
    for lvl, title, pg in doc.get_toc():
        if lvl == 1:
            mode = {"Figures": "F", "Tables": "T"}.get(title.strip())
            continue
        if mode == "F" and pg - 1 < 197:
            figs.append(title.strip())
            continue
        if mode == "T" and pg - 1 < 197:
            tabs.append(title.strip())
            continue
        if mode is None and R.FIRST <= pg - 1 <= R.LAST:
            m = re.match(r"^([A-Z]{0,2}\d{3}(?:\.\d+)*)\b", title.strip())
            if m:
                paras.setdefault(m.group(1), pg - 1)
    return paras, figs, tabs


def main():
    doc = pymupdf.open(b313.SRC)
    idx = load("chapters_index.json")
    problems = []

    sections, pages_of = {}, {}
    for c in idx["chapters"]:
        js = load(c["file"])
        for s in walk(js["sections"]):
            if s["id"]:
                sections.setdefault(s["id"], s["page"])
        pages_of[c["chapter"]] = js

    toc_paras, toc_figs, toc_tabs = toc_lists(doc)

    # -- 1. every paragraph the bookmark tree names must be in a section tree
    missing = sorted(p for p in toc_paras if p not in sections)
    print("paragraphs in bookmark tree: %d   present in output: %d   missing: %d"
          % (len(toc_paras), len(toc_paras) - len(missing), len(missing)))
    if missing:
        print("   missing:", ", ".join(missing[:40]))
        problems.append("%d bookmarked paragraphs absent" % len(missing))

    # -- 2. and it must be on the page the book says
    off = [(p, sections[p], toc_paras[p] + 1) for p in toc_paras
           if p in sections and abs(sections[p] - (toc_paras[p] + 1)) > 1]
    print("paragraphs on an unexpected page: %d" % len(off))
    for p, got, want in off[:10]:
        print("   %-12s output p%-4d book p%d" % (p, got, want))
    if off:
        problems.append("%d paragraphs on the wrong page" % len(off))

    # -- 3. figures and tables against the Code's own lists
    got_figs = {f["figure_id"] for c in idx["chapters"] for f in c["figures"]}
    want_figs = {re.match(r"^(Figure [^\s]+)", t).group(1) for t in toc_figs
                 if t.startswith("Figure")}
    print("figures: list of figures %d, extracted %d, missing %s, extra %s"
          % (len(want_figs), len(got_figs), sorted(want_figs - got_figs) or "none",
             sorted(got_figs - want_figs) or "none"))
    if want_figs - got_figs:
        problems.append("figures missing: %s" % sorted(want_figs - got_figs))

    # The Code's List of Tables names 32 "Table x" sheets plus one "Criterion Value Notes
    # for Table K341.3.2-1". The matching sheet for Table 341.3.2-1 is printed too but is
    # not listed, so it is expected in the output without being expected in the list.
    got_tabs = {t["table_id"] for c in idx["chapters"] for t in c["tables"]}
    want_tabs = {re.sub(r"^(Table [^\s]+|Criterion Value Notes for Table [^\s]+).*$", r"\1", t)
                 for t in toc_tabs}
    unlisted = {"Criterion Value Notes for Table 341.3.2-1"}
    missing = want_tabs - got_tabs
    extra = got_tabs - want_tabs - unlisted
    print("tables:  list of tables %d, extracted %d (+%d printed but unlisted), "
          "missing %s, unexpected %s"
          % (len(want_tabs), len(got_tabs), len(got_tabs & unlisted),
             sorted(missing) or "none", sorted(extra) or "none"))
    if missing:
        problems.append("tables missing: %s" % sorted(missing))
    if extra:
        problems.append("unexpected tables: %s" % sorted(extra))

    # -- 4. no page furniture anywhere in the text
    bad = []
    for name in [c["file"] for c in idx["chapters"]] + ["definitions.json"]:
        blob = io.open(os.path.join(DEST, name), encoding="utf-8").read()
        for pat, what in ((r"ð\s*\d\d\s*Þ", "change marker"),
                          (r"ASME B31\.3-2024", "running head"),
                          (r"\(Cont’d\)", "continuation marker")):
            n = len(re.findall(pat, blob))
            if n:
                bad.append("%s: %d x %s" % (name, n, what))
    print("page furniture in output: %s" % ("; ".join(bad) if bad else "none"))
    if bad:
        problems.append("page furniture present")

    # -- 5. every referenced file exists, and row counts match
    dangling, rowbad = [], []
    for c in idx["chapters"]:
        for f in c["figures"]:
            for rel in [f["figure_image"]] + f.get("figure_images", []):
                if not os.path.exists(os.path.join(DEST, rel.replace("/", os.sep))):
                    dangling.append(rel)
        for t in c["tables"]:
            path = os.path.join(DEST, t["file"].replace("/", os.sep))
            if not os.path.exists(path):
                dangling.append(t["file"])
                continue
            js = json.load(io.open(path, encoding="utf-8"))
            if js["row_count"] != len(js["rows"]) or js["row_count"] != t["rows"]:
                rowbad.append(t["table_id"])
    print("dangling file references: %d %s" % (len(dangling), dangling[:5] or ""))
    print("tables whose row_count disagrees: %d %s" % (len(rowbad), rowbad[:5] or ""))
    if dangling:
        problems.append("dangling references")
    if rowbad:
        problems.append("row_count mismatch")

    # -- 6. residual line-break hyphens in prose
    hy = 0
    for c in idx["chapters"]:
        blob = io.open(os.path.join(DEST, c["file"]), encoding="utf-8").read()
        hy += len(re.findall(r"\w[-‐] \w", blob))
    print("residual line-break hyphens: %d" % hy)

    print()
    print("PROBLEMS: %s" % ("; ".join(problems) if problems else "none"))
    return problems


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.exit(1 if main() else 0)
