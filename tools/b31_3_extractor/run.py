# -*- coding: utf-8 -*-
"""Extract ASME B31.3-2024 appendices E through Z to JSON (+ rendered figure PNGs).

Appendices A, B and C already exist in the destination and are left untouched.
"""
import io
import json
import os
import re
import sys

import pymupdf

import b313
import build as B
import figures as F
import prose as P
import special as S
import tables as T
from b313 import SOURCE_BLOCK, clean, page_lines

DEST = (r"C:\Users\User\Google Drive Streaming\My Drive\DANIEL\Cloude\ASME PCC"
        r"\resources\ASME B31\ASME B31.3\APPEX")

# appendix, title, first idx, last idx (0-based)
APPENDICES = [
    ("E", "Reference Standards", 462, 466),
    ("F", "Guidance and Precautionary Considerations", 467, 474),
    ("G", "Safeguarding", 475, 476),
    ("H", "Sample Calculations for Branch Reinforcement", 477, 484),
    ("J", "Nomenclature", 485, 501),
    ("K", "Allowable Stresses for High Pressure Piping", 502, 531),
    ("L", "Aluminum Alloy Pipe Flanges", 532, 534),
    ("M", "Guide to Classifying Fluid Services", 535, 536),
    ("N", "Application of ASME B31.3 Internationally", 537, 540),
    ("Q", "Quality System Program", 541, 541),
    ("R", "Use of Alternative Ultrasonic Acceptance Criteria", 542, 544),
    ("S", "Piping System Stress Analysis Examples", 545, 559),
    ("V", "Allowable Variations in Elevated Temperature Service", 560, 562),
    ("W", "High-Cycle Fatigue Assessment of Piping Systems", 563, 567),
    ("X", "Metallic Bellows Expansion Joints", 568, 571),
    ("Z", "Preparation of Technical Inquiries and Suggestions for Code Revision", 572, 572),
]

# Appendices handled entirely by a bespoke extractor - the generic prose/table pass is
# skipped for the pages they own.
BESPOKE_PAGES = {
    "E": set(range(462, 467)),
    "J": set(range(486, 502)),
    "K": set(range(503, 532)),
}


def slug(s):
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


def write(path, obj):
    with io.open(path, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=1)
    return os.path.basename(path)


def envelope(appendix, **kw):
    d = {"source": SOURCE_BLOCK, "appendix": appendix}
    d.update(kw)
    return d


def main():
    doc = pymupdf.open(b313.SRC)
    dehyph = P.Dehyphenator(doc, range(0, 594))
    manifest = []

    for code, title, first, last in APPENDICES:
        folder = os.path.join(DEST, "appendix_%s" % code.lower())
        os.makedirs(folder, exist_ok=True)
        pages = list(range(first, last + 1))
        skip = BESPOKE_PAGES.get(code, set())
        entry = {"appendix": code, "title": title, "folder": os.path.basename(folder),
                 "pdf_pages": [first + 1, last + 1], "tables": [], "figures": [],
                 "documents": []}

        # ---- figures first: their pages are excluded from the prose pass
        figs = B.collect_figures(doc, pages, folder, "appendix_%s" % code.lower())
        fig_pages = {f["pdf_page"] - 1 for f in figs}
        full_page_figs = {i for i in fig_pages
                          if len(page_lines(doc[i])) < 90 and doc[i].rotation}
        entry["figures"] = figs

        # ---- tables (generic ruled engine)
        tab_pages = [i for i in pages if i not in skip]
        found = B.collect_tables(doc, tab_pages, dehyph=dehyph)
        regions_by_page = {}
        for i in tab_pages:
            lines = page_lines(doc[i])
            regs = T.find_tables(doc[i], lines)
            if regs:
                regions_by_page[i] = [dict(r, y_top=r["y_top"] - 60) for r in regs]
        for e in found:
            js = B.table_json(e, code)
            fname = write(os.path.join(folder, "table_%s.json" % slug(e["label"])), js)
            entry["tables"].append({"file": fname, "table_id": e["label"],
                                    "title": e["title"], "rows": js["row_count"],
                                    "pdf_pages": js["pdf_pages"]})

        # ---- prose
        prose_pages = [i for i in pages if i not in skip and i not in fig_pages]
        for i in fig_pages:
            if i not in skip and i not in full_page_figs and len(page_lines(doc[i])) > 40:
                prose_pages.append(i)
        prose_pages.sort()
        pre, secs = B.collect_prose(doc, prose_pages, dehyph, regions_by_page)
        if secs or pre:
            js = envelope(code, title=title, document="Appendix %s prose" % code,
                          pdf_pages=[first + 1, last + 1],
                          section_count=len(secs),
                          paragraph_count=sum(len(s["paragraphs"]) for s in secs) + len(pre),
                          preamble=pre, sections=secs)
            fname = write(os.path.join(folder, "appendix_%s_text.json" % code.lower()), js)
            entry["documents"].append({"file": fname, "kind": "prose",
                                       "sections": len(secs),
                                       "paragraphs": js["paragraph_count"]})

        # ---- bespoke
        if code == "E":
            e = S.appendix_e(doc)
            js = envelope("E", title=title, pdf_pages=[463, 467],
                          introduction=[dehyph(p) for p in e["introduction"]],
                          general_notes=[dehyph(p) for p in e["general_notes"]],
                          publisher_count=len(e["standards"]),
                          standard_count=sum(len(p["entries"]) for p in e["standards"]),
                          standards=e["standards"],
                          organizations_note=dehyph(e["organizations_note"] or ""),
                          organizations=e["organizations"])
            fname = write(os.path.join(folder, "reference_standards.json"), js)
            entry["documents"].append({"file": fname, "kind": "standards_list",
                                       "standards": js["standard_count"],
                                       "organizations": len(e["organizations"])})
        elif code == "J":
            ents, notes = S.appendix_j(doc)
            js = envelope("J", title=title, document="Nomenclature",
                          pdf_pages=[487, 502], entry_count=len(ents),
                          notes=[dehyph(n) for n in notes], entries=ents)
            fname = write(os.path.join(folder, "nomenclature.json"), js)
            entry["documents"].append({"file": fname, "kind": "nomenclature",
                                       "entries": len(ents)})
        elif code == "K":
            spec = S.k_spec_index(doc)
            js = envelope("K", document="Specification Index for Appendix K",
                          pdf_pages=[504, 504],
                          entry_count=sum(len(g["entries"]) for g in spec), groups=spec)
            fname = write(os.path.join(folder, "spec_index_k.json"), js)
            entry["documents"].append({"file": fname, "kind": "specification_index",
                                       "entries": js["entry_count"]})

            notes_doc = k_notes(doc, dehyph)
            fname = write(os.path.join(folder, "notes_tables_k_1_k_1c.json"), notes_doc)
            entry["documents"].append({"file": fname, "kind": "table_notes"})

            for tid, a, b, units, comp in [("Table K-1", 506, 517, "SI", "Table K-1C"),
                                           ("Table K-1C", 518, 531, "U.S. Customary",
                                            "Table K-1")]:
                t = S.k_spread_table(doc, a, b, tid)
                js = envelope("K", table_id=tid, title=t["title"],
                              pdf_pages=[a + 1, b + 1], units=units,
                              companion_table=comp, row_count=len(t["rows"]),
                              header_banners=t["banners"],
                              stress_columns=t["stress_columns"], rows=t["rows"])
                fname = write(os.path.join(folder, "table_%s.json" % slug(tid)), js)
                entry["tables"].append({"file": fname, "table_id": tid,
                                        "title": t["title"], "rows": len(t["rows"]),
                                        "pdf_pages": [a + 1, b + 1],
                                        "companion_table": comp, "units": units})

        manifest.append(entry)

    write_index(manifest)
    return manifest


def k_notes(doc, dehyph):
    """The NOTES FOR TABLES K-1 and K-1C page, plus the small grade-equivalents table."""
    i = 504
    lines = page_lines(doc[i])
    left = [L for L in lines if L["bbox"][0] < 300 and L["bbox"][1] > 150]
    general, numbered, cur = [], [], None
    for L in sorted(left, key=lambda L: L["bbox"][1]):
        t = L["text"]
        m = re.match(r"^\((\d{1,2})\)\s*(.*)$", t)
        g = re.match(r"^\(([a-z])\)\s*(.*)$", t)
        if t.startswith("GENERAL NOTES"):
            cur = None
            continue
        if g:
            cur = {"id": "(%s)" % g.group(1), "text": g.group(2)}
            general.append(cur)
        elif m:
            cur = {"id": "(%s)" % m.group(1), "text": m.group(2)}
            numbered.append(cur)
        elif cur is not None:
            cur["text"] = clean(cur["text"] + " " + t)
    for c in general + numbered:
        c["text"] = dehyph(clean(c["text"]))

    grade = []
    right = [L for L in lines if L["bbox"][0] >= 300]
    rows = T._rows([L for L in right if L["bbox"][1] > 200], [340.0, 440.0, 560.0])
    for r in rows:
        if r[0] and r[1]:
            grade.append({"x_grade": r[0], "l_grade": r[1]})
    return envelope("K", document="NOTES FOR TABLES K-1 and K-1C",
                    applies_to=["Table K-1", "Table K-1C"], pdf_pages=[505, 505],
                    general_notes=general, notes=numbered,
                    api_5l_grade_equivalents=grade)


def write_index(manifest):
    path = os.path.join(DEST, "asme_b31_3_2024_index.json")
    with io.open(path, encoding="utf-8") as fh:
        idx = json.load(fh)
    keep = [a for a in idx["appendices"] if a["appendix"] in ("A", "B", "C")]
    idx["appendices"] = keep + manifest
    write(path, idx)


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    m = main()
    for e in m:
        print("%-2s %-52s tables=%2d figures=%d docs=%d"
              % (e["appendix"], e["title"][:52], len(e["tables"]),
                 len(e["figures"]), len(e["documents"])))
