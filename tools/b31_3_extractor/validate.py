# -*- coding: utf-8 -*-
"""Post-extraction checks: counts, line-number continuity, null ratios, file references."""
import io
import json
import os
import sys

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "out")


def load(p):
    with io.open(p, encoding="utf-8") as fh:
        return json.load(fh)


def main():
    idx = load(os.path.join(OUT, "asme_b31_3_2024_index.json"))
    tot_tables = tot_rows = tot_figs = tot_docs = 0
    problems = []
    files = 0

    for a in idx["appendices"]:
        if a["appendix"] in ("A", "B", "C"):
            continue
        folder = os.path.join(OUT, a["folder"])
        for t in a.get("tables", []):
            p = os.path.join(folder, t["file"])
            if not os.path.exists(p):
                problems.append("missing table file %s" % t["file"])
                continue
            files += 1
            js = load(p)
            rows = js.get("rows", [])
            tot_tables += 1
            tot_rows += len(rows)
            if len(rows) != t["rows"]:
                problems.append("%s row count mismatch" % t["table_id"])
            if not rows:
                problems.append("%s has no rows" % t["table_id"])
        for f in a.get("figures", []):
            p = os.path.join(folder, f["figure_image"])
            if not os.path.exists(p):
                problems.append("missing figure %s" % f["figure_image"])
            else:
                files += 1
                tot_figs += 1
        for dcm in a.get("documents", []):
            p = os.path.join(folder, dcm["file"])
            if not os.path.exists(p):
                problems.append("missing document %s" % dcm["file"])
            else:
                files += 1
                tot_docs += 1

    # Table K-1 / K-1C: line numbers must run 1..N inside every spread, and the two
    # editions must list the same number of materials.
    counts = {}
    for tid, fn in (("Table K-1", "table_table_k_1.json"),
                    ("Table K-1C", "table_table_k_1c.json")):
        js = load(os.path.join(OUT, "appendix_k", fn))
        rows = js["rows"]
        counts[tid] = len(rows)
        by_spread = {}
        for r in rows:
            by_spread.setdefault(tuple(r["pdf_pages"]), []).append(r["line_no"])
        for k, v in by_spread.items():
            if v != list(range(min(v), min(v) + len(v))):
                problems.append("%s spread %s line numbers not contiguous" % (tid, k))
        nostress = sum(1 for r in rows if not r.get("allowable_stress"))
        if nostress:
            problems.append("%s: %d rows with no stress values" % (tid, nostress))
    if counts["Table K-1"] != counts["Table K-1C"]:
        problems.append("K-1 (%d) and K-1C (%d) row counts differ"
                        % (counts["Table K-1"], counts["Table K-1C"]))

    prose_secs = prose_words = 0
    for a in idx["appendices"]:
        if a["appendix"] in ("A", "B", "C"):
            continue
        for dcm in a.get("documents", []):
            if dcm.get("kind") != "prose":
                continue
            js = load(os.path.join(OUT, a["folder"], dcm["file"]))
            prose_secs += len(js["sections"])
            for s in js["sections"]:
                for par in s["paragraphs"]:
                    prose_words += len(par["text"].split())
            for s in js["sections"]:
                for par in s["paragraphs"]:
                    if "- " in par["text"] and par["kind"] == "paragraph":
                        problems.append("residual line-break hyphen in %s %s"
                                        % (a["appendix"], (s["id"] or s["heading"])))
                        break

    print("files written      : %d" % files)
    print("tables             : %d  (%d data rows)" % (tot_tables, tot_rows))
    print("figures rendered   : %d" % tot_figs)
    print("documents          : %d" % tot_docs)
    print("prose sections     : %d  (%d words)" % (prose_secs, prose_words))
    print("K-1 / K-1C rows    : %d / %d" % (counts["Table K-1"], counts["Table K-1C"]))
    print("problems           : %d" % len(problems))
    for p in problems[:40]:
        print("   -", p)


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
