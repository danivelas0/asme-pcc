# -*- coding: utf-8 -*-
"""Generic per-appendix assembly: tables, figures, prose."""
import re
import pymupdf
from b313 import page_lines, clean, num, SOURCE_BLOCK
import tables as T
import figures as F
import prose as P

CONTD = re.compile(r"\s*\(Cont[\u2019']d\)\s*$", re.I)
NOTE_START = re.compile(r"^(GENERAL NOTES?:|NOTES?:|Legend:)", re.I)


def _colnames(header_rows, ncols):
    """Stack the header rows into one name per column."""
    names = []
    for c in range(ncols):
        parts = [r[c] for r in header_rows if c < len(r) and r[c]]
        seen, uniq = set(), []
        for p in parts:
            if p not in seen:
                seen.add(p); uniq.append(p)
        names.append(clean(" ".join(uniq)) or None)
    return names


GREEK = {"Γ": "gamma", "α": "alpha", "β": "beta", "γ": "gamma",
         "Δ": "delta", "δ": "delta", "ε": "epsilon", "θ": "theta",
         "λ": "lambda", "μ": "mu", "ν": "nu", "ρ": "rho",
         "σ": "sigma", "Σ": "sigma", "τ": "tau", "φ": "phi",
         "ω": "omega", "∑": "sum"}


def _key(name, i):
    if not name:
        return "column_%d" % (i + 1)
    k = name.lower()
    for g, w in GREEK.items():
        k = k.replace(g, w).replace(g.lower(), w)
    k = re.sub(r"\[note[^\]]*\]", "", k)
    k = k.replace("\u00b0", "deg_").replace("%", "pct")
    k = re.sub(r"[^a-z0-9]+", "_", k).strip("_")
    return k or "column_%d" % (i + 1)


def _uniq(keys):
    out, seen = [], {}
    for k in keys:
        if k in seen:
            seen[k] += 1
            k = "%s_%d" % (k, seen[k])
        else:
            seen[k] = 1
        out.append(k)
    return out


def merge_wrapped(rows, key_col=0):
    """Fold a wrapped continuation line back into the row above it.

    A continuation carries no value in the key column and only fills the one cell whose text
    wrapped, so it is safe to recognise structurally rather than by looking at the text.
    """
    out = []
    for r in rows:
        filled = [i for i, c in enumerate(r) if c]
        if (out and filled == [key_col] and r[key_col]
                and (r[key_col].startswith("[") or r[key_col][:1].islower())):
            # a wrapped key cell, e.g. a material name whose "[Note (2)]" ran to a new line
            out[-1][key_col] = clean(out[-1][key_col] + " " + r[key_col])
            continue
        if out and not r[key_col] and filled and len(filled) <= 2:
            for i in filled:
                out[-1][i] = clean((out[-1][i] + " " + r[i]).strip())
        else:
            out.append(list(r))
    return out


NOTE_ITEM = re.compile(r"^\((?:[a-z]|\d{1,2})\)\s")


def table_notes(page, region, lines, next_top=None, dehyph=None):
    """Notes printed under a table, joined back into whole notes.

    The lines are collected in the table's own column band so a neighbouring column's prose
    cannot leak in, and each note runs until the next NOTES: heading or (a)/(1) marker.
    """
    lo, hi = region["y_bot"], next_top if next_top else 10_000
    got, started = [], False
    for L in sorted(lines, key=lambda L: (L["bbox"][1], L["bbox"][0])):
        cy = (L["bbox"][1] + L["bbox"][3]) / 2
        if not (lo < cy < hi):
            continue
        if L["bbox"][2] < region["x0"] - 6 or L["bbox"][0] > region["x1"] + 6:
            continue
        if started and L["font"].startswith("Ronnia"):
            break              # a section heading: the notes have ended and prose resumes
        t = L["text"]
        if NOTE_START.match(t):
            started = True
            got.append(t)
        elif started:
            if NOTE_ITEM.match(t):
                got.append(t)
            elif got:
                got[-1] = clean(got[-1] + " " + t)
    if dehyph:
        got = [dehyph(g) for g in got]
    # a bare "NOTES:" heading with nothing after it carries no content
    return [g for g in got if g and not g.rstrip().endswith(":")]


def collect_tables(doc, pages, drop=(), dehyph=None):
    """All ruled tables over a page range, joined across their continuation pages."""
    found = {}
    order = []
    for i in pages:
        pg = doc[i]
        lines = page_lines(pg)
        regs = T.find_tables(pg, lines)
        for n, reg in enumerate(regs):
            label, title = T.caption_above(lines, reg["y_top"], x0=reg["x0"], x1=reg["x1"])
            # Line-art figures carry internal rules of their own, so a ruled region is
            # only a table when the caption above it actually says "Table".
            if not label or not label.startswith("Table"):
                continue
            label = CONTD.sub("", label).strip()
            if label in drop:
                continue
            title = CONTD.sub("", title).strip() if title else None
            head, body = T.cells_in(pg, reg, lines)
            nxt = regs[n + 1]["y_top"] if n + 1 < len(regs) else None
            notes = table_notes(pg, reg, lines, nxt, dehyph)
            if label not in found:
                found[label] = {"label": label, "title": title, "pages": [], "head": [],
                                "body": [], "notes": [], "ncols": len(reg["col_bounds"]) - 1}
                order.append(label)
            e = found[label]
            e["pages"].append(i + 1)
            if title and not e["title"]:
                e["title"] = title
            if not e["head"]:
                e["head"] = head
            e["body"].extend(body)
            for nt in notes:
                if nt not in e["notes"]:
                    e["notes"].append(nt)
    return [found[k] for k in order]


def table_json(e, appendix):
    names = _colnames(e["head"], e["ncols"])
    keys = _uniq([_key(n, i) for i, n in enumerate(names)])
    rows = merge_wrapped(e["body"])
    rows = [r for r in rows if any(c for c in r)]
    data = []
    for r in rows:
        data.append({k: num(v) for k, v in zip(keys, r)})
    return {
        "source": SOURCE_BLOCK,
        "appendix": appendix,
        "table_id": e["label"],
        "title": e["title"],
        "pdf_pages": [min(e["pages"]), max(e["pages"])],
        "columns": [{"key": k, "header": n} for k, n in zip(keys, names)],
        "row_count": len(data),
        "notes": e["notes"] or [],
        "rows": data,
    }


def collect_figures(doc, pages, out_dir, prefix):
    """Render every technical figure in the range; returns metadata records."""
    import os
    out = []
    for i in pages:
        pg = doc[i]
        rot = pg.rotation
        if rot:
            pg.set_rotation(0)
        lines = page_lines(pg)
        caps = F.figure_captions(pg, lines)
        if not caps:
            continue
        regs = T.find_tables(pg, lines)
        for cap in caps:
            label, title = F.caption_text(pg, cap, lines)
            rect = F.figure_rect(pg, cap, regs)
            if rect is None or rect.is_empty:
                continue
            if rot:                        # landscape: take the whole content band
                for L in lines:
                    rect |= pymupdf.Rect(*L["bbox"])
            fname = "%s_%s.png" % (prefix, re.sub(r"[^A-Za-z0-9]+", "_", label).strip("_").lower())
            info = F.render(pg, rect, os.path.join(out_dir, fname), rotate=rot)
            out.append({"figure_id": label, "title": title, "pdf_page": i + 1,
                        "figure_image": fname, "rotated": bool(rot),
                        "width_px": info["width"], "height_px": info["height"],
                        "dpi": info["dpi"]})
    return out


def collect_prose(doc, pages, dehyph, table_regions_by_page, fig_pages=()):
    """Section tree for the narrative part of an appendix."""
    blocks = []
    for i in pages:
        if i in fig_pages:
            continue
        pg = doc[i]
        excl = table_regions_by_page.get(i, [])
        for b in P.page_blocks(pg, exclude=excl):
            blocks.append((b, i + 1))
    pre, secs = P.assemble(blocks, None)
    for s in secs:
        for p in s["paragraphs"]:
            p["text"] = dehyph(p["text"])
    for p in pre:
        p["text"] = dehyph(p["text"])
    return pre, secs
