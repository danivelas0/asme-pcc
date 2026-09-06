# -*- coding: utf-8 -*-
"""Body-of-the-Code extraction: Chapters I-X of ASME B31.3-2024.

The appendices are tables, lists and worked examples. The body is the Code itself - numbered
requirements in a strict paragraph hierarchy, with tables and figures hung off them - so the
engines carry over but the assembly does not. Three things here are not in the appendix
range at all, and each is handled below rather than in the shared modules, so that the
appendix extraction keeps producing byte-identical output:

1. Landscape sheets carry *tables*, not only figures (Table 302.3.5-1, Table 341.3.2-1).
   They come in through rotate.UprightPage.
2. Figures are embedded raster images on 20 pages, at up to 3,892 px across. Rendering
   those at the appendices' 300 dpi would throw away more than half the resolution the
   source actually holds, so the render dpi is raised to match the image.
3. ASME sets in-body group banners in the same bold face as the column labels, so the
   appendix rule for the head/body split ("first rule below the last bold label") walks
   into the data. See body_header_rule.
"""
import os
import re

import pymupdf

import figures as F
import prose as P
import tables as T
from b313 import clean, num, page_lines
from rotate import upright

CONTD = re.compile(r"\s*\(Cont[’']d\)\s*$", re.I)
# "300", "A302.3.2", "K323.3.5", "U335.8" - the paragraph numbers the Code is cited by.
PARA_ID = re.compile(r"^([A-Z]{0,2}\d{3}(?:\.\d+)*)\b")
PART_HEAD = re.compile(r"^PART\s+(\d+)\b\s*(.*)$")
CHAPTER_HEAD = re.compile(r"^Chapter\s+([IVX]+)\b\s*(.*)$", re.I)
# A table caption is normally "Table 330.1.1-1", but the acceptance-criteria tables define
# their value symbols on a companion sheet captioned "Criterion Value Notes for Table
# K341.3.2-1" - a ruled table in its own right, and one the Code's own List of Tables names.
TABLE_CAP = re.compile(r"^(?:Table\s+[A-Z]?[0-9A-Za-z.\-]+"
                       r"|Criterion Value Notes for Table\s+[A-Z]?[0-9A-Za-z.\-]+)")


# --------------------------------------------------------------------------- tables
def find_tables(page, lines):
    """Table regions on a body page, with regions that span a section heading discarded.

    T.find_tables treats any two horizontal rules on a page as the top and bottom of one
    table unless a caption falls between them. In the body that is not safe. Page 89 carries
    the closing rule of a table continued from the page before (y=101) and, 466 pt lower, the
    rule under an unruled inline list inside para. 323.3.2 (y=567), with two columns of
    ordinary prose between them. The two got glued into a single region covering the whole
    left column, the prose pass was told to skip it, and para. 323.3 - a numbered requirement
    of the Code - disappeared from the output entirely.

    A table never has a numbered section heading inside it, so a Ronnia-Bold paragraph
    heading between two rules proves they belong to different things. The rules are split
    there, and a fragment left with fewer than two rules is not a table at all.
    """
    regs = T.find_tables(page, lines)
    heads = [_cy(L) for L in lines
             if L["font"].startswith("Ronnia-Bold") and PARA_ID.match(L["text"])]
    if not heads:
        return regs
    out = []
    for reg in regs:
        cuts = [h for h in heads if reg["y_top"] < h < reg["y_bot"]]
        if not cuts:
            out.append(reg)
            continue
        edges = [reg["y_top"] - 1] + sorted(cuts) + [reg["y_bot"] + 1]
        for lo, hi in zip(edges, edges[1:]):
            rules = [y for y in reg["rules"] if lo < y < hi]
            if len(rules) < 2:
                continue
            out.append(dict(reg, rules=rules, y_top=rules[0], y_bot=rules[-1]))
    return out


def body_header_rule(region, lines):
    """Where a body table's head ends and its data begins.

    The appendix rule - the first rule below the last bold label - cannot be used here.
    Table 326.1.1-1 and Table A326.1-1 group their entries under bold banners ("Bolting",
    "Metallic Fittings, Valves, and Flanges") set in the same face as the column labels, so
    "the last bold label" sits deep inside the data and seven real rows are silently
    reclassified as header.

    What separates head from body is not the last bold line but the first non-bold one: the
    head is bold all the way across, the data never is. So the split is the last rule that
    still has nothing but bold above it - and then, because ASME rules above and below each
    group banner too, a trailing head row that labels a single column other than the first
    is given back to the body, where group_banners picks it up as grouping.
    """
    inside = _in_region(region, lines)
    best = None
    for y in region["rules"][1:]:
        above = [L for L in inside if _cy(L) < y]
        if not above:
            continue
        if all(_bold(L) for L in above):
            best = y
        else:
            break
    if best is None:
        return T._header_rule(region)
    # Walk the split back up over any group banner it swallowed.
    rules = [y for y in region["rules"] if y <= best]
    while len(rules) >= 2:
        band = [L for L in inside if rules[-2] < _cy(L) < rules[-1]]
        cols = {_col_of(L, region["col_bounds"]) for L in band}
        if band and len(cols) == 1 and 0 not in cols:
            rules.pop()                      # that tier labels one column: a group banner
            best = rules[-1]
        else:
            break
    return best


def _cy(L):
    return (L["bbox"][1] + L["bbox"][3]) / 2


def _bold(L):
    return L["font"].endswith(("Bold", "Bold-Italic")) or L["font"].startswith("Ronnia")


def _col_of(L, bounds):
    return T._colof(L["bbox"][0], L["bbox"][2], bounds)


def _in_region(region, lines):
    return [L for L in lines
            if region["y_top"] - 1 <= _cy(L) <= region["y_bot"] + 1
            and L["bbox"][2] > region["x0"] - 4 and L["bbox"][0] < region["x1"] + 4]


def body_rows(region, lines, hr, bounds, ytol=4.0):
    """Data rows for a table region, with ASME's bold group banners kept as grouping.

    Table 326.1.1-1 sorts its component standards under bold banners - "Bolting", "Metallic
    Fittings, Valves, and Flanges", "Gaskets" - printed as rows inside the data. They are
    not column labels and they are not entries either, and letting them through as rows is
    actively harmful: a banner row leaves the key column empty, so the wrapped-continuation
    merge would fold "Metallic Fittings, Valves, and Flanges" into the title of the standard
    printed above it. They are recorded on the rows they head instead.
    """
    from b313 import words_with_x
    items = [L for L in _in_region(region, lines) if _cy(L) >= hr
             and not (L["bbox"][2] < bounds[0] - 4 or L["bbox"][0] > bounds[-1] + 4)]
    clusters = []
    for L in sorted(items, key=lambda L: (L["bbox"][1], L["bbox"][0])):
        if clusters and abs(_cy(L) - clusters[-1][0]) <= ytol:
            clusters[-1][1].append(L)
        else:
            clusters.append([_cy(L), [L]])

    rows, group = [], None
    for _, ls in clusters:
        cols = {_col_of(L, bounds) for L in ls}
        if all(_bold(L) for L in ls) and len(cols) == 1 and 0 not in cols:
            group = CONTD.sub("", clean(
                " ".join(L["text"] for L in sorted(ls, key=lambda L: L["bbox"][0])))).strip()
            continue
        cells = [""] * (len(bounds) - 1)
        runs = []
        for L in ls:
            runs.extend(words_with_x(L, bounds))
        for text, x0, x1 in sorted(runs, key=lambda r: r[1]):
            c = T._colof(x0, x1, bounds)
            cells[c] = (cells[c] + " " + text).strip() if cells[c] else text
        first = [L["bbox"][0] for L in ls if _col_of(L, bounds) == 0]
        rows.append({"cells": [clean(c) for c in cells], "group": group,
                     "x0": min(first) if first else bounds[0]})
    return rows


# A header cell wrapped mid-number: "(1,00-" over "0)" is one temperature, 1,000.
_WRAP = re.compile(r"[-‐]$")


def head_columns(region, lines, hr, bounds):
    """Column names for a body table, built from where each head run actually sits.

    Stacking the head row by row and column by column is what the appendix tables needed,
    but it mis-assigns every spanning label here. ASME writes "Greater Material Thickness"
    once, centred over the mm and in. columns, and "Component Temperature, Ti, deg C (deg F)"
    once, centred over sixteen of them. Placed by column index, that text lands on whichever
    single column happens to contain its midpoint - Table 302.3.5-1 came out with a column
    called "Component 593 (1,100)" and another called "Temperature, Ti, °C (°F) 649 (1,200)",
    while the fourteen other temperature columns carried none of it.

    So each run is assigned to every column it physically overlaps: a run over one column is
    that column's own label, a run over several is a banner qualifying all of them. A banner
    too long for its span wraps onto a second line, and the two lines cover different column
    sets ("Category D Fluid" over two columns, "Service [Note (5)]" over the next two), so a
    run is first merged with a banner directly above it on the same centre.

    Header cells are not split at column gaps the way data rows are. Each printed header
    cell is its own text line here, and splitting broke the one banner that is set with wide
    internal spacing: "Component Temperature, Ti, deg C (deg F)" came apart into "Component"
    over two columns and "Temperature, Ti, deg C (deg F)" over three, leaving the other
    eleven temperature columns unqualified.

    Returns (labels, banners_per_column, banners_seen).
    """
    items = [L for L in _in_region(region, lines) if _cy(L) < hr
             and not (L["bbox"][2] < bounds[0] - 4 or L["bbox"][0] > bounds[-1] + 4)]
    ncol = len(bounds) - 1

    clusters = []
    for L in sorted(items, key=lambda L: (L["bbox"][1], L["bbox"][0])):
        if clusters and abs(_cy(L) - clusters[-1][0]) <= 4.0:
            clusters[-1][1].append(L)
        else:
            clusters.append([_cy(L), [L]])

    labels = [[] for _ in range(ncol)]
    spans = [[] for _ in range(ncol)]
    seen, prev = [], []
    for _, ls in clusters:
        runs = [(L["text"], L["bbox"][0], L["bbox"][2]) for L in ls]
        cur = []
        for text, x0, x1 in sorted(runs, key=lambda r: r[1]):
            hit = [i for i in range(ncol)
                   if min(x1, bounds[i + 1]) - max(x0, bounds[i]) > 1.0]
            if not hit:
                hit = [T._colof(x0, x1, bounds)]
            if len(hit) == 1:
                _append(labels[hit[0]], text)
                continue
            cx = (x0 + x1) / 2
            host = next((b for b in prev if abs(b["cx"] - cx) <= 45), None)
            if host is not None:                    # second line of a wrapped banner
                host["text"] = clean(host["text"] + " " + text)
                host["hit"] = sorted(set(host["hit"]) | set(hit))
                host["cx"] = cx
                cur.append(host)
            else:
                cur.append({"text": text, "hit": hit, "cx": cx})
        prev = cur
        for b in cur:
            if b not in seen:
                seen.append(b)
    for b in seen:
        for i in b["hit"]:
            _append(spans[i], b["text"])
    return ([clean(" ".join(c)) or None for c in labels],
            [list(c) for c in spans],
            [b["text"] for b in seen])


def _append(out, text):
    if out and _WRAP.search(out[-1]):
        out[-1] = out[-1][:-1] + text       # "(1,00-" over "0)" is one temperature, 1,000
    elif text not in out:
        out.append(text)


def collect_tables(doc, pages, dehyph=None):
    """Every ruled table over a page range, joined across its continuation pages."""
    found, order = {}, []
    carried = {}
    for i in pages:
        pg = upright(doc[i])
        lines = page_lines(pg)
        cont = continued_notes(pg, lines)
        if cont:
            carried.setdefault(cont[0], []).extend(
                [dehyph(n) for n in cont[1]] if dehyph else cont[1])
        regs = find_tables(pg, lines)
        for n, reg in enumerate(regs):
            label, title = T.caption_above(lines, reg["y_top"], x0=reg["x0"], x1=reg["x1"])
            # Line art carries rules of its own, so a ruled region is a table only when it
            # is captioned as one.
            if not label or not TABLE_CAP.match(label):
                continue
            label = CONTD.sub("", label).strip()
            title = CONTD.sub("", title).strip() if title else None
            hr = body_header_rule(reg, lines)
            _, _, _, bounds = T.header_and_body(pg, reg, lines, header_rule=hr)
            head, spans, banners = head_columns(reg, lines, hr, bounds)
            body = body_rows(reg, lines, hr, bounds)
            nxt = regs[n + 1]["y_top"] if n + 1 < len(regs) else None
            notes = _table_notes(reg, lines, nxt, dehyph)
            if label not in found:
                found[label] = {"label": label, "title": title, "pages": [], "head": [],
                                "spans": [], "banners": [], "body": [], "notes": [],
                                "ncols": len(bounds) - 1}
                order.append(label)
            e = found[label]
            e["pages"].append(i + 1)
            if title and not e["title"]:
                e["title"] = title
            if not e["head"]:
                e["head"], e["spans"] = head, spans
                e["ncols"] = len(bounds) - 1
            for b in banners:
                if b not in e["banners"]:
                    e["banners"].append(b)
            e["body"].extend(body)
            for nt in notes:
                if nt not in e["notes"]:
                    e["notes"].append(nt)
    for label, notes in carried.items():
        if label in found:
            for nt in notes:
                if nt not in found[label]["notes"]:
                    found[label]["notes"].append(nt)
    return [found[k] for k in order]


NOTE_START = re.compile(r"^(GENERAL NOTES?:|NOTES?:|Legend:)", re.I)
NOTE_ITEM = re.compile(r"^\((?:[a-z]|\d{1,2})\)\s")


def _table_notes(region, lines, next_top=None, dehyph=None):
    lo, hi = region["y_bot"], next_top if next_top else 10_000
    got, started = [], False
    for L in sorted(lines, key=lambda L: (L["bbox"][1], L["bbox"][0])):
        cy = (L["bbox"][1] + L["bbox"][3]) / 2
        if not (lo < cy < hi):
            continue
        if L["bbox"][2] < region["x0"] - 6 or L["bbox"][0] > region["x1"] + 6:
            continue
        if started and L["font"].startswith("Ronnia"):
            break
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
    return [g for g in got if g and not g.rstrip().endswith(":")]


def merge_rows(rows):
    """Fold wrapped continuation lines back into the row above them.

    Two shapes occur. Usually a continuation leaves the key column empty and fills only the
    cell whose text wrapped. But a long entry in the key column wraps in place, and ASME
    hangs those continuation lines by one em: in Table 302.3.5-1 every row opens at x=66.9
    and every continuation sits at x=74.9. That indent is what separates them, not the text.
    Judging by the text instead is wrong in both directions - "N088xx and N066xx nickel" is
    a continuation though it opens with a capital, and "Other materials [Note (9)]" is a row
    in its own right though it carries no values at all.
    """
    out = []
    for r in rows:
        cells, filled = r["cells"], [i for i, c in enumerate(r["cells"]) if c]
        if out and filled == [0] and cells[0] and r["x0"] > out[-1]["x0"] + 2.5:
            out[-1]["cells"][0] = clean(out[-1]["cells"][0] + " " + cells[0])
        elif out and not cells[0] and filled and len(filled) <= 2:
            for i in filled:
                out[-1]["cells"][i] = clean((out[-1]["cells"][i] + " " + cells[i]).strip())
        else:
            out.append({"cells": list(cells), "group": r["group"], "x0": r["x0"]})
    return out


# --------------------------------------------------------------------------- figures
MAX_DPI = 600


def _art_rect(page, caption, table_regions=()):
    """Extent of a figure's artwork, whether it is drawn or placed.

    The appendix range has no raster images at all, so figures there are found purely from
    the drawing list. Twenty body pages instead carry the figure as a single embedded image,
    and on those pages get_drawings() is empty, so a drawings-only search finds no figure and
    silently skips it. Both sources are unioned here.

    Clustering has to be two-dimensional. Growing the artwork by vertical proximity alone is
    safe on a page that is nothing but a figure, which is all the appendices contain, but the
    body sets column-width figures beside running text: on page 57 the miter-bend image
    occupies x 71-271 in the left column, and vertical clustering pulled in vector strokes
    from the right column at the same height, stretching the clip to x 518. The prose pass
    was then told to skip both columns and the requirements of para. 304.2.3 vanished with
    the figure. Boxes therefore join a cluster only when they are close in x as well as y.
    """
    head_y, foot_y = _band(page)
    boxes = []
    for dr in page.get_drawings():
        b = dr["rect"]
        if b.is_empty or b.is_infinite or b.y1 < head_y or b.y0 > foot_y:
            continue
        if b.height < 2.5 and b.width > 20 and any(
                reg["y_top"] - 2 <= b.y0 <= reg["y_bot"] + 2 for reg in table_regions):
            continue
        if b.height < 3.0 and b.width < 20.0:
            continue        # a minus sign or fraction bar in an equation, not line art
        boxes.append(pymupdf.Rect(b))
    for im in page.get_image_info():
        b = pymupdf.Rect(im["bbox"])
        if b.is_empty or b.y1 < head_y or b.y0 > foot_y:
            continue
        boxes.append(b)
    if not boxes:
        return pymupdf.Rect()

    clusters = []
    for b in sorted(boxes, key=lambda b: (b.y0, b.x0)):
        joined = None
        for c in clusters:
            if _near(b, c):
                c |= b
                joined = c
                break
        if joined is None:
            clusters.append(pymupdf.Rect(b))
        else:                       # one box can bridge two clusters
            merged, rest = joined, []
            for c in clusters:
                if c is joined:
                    continue
                if _near(c, merged):
                    merged |= c
                else:
                    rest.append(c)
            clusters = rest + [merged]

    cap = pymupdf.Rect(*caption["bbox"])
    below = [c for c in clusters if c.y1 > cap.y1]
    if not below:
        return pymupdf.Rect()
    # the artwork a caption belongs to is the nearest cluster under it that it sits over
    over = [c for c in below if min(c.x1, cap.x1) - max(c.x0, cap.x0) > -24]
    keep = min(over or below, key=lambda c: c.y0)
    grown = True
    while grown:
        grown = False
        for c in clusters:
            if c is keep or keep.contains(c):
                continue
            if (min(c.x1, keep.x1) - max(c.x0, keep.x0) > 0
                    and -60 <= c.y0 - keep.y1 <= 60):
                keep |= c
                grown = True
    return keep


def _near(a, b, xgap=20.0, ygap=25.0):
    return (min(a.x1, b.x1) - max(a.x0, b.x0) > -xgap
            and min(a.y1, b.y1) - max(a.y0, b.y0) > -ygap)


def _band(page):
    box = getattr(page, "content_box", None)
    if box is None:
        from b313 import HEAD_Y, FOOT_Y
        return HEAD_Y, FOOT_Y
    return 55.0, page.rect.height - 55.0


def figure_rect(page, caption, table_regions=()):
    """Clip rect for one figure: its artwork, the labels inside it, and its caption."""
    art = _art_rect(page, caption, table_regions)
    if art.is_empty:
        return None
    head_y, foot_y = _band(page)
    band = pymupdf.Rect(0, head_y, page.rect.x1, foot_y)
    art &= band
    box = pymupdf.Rect(art)
    lines = page_lines(page)
    for L in lines:
        lb = pymupdf.Rect(*L["bbox"])
        if art.intersects(lb) or (art.y0 - 4 <= lb.y0 and lb.y1 <= art.y1 + 4
                                  and art.x0 - 8 <= lb.x0 and lb.x1 <= art.x1 + 8):
            box |= lb
    # Part labels - "(c) Subsurface Flaw" - sit just under the last drawing and would
    # otherwise be cropped off. Only short lines qualify, so body prose is never pulled in.
    for L in lines:
        lb = pymupdf.Rect(*L["bbox"])
        if (0 <= lb.y0 - art.y1 <= 32 and lb.width < 250
                and lb.x1 > art.x0 - 20 and lb.x0 < art.x1 + 20):
            box |= lb
    box |= pymupdf.Rect(*caption["bbox"])
    box &= band
    box.x0 = max(box.x0 - 6, 0); box.y0 = max(box.y0 - 4, head_y)
    box.x1 = min(box.x1 + 6, page.rect.x1); box.y1 = min(box.y1 + 4, foot_y)
    return box


def render_dpi(page, rect):
    """Enough dpi to keep everything the source actually holds, and no more.

    A placed image carries a fixed number of pixels. Figure 328.4.2-1 is 3,892 px across a
    468 pt box, which is 599 dpi; rendering it at the appendices' 300 dpi would throw away
    more than half of it. So the dpi is raised to the densest image under the clip, and
    capped, because nothing is gained by resampling a 600 dpi scan higher.
    """
    best = F.DPI
    for im in page.get_image_info():
        b = pymupdf.Rect(im["bbox"])
        if b.is_empty or not rect.intersects(b) or b.width <= 0:
            continue
        best = max(best, 72.0 * im["width"] / b.width)
    return int(min(round(best), MAX_DPI))


def figure_notes(page, rect, lines, dehyph=None):
    """The GENERAL NOTE and numbered notes printed with a figure.

    These carry real requirements - Figure 304.3.3-1's notes say when a reinforcing saddle
    may be used and how A4 is counted - and they are set as ordinary prose, so the prose pass
    would take them for running text if the figure area were not excluded, and nothing at all
    would hold them once it is. They are read here and kept with the figure.
    """
    got = []
    for L in sorted(lines, key=lambda L: (L["bbox"][1], L["bbox"][0])):
        cy = (L["bbox"][1] + L["bbox"][3]) / 2
        if not (rect.y0 - 2 <= cy <= rect.y1 + 2):
            continue
        if L["bbox"][2] < rect.x0 - 4 or L["bbox"][0] > rect.x1 + 4:
            continue
        if L["font"].startswith("Ronnia"):
            continue                         # the caption
        t = L["text"]
        if NOTE_START.match(t) or NOTE_ITEM.match(t):
            got.append(t)
        elif got:
            got[-1] = clean(got[-1] + " " + t)
    if dehyph:
        got = [dehyph(g) for g in got]
    return [g for g in got if g and not g.rstrip().endswith(":")]


def collect_figures(doc, pages, out_dir, prefix="", dehyph=None):
    """Render every technical figure over a page range; returns one record per figure.

    Figure 304.3.3-1 and Figure 304.3.4-1 are each printed over two sheets, the second
    captioned "(Cont'd)". They are one figure, so the sheets are collected under one record
    and listed in figure_images rather than overwriting each other's file.
    """
    out, seen = [], {}
    for i in pages:
        pg = upright(doc[i])
        lines = page_lines(pg)
        caps = F.figure_captions(pg, lines)
        if not caps:
            continue
        regs = find_tables(pg, lines)
        for cap in caps:
            label, title = F.caption_text(pg, cap, lines)
            label = CONTD.sub("", label).strip()
            contd = bool(title and CONTD.search(title))
            title = CONTD.sub("", title).strip() if title else None
            rect = figure_rect(pg, cap, regs)
            if rect is None or rect.is_empty:
                continue
            if getattr(pg, "source_rotation", 0):
                # A landscape sheet carries one figure and nothing else, and its callouts
                # sit well outside the artwork, so the clip is the whole content band.
                for L in lines:
                    rect |= pymupdf.Rect(*L["bbox"])
                rect.x0 = max(rect.x0 - 6, 0); rect.y0 = max(rect.y0 - 6, 0)
                rect.x1 = min(rect.x1 + 6, pg.rect.x1)
                rect.y1 = min(rect.y1 + 6, pg.rect.y1)
            dpi = render_dpi(pg, rect)
            stem = (prefix + "_" if prefix else "") + slug(label)
            rec = seen.get(label)
            n = len(rec["_sheets"]) + 1 if rec else 1
            fname = "%s%s.png" % (stem, "" if n == 1 else "_sheet_%d" % n)
            mat = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
            pix = pg.get_pixmap(clip=rect, matrix=mat, colorspace=pymupdf.csRGB, alpha=False)
            pix.save(os.path.join(out_dir, fname))
            sheet = {"file": fname, "pdf_page": i + 1, "continuation": contd,
                     "_rect": (rect.x0, rect.y0, rect.x1, rect.y1),
                     "_notes": figure_notes(pg, rect, lines, dehyph),
                     "width_px": pix.width, "height_px": pix.height, "dpi": dpi}
            if rec is None:
                rec = {"figure_id": label, "title": title, "pdf_page": i + 1,
                       "figure_image": fname,
                       "landscape": bool(getattr(pg, "source_rotation", 0)),
                       "_sheets": [sheet]}
                seen[label] = rec
                out.append(rec)
            else:
                rec["_sheets"].append(sheet)
                if title and not rec["title"]:
                    rec["title"] = title
    for rec in out:
        # _sheets is kept on every figure, single-sheet ones included: it carries the clip
        # rect the prose pass needs in order to skip the figure. Dropping it for single-sheet
        # figures left all 39 of them un-excluded, so their captions and internal callouts
        # came through as stray sections of the Code.
        rec["sheets"] = len(rec["_sheets"])
        notes = []
        for sh in rec["_sheets"]:
            for n in sh.pop("_notes", []):
                if n not in notes:
                    notes.append(n)
        rec["notes"] = notes
        if rec["sheets"] > 1:
            rec["figure_images"] = list(rec["_sheets"])
    return out


def slug(s):
    return re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")


# --------------------------------------------------------------------------- prose
def split_runin(line):
    """Split a line at the bold/roman boundary of an ASME run-in heading.

    "326.1.1 Listed Piping Components. Dimensional stan-" is one printed line carrying two
    different things, and the heading is the longer half, so a whole-line font vote calls
    the entire line a heading. The paragraph then begins "dards for piping components ...".
    Splitting at the font boundary gives the heading and the text back intact.

    Only a bold run at the very start of the line counts. Bold appearing later is emphasis
    inside a sentence, not a heading.
    """
    from b313 import _join
    chars = line["chars"]
    if not chars or not chars[0].get("font", "").startswith("Ronnia-Bold"):
        return [line]
    n = 0
    while n < len(chars) and (chars[n].get("font", "").startswith("Ronnia-Bold")
                              or chars[n]["c"] == " "):
        n += 1
    while n and chars[n - 1]["c"] == " ":
        n -= 1
    if n == 0 or n >= len(chars):
        return [line]
    parts = []
    for seg, font in ((chars[:n], chars[0]["font"]), (chars[n:], chars[n]["font"])):
        while seg and seg[0]["c"] == " ":
            seg = seg[1:]
        if not seg:
            continue
        text = clean(_join(seg, seg[0].get("size", 9.6)))
        if not text:
            continue
        d = dict(line)
        d["text"] = text
        d["chars"] = seg
        d["font"] = font
        d["size"] = round(seg[0].get("size", line["size"]), 1)
        d["bbox"] = (seg[0]["bbox"][0], line["bbox"][1], seg[-1]["bbox"][2], line["bbox"][3])
        d["full_width"] = (d["bbox"][2] - d["bbox"][0]) > 300
        parts.append(d)
    return parts or [line]


def page_blocks(page, exclude=()):
    """Ordered content blocks for one body page, with run-in headings separated."""
    lines = page_lines(page)
    lines = [L for L in lines if not P.in_regions(L, exclude)]
    split = []
    for L in lines:
        split.extend(split_runin(L))
    split = P.merge_heading_fragments(split)
    base = P.column_base(split)
    strokes = [(d["rect"].x0, d["rect"].y0, d["rect"].x1, d["rect"].y1)
               for d in page.get_drawings()
               if d["rect"].height < 2.0 and 3.0 < d["rect"].width < 12.0]
    split = P._group_equations(split, base, strokes)
    blocks = []
    for L in P.reading_order(split):
        b = base.get(L["col"], 72.0)
        blocks.append({
            "text": L["text"], "heading": L["font"].startswith("Ronnia-Bold"),
            "indent": max(0, int(round((L["bbox"][0] - b) / P.EM))),
            "equation": P._is_equation(L), "size": L["size"],
            "y": L["bbox"][1], "col": L["col"],
            "appendix_title": L["size"] > 14,   # the chapter title, in assemble()'s terms
        })
    return blocks


# --------------------------------------------------------------------------- definitions
def definitions(doc, first=35, last=42, dehyph=None):
    """Para. 300.2 Definitions as a glossary.

    The Code's terms are set in italic and closed with a colon, the definition following in
    roman on the same line, so an entry is recognised by an italic run at the very start of
    a line. Italic appearing later in a line is a cross-reference inside a definition ("See
    also oxygen-arc cutting.") and must not open a new entry.
    """
    entries, cur, started = [], None, False
    for i in range(first, last + 1):
        for L in page_lines(doc[i]):
            t = L["text"]
            if not started:
                if L["font"].startswith("Ronnia") and t.strip() == "Definitions":
                    started = True
                continue
            if L["font"].startswith("Ronnia"):
                if PARA_ID.match(t) and not t.startswith("300.2"):
                    started = False           # 300.3 Nomenclature: the glossary has ended
                    break
                continue
            chars = L["chars"]
            head = _italic_head(chars)
            if head and ":" in t[:len(head) + 3]:
                term, _, rest = t.partition(":")
                cur = {"term": clean(term), "definition": clean(rest)}
                entries.append(cur)
            elif cur is not None:
                cur["definition"] = clean(cur["definition"] + " " + t)
    for e in entries:
        if dehyph:
            e["definition"] = dehyph(e["definition"])
            e["term"] = dehyph(e["term"])
        m = re.match(r"^(.*?)\s*\(([^()]{1,12})\)$", e["term"])
        e["acronym"] = m.group(2) if m else None
        if m:
            e["term"] = clean(m.group(1))
        e["see_also"] = re.findall(r"[Ss]ee(?: also)? ([^.;]+)", e["definition"])
    return entries


def _italic_head(chars):
    """The leading run of italic characters on a line, if it starts with one."""
    if not chars or "Italic" not in chars[0].get("font", ""):
        return ""
    n = 0
    while n < len(chars) and ("Italic" in chars[n].get("font", "") or chars[n]["c"] == " "):
        n += 1
    return "".join(c["c"] for c in chars[:n]).strip()


# --------------------------------------------------------------------------- assembly
def excluded_regions(doc, pages, figs=()):
    """Page areas the prose pass must skip: every table with its caption and notes, and
    every figure with its caption.

    The table's own rules bound only its grid. Its caption sits above the first rule and its
    NOTES below the last one, and both are set in the same faces as the surrounding prose,
    so leaving them out of the exclusion drops a table's notes into the running text of
    whatever paragraph happens to follow.
    """
    by_page = {}
    for i in pages:
        pg = upright(doc[i])
        lines = page_lines(pg)
        out = []
        for reg in find_tables(pg, lines):
            top = reg["y_top"]
            caps = [L["bbox"][1] for L in lines
                    if L["font"].startswith("Ronnia") and top - 70 < L["bbox"][3] <= top + 1]
            out.append({"x0": reg["x0"] - 6, "x1": reg["x1"] + 6,
                        "y_top": (min(caps) if caps else top) - 4,
                        "y_bot": _notes_extent(reg, lines)})
        cont = continued_notes(pg, lines)
        if cont:
            out.append({"x0": -1e9, "x1": 1e9, "y_top": 55.0, "y_bot": cont[2] + 2})
        by_page[i] = out
    for f in figs:
        for sh in f.get("_sheets", ()):
            r = sh.get("_rect")
            if r is None:
                continue
            by_page.setdefault(sh["pdf_page"] - 1, []).append(
                {"x0": r[0], "x1": r[2], "y_top": r[1], "y_bot": r[3]})
    return by_page


def _notes_extent(region, lines):
    """How far below its last rule a table's NOTES run."""
    y = region["y_bot"]
    started = False
    for L in sorted(lines, key=lambda L: L["bbox"][1]):
        cy = _cy(L)
        if cy <= region["y_bot"] or L["bbox"][2] < region["x0"] - 6 or L["bbox"][0] > region["x1"] + 6:
            continue
        if L["font"].startswith("Ronnia"):
            break                     # a heading: prose has resumed
        if cy - y > 20:
            break        # notes are set tight under the table; a gap that big ends them
        if NOTE_START.match(L["text"]) or started:
            started = True
        y = max(y, L["bbox"][3])
    return y + 2


def collect_prose(doc, pages, dehyph, exclude):
    """Flat section list for a page range, table and figure areas removed."""
    blocks = []
    for i in pages:
        pg = upright(doc[i])
        for b in page_blocks(pg, exclude=exclude.get(i, [])):
            blocks.append((b, i + 1))
    pre, secs = P.assemble(blocks, None)
    for s in secs:
        # headings break across lines just as paragraphs do - "Reinforcement of Welded
        # Branch Connec- tions" - so they need the same treatment
        if s.get("heading"):
            s["heading"] = dehyph(s["heading"])
        for par in s["paragraphs"]:
            par["text"] = dehyph(par["text"])
    for par in pre:
        par["text"] = dehyph(par["text"])
    return pre, secs


def build_tree(sections):
    """Nest the flat section list on the printed paragraph numbers.

    "304", "304.1", "304.1.1" is the hierarchy an engineer cites, and it is carried in the
    numbers themselves, so the tree is built from them rather than from heading sizes -
    which vary with how the printer set a given heading. PART headings open a container of
    their own; chapters without PARTs put their sections at the top level.
    """
    root, part, open_ = [], None, []
    for s in sections:
        head = s.get("heading") or ""
        m = PART_HEAD.match(head)
        if not s["id"] and m:
            part = {"kind": "part", "part": "PART %s" % m.group(1),
                    "title": clean(m.group(2)) or None, "page": s["page"], "sections": []}
            root.append(part)
            open_ = []
            continue
        node = {"kind": "section", "id": s["id"], "heading": head or None,
                "page": s["page"], "paragraphs": s["paragraphs"], "sections": []}
        if s["id"]:
            while open_ and not s["id"].startswith(open_[-1][0] + "."):
                open_.pop()
        host = open_[-1][1]["sections"] if open_ else (part["sections"] if part else root)
        host.append(node)
        if s["id"]:
            open_.append((s["id"], node))
    return root


def count_tree(nodes):
    """(sections, paragraphs, equations, words) over a section tree."""
    sec = par = eq = words = 0
    for n in nodes:
        if n["kind"] == "section":
            sec += 1
            for p in n["paragraphs"]:
                if p.get("kind") == "equation":
                    eq += 1
                else:
                    par += 1
                    words += len(p["text"].split())
        a, b, c, w = count_tree(n["sections"])
        sec += a; par += b; eq += c; words += w
    return sec, par, eq, words


def continued_notes(page, lines):
    """A table's NOTES carried over onto the next sheet, and the band they occupy.

    Tables 302.3.5-1, 323.2.2-1 and 341.3.2-1 run out of room and finish their notes on the
    following page, under a repeated caption and a "NOTES: (Cont'd)" line. There are no
    rules on that page - they were all on the previous one - so the table engine sees
    nothing, and the notes fell through to the prose pass, where they arrived as the single
    fragment "'d) NOTES: (Cont'd)" and the note text itself was scattered.

    The notes are set full measure at the head of the page and ordinary two-column prose
    resumes below them, so the block ends at the first baseline carrying two columns.

    Returns (table_id, notes, y_bottom) or None.
    """
    caps = [L for L in lines if L["font"].startswith("Ronnia") and CONTD.search(L["text"])]
    if not caps:
        return None
    ids = []
    for L in lines:
        if not L["font"].startswith("Ronnia"):
            continue
        m = re.search(r"Table\s+([A-Z]?[0-9A-Za-z.\-]+)", L["text"])
        if m:
            ids.append("Table " + m.group(1).rstrip("."))
    if not ids:
        return None
    start = next((L for L in lines if NOTE_START.match(L["text"])), None)
    if start is None:
        return None

    rest = [L for L in lines if L["bbox"][1] > start["bbox"][1] - 0.5
            and not L["font"].startswith("Ronnia")]
    clusters = []
    for L in sorted(rest, key=lambda L: (round(L["bbox"][1], 1), L["bbox"][0])):
        if clusters and abs(L["bbox"][1] - clusters[-1][0]) <= 3.0:
            clusters[-1][1].append(L)
        else:
            clusters.append([L["bbox"][1], [L]])
    got, y_bot = [], start["bbox"][3]
    for y, ls in clusters:
        if len(ls) > 1:
            break                    # two columns on one baseline: the prose has resumed
        t = ls[0]["text"]
        if NOTE_START.match(t):
            y_bot = ls[0]["bbox"][3]
            continue
        if NOTE_ITEM.match(t):
            got.append(t)
        elif got:
            got[-1] = clean(got[-1] + " " + t)
        else:
            continue
        y_bot = ls[0]["bbox"][3]
    if not got:
        return None
    return ids[0], got, y_bot
