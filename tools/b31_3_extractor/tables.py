# -*- coding: utf-8 -*-
"""Ruled-table extraction for B31.3 appendices.

ASME sets these tables with horizontal rules only, but the rules are emitted as one segment
per column, so the rule with the most segments hands us the exact column grid. That is far
safer than inferring columns from whitespace, which is what the A/B/C run had to do.
"""
import re
from collections import defaultdict
from b313 import clean, num, page_lines


def rule_groups(page, tol=1.2):
    """Horizontal rules grouped by y -> {y: [(x0,x1), ...]} sorted."""
    segs = defaultdict(list)
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.height < 2.5 and r.width > 20 and r.x1 > 0 and r.x0 < page.rect.x1:
            segs[round(r.y0, 1)].append((max(r.x0, 0.0), min(r.x1, page.rect.x1)))
    ys = sorted(segs)
    merged = {}
    for y in ys:
        hit = next((k for k in merged if abs(k - y) <= tol), None)
        if hit is None:
            merged[y] = list(segs[y])
        else:
            merged[hit].extend(segs[y])
    # Displayed equations draw their fraction bars as short horizontal lines. A table rule
    # always runs the width of the table, so anything spanning less than 100 pt is math.
    out = {}
    for y, v in sorted(merged.items()):
        v = sorted(v)
        if max(x[1] for x in v) - min(x[0] for x in v) >= 100.0:
            out[y] = v
    return out


def _xspan(segs):
    return min(s[0] for s in segs), max(s[1] for s in segs)

TABLE_LABEL = re.compile(r"^(Table\s+[A-Z]?[0-9A-Za-z.\-]+)(?![0-9A-Za-z])")


def table_labels(lines):
    """Ronnia-Bold 'Table X-n' captions on the page, with their y and x centre."""
    out = []
    for L in lines:
        if L["font"].startswith("Ronnia") and TABLE_LABEL.match(L["text"]):
            out.append((L["bbox"][1], (L["bbox"][0] + L["bbox"][2]) / 2, L["text"]))
    return sorted(out)


def _split_bands(segs, min_gap=12.0):
    """Side-by-side tables: a rule row broken by a gap wider than a column separator."""
    segs = sorted(segs)
    bands, cur = [], [segs[0]]
    for s in segs[1:]:
        if s[0] - cur[-1][1] > min_gap:
            bands.append(cur); cur = []
        cur.append(s)
    bands.append(cur)
    return bands


def _merge_bounds(bounds, tol=2.5):
    """Collapse rule ends that differ only by hairline width into one column boundary."""
    out = []
    for x in bounds:
        if out and x - out[-1] <= tol:
            out[-1] = (out[-1] + x) / 2
        else:
            out.append(x)
    return [round(x, 1) for x in out]


def find_tables(page, lines=None):
    """Table regions on a page.

    Stacked tables are separated by their printed captions (rule spacing alone cannot tell a
    tall table from two short ones); side-by-side tables are separated by gaps in the rules.
    """
    lines = lines if lines is not None else page_lines(page)
    rg = rule_groups(page)
    if not rg:
        return []
    caps = table_labels(lines)
    ys = sorted(rg)
    # cut points: the first rule at or below each caption after the first
    cuts = []
    for cy, cx, _ in caps:
        below = [y for y in ys if y > cy]
        if below:
            cuts.append(below[0])
    cuts = sorted(set(cuts))
    groups, cur = [], []
    for y in ys:
        if cur and y in cuts and y != cur[0]:
            groups.append(cur); cur = []
        cur.append(y)
    if cur:
        groups.append(cur)

    out = []
    for g in groups:
        if len(g) < 2:
            continue
        widest = max(g, key=lambda y: len(rg[y]))
        for band in _split_bands(rg[widest]):
            bx0, bx1 = band[0][0], band[-1][1]
            rules = [y for y in g
                     if any(min(s[1], bx1) - max(s[0], bx0) > 10 for s in rg[y])]
            if len(rules) < 2:
                continue
            bounds = _merge_bounds(sorted({round(x, 1) for s in band for x in s}))
            out.append({
                "y_top": rules[0], "y_bot": rules[-1], "x0": bx0, "x1": bx1,
                "rules": rules, "col_bounds": bounds,
            })
    return sorted(out, key=lambda r: (r["y_top"], r["x0"]))


def cells_in(page, region, lines=None, header_rule=None):
    """Rows of cells for a table region. Rows cluster by y; columns by the rule grid."""
    lines = lines if lines is not None else page_lines(page)
    bounds = region["col_bounds"]
    if len(bounds) < 2:
        return [], []
    items = []
    for L in lines:
        bx0, by0, bx1, by1 = L["bbox"]
        cy = (by0 + by1) / 2
        if not (region["y_top"] - 0.5 <= cy <= region["y_bot"] + 0.5):
            continue
        if bx1 < bounds[0] - 4 or bx0 > bounds[-1] + 4:
            continue
        items.append(L)
    bounds = _extend_bounds(bounds, items)
    hr = header_rule if header_rule is not None else header_rule_for(region, lines)
    head = [L for L in items if (L["bbox"][1] + L["bbox"][3]) / 2 < hr]
    body = [L for L in items if (L["bbox"][1] + L["bbox"][3]) / 2 >= hr]
    head, _ = _split_banners(head, bounds)
    return _rows(head, bounds), _rows(body, bounds)


def _split_banners(head, bounds):
    """Separate spanning header banners from the per-column labels.

    ASME runs a caption across the whole table head ("Allowable Stress, S, MPa, at Metal
    Temperature ... Not Exceeding"). It is real content, but it belongs to the table, not to
    any one column, and folding it into the column names corrupts every one it crosses.
    """
    widest = max((bounds[i + 1] - bounds[i]) for i in range(len(bounds) - 1))
    labels, banners = [], []
    for L in head:
        (banners if (L["bbox"][2] - L["bbox"][0]) > 1.8 * widest else labels).append(L)
    return labels, banners


def header_and_body(page, region, lines=None, header_rule=None):
    """Column-label rows, spanning banner text, and body rows."""
    lines = lines if lines is not None else page_lines(page)
    bounds = _extend_bounds(region["col_bounds"], [
        L for L in lines
        if region["y_top"] - 0.5 <= (L["bbox"][1] + L["bbox"][3]) / 2 <= region["y_bot"] + 0.5
        and not (L["bbox"][2] < region["col_bounds"][0] - 4
                 or L["bbox"][0] > region["col_bounds"][-1] + 4)])
    hr = header_rule if header_rule is not None else header_rule_for(region, lines)
    items = [L for L in lines
             if region["y_top"] - 0.5 <= (L["bbox"][1] + L["bbox"][3]) / 2 <= region["y_bot"] + 0.5
             and not (L["bbox"][2] < bounds[0] - 4 or L["bbox"][0] > bounds[-1] + 4)]
    head = [L for L in items if (L["bbox"][1] + L["bbox"][3]) / 2 < hr]
    body = [L for L in items if (L["bbox"][1] + L["bbox"][3]) / 2 >= hr]
    labels, banners = _split_banners(head, bounds)
    seen, btexts = set(), []
    for L in sorted(banners, key=lambda L: L["bbox"][1]):
        if L["text"] not in seen:
            seen.add(L["text"]); btexts.append(L["text"])
    return _rows(labels, bounds), btexts, _rows(body, bounds), bounds


def _header_rule(region):
    """The rule that closes the head and opens the body.

    These tables are ruled at the top, under each tier of the head, and at the foot, with no
    rules between data rows - so the last rule before the closing one is the head/body split,
    however many tiers the head has.
    """
    r = region["rules"]
    return r[-2] if len(r) >= 3 else r[0]


def header_rule_for(region, lines):
    """Where the head ends and the body begins.

    The head is set in bold and the body is not, so the split is the first rule below the
    last bold label. Rule counting alone cannot do it: heads run to one, two or three tiers
    depending on the table.
    """
    bold = [L["bbox"][3] for L in lines
            if L["font"].endswith("Bold") and not L["font"].startswith("Ronnia")
            and region["y_top"] - 1 <= (L["bbox"][1] + L["bbox"][3]) / 2 <= region["y_bot"] + 1
            and L["bbox"][2] > region["x0"] - 4 and L["bbox"][0] < region["x1"] + 4
            and (L["bbox"][2] - L["bbox"][0]) <= 1.8 * (region["x1"] - region["x0"])]
    if bold:
        below = [y for y in region["rules"] if y > max(bold) - 1.5]
        if below:
            return below[0]
    return _header_rule(region)


def _extend_bounds(bounds, items, min_hits=3):
    """Widen the rule grid to cover a column the rules do not reach.

    On several Table K-1C pages the printed rules start at the material column and the line
    numbers sit to their left, outside the grid entirely. The test is done on text runs, not
    whole lines: the line number and the material name are usually emitted as a single text
    line that straddles the grid edge, so a line-level test would never see it.
    """
    from b313 import words_with_x
    left, right = [], []
    for L in items:
        for text, x0, x1 in words_with_x(L):
            if x1 < bounds[0] - 1:
                left.append(x0)
            elif x0 > bounds[-1] + 1:
                right.append(x1)
    if len(left) >= min_hits:
        bounds = [round(min(left) - 2, 1)] + bounds
    if len(right) >= min_hits:
        bounds = bounds + [round(max(right) + 2, 1)]
    return bounds


def _colof(x0, x1, bounds):
    cx = (x0 + x1) / 2
    for i in range(len(bounds) - 1):
        if bounds[i] - 3 <= cx <= bounds[i + 1] + 3:
            return i
    return 0 if cx < bounds[0] else len(bounds) - 2


def _rows(lines, bounds, ytol=4.0):
    """Cluster lines into rows, then place each run of text in its column.

    Placement is by run, not by line: PyMuPDF often emits a whole printed row as one text
    line ("1 Carbon steel"), so assigning the line as a unit would drop the line number into
    the material column. Runs are split only at real column gaps, which leaves a spanning
    header phrase intact to be placed by its midpoint.
    """
    from b313 import words_with_x
    rows = []
    for L in sorted(lines, key=lambda L: (L["bbox"][1], L["bbox"][0])):
        cy = (L["bbox"][1] + L["bbox"][3]) / 2
        if rows and abs(cy - rows[-1][0]) <= ytol:
            rows[-1][1].append(L)
        else:
            rows.append([cy, [L]])
    out = []
    for cy, ls in rows:
        cells = [""] * (len(bounds) - 1)
        runs = []
        for L in ls:
            runs.extend(words_with_x(L, bounds))
        for text, x0, x1 in sorted(runs, key=lambda r: r[1]):
            c = _colof(x0, x1, bounds)
            cells[c] = (cells[c] + " " + text).strip() if cells[c] else text
        out.append([clean(c) for c in cells])
    return out


def caption_above(lines, y_top, limit=70.0, x0=None, x1=None):
    """The 'Table X-n' label and its title, set in Ronnia-Bold just above the rules."""
    cands = [L for L in lines
             if y_top - limit < L["bbox"][3] <= y_top + 1 and L["font"].startswith("Ronnia")]
    if x0 is not None:
        cands = [L for L in cands
                 if min(L["bbox"][2], x1) - max(L["bbox"][0], x0) > 0.35 * (L["bbox"][2] - L["bbox"][0])]
    cands.sort(key=lambda L: L["bbox"][1])
    if not cands:
        return None, None
    label = cands[0]["text"]
    title = " ".join(L["text"] for L in cands[1:]) or None
    m = re.match(r"^(Table\s+[A-Z]?[\w.\-]+)\s*(.*)$", label)
    if m:
        return m.group(1), clean((m.group(2) + " " + (title or "")).strip()) or None
    return label, title
