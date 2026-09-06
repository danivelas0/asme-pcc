# -*- coding: utf-8 -*-
"""Prose extraction for the narrative B31.3 appendices.

Structure comes from two signals the typesetting makes explicit: headings are set in
Ronnia-Bold while body text is Cambria, and every paragraph or list item starts with a
first-line indent of one em (9.6 pt) or a multiple of it, while continuation lines sit flush
on the column margin. That is enough to recover the paragraph breaks and the (a)/(1)/(-a)
nesting without guessing from the text.
"""
import re
from b313 import page_lines, clean, is_heading, COL_SPLIT

EM = 9.6
MARKER = re.compile(r"^(\((?:-?[a-z]{1,2}|\d{1,2}|[ivxl]{1,4})\))\s*")
HEAD_ID = re.compile(r"^([A-Z]{0,2}\d{3}(?:\.\d+)*)\s+(.*)$")


def merge_heading_fragments(lines, ytol=2.0):
    """'F306' and 'FITTINGS, BENDS, ...' are separate blocks on one baseline; rejoin them."""
    out, used = [], set()
    for i, L in enumerate(lines):
        if i in used:
            continue
        if not is_heading(L):
            out.append(L); continue
        group = [L]
        for j in range(i + 1, len(lines)):
            M = lines[j]
            if j in used or not is_heading(M):
                continue
            if abs(M["bbox"][1] - L["bbox"][1]) <= ytol and M["col"] == L["col"]:
                group.append(M); used.add(j)
        group.sort(key=lambda G: G["bbox"][0])
        merged = dict(L)
        merged["text"] = clean(" ".join(G["text"] for G in group))
        merged["bbox"] = (group[0]["bbox"][0], L["bbox"][1],
                          group[-1]["bbox"][2], L["bbox"][3])
        out.append(merged)
    return out


def reading_order(lines):
    """Two-column reading order; full-width lines (the appendix title) lead."""
    full = [L for L in lines if L["full_width"] and L["bbox"][1] < 190]
    rest = [L for L in lines if L not in full]
    full.sort(key=lambda L: L["bbox"][1])
    # Displayed maths is a scatter of separately positioned glyph runs, so ties on a
    # baseline must be ordered left to right or an equation comes out shuffled.
    rest.sort(key=lambda L: (L["col"], round(L["bbox"][1] / 3.0), L["bbox"][0]))
    return full + rest


def column_base(lines):
    """Left margin of each column on this page (54/72 alternate with the recto/verso)."""
    base = {}
    for c in (0, 1):
        xs = [round(L["bbox"][0], 1) for L in lines if L["col"] == c and not is_heading(L)]
        base[c] = min(xs) if xs else (72.0 if c == 0 else 324.0)
    return base


def in_regions(L, regions, pad_top=0.0):
    for r in regions:
        if (r["y_top"] - pad_top <= (L["bbox"][1] + L["bbox"][3]) / 2 <= r["y_bot"]
                and L["bbox"][2] > r["x0"] - 4 and L["bbox"][0] < r["x1"] + 4):
            return True
    return False


def _is_equation(L):
    if L.get("is_equation"):
        return True
    """A displayed equation: math fonts, or a short centred line built around '='."""
    if L["font"].startswith(("STIXGeneral", "ArnoPro", "CMMI", "CMEX", "CMSY")):
        return True
    t = L["text"]
    if len(t) <= 60 and ("=" in t or "\u2264" in t or "\u2265" in t):
        letters = sum(ch.isalpha() for ch in t)
        return letters < 0.55 * max(1, len(t.replace(" ", "")))
    return False


MATH_FONTS = ("STIXGeneral", "ArnoPro", "CMMI", "CMEX", "CMSY", "Mathematica")


def _equation_text(group):
    """Linearise one displayed equation into reading order.

    The maths is a scatter of independently positioned glyph runs - numerator, fraction rule,
    denominator, subscripts - so the runs are clustered into baselines and then read left to
    right. This is the printed expression in reading order, not a parsed formula: the text
    layer carries positions, not structure, and inventing structure would risk changing what
    the equation says.
    """
    rows = []
    for L in sorted(group, key=lambda L: L["bbox"][1]):
        cy = (L["bbox"][1] + L["bbox"][3]) / 2
        if rows and abs(cy - rows[-1][0]) <= 4.5:
            rows[-1][1].append(L)
        else:
            rows.append([cy, [L]])
    parts = []
    for _, ls in rows:
        parts.append(" ".join(L["text"] for L in sorted(ls, key=lambda L: L["bbox"][0])))
    return clean(" ".join(parts))


def _group_equations(lines, base, strokes=()):
    """Fold each displayed equation into a single block.

    An equation is found geometrically, not by font: the maths fonts locate it, but the
    expression also contains ordinary Cambria runs (units, "W = 1.0"), and picking only the
    maths fonts would leave those behind to be shuffled into the surrounding prose. So each
    maths run is grown into a vertical band within its column, overlapping bands are merged,
    and every short indented run inside a band belongs to that equation. Full-measure lines
    starting on the column margin are prose and are never absorbed.
    """
    math = [L for L in lines if L["font"].startswith(MATH_FONTS)]
    if not math:
        return lines
    bands = []
    for L in sorted(math, key=lambda L: (L["col"], L["bbox"][1])):
        b = [L["col"], L["bbox"][1], L["bbox"][3]]
        if bands and bands[-1][0] == b[0] and b[1] - bands[-1][2] <= 12:
            bands[-1][2] = max(bands[-1][2], b[2])
        else:
            bands.append(b)

    def prose_like(L):
        col_base = base.get(L["col"], 72.0)
        return (L["bbox"][2] - L["bbox"][0]) > 120 and L["bbox"][0] <= col_base + 2

    used, merged = set(), []
    for col, y0, y1 in bands:
        group = [L for L in lines
                 if L["col"] == col and not prose_like(L) and not is_heading(L)
                 and y0 - 2 <= (L["bbox"][1] + L["bbox"][3]) / 2 <= y1 + 2]
        for sx0, sy0, sx1, sy1 in strokes:
            if y0 - 2 <= (sy0 + sy1) / 2 <= y1 + 2 and                     (0 if sx0 < COL_SPLIT else 1) == col:
                group.append({"text": "−", "bbox": (sx0, sy0 - 3.2, sx1, sy1 + 1.0),
                              "font": "STIXGeneral-Regular", "size": 9.6,
                              "chars": [], "col": col, "full_width": False})
        if not group:
            continue
        used.update(id(L) for L in group)
        merged.append({"text": _equation_text(group), "font": "equation",
                       "bbox": (min(x["bbox"][0] for x in group), y0,
                                max(x["bbox"][2] for x in group), y1),
                       "size": group[0]["size"], "chars": [], "col": col,
                       "full_width": False, "is_equation": True})
    return [L for L in lines if id(L) not in used] + merged


def page_blocks(page, exclude=(), title_y=190):
    """Ordered content blocks for one prose page: headings, paragraph lines, equations."""
    lines = page_lines(page)
    lines = [L for L in lines if not in_regions(L, exclude)]
    lines = merge_heading_fragments(lines)
    base = column_base(lines)
    strokes = [(d["rect"].x0, d["rect"].y0, d["rect"].x1, d["rect"].y1)
               for d in page.get_drawings()
               if d["rect"].height < 2.0 and 3.0 < d["rect"].width < 12.0]
    lines = _group_equations(lines, base, strokes)
    blocks = []
    for L in reading_order(lines):
        b = base.get(L["col"], 72.0)
        indent = max(0, int(round((L["bbox"][0] - b) / EM)))
        blocks.append({
            "text": L["text"], "heading": is_heading(L), "indent": indent,
            "equation": _is_equation(L), "size": L["size"],
            "y": L["bbox"][1], "col": L["col"],
            "appendix_title": L["size"] > 14,
        })
    return blocks


def assemble(blocks, page_numbers):
    """Fold the per-page block stream into sections with paragraphs."""
    sections, cur = [], None
    para = None

    def flush():
        nonlocal para
        if para and para["text"].strip():
            para["text"] = clean(para["text"])
            (cur["paragraphs"] if cur else sections_pre).append(para)
        para = None

    sections_pre = []
    for blk, pg in blocks:
        if blk["appendix_title"]:
            continue
        if blk["heading"]:
            flush()
            m = HEAD_ID.match(blk["text"])
            sid, htext = (m.group(1), m.group(2)) if m else (None, blk["text"])
            if cur and cur["heading"] and not cur["paragraphs"] and \
               cur["id"] and sid is None and blk["size"] == cur["size"]:
                cur["heading"] = clean(cur["heading"] + " " + htext)   # wrapped heading
                continue
            cur = {"id": sid, "heading": htext or None,
                   "level": 1 if (htext or "").isupper() else 2,
                   "page": pg, "size": blk["size"], "paragraphs": []}
            sections.append(cur)
            continue
        if blk["equation"]:
            flush()
            target = cur["paragraphs"] if cur else sections_pre
            # One printed equation arrives as many glyph runs (numerator, rule, denominator,
            # subscripts). Runs that continue in the same column within a couple of lines are
            # the same equation, so they are joined rather than emitted one fragment each.
            if (target and target[-1].get("kind") == "equation"
                    and blk["col"] == target[-1].get("_col")
                    and 0 <= blk["y"] - target[-1]["_y"] <= 22):
                target[-1]["text"] = clean(target[-1]["text"] + " " + blk["text"])
                target[-1]["_y"] = blk["y"]
            else:
                target.append({"kind": "equation", "text": blk["text"],
                               "indent": blk["indent"], "_y": blk["y"],
                               "_col": blk["col"]})
            continue
        starts = blk["indent"] >= 1
        if starts or para is None:
            flush()
            m = MARKER.match(blk["text"])
            para = {"kind": "paragraph", "marker": m.group(1) if m else None,
                    "indent": blk["indent"], "text": blk["text"]}
        else:
            para["text"] += " " + blk["text"]
    flush()
    for s in sections:
        s.pop("size", None)
        for par in s["paragraphs"]:
            par.pop("_y", None)
            par.pop("_col", None)
    for par in sections_pre:
        par.pop("_y", None)
        par.pop("_col", None)
    return sections_pre, sections


class Dehyphenator:
    """Rejoin words broken by a line-break hyphen, without welding real compounds shut.

    'selec- tively' must become 'selectively', but 'tongue-and- groove' must stay hyphenated.
    Deciding by rule is unreliable, so both candidate forms are tested against a vocabulary
    built from the whole document: whichever form the code itself uses elsewhere wins.
    """
    BREAK = re.compile(r"(\w[\w\u2010\-]*)[-\u2010]\s+(\w+)")

    def __init__(self, doc, pages):
        from b313 import page_lines
        self.plain, self.hyph = set(), set()
        for i in pages:
            for L in page_lines(doc[i]):
                for w in re.findall(r"[A-Za-z][A-Za-z\u2010\-]*[A-Za-z]", L["text"]):
                    (self.hyph if ("-" in w or "\u2010" in w) else self.plain).add(w.lower())

    SUSPENDED = {"and", "or", "to", "nor", "but", "through"}

    def _fix(self, m):
        a, b = m.group(1), m.group(2)
        if b.lower() in self.SUSPENDED:
            return m.group(0)          # suspended compound: "low- and high-cycle"
        if "-" in a or "‐" in a:
            joined_c = (a + b).lower()
            if joined_c not in self.plain:
                return a + "-" + b     # already a compound; the break is its next hyphen
        joined = (a + b)
        hyphenated = a + "-" + b
        jl, hl = joined.lower(), hyphenated.lower()
        if jl in self.plain and hl not in self.hyph:
            return joined
        if hl in self.hyph and jl not in self.plain:
            return hyphenated
        if hl in self.hyph and jl in self.plain:
            return hyphenated          # the code prints it hyphenated elsewhere
        return joined                  # unseen: a line-break hyphen is the common case

    def __call__(self, text):
        prev = None
        while prev != text:
            prev = text
            text = self.BREAK.sub(self._fix, text)
        return text
