# -*- coding: utf-8 -*-
"""Shared extraction core for ASME B31.3-2024 appendices E-Z.

The PDF has a real text layer, but it is justified typesetting: on many lines the space
glyphs are simply not emitted and the words are separated by letter-spacing instead. Every
text value therefore has to be rebuilt from character positions rather than taken from
get_text(), or roughly a third of the prose comes out as runtogetherwords.
"""
import re, unicodedata

SRC = (r"C:\Users\User\OneDrive - INSPECTRA SA\Engineering\0-STANDARDS AND CODES"
       r"\0-CODES\B31- PRESSURE PIPING\ASME B31.3 2024 Process Piping.pdf")

HEAD_Y, FOOT_Y = 55.0, 735.0     # running head above, printed folio below
COL_SPLIT = 306.0

SOURCE_BLOCK = {
    "code": "ASME B31.3",
    "edition": "2024",
    "title": "Process Piping - ASME Code for Pressure Piping, B31",
    "extracted_from": "ASME B31.3 2024 Process Piping.pdf",
    "note": ("Technical content only. Publisher front/back matter, running heads, printed "
             "folios and edition change markers are omitted. Values are as printed; the "
             "ASME ellipsis (no value / not permitted) is represented as null."),
}

# ð24Þ / ð23Þ etc: the ASME revision change-marker, not technical content.
CHANGE_MARK = re.compile(r"[\u00f0\u00f1]\s*\d{2}\s*[\u00de\u00fe]")
RUN_HEAD = re.compile(r"^ASME\s+B31\.3-2024$")


def clean(s):
    s = CHANGE_MARK.sub("", s)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def is_ellipsis(s):
    """ASME prints U+2026 (or spaced dots) to mean 'no value / not permitted'."""
    return bool(s) and set(s.replace(" ", "")) <= {"\u2026", "."} and "\u2026" in s or \
           s.strip() in {"...", ". . .", "\u2026"}


def num(s):
    """Type a printed cell. Returns int/float, or the string, or None for the ellipsis."""
    if s is None:
        return None
    t = clean(s)
    if not t or is_ellipsis(t):
        return None
    t2 = t.replace("\u2212", "-").replace("\u2013", "-").replace("\u2010", "-")
    t2 = t2.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", t2) else t2
    # SI tables separate thousands with a space ("1 895"); that is one number, not two.
    if re.fullmatch(r"-?\d{1,3}(\s\d{3})+", t2):
        t2 = re.sub(r"\s", "", t2)
    if re.fullmatch(r"-?\d+", t2):
        return int(t2)
    if re.fullmatch(r"-?\d*\.\d+", t2) or re.fullmatch(r"-?\d+\.", t2):
        return float(t2.rstrip("."))
    if re.fullmatch(r"-?\d+(\.\d+)?[eE][-+]?\d+", t2):
        return float(t2)
    return t


def _chars_of(line):
    """Characters of a line, each tagged with the font of the span it came from.

    The tag costs nothing here and is what lets a caller split a line at a font boundary.
    ASME sets its subsection headings run-in - "326.1.1 Listed Piping Components." in bold,
    the requirement continuing in roman on the same line - so a whole-line font decision
    puts the first words of the text into the heading and starts the paragraph mid-word.
    """
    out = []
    for sp in line["spans"]:
        for c in sp["chars"]:
            c["font"] = sp["font"]
            c["size"] = sp["size"]
            out.append(c)
    return out


def _join(chars, size=9.6):
    """Rebuild a string from characters, restoring space glyphs the typesetter dropped.

    Two opposite failures have to be told apart. On some justified lines the space glyph is
    simply absent and the words are pushed apart by ~1.4 pt ("CodeCompliance"); on others the
    spaces are present but every letter is tracked out by ~0.6 pt to fill the measure, and a
    fixed threshold would then split every word into letters. So the threshold is derived per
    line from that line's own letter tracking: the median gap between adjacent non-space
    characters is the baseline, and only a gap clearly above it means a missing space.
    """
    gaps = []
    for a, b in zip(chars, chars[1:]):
        if a["c"] != " " and b["c"] != " ":
            gaps.append(b["bbox"][0] - a["bbox"][2])
    nspace = sum(1 for c in chars if c["c"] == " ")
    if nspace and nspace / len(chars) > 0.45:
        # Fully tracked-out setting: the typesetter put a real space between every letter
        # (URLs in Appendix N are set this way). Keep only the spaces that are wider than
        # the uniform letter tracking, which are the genuine word breaks.
        spans = []
        prev_ns = None
        for c in chars:
            if c["c"] == " ":
                continue
            if prev_ns is not None:
                spans.append(c["bbox"][0] - prev_ns["bbox"][2])
            prev_ns = c
        if spans:
            med = sorted(spans)[len(spans) // 2]
            keep = med + max(1.0, 0.6 * abs(med))
            out, prev_ns = [], None
            for c in chars:
                if c["c"] == " ":
                    continue
                if prev_ns is not None and c["bbox"][0] - prev_ns["bbox"][2] > keep:
                    out.append(" ")
                out.append(c["c"])
                prev_ns = c
            return "".join(out)
    base = 0.0
    if len(gaps) >= 5:
        g = sorted(gaps)
        base = max(0.0, g[len(g) // 2])
    thr = base + max(0.45, 0.05 * size)
    out, prev = [], None
    for c in chars:
        ch = c["c"]
        if prev is not None and ch != " " and prev["c"] != " ":
            if c["bbox"][0] - prev["bbox"][2] > thr:
                out.append(" ")
        out.append(ch)
        prev = c
    return "".join(out)


def page_lines(page, keep_furniture=False):
    """Reconstructed text lines, ordered reading-wise (column 0 top-to-bottom, then column 1).

    Each line: text, bbox, font/size of its dominant span, column index, and the raw spans
    so callers that need per-cell x positions (tables) can work from characters.
    """
    rot = page.rotation
    # A landscape sheet seen through rotate.UprightPage keeps its running head and folio in
    # the physical sheet's orientation, so there they bound the content in x, not in y. The
    # band is therefore carried on the page as a box. An ordinary page has no such attribute
    # and gets the portrait band, unchanged.
    box = getattr(page, "content_box", (float("-inf"), HEAD_Y, float("inf"), FOOT_Y))
    col_split = getattr(page, "col_split", COL_SPLIT)
    raw = page.get_text("rawdict")
    lines = []
    for b in raw["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            chars = _chars_of(l)
            if not chars:
                continue
            txt = clean(_join(chars, l["spans"][0]["size"]))
            if not txt:
                continue
            x0, y0, x1, y1 = l["bbox"]
            if not keep_furniture:
                if y1 < box[1] or y0 > box[3] or x1 < box[0] or x0 > box[2]:
                    continue
                # The printed folio sits at y=745, already outside the content band, so
                # nothing here may reject a line merely for being a bare number: a table's
                # last row can fall well below 700 and is data, not furniture.
                if RUN_HEAD.match(txt):
                    continue
            sizes = {}
            for sp in l["spans"]:
                sizes[(sp["font"], round(sp["size"], 1))] = sizes.get(
                    (sp["font"], round(sp["size"], 1)), 0) + len(sp["chars"])
            font, size = max(sizes.items(), key=lambda kv: kv[1])[0]
            lines.append({
                "text": txt, "bbox": (x0, y0, x1, y1), "font": font, "size": size,
                "chars": chars, "col": 0 if x0 < col_split else 1,
                "full_width": (x1 - x0) > 300,
            })
    if rot:                                   # landscape pages: reading order is by x then y
        lines.sort(key=lambda L: (round(L["bbox"][0], 0), L["bbox"][1]))
    else:
        lines.sort(key=lambda L: (0 if L["full_width"] else 1 + L["col"], L["bbox"][1]))
        # full-width lines interleave by y with whichever column they precede
        lines.sort(key=lambda L: (L["bbox"][1] // 400, 0 if L["full_width"] else 1 + L["col"],
                                  L["bbox"][1]))
    return lines


def is_heading(line):
    return line["font"].startswith("Ronnia-Bold")


def words_with_x(line, bounds=None):
    """Split a reconstructed line into (text, x0, x1) runs separated by real column gaps.

    A wide gap is a column break on its own. Where the table's rule grid is known, a much
    narrower gap also breaks the run if a column boundary falls inside it - the line number
    and the material name are set only about two points apart in Table K-1, far less than an
    ordinary column gap, and would otherwise land in the same cell. Gaps are measured between
    non-space characters, because a space glyph carries its own advance width and would
    otherwise hide the gap it sits in.
    """
    runs, cur, prev = [], [], None
    for c in line["chars"]:
        if c["c"] == " ":
            if cur:
                cur.append(c)
            continue
        if prev is not None:
            gap = c["bbox"][0] - prev["bbox"][2]
            straddles = bounds is not None and gap > 0.8 and any(
                prev["bbox"][2] - 0.5 <= b <= c["bbox"][0] + 0.5 for b in bounds)
            if gap > 3.0 or straddles:
                runs.append(cur); cur = []
        cur.append(c)
        prev = c
    if cur:
        runs.append(cur)
    out = []
    for r in runs:
        while r and r[0]["c"] == " ":
            r = r[1:]
        while r and r[-1]["c"] == " ":
            r = r[:-1]           # a trailing space carries advance width and would
        if not r:                # otherwise push the run's right edge into the next column
            continue
        t = clean(_join(r))
        if t:
            out.append((t, r[0]["bbox"][0], r[-1]["bbox"][2]))
    return out
