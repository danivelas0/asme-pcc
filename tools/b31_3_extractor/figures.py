# -*- coding: utf-8 -*-
"""Figure rendering.

Every figure in these appendices is vector line-art - the pages carry no raster images at
all - so a figure has to be rasterised from a clip rectangle. The rectangle is the union of
the page's drawing operations and the text sitting among them, taken inside the content band
so the running head and folio never enter the image.
"""
import re
import pymupdf
from b313 import page_lines, HEAD_Y, FOOT_Y

FIG_LABEL = re.compile(r"^(Figure\s+[A-Z]?[0-9A-Za-z.\-]+)(?![0-9A-Za-z])")
DPI = 300


def figure_captions(page, lines=None):
    lines = lines if lines is not None else page_lines(page)
    return [L for L in lines
            if L["font"].startswith("Ronnia") and FIG_LABEL.match(L["text"])]


def _art_rect(page, caption, table_regions=()):
    """Extent of the figure's own line art.

    A page carries stray hairlines well below a figure - rules under headings, fraction bars
    in equations - and unioning every drawing on the page stretches the clip down over the
    prose that follows. So the drawings are clustered vertically and only the run of clusters
    continuous with the caption is kept.
    """
    boxes = []
    for dr in page.get_drawings():
        b = dr["rect"]
        if b.is_empty or b.is_infinite or b.y1 < HEAD_Y or b.y0 > FOOT_Y:
            continue
        if b.height < 2.5 and b.width > 20 and any(
                reg["y_top"] - 2 <= b.y0 <= reg["y_bot"] + 2 for reg in table_regions):
            continue
        boxes.append(b)
    if not boxes:
        return pymupdf.Rect()
    boxes.sort(key=lambda b: b.y0)
    clusters = []
    for b in boxes:
        if clusters and b.y0 - clusters[-1][1] <= 25:
            clusters[-1][1] = max(clusters[-1][1], b.y1)
            clusters[-1][2] |= b
        else:
            clusters.append([b.y0, b.y1, pymupdf.Rect(b)])
    cap_y = caption["bbox"][3]
    start = next((i for i, c in enumerate(clusters) if c[1] > cap_y), 0)
    keep = pymupdf.Rect(clusters[start][2])
    last = clusters[start][1]
    for c in clusters[start + 1:]:
        if c[0] - last > 60:
            break
        keep |= c[2]
        last = c[1]
    return keep


def figure_rect(page, caption, table_regions=()):
    """Clip rect for one figure: its ink, the labels inside it, and its caption."""
    art = _art_rect(page, caption, table_regions)
    if art.is_empty:
        return None
    band = pymupdf.Rect(0, HEAD_Y, page.rect.x1, FOOT_Y)
    art &= band
    box = pymupdf.Rect(art)
    for L in page_lines(page):
        lb = pymupdf.Rect(*L["bbox"])
        if art.intersects(lb) or (art.y0 - 4 <= lb.y0 and lb.y1 <= art.y1 + 4
                                  and art.x0 - 8 <= lb.x0 and lb.x1 <= art.x1 + 8):
            box |= lb
    # Part labels - "(c) Subsurface Flaw" - are set just under the last drawing and would
    # otherwise be cropped off. Only short lines qualify, so body prose is never pulled in.
    for L in page_lines(page):
        lb = pymupdf.Rect(*L["bbox"])
        if (0 <= lb.y0 - art.y1 <= 32 and lb.width < 250
                and lb.x1 > art.x0 - 20 and lb.x0 < art.x1 + 20):
            box |= lb
    box |= pymupdf.Rect(*caption["bbox"])
    box &= band
    box.x0 = max(box.x0 - 6, 0); box.y0 = max(box.y0 - 4, HEAD_Y)
    box.x1 = min(box.x1 + 6, page.rect.x1); box.y1 = min(box.y1 + 4, FOOT_Y)
    return box


def render(page, rect, out_path, dpi=DPI, rotate=0):
    """Rasterise a clip rect.

    Landscape figures are set rotated on a portrait sheet, and PyMuPDF reports both drawing
    and text coordinates in unrotated page space, so the clip is computed there and the page
    rotation is reapplied at render time to bring the image out upright.
    """
    mat = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
    if rotate:
        mat = mat * pymupdf.Matrix(1, 1).prerotate(rotate)
    pix = page.get_pixmap(clip=rect, matrix=mat, colorspace=pymupdf.csRGB, alpha=False)
    pix.save(out_path)
    return {"width": pix.width, "height": pix.height, "dpi": dpi}


def caption_text(page, caption, lines=None):
    """The label line plus the title lines set beneath it, in Ronnia-Bold."""
    lines = lines if lines is not None else page_lines(page)
    y = caption["bbox"][1]
    # On a landscape page the caption's continuation lines sit alongside it in x, not under
    # it in y, because the whole block is rotated on the sheet.
    if page.rotation or (caption["bbox"][3] - caption["bbox"][1]) > 40:
        x = caption["bbox"][0]
        same = [L for L in lines if L["font"].startswith("Ronnia")
                and x - 30 <= L["bbox"][0] <= x + 30]
        same.sort(key=lambda L: (-L["bbox"][0], L["bbox"][1]))
    else:
        same = [L for L in lines if L["font"].startswith("Ronnia")
                and y - 1 <= L["bbox"][1] <= y + 30
                and abs(((L["bbox"][0] + L["bbox"][2]) / 2)
                        - ((caption["bbox"][0] + caption["bbox"][2]) / 2)) < 160]
        same.sort(key=lambda L: (L["bbox"][1], L["bbox"][0]))
    label = FIG_LABEL.match(caption["text"]).group(1)
    rest = " ".join(L["text"] for L in same)
    title = rest.replace(label, " ", 1).strip()   # the label can lead or trail the title
    title = " ".join(title.split())
    return label, (title or None)
