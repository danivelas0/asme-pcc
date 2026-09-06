# -*- coding: utf-8 -*-
"""Upright view of a landscape page.

PyMuPDF reports text *and* drawings in unrotated page space whatever the page's /Rotate
value is - only page.rect changes. So on the landscape sheets (Table 302.3.5-1, Table
341.3.2-1, Figure 304.3.3-1 and their continuations) the printed horizontal rules arrive as
tall thin vertical rectangles and the rule-grid table engine finds nothing at all.

Measured on idx 120: the caption "Table 341.3.2-1" sits at x 58-68, y 366-432 with text
direction (0, -1), i.e. the content is turned 90 degrees counter-clockwise inside a 612x792
media box. Rotating it back gives

    upright_x = page_height - y        upright_y = x

which puts that caption at x 360-426 of a 792-wide sheet, y 58-68 - centred just under the
head margin, where a table caption belongs.

This module wraps a page so every coordinate it hands out is already upright. Nothing
downstream has to know: the table, prose and figure engines see an ordinary portrait-shaped
page whose rules are horizontal again.
"""
import pymupdf


class UprightPage(object):
    """Read-only proxy presenting a rotated page's content in upright coordinates."""

    def __init__(self, page):
        if page.rotation != 90:
            # The transform below was measured on the /Rotate 90 sheets, which is every
            # landscape page in this document. A 270 sheet is its mirror and would come out
            # upside down, so refuse rather than emit a wrong figure.
            raise NotImplementedError(
                "page %d has rotation %d; only 90 is handled" % (page.number, page.rotation))
        self._page = page
        self._h = page.mediabox.height if page.rotation in (90, 270) else page.rect.height
        w = page.mediabox.height
        h = page.mediabox.width
        self.rect = pymupdf.Rect(0, 0, w, h)
        self.rotation = 0            # coordinates handed out are already upright
        # The running head and the printed folio do not turn with the table: they stay in
        # the physical sheet's orientation, so in upright space they sit at the right and
        # left edges. The portrait band (unrotated y between 55 and 735) maps to upright x
        # between 57 and 737; nothing bounds the content in upright y. Measured on idx 120:
        # folio at x 37-47, running head at x 745-753, table content 66-737.
        self.content_box = (w - 735.0, float("-inf"), w - 55.0, float("inf"))
        # A landscape sheet carries one full-width table or figure, never two prose columns.
        self.col_split = w
        self.source_rotation = page.rotation
        self.number = page.number

    # -- coordinate transform ------------------------------------------------
    def _pt(self, x, y):
        return (self._h - y, x)

    def _box(self, b):
        x0, y0, x1, y1 = b
        return (self._h - y1, x0, self._h - y0, x1)

    def unrotate_rect(self, r):
        """Map an upright rect back to unrotated page space, for get_pixmap clipping."""
        return pymupdf.Rect(r.y0, self._h - r.x1, r.y1, self._h - r.x0)

    # -- proxied content -----------------------------------------------------
    def get_text(self, kind="text", **kw):
        d = self._page.get_text(kind, **kw)
        if kind not in ("dict", "rawdict"):
            return d
        for b in d["blocks"]:
            if b.get("type") != 0:
                b["bbox"] = self._box(b["bbox"])
                continue
            b["bbox"] = self._box(b["bbox"])
            for l in b["lines"]:
                l["bbox"] = self._box(l["bbox"])
                l["dir"] = (1.0, 0.0)
                for s in l["spans"]:
                    s["bbox"] = self._box(s["bbox"])
                    s["origin"] = self._pt(*s["origin"])
                    for c in s.get("chars", ()):
                        c["bbox"] = self._box(c["bbox"])
                        c["origin"] = self._pt(*c["origin"])
        return d

    def get_drawings(self, **kw):
        out = []
        for dr in self._page.get_drawings(**kw):
            d = dict(dr)
            d["rect"] = pymupdf.Rect(*self._box(tuple(dr["rect"])))
            out.append(d)
        return out

    def get_images(self, *a, **kw):
        return self._page.get_images(*a, **kw)

    def get_image_info(self, **kw):
        out = []
        for im in self._page.get_image_info(**kw):
            d = dict(im)
            d["bbox"] = self._box(tuple(im["bbox"]))
            out.append(d)
        return out

    def get_pixmap(self, clip=None, matrix=None, **kw):
        """Rasterise an upright clip.

        Rendering, unlike text and drawing extraction, already honours the page's /Rotate,
        and MuPDF's rotated space turns out to be exactly the upright space computed here -
        checked by clipping to a caption's upright bbox and getting that caption back. So
        the clip is passed straight through and no correcting matrix is applied; supplying
        one produced figures lying on their side.
        """
        return self._page.get_pixmap(clip=clip, matrix=matrix or pymupdf.Identity, **kw)


def upright(page):
    """The page itself when portrait, an upright proxy when landscape."""
    return UprightPage(page) if page.rotation else page
