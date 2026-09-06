# -*- coding: utf-8 -*-
"""Extractors for the three appendices whose layout the generic engines do not cover:
E (a multi-column standards list with no rules), J (a four-column nomenclature list) and
K (two-page stress-table spreads joined by line number)."""
import re
from collections import defaultdict
from b313 import page_lines, clean, num
import tables as T
import build as B

CONTD = re.compile(r"\s*\(Cont[’']d\)\s*$", re.I)


def _bands(lines, min_gap=40.0):
    """Cluster line start positions into list columns (Appendix E has two or three)."""
    xs = sorted({round(L["bbox"][0], 1) for L in lines})
    if not xs:
        return []
    groups, cur = [], [xs[0]]
    for x in xs[1:]:
        if x - cur[-1] > min_gap:
            groups.append(cur)
            cur = []
        cur.append(x)
    groups.append(cur)
    return [min(g) for g in groups]


def _band_of(L, starts):
    x = L["bbox"][0]
    best = 0
    for i, s in enumerate(starts):
        if x >= s - 6:
            best = i
    return best


# --------------------------------------------------------------------------- Appendix E
def appendix_e(doc):
    intro, orgs, general_notes, org_intro = [], [], [], None
    publishers, order = {}, []

    for i in range(462, 466):
        lines = page_lines(doc[i])
        body = [L for L in lines if not L["font"].startswith("Ronnia")]
        if i == 462:                       # the appendix opens with explanatory prose
            intro_lines = [L for L in body if L["bbox"][1] < 320]
            body = [L for L in body if L["bbox"][1] >= 320]
            para = None
            for L in sorted(intro_lines, key=lambda L: L["bbox"][1]):
                if L["bbox"][0] > 58 and para is not None:
                    intro.append(para)
                    para = None
                para = L["text"] if para is None else para + " " + L["text"]
            if para:
                intro.append(para)
        starts = _bands(body)
        cols = defaultdict(list)
        for L in body:
            cols[_band_of(L, starts)].append(L)
        for b in sorted(cols):
            cur = None
            for L in sorted(cols[b], key=lambda L: L["bbox"][1]):
                if L["font"].endswith("Bold"):
                    name = CONTD.sub("", L["text"]).strip()
                    if name not in publishers:
                        publishers[name] = {"publisher": name, "entries": []}
                        order.append(name)
                    cur = publishers[name]
                elif cur is not None:
                    cur["entries"].append(L["text"])

    lines = page_lines(doc[466])
    for L in sorted(lines, key=lambda L: L["bbox"][1]):
        if L["bbox"][1] < 110:
            general_notes.append(L["text"])
        elif L["bbox"][1] < 300:
            org_intro = clean((org_intro + " " + L["text"]) if org_intro else L["text"])
    rows = T._rows([L for L in lines if L["bbox"][1] >= 300], [50.0, 140.0, 450.0, 560.0])
    for r in rows:
        if r[0]:
            orgs.append({"abbreviation": r[0], "organization": r[1] or None,
                         "website": r[2] or None})
    return {
        "introduction": [clean(p) for p in intro],
        "general_notes": [clean(" ".join(general_notes))] if general_notes else [],
        "standards": [publishers[k] for k in order],
        "organizations_note": org_intro,
        "organizations": orgs,
    }


# --------------------------------------------------------------------------- Appendix J
J_FIELDS = ["symbol", "definition", "units_si", "units_us_customary",
            "paragraph", "table_fig_app", "equation"]
# Column grid measured on the first page of the list; every page prints the same grid,
# shifted bodily left or right with the recto/verso margin, so one offset re-registers it.
J_BOUNDS = [54.0, 90.0, 262.0, 295.0, 347.0, 402.0, 484.0, 560.0]
J_SYMBOL_X = 57.2


def appendix_j(doc, pages=range(486, 502)):
    entries, notes = [], []
    for i in pages:
        lines = page_lines(doc[i])
        hdr = [L for L in lines if L["text"] == "Symbol"]
        if not hdr:
            continue
        off = hdr[0]["bbox"][0] - J_SYMBOL_X
        bounds = [x + off for x in J_BOUNDS]
        rules = [y for y in T.rule_groups(doc[i])]
        top = max([y for y in rules if y < 130] or [107.5])
        bot = max(rules) if max(rules) > top else 720.0
        body = [L for L in lines if top < (L["bbox"][1] + L["bbox"][3]) / 2 < bot]
        for L in sorted((L for L in lines if L["bbox"][1] > bot),
                        key=lambda L: L["bbox"][1]):
            if re.match(r"^(GENERAL NOTE|NOTE)", L["text"]):
                notes.append(L["text"])
            elif notes:
                notes[-1] = clean(notes[-1] + " " + L["text"])
        for row in T._rows(body, bounds):
            cells = (row + [""] * 7)[:7]
            if cells[0]:
                rec = {k: (num(v) if k not in ("symbol", "definition") else (v or None))
                       for k, v in zip(J_FIELDS, cells)}
                rec["pdf_page"] = i + 1
                entries.append(rec)
            elif entries and any(cells[1:]):
                e = entries[-1]
                if cells[1]:
                    e["definition"] = clean(((e["definition"] or "") + " " + cells[1]).strip())
                for k, v in zip(J_FIELDS[2:], cells[2:]):
                    v = num(v)
                    if v is not None:
                        e[k] = v if e.get(k) in (None, "") else \
                            clean("%s %s" % (e[k], v))
    return entries, notes


# --------------------------------------------------------------------------- Appendix K
TEMP_KEY = re.compile(r"(-?\d+)\s*$")
TEMP_COL = re.compile(r"^(Min\. Temp\. to\s+)?-?\d+$")


def _spread_half(doc, i, header_rule=158.8):
    pg = doc[i]
    lines = page_lines(pg)
    regs = T.find_tables(pg, lines)
    if not regs:
        return None
    reg = regs[0]
    head, banners, body, bounds = T.header_and_body(pg, reg, lines, header_rule)
    names = B._colnames(head, len(bounds) - 1)
    label, title = T.caption_above(lines, reg["y_top"], x0=reg["x0"], x1=reg["x1"])
    return {"page": i + 1, "names": names, "rows": B.merge_wrapped(body),
            "banners": banners, "label": CONTD.sub("", label or "").strip(),
            "title": CONTD.sub("", title).strip() if title else None}


def k_spread_table(doc, first, last, table_id):
    """Join the material half and the stress half of each two-page spread by line number."""
    rows, banners, title, pages, stress_headers = [], [], None, [], {}
    for i in range(first, last + 1, 2):
        left, right = _spread_half(doc, i), _spread_half(doc, i + 1)
        if left is None or right is None:
            continue
        pages += [left["page"], right["page"]]
        title = title or left["title"] or right["title"]
        for b in left["banners"] + right["banners"]:
            if b not in banners:
                banners.append(b)
        lkeys = B._uniq([B._key(n, k) for k, n in enumerate(left["names"])])
        lrows = {r[0]: r for r in left["rows"] if r and r[0]}
        rrows = {r[0]: r for r in right["rows"] if r and r[0]}
        for ln in sorted(lrows, key=lambda v: int(v) if v.isdigit() else 0):
            lr, rr = lrows[ln], rrows.get(ln)
            rec = {"line_no": num(ln), "pdf_pages": [left["page"], right["page"]]}
            for k, v in list(zip(lkeys, lr))[1:]:
                rec[k] = num(v)
            stress = {}
            if rr:
                # The stress half is not laid out identically on every spread: on two of
                # them the tensile/yield strength columns are printed on the right-hand page
                # instead of the left, so columns are classified by their printed heading
                # rather than by position.
                for name, val in list(zip(right["names"], rr))[1:]:
                    if not name:
                        continue
                    if TEMP_COL.match(name):
                        key = TEMP_KEY.search(name).group(1)
                        stress[key] = num(val)
                        stress_headers.setdefault(key, name)
                    elif name.startswith("Max. Temp"):
                        rec["max_temp"] = num(val)
                        rec.setdefault("max_temp_header", name)
                    else:
                        rec[B._key(name, 0)] = num(val)
            rec["allowable_stress"] = stress
            rows.append(rec)
    return {"table_id": table_id, "title": title, "pages": pages, "rows": rows,
            "banners": banners,
            "stress_columns": [{"temperature": k, "header": v}
                               for k, v in sorted(stress_headers.items(),
                                                  key=lambda kv: int(kv[0]))]}


def k_spec_index(doc, i=503):
    lines = page_lines(doc[i])
    body = [L for L in lines if L["bbox"][1] > 112]
    out, cur = [], None
    for row in T._rows(body, [66.0, 108.0, 560.0]):
        spec, title = (row + ["", ""])[:2]
        if spec and not title:
            cur = {"publisher": spec, "entries": []}
            out.append(cur)
        elif spec:
            if cur is None:
                cur = {"publisher": None, "entries": []}
                out.append(cur)
            cur["entries"].append({"spec_no": spec, "title": title})
        elif title and cur and cur["entries"]:
            e = cur["entries"][-1]
            e["title"] = clean(e["title"] + " " + title)
    return out
