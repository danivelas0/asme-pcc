# ASME Boiler and Pressure Vessel Code

**Section II - Materials — Part B - Nonferrous Material Specifications**

- **Designation:** ASME BPVC.II.B-2025
- **Edition:** 2025 Edition (edition year 2025)
- **Date of issuance:** 2025-07-01
- **Publisher:** ASME (The American Society of Mechanical Engineers)
- **Source:** `B-2025.pdf`, 1319 pages, complete embedded text layer
- **Specifications:** 143 (SB-26/SB-26M through SB/EN 1706)
- **Of which not SB-:** 6 (SA-494/SA-494M, SF-467, SF-467M, SF-468, SF-468M, SB/EN 1706)

Part B is the **nonferrous** half of Section II — aluminium, copper, nickel, titanium and zirconium — and is filed as `SEC II/B/`. It sits beside `SEC II/A1/` and `SEC II/A2/`, which together hold Part A (ferrous). Parts C and D are separate PDFs and are not in this drop; each would sit beside these as its own folder under `SEC II/`.

## Layout

**One folder per material specification**, named as the spec is cited with `/` replaced by `-` (a path cannot contain `/`). Each holds its own JSON and, where the spec has any, a `figures/` subfolder.

```
SB-395-SB-395M/
  SB-395-SB-395M.json    the specification, page by page
  figures/FIG-1.png      its figures, 300 DPI
```

| File / folder | What it holds |
| --- | --- |
| `index.json` | Every folder in Code order: spec id, equivalent standard, page ranges, figure counts |
| `meta.json` | Conversion settings, measured quality, figure-extraction method |
| `content.json` | Full-fidelity single-file conversion, all 1319 pages |
| `SB-*/`, `SA-*/`, `SF-*/` | The 143 specifications, including SA-494/SA-494M, the SF- fastener specs and the SB/EN 1706 adoption |
| `98-Appendices/` | Mandatory Appendices I-IV and Nonmandatory Appendix A |
| `_tools/` | The scripts that produced this drop |

Each spec's JSON carries the edition block, the spec id and title, its **source-standard equivalence** — ASTM for almost every spec (`SB-395/SB-395M` ↔ `B395/B395M-18`), the originating body for the one adoption (`SB/EN 1706` ↔ `EN1706:2020`), with the full "identical with … except …" note — plus PDF and printed page ranges, its figure list, and the page-by-page block tree.

**Folders sort alphabetically, not numerically** — `SB-108` lands before `SB-26`. Walk `index.json` to get them in Code order.

## Specifications

| Spec | Equivalent standard | Pages | Figures |
| --- | --- | ---: | ---: |
| `SB-26/SB-26M` | B26/B26M-11 | 14 | 2 |
| `SB-42` | B42-20 | 8 |  |
| `SB-43` | B43-20 | 8 |  |
| `SB-61` | B61-15(2021) | 4 |  |
| `SB-62` | B62-17 | 4 |  |
| `SB-75/SB-75M` | B75/B75M-19 | 10 |  |
| `SB-96/SB-96M` | B96/B96M-20 | 6 |  |
| `SB-98/SB-98M` | B98/B98M-13(2019) | 6 |  |
| `SB-108/SB-108M` | B108/B108M-12ε1 | 18 | 4 |
| `SB-111/SB-111M` | B111/B111M-18a | 14 |  |
| `SB-127` | B127-05(2014) | 10 |  |
| `SB-135/SB-135M` | B135/B135M-17 | 10 |  |
| `SB-148` | B148-18 | 8 |  |
| `SB-150/SB-150M` | B150/B150M-12(2017) | 8 |  |
| `SB-151/SB-151M` | B151/B151M-20 | 6 |  |
| `SB-152/SB-152M` | B152/B152M-19 | 8 |  |
| `SB-160` | B160-05(2014) | 10 |  |
| `SB-161` | B161-05(2014) | 6 |  |
| `SB-162` | B162-99(2014) | 16 |  |
| `SB-163` | B163-19 | 14 | 1 |
| `SB-164` | B164-03(2014) | 14 |  |
| `SB-165` | B165-19 | 8 |  |
| `SB-166` | B166-19 | 10 |  |
| `SB-167` | B167-18 | 10 |  |
| `SB-168` | B168-19 | 14 |  |
| `SB-169/SB-169M` | B169/B169M-20 | 6 |  |
| `SB-171/SB-171M` | B171/B171M-18 | 8 |  |
| `SB-187/SB-187M` | B187/B187M-20 | 10 | 3 |
| `SB-209/SB-209M` | B209/B209M-21a | 44 |  |
| `SB-210` | B210-12 | 12 | 2 |
| `SB-211/SB-211M` | B211/B211M-19 | 14 |  |
| `SB-221` | B221-21 | 18 |  |
| `SB-221M` | B221M-21 | 20 |  |
| `SB-234` | B234-10 | 8 |  |
| `SB-241/SB-241M` | B241/B241M-22 | 22 |  |
| `SB-247` | B247-09 | 18 |  |
| `SB-248` | B248-17 | 14 | 3 |
| `SB-249/SB-249M` | B249/B249M-20 | 14 | 3 |
| `SB-251/SB-251M` | B251/B251M-17 | 10 |  |
| `SB-265` | B265-20a | 14 |  |
| `SB-271/SB-271M` | B271/B271M-18 | 8 |  |
| `SB-283/SB-283M` | B283/B283M-22 | 18 | 8 |
| `SB-308/SB-308M` | B308/B308M-20 | 6 |  |
| `SB-315` | B315-19 | 12 |  |
| `SB-333` | B333-03(2018) | 4 |  |
| `SB-335` | B335-03(2018) | 6 |  |
| `SB-338` | B338-17 | 10 |  |
| `SB-348/SB-348M` | B348/B348M-19 | 10 |  |
| `SB-359/SB-359M` | B359/B359M-23 | 12 | 2 |
| `SB-363` | B363-14 | 6 |  |
| `SB-366/SB-366M` | B366/B366M-17 | 10 |  |
| `SB-367` | B367-13(2017) | 8 |  |
| `SB-369` | B369-09(2016) | 6 | 2 |
| `SB-381` | B381-13(2019) | 10 |  |
| `SB-395/SB-395M` | B395/B395M-18 | 14 | 2 |
| `SB-407` | B407-08a(2019) | 6 |  |
| `SB-408` | B408-06(2011) | 6 |  |
| `SB-409` | B409-06(2011) | 6 |  |
| `SB-423` | B423-05(2009) | 6 |  |
| `SB-424` | B424-11 | 6 |  |
| `SB-425` | B425-99(2009) | 8 |  |
| `SB-434` | B434-06(2011) | 4 |  |
| `SB-435` | B435-06(2016) | 4 |  |
| `SB-443` | B443-00(2014) | 12 |  |
| `SB-444` | B444-06(2011) | 4 |  |
| `SB-446` | B446-19 | 6 |  |
| `SB-462` | B462-18ε1 | 6 |  |
| `SB-463` | B463-10(2016) | 4 |  |
| `SB-464` | B464-05(2009) | 4 |  |
| `SB-466/SB-466M` | B466/B466M-18 | 8 |  |
| `SB-467` | B467-14(2022) | 10 |  |
| `SB-468` | B468-04(2009) | 4 |  |
| `SB-473` | B473-07(2013) | 10 |  |
| `SB-493/SB-493M` | B493/B493M-14(2019) | 4 |  |
| `SA-494/SA-494M` | A494/A494M-15 | 8 | 3 |
| `SB-505/SB-505M` | B505/B505M-22 | 12 |  |
| `SB-511` | B511-01(2009) | 10 |  |
| `SB-514` | B514-05(2019) | 4 |  |
| `SB-515` | B515-95(2014) | 4 |  |
| `SB-516` | B516-18 | 4 |  |
| `SB-517` | B517-98 | 4 |  |
| `SB-523/SB-523M` | B523/B523M-18 | 6 |  |
| `SB-535` | B535-99 | 4 |  |
| `SB-536` | B536-95 | 12 |  |
| `SB-543/SB-543M` | B543/B543M-18 | 14 |  |
| `SB-548` | B548-03(2009) | 6 | 1 |
| `SB-550/SB-550M` | B550/B550M-07(2019) | 6 |  |
| `SB-551/SB-551M` | B551/B551M-12(2017) | 12 |  |
| `SB-564` | B564-22 | 12 |  |
| `SB-572` | B572-06(2011) | 6 |  |
| `SB-573` | B573-06(2016) | 4 |  |
| `SB-574` | B574-17 | 6 |  |
| `SB-575` | B575-17 | 6 |  |
| `SB-581` | B581-02(2008) | 6 |  |
| `SB-582` | B582-07(2013) | 4 |  |
| `SB-584` | B584-22 | 10 |  |
| `SB-599` | B599-92ε1 | 14 |  |
| `SB-619/SB-619M` | B619/B619M-17 | 8 |  |
| `SB-620` | B620-03(2013) | 4 |  |
| `SB-621` | B621-02(2011) | 4 |  |
| `SB-622` | B622-17b | 8 |  |
| `SB-625` | B625-17 | 6 |  |
| `SB-626` | B626-17 | 6 |  |
| `SB-637` | B637-18 | 8 |  |
| `SB-649` | B649-17 | 10 |  |
| `SB-653/SB-653M` | B653/B653M-11(2020) | 4 |  |
| `SB-658/SB-658M` | B658/B658M-11(2020) | 6 |  |
| `SB-666/SB-666M` | B666/B666M-15 | 8 | 12 |
| `SB-668` | B668-99 | 4 |  |
| `SB-672` | B672-95 | 8 |  |
| `SB-673` | B673-05(2016) | 4 |  |
| `SB-674` | B674-05 | 4 |  |
| `SB-675` | B675-02(2013) | 4 |  |
| `SB-676` | B676-03(2014) | 4 |  |
| `SB-677` | B677-21 | 4 |  |
| `SB-688` | B688-96(2014) | 10 |  |
| `SB-690` | B690-02(2013) | 8 |  |
| `SB-691` | B691-02(2013) | 8 |  |
| `SB-704` | B704-00 | 4 |  |
| `SB-705` | B705-05(2014) | 4 |  |
| `SB-706` | B706-18 | 8 |  |
| `SB-709` | B709-17 | 6 |  |
| `SB-710` | B710-20a | 4 |  |
| `SB-729` | B729-20 | 4 |  |
| `SB-751` | B751-08(2013) | 8 |  |
| `SB-752/SB-752M` | B752/B752M-22 | 6 |  |
| `SB-775` | B775-08 | 6 |  |
| `SB-804` | B804-02(2013) | 8 | 4 |
| `SB-815` | B815-02(2011) | 4 |  |
| `SB-818` | B818-03(2013) | 4 |  |
| `SB-824` | B824-17 | 8 |  |
| `SB-829` | B829-19 | 8 |  |
| `SB-834` | B834-17 | 6 |  |
| `SB-861` | B861-19 | 10 |  |
| `SB-862` | B862-14 | 12 |  |
| `SB-906` | B906-02(2012) | 18 |  |
| `SB-928/SB-928M` | B928/B928M-13 | 14 | 5 |
| `SB-956` | B956-19ε1 | 12 | 3 |
| `SF-467` | F467-13(2018) | 10 |  |
| `SF-467M` | F467M-06a(2018) | 10 |  |
| `SF-468` | F468-16 | 16 | 4 |
| `SF-468M` | F468M-06(2018) | 12 |  |
| `SB/EN 1706` | EN1706:2020 | 2 |  |

## Figures

65 figures across 18 specifications, PNG at 300 DPI, filed under the spec they belong to. Figure numbers restart in every spec, so `FIG-1.png` exists many times over — always scoped by its folder. SB-666/SB-666M prints each of its marking diagrams twice, inch-pound and metric, so it carries both `FIG-2.png` and `FIG-2-M.png`.

**Method.** Section II's ASTM figures have no ruled frame, caption placement varies between specs, and most artwork sits inside Form XObjects whose child bounds pdfium reports in form space rather than page space. Neither Section VIII's frame detector nor raw object geometry works, so the artwork is located from pixels instead. Sparse line art segments into several bands, so a paired band is grown to cover every band within the caption's prose-bounded region; a numbered table claims its own band first, so a figure cannot grow upward into one. A caption that wraps is followed onto its further lines, which are centred on the caption's axis and inset from the body's left margin - running text, justified to the column, never is.

**Excluded.** Publisher and commercial branding: cover and back-cover ASME logos, the ANSI accreditation badge, scan watermarks and QR codes. Only artwork paired with a numbered 'FIG. n' caption is exported, so none are candidates. Sentences that merely open with a figure reference ('Figure 1a has discontinuous grain boundary precipitation ...') are prose, not captions, and are rejected.

**Audited.** Every one of the 65 crops was rendered beside its source page, with the crop region outlined, and checked by eye - a clean run (right count, nothing unmatched) is satisfied just as well by a crop of the wrong region.

## Known limits

- Text coverage against the PDF's own text layer is **95.2%** overall, or **96.7%** once the running header — which marker drops by design — is discounted.
- The **80 landscape pages average 74%** against **98%** for the 1,239 upright ones. Those are the wide alloy/temper property matrices, and they are where nearly all of the loss is: the weakest specifications are the SF- fastener specs and `SB-626`, around 88%. Treat the PDF as authoritative for values read off a landscape table.
- Figure crops are bounded by measured ink, not by a drawn frame, so a few carry an adjoining note or legend that belongs with the drawing. Nothing is cut off; some crops are simply more generous than others.
- Tables are preserved as structured blocks inside the JSON, not re-exported as images.
- **Front and back matter are not included.** The cover, contents, list of sections, foreword, statement of policy, personnel, correspondence, summary of changes, cross-referencing pages and the back cover were removed from this drop, so the page counts above sum to less than the source PDF. `content.json` still holds those pages.
