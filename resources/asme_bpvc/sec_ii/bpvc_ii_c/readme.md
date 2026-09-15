# ASME Boiler and Pressure Vessel Code

**Section II - Materials — Part C - Specifications for Welding Rods, Electrodes, and Filler Metals**

- **Designation:** ASME BPVC.II.C-2025
- **Edition:** 2025 Edition (edition year 2025)
- **Date of issuance:** 2025-07-01
- **Publisher:** ASME (The American Society of Mechanical Engineers)
- **Source:** `C-2025.pdf`, 1153 pages, complete embedded text layer
- **Specifications:** 35 (SFA-5.01M/SFA-5.01 through SFA-5.39/SFA-5.39M)

Part C holds the **welding consumables** — rods, electrodes, filler metals, fluxes and shielding gases — and is filed as `SEC II/C/`. Every specification in it is a reprint of an **AWS** standard, which is what makes it read differently from Parts A and B: the equivalence is to `A5.x`, not to ASTM, and the figures are titled in the AWS house style. It sits beside `SEC II/A1/`, `SEC II/A2/` (Part A, ferrous) and `SEC II/B/` (Part B, nonferrous). Part D is a separate PDF and is not in this drop.

## Layout

**One folder per specification**, named as the spec is cited with `/` replaced by `-` (a path cannot contain `/`). Each holds its own JSON and, where the spec has any, a `figures/` subfolder.

```
SFA-5.4-SFA-5.4M/
  SFA-5.4-SFA-5.4M.json    the specification, page by page
  figures/Figure-1.png       its figures, 300 DPI
```

| File / folder | What it holds |
| --- | --- |
| `index.json` | Every folder in Code order: spec id, equivalent standard, page ranges, figure counts |
| `meta.json` | Conversion settings, measured quality, figure-extraction method |
| `content.json` | Full-fidelity single-file conversion, all 1153 pages |
| `SFA-*/` | The 35 specifications |
| `98-Appendices/` | Mandatory Appendix I |
| `_tools/` | The scripts that produced this drop |

Each spec's JSON carries the edition block, the spec id and title, its **source-standard equivalence** — the AWS specification it reprints (`SFA-5.4/SFA-5.4M` ↔ `A5.4/A5.4M:2012 (R2022)`), with the full "identical with … in case of dispute …" note — plus PDF and printed page ranges, its figure list, and the page-by-page block tree.

**Folders sort alphabetically, not numerically** — `SFA-5.10` lands before `SFA-5.2`. Walk `index.json` to get them in Code order.

## Specifications

| Spec | Equivalent AWS standard | Pages | Figures |
| --- | --- | ---: | ---: |
| `SFA-5.01M/SFA-5.01` | A5.01M/A5.01:2019 (R2024) | 24 |  |
| `SFA-5.02/SFA-5.02M` | A5.02/A5.02M:2025 | 18 | 4 |
| `SFA-5.1/SFA-5.1M` | A5.1/A5.1M:2025 | 54 | 11 |
| `SFA-5.2/SFA-5.2M` | A5.2/A5.2M:2018 | 16 | 1 |
| `SFA-5.3/SFA-5.3M` | A5.3/A5.3M:2023 | 16 | 1 |
| `SFA-5.4/SFA-5.4M` | A5.4/A5.4M:2012 (R2022) | 40 | 10 |
| `SFA-5.5/SFA-5.5M` | A5.5/A5.5M:2022 | 64 | 10 |
| `SFA-5.6/SFA-5.6M` | A5.6/A5.6M:2008 | 20 | 7 |
| `SFA-5.7/SFA-5.7M` | A5.7/A5.7M:2007 | 14 | 1 |
| `SFA-5.8/SFA-5.8M` | A5.8/A5.8M:2011 | 42 | 4 |
| `SFA-5.9/SFA-5.9M` | A5.9/A5.9M:2022 | 34 | 1 |
| `SFA-5.10/SFA-5.10M` | A5.10/A5.10M:2023 | 42 | 6 |
| `SFA-5.11/SFA-5.11M` | A5.11/A5.11M:2018 | 40 | 8 |
| `SFA-5.12/SFA-5.12M` | A5.12/A5.12M:2024 | 22 | 1 |
| `SFA-5.13/SFA-5.13M` | A5.13/A5.13M:2024 | 34 | 1 |
| `SFA-5.14/SFA-5.14M` | A5.14/A5.14M:2024 | 30 |  |
| `SFA-5.15` | A5.15-1990 (S2023) | 26 | 3 |
| `SFA-5.16/SFA-5.16M` | A5.16/A5.16M:2023 | 22 | 1 |
| `SFA-5.17/SFA-5.17M` | A5.17/A5.17M:2019 | 32 | 5 |
| `SFA-5.18/SFA-5.18M` | A5.18/A5.18M:2023 | 36 | 5 |
| `SFA-5.20/SFA-5.20M` | A5.20/A5.20M:2021 | 40 | 6 |
| `SFA-5.21/SFA-5.21M` | A5.21/A5.21M:2024 | 32 | 1 |
| `SFA-5.22/SFA-5.22M` | A5.22/A5.22M:2024 | 54 | 10 |
| `SFA-5.23/SFA-5.23M` | A5.23/A5.23M:2021 | 48 | 8 |
| `SFA-5.24/SFA-5.24M` | A5.24/A5.24M:2023 | 14 |  |
| `SFA-5.25/SFA-5.25M` | A5.25/A5.25M:2023 | 24 | 4 |
| `SFA-5.26/SFA-5.26M` | A5.26/A5.26M:2020 | 24 | 4 |
| `SFA-5.28/SFA-5.28M` | A5.28/A5.28M:2022 | 46 | 4 |
| `SFA-5.29/SFA-5.29M` | A5.29/A5.29M:2022 | 46 | 5 |
| `SFA-5.30/SFA-5.30M` | A5.30/A5.30M:2022 | 26 | 6 |
| `SFA-5.31M/SFA-5.31` | A5.31M/A5.31:2022 | 20 | 2 |
| `SFA-5.32M/SFA-5.32` | A5.32M/A5.32:2021 | 24 |  |
| `SFA-5.34/SFA-5.34M` | A5.34/A5.34M:2018 | 30 | 8 |
| `SFA-5.35/SFA-5.35M` | A5.35/A5.35M:2015-AMD-1 | 16 | 3 |
| `SFA-5.39/SFA-5.39M` | A5.39/A5.39M:2020 | 39 | 9 |

## Figures

150 figures across 31 specifications, PNG at 300 DPI, filed under the spec they belong to. Figure numbers restart in every spec, so `Figure-1.png` exists many times over — always scoped by its folder. A figure printed in both unit systems keeps both in its id (`Figure-1-1M.png`), and a figure continued onto a second sheet is filed as the page titles it (`Figure-3-Continued.png`).

**Method.** Part C's AWS figures have no ruled frame, and most artwork sits inside Form XObjects whose child bounds pdfium reports in form space rather than page space, so neither Section VIII's frame detector nor raw object geometry works and the artwork is located from pixels instead. Three things are particular to this volume. A caption is 'Figure 1-Title' with an em dash, and the dash is what separates it from a cross-reference: 28 of the 170 lines opening with a figure word are references such as 'Figure 2).' or 'Figure 4, except when ...'. A figure usually has a page to itself - drawing, dimension table, notes, caption last - so the caption can sit most of the type area from its own ink, and the figure is several stacked blocks that neither ink weight nor blank paper can join; what does bound it is the region between its caption and the next. Which side of the caption that region lies on follows the reprinted body's house style, measured here rather than assumed: all 142 AWS captions sit below their artwork, and all 8 in the older ASTM style (SFA-5.6, SFA-5.7) above theirs. A numbered table claims its own grid before any caption can, which is what keeps a figure out of the table above it.

**Excluded.** Publisher and commercial branding: cover and back-cover ASME logos, the ANSI accreditation badge, scan watermarks and QR codes. Only artwork carrying a numbered figure title is exported, so none are candidates. Lines that merely open with a figure reference - 'Figure 2).', 'Figure 4, except when ...' - are body text, not captions, and are rejected: 28 of the 170 lines here that begin with a figure word are of that kind.

**Audited.** Every one of the 150 crops was rendered beside its source page, with the crop region outlined, and checked by eye - a clean run (right count, nothing unmatched) is satisfied just as well by a crop of the wrong region. Fifteen crops were wrong across runs that each reported nothing unmatched.

## Known limits

- Text coverage against the PDF's own text layer is **95.5%** overall, or **97.8%** once the running header — which marker drops by design — is discounted. That is the best of the Section II drops.
- The **92 landscape pages average 91%** against **98.5%** for the 1,061 upright ones — a far smaller gap than Part B's, because Part C's landscape pages are chemical-composition tables rather than wide property matrices. The weakest specification is `SFA-5.30/SFA-5.30M` at 91%. Treat the PDF as authoritative for values read off a landscape table.
- Figure crops are bounded by measured ink, not by a drawn frame. Part C sets most figures on a page of their own — drawing, dimension table, notes, caption — and the crop takes all of it, which is what the page titles.
- Tables are preserved as structured blocks inside the JSON, not re-exported as images.
- **Front and back matter are not included.** The cover, contents, list of sections, foreword, statement of policy, personnel, correspondence, summary of changes, cross-referencing pages and the back cover were removed from this drop, so the page counts above sum to less than the source PDF. `content.json` still holds those pages.
