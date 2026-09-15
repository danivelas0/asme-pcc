# ASME Boiler and Pressure Vessel Code

**Section II - Materials — Part A - Ferrous Material Specifications (Beginning to SA-450)**

- **Designation:** ASME BPVC.II.A-2025
- **Edition:** 2025 Edition (edition year 2025)
- **Date of issuance:** 2025-07-01
- **Publisher:** ASME (The American Society of Mechanical Engineers)
- **Source:** `A1-2025.pdf`, 803 pages, complete embedded text layer
- **Specifications:** 76 (SA-6/SA-6M through SA-450/SA-450M)

This is **Volume 1 of Part A**, filed as `SEC II/A1/`. The companion volume (SA-451 onward) and Parts B, C and D are separate PDFs and are not in this drop; each would sit beside this one as its own folder under `SEC II/`.

## Layout

**One folder per material specification**, named as the spec is cited with `/` replaced by `-` (a path cannot contain `/`). Each holds its own JSON and, where the spec has any, a `figures/` subfolder.

```
SA-370/
  SA-370.json          the specification, page by page
  figures/FIG-1.png    its figures, 300 DPI
```

| File / folder | What it holds |
| --- | --- |
| `index.json` | Every folder in Code order: spec id, ASTM designation, page ranges, figure counts |
| `meta.json` | Conversion settings, measured quality, figure-extraction method |
| `content.json` | Full-fidelity single-file conversion, all 803 pages |
| `SA-*/` | The 76 specifications |
| `_tools/` | The scripts that produced this drop |

Each spec's JSON carries the edition block, the spec id and title, its **ASTM equivalence** (e.g. SA-370 ↔ `A370-21`, with the full "identical with … except …" note), PDF and printed page ranges, its figure list, and the page-by-page block tree.

**Folders sort alphabetically, not numerically** — `SA-105` lands before `SA-20`. Walk `index.json` to get them in Code order.

## Specifications

| Spec | ASTM | Pages | Figures |
| --- | --- | ---: | ---: |
| `SA-6/SA-6M` | A6/A6M-21 | 64 |  |
| `SA-20/SA-20M` | A20/A20M-20 | 34 |  |
| `SA-29/SA-29M` | A29/A29M-20 | 18 |  |
| `SA-31` | A31-14 | 6 | 2 |
| `SA-36/SA-36M` | A36/A36M-19 | 4 |  |
| `SA-47/SA-47M` | A47/A47M-99(2018) | 10 | 3 |
| `SA-53/SA-53M` | A53/A53M-20 | 24 | 1 |
| `SA-105/SA-105M` | A105/A105M-21 | 6 |  |
| `SA-106/SA-106M` | A106/A106M-19a | 10 |  |
| `SA-134/SA-134M` | A134/A134M-19 | 6 |  |
| `SA-135/SA-135M` | A135/A135M-21 | 10 |  |
| `SA-178/SA-178M` | A178/A178M-19 | 6 |  |
| `SA-179/SA-179M` | A179/A179M-19 | 4 |  |
| `SA-181/SA-181M` | A181/A181M-06 | 6 |  |
| `SA-182/SA-182M` | A182/A182M-23 | 18 |  |
| `SA-192/SA-192M` | A192/A192M-17 | 4 |  |
| `SA-193/SA-193M` | A193/A193M-20 | 14 |  |
| `SA-194/SA-194M` | A194/A194M-22 | 14 |  |
| `SA-203/SA-203M` | A203/A203M-17 | 4 |  |
| `SA-204/SA-204M` | A204/A204M-17 | 4 |  |
| `SA-209/SA-209M` | A209/A209M-03(2017) | 4 |  |
| `SA-210/SA-210M` | A210/A210M-19 | 4 |  |
| `SA-213/SA-213M` | A213/A213M-23 | 18 |  |
| `SA-214/SA-214M` | A214/A214M-19 | 4 |  |
| `SA-216/SA-216M` | A216/A216M-21 | 6 |  |
| `SA-217/SA-217M` | A217/A217M-22 | 6 |  |
| `SA-225/SA-225M` | A225/A225M-17 | 4 |  |
| `SA-231/SA-231M` | A231/A231M-96 | 6 |  |
| `SA-232/SA-232M` | A232/A232M-91 | 6 |  |
| `SA-234/SA-234M` | A234/A234M-23a | 12 |  |
| `SA-240/SA-240M` | A240/A240M-17 | 16 |  |
| `SA-249/SA-249M` | A249/A249M-18a | 12 |  |
| `SA-250/SA-250M` | A250/A250M-05(2014) | 6 |  |
| `SA-263` | A263-12(2019) | 6 | 3 |
| `SA-264` | A264-12(2019) | 6 | 3 |
| `SA-265` | A265-12(2019) | 8 | 3 |
| `SA-266/SA-266M` | A266/A266M-21 | 6 |  |
| `SA-268/SA-268M` | A268/A268M-20 | 8 |  |
| `SA-276` | A276/A276M-17 | 10 |  |
| `SA-278/SA-278M` | A278/A278M-01(2015) | 6 | 2 |
| `SA-283/SA-283M` | A283/A283M-18 | 4 |  |
| `SA-285/SA-285M` | A285/A285M-17 | 4 |  |
| `SA-299/SA-299M` | A299/A299M-17 | 4 |  |
| `SA-302/SA-302M` | A302/A302M-17 | 4 |  |
| `SA-307` | A307-10 | 6 |  |
| `SA-311/SA-311M` | A311/A311M-04(2015) | 6 |  |
| `SA-312/SA-312M` | A312/A312M-24 | 14 |  |
| `SA-320/SA-320M` | A320/A320M-22 | 8 |  |
| `SA-325` | A325-10 | 10 |  |
| `SA-333/SA-333M` | A333/A333M-16 | 10 |  |
| `SA-334/SA-334M` | A334/A334M-04a(2016) | 12 |  |
| `SA-335/SA-335M` | A335/A335M-24 | 12 |  |
| `SA-336/SA-336M` | A336/A336M-23 | 10 |  |
| `SA-350/SA-350M` | A350/A350M-23 | 12 | 4 |
| `SA-351/SA-351M` | A351/A351M-18 | 8 |  |
| `SA-352/SA-352M` | A352/A352M-06(2012) | 6 |  |
| `SA-353/SA-353M` | A353/A353M-17 | 4 |  |
| `SA-354` | A354-11 | 8 |  |
| `SA-358/SA-358M` | A358/A358M-19 | 10 |  |
| `SA-369/SA-369M` | A369/A369M-23 | 8 |  |
| `SA-370` | A370-21 | 50 | 15 |
| `SA-372/SA-372M` | A372/A372M-20 | 6 |  |
| `SA-376/SA-376M` | A376/A376M-19 | 8 |  |
| `SA-387/SA-387M` | A387/A387M-17a | 6 |  |
| `SA-395/SA-395M` | A395/A395M-99(2018) | 14 | 6 |
| `SA-401/SA-401M` | A401/A401M-18 | 6 |  |
| `SA-403/SA-403M` | A403/A403M-19a | 12 |  |
| `SA-409/SA-409M` | A409/A409M-19 | 8 |  |
| `SA-414/SA-414M` | A414/A414M-14(2019) | 6 | 1 |
| `SA-420/SA-420M` | A420/A420M-19a | 8 |  |
| `SA-423/SA-423M` | A423/A423M-19 | 4 |  |
| `SA-426/SA-426M` | A426/A426M-13 | 6 |  |
| `SA-437/SA-437M` | A437/A437M-15(2021) | 4 |  |
| `SA-439/SA-439M` | A439/A439M-18 | 8 | 5 |
| `SA-449` | A449-10 | 8 |  |
| `SA-450/SA-450M` | A450/A450M-21 | 12 |  |

## Figures

48 figures across 12 specifications, PNG at 300 DPI, filed under the spec they belong to. Figure numbers restart in every spec, so `FIG-1.png` exists many times over — always scoped by its folder.

**Method.** Section II's ASTM figures have no ruled frame, caption placement varies between specs, and most artwork sits inside Form XObjects whose child bounds pdfium reports in form space rather than page space. Neither Section VIII's frame detector nor raw object geometry works, so the artwork is located from pixels instead.

**Excluded.** Publisher and commercial branding: cover and back-cover ASME logos, the ANSI accreditation badge, scan watermarks and QR codes. Only artwork paired with a numbered 'FIG. n' caption is exported, so none are candidates.

**Not extracted (4).** These captions were found but their artwork could not be isolated — each shares a run of ink with the figure next to it. The figures themselves are in `content.json`'s page tree and in the source PDF:

- PDF p642 — FIG. 3 Rectangular Tension Test Specimens
- PDF p643 — FIG. 4 Standard 0.500-in. (12.5 mm) Round Tension Test Specimen With 2
- PDF p716 — FIG. 6 STANDARD
- PDF p781 — FIG. 6 Examples of Small

## Known limits

- Text coverage against the PDF's own text layer is **97.0%** overall — the highest of the BPVC drops so far.
- The 41 landscape pages average **88.1%** against **97.5%** for the 762 upright pages; the weakest single spec is `SA-240/SA-240M` at 87%. Those are the wide alloy/mechanical-property tables — treat the PDF as authoritative for values read off them.
- Figure crops are bounded by measured ink, not by a drawn frame, so a few carry an adjoining metric-equivalents table or a stray line of text. Nothing is cut off; some crops are simply more generous than others.
- Tables are preserved as structured blocks inside the JSON, not re-exported as images.
- **Front and back matter are not included.** The cover, contents, list of sections, foreword, statement of policy, personnel, correspondence, summary of changes, cross-referencing pages and the back cover were removed from this drop, so the page counts above sum to less than the source PDF. `content.json` still holds those pages.
