# ASME Boiler and Pressure Vessel Code

**Section II - Materials — Part A - Ferrous Material Specifications (SA-451 to End)**

- **Designation:** ASME BPVC.II.A-2025
- **Edition:** 2025 Edition (edition year 2025)
- **Date of issuance:** 2025-07-01
- **Publisher:** ASME (The American Society of Mechanical Engineers)
- **Source:** `A2-2025.pdf`, 943 pages, complete embedded text layer
- **Specifications:** 114 (SA-451/SA-451M through SA/JIS G5504)
- **Of which international adoptions:** 18 (SA/AS, SA/CSA, SA/EN, SA/GB, SA/IS, SA/ISO, SA/JIS)

This is **Volume 2 of Part A**, filed as `SEC II/A2/`, and it completes Part A alongside `SEC II/A1/` (SA-6 through SA-450). Parts B, C and D are separate PDFs and are not in this drop; each would sit beside these as its own folder under `SEC II/`.

## Layout

**One folder per material specification**, named as the spec is cited with `/` replaced by `-` (a path cannot contain `/`). Each holds its own JSON and, where the spec has any, a `figures/` subfolder.

```
SA-453-SA-453M/
  SA-453-SA-453M.json    the specification, page by page
  figures/FIG-1.png      its figures, 300 DPI
```

| File / folder | What it holds |
| --- | --- |
| `index.json` | Every folder in Code order: spec id, equivalent standard, page ranges, figure counts |
| `meta.json` | Conversion settings, measured quality, figure-extraction method |
| `content.json` | Full-fidelity single-file conversion, all 943 pages |
| `SA-*/` | The 114 specifications, including the SA/EN, SA/GB, SA/IS, SA/ISO, SA/JIS, SA/AS and SA/CSA adoptions |
| `98-Appendices/` | Mandatory Appendices I-IV and Nonmandatory Appendix A |
| `_tools/` | The scripts that produced this drop |

Each spec's JSON carries the edition block, the spec id and title, its **source-standard equivalence** — ASTM for the SA- specs (`SA-451/SA-451M` ↔ `A451/A451M-06(2010)`) and the originating body for the adoptions (`SA/JIS G4303` ↔ `JIS G4303:2012`), with the full "identical with … except …" note — plus PDF and printed page ranges, its figure list, and the page-by-page block tree.

**Folders sort alphabetically, not numerically** — `SA-105` lands before `SA-20`. Walk `index.json` to get them in Code order.

## Specifications

| Spec | Equivalent standard | Pages | Figures |
| --- | --- | ---: | ---: |
| `SA-451/SA-451M` | A451/A451M-06(2010) | 6 |  |
| `SA-453/SA-453M` | A453/A453M-17 | 8 | 1 |
| `SA-455/SA-455M` | A455/A455M-12a(2017) | 4 |  |
| `SA-476/SA-476M` | A476/A476M-00(2018) | 8 | 4 |
| `SA-479/SA-479M` | A479/A479M-21 | 10 |  |
| `SA-480/SA-480M` | A480/A480M-23b | 28 |  |
| `SA-484/SA-484M` | A484/A484M-21 | 14 | 2 |
| `SA-487/SA-487M` | A487/A487M-21 | 8 |  |
| `SA-508/SA-508M` | A508/A508M-18 | 10 |  |
| `SA-513/SA-513M` | A513/A513M-20a | 22 |  |
| `SA-515/SA-515M` | A515/A515M-17 | 4 |  |
| `SA-516/SA-516M` | A516/A516M-17 | 4 |  |
| `SA-517/SA-517M` | A517/A517M-17 | 4 |  |
| `SA-522/SA-522M` | A522/A522M-07 | 6 |  |
| `SA-524/SA-524M` | A524/A524M-21 | 8 |  |
| `SA-530/SA-530M` | A530/A530M-18 | 10 |  |
| `SA-533/SA-533M` | A533/A533M-16 | 4 |  |
| `SA-537/SA-537M` | A537/A537M-20 | 4 |  |
| `SA-540/SA-540M` | A540/A540M-15(2021) | 8 | 1 |
| `SA-541/SA-541M` | A541/A541M-05(2015) | 10 |  |
| `SA-542/SA-542M` | A542/A542M-19 | 6 | 1 |
| `SA-543/SA-543M` | A543/A543M-09(2014) | 4 |  |
| `SA-553/SA-553M` | A553/A553M-17 | 4 |  |
| `SA-556/SA-556M` | A556/A556M-18 | 6 | 1 |
| `SA-562/SA-562M` | A562/A562M-10 | 4 |  |
| `SA-563` | A563-07a(2014) | 12 |  |
| `SA-564/SA-564M` | A564/A564M-04(2009) | 12 |  |
| `SA-568/SA-568M` | A568/A568M-19a | 32 | 5 |
| `SA-572/SA-572M` | A572/A572M-21 | 4 |  |
| `SA-574` | A574-04 | 10 | 3 |
| `SA-587` | A587-96(2005) | 6 |  |
| `SA-592/SA-592M` | A592/A592M-04(2009) | 4 |  |
| `SA-609/SA-609M` | A609/A609M-91(2007) | 14 | 4 |
| `SA-612/SA-612M` | A612/A612M-20 | 4 |  |
| `SA-638/SA-638M` | A638/A638M-00(2004) | 6 |  |
| `SA-645/SA-645M` | A645/A645M-10(2016) | 4 |  |
| `SA-649/SA-649M` | A649/A649M-10 | 6 |  |
| `SA-656/SA-656M` | A656/A656M-18 | 4 |  |
| `SA-660` | A660-96(2010) | 6 |  |
| `SA-662/SA-662M` | A662/A662M-17 | 4 |  |
| `SA-666` | A666-03 | 12 |  |
| `SA-667/SA-667M` | A667/A667M-87(2018) | 4 |  |
| `SA-671/SA-671M` | A671/A671M-20 | 8 |  |
| `SA-672/SA-672M` | A672/A672M-19 | 8 |  |
| `SA-675/SA-675M` | A675/A675M-03(2009) | 6 |  |
| `SA-688/SA-688M` | A688/A688M-15 | 10 | 2 |
| `SA-691/SA-691M` | A691/A691M-19 | 8 |  |
| `SA-693` | A693-22 | 10 |  |
| `SA-696` | A696-90a(2012) | 4 |  |
| `SA-703/SA-703M` | A703/A703M-18a | 14 | 2 |
| `SA-705/SA-705M` | A705/A705M-95(2009) | 10 |  |
| `SA-723/SA-723M` | A723/A723M-10(2015) | 6 |  |
| `SA-724/SA-724M` | A724/A724M-09(2018) | 6 |  |
| `SA-727/SA-727M` | A727/A727M-14(2019) | 6 |  |
| `SA-736/SA-736M` | A736/A736M-17 | 4 |  |
| `SA-737/SA-737M` | A737/A737M-17 | 4 |  |
| `SA-738/SA-738M` | A738/A738M-19 | 6 |  |
| `SA-739` | A739-90a(2016) | 4 |  |
| `SA-747/SA-747M` | A747/A747M-04 | 6 |  |
| `SA-748/SA-748M` | A748/A748M-87(2018) | 4 |  |
| `SA-749/SA-749M` | A749/A749M-97(2002) | 12 |  |
| `SA-751` | A751-21 | 6 |  |
| `SA-765/SA-765M` | A765/A765M-07(2017) | 6 |  |
| `SA-770/SA-770M` | A770/A770M-03(2018) | 6 | 2 |
| `SA-781/SA-781M` | A781/A781M-06 | 20 | 9 |
| `SA-788/SA-788M` | A788/A788M-15 | 14 | 1 |
| `SA-789/SA-789M` | A789/A789M-24 | 8 |  |
| `SA-790/SA-790M` | A790/A790M-24 | 12 |  |
| `SA-803/SA-803M` | A803/A803M-16 | 8 | 2 |
| `SA-813/SA-813M` | A813/A813M-14(2019) | 10 |  |
| `SA-814/SA-814M` | A814/A814M-15(2019) | 8 |  |
| `SA-815/SA-815M` | A815/A815M-10a | 8 |  |
| `SA-832/SA-832M` | A832/A832M-17 | 6 | 1 |
| `SA-834` | A834-95(2015) | 6 |  |
| `SA-836/SA-836M` | A836/A836M-14(2020) | 4 |  |
| `SA-841/SA-841M` | A841/A841M-17 | 10 | 1 |
| `SA-859/SA-859M` | A859/A859M-04(2019) | 6 |  |
| `SA-874/SA-874M` | A874/A874M-98(2018) | 4 |  |
| `SA-905` | A905-19 | 4 |  |
| `SA-941` | A941-22a | 10 |  |
| `SA-960/SA-960M` | A960/A960M-20 | 12 |  |
| `SA-961/SA-961M` | A961/A961M-21 | 10 |  |
| `SA-962/SA-962M` | A962/A962M-22 | 14 | 2 |
| `SA-965/SA-965M` | A965/A965M-21a | 8 |  |
| `SA-985/SA-985M` | A985/A985M-04a | 22 | 10 |
| `SA-988/SA-988M` | A988/A988M-17 | 14 |  |
| `SA-989/SA-989M` | A989/A989M-18 | 10 |  |
| `SA-995/SA-995M` | A995/A995M-20 | 6 |  |
| `SA-999/SA-999M` | A999/A999M-18 | 12 |  |
| `SA-1008/SA-1008M` | A1008/A1008M-24 | 14 | 1 |
| `SA-1010/SA-1010M` | A1010/A1010M-01(2009) | 4 |  |
| `SA-1011/SA-1011M` | A1011/A1011M-23 | 10 |  |
| `SA-1016/SA-1016M` | A1016/A1016M-17a | 12 |  |
| `SA-1017/SA-1017M` | A1017/A1017M-17 | 4 |  |
| `SA-1058` | A1058-19 | 12 |  |
| `SA-1091/SA-1091M` | A1091/A1091M-24 | 10 |  |
| `SA/AS 1548` | AS 1548-2008 | 2 |  |
| `SA/CSA-G40.21` | CSA-G40.21-2013 | 2 |  |
| `SA/EN 10025-2` | EN 10025-2:2019 | 4 |  |
| `SA/EN 10028-2` | EN 10028-2:2017 | 4 |  |
| `SA/EN 10028-3` | EN 10028-3:2017 | 2 |  |
| `SA/EN 10028-4` | EN 10028-4:2017 | 4 |  |
| `SA/EN 10028-7` | EN 10028-7:2016 | 2 |  |
| `SA/EN 10088-2` | EN 10088-2:2014 | 2 |  |
| `SA/EN 10088-3` | EN 10088-3:2014 | 4 |  |
| `SA/EN 10216-2` | EN 10216-2:2013 | 2 |  |
| `SA/EN 10217-1` | EN 10217-1:2019 | 2 |  |
| `SA/EN 10222-2` | EN 10222-2:2017 | 2 |  |
| `SA/GB 713` | GB 713-2014 | 2 |  |
| `SA/IS 2062` | IS 2062-2011 | 2 |  |
| `SA/ISO 898-1` | ISO 898-1:2013 | 2 |  |
| `SA/JIS G3118` | JIS G3118:2017 | 4 |  |
| `SA/JIS G4303` | JIS G4303:2012 | 2 |  |
| `SA/JIS G5504` | JIS G5504:2005 | 2 |  |

## Figures

55 figures across 20 specifications, PNG at 300 DPI, filed under the spec they belong to. Figure numbers restart in every spec, so `FIG-1.png` exists many times over — always scoped by its folder.

**Method.** Section II's ASTM figures have no ruled frame, caption placement varies between specs, and most artwork sits inside Form XObjects whose child bounds pdfium reports in form space rather than page space. Neither Section VIII's frame detector nor raw object geometry works, so the artwork is located from pixels instead.

**Excluded.** Publisher and commercial branding: cover and back-cover ASME logos, the ANSI accreditation badge, scan watermarks and QR codes. Only artwork paired with a numbered 'FIG. n' caption is exported, so none are candidates.

**Not extracted (1).** These captions were found but their artwork could not be isolated — each shares a run of ink with the figure next to it. The figures themselves are in `content.json`'s page tree and in the source PDF:

- PDF p927 — Figure II-400-1

## Known limits

- Text coverage against the PDF's own text layer is **96.8%** overall.
- The 34 landscape pages average **85.7%** against **97.1%** for the 909 upright pages; the weakest single spec is `SA-705/SA-705M` at 89%. Those are the wide alloy/mechanical-property tables — treat the PDF as authoritative for values read off them.
- Figure crops are bounded by measured ink, not by a drawn frame, so a few carry an adjoining metric-equivalents table or a stray line of text. Nothing is cut off; some crops are simply more generous than others.
- **One further figure is not extracted and is not in the list above:** `FIG. X1.2` on PDF p669 (`SA-841/SA-841M`). Its page is landscape, and the caption assembler groups characters by shared baseline, which only holds for upright text — so on a rotated page the caption is never recognised in the first place and never reaches the unmatched list. 57 captions exist in this volume; 55 were extracted.
- Tables are preserved as structured blocks inside the JSON, not re-exported as images.
- **Front and back matter are not included.** The cover, contents, list of sections, foreword, statement of policy, personnel, correspondence, summary of changes, cross-referencing pages and the back cover were removed from this drop, so the page counts above sum to less than the source PDF. `content.json` still holds those pages.
