# ASME B31.3-2024 extractor

Extracts the narrative, tabular and figure content of ASME B31.3-2024 to JSON, with the
figures rendered to PNG. Two runs share one core:

- **Appendices E-Z** (`run.py`) -> `APPEX/`. Appendices A, B and C were extracted earlier by
  a different, tables-only script and are not touched.
- **The body of the Code, Chapters I-X** (`run_body.py`) -> `CHAPTERS/`. PDF idx 33-196.
  Publisher front matter and the INDEX are out of scope.

## Running it

```bash
python run.py                 # appendices E-Z -> APPEX/, extends the index manifest
python validate.py <dir>      # appendix checks

python run_body.py            # Chapters I-X -> CHAPTERS/
python validate_body.py       # body checks, against the PDF's own bookmark tree
```

`b313.SRC` is the source PDF path and `run.DEST` the output folder; both are constants at the
top of their modules. `run.py` re-reads the existing `asme_b31_3_2024_index.json`, keeps the
A/B/C entries byte-identical, and replaces the rest.

Requires `pymupdf` only.

## Modules

| File | What it does |
|---|---|
| `b313.py` | Source constants, and the line reconstruction everything else is built on: character-level text rebuilding, word-space restoration, page-furniture stripping, cell typing |
| `tables.py` | Ruled-table detection - column grid from the printed rules, region splitting, header/body separation |
| `prose.py` | Section tree from the heading fonts, paragraph and list-item structure, equation assembly, vocabulary-checked de-hyphenation |
| `figures.py` | Figure location, clip-rect derivation from the line art, rendering (including the landscape page) |
| `build.py` | Generic per-appendix assembly: tables, figures, prose, note collection |
| `special.py` | Appendices E, J and K, whose layouts the generic engines do not cover |
| `run.py` | Appendix driver: inventory, output schemas, index manifest |
| `rotate.py` | Upright view of a landscape sheet (see below) |
| `body.py` | The body of the Code: rotation-aware tables, raster-aware figures, run-in headings, the paragraph-number tree, the 300.2 glossary |
| `run_body.py` | Body driver: chapter inventory, output schemas, chapters manifest |
| `validate.py` | Post-extraction checks, appendices |
| `validate_body.py` | Post-extraction checks, body |

## The five source facts that drive the design

These are properties of this PDF, established by measurement. Anything reused against another
ASME document should be re-measured rather than assumed.

1. **The horizontal rules are the column grid.** Each rule is emitted as one segment per
   column, so the rule with the most segments gives the exact column boundaries. This is much
   safer than inferring columns from whitespace.
2. **Justified lines fail in two opposite ways.** Some drop the space glyph and push words
   apart by ~1.4 pt; others keep the spaces and track every letter out by ~0.6 pt. The
   threshold is therefore derived per line from that line's own median letter gap. Fully
   tracked-out lines (URLs in Appendix N) are detected by space ratio and collapsed.
3. **Minus signs in equations are drawn as vector strokes, not text.** They must be restored
   from the drawing list or the equations silently lose their subtraction.
4. **Displayed maths is a scatter of positioned glyph runs**, mixing maths fonts with plain
   Cambria for units. Equations are found geometrically and linearised into reading order;
   they are stored as strings, never parsed.
5. **There are no raster images in the appendix range.** Every figure is line art and has to
   be rendered from a clip rect.

## Traps that cost real data, worth keeping in mind

- Do not strip "a bare number low on the page" as a printed folio. The folio sits at y=745,
  outside the content band already; a rule like that deletes the last row of any table that
  falls low on the page, and those rows are all-numeric.
- Fraction bars in equations look like table rules. A table rule spans the table; anything
  under 100 pt wide is maths.
- The head/body split is not a fixed rule index - heads run to one, two or three tiers. Take
  the first rule below the last bold label.
- Line art contains its own rules, so a ruled region is only a table if the caption above it
  says "Table".
- A page's stray hairlines will stretch a figure's clip down over the prose beneath it.
  Cluster the drawings and keep only the run continuous with the caption.


## What the body of the Code needed that the appendices did not

The appendices are tables, lists and worked examples. The body is the Code itself, and five
further properties of the source had to be handled. All were measured, not assumed.

1. **PyMuPDF reports text and drawings in unrotated page space whatever the page's /Rotate
   says** - only `page.rect` changes. On the seven landscape sheets the printed horizontal
   rules therefore arrive as tall thin vertical rectangles and the rule-grid table engine
   finds nothing at all. `rotate.UprightPage` turns the coordinates back
   (`upright_x = page_height - y`, `upright_y = x`, measured on idx 120) so the table, prose
   and figure engines see an ordinary page. Rendering is the exception: it already honours
   /Rotate, and MuPDF's rotated space is exactly this upright space, so the clip is passed
   through untouched and no correcting matrix is applied.
2. **The running head and folio do not turn with the table.** They stay in the sheet's
   orientation, so on a landscape page they bound the content in x, not in y. The band is
   carried on the page as `content_box` for that reason.
3. **Twenty pages carry the figure as an embedded raster image** and have an empty drawing
   list; the appendix code, which looks only at drawings, finds no figure there at all. Both
   sources are unioned, and the render dpi is raised to whatever the placed image actually
   holds (up to 600), because rendering a 3,892 px image at the appendices' 300 dpi throws
   away more than half of it.
4. **Headings are set run-in** - "326.1.1 Listed Piping Components." in bold with the
   requirement continuing in roman on the same line - so a whole-line font decision puts the
   first words of the text into the heading and starts the paragraph mid-word. Characters
   are tagged with their span's font and the line is split at the boundary.
5. **Group banners are set in the same bold face as the column labels.** Table 326.1.1-1
   sorts its standards under "Bolting", "Metallic Fittings, Valves, and Flanges" and so on.
   The appendix head/body rule ("the first rule below the last bold label") walks into the
   data on those tables, so the split is taken as the last rule with nothing but bold above
   it, and the banners are recorded as a `group` on the rows they head.

## Traps that cost real data in the body run

Three of these deleted content silently. They are listed because none of them failed - they
all produced plausible-looking output.

- **Any two rules on a page were treated as one table.** Page 89 carries the closing rule of
  a table continued from the page before and, 466 pt lower, the rule under an unruled inline
  list, with two columns of prose between. The two were glued into one region covering the
  whole left column, the prose pass was told to skip it, and para. 323.3 disappeared. A
  numbered section heading between two rules proves they belong to different things.
- **Figure artwork was clustered by vertical proximity alone.** That is safe on a page that
  is nothing but a figure, which is all the appendices contain. Page 57 sets the miter-bend
  image in the left column, and vertical clustering pulled in vector strokes from the right
  column at the same height, stretching the clip across both columns and taking para.
  304.2.3 with it. Cluster in x as well as y.
- **The figure clip rect was kept only for multi-sheet figures**, so all 39 single-sheet
  figures were never excluded from the prose pass and their captions and internal callouts
  came through as stray sections of the Code.
- A table's caption sits above its first rule and its NOTES below the last one, both set in
  the same faces as the surrounding prose. Excluding only the ruled grid drops a table's
  notes into the running text of whatever paragraph follows.
- Continuation lines in a table's key column are not distinguishable by their text -
  "N088xx and N066xx nickel" is a continuation though it opens with a capital, and "Other
  materials [Note (9)]" is a row of its own though it carries no values. ASME hangs the
  continuations by one em; the indent is the signal.
- The de-hyphenator was applied to paragraph text but not to headings, leaving twelve
  headings reading "Reinforcement of Welded Branch Connec- tions".
