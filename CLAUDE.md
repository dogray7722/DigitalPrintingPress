# Digital Printing Press

A Vite + React + TypeScript web app that generates downloadable travel planning `.xlsx` files for resale on Etsy/Gumroad. The user fills out a 4-step wizard; the app builds a multi-sheet workbook and triggers a browser download.

## Tech Stack

| Layer          | Technology                                                                                    |
| -------------- | --------------------------------------------------------------------------------------------- |
| Build          | Vite 6 (ESM)                                                                                  |
| UI             | React 18 + TypeScript 5                                                                       |
| Styling        | Tailwind CSS 3 (no component library — built from scratch)                                    |
| State          | Zustand 5 (wizard state)                                                                      |
| Spreadsheet    | ExcelJS 4 + JSZip 3 (chart injection)                                                         |
| Browser compat | vite-plugin-node-polyfills (installed with `--legacy-peer-deps` due to Vite 6 peer conflict)  |
| Icons          | lucide-react                                                                                  |
| Date math      | date-fns                                                                                      |
| AI             | @anthropic-ai/sdk — server-side via custom Vite middleware (`vite-plugin-recommendations.ts`) |

**Dev server:** `npm run dev` → http://localhost:5173
**Build:** `npm run build` (runs `tsc -b && vite build`)

## Project Purpose

A "digital printing press" — the user configures a trip (destination, dates, budget, style, sheet selection) and receives a professional multi-sheet travel planner spreadsheet as an instant download. Target buyers are travelers; the seller lists generated files on Etsy/Gumroad.

## Wizard Flow

Four steps, state in `src/store/wizardStore.ts` (Zustand):

1. **Step 1 — Destination** (`Step1Destination.tsx`): destination free-text, trip dates, number of adults, AI recommendations toggle (`useRecommendations`)
2. **Step 2 — Budget** (`Step2Budget.tsx`): per-category estimated budgets, currency selector
3. **Step 3 — Style** (`Step3Style.tsx`): color theme, font family, chart style, sheet selection toggles
4. **Step 4 — Preview & Generate** (`Step4Preview.tsx`): summary card, "Generate Spreadsheet" button

Navigation and validation live in `src/hooks/useWizardNavigation.ts`. When AI recs are on, only `destination` is required to advance from Step 1 (departure/return cities were removed).

## Generated Sheets (tab order)

All sheet builders are in `src/lib/excel/sheets/`. The entry point is `src/lib/excel/index.ts`. Tab order matches build-call order in `index.ts`; UI toggle order in `src/types/wizard.ts` (`ALL_SHEET_IDS`) mirrors it minus always-on sheets.

| #   | Sheet          | Always on | Notes                                                                                                                                            |
| --- | -------------- | --------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| 1   | QUICK START    | ✓         | Always **first**; not toggleable. The short onboarding page — numbered steps for dates (OVERVIEW B6/D6), estimated budget (H15:H20), the ITINERARY draft, the tracker "★ Recommendations →" chips, and logging spend. Steps are conditional and auto-numbered by `stepHeader()`, so copy must never name a step by number. Left unprotected (no formulas). No gridlines. `hasRecommendationsFor()` from `recommendationInsert.ts` gates the recommendations step, so it can only describe guides that actually exist |
| 2   | OVERVIEW       | ✓         | Named ranges: TripStart B6, TripEnd D6, NumAdults E9, TotalBudget D16; rows 1–14 are a dashboard "hero band" (dark B:E text panel + G1:L11 cover photo, no gridlines/freeze — see below) instead of a title bar + label/value ledger; native DrawingML chart(s) injected via JSZip post-processing; trip-readiness doughnut ring (B25:E34) with a centered `T1/T3` "% ready" label floated over its transparent hole, quick-nav internal-hyperlink strip (B38+, each link prefixed with its `SHEET_ICONS` emoji — the same map the Step 3 toggle grid uses, exported from `src/types/wizard.ts`), neighborhood guide from AI regions (G38+, when recs enabled). Exactly one image (the cover photo) — a second upload + multi-image strip was tried and removed: it served no purpose and arbitrary source dimensions made the results unpredictable |
| 3   | ITINERARY      | ✓         | Auto-dates from TripStart+n; AI-prefilled when recs on                                                                                           |
| 4   | TRANSPORT      | toggle    | SheetId = `flights`; Mode dropdown: Air/Train/Bus/Car/Ferry/Taxi/Other                                                                           |
| 5   | HOTELS         | toggle    | Date col auto-fills from TripStart; AI recommendation guide table appended below tracker                                                         |
| 6   | RESTAURANTS    | toggle    | AI recommendation guide table appended below tracker                                                                                             |
| 7   | EXCURSIONS     | toggle    | Per-person × NumAdults; AI recommendation guide table appended below tracker                                                                     |
| 8   | BUDGET TRACKER | toggle    | Date / Category dropdown / Description / Amount / Running Total / Notes                                                                          |
| 9   | PACKING LIST   | toggle    | Seasonal packing + COUNTIF progress formula; `Packed?` (C5:C500) is a one-option `"✓"` dropdown (see below)                                       |
| 10  | TASKS          | toggle    | `Done?` (D5:D500) is a one-option `"✓"` dropdown (see below); Priority dropdown on E5:E500                                                        |
| 11  | ANNUAL EVENTS  | AI-gated  | Only generated when `useRecommendations` is on AND data returned                                                                                 |
| 12  | INSTRUCTIONS   | ✓         | Always last; not toggleable                                                                                                                      |

### Check-mark dropdowns (PACKING LIST / TASKS)

`Packed?` (`C5:C500`) and `Done?` (`D5:D500`) each carry a **single-option list data
validation whose only value is `✓`** — picking from a dropdown beats making the buyer find
the character in Windows' emoji picker or Mac's Character Viewer, and INSTRUCTIONS no
longer explains how to type it. `allowBlank: true` is what lets a cell be cleared to
un-check a row, so keep it.

The literal is load-bearing in three places that must stay in sync: the dropdown option,
each sheet's `A2` COUNTIF progress formula, and OVERVIEW's readiness-ring counts
(`overview.ts` `T1`). Changing the glyph in one place silently pins the progress bars at
0%. `scripts/verify-check-dropdowns.mjs` asserts the dropdown, the COUNTIF, and the raw
sheet XML all agree — note ExcelJS XML-escapes the wrapping quotes (`&quot;✓&quot;`), which
is the same shape Excel writes for an inline list, so assert on either form.

Excel needs the target cells **unlocked** for a dropdown to be usable on a protected
sheet — both columns already are (see `protection.ts`).

## AI Recommendations

Backend: `vite-plugin-recommendations.ts` (root) registers a `/api/recommendations` POST endpoint in Vite's dev middleware. Uses `@anthropic-ai/sdk` with model `claude-sonnet-4-6` and the `web_search_20250305` tool.

Requires `ANTHROPIC_API_KEY` in `.env` (see `.env.example`).

**What the API returns:** `{ regions, itinerary, events }`

- `regions`: `{ region, description, hotels[], restaurants[], excursions[] }[]` — each place is a `PlaceRec { name, price?, rating?, cuisine?, duration?, description? }`
- AI output is **informational only** — it does NOT prefill HOTELS/TRANSPORT/RESTAURANTS/EXCURSIONS tracker rows
- ITINERARY IS still AI-prefilled
- ANNUAL EVENTS tab is still built from AI data

**Recommendation insert** (`src/lib/excel/sheets/recommendationInsert.ts`): `buildRecommendationInsert(ws, startRow, ts, regions, kind)` renders a region-grouped guide table BELOW the working tracker on each of HOTELS/RESTAURANTS/EXCURSIONS. The `kind` picks the array and attribute columns (hotels=Price+Stars, restaurants=Cuisine+Price, excursions=Duration+Price). The places loop **must** `row++` per place or ExcelJS throws "Cannot merge already merged cells."

**Recommendations link.** The guide lands far below the tracker (row 28 on ACCOMMODATION, 58 on DINING, 38 on EXCURSIONS), so each tracker carries a "★ Recommendations →" chip to it. It's a **native internal hyperlink** (`cell.value = { text, hyperlink: "'SHEET'!A28" }`), the same form OVERVIEW's quick-nav strip uses — see "Native Internal Hyperlinks" below. No formula means no cached `result:` obligation. (There is deliberately **no** "back to top" link — it was tried and removed.)

- The link sits in the **last two cells of the row-2 description band**, merged — `F2:G2` / `G2:H2` / `I2:J2` — which is inside every tracker's frozen pane (`ySplit: 5`), so it stays on screen at any scroll depth.
- **Row 2's merge stops two columns short** (`A2:E2` / `A2:F2` / `A2:H2`) to free that range. The ACCOMMODATION and DINING descriptions were **trimmed** to fit: at the band's `sectionHeader` size (13pt) a width unit holds only ~0.8 chars, and there is no wrap on this row (height 22). Caps are roughly **73 / 93 / 102** characters. Re-lengthening a description past its cap will clip it mid-word.
- `reserveRecommendationsSlot(ws, ts, range)` is called **unconditionally**, right after `styleSectionHeader(ws.getRow(2), ts)`. It merges the range and paints the band fill — `styleSectionHeader` skips empty cells, so without it the band shows a white notch whenever recommendations are off.
- `buildRecommendationsLink(ws, ts, ref, targetRow)` styles it as a **button chip** that deliberately contrasts with the band: dark `primary` fill against the band's mid-tone `secondary`, bold accent text at `sectionHeader` size with a matching thin frame on all four sides, centered, **no underline** (it read as a dated hyperlink squiggle; fill + frame + ★/→ already signal clickable). Must run **after** `styleSectionHeader` so the overrides win. The label `★   Recommendations   →` uses plain BMP characters, not emoji, so the star and arrow take the accent color; they stay in one string because a cell carries a single hyperlink across its whole value (rich-text runs would drop the link).
- **The chip's accent color is measured, not fixed.** Tinting text and frame with the band's own `secondary` ties the chip to the band and is the intended look on **inkwell** (gold on charcoal, 6.68:1) — but that pairing is illegible on every other theme (1.60–2.84:1, below even the 3:1 large-text floor; `desert` is worst). `buildRecommendationsLink` calls `contrastRatio()` from `styleFactory.ts` and falls back to `primaryText` when `secondary` doesn't clear `AA_CONTRAST`, so inkwell gets the designed chip and everything else stays readable. `accent` was evaluated as a middle option and rejected — still fails on sakura/desert/parchment. The frame follows the same decision rather than staying gold unconditionally: a frame at 1.6:1 isn't a frame. **A theme added later is handled automatically** — and `verify-rec-jump-links.mjs` builds a workbook per theme and asserts the resulting pair clears 3:1, so a low-contrast `secondary` fails the suite instead of silently shipping unreadable text.
- `verify-rec-jump-links.mjs` also asserts the chip's fill and text color **differ from `A2`'s** rather than checking hardcoded hex, so the chip-vs-band contrast holds for every theme.

`buildRecommendationInsert` returns `number | null` — the section-band row, or `null` on either of its two no-op paths. **Callers must gate the link on that return value, not on `regions`**: the guide also no-ops when `regions` is non-empty but holds no places of that `kind`, and a link into a blank row is worse than no link. Verified by `scripts/verify-rec-jump-links.mjs`, including the negative case (no link, but the band fill survives).

## Native Internal Hyperlinks

`src/lib/excel/nativeHyperlinks.ts` (`injectNativeHyperlinks`, runs in `index.ts` after `injectChart`) rewrites ExcelJS's internal (same-workbook) hyperlinks into the shape Excel itself writes — **the only shape Google Sheets imports correctly.**

There are two ways to make an internal link, and they are not equivalent:

| | `HYPERLINK("#'SHEET'!A1", …)` formula | native `{ text, hyperlink: "'SHEET'!A1" }` |
| --- | --- | --- |
| Excel | navigates | navigates |
| Google Sheets | **inert text** — Sheets doesn't resolve the fragment | navigates |
| calcChain contract | needs a correct cached `result:` | n/a — not a formula |

ExcelJS *does* detect an internal link (its `isInternalLink` regex matches `Sheet!A1` / `'Sheet Name'!A1`) and emits `location`, but it gets two things wrong:

1. It **also** emits an `r:id` plus a matching `TargetMode="External"` relationship whose Target is the sheet reference. A hyperlink is either internal (`location`) or external (`r:id`), never both — Excel tries to resolve `'ACCOMMODATION'!A28` as a file.
2. It never emits **`display`**. That one is why this matters for Sheets: without it, Google renders the link's text as a literal `#gid=xxxx` instead of the label (PHPOffice/PhpSpreadsheet#807 — same bug class, different library).

```
ExcelJS:  <hyperlink ref="F2" r:id="rId1" location="'ACCOMMODATION'!A28"/>
          + <Relationship Id="rId1" Type=".../hyperlink" Target="'ACCOMMODATION'!A28" TargetMode="External"/>
Excel:    <hyperlink ref="F2" location="'ACCOMMODATION'!A28" display="★   Recommendations   →"/>
```

The pass recovers `display` from the cell's own value (via `sharedStrings.xml`), so it is self-contained — every internal hyperlink in the workbook is fixed up with no registry threaded through the sheet builders.

**Ordering: it must run AFTER `injectChart`.** Chart injection allocates its worksheet rel as `maxRelId(rels) + 1`; deleting hyperlink rels first would lower that max and let the new rel collide with the `<drawing r:id="rIdN"/>` reference ExcelJS already wrote into the sheet XML. Deleting afterwards is safe — rIds needn't be contiguous, and nothing is renumbered.

**ExcelJS's reader can't see these links.** It binds a hyperlink to its cell only when the element carries an `r:id` (`worksheet-xform.js`: `if (hyperlink.rId)`), so an r:id-less native link is parsed but never surfaced on the cell — this affects genuine Excel-authored files too. Assert against the raw XML, not a `wb.xlsx.load()` round-trip (`verify-rec-jump-links.mjs` does exactly this). The label still round-trips as the cell's own string value, so a link-ignoring consumer shows the words, not a blank cell.

**Confirmed working in Google Sheets** (manual Drive import, Aug 2026): the tracker links import as real internal links — Sheets resolves the target (`DINING!A58`) and renders the `display` label, not `#gid=xxxx`.

**Google Sheets needs two clicks to follow ANY link, and the file cannot change that.** Clicking a linked cell opens a preview chip; following the link is a second click on the chip. This is universal Sheets behavior for every link in every spreadsheet — native or `HYPERLINK()` formula — not something `location`/`display` influences. The only mitigations are viewer-side and partial (Docs → Tools → Preferences → "Show Link Details" off shrinks the card but a small chip remains; publish-to-web gives single-click but doesn't apply to a downloaded file). Excel follows on one click. **Don't try to "fix" this in the workbook** — there is nothing to fix.

**Every internal link in the workbook now uses this form** — the three tracker chips, OVERVIEW's 8-row quick-nav strip, and QUICK START's "where to go next" rows. No `HYPERLINK()` formula links remain; don't reintroduce one.

**OVERVIEW is the delicate case, and it works because nothing is renumbered.** ExcelJS assigns rIds to hyperlinks *before* drawings (`worksheet-xform.js`), so the cover-photo drawing lands on `rId9` behind the 8 quick-nav rels, and `chartInjection` then merges its charts into that existing drawing part. Deleting the hyperlink rels afterwards leaves `<drawing r:id="rId9"/>` resolving fine — rIds needn't be contiguous. `verify-quick-nav-icons.mjs` asserts exactly that (`OVERVIEW r:id rId9 still resolves`), so a future change that renumbers instead of deleting fails the suite rather than shipping a repair prompt.

**OVERVIEW is not `sheet1.xml`** — QUICK START is the first tab, so OVERVIEW's part is `sheet2.xml`. Nothing in the pipeline cares (`chartInjection` resolves its target by name through `workbook.xml` + its rels, and `injectNativeHyperlinks` walks every sheet part), but a verify script that hardcodes a part path will silently assert against the wrong sheet. Resolve by name — `verify-sheet-protection.mjs`, `verify-quick-nav-icons.mjs` and `verify-quick-start.mjs` all carry the same `sheetPathFor(name)` helper.

## Chart Injection

ExcelJS has no `addChart` API. Charts are injected via `src/lib/excel/chartInjection.ts`, which post-processes the `writeBuffer()` output using JSZip: injects `xl/charts/chart1.xml`, `xl/drawings/drawing1.xml`, rels, and content-type overrides.

- `injectChart(buffer, opts)` accepts `kind: 'doughnut' | 'pie' | 'bar'`
- Budget chart bound to `OVERVIEW!$G$15:$G$20` (cats) / `$N$15:$N$20` (hidden helper col, always populated); the bar variant stacks `$O$/$P$/$Q$15:20`
- `plotVisOnly=0` so hidden column still plots
- Theme colors: `CHART_PALETTES` (Record\<ThemeId, string[]>, 6 colors/theme)
- Budget chart anchor: G25:L36, sitting directly under an in-cell "SPENT VS PLANNED" header at `G24:L24`. **Neither OVERVIEW chart passes a `title`** — a chart-space title eats plot height and doesn't align with anything else on the sheet, so both headers are cells instead (which also lets them sit level with the left column's headers).

**Two charts on OVERVIEW**: `index.ts` calls `injectChart` twice — the budget bar chart, then a trip-readiness doughnut (Ready/To-do, bound to hidden cells `T1:T2`/`U1:U2`, anchor B25:E34, gated on `state.sheets.packingList || state.sheets.tasks`). The second call lands on the MERGE branch (a drawing part already exists after the first call) and appends a second `<xdr:twoCellAnchor>` into the same drawing part. `buildChartAnchorXml`/`buildDrawingXml` take a `frameId` param (`chartIdx + 1`) so each chart's `cNvPr id` is unique — two charts sharing a drawing part with the same id produces an Excel "repair" prompt on open.

**Centered ring label**: the readiness doughnut passes `legend: false, dataLabels: false` and no `title` (an in-cell "TRIP READINESS" header + `overview.ts`'s `B29`/`B31` cells carry that instead), so its plot area — and the hole — stay centered in the anchor. Every injected chart also gets `<c:roundedCorners val="0"/>` and a transparent (`noFill`) chart-space background (`chartInjection.ts`), which is what lets the `B29` cell's `=TEXT(T1/T3,"0%")` show through the doughnut's hole instead of sitting behind an opaque white box. This float-a-cell-behind-the-hole technique needs live-rendering verification in Excel and Google Sheets — alignment depends on Excel's automatic plot-area layout math and isn't guaranteed identical across viewers.

Verify charts: `node scripts/verify-chart-injection.mjs` (single-chart XML shape) and `node scripts/verify-two-chart-injection.mjs` (two charts merged into one drawing part, frame-id collision regression)

## OVERVIEW Layout (two columns sharing row numbers)

OVERVIEW is a one-screen dashboard, not a scrollable table: `ws.views = [{ showGridLines:
false }]` — no gridlines and no freeze pane (every other sheet builder freezes its header
rows; this is a deliberate OVERVIEW-only exception).

It stacks **two independent content columns that share row numbers**: a B:E card stack
and a G:L card stack. Column **A is an empty 3-wide left margin** (so the dark panel
floats off the sheet edge) and column **F an empty 6-wide gutter** between the two stacks;
neither is ever filled or populated. Section headers are deliberately kept **level across
that gutter** — row 13 (BUDGET SUMMARY), row 24 (TRIP READINESS ‖ SPENT VS PLANNED),
row 38 (JUMP TO ‖ WHERE TO BASE YOURSELF) — which is what gives the sheet a horizontal
rhythm rather than two independently drifting columns.

| Rows  | B:E (left)                                     | G:L (right)                           |
| ----- | ---------------------------------------------- | ------------------------------------- |
| 1–14  | dark hero panel (see below)                    | 1–11 cover photo, 12 gutter           |
| 13–14 | countdown caption / panel padding              | BUDGET SUMMARY header, column headers |
| 15–16 | CURRENCY / TOTAL BUDGET card                   | six budget category rows (15–20)      |
| 17–22 | currency-exchange widget (when built)          | TOTAL row (21), gutter (22)           |
| 23    | blank gutter under the exchange widget         | blank gutter                          |
| 24    | TRIP READINESS header                          | SPENT VS PLANNED header               |
| 25–36 | readiness doughnut (25–34), count caption (35) | spend bar chart (25–36)               |
| 38    | JUMP TO header                                 | WHERE TO BASE YOURSELF header         |
| 39–48 | quick-nav hyperlink rows (39–46)               | five neighborhoods, two rows each     |

**All row heights are set in one `rowHeights` map** at the top of `buildOverviewSheet`,
not by each block. Because the two columns share row numbers, every height is a
*resolution* of both sides' needs; letting a block set its own is how one side ends up
squashing the other. The column-scoped style helpers deliberately do **not** touch
heights (their `styleFactory` originals do).

Hidden helper columns: **N–Q** (chart series), **S** (currency dropdown options),
**T–U** (readiness ring values/labels).

### Hero band (rows 1–14)

Excel can't composite editable cell text on top of an opaque floating image (drawings
always render above the grid — no CSS-style layering), so this is a dark text panel
**beside** the cover photo, not overlaid on it:

- **B1:E14** — solid `ts.palette.primary` fill, no borders. Destination headline (`B2:E3`,
  steps down a font size past 18 chars); then two **label-over-value cards** — `B5:C5`
  START DATE / `D5:E5` END DATE over `B6:C6` TripStart / `D6:E6` TripEnd, and `B8`/`C8:D8`/
  `E8` DAYS / PARTY / TRAVELERS over `B9` derived day count / `C9:D9` party-type dropdown /
  `E9` travelers; then the countdown badge (`B11:E12` oversized count + `B13:E13` unit
  caption, accent-filled across all three rows — the mockup's big circular "292 / DAYS TO
  GO" marker, which can't be drawn as a circle over the photo). The count and unit are two
  cells rather than one sentence so the number can carry its own display size.
- **The captions are what keep the values honest.** An earlier pass had no labels at all
  and folded the punctuation into each value's number format instead (`"–  "mmm d, yyyy`
  on TripEnd, `0" days"`, `0" travelers"`). Now that every value is captioned, the number
  formats are plain (`mmm d, yyyy`, `0`) and the cells stay genuinely editable.
- **No separator columns**, and no per-value single-column layout: values are merged
  across half the panel (B:C / D:E) so a long value can't clip its neighbor. A narrow
  separator column would be shared by every other B:E block down the sheet (the row 15–16
  card, currency widget, quick-nav, readiness ring).
- Caption styling inside the panel is **not** `styleLabelCell` — that helper's `lightBg`
  fill would punch holes in the dark band. It's the same small/bold treatment in
  `ts.palette.primaryText`. The rows 15–16 card sits *below* the panel and does use the
  shared `styleLabelCell`/`styleValueCell` pair.
- **A–F widths are set once, with the hero band** — sized for its title/header fonts, well
  above the 11pt default the width unit assumes. The later "Column widths" block sets G–L
  only; re-adding A–F there silently discards the hero sizing (this regressed once and
  produced `###` and clipped labels everywhere).
- **G1:L11** — cover photo (`state.overviewImage`), embedded via the `addDataUrlImage`
  helper. This is the workbook's **only** image. It stops three rows short of the panel's
  bottom so the BUDGET SUMMARY header at row 13 gets a gutter instead of butting into the
  photo. `PictureUploader.tsx`'s crop target (`TARGET_W`/`TARGET_H`) must match this
  anchor's aspect ratio — G:L is 99 width units (~693px), rows 1–11 total 260pt (~347px),
  so ~2:1, currently 1400×700 — or the embedded photo stretches. **Any row-height change
  in rows 1–11 changes that ratio**, so the two must move together.

All hero-band styling is bespoke/inline in `overview.ts`, not the shared `styleFactory`
helpers — scoped to this sheet so every other sheet keeps its existing table look.

**Never use `styleSectionHeader`/`styleColumnHeader`/`styleDataRow`/`styleTotalRow` on
OVERVIEW.** Those walk *every* populated cell in a row (and set the row height), and the
two columns share row numbers — so a header on one side restyles the other side's data
(this shipped once: the TRIP READINESS header bolded a budget row, and the neighborhood
guide restyled the quick-nav links). The Budget Summary table overlaps the left column's
CURRENCY/TOTAL BUDGET card *and* the currency-exchange widget, so this is not theoretical.
Use the column-scoped `styleSectionHeaderCells` / `styleColumnHeaderCells` /
`styleDataRowCells` / `styleTotalRowCells` helpers in `overview.ts` with `LEFT_COLS` /
`RIGHT_COLS` instead.

**The readiness ring's centered % must be centered on the chart FRAME, not eyeballed.**
The frame is anchored `B25:E34` whose `toCol: 5` is an *exclusive* right edge, so it spans
columns B–E only; merging `B29:E30` therefore spans exactly the frame width and lands the
text on the donut hole. (A merge starting one column further right sits ~6 width units off
center, which pushes the text under the ring — that shipped once.) Vertically, rows 25–34
are all `RING_ROW_H` (18pt) = 180pt, so the frame's mid-line at 90pt falls exactly on the
29/30 boundary. Keep that arithmetic in sync with the anchor, `RING_ROW_H`, and the B–E
widths.

**The quick-nav strip and the neighborhood guide share rows 39–46.** Both use
`BOTTOM_ROW_H`, and neighborhood descriptions are truncated to a single line at the G:L
merged width (`NEIGHBORHOOD_DESC_CHARS`) rather than given a taller wrapped row — a
two-height right column would show through as a ragged left one.

## Recalculation Contract (Excel for Mac chart repaint)

ExcelJS never writes `xl/calcChain.xml`. Without it, Excel for Mac mis-orders its chart-refresh hooks and the injected chart repaints **one calculation behind the cells** (each typed cost shows the *previous* entry's bars). The fix is a three-part contract — **keep all three in sync**:

1. **`injectCalcChain()`** (`src/lib/excel/calcChain.ts`, runs after `injectChart` in `index.ts`) writes the calcChain part and pins `<calcPr calcId="191029"/>` (ExcelJS hardcodes 171027 with no API).
2. **Never set `wb.calcProperties.fullCalcOnLoad`** — the full rebuild it forces on load re-breaks the repaint ordering.
3. **Every formula cell must carry a correct cached `result:`** — with a native calcId, Excel trusts caches on open and won't recalc, so an uncached formula renders blank.

Also: OVERVIEW "Actual Spent" formulas must sum tracker **data rows directly** (e.g. `SUM('ACCOMMODATION'!$E$6:$E$25)`), never tracker TOTAL cells; and the injected chart series carry `<c:strCache>`/`<c:numCache>` data points matching the cells' cached results (`categoryBudgetAmounts()` in `overview.ts` feeds both). Diagnosed by diffing our output against the same workbook re-saved natively by Excel; verified by typing tests in Excel for Mac 16.110.

## Themes, Fonts, Chart Styles

- **4 color themes:** Sakura, Ocean, Forest, Desert (defined in `src/types/theme.ts`; `ThemePicker` auto-lists from `Object.keys(THEMES)`)
- **3 font families:** sans (Calibri), serif (Cambria), mono (Courier New)
- **3 chart styles:** bar, pie, donut (all native DrawingML)

## Named Ranges

Defined in `src/lib/excel/workbookConfig.ts` via `addNamedRanges`. **ExcelJS arg order is `wb.definedNames.add(location, name)` — location first, name second.** Reversing these silently produces an empty `<definedNames>` block (formulas silently resolve to blank via `IFERROR` wrappers).

Each location is the **anchor cell of a merged card value** on OVERVIEW: `TripStart` B6, `TripEnd` D6, `NumAdults` E9, `TotalBudget` D16. Moving one of those cards means updating `workbookConfig.ts`, the INSTRUCTIONS named-range table (`instructions.ts`), and `scripts/verify-sheet-protection.mjs` together.

Freshly written formula cells need a cached `result:` value or they render blank until the first recalc — set `cell.value = { formula: '...', result: computedValue }`.

## Excel + Google Sheets Compatibility Constraints

The generated `.xlsx` must open correctly in **both Excel and Google Sheets**. This eliminates entire categories of techniques:

- **No macros/VBA** — `.xlsx` can't carry them; any dynamic behavior must be static formulas or data validation
- **No row hiding driven by cell/dropdown** — needs VBA (Excel) or Apps Script (Sheets); outline groups are OK (manual [+]/[−])
- **No currency-symbol-from-dropdown** — `numFmt` strings are static; can't reference a cell value
- **Cell/sheet protection is Excel-only** — every generated sheet except INSTRUCTIONS and ANNUAL EVENTS is locked via `protectWorkbook()` (`src/lib/excel/protection.ts`, called in `index.ts` right before `writeBuffer()`). Each sheet builder marks its own data-entry cells `protection: { locked: false }` and its formula/total cells `protection: { hidden: true }` at the point those cells are written. No password — this is a guardrail against accidentally overwriting a formula, not real security. Google Sheets ignores `sheetProtection`/`cell.protection` on upload, so every cell becomes editable again there; don't imply otherwise in copy or INSTRUCTIONS text
- **A cell's `locked` flag is inert unless the sheet carries `<sheetProtection>`** — "unprotected sheet with one protected column" is not expressible in OOXML. To make a sheet feel open while guarding a single column, keep `sheetProtection` on, unlock every other cell, and switch off the sheet-level restrictions via `RELAXED_PROTECTION` in `protection.ts` (ExcelJS's defaults disallow formatting, insert/delete, sort and filter). **ITINERARY** is the one sheet set up this way: it's a writing surface, so every data cell in rows 4–33 is unlocked except the auto-Date column **B**, which stays locked + formula-hidden so typing a date can't sever that row's `TripStart` link. Column A's day number is deliberately editable there. Rows 1–3 stay locked as chrome
- **Internal navigation links work in BOTH — but only in the native form.** A `HYPERLINK("#'SHEET'!A1", …)` formula navigates in Excel and is inert text in Google Sheets. A native `{ text, hyperlink: "'SHEET'!A1" }` link works in both, once `nativeHyperlinks.ts` has rewritten it into Excel's `location` + `display` shape (verified by Drive import). Prefer the native form for any new link. Note Sheets always needs a second click via its link-preview chip — a platform behavior the file can't control; see "Native Internal Hyperlinks" above
- **The OVERVIEW B16 currency dropdown** is cosmetic/reference-only (a `type:'list'` data validation of 30 currencies); it does NOT drive `numFmt`. Options written to hidden helper column R (inline list would exceed Excel's 255-char limit). This is acceptable because it doesn't imply false interactivity.

Before proposing any in-spreadsheet interactivity, verify it works statically in Google Sheets. If a control can't actually do what it implies, say so and offer honest alternatives.

## Key File Map

```
src/
  store/wizardStore.ts          # Zustand state + DEFAULT_STATE
  types/wizard.ts               # WizardState type, ALL_SHEET_IDS
  types/theme.ts                # ThemeId, THEMES, CHART_PALETTES
  hooks/useWizardNavigation.ts  # Step validation + navigation
  hooks/useExcelGeneration.ts   # Orchestrates generation + download
  lib/
    excel/
      index.ts                  # Main buildWorkbook() entry point
      workbookConfig.ts         # Named ranges, workbook defaults
      styleFactory.ts           # ThemeStyle, helpers: truncate(), wrappedLineCount(), rowHeightForLines()
      chartInjection.ts         # JSZip-based DrawingML chart injection
      nativeHyperlinks.ts       # Rewrites internal links into Excel's location+display shape
      sheets/
        overview.ts             # OVERVIEW sheet + buildActualSpentFormula()
        itinerary.ts
        flights.ts              # TRANSPORT sheet (SheetId stays `flights`)
        hotels.ts
        restaurants.ts
        excursions.ts
        budgetTracker.ts
        packingList.ts
        tasks.ts
        events.ts               # ANNUAL EVENTS (AI-gated)
        instructions.ts
        quickStart.ts           # QUICK START (always the first tab)
        recommendationInsert.ts # AI guide table injected below trackers
    recommendations/
      client.ts                 # Browser-side fetch to /api/recommendations
      types.ts                  # PlaceRec, RegionRec, RecommendationResponse
  data/
    currencies.ts               # 30 currencies with code + symbol
    packingLists.ts
    tripStylePresets.ts
  components/
    layout/                     # WizardShell, StepProgress
    shared/                     # ThemePicker, FontPicker, ChartStylePicker, SheetToggleGrid, CurrencySelector, PictureUploader
    steps/                      # Step1–Step4

vite-plugin-recommendations.ts  # Vite middleware: /api/recommendations POST
scripts/
  verify-chart-injection.mjs
  verify-currency-dropdown.mjs
  verify-image-chart-injection.mjs
  verify-two-chart-injection.mjs
  verify-sheet-protection.mjs
  verify-quick-nav-icons.mjs
  verify-rec-jump-links.mjs
  verify-check-dropdowns.mjs
  verify-quick-start.mjs
```

## Row Height / Text Wrapping (ExcelJS)

Long text in merged cells does NOT auto-grow in Excel desktop (Google Sheets and Numbers do auto-grow, so bugs only surface when opening in Excel). The fix is dynamic row height + wrapping + length caps — NOT column-width autofit (ExcelJS has no real autofit).

Helper functions in `styleFactory.ts`:

- `truncate(text, max)` — word-boundary truncation with ellipsis
- `wrappedLineCount(text, widthChars)` — `ceil(len / floor(width * 0.95))`, honors `\n`
- `rowHeightForLines(lines, ts, { min, maxLines })` — `lineHeight = ts.sizes.data * 1.4 + 6`

Apply `wrap: true` and `vertical: 'top'` per-cell (not via `row.alignment` — unreliable in ExcelJS).

## Environment

Copy `.env.example` → `.env` and add `ANTHROPIC_API_KEY` to test AI recommendations locally. The key is read server-side only (Vite middleware); it is never bundled into the client.

## Verification Scripts

```bash
node scripts/verify-chart-injection.mjs      # bar, pie, donut → /tmp/chart-*.xlsx
node scripts/verify-currency-dropdown.mjs    # currency dropdown round-trip
node scripts/verify-image-chart-injection.mjs
node scripts/verify-two-chart-injection.mjs  # budget chart + readiness doughnut share one drawing part
node scripts/verify-sheet-protection.mjs
node scripts/verify-quick-nav-icons.mjs      # OVERVIEW JUMP TO strip: per-sheet emoji survive into the hyperlink `display` attr; no r:id; cover-photo drawing rel intact
node scripts/verify-check-dropdowns.mjs      # PACKING LIST C / TASKS D one-option "✓" dropdown matches the COUNTIF literal, in-model and in raw XML
node scripts/verify-quick-start.mjs          # QUICK START is tab 1, its jump links are native location+display, its step numbering has no gaps, and it only advertises recommendation guides that exist
node scripts/verify-rec-jump-links.mjs       # tracker row-2 "★ Recommendations →" button: native location+display XML, no r:id, no leftover hyperlink rels, chip styling contrasts the band, and band fill preserved when there's no guide
```

## Git Workflow

Never force-push, never skip hooks, never amend published commits.
