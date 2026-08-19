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
| 1   | OVERVIEW       | ✓         | Named ranges: TripStart A5, TripEnd B5, NumAdults C7, TotalBudget D15; rows 1–14 are a dashboard "hero band" (dark A:D text panel + F1:K14 cover photo, no gridlines/freeze — see below) instead of a title bar + label/value ledger; native DrawingML chart(s) injected via JSZip post-processing; trip-readiness doughnut ring (A24:E40) with a centered `S1/S3` "% ready" label floated over its transparent hole, quick-nav HYPERLINK strip (A42+, each link prefixed with its `SHEET_ICONS` emoji — the same map the Step 3 toggle grid uses, exported from `src/types/wizard.ts`), neighborhood guide from AI regions (F46+, when recs enabled). Exactly one image (the cover photo) — a second upload + multi-image strip was tried and removed: it served no purpose and arbitrary source dimensions made the results unpredictable |
| 2   | ITINERARY      | ✓         | Auto-dates from TripStart+n; AI-prefilled when recs on                                                                                           |
| 3   | TRANSPORT      | toggle    | SheetId = `flights`; Mode dropdown: Air/Train/Bus/Car/Ferry/Taxi/Other                                                                           |
| 4   | HOTELS         | toggle    | Date col auto-fills from TripStart; AI recommendation guide table appended below tracker                                                         |
| 5   | RESTAURANTS    | toggle    | AI recommendation guide table appended below tracker                                                                                             |
| 6   | EXCURSIONS     | toggle    | Per-person × NumAdults; AI recommendation guide table appended below tracker                                                                     |
| 7   | BUDGET TRACKER | toggle    | Date / Category dropdown / Description / Amount / Running Total / Notes                                                                          |
| 8   | PACKING LIST   | toggle    | Seasonal packing + COUNTIF progress formula                                                                                                      |
| 9   | TASKS          | toggle    |                                                                                                                                                  |
| 10  | ANNUAL EVENTS  | AI-gated  | Only generated when `useRecommendations` is on AND data returned                                                                                 |
| 11  | INSTRUCTIONS   | ✓         | Always last; not toggleable                                                                                                                      |

## AI Recommendations

Backend: `vite-plugin-recommendations.ts` (root) registers a `/api/recommendations` POST endpoint in Vite's dev middleware. Uses `@anthropic-ai/sdk` with model `claude-sonnet-4-6` and the `web_search_20250305` tool.

Requires `ANTHROPIC_API_KEY` in `.env` (see `.env.example`).

**What the API returns:** `{ regions, itinerary, events }`

- `regions`: `{ region, description, hotels[], restaurants[], excursions[] }[]` — each place is a `PlaceRec { name, price?, rating?, cuisine?, duration?, description? }`
- AI output is **informational only** — it does NOT prefill HOTELS/TRANSPORT/RESTAURANTS/EXCURSIONS tracker rows
- ITINERARY IS still AI-prefilled
- ANNUAL EVENTS tab is still built from AI data

**Recommendation insert** (`src/lib/excel/sheets/recommendationInsert.ts`): `buildRecommendationInsert(ws, startRow, ts, regions, kind)` renders a region-grouped guide table BELOW the working tracker on each of HOTELS/RESTAURANTS/EXCURSIONS. The `kind` picks the array and attribute columns (hotels=Price+Stars, restaurants=Cuisine+Price, excursions=Duration+Price). The places loop **must** `row++` per place or ExcelJS throws "Cannot merge already merged cells."

## Chart Injection

ExcelJS has no `addChart` API. Charts are injected via `src/lib/excel/chartInjection.ts`, which post-processes the `writeBuffer()` output using JSZip: injects `xl/charts/chart1.xml`, `xl/drawings/drawing1.xml`, rels, and content-type overrides.

- `injectChart(buffer, opts)` accepts `kind: 'doughnut' | 'pie' | 'bar'`
- Budget chart bound to `OVERVIEW!$F$18:$F$23` (cats) / `$M$18:$M$23` (hidden helper col, always populated)
- `plotVisOnly=0` so hidden column still plots
- Theme colors: `CHART_PALETTES` (Record\<ThemeId, string[]>, 6 colors/theme)
- Budget chart anchor: F27:L45

**Two charts on OVERVIEW**: `index.ts` calls `injectChart` twice — the budget bar chart, then a trip-readiness doughnut (Ready/To-do, bound to hidden cells `S1:S2`/`T1:T2`, anchor A24:E40, gated on `state.sheets.packingList || state.sheets.tasks`). The second call lands on the MERGE branch (a drawing part already exists after the first call) and appends a second `<xdr:twoCellAnchor>` into the same drawing part. `buildChartAnchorXml`/`buildDrawingXml` take a `frameId` param (`chartIdx + 1`) so each chart's `cNvPr id` is unique — two charts sharing a drawing part with the same id produces an Excel "repair" prompt on open.

**Centered ring label**: the readiness doughnut passes `legend: false, dataLabels: false` and no `title` (an in-cell "TRIP READINESS" header + `overview.ts`'s `B31`/`B33` cells carry that instead), so its plot area — and the hole — stay centered in the anchor. Every injected chart also gets `<c:roundedCorners val="0"/>` and a transparent (`noFill`) chart-space background (`chartInjection.ts`), which is what lets the `B31` cell's `=TEXT(S1/S3,"0%")` show through the doughnut's hole instead of sitting behind an opaque white box. This float-a-cell-behind-the-hole technique needs live-rendering verification in Excel and Google Sheets — alignment depends on Excel's automatic plot-area layout math and isn't guaranteed identical across viewers.

Verify charts: `node scripts/verify-chart-injection.mjs` (single-chart XML shape) and `node scripts/verify-two-chart-injection.mjs` (two charts merged into one drawing part, frame-id collision regression)

## OVERVIEW Hero Band (rows 1–14)

Replaces the old title bar + 12-row label/value ledger. Excel can't composite editable
cell text on top of an opaque floating image (drawings always render above the grid —
no CSS-style layering), so this is a dark text panel **beside** the cover photo, not
overlaid on it:

- **A1:D14** — solid `ts.palette.primary` fill, no borders. Destination headline (`A2:D3`,
  steps down a font size past 18 chars), date line (`A5` TripStart / `B5` TripEnd), a
  duration/party/travelers line (`A7` derived duration / `B7` party-type dropdown / `C7`
  travelers), and the countdown badge (`A9:B10` oversized count + `A11:B11` unit caption,
  accent-filled across all three rows — the mockup's big circular "292 / DAYS TO GO"
  marker, which can't be drawn as a circle over the photo). The count and unit are two
  cells rather than one sentence so the number can carry its own display size; rows 9–10
  are heightened to 26pt to fit it.
- **No separator columns.** Punctuation that belongs to a value is folded into that
  value's number format (`"–  "mmm d, yyyy` on TripEnd, `0" days"`, `0" travelers"`), so
  the cell stays a genuinely editable raw date/number while displaying the decoration.
  A narrow separator column would be shared by every other A:D block down the sheet (the
  row-15 strip, currency widget, quick-nav, readiness ring) and is what forced the
  row-15 labels to clip.
- **Column E** is a deliberately empty, unfilled 3-wide spacer between the panel and the
  photo. **A–E widths are set once, with the hero band** — sized for its title/header
  fonts, well above the 11pt default the width unit assumes. The later "Column widths"
  block sets F–K only; re-adding A–E there silently discards the hero sizing (this
  regressed once and produced `###` and clipped labels everywhere).
- **F1:K14** — cover photo (`state.overviewImage`), embedded via the `addDataUrlImage`
  helper. This is the workbook's **only** image. `PictureUploader.tsx`'s crop target
  (`TARGET_W`/`TARGET_H`) must match this anchor's aspect ratio (currently 1400×756,
  ~1.85:1) or the embedded photo stretches.
- **Row 15** reuses the old ledger's trailing spacer row for a compact Currency/Total
  Budget strip (`A15`/`B15`/`C15`/`D15`) — sized to fit exactly where the old block ended
  so row 16 onward (Budget Summary and everything below) never needs to move.
- `ws.views = [{ showGridLines: false }]` — no gridlines and no freeze pane (every other
  sheet builder freezes header rows; this is a deliberate OVERVIEW-only exception since
  it's a one-screen dashboard, not a scrollable table).

All hero-band styling is bespoke/inline in `overview.ts`, not the shared `styleFactory`
helpers — scoped to this sheet so every other sheet keeps its existing table look.

**Never use `styleSectionHeader`/`styleDataRow` on OVERVIEW.** Those walk *every*
populated cell in a row, and OVERVIEW stacks two independent content columns (the A:D
panel and the F:K cards) that share row numbers — so a header on one side restyles the
other side's data (this shipped once: the TRIP READINESS header at row 23 bolded the
Miscellaneous budget row, and the neighborhood guide at rows 46+ restyled the quick-nav
links). Use the column-scoped `styleSectionHeaderCells` / `styleDataRowCells` helpers in
`overview.ts` with `LEFT_COLS` / `RIGHT_COLS` instead.

**The readiness ring's centered % must be centered on the chart FRAME, not eyeballed.**
The frame is anchored `A24:E40` whose `toCol: 4` is an *exclusive* right edge, so it spans
columns A–D only; merging `A31:D32` therefore spans exactly the frame width and lands the
text on the donut hole. (`B31:D32` sits ~6 width units right of center, which pushes the
text under the ring — that shipped once too.) Vertically, rows 24–39 total ~238pt (row 24
= 20pt total row, row 25 = 8pt spacer, rest default 15pt), so the mid-line falls on the
31/32 boundary. Keep that arithmetic in sync with the anchor and the A–D widths.

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

Freshly written formula cells need a cached `result:` value or they render blank until the first recalc — set `cell.value = { formula: '...', result: computedValue }`.

## Excel + Google Sheets Compatibility Constraints

The generated `.xlsx` must open correctly in **both Excel and Google Sheets**. This eliminates entire categories of techniques:

- **No macros/VBA** — `.xlsx` can't carry them; any dynamic behavior must be static formulas or data validation
- **No row hiding driven by cell/dropdown** — needs VBA (Excel) or Apps Script (Sheets); outline groups are OK (manual [+]/[−])
- **No currency-symbol-from-dropdown** — `numFmt` strings are static; can't reference a cell value
- **Cell/sheet protection is Excel-only** — every generated sheet except INSTRUCTIONS and ANNUAL EVENTS is locked via `protectWorkbook()` (`src/lib/excel/protection.ts`, called in `index.ts` right before `writeBuffer()`). Each sheet builder marks its own data-entry cells `protection: { locked: false }` and its formula/total cells `protection: { hidden: true }` at the point those cells are written. No password — this is a guardrail against accidentally overwriting a formula, not real security. Google Sheets ignores `sheetProtection`/`cell.protection` on upload, so every cell becomes editable again there; don't imply otherwise in copy or INSTRUCTIONS text
- **The OVERVIEW D11 currency dropdown** is cosmetic/reference-only (a `type:'list'` data validation of 30 currencies); it does NOT drive `numFmt`. Options written to hidden helper column R (inline list would exceed Excel's 255-char limit). This is acceptable because it doesn't imply false interactivity.

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
node scripts/verify-quick-nav-icons.mjs      # OVERVIEW JUMP TO strip: per-sheet emoji survive the HYPERLINK formula round-trip
```

## Git Workflow

Never force-push, never skip hooks, never amend published commits.
