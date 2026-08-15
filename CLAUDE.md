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
| 1   | OVERVIEW       | ✓         | Named ranges: TripStart D3, TripEnd D4, NumAdults D9, TotalBudget D14; native DrawingML chart (bar/pie/donut) injected via JSZip post-processing |
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
- All charts bound to `OVERVIEW!$F$18:$F$23` (cats) / `$M$18:$M$23` (hidden helper col, always populated)
- `plotVisOnly=0` so hidden column still plots
- Theme colors: `CHART_PALETTES` (Record\<ThemeId, string[]>, 6 colors/theme)
- Anchor: F27:L45

Verify charts: `node scripts/verify-chart-injection.mjs`

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
```

## Git Workflow

Never force-push, never skip hooks, never amend published commits.
