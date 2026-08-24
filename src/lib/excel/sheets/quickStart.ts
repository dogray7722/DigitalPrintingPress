import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { Recommendations } from '../../recommendations/types'
import type { ThemeStyles } from '../styleFactory'
import { wrappedLineCount, rowHeightForLines } from '../styleFactory'
import { hasRecommendationsFor } from './recommendationInsert'

// QUICK START is the workbook's first tab: the five things a buyer has to do before the
// rest of the file is theirs. It is deliberately short and points at specific cells —
// INSTRUCTIONS (last tab) is the long-form reference, and this sheet must not grow into
// a second copy of it.
//
// Like INSTRUCTIONS it is left unprotected (see protection.ts): it holds no formulas, so
// there is nothing here to guard, and a locked sheet reads as "don't touch" on the very
// first thing the buyer opens.

// Body/bullet text is merged across A:F, so the merged width below drives wrapping.
const COL_WIDTHS: Record<string, number> = {
  A: 6, // step-number badge column — narrow so the numeral sits in its own gutter
  B: 24,
  C: 24,
  D: 22,
  E: 22,
  F: 20,
}
const MERGE_WIDTH = Object.values(COL_WIDTHS).reduce((a, b) => a + b, 0)
// Step bodies/bullets are merged B:F (indented past the badge column).
const INDENT_WIDTH = MERGE_WIDTH - COL_WIDTHS.A

export function buildQuickStartSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  recommendations?: Recommendations
): void {
  const ws = wb.addWorksheet('QUICK START')
  ws.properties.tabColor = { argb: ts.palette.accent }
  ws.views = [{ showGridLines: false }]

  const regions = recommendations?.regions
  // Which trackers actually ended up with a "★ Recommendations →" chip. Gate on the
  // same predicate the insert builder uses — a tab can be enabled and still have no
  // guide when the AI returned no places of that kind.
  const recTabs: [enabled: boolean, sheet: string, blurb: string][] = [
    [
      state.sheets.hotels && hasRecommendationsFor(regions, 'hotels'),
      'ACCOMMODATION',
      'places to stay grouped by neighbourhood, with a price band and star rating — use it to choose the area you want to be based in, then log whatever you actually book in the tracker at the top of the sheet.',
    ],
    [
      state.sheets.restaurants && hasRecommendationsFor(regions, 'restaurants'),
      'DINING',
      'restaurants and cafés by neighbourhood, with cuisine and price band — shortlist two or three near where you are staying so you are not deciding on an empty stomach.',
    ],
    [
      state.sheets.excursions && hasRecommendationsFor(regions, 'excursions'),
      'EXCURSIONS',
      'tours, sights and activities with a rough duration and price — the duration column is what makes these easy to slot into a free morning or afternoon on ITINERARY.',
    ],
  ]
  const liveRecTabs = recTabs.filter(([on]) => on)

  let row = 1
  let step = 0

  // ── Local styling helpers ───────────────────────────────────────────────────

  function solid(argb: string): ExcelJS.Fill {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb } } as ExcelJS.Fill
  }

  /** Full-width banner across A:F. */
  function title(text: string, subtitle: string) {
    const cell = ws.getCell(`A${row}`)
    cell.value = text
    ws.mergeCells(`A${row}:F${row}`)
    cell.fill = solid(ts.palette.primary)
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.title,
      bold: true,
      color: { argb: ts.palette.primaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'center' }
    ws.getRow(row).height = 42
    row++

    const sub = ws.getCell(`A${row}`)
    sub.value = subtitle
    ws.mergeCells(`A${row}:F${row}`)
    sub.fill = solid(ts.palette.primary)
    sub.font = {
      name: ts.fontName,
      size: ts.sizes.data,
      color: { argb: ts.palette.primaryText },
    }
    sub.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    ws.getRow(row).height = rowHeightForLines(wrappedLineCount(subtitle, MERGE_WIDTH), ts, {
      min: 22,
    })
    row++
  }

  /**
   * Numbered step header: the numeral sits alone in the narrow column A badge, the
   * heading runs across B:F. Auto-increments, so reordering steps can't leave a gap.
   */
  function stepHeader(heading: string) {
    step++
    const badge = ws.getCell(`A${row}`)
    badge.value = step
    badge.fill = solid(ts.palette.primary)
    badge.font = {
      name: ts.fontName,
      size: ts.sizes.sectionHeader,
      bold: true,
      color: { argb: ts.palette.primaryText },
    }
    badge.alignment = { vertical: 'middle', horizontal: 'center' }

    const cell = ws.getCell(`B${row}`)
    cell.value = heading
    ws.mergeCells(`B${row}:F${row}`)
    cell.fill = solid(ts.palette.secondary)
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.sectionHeader,
      bold: true,
      color: { argb: ts.palette.secondaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
    ws.getRow(row).height = 24
    row++
  }

  /** Section header with no step number (used for the closing "where next" block). */
  function sectionHeader(heading: string) {
    const cell = ws.getCell(`A${row}`)
    cell.value = heading
    ws.mergeCells(`A${row}:F${row}`)
    cell.fill = solid(ts.palette.secondary)
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.sectionHeader,
      bold: true,
      color: { argb: ts.palette.secondaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
    ws.getRow(row).height = 24
    row++
  }

  /** Lead paragraph under a step header, indented to B:F. */
  function bodyLine(text: string) {
    const cell = ws.getCell(`B${row}`)
    cell.value = text
    ws.mergeCells(`B${row}:F${row}`)
    cell.font = { name: ts.fontName, size: ts.sizes.data, bold: true }
    cell.alignment = { vertical: 'middle', wrapText: true }
    ws.getRow(row).height = rowHeightForLines(wrappedLineCount(text, INDENT_WIDTH), ts, {
      min: 18,
    })
    row++
  }

  function bulletLine(text: string) {
    const full = `•   ${text}`
    const cell = ws.getCell(`B${row}`)
    cell.value = full
    ws.mergeCells(`B${row}:F${row}`)
    cell.font = { name: ts.fontName, size: ts.sizes.data }
    cell.fill = solid(ts.palette.lightBg)
    cell.alignment = { vertical: 'middle', wrapText: true }
    ws.getRow(row).height = rowHeightForLines(wrappedLineCount(full, INDENT_WIDTH), ts, {
      min: 18,
    })
    row++
  }

  /**
   * A native internal hyperlink row — `{ text, hyperlink }`, NOT a HYPERLINK() formula.
   * The formula form navigates in Excel but is inert in Google Sheets; nativeHyperlinks.ts
   * rewrites this form into Excel's `location` + `display` shape, which both read. (Sheets
   * still needs a second click via its link-preview chip — a platform behavior no file can
   * change.) No formula also means no cached `result:` obligation.
   */
  function linkRow(sheet: string, label: string, blurb: string) {
    const cell = ws.getCell(`B${row}`)
    cell.value = { text: label, hyperlink: `'${sheet}'!A1` }
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.data,
      bold: true,
      color: { argb: ts.palette.primary },
    }
    cell.alignment = { vertical: 'middle' }

    const desc = ws.getCell(`C${row}`)
    desc.value = blurb
    ws.mergeCells(`C${row}:F${row}`)
    desc.font = { name: ts.fontName, size: ts.sizes.data }
    desc.alignment = { vertical: 'middle', wrapText: true }

    const descWidth = COL_WIDTHS.C + COL_WIDTHS.D + COL_WIDTHS.E + COL_WIDTHS.F
    ws.getRow(row).height = rowHeightForLines(wrappedLineCount(blurb, descWidth), ts, {
      min: 20,
    })
    row++
  }

  function spacer() {
    ws.getRow(row).height = 10
    row++
  }

  // ── Title ───────────────────────────────────────────────────────────────────
  const dest = (state.destination || 'your trip').toUpperCase()
  // Deliberately no step count in the subtitle: the numbered steps are conditional
  // (the itinerary and recommendation steps drop out when those tabs aren't built),
  // so a hardcoded "five steps" would go stale.
  title(
    `QUICK START  —  ${dest}`,
    'Work down this page once and the rest of the workbook looks after itself. Every step points at a specific cell. The INSTRUCTIONS tab (last) has the long version.'
  )
  spacer()

  // ── Step: dates ─────────────────────────────────────────────────────────────
  stepHeader('SET YOUR TRIP DATES')
  bodyLine('Go to the OVERVIEW tab and change the two dates in the dark panel at the top.')
  bulletLine(
    'Start Date is cell B6, End Date is cell D6. Type over them — in Google Sheets you can double-click for a calendar; Excel has no cell date picker, so type the date instead.'
  )
  bulletLine(
    'Use a real date, not text. "15 Jun 2027" or your own short format both work — Excel reformats it for you.'
  )
  bulletLine(
    'These two cells drive the whole file: the day count, the countdown badge, and every date elsewhere in the workbook.'
  )
  bulletLine(
    'While you are up there, set the number of travellers in E9 — per-person costs on EXCURSIONS multiply by it.'
  )
  spacer()

  // ── Step: budget ────────────────────────────────────────────────────────────
  stepHeader('UPDATE YOUR ESTIMATED BUDGET')
  bodyLine(
    'The Budget Summary on the right of OVERVIEW is seeded with estimates from the wizard. Replace them with your own numbers.'
  )
  bulletLine(
    'Type over the Estimated Budget column, cells H15 to H20 — Transportation, Accommodation, Food, Activities, Shopping and Miscellaneous.'
  )
  bulletLine(
    'Do not type in the TOTAL row or in Total Budget (D16): those are formulas and will recalculate on their own.'
  )
  bulletLine(
    // No step number in the cross-reference: the numbering is conditional (see title()).
    'Actual Spent, Remaining and % Used stay empty until you start logging costs — that is the last step on this page.'
  )
  bulletLine(
    'The SPENT VS PLANNED chart underneath redraws itself as those numbers change; nothing to set up.'
  )
  spacer()

  // ── Step: itinerary ─────────────────────────────────────────────────────────
  if (state.sheets.itinerary) {
    const prefilled = !!recommendations?.itinerary?.length
    stepHeader('ADJUST THE ITINERARY')
    bodyLine(
      prefilled
        ? 'The ITINERARY tab arrives as a sample plan for your dates. It is a starting draft to edit, not a booked schedule — expect to rewrite most of it.'
        : 'The ITINERARY tab is laid out for your dates and waiting to be filled in, one row per day.'
    )
    bulletLine(
      prefilled
        ? 'Overwrite the Morning, Afternoon and Evening cells with what you actually want to do, and clear any suggested day you have no interest in.'
        : 'Fill in Morning, Afternoon and Evening for each day. Location, Accommodation, Transportation and Notes are there for the details around them.'
    )
    bulletLine(
      'Leave the Date column (B) alone — it fills itself from your Trip Start and is the only locked column on that tab. Everything else is yours to rewrite.'
    )
    bulletLine(
      'Need more days? Push Trip End (OVERVIEW D6) further out and the extra day rows appear on their own, up to 30.'
    )
    spacer()
  }

  // ── Step: recommendation links ──────────────────────────────────────────────
  if (liveRecTabs.length) {
    stepHeader('USE THE RECOMMENDATION LINKS')
    bodyLine(
      `Each of these tabs carries a research guide written for ${state.destination || 'your destination'}, sitting below the tracker you type into. A "★  Recommendations  →" button in the coloured band at the top of the sheet (row 2) jumps you straight down to it.`
    )
    bulletLine(
      'The button stays put when you scroll, so you can jump to the guide from anywhere on the sheet.'
    )
    bulletLine(
      'In Google Sheets one click opens a small preview card and a second click follows the link. That is how Sheets treats every link in every file — in Excel one click is enough.'
    )
    liveRecTabs.forEach(([, sheet, blurb]) => {
      bulletLine(`${sheet} — ${blurb}`)
    })
    bulletLine(
      'Work top-down: pick your neighbourhood on ACCOMMODATION first, then choose places to eat and things to do near it. OVERVIEW’s "WHERE TO BASE YOURSELF" panel is the short version of that same decision.'
    )
    bulletLine(
      'The guides are suggestions, not bookings. Nothing in them feeds your budget totals — only the rows you type into the tracker above do that. Ctrl+Home (Fn+Ctrl+Left on a Mac) takes you back to the top of the sheet.'
    )
    spacer()
  }

  // ── Step: tracking ──────────────────────────────────────────────────────────
  stepHeader('LOG THINGS AS YOU BOOK AND SPEND')
  bodyLine(
    'From here the workbook keeps itself up to date. Enter each cost once, on the tab it belongs to.'
  )
  const trackers = [
    [state.sheets.flights, 'TRANSPORTATION', 'flights, trains and transfers'],
    [state.sheets.hotels, 'ACCOMMODATION', 'nightly rates, taxes and fees'],
    [state.sheets.restaurants, 'DINING', 'meals out'],
    [state.sheets.excursions, 'EXCURSIONS', 'tours and activities'],
  ].filter(([on]) => on) as [boolean, string, string][]
  if (trackers.length) {
    bulletLine(
      `Costs go on the tracker tabs — ${trackers
        .map(([, sheet, what]) => `${sheet} (${what})`)
        .join(', ')} — and roll straight up into Actual Spent on OVERVIEW.`
    )
  }
  if (state.sheets.budgetTracker) {
    bulletLine(
      'OTHER EXPENSES is for the incidentals with no tab of their own — shopping, tips, fees. Pick a category from the dropdown in column B so the amount lands in the right budget line.'
    )
  }
  if (state.sheets.packingList || state.sheets.tasks) {
    const both = state.sheets.packingList && state.sheets.tasks
    const which = both
      ? 'PACKING LIST and TASKS'
      : state.sheets.packingList
        ? 'PACKING LIST'
        : 'TASKS'
    bulletLine(
      `Tick items off on ${which} by clicking the cell and choosing ✓ from the dropdown — clear the cell to untick. The TRIP READINESS ring on OVERVIEW counts those ticks.`
    )
  }
  spacer()

  // ── Where to go next ────────────────────────────────────────────────────────
  sectionHeader('WHERE TO GO NEXT')
  linkRow('OVERVIEW', 'OVERVIEW', 'Start here — dates, budget, readiness, and a JUMP TO list of every tab.')
  if (state.sheets.itinerary) {
    linkRow('ITINERARY', 'ITINERARY', 'Your day-by-day plan.')
  }
  linkRow(
    'INSTRUCTIONS',
    'INSTRUCTIONS',
    'The full reference: how the budget rolls up, entering dates, changing currency symbols, and what each tab is for.'
  )

  Object.entries(COL_WIDTHS).forEach(([col, width]) => {
    ws.getColumn(col).width = width
  })
}
