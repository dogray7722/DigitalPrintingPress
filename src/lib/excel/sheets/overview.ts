import type ExcelJS from 'exceljs'
import type { SheetId, WizardState } from '../../../types/wizard'
import { SHEET_ICONS } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import type { ExchangeRate, Recommendations } from '../../recommendations/types'
import {
  styleDataCell,
  styleLabelCell,
  styleValueCell,
  truncate,
} from '../styleFactory'
import { computeTotalBudget } from '../../utils'
import { THEMES } from '../../../types/theme'
import { CURRENCIES, getCurrencySymbol, getCurrencyNumFmt } from '../../../data/currencies'
import { getPackingList } from '../../../data/packingLists'
import { DEFAULT_TASKS } from './tasks'
import { addDays } from 'date-fns'

const BUDGET_CATEGORIES = [
  { key: 'Transportation', budget: 'transport', perTrip: false },
  { key: 'Accommodation', budget: 'hotel', perTrip: false },
  { key: 'Food', budget: 'food', perTrip: false },
  { key: 'Activities', budget: 'activities', perTrip: false },
  // Shopping is a per-trip total (like flights), seeded from the wizard's Shopping input.
  { key: 'Shopping', budget: 'shopping', perTrip: true },
  { key: 'Miscellaneous', budget: 'misc', perTrip: false },
] as const

// Per-category estimated budget amounts, in Budget Summary (G15:G20) row order.
// Used to seed the table's Estimated Budget column AND the injected chart's cached
// data points (index.ts) — the two must agree or the chart renders stale numbers.
export function categoryBudgetAmounts(
  state: WizardState
): { key: string; amount: number }[] {
  return BUDGET_CATEGORIES.map(({ key, budget, perTrip }) => ({
    key,
    // Transportation combines the flights lump-sum with the per-day transport budget.
    amount:
      key === 'Transportation'
        ? state.budgets.flights + state.budgets.transport * state.duration
        : perTrip
          ? state.budgets[budget as keyof typeof state.budgets]
          : (state.budgets[budget as keyof typeof state.budgets] as number) * state.duration,
  }))
}

// "Actual Spent" for a category = the sum of the relevant tracker sheet's cost DATA
// ROWS (when that sheet is enabled) PLUS any matching rows logged in OTHER EXPENSES.
// Each term is added only when its source sheet exists, so the formula never
// references a missing sheet (which would corrupt the workbook). Falls back to 0
// when no source applies.
//
// IMPORTANT: sum the tracker DATA ROWS directly — do NOT reference the tracker's
// TOTAL/summary cells (e.g. 'ACCOMMODATION'!$E$26). Both compute the same number,
// but the intermediate cross-sheet formula hop breaks Excel's (Mac) chart repaint:
// the OVERVIEW cells recalc correctly, yet the injected chart doesn't redraw until
// a forced recalc. Direct data-row sums (the Transportation pattern) repaint fine.
// Ranges must match the data rows written by each tracker builder.
function buildActualSpentFormula(category: string, state: WizardState): string {
  const terms: string[] = []

  if (category === 'Accommodation' && state.sheets.hotels) {
    terms.push(`SUM('ACCOMMODATION'!$E$6:$E$25)`)
  }
  if (category === 'Food' && state.sheets.restaurants) {
    terms.push(`SUM('DINING'!$G$6:$G$55)`)
  }
  if (category === 'Activities' && state.sheets.excursions) {
    // EXCURSIONS cost cells are per-row formulas (unit × travelers); summing the H
    // column keeps typed-over values working. One formula hop remains here by design.
    terms.push(`SUM('EXCURSIONS'!$H$6:$H$35)`)
  }
  if (category === 'Transportation' && state.sheets.flights) {
    // All transport (air + ground) from the TRANSPORTATION sheet feeds this single category.
    terms.push(`SUM('TRANSPORTATION'!$H$6:$H$25)`)
  }
  if (state.sheets.budgetTracker) {
    terms.push(
      `SUMIF('OTHER EXPENSES'!$B$8:$B$500,"${category}",'OTHER EXPENSES'!$D$8:$D$500)`
    )
  }

  if (terms.length === 0) return '0'
  return `IFERROR(${terms.join('+')},0)`
}

// startRow = first row of the widget block (header row). Widget spans B:E only.
// E{startRow+2} holds the live exchange rate — user can edit it directly to recalculate.
function buildCurrencyWidget(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  er: ExchangeRate,
  startRow: number,
): void {
  const { from, to, rate, fetchedAt } = er
  const bc = ts.palette.border
  const fromSym = getCurrencySymbol(from)
  const toSym = getCurrencySymbol(to)
  // Strip the quoted symbol prefix from getCurrencyNumFmt to get the plain number format
  // (e.g. '"$"#,##0.00' → '#,##0.00', '"¥"#,##0' → '#,##0').
  const fromNumFmt = getCurrencyNumFmt(from).replace(/^"[^"]*"/, '')
  const toNumFmt = getCurrencyNumFmt(to).replace(/^"[^"]*"/, '')

  const solid = (argb: string): ExcelJS.Fill =>
    ({ type: 'pattern', pattern: 'solid', fgColor: { argb } } as ExcelJS.Fill)
  const thin = (argb: string) => ({ style: 'thin' as const, color: { argb } })
  const dv = ws as ExcelJS.Worksheet & { dataValidations: { add: (a: string, d: ExcelJS.DataValidation) => void } }

  const r0 = startRow       // header
  const r1 = startRow + 1   // left dropdown (B:D) | auto-opposite label (E)
  const r2 = startRow + 2   // rate label (B:D) | editable rate number (E)
  const r3 = startRow + 3   // timestamp
  const r4 = startRow + 4   // input label (B:C) | live symbol (D) | number input (E)
  const r5 = startRow + 5   // result label (B:C) | live symbol (D) | result formula (E)

  // ── Header ──────────────────────────────────────────────────────────────────
  ws.mergeCells(`B${r0}:E${r0}`)
  const hdr = ws.getCell(`B${r0}`)
  hdr.value = 'CURRENCY EXCHANGE'
  hdr.fill = solid(ts.palette.primary)
  hdr.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.primaryText } }
  hdr.alignment = { vertical: 'middle', horizontal: 'center' }

  // ── Left dropdown (user picks either currency); right cell auto-shows the other ─
  ws.mergeCells(`B${r1}:D${r1}`)
  const fromDrop = ws.getCell(`B${r1}`)
  fromDrop.value = from
  fromDrop.fill = solid(ts.palette.secondary)
  fromDrop.font = { name: ts.fontName, size: ts.sizes.header, bold: true, color: { argb: ts.palette.secondaryText } }
  fromDrop.alignment = { vertical: 'middle', horizontal: 'center' }
  fromDrop.border = { bottom: thin(bc) }
  fromDrop.protection = { locked: false }
  dv.dataValidations.add(`B${r1}`, {
    type: 'list', allowBlank: false, formulae: [`"${from},${to}"`], showErrorMessage: false,
  } as ExcelJS.DataValidation)

  const oppCell = ws.getCell(`E${r1}`)
  oppCell.value = { formula: `IF(B${r1}="${from}","${to}","${from}")`, result: to }
  oppCell.fill = solid(ts.palette.secondary)
  oppCell.font = { name: ts.fontName, size: ts.sizes.header, bold: true, color: { argb: ts.palette.secondaryText } }
  oppCell.alignment = { vertical: 'middle', horizontal: 'center' }
  oppCell.border = { bottom: thin(bc) }
  oppCell.protection = { hidden: true }

  // ── Rate row: static "1 FROM =" label (B:D) + user-editable rate number (E) ──
  // E{r2} is the canonical from→to rate. Edit it directly in the sheet to
  // recalculate — the conversion formula below references this cell, not a hidden cell.
  ws.mergeCells(`B${r2}:C${r2}`)
  const rateLbl = ws.getCell(`B${r2}`)
  rateLbl.value = `1 ${from}  =`
  rateLbl.fill = solid(ts.palette.mediumBg)
  rateLbl.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.secondaryText } }
  rateLbl.alignment = { vertical: 'middle', horizontal: 'right' }

  const rateSymCell = ws.getCell(`D${r2}`)
  rateSymCell.value = toSym
  rateSymCell.fill = solid(ts.palette.mediumBg)
  rateSymCell.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.secondaryText } }
  rateSymCell.alignment = { vertical: 'middle', horizontal: 'right' }

  const rateCell = ws.getCell(`E${r2}`)
  rateCell.value = rate
  rateCell.numFmt = '0.0000'
  rateCell.fill = solid(ts.palette.mediumBg)
  rateCell.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.secondaryText } }
  rateCell.alignment = { vertical: 'middle', horizontal: 'left' }
  rateCell.protection = { locked: false }

  // ── Timestamp ─────────────────────────────────────────────────────────────────
  ws.mergeCells(`B${r3}:E${r3}`)
  const stamp = ws.getCell(`B${r3}`)
  stamp.value = `Rate as of ${fetchedAt}`
  stamp.fill = solid(ts.palette.lightBg)
  stamp.font = { name: ts.fontName, size: Math.max(7, ts.sizes.data - 1), italic: true, color: { argb: ts.palette.secondaryText } }
  stamp.alignment = { vertical: 'middle', horizontal: 'center' }

  // ── Input row: label (B:C) | live symbol (D) | number input (E) ─────────────
  // D{r4} is a formula cell that mirrors E{r1} to show the correct symbol
  // (¥ or $) even when the user flips the left dropdown.
  ws.mergeCells(`B${r4}:C${r4}`)
  const inputLbl = ws.getCell(`B${r4}`)
  inputLbl.value = { formula: `CONCATENATE("Amount in ",B${r1},":")`, result: `Amount in ${from}:` }
  inputLbl.fill = solid(ts.palette.lightBg)
  inputLbl.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  inputLbl.alignment = { vertical: 'middle', horizontal: 'right' }
  inputLbl.protection = { hidden: true }

  const inputSym = ws.getCell(`D${r4}`)
  inputSym.value = { formula: `IF(B${r1}="${from}","${fromSym}","${toSym}")`, result: fromSym }
  inputSym.fill = solid(ts.palette.lightBg)
  inputSym.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  inputSym.alignment = { vertical: 'middle', horizontal: 'right' }
  inputSym.protection = { hidden: true }

  const inputCell = ws.getCell(`E${r4}`)
  inputCell.numFmt = fromNumFmt
  inputCell.fill = solid('FFFFFFFF')
  inputCell.font = { name: ts.fontName, size: ts.sizes.data }
  inputCell.alignment = { vertical: 'middle', horizontal: 'left' }
  inputCell.border = { bottom: thin(bc) }
  inputCell.protection = { locked: false }

  // ── Result row: label (B:C) | live symbol (D) | converted amount (E) ─────────
  // D{r5} mirrors the LEFT dropdown so the symbol flips with the direction.
  // E{r5} references E{r2} (the editable rate) — recalculates on any rate change.
  ws.mergeCells(`B${r5}:C${r5}`)
  const resultLbl = ws.getCell(`B${r5}`)
  resultLbl.value = { formula: `CONCATENATE("Converts to ",E${r1},":")`, result: `Converts to ${to}:` }
  resultLbl.fill = solid(ts.palette.mediumBg)
  resultLbl.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  resultLbl.alignment = { vertical: 'middle', horizontal: 'right' }
  resultLbl.protection = { hidden: true }

  const resultSym = ws.getCell(`D${r5}`)
  resultSym.value = { formula: `IF(B${r1}="${from}","${toSym}","${fromSym}")`, result: toSym }
  resultSym.fill = solid(ts.palette.mediumBg)
  resultSym.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  resultSym.alignment = { vertical: 'middle', horizontal: 'right' }
  resultSym.protection = { hidden: true }

  const resultCell = ws.getCell(`E${r5}`)
  resultCell.value = {
    formula: `IF(E${r4}="","",IF(B${r1}="${from}",IFERROR(E${r4}*E${r2},""),IFERROR(E${r4}/E${r2},"")))`,
    result: '',
  }
  resultCell.numFmt = toNumFmt
  resultCell.fill = solid(ts.palette.mediumBg)
  resultCell.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  resultCell.alignment = { vertical: 'middle', horizontal: 'left' }
  resultCell.protection = { hidden: true }
}

// Trip-readiness ring data (hidden cols T/U), feeding an injected doughnut (index.ts).
// Ready = checked packing items + checked tasks; Total = all such items. Each term is
// guarded (IFERROR + sheet toggle) so a disabled sheet never leaves a #REF. A fresh
// workbook's packing/task rows are already pre-populated with real items (only the
// checkmark column starts empty), so the true fresh-workbook total is the actual item
// count — computed here so the cached formula results (and the centered % label, before
// first recalc) are accurate rather than a placeholder 0.
function buildReadinessHelpers(ws: ExcelJS.Worksheet, ts: ThemeStyles, state: WizardState): void {
  if (!state.sheets.packingList && !state.sheets.tasks) return
  const packDone = state.sheets.packingList ? `IFERROR(COUNTIF('PACKING LIST'!$C$5:$C$500,"✓"),0)` : '0'
  const taskDone = state.sheets.tasks ? `IFERROR(COUNTIF('TASKS'!$D$5:$D$500,"✓"),0)` : '0'
  const packTot = state.sheets.packingList ? `IFERROR(COUNTA('PACKING LIST'!$B$5:$B$500),0)` : '0'
  const taskTot = state.sheets.tasks ? `IFERROR(COUNTA('TASKS'!$B$5:$B$500),0)` : '0'

  const packCount = state.sheets.packingList ? getPackingList(state.travelMonth).length : 0
  const taskCount = state.sheets.tasks ? DEFAULT_TASKS.length : 0
  const total = packCount + taskCount

  ws.getCell('U1').value = 'Ready'
  ws.getCell('U2').value = 'To do'
  const t1 = ws.getCell('T1')
  t1.value = { formula: `${packDone}+${taskDone}`, result: 0 } // ready
  t1.protection = { hidden: true }
  const t3 = ws.getCell('T3')
  t3.value = { formula: `${packTot}+${taskTot}`, result: total } // total
  t3.protection = { hidden: true }
  const t2 = ws.getCell('T2')
  t2.value = { formula: `MAX(T3-T1,0)`, result: total } // remaining
  t2.protection = { hidden: true }
  ws.getColumn('T').hidden = true
  ws.getColumn('U').hidden = true

  // "TRIP READINESS" section header. The chart itself (index.ts) carries no title/legend
  // so its plot area — and the doughnut's hole — stays centered in its anchor for the
  // label overlay below. Row 24 is shared with the right column's "SPENT VS PLANNED"
  // header (G24:L24), so style only B:E; the height is set with the right column.
  ws.getCell('B24').value = 'TRIP READINESS'
  ws.mergeCells('B24:E24')
  styleSectionHeaderCells(ws, ts, 24, LEFT_COLS)

  // Centered "% ready" label floated over the doughnut's transparent hole.
  //
  // The overlay must be centered on the CHART FRAME, not eyeballed: the frame is anchored
  // B25:E34 (index.ts) whose `toCol: 5` is an exclusive right edge, so it spans columns
  // B–E only. Merging B29:E30 therefore spans exactly the frame's width, putting the
  // merged cell's center on the frame's center — and, with no title/legend to offset the
  // plot area, on the doughnut hole's center. (Merging C29:E30 instead lands ~6 width
  // units right of center, which pushes the text under the ring.) Vertically, rows 25–34
  // are all RING_ROW_H (18pt) = 180pt, so the frame's mid-line at 90pt falls exactly on
  // the 29/30 boundary — the center of a 29:30 merge. Keep this arithmetic in sync with
  // the anchor, RING_ROW_H, and the B–E column widths.
  const pctCell = ws.getCell('B29')
  pctCell.value = { formula: 'IFERROR(TEXT(T1/T3,"0%"),"0%")', result: '0%' }
  pctCell.protection = { hidden: true }
  ws.mergeCells('B29:E30')
  pctCell.font = { name: ts.fontName, size: ts.sizes.title, bold: true, color: { argb: ts.palette.secondaryText } }
  pctCell.alignment = { vertical: 'middle', horizontal: 'center' }

  const readyCaption = ws.getCell('B31')
  readyCaption.value = 'ready'
  ws.mergeCells('B31:E31')
  readyCaption.font = { name: ts.fontName, size: ts.sizes.data, color: { argb: ts.palette.secondaryText } }
  readyCaption.alignment = { vertical: 'top', horizontal: 'center' }

  // Count caption BELOW the frame (row 35 — the anchor's `toRow: 34` ends at the top of
  // row 35). A freshly generated workbook is always 0% ready, so the ring alone is a
  // featureless single-color circle; the counts make it read as a tracker at zero rather
  // than a failed render.
  const countCaption = ws.getCell('B35')
  countCaption.value = {
    formula: 'IFERROR(T1&" of "&T3&" items complete","")',
    result: `0 of ${total} items complete`,
  }
  countCaption.protection = { hidden: true }
  ws.mergeCells('B35:E35')
  countCaption.font = { name: ts.fontName, size: ts.sizes.data, italic: true }
  countCaption.alignment = { vertical: 'middle', horizontal: 'center' }
}

// Neighborhood guide from recommendations.regions (same data hotels/dining/excursions use).
// Each region takes TWO full-width rows — a bold name row over a muted description row.
// The old one-row form put the name in the 20-wide first column with the description
// merged beside it, so any name longer than ~18 chars ran into (and was clipped by) its
// own description. Stacking them gives both the full G:L width.
//
// These rows share row numbers with the quick-nav link strip in B:E (both blocks start at
// the same header row + 1), so every row here is BOTTOM_ROW_H — an uneven right column
// would show through as a ragged left one. Descriptions are therefore capped to a single
// line at the merged width rather than given a taller wrapped row.
const BOTTOM_ROW_H = 20
// G:L merged width in width-units (20+20+16+16+11+16); ~0.95 chars per unit at data size.
const NEIGHBORHOOD_DESC_CHARS = 92

function buildNeighborhoodGuide(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  regions: Recommendations['regions'],
  startRow: number
): number {
  ws.getCell(`G${startRow}`).value = 'WHERE TO BASE YOURSELF'
  ws.mergeCells(`G${startRow}:L${startRow}`)
  ws.getRow(startRow).height = 22
  styleSectionHeaderCells(ws, ts, startRow, RIGHT_COLS)
  let r = startRow + 1
  regions.slice(0, 5).forEach((region, i) => {
    // Name row — full width, bold.
    styleDataRowCells(ws, ts, r, RIGHT_COLS, i % 2 === 0)
    const nameCell = ws.getCell(`G${r}`)
    nameCell.value = region.region
    nameCell.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
    nameCell.alignment = { vertical: 'middle', horizontal: 'left' }
    ws.mergeCells(`G${r}:L${r}`)
    ws.getRow(r).height = BOTTOM_ROW_H
    r++

    // Description row — same stripe, muted, directly beneath its name.
    styleDataRowCells(ws, ts, r, RIGHT_COLS, i % 2 === 0)
    const descCell = ws.getCell(`G${r}`)
    descCell.value = truncate(region.description ?? '', NEIGHBORHOOD_DESC_CHARS)
    descCell.font = { name: ts.fontName, size: ts.sizes.data, italic: true }
    descCell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' }
    ws.mergeCells(`G${r}:L${r}`)
    ws.getRow(r).height = BOTTOM_ROW_H
    r++
  })
  return r + 1
}

// OVERVIEW stacks two INDEPENDENT columns of content (the B:E panel and the G:L cards)
// that share row numbers, so styleFactory's row-level helpers can't be used here — they
// walk every populated cell in the row and would restyle the neighbouring column's
// content (e.g. a section header on the left bolding a budget data row on the right).
// These apply the same styling to an explicit column range only.
// Column A is an empty left margin and column F the gutter between the two stacks —
// neither is ever styled or populated.
const LEFT_COLS = ['B', 'C', 'D', 'E'] as const
const RIGHT_COLS = ['G', 'H', 'I', 'J', 'K', 'L'] as const

function styleSectionHeaderCells(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  row: number,
  cols: readonly string[]
): void {
  cols.forEach((col) => {
    const cell = ws.getCell(`${col}${row}`)
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.secondary } } as ExcelJS.Fill
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.sectionHeader,
      bold: true,
      color: { argb: ts.palette.secondaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
  })
}

function styleDataRowCells(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  row: number,
  cols: readonly string[],
  isEven: boolean
): void {
  cols.forEach((col) => styleDataCell(ws.getCell(`${col}${row}`), ts, isEven))
}

// Column-scoped twins of styleFactory's styleColumnHeader / styleTotalRow. The Budget
// Summary table shares rows 14–21 with the left column's CURRENCY / TOTAL BUDGET card
// and the currency-exchange widget, so the row-level helpers would restyle those.
function styleColumnHeaderCells(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  row: number,
  cols: readonly string[]
): void {
  const side = { style: 'thin' as const, color: { argb: ts.palette.border } }
  cols.forEach((col) => {
    const cell = ws.getCell(`${col}${row}`)
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.secondary } } as ExcelJS.Fill
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.header,
      bold: true,
      color: { argb: ts.palette.secondaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'center' }
    cell.border = { top: side, bottom: side, left: side, right: side }
  })
}

function styleTotalRowCells(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  row: number,
  cols: readonly string[]
): void {
  const side = { style: 'thin' as const, color: { argb: ts.palette.border } }
  cols.forEach((col) => {
    const cell = ws.getCell(`${col}${row}`)
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.mediumBg } } as ExcelJS.Fill
    cell.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
    cell.border = { top: side, bottom: side, left: side, right: side }
  })
}

// Embeds a base64 data-URL image (JPEG or PNG) into the sheet at the given cell range.
// exceljs's `Image.buffer` type is a module-local `interface Buffer extends ArrayBuffer
// {}`, structurally distinct from Node/polyfilled `Buffer` (Uint8Array) — cast through
// ArrayBuffer, which it's structurally identical to.
function addDataUrlImage(
  wb: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  dataUrl: string,
  extension: 'jpeg' | 'png',
  range: string
): void {
  const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
  const imageId = wb.addImage({ buffer: Buffer.from(base64, 'base64') as unknown as ArrayBuffer, extension })
  ws.addImage(imageId, range)
}

// Quick-nav: one internal HYPERLINK per enabled sheet, merged across B:E. Only enabled
// sheets get a link, so none dangle. Sheet display names come from the builders' tab
// names. Returns the next free row.
function buildQuickNav(ws: ExcelJS.Worksheet, state: WizardState, ts: ThemeStyles, startRow: number): number {
  const links: [boolean, string, string, SheetId][] = [
    [state.sheets.itinerary, 'ITINERARY', 'Itinerary', 'itinerary'],
    [state.sheets.flights, 'TRANSPORTATION', 'Transportation', 'flights'],
    [state.sheets.hotels, 'ACCOMMODATION', 'Accommodation', 'hotels'],
    [state.sheets.restaurants, 'DINING', 'Dining', 'restaurants'],
    [state.sheets.excursions, 'EXCURSIONS', 'Excursions', 'excursions'],
    [state.sheets.packingList, 'PACKING LIST', 'Packing List', 'packingList'],
    [state.sheets.tasks, 'TASKS', 'Tasks', 'tasks'],
    [state.sheets.budgetTracker, 'OTHER EXPENSES', 'Other Expenses', 'budgetTracker'],
  ]
  ws.getCell(`B${startRow}`).value = 'JUMP TO'
  ws.mergeCells(`B${startRow}:E${startRow}`)
  ws.getRow(startRow).height = 22
  styleSectionHeaderCells(ws, ts, startRow, LEFT_COLS)
  let r = startRow + 1
  links.filter(([on]) => on).forEach(([, sheet, label, id], i) => {
    styleDataRowCells(ws, ts, r, LEFT_COLS, i % 2 === 0)
    const cell = ws.getCell(`B${r}`)
    // Icon comes from the shared SHEET_ICONS map (same emoji the Step 3 toggle grid
    // shows). It renders in its own colors and ignores the font color below —
    // only the label text picks up the theme primary.
    const text = `${SHEET_ICONS[id]}  ${label}`
    cell.value = { formula: `HYPERLINK("#'${sheet}'!A1","${text}")`, result: text }
    cell.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.primary } }
    ws.mergeCells(`B${r}:E${r}`)
    ws.getRow(r).height = BOTTOM_ROW_H
    r++
  })
  return r
}

export function buildOverviewSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  exchangeRate?: ExchangeRate,
  recommendations?: Recommendations
): void {
  const ws = wb.addWorksheet('OVERVIEW')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  const dest = state.destination || 'My Trip'
  const startDate = state.startDate ? new Date(state.startDate) : new Date()
  const endDate = addDays(startDate, state.duration - 1)
  const totalBudget = computeTotalBudget(state.budgets, state.duration)
  const currSym = getCurrencySymbol(state.currency)

  // The currency-exchange widget occupies B17:E22 when it's built (skipped when source
  // and target currency match). Row 22 is otherwise just part of the gutter above the
  // row-24 header band, so it only needs full height when the widget's last row lands
  // there. Row 23 is a blank gutter either way, so the widget never butts into the
  // TRIP READINESS / SPENT VS PLANNED header.
  const hasCurrencyWidget = !!(exchangeRate && exchangeRate.from !== exchangeRate.to)

  // ── Row heights ─────────────────────────────────────────────────────────────
  // OVERVIEW stacks two INDEPENDENT columns of content (the B:E card stack and the G:L
  // cards) that nonetheless SHARE row numbers, so every height here is the resolution of
  // both sides' needs — set them in one place rather than letting each block set its own,
  // or one side silently squashes the other.
  //
  // The left column is a stack of label/value cards: a short label row, a taller value
  // row, a thin gutter, repeat. The right column is the cover photo, the Budget Summary
  // table, the spend chart and the neighborhood guide. Section headers are deliberately
  // LEVEL across the F gutter — row 13 (BUDGET SUMMARY), row 24 (TRIP READINESS |
  // SPENT VS PLANNED), row 38 (JUMP TO | WHERE TO BASE YOURSELF) — which is what gives
  // the sheet a horizontal rhythm instead of two independently drifting columns.
  //
  // Rows 1–11 are also the cover photo's anchor, so their TOTAL (260pt) sets the photo's
  // aspect ratio against the G:L width — see the addDataUrlImage call at the bottom, and
  // keep PictureUploader's crop target in sync with any change here.
  //
  // RING_ROW_H is load-bearing: it sets the readiness doughnut's frame height (rows
  // 25–34 = 180pt), and the centered "% ready" overlay in buildReadinessHelpers is
  // merged across the row pair straddling that frame's mid-line. Changing it means
  // re-deriving that merge.
  const RING_ROW_H = 18
  const rowHeights: Record<number, number> = {
    1: 24, // hero panel top padding | cover photo
    2: 34,
    3: 34, // destination headline (B2:E3)
    4: 8, // gutter
    5: 18, // labels: START DATE | END DATE
    6: 30, // values: TripStart | TripEnd
    7: 10, // gutter
    8: 18, // labels: DAYS | PARTY | TRAVELERS
    9: 30, // values
    10: 14, // gutter
    11: 40, // ← last row of the cover photo
    12: 24, // countdown number (B11:E12) | gutter under the photo
    13: 22, // "DAYS TO GO" caption | BUDGET SUMMARY header
    14: 20, // hero panel bottom padding | Budget Summary column headers
    15: 20, // labels: CURRENCY | TOTAL BUDGET ‖ budget category rows 15–20
    16: 20, // values: currency | TotalBudget
    17: 20, // currency-exchange widget starts here (B17:E22)
    18: 20,
    19: 20,
    20: 20,
    21: 20, // Budget Summary TOTAL row
    22: hasCurrencyWidget ? 18 : 8, // widget's last row, else gutter
    23: 8, // blank gutter under the currency-exchange widget
    24: 22, // TRIP READINESS | SPENT VS PLANNED headers
    35: 20, // "N of M items complete" caption, just below the ring frame
    36: RING_ROW_H, // last row of the spend chart frame
    37: 8, // gutter
    38: 22, // JUMP TO | WHERE TO BASE YOURSELF headers
  }
  // Rows 25–34: readiness doughnut (B:E) beside the spend chart (G:L, which runs to 36).
  for (let r = 25; r <= 34; r++) rowHeights[r] = RING_ROW_H
  Object.entries(rowHeights).forEach(([r, h]) => {
    ws.getRow(Number(r)).height = h
  })

  // ── Hero band (rows 1–14): dark text panel (B:E) beside the cover photo (G:L) ──
  // Excel can't truly composite editable cell text on top of an opaque floating image
  // (drawings always render above the grid, unlike CSS layering), so the destination/
  // dates/party/countdown live in a panel immediately BESIDE the photo instead. Column A
  // is an empty left margin so the panel floats off the sheet edge, and column F is a
  // deliberately empty, unfilled spacer separating the panel from the photo.
  //
  // Each card is a LABEL row over a VALUE row, both merged across whole columns. The
  // labels are what let the values be plain, genuinely editable raw dates/numbers: an
  // earlier pass had no captions at all and folded the punctuation into each value's
  // number format ("–  "mmm d, "0 days") to compensate, which read as guesswork.
  const heroFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.primary } } as ExcelJS.Fill
  const heroFont = (size: number, bold = true) =>
    ({ name: ts.fontName, size, bold, color: { argb: ts.palette.primaryText } }) as ExcelJS.Font
  // Caption style inside the panel: the light label/value pair used lower down (rows
  // 15–16) can't be reused here — its lightBg fill would punch holes in the dark band —
  // so it's the same small/bold treatment in the panel's own foreground color.
  const heroLabelFont = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.primaryText } } as ExcelJS.Font
  for (let r = 1; r <= 14; r++) {
    LEFT_COLS.forEach((col) => {
      ws.getCell(`${col}${r}`).fill = heroFill
    })
  }

  // Caption cell inside the hero panel. `range` may be a single cell or a merge range.
  const heroLabel = (range: string, text: string): void => {
    const anchor = range.split(':')[0]
    const cell = ws.getCell(anchor)
    cell.value = text
    cell.font = heroLabelFont
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
    if (range.includes(':')) ws.mergeCells(range)
  }

  // Destination headline. Long names step down a size rather than clipping — the panel
  // is ~58 width units and the title size is set for a short name like "NEW YORK".
  ws.mergeCells('B2:E3')
  const headline = ws.getCell('B2')
  const headlineText = truncate(dest.toUpperCase(), 34)
  headline.value = headlineText
  headline.font = heroFont(headlineText.length > 18 ? Math.round(ts.sizes.title * 0.72) : ts.sizes.title)
  headline.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }

  // ── Rows 5–6: date card ─────────────────────────────────────────────────────
  // Two independently-editable date cells, each captioned and merged across half the
  // panel. Both carry the year so a range spanning a year boundary still reads correctly.
  const dateFont = heroFont(Math.round(ts.sizes.title * 0.65))

  heroLabel('B5:C5', 'START DATE')
  heroLabel('D5:E5', 'END DATE')

  // B6 = TripStart (named range anchor) — user-editable
  const tripStartCell = ws.getCell('B6')
  tripStartCell.value = startDate
  tripStartCell.numFmt = 'mmm d, yyyy'
  tripStartCell.protection = { locked: false }
  tripStartCell.font = dateFont
  tripStartCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.mergeCells('B6:C6')

  // D6 = TripEnd (named range anchor) — user-editable
  const tripEndCell = ws.getCell('D6')
  tripEndCell.value = endDate
  tripEndCell.numFmt = 'mmm d, yyyy'
  tripEndCell.protection = { locked: false }
  tripEndCell.font = dateFont
  tripEndCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.mergeCells('D6:E6')

  // ── Rows 8–9: trip-shape card ───────────────────────────────────────────────
  // Duration is derived (inclusive day count, not nights — TripEnd-TripStart+1); party
  // type and traveler count stay editable.
  const lineFont = heroFont(ts.sizes.header)

  heroLabel('B8', 'DAYS')
  heroLabel('C8:D8', 'PARTY')
  heroLabel('E8', 'TRAVELERS')

  const daysCell = ws.getCell('B9')
  daysCell.value = { formula: 'IFERROR(TripEnd-TripStart+1,"")', result: state.duration }
  daysCell.numFmt = '0'
  daysCell.protection = { hidden: true }
  daysCell.font = lineFont
  daysCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }

  const partyCell = ws.getCell('C9')
  partyCell.value = state.partyType.charAt(0).toUpperCase() + state.partyType.slice(1)
  partyCell.protection = { locked: false }
  partyCell.font = lineFont
  partyCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.mergeCells('C9:D9')
  ;(ws as any).dataValidations.add('C9', {
    type: 'list',
    allowBlank: false,
    formulae: ['"Solo,Couple,Family,Group"'],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // E9 = NumAdults (named range anchor) — user-editable
  const travelersCell = ws.getCell('E9')
  travelersCell.value = state.partySize
  travelersCell.numFmt = '0'
  travelersCell.protection = { locked: false }
  travelersCell.font = lineFont
  travelersCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }

  // ── Rows 11–13: countdown badge ─────────────────────────────────────────────
  // The mockup's big circular "292 / DAYS TO GO" marker. Excel can't draw it as a circle
  // over the photo, so it's an accent-filled block spanning the full panel width: the
  // count as one oversized number (B11:E12), the unit as a caption beneath (B13:E13).
  // Two cells rather than one sentence so the number can carry its own display size.
  // Precomputed cached results so both render immediately on open (Excel/Sheets won't
  // trigger an initial recalc on this freshly written file).
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  const startDay = new Date(startDate.getTime())
  startDay.setHours(0, 0, 0, 0)
  const daysUntil = Math.round((startDay.getTime() - today.getTime()) / 86400000)
  const badgeFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.secondary } } as ExcelJS.Fill
  LEFT_COLS.forEach((col) => {
    ;[11, 12, 13].forEach((r) => {
      ws.getCell(`${col}${r}`).fill = badgeFill
    })
  })

  const badgeCount = ws.getCell('B11')
  badgeCount.value = {
    formula: 'IFERROR(ABS(INT(TripStart-TODAY())),"—")',
    result: Math.abs(daysUntil),
  }
  badgeCount.protection = { hidden: true }
  badgeCount.font = {
    name: ts.fontName,
    size: Math.round(ts.sizes.title * 1.6),
    bold: true,
    color: { argb: ts.palette.secondaryText },
  }
  badgeCount.alignment = { vertical: 'middle', horizontal: 'center' }
  ws.mergeCells('B11:E12')

  const badgeUnit = ws.getCell('B13')
  badgeUnit.value = {
    formula: 'IFERROR(IF(TripStart=TODAY(),"TODAY!",IF(TripStart>TODAY(),"DAYS TO GO","DAYS AGO")),"")',
    result: daysUntil === 0 ? 'TODAY!' : daysUntil > 0 ? 'DAYS TO GO' : 'DAYS AGO',
  }
  badgeUnit.protection = { hidden: true }
  badgeUnit.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  badgeUnit.alignment = { vertical: 'top', horizontal: 'center' }
  ws.mergeCells('B13:E13')

  // Column widths for the whole B:E stack — the hero band, the row 15–16 card, the
  // currency widget, the readiness ring and quick-nav all share these, so each column is
  // sized for the WIDEST thing that lands in it, at the largest font that lands in it
  // (the hero rows run at title/header sizes, well above the 11pt default these width
  // units assume). Undersizing here is what produced "###" and clipped labels.
  //   A: empty left margin — keeps the dark panel off the sheet edge (never filled)
  //   B: "DAYS" · day count · quick-nav links · left half of the date/currency cards
  //   C: rest of the date/currency card merges · party type
  //   D: "END DATE" / "TOTAL BUDGET" merge anchors
  //   E: travelers · right half of those merges
  //   F: empty spacer between the panel and the cover photo (never filled)
  ws.getColumn('A').width = 3
  ws.getColumn('B').width = 10
  ws.getColumn('C').width = 18
  ws.getColumn('D').width = 16
  ws.getColumn('E').width = 14
  ws.getColumn('F').width = 6

  // ── Rows 15–16: Currency + Total Budget card ────────────────────────────────
  // Same label-over-value shape as the hero cards, but below the dark panel, so it uses
  // the shared light caption/value styles. Merged in halves (B:C | D:E) so the two
  // labels can't clip each other the way they did when all four sat on one row.
  const cardLabel = (range: string, text: string): void => {
    const anchor = range.split(':')[0]
    const cell = ws.getCell(anchor)
    cell.value = text
    styleLabelCell(cell, ts)
    ws.mergeCells(range)
  }
  cardLabel('B15:C15', 'CURRENCY')
  cardLabel('D15:E15', 'TOTAL BUDGET')

  // Currency — informational label (chosen in the wizard), a dropdown so the user can
  // pick a different reference currency. Options are written to a hidden helper column
  // (S) because the full "CODE (symbol)" list exceeds Excel's 255-char inline list-
  // formula limit. Purely a label — does not affect any numFmt elsewhere (those are
  // static, baked in at generation time).
  const currValue = ws.getCell('B16')
  currValue.value = `${state.currency} (${currSym})`
  currValue.protection = { locked: false }
  styleValueCell(currValue, ts)
  ws.mergeCells('B16:C16')

  CURRENCIES.forEach((c, i) => {
    ws.getCell(`S${i + 1}`).value = `${c.code} (${c.symbol})`
  })
  ;(ws as any).dataValidations.add('B16', {
    type: 'list',
    allowBlank: true,
    formulae: [`$S$1:$S${CURRENCIES.length}`],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // D16 = TotalBudget (named range anchor) — live sum of the Estimated Budget column,
  // so it reflects any edits the user types into those cells. Cached result keeps the
  // cell populated before the first recalc.
  const budgetValue = ws.getCell('D16')
  budgetValue.value = { formula: 'SUM(H15:H20)', result: totalBudget }
  budgetValue.numFmt = ts.numFmtCurrency
  budgetValue.protection = { hidden: true }
  budgetValue.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  budgetValue.alignment = { vertical: 'middle', horizontal: 'left' }
  budgetValue.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.mediumBg } } as ExcelJS.Fill
  ws.mergeCells('D16:E16')

  // ── Budget Summary Table (rows 13–21) ───────────────────────────────────────
  // Header sits at row 13 with a gutter (row 12) between it and the cover photo, which
  // stops at row 11 instead of butting straight into this band.
  ws.getCell('G13').value = 'BUDGET SUMMARY'
  ws.mergeCells('G13:L13')
  styleSectionHeaderCells(ws, ts, 13, RIGHT_COLS)

  ws.getCell('G14').value = 'Category'
  ws.getCell('H14').value = 'Estimated Budget'
  ws.getCell('I14').value = 'Actual Spent'
  ws.getCell('J14').value = 'Remaining'
  ws.getCell('K14').value = '% Used'
  ws.getCell('L14').value = 'Status'
  styleColumnHeaderCells(ws, ts, 14, RIGHT_COLS)

  let rowIdx = 15
  categoryBudgetAmounts(state).forEach(({ key, amount: budgetAmt }, i) => {
    ws.getCell(`G${rowIdx}`).value = key

    // Estimated Budget (col H) — a plain, editable starting value derived from the
    // wizard's budget selections. It is NOT a formula: it doesn't change when the trip
    // dates change. Users simply type their own figure over it.
    const hCell = ws.getCell(`H${rowIdx}`)
    hCell.value = budgetAmt
    hCell.numFmt = ts.numFmtCurrency
    hCell.protection = { locked: false }

    // Actual spent = per-category sheet totals + matching OTHER EXPENSES entries.
    // Cached results throughout this table = the empty-tracker evaluation, so cells
    // render immediately in viewers that don't recalc on open (see countdown note).
    const iCell = ws.getCell(`I${rowIdx}`)
    iCell.value = { formula: buildActualSpentFormula(key, state), result: 0 }
    iCell.numFmt = ts.numFmtCurrency
    iCell.protection = { hidden: true }

    // Remaining
    const jCell = ws.getCell(`J${rowIdx}`)
    jCell.value = { formula: `IFERROR(H${rowIdx}-I${rowIdx},H${rowIdx})`, result: budgetAmt }
    jCell.numFmt = ts.numFmtCurrency
    jCell.protection = { hidden: true }

    // % Used
    const kCell = ws.getCell(`K${rowIdx}`)
    kCell.value = { formula: `IFERROR(I${rowIdx}/H${rowIdx},0)`, result: 0 }
    kCell.numFmt = ts.numFmtPercent
    kCell.protection = { hidden: true }

    // Status
    const lCell = ws.getCell(`L${rowIdx}`)
    lCell.value = {
      formula: `IFERROR(IF(I${rowIdx}>H${rowIdx},"Over budget",IF(I${rowIdx}/H${rowIdx}>0.8,"Near limit","On track")),"On track")`,
      result: 'On track',
    }
    lCell.protection = { hidden: true }

    // Col N: pie/donut chart helper — actual spend once logged, else budget estimate so
    // the chart is never empty when the user hasn't entered spending yet.
    const nCell = ws.getCell(`N${rowIdx}`)
    nCell.value = { formula: `IF(I${rowIdx}>0,I${rowIdx},H${rowIdx})`, result: budgetAmt }
    nCell.numFmt = ts.numFmtCurrency
    nCell.protection = { hidden: true }

    // Col O: bar chart helper — MIN(actual, budget), the colored "actual spent" segment
    // (base of the stacked bar).
    const oCell = ws.getCell(`O${rowIdx}`)
    oCell.value = { formula: `MIN(I${rowIdx},H${rowIdx})`, result: 0 }
    oCell.numFmt = ts.numFmtCurrency
    oCell.protection = { hidden: true }

    // Col P: bar chart helper — MAX(budget - actual, 0), the light "remaining budget"
    // segment stacked after col O so one bar = full budget (matches the Excel look in
    // Google Sheets, which ignores overlap on clustered charts).
    const pCell = ws.getCell(`P${rowIdx}`)
    pCell.value = { formula: `MAX(H${rowIdx}-I${rowIdx},0)`, result: budgetAmt }
    pCell.numFmt = ts.numFmtCurrency
    pCell.protection = { hidden: true }

    // Col Q: bar chart helper — MAX(actual - budget, 0), the red "over budget" overage
    // segment stacked after col P. Zero (invisible) when on/under budget; when spending
    // exceeds the estimate it extends the bar past the budget length in red, mirroring
    // the "turns red when over budget" preview shown in the wizard's chart picker.
    const qCell = ws.getCell(`Q${rowIdx}`)
    qCell.value = { formula: `MAX(I${rowIdx}-H${rowIdx},0)`, result: 0 }
    qCell.numFmt = ts.numFmtCurrency
    qCell.protection = { hidden: true }

    styleDataRowCells(ws, ts, rowIdx, RIGHT_COLS, i % 2 === 0)
    rowIdx++
  })

  // Totals row (row 21)
  ws.getCell(`G${rowIdx}`).value = 'TOTAL'
  const hTot = ws.getCell(`H${rowIdx}`)
  hTot.value = { formula: `SUM(H15:H${rowIdx - 1})`, result: totalBudget }
  hTot.numFmt = ts.numFmtCurrency
  hTot.protection = { hidden: true }
  const iTot = ws.getCell(`I${rowIdx}`)
  iTot.value = { formula: `SUM(I15:I${rowIdx - 1})`, result: 0 }
  iTot.numFmt = ts.numFmtCurrency
  iTot.protection = { hidden: true }
  const jTot = ws.getCell(`J${rowIdx}`)
  jTot.value = { formula: `SUM(J15:J${rowIdx - 1})`, result: totalBudget }
  jTot.numFmt = ts.numFmtCurrency
  jTot.protection = { hidden: true }
  styleTotalRowCells(ws, ts, rowIdx, RIGHT_COLS)

  // ── Currency Exchange Widget (cols B–E, rows 17–22) ─────────────────────────
  // Sits left of the budget table (G13:L21) — same row band, different columns; the
  // negative space beside the table is exactly the six rows it needs, and row 23 below
  // it is a blank gutter.
  // Built after the budget table so its fills win on any shared row.
  // Skip when source/target match — a same-currency converter (rate 1.0) is meaningless.
  if (hasCurrencyWidget && exchangeRate) {
    buildCurrencyWidget(ws, ts, exchangeRate, 17)
  }

  // ── Spend chart (rows 24–36) ────────────────────────────────────────────────
  // In-cell title so the injected chart itself carries none — a chart-space title eats
  // plot height and doesn't line up with anything else on the sheet, whereas this band
  // sits level with TRIP READINESS across the F gutter. Chart anchored G25:L36 (index.ts).
  ws.getCell('G24').value = 'SPENT VS PLANNED'
  ws.mergeCells('G24:L24')
  styleSectionHeaderCells(ws, ts, 24, RIGHT_COLS)

  // Trip-readiness ring: header row 24, chart injected in index.ts anchored B25:E34.
  buildReadinessHelpers(ws, ts, state)

  // Quick-nav HYPERLINK list, lower-left — header level with the neighborhood guide's.
  buildQuickNav(ws, state, ts, 38)

  // Neighborhood guide, right column below the spend chart.
  if (state.useRecommendations && recommendations?.regions?.length) {
    buildNeighborhoodGuide(ws, ts, recommendations.regions, 38)
  }

  // ── Visual Budget Breakdown ─────────────────────────────────────────────────
  // The breakdown is a native chart object injected after export for every style
  // (bar / pie / donut) — see chartInjection.ts. Nothing to render in-cell here.

  // ── Column widths ───────────────────────────────────────────────────────────
  // A–F are set with the hero band above (they're sized for its larger fonts) — do not
  // re-set them here or that sizing is silently discarded.
  ws.getColumn('G').width = 20
  ws.getColumn('H').width = 20
  ws.getColumn('I').width = 16
  ws.getColumn('J').width = 16
  ws.getColumn('K').width = 11
  ws.getColumn('L').width = 16

  // Hide the chart's helper columns (chart plots hidden cells — see plotVisOnly in chartInjection)
  ws.getColumn('N').hidden = true
  ws.getColumn('O').hidden = true
  ws.getColumn('P').hidden = true
  ws.getColumn('Q').hidden = true

  // Hide the currency-dropdown options helper column (S)
  ws.getColumn('S').hidden = true

  // ── Cover photo (optional, anchored G1:L11) ─────────────────────────────────
  // Stops at row 11, three rows short of the text panel's bottom, so the BUDGET SUMMARY
  // header at row 13 gets a gutter instead of butting into the photo. PictureUploader's
  // crop target (TARGET_W/TARGET_H) must match THIS range's aspect ratio — G:L is 99
  // width units (~693px) and rows 1–11 total 260pt (~347px), i.e. ~2:1 — or the embedded
  // photo stretches. Changing any row height in 1–11 changes that ratio.
  if (state.overviewImage) {
    addDataUrlImage(wb, ws, state.overviewImage.dataUrl, 'jpeg', 'G1:L11')
  }

  // OVERVIEW is a one-screen dashboard, not a scrollable table — no freeze pane (every
  // other sheet builder freezes its header rows; this is a deliberate OVERVIEW-only
  // exception) and no gridlines, so the hero band + cards read as a dashboard.
  ws.views = [{ showGridLines: false }]
}
