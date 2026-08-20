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

// Per-category estimated budget amounts, in Budget Summary (F15:F20) row order.
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

// startRow = first row of the widget block (header row). Widget spans A:D only.
// D{startRow+2} holds the live exchange rate — user can edit it directly to recalculate.
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
  const r1 = startRow + 1   // left dropdown (A:C) | auto-opposite label (D)
  const r2 = startRow + 2   // rate label (A:C) | editable rate number (D)
  const r3 = startRow + 3   // timestamp
  const r4 = startRow + 4   // input label (A:B) | live symbol (C) | number input (D)
  const r5 = startRow + 5   // result label (A:B) | live symbol (C) | result formula (D)

  // ── Header ──────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${r0}:D${r0}`)
  const hdr = ws.getCell(`A${r0}`)
  hdr.value = 'CURRENCY EXCHANGE'
  hdr.fill = solid(ts.palette.primary)
  hdr.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.primaryText } }
  hdr.alignment = { vertical: 'middle', horizontal: 'center' }

  // ── Left dropdown (user picks either currency); right cell auto-shows the other ─
  ws.mergeCells(`A${r1}:C${r1}`)
  const fromDrop = ws.getCell(`A${r1}`)
  fromDrop.value = from
  fromDrop.fill = solid(ts.palette.secondary)
  fromDrop.font = { name: ts.fontName, size: ts.sizes.header, bold: true, color: { argb: ts.palette.secondaryText } }
  fromDrop.alignment = { vertical: 'middle', horizontal: 'center' }
  fromDrop.border = { bottom: thin(bc) }
  fromDrop.protection = { locked: false }
  dv.dataValidations.add(`A${r1}`, {
    type: 'list', allowBlank: false, formulae: [`"${from},${to}"`], showErrorMessage: false,
  } as ExcelJS.DataValidation)

  const oppCell = ws.getCell(`D${r1}`)
  oppCell.value = { formula: `IF(A${r1}="${from}","${to}","${from}")`, result: to }
  oppCell.fill = solid(ts.palette.secondary)
  oppCell.font = { name: ts.fontName, size: ts.sizes.header, bold: true, color: { argb: ts.palette.secondaryText } }
  oppCell.alignment = { vertical: 'middle', horizontal: 'center' }
  oppCell.border = { bottom: thin(bc) }
  oppCell.protection = { hidden: true }

  // ── Rate row: static "1 FROM =" label (A:C) + user-editable rate number (D) ──
  // D{r2} is the canonical from→to rate. Edit it directly in the sheet to
  // recalculate — the conversion formula below references this cell, not a hidden cell.
  ws.mergeCells(`A${r2}:B${r2}`)
  const rateLbl = ws.getCell(`A${r2}`)
  rateLbl.value = `1 ${from}  =`
  rateLbl.fill = solid(ts.palette.mediumBg)
  rateLbl.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.secondaryText } }
  rateLbl.alignment = { vertical: 'middle', horizontal: 'right' }

  const rateSymCell = ws.getCell(`C${r2}`)
  rateSymCell.value = toSym
  rateSymCell.fill = solid(ts.palette.mediumBg)
  rateSymCell.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.secondaryText } }
  rateSymCell.alignment = { vertical: 'middle', horizontal: 'right' }

  const rateCell = ws.getCell(`D${r2}`)
  rateCell.value = rate
  rateCell.numFmt = '0.0000'
  rateCell.fill = solid(ts.palette.mediumBg)
  rateCell.font = { name: ts.fontName, size: ts.sizes.sectionHeader, bold: true, color: { argb: ts.palette.secondaryText } }
  rateCell.alignment = { vertical: 'middle', horizontal: 'left' }
  rateCell.protection = { locked: false }

  // ── Timestamp ─────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${r3}:D${r3}`)
  const stamp = ws.getCell(`A${r3}`)
  stamp.value = `Rate as of ${fetchedAt}`
  stamp.fill = solid(ts.palette.lightBg)
  stamp.font = { name: ts.fontName, size: Math.max(7, ts.sizes.data - 1), italic: true, color: { argb: ts.palette.secondaryText } }
  stamp.alignment = { vertical: 'middle', horizontal: 'center' }

  // ── Input row: label (A:B) | live symbol (C) | number input (D) ─────────────
  // C{r4} is a formula cell that mirrors D{r1} to show the correct symbol
  // (¥ or $) even when the user flips the left dropdown.
  ws.mergeCells(`A${r4}:B${r4}`)
  const inputLbl = ws.getCell(`A${r4}`)
  inputLbl.value = { formula: `CONCATENATE("Amount in ",A${r1},":")`, result: `Amount in ${from}:` }
  inputLbl.fill = solid(ts.palette.lightBg)
  inputLbl.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  inputLbl.alignment = { vertical: 'middle', horizontal: 'right' }
  inputLbl.protection = { hidden: true }

  const inputSym = ws.getCell(`C${r4}`)
  inputSym.value = { formula: `IF(A${r1}="${from}","${fromSym}","${toSym}")`, result: fromSym }
  inputSym.fill = solid(ts.palette.lightBg)
  inputSym.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  inputSym.alignment = { vertical: 'middle', horizontal: 'right' }
  inputSym.protection = { hidden: true }

  const inputCell = ws.getCell(`D${r4}`)
  inputCell.numFmt = fromNumFmt
  inputCell.fill = solid('FFFFFFFF')
  inputCell.font = { name: ts.fontName, size: ts.sizes.data }
  inputCell.alignment = { vertical: 'middle', horizontal: 'left' }
  inputCell.border = { bottom: thin(bc) }
  inputCell.protection = { locked: false }

  // ── Result row: label (A:B) | live symbol (C) | converted amount (D) ─────────
  // C{r5} mirrors the LEFT dropdown so the symbol flips with the direction.
  // D{r5} references D{r2} (the editable rate) — recalculates on any rate change.
  ws.mergeCells(`A${r5}:B${r5}`)
  const resultLbl = ws.getCell(`A${r5}`)
  resultLbl.value = { formula: `CONCATENATE("Converts to ",D${r1},":")`, result: `Converts to ${to}:` }
  resultLbl.fill = solid(ts.palette.mediumBg)
  resultLbl.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  resultLbl.alignment = { vertical: 'middle', horizontal: 'right' }
  resultLbl.protection = { hidden: true }

  const resultSym = ws.getCell(`C${r5}`)
  resultSym.value = { formula: `IF(A${r1}="${from}","${toSym}","${fromSym}")`, result: toSym }
  resultSym.fill = solid(ts.palette.mediumBg)
  resultSym.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  resultSym.alignment = { vertical: 'middle', horizontal: 'right' }
  resultSym.protection = { hidden: true }

  const resultCell = ws.getCell(`D${r5}`)
  resultCell.value = {
    formula: `IF(D${r4}="","",IF(A${r1}="${from}",IFERROR(D${r4}*D${r2},""),IFERROR(D${r4}/D${r2},"")))`,
    result: '',
  }
  resultCell.numFmt = toNumFmt
  resultCell.fill = solid(ts.palette.mediumBg)
  resultCell.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  resultCell.alignment = { vertical: 'middle', horizontal: 'left' }
  resultCell.protection = { hidden: true }
}

// Trip-readiness ring data (hidden cols S/T), feeding an injected doughnut (index.ts).
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

  ws.getCell('T1').value = 'Ready'
  ws.getCell('T2').value = 'To do'
  const s1 = ws.getCell('S1')
  s1.value = { formula: `${packDone}+${taskDone}`, result: 0 } // ready
  s1.protection = { hidden: true }
  const s3 = ws.getCell('S3')
  s3.value = { formula: `${packTot}+${taskTot}`, result: total } // total
  s3.protection = { hidden: true }
  const s2 = ws.getCell('S2')
  s2.value = { formula: `MAX(S3-S1,0)`, result: total } // remaining
  s2.protection = { hidden: true }
  ws.getColumn('S').hidden = true
  ws.getColumn('T').hidden = true

  // "TRIP READINESS" section header. The chart itself (index.ts) carries no title/legend
  // so its plot area — and the doughnut's hole — stays centered in its anchor for the
  // label overlay below. Row 23 is shared with the right column's "SPENT VS PLANNED"
  // header (F23:K23), so style only A:D; the height is set with the right column.
  ws.getCell('A23').value = 'TRIP READINESS'
  ws.mergeCells('A23:D23')
  styleSectionHeaderCells(ws, ts, 23, LEFT_COLS)

  // Centered "% ready" label floated over the doughnut's transparent hole.
  //
  // The overlay must be centered on the CHART FRAME, not eyeballed: the frame is anchored
  // A24:D33 (index.ts) whose `toCol: 4` is an exclusive right edge, so it spans columns
  // A–D only. Merging A28:D29 therefore spans exactly the frame's width, putting the
  // merged cell's center on the frame's center — and, with no title/legend to offset the
  // plot area, on the doughnut hole's center. (Merging B28:D29 instead lands ~6 width
  // units right of center, which pushes the text under the ring.) Vertically, rows 24–33
  // are all RING_ROW_H (18pt) = 180pt, so the frame's mid-line at 90pt falls exactly on
  // the 28/29 boundary — the center of a 28:29 merge. Keep this arithmetic in sync with
  // the anchor, RING_ROW_H, and the A–D column widths.
  const pctCell = ws.getCell('A28')
  pctCell.value = { formula: 'IFERROR(TEXT(S1/S3,"0%"),"0%")', result: '0%' }
  pctCell.protection = { hidden: true }
  ws.mergeCells('A28:D29')
  pctCell.font = { name: ts.fontName, size: ts.sizes.title, bold: true, color: { argb: ts.palette.secondaryText } }
  pctCell.alignment = { vertical: 'middle', horizontal: 'center' }

  const readyCaption = ws.getCell('A30')
  readyCaption.value = 'ready'
  ws.mergeCells('A30:D30')
  readyCaption.font = { name: ts.fontName, size: ts.sizes.data, color: { argb: ts.palette.secondaryText } }
  readyCaption.alignment = { vertical: 'top', horizontal: 'center' }

  // Count caption BELOW the frame (row 34 — the anchor's `toRow: 33` ends at the top of
  // row 34). A freshly generated workbook is always 0% ready, so the ring alone is a
  // featureless single-color circle; the counts make it read as a tracker at zero rather
  // than a failed render.
  const countCaption = ws.getCell('A34')
  countCaption.value = {
    formula: 'IFERROR(S1&" of "&S3&" items complete","")',
    result: `0 of ${total} items complete`,
  }
  countCaption.protection = { hidden: true }
  ws.mergeCells('A34:D34')
  countCaption.font = { name: ts.fontName, size: ts.sizes.data, italic: true }
  countCaption.alignment = { vertical: 'middle', horizontal: 'center' }
}

// Neighborhood guide from recommendations.regions (same data hotels/dining/excursions use).
// Each region takes TWO full-width rows — a bold name row over a muted description row.
// The old one-row form put the name in the 20-wide F column with the description merged
// across G:K beside it, so any name longer than ~18 chars ran into (and was clipped by)
// its own description. Stacking them gives both the full F:K width.
//
// These rows share row numbers with the quick-nav link strip in A:D (both blocks start
// at BOTTOM_BLOCK_ROW + 1), so every row here is BOTTOM_ROW_H — an uneven right column
// would show through as a ragged left one. Descriptions are therefore capped to a single
// line at the merged width rather than given a taller wrapped row.
const BOTTOM_ROW_H = 20
// F:K merged width in width-units (20+20+16+16+11+16); ~0.95 chars per unit at data size.
const NEIGHBORHOOD_DESC_CHARS = 92

function buildNeighborhoodGuide(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  regions: Recommendations['regions'],
  startRow: number
): number {
  ws.getCell(`F${startRow}`).value = 'WHERE TO BASE YOURSELF'
  ws.mergeCells(`F${startRow}:K${startRow}`)
  ws.getRow(startRow).height = 22
  styleSectionHeaderCells(ws, ts, startRow, RIGHT_COLS)
  let r = startRow + 1
  regions.slice(0, 5).forEach((region, i) => {
    // Name row — full width, bold.
    styleDataRowCells(ws, ts, r, RIGHT_COLS, i % 2 === 0)
    const nameCell = ws.getCell(`F${r}`)
    nameCell.value = region.region
    nameCell.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
    nameCell.alignment = { vertical: 'middle', horizontal: 'left' }
    ws.mergeCells(`F${r}:K${r}`)
    ws.getRow(r).height = BOTTOM_ROW_H
    r++

    // Description row — same stripe, muted, directly beneath its name.
    styleDataRowCells(ws, ts, r, RIGHT_COLS, i % 2 === 0)
    const descCell = ws.getCell(`F${r}`)
    descCell.value = truncate(region.description ?? '', NEIGHBORHOOD_DESC_CHARS)
    descCell.font = { name: ts.fontName, size: ts.sizes.data, italic: true }
    descCell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' }
    ws.mergeCells(`F${r}:K${r}`)
    ws.getRow(r).height = BOTTOM_ROW_H
    r++
  })
  return r + 1
}

// OVERVIEW stacks two INDEPENDENT columns of content (the A:D panel and the F:K cards)
// that share row numbers, so styleFactory's row-level helpers can't be used here — they
// walk every populated cell in the row and would restyle the neighbouring column's
// content (e.g. a section header on the left bolding a budget data row on the right).
// These apply the same styling to an explicit column range only.
const LEFT_COLS = ['A', 'B', 'C', 'D'] as const
const RIGHT_COLS = ['F', 'G', 'H', 'I', 'J', 'K'] as const

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
// Summary table now shares rows 14–21 with the left column's CURRENCY / TOTAL BUDGET
// card and the currency-exchange widget, so the row-level helpers would restyle those.
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

// Quick-nav: one internal HYPERLINK per enabled sheet, merged across A:D. Only enabled
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
  ws.getCell(`A${startRow}`).value = 'JUMP TO'
  ws.mergeCells(`A${startRow}:D${startRow}`)
  ws.getRow(startRow).height = 22
  styleSectionHeaderCells(ws, ts, startRow, LEFT_COLS)
  let r = startRow + 1
  links.filter(([on]) => on).forEach(([, sheet, label, id], i) => {
    styleDataRowCells(ws, ts, r, LEFT_COLS, i % 2 === 0)
    const cell = ws.getCell(`A${r}`)
    // Icon comes from the shared SHEET_ICONS map (same emoji the Step 3 toggle grid
    // shows). It renders in its own colors and ignores the font color below —
    // only the label text picks up the theme primary.
    const text = `${SHEET_ICONS[id]}  ${label}`
    cell.value = { formula: `HYPERLINK("#'${sheet}'!A1","${text}")`, result: text }
    cell.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.primary } }
    ws.mergeCells(`A${r}:D${r}`)
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

  // The currency-exchange widget occupies A17:D22 when it's built (skipped when source
  // and target currency match). Row 22 is otherwise just a gutter above the row-23
  // header band, so it only needs full height when the widget's last row lands there.
  const hasCurrencyWidget = !!(exchangeRate && exchangeRate.from !== exchangeRate.to)

  // ── Row heights ─────────────────────────────────────────────────────────────
  // OVERVIEW stacks two INDEPENDENT columns of content (the A:D card stack and the F:K
  // cards) that nonetheless SHARE row numbers, so every height here is the resolution of
  // both sides' needs — set them in one place rather than letting each block set its own,
  // or one side silently squashes the other.
  //
  // The left column is a stack of label/value cards: a short label row, a taller value
  // row, a thin gutter, repeat. The right column is the cover photo, the Budget Summary
  // table, the spend chart and the neighborhood guide. Section headers are deliberately
  // LEVEL across the E gutter — row 13 (BUDGET SUMMARY), row 23 (TRIP READINESS |
  // SPENT VS PLANNED), row 37 (JUMP TO | WHERE TO BASE YOURSELF) — which is what gives
  // the sheet a horizontal rhythm instead of two independently drifting columns.
  //
  // RING_ROW_H is load-bearing: it sets the readiness doughnut's frame height (rows
  // 24–33 = 180pt), and the centered "% ready" overlay in buildReadinessHelpers is
  // merged across the row pair straddling that frame's mid-line. Changing it means
  // re-deriving that merge.
  const RING_ROW_H = 18
  const rowHeights: Record<number, number> = {
    1: 20, // hero panel top padding | cover photo
    2: 20,
    3: 20, // destination headline (A2:D3)
    4: 6, // gutter
    5: 16, // labels: START DATE | END DATE
    6: 24, // values: TripStart | TripEnd
    7: 8, // gutter
    8: 16, // labels: DAYS | PARTY | TRAVELERS
    9: 24, // values
    10: 10, // gutter
    11: 34,
    12: 34, // countdown number (A11:D12)
    13: 22, // "DAYS TO GO" caption | BUDGET SUMMARY header
    14: 20, // hero panel bottom padding | Budget Summary column headers
    15: 20, // labels: CURRENCY | TOTAL BUDGET ‖ budget category rows 15–20
    16: 20, // values: currency | TotalBudget
    17: 20, // currency-exchange widget starts here (A17:D22)
    18: 20,
    19: 20,
    20: 20,
    21: 20, // Budget Summary TOTAL row
    22: hasCurrencyWidget ? 18 : 8, // gutter (or the widget's last row)
    23: 22, // TRIP READINESS | SPENT VS PLANNED headers
    34: 20, // "N of M items complete" caption, just below the ring frame
    35: RING_ROW_H, // last row of the spend chart frame
    36: 8, // gutter
    37: 22, // JUMP TO | WHERE TO BASE YOURSELF headers
  }
  // Rows 24–33: readiness doughnut (A:D) beside the spend chart (F:K, which runs to 35).
  for (let r = 24; r <= 33; r++) rowHeights[r] = RING_ROW_H
  Object.entries(rowHeights).forEach(([r, h]) => {
    ws.getRow(Number(r)).height = h
  })

  // ── Hero band (rows 1–14): dark text panel (A:D) beside the cover photo (F:K) ──
  // Excel can't truly composite editable cell text on top of an opaque floating image
  // (drawings always render above the grid, unlike CSS layering), so the destination/
  // dates/party/countdown live in a panel immediately BESIDE the photo instead. Column E
  // is a deliberately empty, unfilled spacer separating the two so they don't butt
  // together.
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

  // Label + value card inside the hero panel. `span` is the merge range suffix pair.
  const heroLabel = (range: string, text: string): void => {
    const anchor = range.split(':')[0]
    const cell = ws.getCell(anchor)
    cell.value = text
    cell.font = heroLabelFont
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
    if (range.includes(':')) ws.mergeCells(range)
  }

  // Destination headline. Long names step down a size rather than clipping — the panel
  // is ~53 width units and the title size is set for a short name like "NEW YORK".
  ws.mergeCells('A2:D3')
  const headline = ws.getCell('A2')
  const headlineText = truncate(dest.toUpperCase(), 34)
  headline.value = headlineText
  headline.font = heroFont(headlineText.length > 18 ? Math.round(ts.sizes.title * 0.72) : ts.sizes.title)
  headline.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }

  // ── Rows 5–6: date card ─────────────────────────────────────────────────────
  // Two independently-editable date cells, each captioned and merged across half the
  // panel. Both carry the year so a range spanning a year boundary still reads correctly.
  const dateFont = heroFont(Math.round(ts.sizes.title * 0.65))

  heroLabel('A5:B5', 'START DATE')
  heroLabel('C5:D5', 'END DATE')

  // A6 = TripStart (named range anchor) — user-editable
  const tripStartCell = ws.getCell('A6')
  tripStartCell.value = startDate
  tripStartCell.numFmt = 'mmm d, yyyy'
  tripStartCell.protection = { locked: false }
  tripStartCell.font = dateFont
  tripStartCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.mergeCells('A6:B6')

  // C6 = TripEnd (named range anchor) — user-editable
  const tripEndCell = ws.getCell('C6')
  tripEndCell.value = endDate
  tripEndCell.numFmt = 'mmm d, yyyy'
  tripEndCell.protection = { locked: false }
  tripEndCell.font = dateFont
  tripEndCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.mergeCells('C6:D6')

  // ── Rows 8–9: trip-shape card ───────────────────────────────────────────────
  // Duration is derived (inclusive day count, not nights — TripEnd-TripStart+1); party
  // type and traveler count stay editable.
  const lineFont = heroFont(ts.sizes.header)

  heroLabel('A8', 'DAYS')
  heroLabel('B8:C8', 'PARTY')
  heroLabel('D8', 'TRAVELERS')

  const daysCell = ws.getCell('A9')
  daysCell.value = { formula: 'IFERROR(TripEnd-TripStart+1,"")', result: state.duration }
  daysCell.numFmt = '0'
  daysCell.protection = { hidden: true }
  daysCell.font = lineFont
  daysCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }

  const partyCell = ws.getCell('B9')
  partyCell.value = state.partyType.charAt(0).toUpperCase() + state.partyType.slice(1)
  partyCell.protection = { locked: false }
  partyCell.font = lineFont
  partyCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.mergeCells('B9:C9')
  ;(ws as any).dataValidations.add('B9', {
    type: 'list',
    allowBlank: false,
    formulae: ['"Solo,Couple,Family,Group"'],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // D9 = NumAdults (named range anchor) — user-editable
  const travelersCell = ws.getCell('D9')
  travelersCell.value = state.partySize
  travelersCell.numFmt = '0'
  travelersCell.protection = { locked: false }
  travelersCell.font = lineFont
  travelersCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }

  // ── Rows 11–13: countdown badge ─────────────────────────────────────────────
  // The mockup's big circular "292 / DAYS TO GO" marker. Excel can't draw it as a circle
  // over the photo, so it's an accent-filled block spanning the full panel width: the
  // count as one oversized number (A11:D12), the unit as a caption beneath (A13:D13).
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

  const badgeCount = ws.getCell('A11')
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
  ws.mergeCells('A11:D12')

  const badgeUnit = ws.getCell('A13')
  badgeUnit.value = {
    formula: 'IFERROR(IF(TripStart=TODAY(),"TODAY!",IF(TripStart>TODAY(),"DAYS TO GO","DAYS AGO")),"")',
    result: daysUntil === 0 ? 'TODAY!' : daysUntil > 0 ? 'DAYS TO GO' : 'DAYS AGO',
  }
  badgeUnit.protection = { hidden: true }
  badgeUnit.font = { name: ts.fontName, size: ts.sizes.data, bold: true, color: { argb: ts.palette.secondaryText } }
  badgeUnit.alignment = { vertical: 'top', horizontal: 'center' }
  ws.mergeCells('A13:D13')

  // Column widths for the whole A:D stack — the hero band, the row 15–16 card, the
  // currency widget, the readiness ring and quick-nav all share these, so each column is
  // sized for the WIDEST thing that lands in it, at the largest font that lands in it
  // (the hero rows run at title/header sizes, well above the 11pt default these width
  // units assume). Undersizing here is what produced "###" and clipped labels.
  //   A: "DAYS" · day count · quick-nav links · left half of the date/currency cards
  //   B: rest of the date/currency card merges · party type
  //   C: "END DATE" / "TOTAL BUDGET" merge anchors
  //   D: travelers · right half of those merges
  //   E: empty spacer between the panel and the cover photo (never filled)
  ws.getColumn('A').width = 10
  ws.getColumn('B').width = 18
  ws.getColumn('C').width = 16
  ws.getColumn('D').width = 14
  ws.getColumn('E').width = 3

  // ── Rows 15–16: Currency + Total Budget card ────────────────────────────────
  // Same label-over-value shape as the hero cards, but below the dark panel, so it uses
  // the shared light caption/value styles. Merged in halves (A:B | C:D) so the two
  // labels can't clip each other the way they did when all four sat on one row.
  const cardLabel = (range: string, text: string): void => {
    const anchor = range.split(':')[0]
    const cell = ws.getCell(anchor)
    cell.value = text
    styleLabelCell(cell, ts)
    ws.mergeCells(range)
  }
  cardLabel('A15:B15', 'CURRENCY')
  cardLabel('C15:D15', 'TOTAL BUDGET')

  // Currency — informational label (chosen in the wizard), a dropdown so the user can
  // pick a different reference currency. Options are written to a hidden helper column
  // (R) because the full "CODE (symbol)" list exceeds Excel's 255-char inline list-
  // formula limit. Purely a label — does not affect any numFmt elsewhere (those are
  // static, baked in at generation time).
  const currValue = ws.getCell('A16')
  currValue.value = `${state.currency} (${currSym})`
  currValue.protection = { locked: false }
  styleValueCell(currValue, ts)
  ws.mergeCells('A16:B16')

  CURRENCIES.forEach((c, i) => {
    ws.getCell(`R${i + 1}`).value = `${c.code} (${c.symbol})`
  })
  ;(ws as any).dataValidations.add('A16', {
    type: 'list',
    allowBlank: true,
    formulae: [`$R$1:$R${CURRENCIES.length}`],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // C16 = TotalBudget (named range anchor) — live sum of the Estimated Budget column,
  // so it reflects any edits the user types into those cells. Cached result keeps the
  // cell populated before the first recalc.
  const budgetValue = ws.getCell('C16')
  budgetValue.value = { formula: 'SUM(G15:G20)', result: totalBudget }
  budgetValue.numFmt = ts.numFmtCurrency
  budgetValue.protection = { hidden: true }
  budgetValue.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  budgetValue.alignment = { vertical: 'middle', horizontal: 'left' }
  budgetValue.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.mediumBg } } as ExcelJS.Fill
  ws.mergeCells('C16:D16')

  // ── Budget Summary Table (rows 13–21) ───────────────────────────────────────
  // Header sits at row 13 with an 8pt gutter (row 12) between it and the cover photo,
  // which now stops at row 11 instead of butting straight into this band.
  ws.getCell('F13').value = 'BUDGET SUMMARY'
  ws.mergeCells('F13:K13')
  styleSectionHeaderCells(ws, ts, 13, RIGHT_COLS)

  ws.getCell('F14').value = 'Category'
  ws.getCell('G14').value = 'Estimated Budget'
  ws.getCell('H14').value = 'Actual Spent'
  ws.getCell('I14').value = 'Remaining'
  ws.getCell('J14').value = '% Used'
  ws.getCell('K14').value = 'Status'
  styleColumnHeaderCells(ws, ts, 14, RIGHT_COLS)

  let rowIdx = 15
  categoryBudgetAmounts(state).forEach(({ key, amount: budgetAmt }, i) => {
    ws.getCell(`F${rowIdx}`).value = key

    // Estimated Budget (col G) — a plain, editable starting value derived from the
    // wizard's budget selections. It is NOT a formula: it doesn't change when the trip
    // dates change. Users simply type their own figure over it.
    const gCell = ws.getCell(`G${rowIdx}`)
    gCell.value = budgetAmt
    gCell.numFmt = ts.numFmtCurrency
    gCell.protection = { locked: false }

    // Actual spent = per-category sheet totals + matching OTHER EXPENSES entries.
    // Cached results throughout this table = the empty-tracker evaluation, so cells
    // render immediately in viewers that don't recalc on open (see countdown note).
    const hCell = ws.getCell(`H${rowIdx}`)
    hCell.value = { formula: buildActualSpentFormula(key, state), result: 0 }
    hCell.numFmt = ts.numFmtCurrency
    hCell.protection = { hidden: true }

    // Remaining
    const iCell = ws.getCell(`I${rowIdx}`)
    iCell.value = { formula: `IFERROR(G${rowIdx}-H${rowIdx},G${rowIdx})`, result: budgetAmt }
    iCell.numFmt = ts.numFmtCurrency
    iCell.protection = { hidden: true }

    // % Used
    const jCell = ws.getCell(`J${rowIdx}`)
    jCell.value = { formula: `IFERROR(H${rowIdx}/G${rowIdx},0)`, result: 0 }
    jCell.numFmt = ts.numFmtPercent
    jCell.protection = { hidden: true }

    // Status
    const kCell = ws.getCell(`K${rowIdx}`)
    kCell.value = {
      formula: `IFERROR(IF(H${rowIdx}>G${rowIdx},"Over budget",IF(H${rowIdx}/G${rowIdx}>0.8,"Near limit","On track")),"On track")`,
      result: 'On track',
    }
    kCell.protection = { hidden: true }

    // Col M: pie/donut chart helper — actual spend once logged, else budget estimate so
    // the chart is never empty when the user hasn't entered spending yet.
    const mCell = ws.getCell(`M${rowIdx}`)
    mCell.value = { formula: `IF(H${rowIdx}>0,H${rowIdx},G${rowIdx})`, result: budgetAmt }
    mCell.numFmt = ts.numFmtCurrency
    mCell.protection = { hidden: true }

    // Col N: bar chart helper — MIN(actual, budget), the colored "actual spent" segment
    // (base of the stacked bar).
    const nCell = ws.getCell(`N${rowIdx}`)
    nCell.value = { formula: `MIN(H${rowIdx},G${rowIdx})`, result: 0 }
    nCell.numFmt = ts.numFmtCurrency
    nCell.protection = { hidden: true }

    // Col O: bar chart helper — MAX(budget - actual, 0), the light "remaining budget"
    // segment stacked after col N so one bar = full budget (matches the Excel look in
    // Google Sheets, which ignores overlap on clustered charts).
    const oCell = ws.getCell(`O${rowIdx}`)
    oCell.value = { formula: `MAX(G${rowIdx}-H${rowIdx},0)`, result: budgetAmt }
    oCell.numFmt = ts.numFmtCurrency
    oCell.protection = { hidden: true }

    // Col P: bar chart helper — MAX(actual - budget, 0), the red "over budget" overage
    // segment stacked after col O. Zero (invisible) when on/under budget; when spending
    // exceeds the estimate it extends the bar past the budget length in red, mirroring
    // the "turns red when over budget" preview shown in the wizard's chart picker.
    const pCell = ws.getCell(`P${rowIdx}`)
    pCell.value = { formula: `MAX(H${rowIdx}-G${rowIdx},0)`, result: 0 }
    pCell.numFmt = ts.numFmtCurrency
    pCell.protection = { hidden: true }

    styleDataRowCells(ws, ts, rowIdx, RIGHT_COLS, i % 2 === 0)
    rowIdx++
  })

  // Totals row (row 21)
  ws.getCell(`F${rowIdx}`).value = 'TOTAL'
  const gTot = ws.getCell(`G${rowIdx}`)
  gTot.value = { formula: `SUM(G15:G${rowIdx - 1})`, result: totalBudget }
  gTot.numFmt = ts.numFmtCurrency
  gTot.protection = { hidden: true }
  const hTot = ws.getCell(`H${rowIdx}`)
  hTot.value = { formula: `SUM(H15:H${rowIdx - 1})`, result: 0 }
  hTot.numFmt = ts.numFmtCurrency
  hTot.protection = { hidden: true }
  const iTot = ws.getCell(`I${rowIdx}`)
  iTot.value = { formula: `SUM(I15:I${rowIdx - 1})`, result: totalBudget }
  iTot.numFmt = ts.numFmtCurrency
  iTot.protection = { hidden: true }
  styleTotalRowCells(ws, ts, rowIdx, RIGHT_COLS)

  // ── Currency Exchange Widget (cols A–D, rows 17–22) ─────────────────────────
  // Sits left of the budget table (F13:K21) — same row band, different columns; the
  // spec's "negative space beside the table" is exactly the six rows it needs.
  // Built after the budget table so its fills win on any shared row.
  // Skip when source/target match — a same-currency converter (rate 1.0) is meaningless.
  if (hasCurrencyWidget && exchangeRate) {
    buildCurrencyWidget(ws, ts, exchangeRate, 17)
  }

  // ── Spend chart (rows 23–35) ────────────────────────────────────────────────
  // In-cell title so the injected chart itself carries none — a chart-space title eats
  // plot height and doesn't line up with anything else on the sheet, whereas this band
  // sits level with TRIP READINESS across the E gutter. Chart anchored F24:K35 (index.ts).
  ws.getCell('F23').value = 'SPENT VS PLANNED'
  ws.mergeCells('F23:K23')
  styleSectionHeaderCells(ws, ts, 23, RIGHT_COLS)

  // Trip-readiness ring: header row 23, chart injected in index.ts anchored A24:D33.
  buildReadinessHelpers(ws, ts, state)

  // Quick-nav HYPERLINK list, lower-left — header level with the neighborhood guide's.
  buildQuickNav(ws, state, ts, 37)

  // Neighborhood guide, right column below the spend chart.
  if (state.useRecommendations && recommendations?.regions?.length) {
    buildNeighborhoodGuide(ws, ts, recommendations.regions, 37)
  }

  // ── Visual Budget Breakdown ─────────────────────────────────────────────────
  // The breakdown is a native chart object injected after export for every style
  // (bar / pie / donut) — see chartInjection.ts. Nothing to render in-cell here.

  // ── Column widths ───────────────────────────────────────────────────────────
  // A–E are set with the hero band above (they're sized for its larger fonts) — do not
  // re-set them here or that sizing is silently discarded.
  ws.getColumn('F').width = 20
  ws.getColumn('G').width = 20
  ws.getColumn('H').width = 16
  ws.getColumn('I').width = 16
  ws.getColumn('J').width = 11
  ws.getColumn('K').width = 16

  // Hide the chart's helper columns (chart plots hidden cells — see plotVisOnly in chartInjection)
  ws.getColumn('M').hidden = true
  ws.getColumn('N').hidden = true
  ws.getColumn('O').hidden = true
  ws.getColumn('P').hidden = true

  // Hide the currency-dropdown options helper column (R)
  ws.getColumn('R').hidden = true

  // ── Cover photo (optional, anchored F1:K11) ─────────────────────────────────
  // Stops at row 11, three rows short of the text panel's bottom, so the BUDGET SUMMARY
  // header at row 13 gets a gutter instead of butting into the photo. PictureUploader's
  // crop target (TARGET_W/TARGET_H) must match THIS range's aspect ratio — F:K is 99
  // width units (~693px) and rows 1–11 total 198pt (~264px), i.e. ~2.63:1 — or the
  // embedded photo stretches. Changing any row height in 1–11 changes that ratio.
  if (state.overviewImage) {
    addDataUrlImage(wb, ws, state.overviewImage.dataUrl, 'jpeg', 'F1:K11')
  }

  // OVERVIEW is a one-screen dashboard, not a scrollable table — no freeze pane (every
  // other sheet builder freezes its header rows; this is a deliberate OVERVIEW-only
  // exception) and no gridlines, so the hero band + cards read as a dashboard.
  ws.views = [{ showGridLines: false }]
}
