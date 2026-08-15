import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import type { ExchangeRate } from '../../recommendations/types'
import {
  styleTitleRow,
  styleSectionHeader,
  styleColumnHeader,
  styleDataRow,
  styleTotalRow,
  styleLabelCell,
  styleValueCell,
} from '../styleFactory'
import { computeTotalBudget } from '../../utils'
import { THEMES } from '../../../types/theme'
import { CURRENCIES, getCurrencySymbol, getCurrencyNumFmt } from '../../../data/currencies'
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

// Per-category estimated budget amounts, in Budget Summary (F18:F23) row order.
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

export function buildOverviewSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  exchangeRate?: ExchangeRate
): void {
  const ws = wb.addWorksheet('OVERVIEW')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  const dest = state.destination || 'My Trip'
  const startDate = state.startDate ? new Date(state.startDate) : new Date()
  const endDate = addDays(startDate, state.duration - 1)
  const totalBudget = computeTotalBudget(state.budgets, state.duration)
  const currSym = getCurrencySymbol(state.currency)

  // ── Row 1: Title ────────────────────────────────────────────────────────────
  const titleRow = ws.getRow(1)
  const titleCell = ws.getCell('A1')
  titleCell.value = `✈  ${dest.toUpperCase()}  —  TRAVEL PLANNER`
  ws.mergeCells('A1:N1')
  styleTitleRow(titleRow, ts)

  // ── Row 2: spacer ───────────────────────────────────────────────────────────
  ws.getRow(2).height = 6

  // ── Key-value info block (rows 3–14, values in col D for named ranges) ──────
  // All rows 3–14 get an explicit height of 20pt (240pt total ≈ 320px @ 96dpi) — both
  // for visual consistency of the key-value block and so the F3:K14 cover-photo anchor,
  // which spans these same rows, renders at a predictable pixel height.
  for (let r = 3; r <= 14; r++) ws.getRow(r).height = 20

  const kvLabels: [number, string][] = [
    [3, 'TRIP START DATE'],
    [4, 'TRIP END DATE'],
    [5, 'DURATION'],
    [6, 'DAYS UNTIL TRIP'],
    [7, ''],
    [8, 'PARTY TYPE'],
    [9, 'TRAVELERS'],
    [10, ''],
    [11, 'CURRENCY'],
    [12, ''],
    [13, ''],
    [14, 'TOTAL BUDGET'],
  ]

  kvLabels.forEach(([rowNum, label]) => {
    if (!label) return
    const labelCell = ws.getCell(`A${rowNum}`)
    labelCell.value = label
    ws.mergeCells(`A${rowNum}:C${rowNum}`)
    styleLabelCell(labelCell, ts)
  })

  // D3 = TripStart (named range anchor) — user-editable
  const d3 = ws.getCell('D3')
  d3.value = startDate
  d3.numFmt = ts.numFmtDate
  d3.protection = { locked: false }
  styleValueCell(d3, ts)

  // D4 = TripEnd (named range anchor) — user-editable
  const d4 = ws.getCell('D4')
  d4.value = endDate
  d4.numFmt = ts.numFmtDate
  d4.protection = { locked: false }
  styleValueCell(d4, ts)

  const d5 = ws.getCell('D5')
  d5.value = {
    formula: 'IFERROR(INT(D4-D3+1)&" days","")',
    result: `${state.duration} days`,
  }
  d5.protection = { hidden: true }
  styleValueCell(d5, ts)

  // Countdown — precompute a cached result so the cell renders immediately on open.
  // (Excel/Sheets won't trigger an initial recalc on this freshly written file, so a
  // formula with no cached result shows blank until the user edits something.)
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  const startDay = new Date(startDate.getTime())
  startDay.setHours(0, 0, 0, 0)
  const daysUntil = Math.round((startDay.getTime() - today.getTime()) / 86400000)
  const daysResult =
    daysUntil === 0
      ? 'Today!'
      : daysUntil > 0
        ? `${daysUntil} days to go`
        : `${Math.abs(daysUntil)} days ago`
  const d6 = ws.getCell('D6')
  d6.value = {
    formula:
      'IFERROR(IF(TripStart=TODAY(),"Today!",IF(TripStart>TODAY(),INT(TripStart-TODAY())&" days to go",INT(TODAY()-TripStart)&" days ago")),"—")',
    result: daysResult,
  }
  d6.protection = { hidden: true }
  styleValueCell(d6, ts)

  const d8 = ws.getCell('D8')
  d8.value = state.partyType.charAt(0).toUpperCase() + state.partyType.slice(1)
  styleValueCell(d8, ts)

  // D9 = NumAdults (named range anchor) — user-editable
  const d9 = ws.getCell('D9')
  d9.value = state.partySize
  d9.protection = { locked: false }
  styleValueCell(d9, ts)

  // Currency — informational label (chosen in the wizard), now a dropdown so the
  // user can pick a different reference currency. Options are written to a hidden
  // helper column (R) because the full "CODE (symbol)" list exceeds Excel's
  // 255-char inline list-formula limit. Purely a label — does not affect any
  // numFmt elsewhere (those are static, baked in at generation time).
  const d11 = ws.getCell('D11')
  d11.value = `${state.currency} (${currSym})`
  d11.protection = { locked: false }
  styleValueCell(d11, ts)

  CURRENCIES.forEach((c, i) => {
    ws.getCell(`R${i + 1}`).value = `${c.code} (${c.symbol})`
  })
  ;(ws as any).dataValidations.add('D11', {
    type: 'list',
    allowBlank: true,
    formulae: [`$R$1:$R${CURRENCIES.length}`],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // D14 = TotalBudget (named range anchor) — live sum of the Estimated Budget column,
  // so it reflects any edits the user types into those cells. Cached result keeps the
  // cell populated before the first recalc.
  const d14 = ws.getCell('D14')
  d14.value = { formula: 'SUM(G18:G23)', result: totalBudget }
  d14.numFmt = ts.numFmtCurrency
  d14.protection = { hidden: true }
  d14.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  d14.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.mediumBg } } as ExcelJS.Fill

  // ── Row 15: spacer ──────────────────────────────────────────────────────────
  ws.getRow(15).height = 8

  // ── Budget Summary Table (rows 16–24) ───────────────────────────────────────
  const summaryHeader = ws.getRow(16)
  ws.getCell('F16').value = 'BUDGET SUMMARY'
  ws.mergeCells('F16:K16')
  styleSectionHeader(summaryHeader, ts)

  const colHeaderRow = ws.getRow(17)
  ws.getCell('F17').value = 'Category'
  ws.getCell('G17').value = 'Estimated Budget'
  ws.getCell('H17').value = 'Actual Spent'
  ws.getCell('I17').value = 'Remaining'
  ws.getCell('J17').value = '% Used'
  ws.getCell('K17').value = 'Status'
  styleColumnHeader(colHeaderRow, ts)

  let rowIdx = 18
  categoryBudgetAmounts(state).forEach(({ key, amount: budgetAmt }, i) => {
    const bRow = ws.getRow(rowIdx)

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

    styleDataRow(bRow, ts, i % 2 === 0)
    rowIdx++
  })

  // Totals row
  const totRow = ws.getRow(rowIdx)
  ws.getCell(`F${rowIdx}`).value = 'TOTAL'
  const gTot = ws.getCell(`G${rowIdx}`)
  gTot.value = { formula: `SUM(G18:G${rowIdx - 1})`, result: totalBudget }
  gTot.numFmt = ts.numFmtCurrency
  gTot.protection = { hidden: true }
  const hTot = ws.getCell(`H${rowIdx}`)
  hTot.value = { formula: `SUM(H18:H${rowIdx - 1})`, result: 0 }
  hTot.numFmt = ts.numFmtCurrency
  hTot.protection = { hidden: true }
  const iTot = ws.getCell(`I${rowIdx}`)
  iTot.value = { formula: `SUM(I18:I${rowIdx - 1})`, result: totalBudget }
  iTot.numFmt = ts.numFmtCurrency
  iTot.protection = { hidden: true }
  styleTotalRow(totRow, ts)
  rowIdx++

  // ── Row spacer ──────────────────────────────────────────────────────────────
  ws.getRow(rowIdx).height = 8
  rowIdx++

  // ── Currency Exchange Widget (cols A–D, rows 17–22) ─────────────────────────
  // Sits left of the budget table (F16:K24) — same row band, different columns.
  // Built after the budget table so widget fills override the row-level style passes.
  // Skip when source/target match — a same-currency converter (rate 1.0) is meaningless.
  if (exchangeRate && exchangeRate.from !== exchangeRate.to) {
    buildCurrencyWidget(ws, ts, exchangeRate, 17)
  }

  // ── Visual Budget Breakdown ─────────────────────────────────────────────────
  // The breakdown is a native chart object injected after export for every style
  // (bar / pie / donut) — see chartInjection.ts. Nothing to render in-cell here.

  // ── Column widths ───────────────────────────────────────────────────────────
  ws.getColumn('A').width = 20
  ws.getColumn('B').width = 3
  ws.getColumn('C').width = 3
  ws.getColumn('D').width = 28
  ws.getColumn('E').width = 4
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

  // ── Cover photo (optional, anchored F3:K14) ─────────────────────────────────
  if (state.overviewImage) {
    const base64 = state.overviewImage.dataUrl.slice(state.overviewImage.dataUrl.indexOf(',') + 1)
    // exceljs's `Image.buffer` type is a module-local `interface Buffer extends
    // ArrayBuffer {}`, structurally distinct from Node/polyfilled `Buffer`
    // (Uint8Array) — cast through ArrayBuffer, which it's structurally identical to.
    const imageId = wb.addImage({ buffer: Buffer.from(base64, 'base64') as unknown as ArrayBuffer, extension: 'jpeg' })
    ws.addImage(imageId, 'F3:K14')
  }

  // Freeze top 2 rows
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2, activeCell: 'A3' }]
}
