import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataRow, styleDataCell, styleTotalRow } from '../styleFactory'
import { THEMES } from '../../../types/theme'

const TRANSPORTATION_MODES = 'Air,Train,Bus,Car / Rental,Ferry,Taxi / Rideshare,Other'
const CLASS_OPTIONS = 'Economy,Premium Economy,Business,First Class,Standard,Sleeper,N/A'

const DATA_ROWS = 20

export function buildFlightsSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles
): void {
  const ws = wb.addWorksheet('TRANSPORTATION')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  ws.getCell('A1').value = 'TRANSPORTATION'
  ws.mergeCells('A1:K1')
  styleTitleRow(ws.getRow(1), ts)

  ws.getCell('A2').value = 'Log all flights, trains, buses, and major transport. Pick a Mode for each leg. Confirmation numbers help at check-in.'
  ws.mergeCells('A2:K2')
  styleSectionHeader(ws.getRow(2), ts)

  // Row 3: total
  ws.getCell('A3').value = 'TOTAL TRAVEL COST'
  ws.getCell('A3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.mergeCells('A3:E3')
  const lastDataRow = 5 + DATA_ROWS
  ws.getCell('H3').value = { formula: `IFERROR(SUM(H6:H${lastDataRow}),0)`, result: 0 }
  ws.getCell('H3').numFmt = ts.numFmtCurrency
  ws.getCell('H3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.getCell('H3').protection = { hidden: true }
  ws.getRow(3).height = 22

  ws.getRow(4).height = 4

  // Column headers
  ws.getCell('A5').value = 'Mode'
  ws.getCell('B5').value = 'Carrier / Operator'
  ws.getCell('C5').value = 'Route'
  ws.getCell('D5').value = 'Departure Date'
  ws.getCell('E5').value = 'Departure Time'
  ws.getCell('F5').value = 'Arrival Date'
  ws.getCell('G5').value = 'Arrival Time'
  ws.getCell('H5').value = 'Cost'
  ws.getCell('I5').value = 'Confirmation #'
  ws.getCell('J5').value = 'Class / Tier'
  ws.getCell('K5').value = 'Notes'
  styleColumnHeader(ws.getRow(5), ts)

  // Transport mode dropdown — bounded to data rows only (no extra rows with dropdowns)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ;(ws as any).dataValidations.add(`A6:A${lastDataRow}`, {
    type: 'list',
    allowBlank: true,
    formulae: [`"${TRANSPORTATION_MODES}"`],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // Class dropdown — bounded to data rows only
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ;(ws as any).dataValidations.add(`J6:J${lastDataRow}`, {
    type: 'list',
    allowBlank: true,
    formulae: [`"${CLASS_OPTIONS}"`],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // Date validation on Departure Date and Arrival Date — triggers the native Google Sheets
  // date-picker calendar on double-click. A date numFmt alone is not enough; a `date`
  // validation rule is required. The wide 2000–2100 range is effectively unbounded for
  // travel use and does not block any realistic entry.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ;(ws as any).dataValidations.add(`D6:D${lastDataRow}`, {
    type: 'date',
    operator: 'between',
    formulae: [new Date(2000, 0, 1), new Date(2100, 0, 1)],
    allowBlank: true,
    showErrorMessage: false,
  } as ExcelJS.DataValidation)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ;(ws as any).dataValidations.add(`F6:F${lastDataRow}`, {
    type: 'date',
    operator: 'between',
    formulae: [new Date(2000, 0, 1), new Date(2100, 0, 1)],
    allowBlank: true,
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  for (let i = 0; i < DATA_ROWS; i++) {
    const rowNum = 6 + i
    const row = ws.getRow(rowNum)

    ws.getCell(`D${rowNum}`).numFmt = ts.numFmtDate
    ws.getCell(`F${rowNum}`).numFmt = ts.numFmtDate
    ws.getCell(`E${rowNum}`).numFmt = 'h:mm AM/PM'
    ws.getCell(`G${rowNum}`).numFmt = 'h:mm AM/PM'
    ws.getCell(`H${rowNum}`).numFmt = ts.numFmtCurrency

    // Every column in this row is user-entered — unlock the whole row.
    for (const col of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) {
      ws.getCell(`${col}${rowNum}`).protection = { locked: false }
    }

    styleDataRow(row, ts, i % 2 === 0)
    // Cost (H) is user-entered, so it's empty in the template — stripe it explicitly
    // so the column matches the populated columns (styleDataRow skips empty cells).
    styleDataCell(ws.getCell(`H${rowNum}`), ts, i % 2 === 0)
    row.height = 20
  }

  const totRow = lastDataRow + 1
  ws.getCell(`A${totRow}`).value = 'TOTAL'
  const hTot = ws.getCell(`H${totRow}`)
  hTot.value = { formula: `IFERROR(SUM(H6:H${lastDataRow}),0)`, result: 0 }
  hTot.numFmt = ts.numFmtCurrency
  hTot.protection = { hidden: true }
  styleTotalRow(ws.getRow(totRow), ts)

  ws.getColumn('A').width = 16
  ws.getColumn('B').width = 22
  ws.getColumn('C').width = 28
  ws.getColumn('D').width = 18
  ws.getColumn('E').width = 16
  ws.getColumn('F').width = 18
  ws.getColumn('G').width = 16
  ws.getColumn('H').width = 12
  ws.getColumn('I').width = 18
  ws.getColumn('J').width = 18
  ws.getColumn('K').width = 30

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 5, activeCell: 'A6' }]
}
