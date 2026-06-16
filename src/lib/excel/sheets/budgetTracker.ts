import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataRow, styleDataCell, styleTotalRow } from '../styleFactory'
import { THEMES } from '../../../types/theme'
import { getCurrencySymbol } from '../../../data/currencies'

export function buildBudgetTrackerSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles
): void {
  const ws = wb.addWorksheet('OTHER EXPENSES')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  const sym = getCurrencySymbol(state.currency)

  // Row 1: Title
  ws.getCell('A1').value = `OTHER EXPENSES  (${sym} ${state.currency})`
  ws.mergeCells('A1:E1')
  styleTitleRow(ws.getRow(1), ts)

  // Row 2: Sub-heading
  ws.getCell('A2').value = 'For additional / incidental expenses (shopping, tips, fees, etc.) not already tracked on the dedicated tabs. Category dropdown in column B feeds the OVERVIEW budget summary.'
  ws.mergeCells('A2:E2')
  ws.getCell('A2').font = { name: ts.fontName, size: ts.sizes.data, italic: true }
  ws.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.lightBg } } as ExcelJS.Fill
  ws.getRow(2).height = 18

  // Row 3: Totals summary — sum the data rows only (8–57), excluding the bottom TOTAL row.
  ws.getCell('A3').value = 'TOTAL SPENT'
  ws.mergeCells('A3:C3')
  ws.getCell('A3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.getCell('D3').value = { formula: 'IFERROR(SUM(D8:D57),0)' }
  ws.getCell('D3').numFmt = ts.numFmtCurrency
  ws.getCell('D3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.getRow(3).height = 22

  // Row 4: spacer
  ws.getRow(4).height = 6

  // Row 5: Section header
  ws.getCell('A5').value = 'EXPENSE LOG'
  ws.mergeCells('A5:E5')
  styleSectionHeader(ws.getRow(5), ts)

  // Row 6: spacer
  ws.getRow(6).height = 4

  // Row 7: Column headers (Date | Category | Description | Cost | Notes)
  ws.getCell('A7').value = 'Date'
  ws.getCell('B7').value = 'Category'
  ws.getCell('C7').value = 'Description'
  ws.getCell('D7').value = 'Cost'
  ws.getCell('E7').value = 'Notes'
  styleColumnHeader(ws.getRow(7), ts)

  // Data validation for category column (B) — bounded to the 50 pre-formatted data
  // rows so Google Sheets doesn't render 450+ spurious dropdown carets below the table.
  ;(ws as any).dataValidations.add('B8:B57', {
    type: 'list',
    allowBlank: true,
    formulae: ['"Transportation,Accommodation,Food,Activities,Shopping,Miscellaneous"'],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // Pre-fill 50 data rows (8–57)
  for (let i = 0; i < 50; i++) {
    const rowNum = 8 + i
    const row = ws.getRow(rowNum)

    // Cost format
    ws.getCell(`D${rowNum}`).numFmt = ts.numFmtCurrency

    // Date format
    ws.getCell(`A${rowNum}`).numFmt = ts.numFmtDate

    styleDataRow(row, ts, i % 2 === 0)
    // Cost (D) is user-entered, so it's empty in the template — stripe it explicitly
    // so the column matches the populated columns (styleDataRow skips empty cells).
    styleDataCell(ws.getCell(`D${rowNum}`), ts, i % 2 === 0)
  }

  // Bottom TOTAL row — a regular total of the data rows, like the other tabs.
  const totRow = 58
  ws.getCell(`A${totRow}`).value = 'TOTAL'
  const dTot = ws.getCell(`D${totRow}`)
  dTot.value = { formula: 'IFERROR(SUM(D8:D57),0)' }
  dTot.numFmt = ts.numFmtCurrency
  styleTotalRow(ws.getRow(totRow), ts)

  // Column widths — totals ~175 so the merged subheading on row 2 (≈173 chars) fits
  // on one line without being clipped (merged cells don't overflow into empty cells).
  ws.getColumn('A').width = 14
  ws.getColumn('B').width = 22
  ws.getColumn('C').width = 58
  ws.getColumn('D').width = 16
  ws.getColumn('E').width = 65

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 7, activeCell: 'A8' }]
}
