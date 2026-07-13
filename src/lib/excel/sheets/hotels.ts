import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import type { RegionRec } from '../../recommendations/types'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataCell, styleTotalRow } from '../styleFactory'
import { THEMES } from '../../../types/theme'
import { buildRecommendationInsert } from './recommendationInsert'

export function buildHotelsSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  regions?: RegionRec[]
): void {
  const ws = wb.addWorksheet('ACCOMMODATION')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  // Title
  ws.getCell('A1').value = 'ACCOMMODATION'
  ws.mergeCells('A1:G1')
  styleTitleRow(ws.getRow(1), ts)

  ws.getCell('A2').value = 'Track all accommodation costs. Use Cost Type to log nightly rates, taxes, fees, parking and other charges as separate rows.'
  ws.mergeCells('A2:G2')
  styleSectionHeader(ws.getRow(2), ts)

  // Summary row — sums the Cost column (E).
  ws.getCell('A3').value = 'TOTAL ACCOMMODATION COST'
  ws.getCell('A3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.mergeCells('A3:D3')
  ws.getCell('E3').value = { formula: 'IFERROR(SUM(E6:E25),0)', result: 0 }
  ws.getCell('E3').numFmt = ts.numFmtCurrency
  ws.getCell('E3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.getRow(3).height = 22

  // Column headers
  ws.getCell('A5').value = 'Date'
  ws.getCell('B5').value = 'City'
  ws.getCell('C5').value = 'Property Name'
  ws.getCell('D5').value = 'Cost Type'
  ws.getCell('E5').value = 'Cost'
  ws.getCell('F5').value = 'Confirmation #'
  ws.getCell('G5').value = 'Notes'
  styleColumnHeader(ws.getRow(5), ts)

  // Date validation (col A) — triggers Google Sheets native date picker on double-click.
  ;(ws as any).dataValidations.add('A6:A25', {
    type: 'date',
    operator: 'between',
    formulae: [new Date(2000, 0, 1), new Date(2100, 0, 1)],
    allowBlank: true,
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // Cost Type dropdown (col D)
  ;(ws as any).dataValidations.add('D6:D25', {
    type: 'list',
    allowBlank: true,
    formulae: ['"Full Cost,Nightly Rate,Taxes,Fees,Parking,Other"'],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // Pre-fill rows
  for (let i = 0; i < 20; i++) {
    const rowNum = 6 + i
    const row = ws.getRow(rowNum)

    const dateCell = ws.getCell(`A${rowNum}`)
    dateCell.numFmt = ts.numFmtDate
    ws.getCell(`E${rowNum}`).numFmt = ts.numFmtCurrency

    // Per the request, only the Cost column is striped — give the rest of the row a
    // plain (white) background, then override the Cost cell with the alternating stripe.
    for (const col of ['A', 'B', 'C', 'D', 'E', 'F', 'G']) {
      styleDataCell(ws.getCell(`${col}${rowNum}`), ts, false)
    }
    styleDataCell(ws.getCell(`E${rowNum}`), ts, i % 2 === 0)
    row.height = 20
  }

  // Totals row
  const totRow = 26
  ws.getCell(`A${totRow}`).value = 'TOTAL'
  const eTot = ws.getCell(`E${totRow}`)
  eTot.value = { formula: 'IFERROR(SUM(E6:E25),0)', result: 0 }
  eTot.numFmt = ts.numFmtCurrency
  styleTotalRow(ws.getRow(totRow), ts)

  ws.getColumn('A').width = 16
  ws.getColumn('B').width = 18
  ws.getColumn('C').width = 28
  ws.getColumn('D').width = 16
  ws.getColumn('E').width = 14
  ws.getColumn('F').width = 20
  ws.getColumn('G').width = 39

  // Informational AI guide below the tracker (TOTAL row is 26; leave a spacer)
  buildRecommendationInsert(ws, totRow + 2, ts, regions, 'hotels')

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 5, activeCell: 'A6' }]
}
