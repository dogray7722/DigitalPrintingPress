import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import type { RegionRec } from '../../recommendations/types'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataRow, styleTotalRow } from '../styleFactory'
import { THEMES } from '../../../types/theme'
import { buildRecommendationInsert } from './recommendationInsert'

export function buildExcursionsSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  regions?: RegionRec[]
): void {
  const ws = wb.addWorksheet('EXCURSIONS')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  ws.getCell('A1').value = 'EXCURSIONS & ACTIVITIES'
  ws.mergeCells('A1:J1')
  styleTitleRow(ws.getRow(1), ts)

  ws.getCell('A2').value = 'Cost = Unit Cost (Per Group) or Unit Cost × # Travelers (Per Person). Select Cost Type per row.'
  ws.mergeCells('A2:J2')
  styleSectionHeader(ws.getRow(2), ts)

  ws.getCell('A3').value = 'TOTAL ACTIVITIES COST'
  ws.getCell('A3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.mergeCells('A3:D3')
  ws.getCell('H3').value = { formula: 'IFERROR(SUM(H6:H35),0)', result: 0 }
  ws.getCell('H3').numFmt = ts.numFmtCurrency
  ws.getCell('H3').font = { name: ts.fontName, size: ts.sizes.header, bold: true }
  ws.getRow(3).height = 22

  // Column headers — A:Date B:Excursion C:Duration D:Type E:CostType F:Cost G:#Travelers H:Total I:Participants J:Notes
  ws.getCell('A5').value = 'Date'
  ws.getCell('B5').value = 'Excursion / Activity'
  ws.getCell('C5').value = 'Duration'
  ws.getCell('D5').value = 'Type'
  ws.getCell('E5').value = 'Cost Type'
  ws.getCell('F5').value = 'Unit Cost'
  ws.getCell('G5').value = '# Travelers'
  ws.getCell('H5').value = 'Cost'
  ws.getCell('I5').value = 'Participants'
  ws.getCell('J5').value = 'Notes / Booking Ref'
  styleColumnHeader(ws.getRow(5), ts)

  // Date validation (col A) — triggers Google Sheets native date picker on double-click.
  ;(ws as any).dataValidations.add('A6:A35', {
    type: 'date',
    operator: 'between',
    formulae: [new Date(2000, 0, 1), new Date(2100, 0, 1)],
    allowBlank: true,
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // Activity type dropdown (col D)
  ;(ws as any).dataValidations.add('D6:D35', {
    type: 'list',
    allowBlank: true,
    formulae: ['"Tour,Museum,Adventure,Cultural,Food & Drink,Nature,Sport,Entertainment,Other"'],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // Cost Type dropdown (col E) — drives the Cost formula
  ;(ws as any).dataValidations.add('E6:E35', {
    type: 'list',
    allowBlank: true,
    formulae: ['"Per Person,Per Group"'],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  // # Travelers (col G) — positive integers only, left blank for user to fill
  ;(ws as any).dataValidations.add('G6:G35', {
    type: 'whole',
    operator: 'greaterThan',
    formulae: [0],
    allowBlank: true,
    showErrorMessage: true,
    errorTitle: 'Invalid value',
    error: 'Please enter a positive whole number.',
  } as ExcelJS.DataValidation)

  for (let i = 0; i < 30; i++) {
    const rowNum = 6 + i
    const row = ws.getRow(rowNum)

    ws.getCell(`A${rowNum}`).numFmt = ts.numFmtDate
    ws.getCell(`F${rowNum}`).numFmt = ts.numFmtCurrency

    // Cost — Per Person: Unit Cost × # Travelers; Per Group (or blank): Unit Cost as entered.
    // IFERROR catches blank inputs and returns "" so empty rows stay clean.
    const hCell = ws.getCell(`H${rowNum}`)
    hCell.value = { formula: `IFERROR(IF(E${rowNum}="Per Person",F${rowNum}*G${rowNum},F${rowNum}),"")`, result: '' }
    hCell.numFmt = ts.numFmtCurrency

    styleDataRow(row, ts, i % 2 === 0)
    row.height = 20
  }

  const totRow = 36
  ws.getCell(`A${totRow}`).value = 'TOTAL'
  const hTot = ws.getCell(`H${totRow}`)
  hTot.value = { formula: 'IFERROR(SUM(H6:H35),0)', result: 0 }
  hTot.numFmt = ts.numFmtCurrency
  styleTotalRow(ws.getRow(totRow), ts)

  ws.getColumn('A').width = 14
  ws.getColumn('B').width = 30
  ws.getColumn('C').width = 12
  ws.getColumn('D').width = 18
  ws.getColumn('E').width = 14
  ws.getColumn('F').width = 14
  ws.getColumn('G').width = 12
  ws.getColumn('H').width = 14
  ws.getColumn('I').width = 25
  ws.getColumn('J').width = 35

  // Informational AI guide below the tracker (TOTAL row is 36; leave a spacer)
  buildRecommendationInsert(ws, totRow + 2, ts, regions, 'excursions')

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 5, activeCell: 'A6' }]
}
