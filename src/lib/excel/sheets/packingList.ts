import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataRow } from '../styleFactory'
import { THEMES } from '../../../types/theme'
import { getPackingList } from '../../../data/packingLists'

export function buildPackingListSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles
): void {
  const ws = wb.addWorksheet('PACKING LIST')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  ws.getCell('A1').value = 'PACKING LIST'
  ws.mergeCells('A1:E1')
  styleTitleRow(ws.getRow(1), ts)

  // Progress formula row
  ws.getCell('A2').value = {
    formula:
      'IFERROR("PACKING PROGRESS: "&TEXT(COUNTIF(C5:C500,"✓")/COUNTA(B5:B500),"0%")&" complete","PACKING PROGRESS")',
  }
  ws.mergeCells('A2:E2')
  ws.getCell('A2').font = {
    name: ts.fontName,
    size: ts.sizes.header,
    bold: true,
    color: { argb: ts.palette.primaryText },
  }
  ws.getCell('A2').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: ts.palette.primary },
  } as ExcelJS.Fill
  ws.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center' }
  ws.getRow(2).height = 24

  ws.getCell('A3').value = 'Add ✓ in the Packed? column as you pack. Progress updates automatically.'
  ws.mergeCells('A3:E3')
  styleSectionHeader(ws.getRow(3), ts)

  // Column headers
  ws.getCell('A4').value = 'Category'
  ws.getCell('B4').value = 'Item'
  ws.getCell('C4').value = 'Packed?'
  ws.getCell('D4').value = 'Quantity'
  ws.getCell('E4').value = 'Notes'
  styleColumnHeader(ws.getRow(4), ts)

  // Generate packing list from data
  const items = getPackingList(state.travelMonth)
  let currentCategory = ''
  let rowNum = 5
  let dataRowIdx = 0

  items.forEach((item) => {
    if (item.category !== currentCategory) {
      // Category separator row
      if (currentCategory !== '') {
        rowNum++ // blank separator
      }
      currentCategory = item.category
      ws.getCell(`A${rowNum}`).value = item.category.toUpperCase()
      ws.mergeCells(`A${rowNum}:E${rowNum}`)
      ws.getCell(`A${rowNum}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: ts.palette.mediumBg },
      } as ExcelJS.Fill
      ws.getCell(`A${rowNum}`).font = {
        name: ts.fontName,
        size: ts.sizes.data,
        bold: true,
      }
      ws.getRow(rowNum).height = 18
      rowNum++
    }

    const row = ws.getRow(rowNum)
    ws.getCell(`A${rowNum}`).value = ''
    ws.getCell(`B${rowNum}`).value = item.item
    ws.getCell(`C${rowNum}`).value = ''
    ws.getCell(`D${rowNum}`).value = 1
    ws.getCell(`E${rowNum}`).value = ''

    styleDataRow(row, ts, dataRowIdx % 2 === 0)
    row.height = 18
    rowNum++
    dataRowIdx++
  })

  // Add blank rows for custom items
  ws.getCell(`A${rowNum + 1}`).value = '— ADD YOUR OWN ITEMS BELOW —'
  ws.mergeCells(`A${rowNum + 1}:E${rowNum + 1}`)
  ws.getCell(`A${rowNum + 1}`).font = {
    name: ts.fontName,
    size: ts.sizes.data,
    italic: true,
    color: { argb: 'FF999999' },
  }
  ws.getRow(rowNum + 1).height = 18

  for (let extra = 0; extra < 20; extra++) {
    const r = rowNum + 2 + extra
    const row = ws.getRow(r)
    ws.getCell(`D${r}`).value = 1
    styleDataRow(row, ts, extra % 2 === 0)
    row.height = 18
  }

  ws.getColumn('A').width = 18
  ws.getColumn('B').width = 36
  ws.getColumn('C').width = 12
  ws.getColumn('D').width = 10
  ws.getColumn('E').width = 30

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 4, activeCell: 'A5' }]
}
