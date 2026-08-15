import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataRow } from '../styleFactory'
import { THEMES } from '../../../types/theme'

const DEFAULT_TASKS = [
  { category: 'Documents', task: 'Renew passport (if expiring within 6 months)', priority: 'High' },
  { category: 'Documents', task: 'Apply for visa / entry permit', priority: 'High' },
  { category: 'Documents', task: 'Purchase travel insurance', priority: 'High' },
  { category: 'Documents', task: 'Scan and email all documents to yourself', priority: 'High' },
  { category: 'Flights', task: 'Book outbound flight', priority: 'High' },
  { category: 'Flights', task: 'Book return flight', priority: 'High' },
  { category: 'Flights', task: 'Online check-in (24–48hrs before departure)', priority: 'Medium' },
  { category: 'Flights', task: 'Download airline app', priority: 'Low' },
  { category: 'Accommodation', task: 'Book first night accommodation', priority: 'High' },
  { category: 'Accommodation', task: 'Book all remaining accommodation', priority: 'High' },
  { category: 'Accommodation', task: 'Confirm all reservations 1 week before travel', priority: 'Medium' },
  { category: 'Money', task: 'Notify bank of travel dates', priority: 'High' },
  { category: 'Money', task: 'Get local currency (cash)', priority: 'Medium' },
  { category: 'Money', task: 'Set up international data plan', priority: 'Medium' },
  { category: 'Health', task: 'Check required vaccinations', priority: 'High' },
  { category: 'Health', task: 'Pack prescription medications (enough supply)', priority: 'High' },
  { category: 'Health', task: 'Pack first aid basics', priority: 'Medium' },
  { category: 'Activities', task: 'Book must-see attractions in advance', priority: 'Medium' },
  { category: 'Activities', task: 'Research restaurant reservations needed', priority: 'Low' },
  { category: 'Activities', task: 'Download offline maps (Google Maps / Maps.me)', priority: 'Medium' },
  { category: 'Packing', task: 'Check airline baggage allowance', priority: 'Medium' },
  { category: 'Packing', task: 'Weigh luggage before departure', priority: 'Low' },
  { category: 'Final Checks', task: 'Arrange airport transfer / taxi', priority: 'Medium' },
  { category: 'Final Checks', task: 'Set out-of-office email', priority: 'Low' },
  { category: 'Final Checks', task: 'Arrange for mail / pet care', priority: 'Medium' },
  { category: 'Final Checks', task: 'Charge all devices & power banks', priority: 'Medium' },
]

export function buildTasksSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles
): void {
  const ws = wb.addWorksheet('TASKS')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  ws.getCell('A1').value = 'PRE-TRIP TASKS'
  ws.mergeCells('A1:F1')
  styleTitleRow(ws.getRow(1), ts)

  // Progress formula
  ws.getCell('A2').value = {
    formula:
      'IFERROR("TASK PROGRESS: "&TEXT(COUNTIF(D5:D500,"✓")/COUNTA(B5:B500),"0%")&" complete","TASK PROGRESS")',
    result: 'TASK PROGRESS: 0% complete',
  }
  ws.getCell('A2').protection = { hidden: true }
  ws.mergeCells('A2:F2')
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

  ws.getCell('A3').value = 'Add ✓ in the Done? column as you complete each task.'
  ws.mergeCells('A3:F3')
  styleSectionHeader(ws.getRow(3), ts)

  ws.getCell('A4').value = '#'
  ws.getCell('B4').value = 'Task'
  ws.getCell('C4').value = 'Category'
  ws.getCell('D4').value = 'Done?'
  ws.getCell('E4').value = 'Priority'
  ws.getCell('F4').value = 'Due / Notes'
  styleColumnHeader(ws.getRow(4), ts)

  // Priority validation
  ;(ws as any).dataValidations.add('E5:E500', {
    type: 'list',
    allowBlank: true,
    formulae: ['"High,Medium,Low"'],
    showErrorMessage: false,
  } as ExcelJS.DataValidation)

  let rowNum = 5
  DEFAULT_TASKS.forEach((task, i) => {
    const row = ws.getRow(rowNum)
    ws.getCell(`A${rowNum}`).value = i + 1
    ws.getCell(`B${rowNum}`).value = task.task
    ws.getCell(`C${rowNum}`).value = task.category
    ws.getCell(`D${rowNum}`).value = ''
    ws.getCell(`E${rowNum}`).value = task.priority
    ws.getCell(`F${rowNum}`).value = ''

    // Done?, Priority, Due/Notes are user-entered; Task/Category stay locked (template content).
    ws.getCell(`D${rowNum}`).protection = { locked: false }
    ws.getCell(`E${rowNum}`).protection = { locked: false }
    ws.getCell(`F${rowNum}`).protection = { locked: false }

    styleDataRow(row, ts, i % 2 === 0)
    row.height = 18
    rowNum++
  })

  // Extra blank rows
  ws.getCell(`A${rowNum + 1}`).value = '— ADD YOUR OWN TASKS BELOW —'
  ws.mergeCells(`A${rowNum + 1}:F${rowNum + 1}`)
  ws.getCell(`A${rowNum + 1}`).font = {
    name: ts.fontName,
    size: ts.sizes.data,
    italic: true,
    color: { argb: 'FF999999' },
  }

  for (let extra = 0; extra < 20; extra++) {
    const r = rowNum + 2 + extra
    const row = ws.getRow(r)
    ws.getCell(`A${r}`).value = DEFAULT_TASKS.length + extra + 1
    // Blank rows for the buyer's own tasks — user-entered except the auto-numbered index.
    for (const col of ['B', 'C', 'D', 'E', 'F']) {
      ws.getCell(`${col}${r}`).protection = { locked: false }
    }
    styleDataRow(row, ts, extra % 2 === 0)
    row.height = 18
  }

  ws.getColumn('A').width = 6
  ws.getColumn('B').width = 45
  ws.getColumn('C').width = 18
  ws.getColumn('D').width = 10
  ws.getColumn('E').width = 12
  ws.getColumn('F').width = 30

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 4, activeCell: 'A5' }]
}
