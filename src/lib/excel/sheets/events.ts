import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import type { EventRec } from '../../recommendations/types'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataRow } from '../styleFactory'
import { THEMES } from '../../../types/theme'

const MAX_EVENTS = 20

const MONTHS = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december',
]

// Maps a free-text "when" string (e.g. "Late March", "July 14", "December–January")
// to a 0–11 month index for chronological sorting. Picks the month token that appears
// earliest in the string (reading order), matching the full name or 3-letter abbreviation
// on word boundaries. Returns 12 for undated/unrecognised entries so they sort last.
function monthIndex(monthOrDate?: string): number {
  if (!monthOrDate) return 12
  const s = monthOrDate.toLowerCase()
  let bestPos = Infinity
  let bestMonth = 12
  for (let m = 0; m < 12; m++) {
    const full = MONTHS[m]
    const re = new RegExp(`\\b(${full}|${full.slice(0, 3)})\\b`)
    const match = re.exec(s)
    if (match && match.index < bestPos) {
      bestPos = match.index
      bestMonth = m
    }
  }
  return bestMonth
}

export function buildEventsSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  events: EventRec[]
): void {
  const ws = wb.addWorksheet('ANNUAL EVENTS')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  const dest = state.destination || 'destination'

  // Row 1: Title
  ws.getCell('A1').value = `MAJOR ANNUAL EVENTS  —  ${dest.toUpperCase()}`
  ws.mergeCells('A1:F1')
  styleTitleRow(ws.getRow(1), ts)

  // Row 2: Sub-heading
  const eventCount = Math.min(events.length, MAX_EVENTS)
  ws.getCell('A2').value = `Top ${eventCount} festivals, holidays, and annual events in ${dest}, listed chronologically.`
  ws.mergeCells('A2:F2')
  styleSectionHeader(ws.getRow(2), ts)

  // Row 4: Column headers
  ws.getCell('A4').value = '#'
  ws.getCell('B4').value = 'Event Name'
  ws.getCell('C4').value = 'When'
  ws.getCell('D4').value = 'Type'
  ws.getCell('E4').value = 'Location / Venue'
  ws.getCell('F4').value = 'Description'
  styleColumnHeader(ws.getRow(4), ts)

  // Keep the AI's top-N selection, then display it chronologically (Jan→Dec).
  // Stable: events in the same month (or both undated) keep their original order.
  const visible = events
    .slice(0, MAX_EVENTS)
    .map((e, i) => ({ e, i }))
    .sort((a, b) => monthIndex(a.e.monthOrDate) - monthIndex(b.e.monthOrDate) || a.i - b.i)
    .map(({ e }) => e)

  visible.forEach((event, i) => {
    const rowNum = 5 + i
    const row = ws.getRow(rowNum)

    ws.getCell(`A${rowNum}`).value = i + 1
    ws.getCell(`B${rowNum}`).value = event.name ?? ''
    ws.getCell(`C${rowNum}`).value = event.monthOrDate ?? ''
    ws.getCell(`D${rowNum}`).value = event.type ?? ''
    ws.getCell(`E${rowNum}`).value = event.location ?? ''
    ws.getCell(`F${rowNum}`).value = event.description ?? ''

    styleDataRow(row, ts, i % 2 === 0)
    row.height = 30
    row.alignment = { vertical: 'top', wrapText: true }
  })

  // Fill remaining rows up to MAX_EVENTS to keep visual structure if we got fewer
  for (let i = visible.length; i < MAX_EVENTS; i++) {
    const rowNum = 5 + i
    const row = ws.getRow(rowNum)
    ws.getCell(`A${rowNum}`).value = i + 1
    styleDataRow(row, ts, i % 2 === 0)
    row.height = 20
  }

  ws.getColumn('A').width = 6
  ws.getColumn('B').width = 32
  ws.getColumn('C').width = 18
  ws.getColumn('D').width = 16
  ws.getColumn('E').width = 26
  ws.getColumn('F').width = 50

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 4, activeCell: 'A5' }]
}
