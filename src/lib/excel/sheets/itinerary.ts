import type ExcelJS from 'exceljs'
import type { WizardState } from '../../../types/wizard'
import type { ThemeStyles } from '../styleFactory'
import type { ItineraryDay } from '../../recommendations/types'
import { styleTitleRow, styleSectionHeader, styleColumnHeader, styleDataRow, truncate, wrappedLineCount, rowHeightForLines } from '../styleFactory'
import { THEMES } from '../../../types/theme'
import { addDays } from 'date-fns'

export function buildItinerarySheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  itinerary?: ItineraryDay[]
): void {
  const ws = wb.addWorksheet('ITINERARY')
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor }

  const startDate = state.startDate ? new Date(state.startDate) : new Date()

  // Row 1: Title
  ws.getCell('A1').value = `ITINERARY  —  ${(state.destination || 'My Trip').toUpperCase()}`
  ws.mergeCells('A1:I1')
  styleTitleRow(ws.getRow(1), ts)

  // Row 2: Section label
  ws.getCell('A2').value = `Dates and day numbers auto-update from TripStart / TripEnd in OVERVIEW`
  ws.mergeCells('A2:I2')
  styleSectionHeader(ws.getRow(2), ts)

  // Row 3: Column headers
  ws.getCell('A3').value = 'Day'
  ws.getCell('B3').value = 'Date'
  ws.getCell('C3').value = 'Location / City'
  ws.getCell('D3').value = 'Morning'
  ws.getCell('E3').value = 'Afternoon'
  ws.getCell('F3').value = 'Evening'
  ws.getCell('G3').value = 'Accommodation'
  ws.getCell('H3').value = 'Transportation'
  ws.getCell('I3').value = 'Notes'
  styleColumnHeader(ws.getRow(3), ts)

  // Data rows — generate the wizard's max day count (30) so the sheet can grow when
  // the user extends the trip in OVERVIEW. Each Day/Date cell is guarded by a TripEnd
  // formula: rows beyond the current trip length render blank and reveal themselves
  // automatically when TripEnd moves out. Rows within state.duration get cached
  // results so they display immediately on open.
  const MAX_DAYS = 30 // wizard slider max; extra rows show/blank via TripEnd formula
  for (let day = 0; day < MAX_DAYS; day++) {
    const rowNum = 4 + day
    const row = ws.getRow(rowNum)
    const withinTrip = day < state.duration

    // Day number — formula so it blanks when the row falls beyond TripEnd.
    ws.getCell(`A${rowNum}`).value = {
      formula: `IF(TripStart+${day}<=TripEnd,${day + 1},"")`,
      result: withinTrip ? day + 1 : '',
    }

    // Auto-date formula using TripStart/TripEnd named ranges. Cache a result so the
    // date renders immediately — a freshly written formula with no cached result shows
    // blank until Excel/Sheets triggers a recalc (see overview.ts countdown cell).
    const dateCell = ws.getCell(`B${rowNum}`)
    dateCell.value = {
      formula: `IF(TripStart+${day}<=TripEnd,TripStart+${day},"")`,
      result: withinTrip ? addDays(startDate, day) : '',
    }
    dateCell.numFmt = 'ddd dd mmm'

    // Prefill from recommendations if available (only within the initial trip length).
    // Truncate verbose entries, then size the row to the longest wrapped cell (col
    // widths below) and wrap each text cell — merged-free here, but Excel still needs
    // an explicit height.
    const rec = withinTrip
      ? (itinerary?.find((d) => d.day === day + 1) ?? itinerary?.[day])
      : undefined
    if (rec) {
      // [value, column letter, column width, max chars]
      const fields: [string, string, number, number][] = [
        [truncate(rec.location, 80), 'C', 22, 80],
        [truncate(rec.morning, 180), 'D', 34, 180],
        [truncate(rec.afternoon, 180), 'E', 34, 180],
        [truncate(rec.evening, 180), 'F', 34, 180],
        [truncate(rec.transport, 180), 'H', 22, 180],
        [truncate(rec.notes, 180), 'I', 36, 180],
      ]
      let maxLines = 1
      fields.forEach(([val, col, width]) => {
        if (val) ws.getCell(`${col}${rowNum}`).value = val
        maxLines = Math.max(maxLines, wrappedLineCount(val, width))
      })

      styleDataRow(row, ts, day % 2 === 0)
      row.height = rowHeightForLines(maxLines, ts, { min: 24, maxLines: 12 })
      // Per-cell wrap + top align (row-level alignment is unreliable in ExcelJS)
      for (const col of ['C', 'D', 'E', 'F', 'G', 'H', 'I']) {
        ws.getCell(`${col}${rowNum}`).alignment = { vertical: 'top', wrapText: true }
      }
    } else {
      styleDataRow(row, ts, day % 2 === 0)
      row.height = 24
    }
  }

  // Column widths — C–I widened ~20% so AI-prefilled content is less cramped;
  // Day (A) and Date (B) left as-is.
  ws.getColumn('A').width = 6
  ws.getColumn('B').width = 14
  ws.getColumn('C').width = 22
  ws.getColumn('D').width = 34
  ws.getColumn('E').width = 34
  ws.getColumn('F').width = 34
  ws.getColumn('G').width = 26
  ws.getColumn('H').width = 22
  ws.getColumn('I').width = 36

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3, activeCell: 'A4' }]
}
