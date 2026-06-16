import type ExcelJS from 'exceljs'
import type { WizardState } from '../../types/wizard'
import { THEMES, FONT_NAMES, FONT_SIZES, type ColorPalette, type FontSizeValues } from '../../types/theme'
import { getCurrencyNumFmt } from '../../data/currencies'

export interface ThemeStyles {
  palette: ColorPalette
  fontName: string
  sizes: FontSizeValues
  numFmtCurrency: string
  numFmtDate: string
  numFmtPercent: string
}

export function getThemeStyles(state: WizardState): ThemeStyles {
  return {
    palette: THEMES[state.theme],
    fontName: FONT_NAMES[state.fontFamily],
    sizes: FONT_SIZES[state.fontSize],
    numFmtCurrency: getCurrencyNumFmt(state.currency),
    numFmtDate: 'dd mmm yyyy',
    numFmtPercent: '0%',
  }
}

// ── Cell styling helpers ─────────────────────────────────────────────────────

type Cell = ExcelJS.Cell
type Row = ExcelJS.Row

function solidFill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } } as ExcelJS.Fill
}

function thinBorder(argb: string): Partial<ExcelJS.Borders> {
  const side = { style: 'thin' as const, color: { argb } }
  return { top: side, bottom: side, left: side, right: side }
}

function bottomBorder(argb: string): Partial<ExcelJS.Borders> {
  return { bottom: { style: 'thin', color: { argb } } }
}

// Apply title row style (large, themed, merged header)
export function styleTitleRow(row: Row, ts: ThemeStyles): void {
  row.height = 42
  row.eachCell((cell) => {
    cell.fill = solidFill(ts.palette.primary)
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.title,
      bold: true,
      color: { argb: ts.palette.primaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false }
  })
}

// Apply section header (sub-heading band)
export function styleSectionHeader(row: Row, ts: ThemeStyles): void {
  row.height = 22
  row.eachCell((cell) => {
    cell.fill = solidFill(ts.palette.secondary)
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.sectionHeader,
      bold: true,
      color: { argb: ts.palette.secondaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
  })
}

// Apply column header style
export function styleColumnHeader(row: Row, ts: ThemeStyles): void {
  row.height = 20
  row.eachCell({ includeEmpty: false }, (cell) => {
    cell.fill = solidFill(ts.palette.secondary)
    cell.font = {
      name: ts.fontName,
      size: ts.sizes.header,
      bold: true,
      color: { argb: ts.palette.secondaryText },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'center' }
    cell.border = thinBorder(ts.palette.border) as ExcelJS.Borders
  })
}

// Apply data row (alternating)
export function styleDataRow(row: Row, ts: ThemeStyles, isEven: boolean): void {
  row.height = 18
  const fillArgb = isEven ? ts.palette.lightBg : 'FFFFFFFF'
  row.eachCell({ includeEmpty: false }, (cell) => {
    cell.fill = solidFill(fillArgb)
    cell.font = { name: ts.fontName, size: ts.sizes.data }
    cell.border = bottomBorder(ts.palette.border) as ExcelJS.Borders
  })
}

// Stripe a single data cell with the same alternating fill / font / border as
// styleDataRow. styleDataRow uses includeEmpty:false, so it skips cells with no
// value — use this to stripe a column whose template cells are empty (e.g. a
// user-entered Cost column) so it matches the populated columns visually.
export function styleDataCell(cell: Cell, ts: ThemeStyles, isEven: boolean): void {
  cell.fill = solidFill(isEven ? ts.palette.lightBg : 'FFFFFFFF')
  cell.font = { name: ts.fontName, size: ts.sizes.data }
  cell.border = bottomBorder(ts.palette.border) as ExcelJS.Borders
}

// Apply total/summary row
export function styleTotalRow(row: Row, ts: ThemeStyles): void {
  row.height = 20
  row.eachCell({ includeEmpty: false }, (cell) => {
    cell.fill = solidFill(ts.palette.mediumBg)
    cell.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
    cell.border = thinBorder(ts.palette.border) as ExcelJS.Borders
  })
}

// Apply a label cell (key-value left column)
export function styleLabelCell(cell: Cell, ts: ThemeStyles): void {
  cell.fill = solidFill(ts.palette.lightBg)
  cell.font = { name: ts.fontName, size: ts.sizes.data, bold: true }
  cell.alignment = { vertical: 'middle', horizontal: 'left' }
}

// Apply a value cell (key-value right column)
export function styleValueCell(cell: Cell, ts: ThemeStyles): void {
  cell.fill = solidFill('FFFFFFFF')
  cell.font = { name: ts.fontName, size: ts.sizes.data }
  cell.alignment = { vertical: 'middle', horizontal: 'left' }
  cell.border = bottomBorder(ts.palette.border) as ExcelJS.Borders
}

// ── Text-fitting helpers ─────────────────────────────────────────────────────
// ExcelJS has no real autofit, and Excel desktop won't auto-grow MERGED cells, so
// for wrapped AI text we estimate how many lines it needs and set an explicit row
// height. Google Sheets honors that height on import, so over-estimating slightly
// (we round up) is safe; under-estimating clips.

// Cap a string at `max` chars on a word boundary and append an ellipsis. Guards
// against one verbose AI entry blowing a row up to a dozen lines.
export function truncate(text: string | undefined, max: number): string {
  const s = text ?? ''
  if (s.length <= max) return s
  return s.slice(0, max - 1).replace(/\s+\S*$/, '').trimEnd() + '…'
}

// Estimate the number of wrapped lines `text` needs in a column `widthChars` wide.
// Honors explicit newlines. Conservative (ceil) so we never under-size.
export function wrappedLineCount(text: string | undefined, widthChars: number): number {
  if (!text) return 1
  const perLine = Math.max(6, Math.floor(widthChars * 0.95))
  return text
    .split('\n')
    .reduce((sum, seg) => sum + Math.max(1, Math.ceil(seg.length / perLine)), 0)
}

// Convert a line count into an ExcelJS row height (points), scaled to the theme's
// data font size, floored at `min` and capped at `maxLines`.
export function rowHeightForLines(
  lines: number,
  ts: ThemeStyles,
  opts: { min?: number; maxLines?: number } = {}
): number {
  const { min = 18, maxLines = 10 } = opts
  const capped = Math.min(Math.max(1, lines), maxLines)
  const lineHeight = ts.sizes.data * 1.4
  return Math.max(min, Math.round(capped * lineHeight + 6))
}

// Auto-fit column widths (approximation)
export function autoFitColumns(ws: ExcelJS.Worksheet): void {
  ws.columns.forEach((col) => {
    if (!col || !col.eachCell) return
    let maxLen = 8
    col.eachCell({ includeEmpty: false }, (cell) => {
      const val = cell.value?.toString() ?? ''
      maxLen = Math.max(maxLen, val.length + 2)
    })
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    col.width = Math.min(maxLen, 40)
  })
}
