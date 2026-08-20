import type ExcelJS from 'exceljs'

// Sheets with no formulas/totals worth guarding — left fully editable.
const UNPROTECTED_SHEETS = new Set(['INSTRUCTIONS', 'ANNUAL EVENTS'])

// ITINERARY is a writing surface, not a calculator, so it's deliberately near-open:
// every data cell is unlocked except the auto-Date column B (itinerary.ts), which stays
// locked so typing over a date can't sever that row's TripStart link.
//
// It still has to carry <sheetProtection> to get that: in OOXML a cell's `locked` flag
// is inert unless the sheet is protected, so "no sheet protection but a protected
// column" isn't expressible — the sheet-level flag IS the mechanism. What we can do is
// switch off every OTHER restriction protection normally brings, so the sheet feels
// unprotected: users can restyle cells, resize and insert/delete rows, sort and filter.
// (ExcelJS's defaults disallow all of these.)
const RELAXED_PROTECTION: Record<string, Partial<ExcelJS.WorksheetProtection>> = {
  ITINERARY: {
    formatCells: true,
    formatColumns: true,
    formatRows: true,
    insertRows: true,
    insertColumns: true,
    insertHyperlinks: true,
    deleteRows: true,
    deleteColumns: true,
    sort: true,
    autoFilter: true,
  },
}

// Locks every worksheet except the ones above. Each sheet builder already marks its
// own data-entry cells `protection: { locked: false }` and its formula/total cells
// `protection: { hidden: true }` at the point those cells are written — this just
// flips on `sheetProtection` so those per-cell flags take effect. No password: this
// is a guardrail against accidentally typing over a formula, not real security (any
// protected .xlsx is trivially unprotected regardless of password), and skipping the
// password avoids ExcelJS's ~600ms/sheet password-hashing cost. Must run before
// `wb.xlsx.writeBuffer()` — protection is an in-memory worksheet property, not part
// of the post-processing (chart/calcChain) pipeline.
//
// Note: Google Sheets ignores sheetProtection/cell.protection on upload, so this is
// an Excel-only guardrail (see CLAUDE.md compatibility constraints).
export async function protectWorkbook(wb: ExcelJS.Workbook): Promise<void> {
  for (const ws of wb.worksheets) {
    if (UNPROTECTED_SHEETS.has(ws.name)) continue
    await ws.protect('', RELAXED_PROTECTION[ws.name] ?? {})
  }
}
