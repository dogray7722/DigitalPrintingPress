import type ExcelJS from 'exceljs'

// Sheets with no formulas/totals worth guarding — left fully editable.
const UNPROTECTED_SHEETS = new Set(['INSTRUCTIONS', 'ANNUAL EVENTS'])

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
    await ws.protect('', {})
  }
}
