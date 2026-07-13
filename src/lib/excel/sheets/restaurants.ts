import type ExcelJS from "exceljs";
import type { WizardState } from "../../../types/wizard";
import type { ThemeStyles } from "../styleFactory";
import type { RegionRec } from "../../recommendations/types";
import {
  styleTitleRow,
  styleSectionHeader,
  styleColumnHeader,
  styleDataRow,
  styleDataCell,
  styleTotalRow
} from "../styleFactory";
import { THEMES } from "../../../types/theme";
import { buildRecommendationInsert } from "./recommendationInsert";

export function buildRestaurantsSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles,
  regions?: RegionRec[]
): void {
  const ws = wb.addWorksheet("DINING");
  ws.properties.tabColor = { argb: THEMES[state.theme].tabColor };

  ws.getCell("A1").value = "DINING";
  ws.mergeCells("A1:H1");
  styleTitleRow(ws.getRow(1), ts);

  ws.getCell("A2").value =
    "Log restaurants, cafes, coffee shops, bars, etc. Rating: 1–5 stars. Type: Breakfast / Lunch / Dinner / Snack.";
  ws.mergeCells("A2:H2");
  styleSectionHeader(ws.getRow(2), ts);

  ws.getCell("A3").value = "TOTAL FOOD COST";
  ws.getCell("A3").font = {
    name: ts.fontName,
    size: ts.sizes.header,
    bold: true
  };
  ws.mergeCells("A3:D3");
  // Cost is now col G; data rows start at 6 (not 8 — that was a bug). Sum the data
  // rows only (6–55) so the bottom TOTAL row at 56 isn't double-counted.
  ws.getCell("G3").value = { formula: "IFERROR(SUM(G6:G55),0)", result: 0 };
  ws.getCell("G3").numFmt = ts.numFmtCurrency;
  ws.getCell("G3").font = {
    name: ts.fontName,
    size: ts.sizes.header,
    bold: true
  };
  ws.getRow(3).height = 22;

  // Column order: Date | Venue Name | Address/Location | Meal Type | Cuisine | Rating | Cost | Notes
  ws.getCell("A5").value = "Date";
  ws.getCell("B5").value = "Venue Name";
  ws.getCell("C5").value = "Address / Location";
  ws.getCell("D5").value = "Meal Type";
  ws.getCell("E5").value = "Cuisine";
  ws.getCell("F5").value = "Rating (1–5)";
  ws.getCell("G5").value = "Cost";
  ws.getCell("H5").value = "Notes / Highlights";
  styleColumnHeader(ws.getRow(5), ts);

  // Date validation (col A) — triggers Google Sheets native date picker on double-click.
  (ws as any).dataValidations.add("A6:A55", {
    type: "date",
    operator: "between",
    formulae: [new Date(2000, 0, 1), new Date(2100, 0, 1)],
    allowBlank: true,
    showErrorMessage: false
  } as ExcelJS.DataValidation);

  // Meal type validation (col D) — bounded to data rows only so Google Sheets
  // doesn't render dropdown carets in the AI recommendation table below row 55.
  (ws as any).dataValidations.add("D6:D55", {
    type: "list",
    allowBlank: true,
    formulae: ['"Breakfast,Lunch,Dinner,Snack,Coffee,Drinks"'],
    showErrorMessage: false
  } as ExcelJS.DataValidation);

  // Rating validation (col F)
  (ws as any).dataValidations.add("F6:F55", {
    type: "whole",
    operator: "between",
    formulae: [1, 5],
    allowBlank: true,
    showErrorMessage: true,
    errorTitle: "Invalid rating",
    error: "Please enter a number from 1 to 5."
  } as ExcelJS.DataValidation);

  for (let i = 0; i < 50; i++) {
    const rowNum = 6 + i;
    const row = ws.getRow(rowNum);

    ws.getCell(`A${rowNum}`).numFmt = ts.numFmtDate;
    ws.getCell(`G${rowNum}`).numFmt = ts.numFmtCurrency;

    styleDataRow(row, ts, i % 2 === 0);
    // Cost (G) is user-entered, so it's empty in the template — stripe it explicitly
    // so the column matches the populated columns (styleDataRow skips empty cells).
    styleDataCell(ws.getCell(`G${rowNum}`), ts, i % 2 === 0);
    row.height = 20;
  }

  // TOTAL row — regular total of the logged costs, consistent with the other tabs.
  const totRow = 56;
  ws.getCell(`A${totRow}`).value = "TOTAL";
  const gTot = ws.getCell(`G${totRow}`);
  gTot.value = { formula: "IFERROR(SUM(G6:G55),0)", result: 0 };
  gTot.numFmt = ts.numFmtCurrency;
  styleTotalRow(ws.getRow(totRow), ts);

  ws.getColumn("A").width = 14;
  ws.getColumn("B").width = 28;
  ws.getColumn("C").width = 30;
  ws.getColumn("D").width = 14;
  ws.getColumn("E").width = 16;
  ws.getColumn("F").width = 14;
  ws.getColumn("G").width = 12;
  ws.getColumn("H").width = 35;

  // Informational AI guide below the tracker (TOTAL row is 56; leave a spacer at 57)
  buildRecommendationInsert(ws, 58, ts, regions, "restaurants");

  ws.views = [{ state: "frozen", xSplit: 0, ySplit: 5, activeCell: "A6" }];
}
