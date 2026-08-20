import type ExcelJS from "exceljs";
import type { WizardState } from "../../../types/wizard";
import type { ThemeStyles } from "../styleFactory";
import { wrappedLineCount, rowHeightForLines } from "../styleFactory";
import { THEMES } from "../../../types/theme";
import { SHEET_LABELS } from "../../../types/wizard";

// Column widths for the INSTRUCTIONS sheet. Body/bullet text is merged across A:F,
// so the merged width (sum below) drives how the text wraps. Column A is also used
// on its own for the Named Range table, so it must be wide enough for "TotalBudget".
const COL_WIDTHS: Record<string, number> = {
  A: 16,
  B: 20,
  C: 22,
  D: 18,
  E: 18,
  F: 18
};
const MERGE_WIDTH = Object.values(COL_WIDTHS).reduce((a, b) => a + b, 0);
const DESC_WIDTH = COL_WIDTHS.C + COL_WIDTHS.D + COL_WIDTHS.E + COL_WIDTHS.F;

export function buildInstructionsSheet(
  wb: ExcelJS.Workbook,
  state: WizardState,
  ts: ThemeStyles
): void {
  const ws = wb.addWorksheet("INSTRUCTIONS");
  ws.properties.tabColor = { argb: "FF888888" };

  let row = 1;

  // Title
  ws.getCell(`A${row}`).value = "HOW TO USE THIS TRAVEL PLANNER";
  ws.mergeCells(`A${row}:F${row}`);
  ws.getCell(`A${row}`).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: THEMES[state.theme].primary }
  } as ExcelJS.Fill;
  ws.getCell(`A${row}`).font = {
    name: ts.fontName,
    size: ts.sizes.title,
    bold: true,
    color: { argb: THEMES[state.theme].primaryText }
  };
  ws.getCell(`A${row}`).alignment = {
    vertical: "middle",
    horizontal: "center"
  };
  ws.getRow(row).height = 42;
  row += 2;

  function sectionHeader(title: string) {
    ws.getCell(`A${row}`).value = title;
    ws.mergeCells(`A${row}:F${row}`);
    ws.getCell(`A${row}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: THEMES[state.theme].secondary }
    } as ExcelJS.Fill;
    ws.getCell(`A${row}`).font = {
      name: ts.fontName,
      size: ts.sizes.sectionHeader,
      bold: true,
      color: { argb: THEMES[state.theme].secondaryText }
    };
    ws.getCell(`A${row}`).alignment = {
      vertical: "middle",
      horizontal: "left"
    };
    ws.getRow(row).height = 22;
  }

  function bodyLine(text: string) {
    const cell = ws.getCell(`A${row}`);
    cell.value = text;
    ws.mergeCells(`A${row}:F${row}`);
    cell.font = { name: ts.fontName, size: ts.sizes.data };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFFFF" }
    } as ExcelJS.Fill;
    cell.alignment = { vertical: "middle", wrapText: true };
    ws.getRow(row).height = rowHeightForLines(
      wrappedLineCount(text, MERGE_WIDTH),
      ts,
      { min: 16 }
    );
  }

  // bulletLine renders a "•" bullet; pass sub:true for an indented sub-item dash.
  function bulletLine(text: string, opts: { sub?: boolean } = {}) {
    const full = opts.sub ? `        –  ${text}` : `  •  ${text}`;
    const cell = ws.getCell(`A${row}`);
    cell.value = full;
    ws.mergeCells(`A${row}:F${row}`);
    cell.font = { name: ts.fontName, size: ts.sizes.data };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: ts.palette.lightBg }
    } as ExcelJS.Fill;
    cell.alignment = { vertical: "middle", wrapText: true };
    ws.getRow(row).height = rowHeightForLines(
      wrappedLineCount(full, MERGE_WIDTH),
      ts,
      { min: 16 }
    );
  }

  function spacer() {
    ws.getRow(row).height = 8;
  }

  // ── Section 1: Getting Started ──────────────────────────────────────────────
  sectionHeader("1.  GETTING STARTED");
  row++;
  bodyLine(
    `This spreadsheet was created for your trip. It works in both Microsoft Excel and Google Sheets.`
  );
  row++;
  bulletLine(
    "Formula and total cells are locked and hidden so you can't accidentally type over them — you can still edit every date, description, and cost field. On ITINERARY only the Date column (B) is locked, since it fills itself from Trip Start; everything else on that tab is yours to rewrite. This is an Excel feature only; if you upload this file to Google Sheets, all cells become editable again."
  );
  row++;
  bodyLine(
    "Key settings live on the OVERVIEW sheet and can be changed at any time:"
  );
  row++;
  bulletLine(
    "B6 — Trip Start: your trip start date. Changing it shifts every date formula across the workbook."
  );
  row++;
  bulletLine(
    "D6 — Trip End: your trip end date. Changing it adds or removes ITINERARY day and date rows."
  );
  row++;
  bulletLine("E9 — Travelers: the number of travelers.");
  row++;
  bulletLine(
    "B16 — Currency: a dropdown of currencies for reference (matches your wizard selection). Changing it here does NOT change the $/€/£ symbols used elsewhere in the workbook — see Section 6 for how to update those manually."
  );
  row++;
  bulletLine(
    "D16 — Total Budget: the sum of the Estimated Budget column, calculated automatically."
  );
  row++;
  spacer();
  row++;

  // ── Section 2: Working with Dates ───────────────────────────────────────────
  sectionHeader("2.  WORKING WITH DATES");
  row++;
  bodyLine("Google Sheets — Date Picker");
  row++;
  bulletLine(
    "OVERVIEW Start Date (B6) and End Date (D6): double-click to open the date picker calendar."
  );
  row++;
  bulletLine(
    "TRANSPORTATION Departure Date and Arrival Date columns: double-click any cell to open the date picker."
  );
  row++;
  bulletLine(
    "ACCOMMODATION Date column: double-click any cell to open the date picker."
  );
  row++;
  bulletLine(
    "DINING Date column: double-click any cell to open the date picker."
  );
  row++;
  bulletLine(
    "EXCURSIONS Date column: double-click any cell to open the date picker."
  );
  row++;
  bulletLine(
    "ITINERARY dates auto-fill from Trip Start — they are formula-driven and do not need to be entered manually."
  );
  row++;
  bulletLine(
    "If a date column shows a number instead of a date: select the column → Format → Number → Date."
  );
  row++;
  spacer();
  row++;
  bodyLine("Excel — Entering Dates");
  row++;
  bulletLine(
    "Excel has no built-in cell date picker. Enter dates by typing directly into the cell."
  );
  row++;
  bulletLine(
    "Recommended format: DD Mon YYYY — for example, 15 Jun 2025. This matches the display format used throughout this workbook."
  );
  row++;
  bulletLine(
    "You can also type your regional short format (e.g., 15/06/2025 or 6/15/2025) and Excel will reformat it automatically."
  );
  row++;
  bulletLine(
    "Always enter a real date value — not plain text — so ITINERARY auto-dates and other date formulas work correctly."
  );
  row++;
  spacer();
  row++;

  // ── Section 3: How Your Budget Rolls Up ─────────────────────────────────────
  sectionHeader("3.  HOW YOUR BUDGET ROLLS UP");
  row++;
  bodyLine(
    "The OVERVIEW Budget Summary tracks six categories: Transportation, Accommodation, Food, Activities, Shopping, and Miscellaneous."
  );
  row++;
  bodyLine(
    "Each row shows Estimated Budget, Actual Spent, Remaining, % Used, and a Status. The budget chart on OVERVIEW updates automatically as you log spending."
  );
  row++;
  bulletLine(
    "Estimated Budget is entered as a default but fully editable — type your own figure over any cell. Total Budget (D16) is the sum of these cells."
  );
  row++;
  bulletLine(
    "Actual Spent is calculated for you. Each category adds its dedicated tab total (when that tab is included) plus any rows you tag with that category on OTHER EXPENSES:"
  );
  row++;
  bulletLine(
    'Transportation  ←  TRANSPORTATION tab  +  OTHER EXPENSES rows tagged "Transportation"',
    { sub: true }
  );
  row++;
  bulletLine(
    'Accommodation  ←  ACCOMMODATION tab  +  OTHER EXPENSES rows tagged "Accommodation"',
    { sub: true }
  );
  row++;
  bulletLine('Food  ←  DINING tab  +  OTHER EXPENSES rows tagged "Food"', {
    sub: true
  });
  row++;
  bulletLine(
    'Activities  ←  EXCURSIONS tab  +  OTHER EXPENSES rows tagged "Activities"',
    { sub: true }
  );
  row++;
  bulletLine(
    'Shopping  ←  OTHER EXPENSES rows tagged "Shopping" (no dedicated tab)',
    { sub: true }
  );
  row++;
  bulletLine(
    'Miscellaneous  ←  OTHER EXPENSES rows tagged "Miscellaneous" (no dedicated tab)',
    { sub: true }
  );
  row++;
  bulletLine(
    'Remaining = Estimated − Actual. When Actual passes Estimated, the Status reads "Over budget" and that bar turns red on the chart.'
  );
  row++;
  spacer();
  row++;

  // ── Section 4: Other Expenses ───────────────────────────────────────────────
  sectionHeader("4.  OTHER EXPENSES");
  row++;
  bodyLine(
    "Use the OTHER EXPENSES sheet for additional / incidental costs (shopping, tips, fees, etc.) that aren’t already captured on a dedicated tab. It is not meant to hold every expense."
  );
  row++;
  bulletLine(
    "Column B: pick a category from the dropdown. It must match a Budget Summary category to roll up — the dropdown guarantees this."
  );
  row++;
  bulletLine(
    "Column D: enter the cost of each expense. The TOTAL row at the bottom sums every entry."
  );
  row++;
  bulletLine(
    "OVERVIEW reads this sheet with SUMIF formulas, so each category total updates in real time."
  );
  row++;
  spacer();
  row++;

  // ── Section 5: Sheet-by-Sheet Guide (only enabled sheets) ──────────────────
  sectionHeader("5.  SHEET GUIDE");
  row++;

  const sheetGuides: { id: keyof typeof SHEET_LABELS; tip: string }[] = [
    {
      id: "itinerary",
      tip: "Plan each day of your trip. Day numbers and dates auto-fill from Trip Start / Trip End on OVERVIEW — extend Trip End to reveal more day rows."
    },
    {
      id: "flights",
      tip: "Log each leg and pick a Mode (Air, Train, Bus, etc.). Enter a Cost per leg; the totals at the top and bottom feed the Transportation budget. Departure and Arrival Date columns support the Google Sheets date picker."
    },
    {
      id: "hotels",
      tip: "Log accommodation costs row by row — use Cost Type to split out nightly rates, taxes, fees, and other charges, or choose Full Cost for a single all-in amount. The total feeds the Accommodation budget. The Date column supports the Google Sheets date picker."
    },
    {
      id: "restaurants",
      tip: "Log dining venues visited (restaurants, cafés, bars, etc.), a 1–5 rating, and a meal type (Breakfast, Lunch, Dinner, etc.). Enter a Cost per visit; the totals feed the Food budget. The Date column supports the date picker."
    },
    {
      id: "excursions",
      tip: "Log tours and activities. Set Cost Type to Per Person (Unit Cost × # Travelers) or Per Group (flat Unit Cost), and add names in Participants. The total feeds the Activities budget; the Date column supports the date picker."
    },
    {
      id: "packingList",
      tip: "Mark items off in the Packed? column — click the cell and pick ✓ from the dropdown. The progress bar at the top updates automatically. To un-pack an item, clear the cell."
    },
    {
      id: "tasks",
      tip: "Pre-trip to-do list. Mark a task done in the Done? column — click the cell and pick ✓ from the dropdown; the progress bar updates automatically. To reopen a task, clear the cell."
    },
    {
      id: "events",
      tip: "Major annual events and festivals at your destination, listed chronologically from January to December."
    }
  ];

  sheetGuides.forEach(({ id, tip }) => {
    if (state.sheets[id]) {
      bulletLine(`${SHEET_LABELS[id]}: ${tip}`);
      row++;
    }
  });

  spacer();
  row++;

  // ── Section 6: Changing Currency Symbols ────────────────────────────────────
  sectionHeader("6.  CHANGING CURRENCY SYMBOLS");
  row++;
  bodyLine(
    "The Currency dropdown on OVERVIEW (B16) is for reference only. Number formats — the $, €, £, etc. shown in Estimated Budget, Actual Spent, Remaining, and every other cost column — are fixed when this workbook is generated and can't follow a dropdown or formula."
  );
  row++;
  bodyLine("To change the symbol on a column or group of cells yourself:");
  row++;
  bulletLine(
    "Select the cells you want to change — for example, the Estimated Budget, Actual Spent, and Remaining columns on OVERVIEW, or a cost column on another tab."
  );
  row++;
  bulletLine(
    'Excel: right-click the selection → Format Cells → Number → Currency, then choose your symbol from the Symbol dropdown. For a symbol not listed, choose Custom and enter a format like "€"#,##0.00.'
  );
  row++;
  bulletLine(
    "Google Sheets: Format → Number → Currency for your locale's default symbol, or Format → Number → More formats → More currencies to pick a different one."
  );
  row++;
  bulletLine(
    "Repeat for each column or sheet where you want the new symbol — there's no single switch that updates the whole workbook at once."
  );
  row++;
  spacer();
  row++;

  // ── Section 7: Named Range Reference ───────────────────────────────────────
  sectionHeader("7.  NAMED RANGE REFERENCE");
  row++;
  bodyLine(
    "These named ranges can be used in any formula across all sheets. All of them point to cells on the OVERVIEW tab:"
  );
  row++;

  const namedRanges = [
    {
      name: "TripStart",
      location: "B6",
      description:
        "Trip start date. Drives ITINERARY auto-dates and all date formulas."
    },
    {
      name: "TripEnd",
      location: "D6",
      description:
        "Trip end date. Controls how many ITINERARY day rows are visible."
    },
    {
      name: "NumAdults",
      location: "E9",
      description:
        "Number of travelers. Used in per-person budget/excursion calculations."
    },
    {
      name: "TotalBudget",
      location: "D16",
      description:
        "Total estimated budget (sum of the Budget Summary Estimated Budget column)."
    }
  ];

  // Named range table header
  ws.getCell(`A${row}`).value = "Name";
  ws.getCell(`B${row}`).value = "OVERVIEW Cell";
  ws.getCell(`C${row}`).value = "Description";
  ws.mergeCells(`C${row}:F${row}`);
  ["A", "B", "C"].forEach((col) => {
    ws.getCell(`${col}${row}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: ts.palette.secondary }
    } as ExcelJS.Fill;
    ws.getCell(`${col}${row}`).font = {
      name: ts.fontName,
      size: ts.sizes.data,
      bold: true
    };
    ws.getCell(`${col}${row}`).alignment = { vertical: "middle" };
  });
  ws.getRow(row).height = 18;
  row++;

  namedRanges.forEach(({ name, location, description }, i) => {
    ws.getCell(`A${row}`).value = name;
    ws.getCell(`B${row}`).value = location;
    ws.getCell(`C${row}`).value = description;
    ws.mergeCells(`C${row}:F${row}`);
    const fillArgb = i % 2 === 0 ? ts.palette.lightBg : "FFFFFFFF";
    ["A", "B", "C"].forEach((col) => {
      ws.getCell(`${col}${row}`).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: fillArgb }
      } as ExcelJS.Fill;
      ws.getCell(`${col}${row}`).font = {
        name: ts.fontName,
        size: ts.sizes.data
      };
      ws.getCell(`${col}${row}`).alignment = {
        vertical: "middle",
        wrapText: true
      };
    });
    ws.getRow(row).height = rowHeightForLines(
      wrappedLineCount(description, DESC_WIDTH),
      ts,
      { min: 18 }
    );
    row++;
  });

  spacer();
  row += 2;

  Object.entries(COL_WIDTHS).forEach(([col, width]) => {
    ws.getColumn(col).width = width;
  });
}
