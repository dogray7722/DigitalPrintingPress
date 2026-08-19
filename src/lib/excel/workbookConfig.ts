import type ExcelJS from "exceljs";
import type { WizardState } from "../../types/wizard";

export function configureWorkbook(
  wb: ExcelJS.Workbook,
  state: WizardState
): void {
  wb.creator = "Travel Planner";
  wb.created = new Date();
  wb.modified = new Date();
  wb.properties.date1904 = false;

  // NB: fullCalcOnLoad is deliberately NOT set. The full rebuild it forces on
  // load re-breaks Excel for Mac's chart repaint ordering (bars lag one edit
  // behind the cells). Instead: every formula cell carries a correct cached
  // `result:` so it renders on open without a recalc, and injectCalcChain()
  // (src/lib/excel/calcChain.ts) writes the calcChain part + native calcId so
  // Excel treats the file like one of its own. Verified by typing tests in
  // Excel for Mac 16.110 — keep all three pieces in sync.

  const dest = state.destination || "My Trip";
  wb.title = `${dest} Travel Planner`;
  wb.subject = "Travel Planning Spreadsheet";
  wb.description = `Generated travel planner for ${dest}`;
  wb.keywords = "travel planner itinerary budget";
}

// Must be called AFTER OVERVIEW sheet is fully populated.
// NB: ExcelJS's signature is add(location, name) — location first. Passing them
// the other way round silently fails to register the name, leaving every
// TripStart-based formula (ITINERARY dates, OVERVIEW countdown) unresolved.
export function addNamedRanges(wb: ExcelJS.Workbook): void {
  wb.definedNames.add("'OVERVIEW'!$A$5", "TripStart");
  wb.definedNames.add("'OVERVIEW'!$B$5", "TripEnd");
  wb.definedNames.add("'OVERVIEW'!$C$7", "NumAdults");
  wb.definedNames.add("'OVERVIEW'!$D$15", "TotalBudget");
}
