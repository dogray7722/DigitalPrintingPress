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
  wb.definedNames.add("'OVERVIEW'!$D$3", "TripStart");
  wb.definedNames.add("'OVERVIEW'!$D$4", "TripEnd");
  wb.definedNames.add("'OVERVIEW'!$D$9", "NumAdults");
  wb.definedNames.add("'OVERVIEW'!$D$14", "TotalBudget");
}
