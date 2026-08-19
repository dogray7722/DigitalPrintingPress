import ExcelJS from 'exceljs'
import type { WizardState } from '../../types/wizard'
import type { Recommendations, ExchangeRate } from '../recommendations/types'
import { configureWorkbook, addNamedRanges } from './workbookConfig'
import { getThemeStyles } from './styleFactory'
import { injectChart, CHART_PALETTES } from './chartInjection'
import { injectCalcChain } from './calcChain'
import { protectWorkbook } from './protection'
import { buildOverviewSheet, categoryBudgetAmounts } from './sheets/overview'
import { buildBudgetTrackerSheet } from './sheets/budgetTracker'
import { buildItinerarySheet } from './sheets/itinerary'
import { buildHotelsSheet } from './sheets/hotels'
import { buildRestaurantsSheet } from './sheets/restaurants'
import { buildExcursionsSheet } from './sheets/excursions'
import { buildPackingListSheet } from './sheets/packingList'
import { buildFlightsSheet } from './sheets/flights'
import { buildTasksSheet } from './sheets/tasks'
import { buildEventsSheet } from './sheets/events'
import { buildInstructionsSheet } from './sheets/instructions'

export async function generateWorkbook(
  state: WizardState,
  recommendations?: Recommendations,
  exchangeRate?: ExchangeRate
): Promise<Blob> {
  const wb = new ExcelJS.Workbook()
  const ts = getThemeStyles(state)

  configureWorkbook(wb, state)

  // OVERVIEW is always first and always included
  buildOverviewSheet(wb, state, ts, exchangeRate, recommendations)

  // Optional sheets in logical order, with recommendations prefill
  if (state.sheets.itinerary) buildItinerarySheet(wb, state, ts, recommendations?.itinerary)
  if (state.sheets.flights) buildFlightsSheet(wb, state, ts)
  if (state.sheets.hotels) buildHotelsSheet(wb, state, ts, recommendations?.regions)
  if (state.sheets.restaurants) buildRestaurantsSheet(wb, state, ts, recommendations?.regions)
  if (state.sheets.excursions) buildExcursionsSheet(wb, state, ts, recommendations?.regions)
  if (state.sheets.budgetTracker) buildBudgetTrackerSheet(wb, state, ts)
  if (state.sheets.packingList) buildPackingListSheet(wb, state, ts)
  if (state.sheets.tasks) buildTasksSheet(wb, state, ts)

  // Events sheet only when recommendations are enabled AND toggled on AND we have data
  if (state.useRecommendations && state.sheets.events && recommendations?.events?.length) {
    buildEventsSheet(wb, state, ts, recommendations.events)
  }

  // Named ranges must be added AFTER OVERVIEW sheet is populated
  addNamedRanges(wb)

  // INSTRUCTIONS is always last
  buildInstructionsSheet(wb, state, ts)

  // Lock formula/total cells and hide their formulas; must run before writeBuffer()
  // since protection is an in-memory worksheet property, not a buffer post-process.
  await protectWorkbook(wb)

  let buffer: ArrayBuffer = (await wb.xlsx.writeBuffer()) as ArrayBuffer

  // ExcelJS can't write native charts — inject a real chart bound to the OVERVIEW
  // budget table so it updates live as the user logs actual spending. Always a
  // stacked bar chart (spent / remaining / over-budget per category).
  // Cached data points for the chart series — must mirror what the referenced cells
  // hold at generation time (i.e. the cached formula results with empty trackers).
  // Without these, Excel's chart redraw runs one edit behind the cells.
  const cats = categoryBudgetAmounts(state)

  buffer = await injectChart(buffer, {
    kind: 'bar',
    sheetName: 'OVERVIEW',
    categoryRange: "'OVERVIEW'!$F$18:$F$23",
    categoryLabels: cats.map(c => c.key),
    // col N = MIN(actual, budget) = 0 with empty trackers.
    values: cats.map(() => 0),
    // col O = MAX(budget - actual, 0) = budget; col P = MAX(actual - budget, 0) = 0.
    remainderValues: cats.map(c => c.amount),
    overValues: cats.map(() => 0),
    // col N = MIN(actual, budget) — colored "spent" segment (base of the stacked bar).
    valueRange: "'OVERVIEW'!$N$18:$N$23",
    // col O = MAX(budget - actual, 0) — light "remaining" segment stacked after
    // col N so each category is one bar whose length = budget.
    remainderRange: "'OVERVIEW'!$O$18:$O$23",
    // col P = MAX(actual - budget, 0) — solid red overage segment stacked after
    // col O, extending the bar past the budget length when a category is overspent.
    overRange: "'OVERVIEW'!$P$18:$P$23",
    anchor: { fromCol: 5, fromRow: 26, toCol: 11, toRow: 44 },
    colors: CHART_PALETTES[state.theme],
    title: `${state.destination || 'Trip'} Budget Breakdown`,
  })

  // Trip-readiness doughnut — ready vs. to-do across PACKING + TASKS. Fills the
  // lower-left. The doughnut auto-labels category + percent (see chartInjection dLbls),
  // so no center-cell label is needed. Skipped when neither source sheet exists.
  if (state.sheets.packingList || state.sheets.tasks) {
    buffer = await injectChart(buffer, {
      kind: 'doughnut',
      sheetName: 'OVERVIEW',
      categoryRange: "'OVERVIEW'!$T$1:$T$2",
      categoryLabels: ['Ready', 'To do'],
      valueRange: "'OVERVIEW'!$S$1:$S$2", // S1 ready, S2 remaining
      values: [0, 1], // fresh workbook = 0% ready
      anchor: { fromCol: 0, fromRow: 23, toCol: 4, toRow: 39 }, // A24:E40
      colors: [ts.palette.secondary, ts.palette.mediumBg],
      // No title/legend/per-slice labels — the "TRIP READINESS" heading and % are
      // in-cell (overview.ts), floating over the doughnut's transparent hole.
      legend: false,
      dataLabels: false,
    })
  }

  // ExcelJS also never writes the calcChain part; without it Excel for Mac
  // repaints the injected chart one calculation behind the cells.
  buffer = await injectCalcChain(buffer)

  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}
