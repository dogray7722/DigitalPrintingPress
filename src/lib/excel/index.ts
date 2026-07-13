import ExcelJS from 'exceljs'
import type { WizardState } from '../../types/wizard'
import type { Recommendations, ExchangeRate } from '../recommendations/types'
import { configureWorkbook, addNamedRanges } from './workbookConfig'
import { getThemeStyles } from './styleFactory'
import { injectChart, CHART_PALETTES } from './chartInjection'
import type { ChartKind } from './chartInjection'
import { injectCalcChain } from './calcChain'
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
  buildOverviewSheet(wb, state, ts, exchangeRate)

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

  let buffer: ArrayBuffer = (await wb.xlsx.writeBuffer()) as ArrayBuffer

  // ExcelJS can't write native charts — inject a real chart bound to the OVERVIEW
  // budget table so it updates live as the user logs actual spending. Every chart
  // style is now a native injected chart (bar / pie / donut).
  const CHART_KIND: Record<WizardState['chartStyle'], ChartKind> = {
    bar: 'bar',
    pie: 'pie',
    donut: 'doughnut',
  }
  // Cached data points for the chart series — must mirror what the referenced cells
  // hold at generation time (i.e. the cached formula results with empty trackers).
  // Without these, Excel's chart redraw runs one edit behind the cells.
  const cats = categoryBudgetAmounts(state)

  buffer = await injectChart(buffer, {
    kind: CHART_KIND[state.chartStyle],
    sheetName: 'OVERVIEW',
    categoryRange: "'OVERVIEW'!$F$18:$F$23",
    categoryLabels: cats.map(c => c.key),
    // bar: col N = MIN(actual, budget) = 0 with empty trackers.
    // pie/donut: col M = IF(actual>0, actual, budget) = budget estimate.
    values: state.chartStyle === 'bar' ? cats.map(() => 0) : cats.map(c => c.amount),
    // bar only: col O = MAX(budget - actual, 0) = budget; col P = MAX(actual - budget, 0) = 0.
    remainderValues: state.chartStyle === 'bar' ? cats.map(c => c.amount) : undefined,
    overValues: state.chartStyle === 'bar' ? cats.map(() => 0) : undefined,
    // bar: col N = MIN(actual, budget) — colored "spent" segment (base of the stacked bar).
    // pie/donut: col M = actual or budget estimate fallback so chart is never empty.
    valueRange: state.chartStyle === 'bar'
      ? "'OVERVIEW'!$N$18:$N$23"
      : "'OVERVIEW'!$M$18:$M$23",
    // bar only: col O = MAX(budget - actual, 0) — light "remaining" segment stacked after
    //   col N so each category is one bar whose length = budget.
    remainderRange: state.chartStyle === 'bar' ? "'OVERVIEW'!$O$18:$O$23" : undefined,
    // bar only: col P = MAX(actual - budget, 0) — solid red overage segment stacked after
    //   col O, extending the bar past the budget length when a category is overspent.
    overRange: state.chartStyle === 'bar' ? "'OVERVIEW'!$P$18:$P$23" : undefined,
    anchor: { fromCol: 5, fromRow: 26, toCol: 11, toRow: 44 },
    colors: CHART_PALETTES[state.theme],
    title: `${state.destination || 'Trip'} Budget Breakdown`,
  })

  // ExcelJS also never writes the calcChain part; without it Excel for Mac
  // repaints the injected chart one calculation behind the cells.
  buffer = await injectCalcChain(buffer)

  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}
