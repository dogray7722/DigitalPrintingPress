// Verifies the OVERVIEW "JUMP TO" strip: every enabled sheet gets one HYPERLINK row
// prefixed with its SHEET_ICONS emoji, the old "→" arrow is gone, and the astral-plane
// emoji survive the round trip through the formula string literal AND the cached
// result (ExcelJS writes both into xl/worksheets/sheet1.xml as UTF-8).
// Run: node scripts/verify-quick-nav-icons.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import { rmSync } from 'node:fs'

const tmpModPath = new URL('./.quicknav.tmp.mjs', import.meta.url)
await build({
  entryPoints: ['src/lib/excel/index.ts'],
  bundle: true,
  format: 'esm',
  outfile: new URL(tmpModPath).pathname,
  platform: 'node',
  external: ['exceljs', 'jszip', 'date-fns'],
})
const { generateWorkbook } = await import(tmpModPath.href)
rmSync(new URL(tmpModPath).pathname)

function assert(cond, msg) {
  if (!cond) {
    console.error('❌ FAIL:', msg)
    process.exitCode = 1
  } else {
    console.log('✓', msg)
  }
}

const state = {
  destination: 'Lisbon',
  useRecommendations: false,
  startDate: '2026-07-01',
  duration: 7,
  partyType: 'couple',
  partySize: 2,
  travelMonth: 7,
  overviewImage: null,
  currency: 'USD',
  tripStyle: 'midrange',
  budgets: { flights: 1040, hotel: 195, food: 78, activities: 65, transport: 39, shopping: 200, misc: 52 },
  theme: 'sakura',
  accentColor: '',
  chartStyle: 'bar',
  fontFamily: 'sans',
  fontSize: 'medium',
  sheets: {
    budgetTracker: true,
    itinerary: true,
    flights: true,
    hotels: true,
    restaurants: true,
    excursions: true,
    packingList: true,
    tasks: true,
    events: false,
  },
}

// Expected in buildQuickNav's link order (overview.ts), icons from SHEET_ICONS.
const EXPECTED = [
  ['ITINERARY', '📅', 'Itinerary'],
  ['TRANSPORTATION', '✈️', 'Transportation'],
  ['ACCOMMODATION', '🏨', 'Accommodation'],
  ['DINING', '🍽️', 'Dining'],
  ['EXCURSIONS', '🎯', 'Excursions'],
  ['PACKING LIST', '🎒', 'Packing List'],
  ['TASKS', '✅', 'Tasks'],
  ['OTHER EXPENSES', '💰', 'Other Expenses'],
]

const blob = await generateWorkbook(state)
const buffer = Buffer.from(await blob.arrayBuffer())

// Read back through ExcelJS so we exercise the real xlsx parse, not a string grep.
const wb = new ExcelJS.Workbook()
await wb.xlsx.load(buffer)
const ws = wb.getWorksheet('OVERVIEW')
assert(!!ws, 'OVERVIEW sheet exists')

assert(ws.getCell('B38').value === 'JUMP TO', 'JUMP TO header still at B38')

EXPECTED.forEach(([sheet, icon, label], i) => {
  const row = 39 + i
  const cell = ws.getCell(`B${row}`)
  const v = cell.value
  const formula = typeof v === 'object' && v ? v.formula : undefined
  const result = typeof v === 'object' && v ? v.result : undefined
  assert(
    typeof formula === 'string' && formula.includes(`#'${sheet}'!A1`),
    `B${row} links to ${sheet}`
  )
  assert(
    typeof formula === 'string' && formula.includes(`"${icon}  ${label}"`),
    `B${row} formula display text is "${icon}  ${label}"`
  )
  assert(result === `${icon}  ${label}`, `B${row} cached result is "${icon}  ${label}"`)
})

// No arrow left anywhere in the strip.
for (let row = 38; row < 39 + EXPECTED.length; row++) {
  const v = ws.getCell(`B${row}`).value
  const text = typeof v === 'object' && v ? `${v.formula ?? ''}${v.result ?? ''}` : String(v ?? '')
  assert(!text.includes('→'), `B${row} has no "→" arrow`)
}

// The row after the last link must be untouched by the strip.
assert(ws.getCell(`B${39 + EXPECTED.length}`).value == null, 'strip ends after the last enabled sheet')

console.log(process.exitCode ? '\nFAILED' : '\nAll quick-nav icon checks passed')
