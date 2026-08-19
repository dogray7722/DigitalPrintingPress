// Verifies the OVERVIEW sheet's B15 currency dropdown: a list-type data validation
// pointing at a hidden helper column (R) containing "CODE (symbol)" options.
// Run: node scripts/verify-currency-dropdown.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import { rmSync } from 'node:fs'

const tmpModPath = new URL('./.overview.tmp.mjs', import.meta.url)
await build({
  entryPoints: ['src/lib/excel/sheets/overview.ts'],
  bundle: true,
  format: 'esm',
  outfile: new URL(tmpModPath).pathname,
  platform: 'node',
  external: ['exceljs'],
})
const { buildOverviewSheet } = await import(tmpModPath.href)
rmSync(new URL(tmpModPath).pathname)

const { CURRENCIES } = await (async () => {
  const tmp = new URL('./.currencies.tmp.mjs', import.meta.url)
  await build({
    entryPoints: ['src/data/currencies.ts'],
    bundle: true,
    format: 'esm',
    outfile: new URL(tmp).pathname,
  })
  const mod = await import(tmp.href)
  rmSync(new URL(tmp).pathname)
  return mod
})()

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
  sheets: { budgetTracker: true, itinerary: true, flights: true, hotels: true, restaurants: true, excursions: true, packingList: true, tasks: true, events: false },
}

const ts = {
  fontName: 'Calibri',
  sizes: { title: 18, sectionHeader: 12, header: 12, data: 11 },
  palette: {
    primary: 'FFE91E63',
    primaryText: 'FFFFFFFF',
    secondary: 'FFFCE4EC',
    secondaryText: 'FF2D1B6B',
    lightBg: 'FFFFF5F7',
    mediumBg: 'FFFCE4EC',
    border: 'FFE0C8D0',
  },
  numFmtCurrency: '"$"#,##0.00',
  numFmtDate: 'dd mmm yyyy',
  numFmtPercent: '0%',
}

const wb = new ExcelJS.Workbook()
buildOverviewSheet(wb, state, ts)
const ws = wb.getWorksheet('OVERVIEW')

const b15 = ws.getCell('B15')
assert(b15.value === 'USD ($)', `B15 initial value is "USD ($)" (got "${b15.value}")`)

const dv = ws.dataValidations.find('B15')
assert(!!dv, 'B15 has a data validation')
assert(dv?.type === 'list', `B15 data validation type is "list" (got "${dv?.type}")`)
assert(
  dv?.formulae?.[0] === `$R$1:$R${CURRENCIES.length}`,
  `B15 data validation formula is "$R$1:$R${CURRENCIES.length}" (got "${dv?.formulae?.[0]}")`
)

CURRENCIES.forEach((c, i) => {
  const cell = ws.getCell(`R${i + 1}`)
  const expected = `${c.code} (${c.symbol})`
  assert(cell.value === expected, `R${i + 1} = "${expected}" (got "${cell.value}")`)
})

assert(ws.getColumn('R').hidden === true, 'Column R is hidden')

// Round-trip through a real xlsx buffer to make sure nothing corrupts on write.
const buf = await wb.xlsx.writeBuffer()
const wb2 = new ExcelJS.Workbook()
await wb2.xlsx.load(buf)
const ws2 = wb2.getWorksheet('OVERVIEW')
const dv2 = ws2.dataValidations.find('B15')
assert(dv2?.type === 'list', 'After round-trip: B15 still has a list data validation')
assert(ws2.getCell('R1').value === 'USD ($)', 'After round-trip: R1 = "USD ($)"')
assert(ws2.getColumn('R').hidden === true, 'After round-trip: column R is hidden')

if (process.exitCode) {
  console.error('\nSome checks failed.')
} else {
  console.log('\nAll checks passed.')
}
