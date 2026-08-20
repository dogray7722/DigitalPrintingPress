// Verifies the PACKING LIST "Packed?" (C) and TASKS "Done?" (D) check-mark dropdowns:
// a single-option list data validation whose literal is exactly "✓" — the same literal
// the sheets' COUNTIF progress formulas and OVERVIEW's readiness ring count. A mismatch
// (or a mangled non-ASCII round-trip) would leave the progress bars stuck at 0%.
// Run: node scripts/verify-check-dropdowns.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { rmSync } from 'node:fs'

async function loadSheet(entry, name) {
  const tmp = new URL(`./.${name}.tmp.mjs`, import.meta.url)
  await build({
    entryPoints: [entry],
    bundle: true,
    format: 'esm',
    outfile: new URL(tmp).pathname,
    platform: 'node',
    external: ['exceljs'],
  })
  const mod = await import(tmp.href)
  rmSync(new URL(tmp).pathname)
  return mod
}

const { buildPackingListSheet } = await loadSheet(
  'src/lib/excel/sheets/packingList.ts',
  'packing'
)
const { buildTasksSheet } = await loadSheet('src/lib/excel/sheets/tasks.ts', 'tasks')

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
buildPackingListSheet(wb, state, ts)
buildTasksSheet(wb, state, ts)

// Pre-write, ExcelJS keys validations by the exact sqref string it was given; only its
// reader expands a range into per-cell entries. Hence `range` here and `probe` below.
const cases = [
  { sheet: 'PACKING LIST', range: 'C5:C500', probe: 'C5', progressCell: 'A2' },
  { sheet: 'TASKS', range: 'D5:D500', probe: 'D5', progressCell: 'A2' },
]

for (const { sheet, range, progressCell } of cases) {
  const ws = wb.getWorksheet(sheet)
  const dv = ws.dataValidations.find(range)
  assert(!!dv, `${sheet} ${range} has a data validation`)
  assert(dv?.type === 'list', `${sheet} ${range} validation type is "list" (got "${dv?.type}")`)
  assert(
    dv?.formulae?.[0] === '"✓"',
    `${sheet} ${range} option list is exactly "✓" (got ${JSON.stringify(dv?.formulae?.[0])})`
  )
  assert(dv?.allowBlank === true, `${sheet} ${range} allows blank (so a row can be un-checked)`)

  // The dropdown literal must match what the progress formula counts.
  const formula = ws.getCell(progressCell).value?.formula ?? ''
  assert(
    formula.includes(`${range.split(':')[0]}:${range.split(':')[1]},"✓"`),
    `${sheet} ${progressCell} COUNTIF counts "✓" over ${range}`
  )
}

// Round-trip through a real xlsx buffer: the ✓ is non-ASCII, so confirm it survives
// XML write/read rather than trusting the in-memory model.
const buf = await wb.xlsx.writeBuffer()
const wb2 = new ExcelJS.Workbook()
await wb2.xlsx.load(buf)
for (const { sheet, probe } of cases) {
  const dv2 = wb2.getWorksheet(sheet).dataValidations.find(probe)
  assert(
    dv2?.type === 'list' && dv2?.formulae?.[0] === '"✓"',
    `After round-trip: ${sheet} ${probe} list option is still "✓"`
  )
}

// And assert against the raw sheet XML — that is what Excel and Google Sheets read.
const zip = await JSZip.loadAsync(buf)
const sheetFiles = Object.keys(zip.files).filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f))
let seen = 0
for (const f of sheetFiles) {
  const xml = await zip.file(f).async('string')
  const m = xml.match(/<dataValidation[^>]*sqref="(C5:C500|D5:D500)"[^>]*>[\s\S]*?<\/dataValidation>/)
  if (!m) continue
  seen++
  // ExcelJS XML-escapes the wrapping quotes (&quot;) — the same shape Excel itself
  // writes for an inline list — so accept either literal or escaped quotes.
  assert(
    /<formula1>(?:"|&quot;)✓(?:"|&quot;)<\/formula1>/.test(m[0]),
    `${f}: sqref ${m[1]} carries the "✓" option in <formula1> in the raw XML`
  )
  assert(/type="list"/.test(m[0]), `${f}: sqref ${m[1]} dataValidation is type="list"`)
}
assert(seen === 2, `Both check-mark dropdowns present in the raw XML (found ${seen})`)

if (process.exitCode) {
  console.error('\nSome checks failed.')
} else {
  console.log('\nAll checks passed.')
}
