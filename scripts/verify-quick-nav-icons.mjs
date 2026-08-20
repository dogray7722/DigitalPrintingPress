// Verifies the OVERVIEW "JUMP TO" strip: every enabled sheet gets one internal-hyperlink
// row prefixed with its SHEET_ICONS emoji, the old "→" arrow is gone, and the
// astral-plane emoji survive the round trip into xl/worksheets/sheet1.xml as UTF-8 —
// both as the cell's own text and inside the hyperlink's `display` attribute.
//
// These are NATIVE internal hyperlinks, not HYPERLINK() formulas, so the targets are
// asserted against the raw XML: ExcelJS's reader only binds a hyperlink to its cell when
// the element carries an r:id (worksheet-xform.js `if (hyperlink.rId)`), and Excel's own
// r:id-less internal form therefore never surfaces on a loaded cell.
// Run: node scripts/verify-quick-nav-icons.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
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

// The cell's own text still round-trips through ExcelJS (that's what a link-ignoring
// consumer renders); the link target lives in the XML, checked below.
EXPECTED.forEach(([, icon, label], i) => {
  const row = 39 + i
  const v = ws.getCell(`B${row}`).value
  const text = typeof v === 'object' && v ? v.text : v
  assert(text === `${icon}  ${label}`, `B${row} label is "${icon}  ${label}"`)
})

// No arrow left anywhere in the strip.
for (let row = 38; row < 39 + EXPECTED.length; row++) {
  const v = ws.getCell(`B${row}`).value
  const text = typeof v === 'object' && v ? String(v.text ?? '') : String(v ?? '')
  assert(!text.includes('→'), `B${row} has no "→" arrow`)
}

// ── Native internal-hyperlink shape in the raw XML ────────────────────────────
// Excel's own form is `location` + `display` with NO r:id. Without `display`, Google
// Sheets renders the link text as a literal "#gid=xxxx" instead of the label.
const zip = await JSZip.loadAsync(buffer)
const overviewXml = await zip.file('xl/worksheets/sheet1.xml').async('string')
const tags = [...overviewXml.matchAll(/<hyperlink\b[^>]*\/>/g)].map((m) => m[0])

assert(
  tags.length === EXPECTED.length,
  `${EXPECTED.length} quick-nav hyperlinks on OVERVIEW, got ${tags.length}`
)

EXPECTED.forEach(([sheet, icon, label], i) => {
  const row = 39 + i
  const tag = tags.find((t) => t.includes(`ref="B${row}"`))
  assert(!!tag, `B${row} has a hyperlink element`)
  if (!tag) return
  assert(tag.includes(`location="&apos;${sheet}&apos;!A1"`), `B${row} links to ${sheet}`)
  // The emoji must survive into the display attribute, not just the cell text.
  assert(tag.includes(`display="${icon}  ${label}"`), `B${row} display is "${icon}  ${label}"`)
  assert(!/r:id=/.test(tag), `B${row} carries no r:id`)
})

// OVERVIEW is the sheet chartInjection rewrites, and the cover-photo drawing rel sits
// AFTER the hyperlink rels. Deleting hyperlink rels must not orphan it.
const rels = (await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('string')) ?? ''
assert(!/relationships\/hyperlink/.test(rels), 'no leftover hyperlink relationships on OVERVIEW')
for (const rid of new Set([...overviewXml.matchAll(/r:id="(rId\d+)"/g)].map((m) => m[1]))) {
  assert(rels.includes(`Id="${rid}"`), `OVERVIEW r:id ${rid} still resolves (drawing intact)`)
}

// The row after the last link must be untouched by the strip.
assert(ws.getCell(`B${39 + EXPECTED.length}`).value == null, 'strip ends after the last enabled sheet')

console.log(process.exitCode ? '\nFAILED' : '\nAll quick-nav icon checks passed')
