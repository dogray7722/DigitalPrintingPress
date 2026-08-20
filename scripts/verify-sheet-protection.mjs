// Verifies worksheet protection end-to-end: generates a full workbook via the real
// generateWorkbook() pipeline, then inspects the raw XML to confirm protected sheets
// carry <sheetProtection>, spot-checked editable cells are unlocked, spot-checked
// formula/total cells are hidden, and INSTRUCTIONS stays fully editable. Also
// confirms chart + calcChain injection still succeed unmodified.
// Run: node scripts/verify-sheet-protection.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { rmSync } from 'node:fs'

const tmpModPath = new URL('./.index.tmp.mjs', import.meta.url)
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

const blob = await generateWorkbook(state)
const buffer = Buffer.from(await blob.arrayBuffer())

const zip = await JSZip.loadAsync(buffer)
const wbXml = await zip.file('xl/workbook.xml').async('string')
const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
const stylesXml = await zip.file('xl/styles.xml').async('string')

function sheetXmlFor(name) {
  const re = new RegExp(`<sheet[^>]*name="${name}"[^>]*r:id="(rId\\d+)"`)
  const m = re.exec(wbXml)
  if (!m) throw new Error(`sheet "${name}" not found in workbook.xml`)
  const target = new RegExp(`Id="${m[1]}"[^>]*Target="([^"]+)"`).exec(wbRels)[1]
  const path = 'xl/' + target.replace(/^\//, '').replace(/^xl\//, '')
  return path
}

async function loadSheetXml(name) {
  return zip.file(sheetXmlFor(name)).async('string')
}

// cellXfs entries, in document order (index = style index `s`).
const cellXfsBlock = /<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/.exec(stylesXml)[1]
const cellXfs = cellXfsBlock.match(/<xf\b[^>]*?\/>|<xf\b[^>]*?>[\s\S]*?<\/xf>/g) || []

function protectionOf(sheetXml, ref) {
  const cellTagMatch = new RegExp(`<c r="${ref}"[^>]*>`).exec(sheetXml)
  const cellTag = cellTagMatch ? cellTagMatch[0] : ''
  const sMatch = /\ss="(\d+)"/.exec(cellTag)
  const xfIndex = sMatch ? Number(sMatch[1]) : 0
  const xf = cellXfs[xfIndex] ?? ''
  const prot = /<protection([^>]*)\/>/.exec(xf)
  const attrs = prot ? prot[1] : ''
  return {
    locked: !/locked="0"/.test(attrs), // OOXML default: locked
    hidden: /hidden="1"/.test(attrs), // OOXML default: not hidden
  }
}

function assertUnlocked(sheetXml, sheetName, ref) {
  const { locked } = protectionOf(sheetXml, ref)
  assert(!locked, `${sheetName}!${ref} is unlocked (editable)`)
}

function assertHidden(sheetXml, sheetName, ref) {
  const { hidden } = protectionOf(sheetXml, ref)
  assert(hidden, `${sheetName}!${ref} formula is hidden`)
}

function assertProtected(sheetXml, sheetName) {
  assert(/<sheetProtection\b/.test(sheetXml), `${sheetName} has <sheetProtection>`)
}

function assertUnprotected(sheetXml, sheetName) {
  assert(!/<sheetProtection\b/.test(sheetXml), `${sheetName} has no <sheetProtection> (fully editable)`)
}

// ── OVERVIEW ──────────────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('OVERVIEW')
  assertProtected(xml, 'OVERVIEW')
  assertUnlocked(xml, 'OVERVIEW', 'B6') // TripStart
  assertUnlocked(xml, 'OVERVIEW', 'D6') // TripEnd
  assertUnlocked(xml, 'OVERVIEW', 'E9') // NumAdults
  assertUnlocked(xml, 'OVERVIEW', 'B16') // Currency
  assertUnlocked(xml, 'OVERVIEW', 'H15') // Estimated Budget, first category row
  assertHidden(xml, 'OVERVIEW', 'D16') // TotalBudget
  assertHidden(xml, 'OVERVIEW', 'I15') // Actual Spent, first category row
}

// ── ITINERARY — near-open: only the auto-Date column B stays locked ──────────
{
  const xml = await loadSheetXml('ITINERARY')
  // <sheetProtection> must stay on: a cell's `locked` flag is inert without it.
  assertProtected(xml, 'ITINERARY')
  for (const col of ['A', 'C', 'D', 'E', 'F', 'G', 'H', 'I']) {
    assertUnlocked(xml, 'ITINERARY', `${col}4`) // first data row
    assertUnlocked(xml, 'ITINERARY', `${col}33`) // last data row (MAX_DAYS = 30)
  }
  assert(protectionOf(xml, 'B4').locked, 'ITINERARY!B4 (auto-Date) stays locked')
  assert(protectionOf(xml, 'B33').locked, 'ITINERARY!B33 (auto-Date) stays locked')
  assertHidden(xml, 'ITINERARY', 'B4')
  // Sheet-level restrictions are relaxed so it behaves like an unprotected sheet.
  const sp = /<sheetProtection[^>]*\/?>/.exec(xml)[0]
  for (const attr of ['formatCells', 'formatRows', 'insertRows', 'deleteRows', 'sort']) {
    assert(new RegExp(`${attr}="0"`).test(sp), `ITINERARY sheetProtection allows ${attr}`)
  }
}

// ── TRANSPORTATION ───────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('TRANSPORTATION')
  assertProtected(xml, 'TRANSPORTATION')
  assertUnlocked(xml, 'TRANSPORTATION', 'A6')
  assertUnlocked(xml, 'TRANSPORTATION', 'H6')
  assertHidden(xml, 'TRANSPORTATION', 'H3')
}

// ── ACCOMMODATION ─────────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('ACCOMMODATION')
  assertProtected(xml, 'ACCOMMODATION')
  assertUnlocked(xml, 'ACCOMMODATION', 'A6')
  assertHidden(xml, 'ACCOMMODATION', 'E3')
}

// ── DINING ────────────────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('DINING')
  assertProtected(xml, 'DINING')
  assertUnlocked(xml, 'DINING', 'A6')
  assertHidden(xml, 'DINING', 'G3')
}

// ── EXCURSIONS ────────────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('EXCURSIONS')
  assertProtected(xml, 'EXCURSIONS')
  assertUnlocked(xml, 'EXCURSIONS', 'A6')
  assertUnlocked(xml, 'EXCURSIONS', 'F6')
  assertHidden(xml, 'EXCURSIONS', 'H6') // per-row cost formula — must stay locked+hidden
  assertHidden(xml, 'EXCURSIONS', 'H3')
}

// ── OTHER EXPENSES ────────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('OTHER EXPENSES')
  assertProtected(xml, 'OTHER EXPENSES')
  assertUnlocked(xml, 'OTHER EXPENSES', 'A8')
  assertHidden(xml, 'OTHER EXPENSES', 'D3')
}

// ── PACKING LIST ──────────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('PACKING LIST')
  assertProtected(xml, 'PACKING LIST')
  // Row 5 is the first category separator header; row 6 is the first actual item row.
  assertUnlocked(xml, 'PACKING LIST', 'C6')
  assertHidden(xml, 'PACKING LIST', 'A2')
}

// ── TASKS ─────────────────────────────────────────────────────────────────────
{
  const xml = await loadSheetXml('TASKS')
  assertProtected(xml, 'TASKS')
  assertUnlocked(xml, 'TASKS', 'D5')
  assertHidden(xml, 'TASKS', 'A2')
}

// ── INSTRUCTIONS — never protected ───────────────────────────────────────────
{
  const xml = await loadSheetXml('INSTRUCTIONS')
  assertUnprotected(xml, 'INSTRUCTIONS')
}

// ── Chart + calcChain injection still intact ─────────────────────────────────
assert(zip.file('xl/charts/chart1.xml'), 'chart1.xml still injected')
assert(zip.file('xl/calcChain.xml'), 'calcChain.xml still injected')

// ── Round-trip: ExcelJS re-opens the protected workbook without throwing ────
try {
  const wb2 = new ExcelJS.Workbook()
  await wb2.xlsx.load(buffer)
  const ws2 = wb2.getWorksheet('OVERVIEW')
  assert(ws2 && ws2.getCell('B6').value instanceof Date, 'ExcelJS re-opens protected workbook intact')
} catch (e) {
  assert(false, 'ExcelJS re-open threw: ' + e.message)
}

if (process.exitCode) {
  console.error('\nSome checks failed.')
} else {
  console.log('\nAll checks passed.')
}
