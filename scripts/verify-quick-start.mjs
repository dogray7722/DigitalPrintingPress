// Verifies the QUICK START tab:
//   • it is the workbook's FIRST sheet (tab order = add order in index.ts)
//   • its "where to go next" links are NATIVE internal hyperlinks — location + display,
//     no r:id, no leftover external hyperlink relationships (the only shape Google Sheets
//     imports correctly; see nativeHyperlinks.ts)
//   • the recommendation step appears only when the trackers actually got a guide, and
//     names exactly the tabs whose "★ Recommendations →" chip exists
//   • pushing QUICK START in front of OVERVIEW didn't disturb OVERVIEW's chart/drawing —
//     chartInjection resolves its sheet by name, not by part index
// Run: node scripts/verify-quick-start.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { rmSync, writeFileSync } from 'node:fs'

const tmpModPath = new URL('./.quickstart.tmp.mjs', import.meta.url)
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

const baseState = {
  destination: 'Lisbon',
  useRecommendations: false,
  startDate: '2026-07-01',
  duration: 7,
  partyType: 'couple',
  partySize: 2,
  travelMonth: 7,
  overviewImage: null,
  currency: 'USD',
  destinationCurrency: '',
  tripStyle: 'midrange',
  budgets: { flights: 1040, hotel: 195, food: 78, activities: 65, transport: 39, shopping: 200, misc: 52 },
  theme: 'sakura',
  accentColor: '',
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

// Regions with hotels + excursions but deliberately NO restaurants: DINING gets no guide
// and no chip, so QUICK START must not advertise one.
const recommendations = {
  regions: [
    {
      region: 'Alfama',
      description: 'The old Moorish quarter, all staircases and viewpoints.',
      hotels: [{ name: 'Casa do Fado', price: '€140/night', rating: '★★★★' }],
      restaurants: [],
      excursions: [{ name: 'São Jorge Castle', duration: '2 hrs', price: '€15' }],
    },
  ],
  itinerary: [{ day: 1, location: 'Alfama', morning: 'Arrive and settle in' }],
  events: [],
}

async function partsOf(buffer) {
  const zip = await JSZip.loadAsync(buffer)
  const wbXml = await zip.file('xl/workbook.xml').async('string')
  const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
  const pathFor = (name) => {
    const rId = new RegExp(`<sheet[^>]*name="${name}"[^>]*r:id="(rId\\d+)"`).exec(wbXml)?.[1]
    if (!rId) throw new Error(`sheet "${name}" not found in workbook.xml`)
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(wbRels)[1]
    return 'xl/' + target.replace(/^\//, '').replace(/^xl\//, '')
  }
  return { zip, wbXml, pathFor }
}

// ── With recommendations ──────────────────────────────────────────────────────
{
  const state = { ...baseState, useRecommendations: true }
  const blob = await generateWorkbook(state, recommendations)
  const buffer = Buffer.from(await blob.arrayBuffer())
  // Layout (row heights, wrapping, the badge gutter) can only really be judged by eye —
  // merged rows don't auto-grow in Excel desktop, so open this one there, not in Numbers.
  writeFileSync('/tmp/quick-start-test.xlsx', buffer)

  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer)

  // Tab order — QUICK START first, INSTRUCTIONS still last.
  const names = wb.worksheets.map((w) => w.name)
  assert(names[0] === 'QUICK START', `QUICK START is the first tab (got "${names[0]}")`)
  assert(names[1] === 'OVERVIEW', `OVERVIEW is still second (got "${names[1]}")`)
  assert(names[names.length - 1] === 'INSTRUCTIONS', 'INSTRUCTIONS is still the last tab')

  const ws = wb.getWorksheet('QUICK START')
  assert(ws.getCell('A1').value === 'QUICK START  —  LISBON', 'title carries the destination')

  // Every populated cell's text, for content assertions.
  const text = []
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      const v = cell.value
      text.push(typeof v === 'object' && v ? String(v.text ?? v.formula ?? '') : String(v ?? ''))
    })
  })
  const all = text.join('\n')

  // Step numbering is auto-incremented — with recs on, all five steps render.
  const badges = text.filter((t) => /^[1-9]$/.test(t)).map(Number)
  assert(
    JSON.stringify(badges) === JSON.stringify([1, 2, 3, 4, 5]),
    `steps numbered 1..5 with no gaps (got ${badges.join(',')})`
  )

  // The cells the copy points at must be the real ones.
  assert(/cell B6/.test(all) && /cell D6/.test(all), 'names TripStart B6 and TripEnd D6')
  assert(/H15 to H20/.test(all), 'names the Estimated Budget column H15:H20')
  assert(/Total Budget \(D16\)/.test(all), 'names TotalBudget D16')
  assert(/E9/.test(all), 'names NumAdults E9')
  assert(/Date column \(B\)/.test(all), 'flags ITINERARY column B as the locked one')

  // Recommendation step: names the chip label and the two tabs that actually got a
  // guide, and NOT the one that did not.
  assert(/★  Recommendations  →/.test(all), 'quotes the chip label verbatim')
  assert(/ACCOMMODATION —/.test(all), 'describes the ACCOMMODATION guide')
  assert(/EXCURSIONS —/.test(all), 'describes the EXCURSIONS guide')
  assert(!/DINING —/.test(all), 'does NOT advertise a DINING guide (no restaurants returned)')
  assert(/preview card/.test(all), 'warns about the Google Sheets two-click link chip')

  // ── Native internal-hyperlink shape ────────────────────────────────────────
  const { zip, pathFor } = await partsOf(buffer)
  const qsPath = pathFor('QUICK START')
  const qsXml = await zip.file(qsPath).async('string')
  const tags = [...qsXml.matchAll(/<hyperlink\b[^>]*\/>/g)].map((m) => m[0])
  assert(tags.length === 3, `3 jump links on QUICK START, got ${tags.length}`)
  for (const [sheet, label] of [
    ['OVERVIEW', 'OVERVIEW'],
    ['ITINERARY', 'ITINERARY'],
    ['INSTRUCTIONS', 'INSTRUCTIONS'],
  ]) {
    const tag = tags.find((t) => t.includes(`location="&apos;${sheet}&apos;!A1"`))
    assert(!!tag, `a link targets '${sheet}'!A1`)
    if (!tag) continue
    assert(tag.includes(`display="${label}"`), `${sheet} link carries display="${label}"`)
    assert(!/r:id=/.test(tag), `${sheet} link carries no r:id`)
  }
  const qsRelsPath = qsPath.replace(/worksheets\/(sheet\d+)\.xml$/, 'worksheets/_rels/$1.xml.rels')
  const qsRels = (await zip.file(qsRelsPath)?.async('string')) ?? ''
  assert(
    !/relationships\/hyperlink/.test(qsRels),
    'no leftover external hyperlink relationships on QUICK START'
  )

  // ── OVERVIEW's chart + drawing survived the reshuffle ──────────────────────
  assert(!!zip.file('xl/charts/chart1.xml'), 'chart1.xml still injected')
  const ovXml = await zip.file(pathFor('OVERVIEW')).async('string')
  const ovRelsPath = pathFor('OVERVIEW').replace(
    /worksheets\/(sheet\d+)\.xml$/,
    'worksheets/_rels/$1.xml.rels'
  )
  const ovRels = (await zip.file(ovRelsPath)?.async('string')) ?? ''
  for (const rid of new Set([...ovXml.matchAll(/r:id="(rId\d+)"/g)].map((m) => m[1]))) {
    assert(ovRels.includes(`Id="${rid}"`), `OVERVIEW r:id ${rid} still resolves`)
  }
}

// ── Without recommendations ───────────────────────────────────────────────────
{
  const blob = await generateWorkbook(baseState)
  const buffer = Buffer.from(await blob.arrayBuffer())
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer)
  const ws = wb.getWorksheet('QUICK START')

  const text = []
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      const v = cell.value
      text.push(typeof v === 'object' && v ? String(v.text ?? '') : String(v ?? ''))
    })
  })
  const all = text.join('\n')

  assert(!/Recommendations/.test(all), 'no recommendation step when there are no guides')
  const badges = text.filter((t) => /^[1-9]$/.test(t)).map(Number)
  assert(
    JSON.stringify(badges) === JSON.stringify([1, 2, 3, 4]),
    `steps renumber to 1..4 with the rec step gone (got ${badges.join(',')})`
  )
  assert(
    /waiting to be filled in/.test(all),
    'itinerary step uses the blank-sheet wording, not "sample plan"'
  )
}

console.log(process.exitCode ? '\nFAILED' : '\nAll QUICK START checks passed')
console.log('Wrote /tmp/quick-start-test.xlsx for manual inspection in Excel/Google Sheets.')
