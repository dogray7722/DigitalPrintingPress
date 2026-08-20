// Verifies the recommendation jump-link pair on ACCOMMODATION / DINING / EXCURSIONS:
//   - a "jump to recommendations" HYPERLINK on the row 4 spacer, targeting the guide's
//     section band, with a cached result (calcId is pinned + fullCalcOnLoad is off, so an
//     uncached formula renders blank in Excel)
//   - a "back to top" HYPERLINK on the right end of that section band, targeting A1
//   - and, critically, NEITHER when there are no recommendations — a link into a blank
//     row is worse than no link at all.
// Run: node scripts/verify-rec-jump-links.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { rmSync } from 'node:fs'

const tmpModPath = new URL('./.recjump.tmp.mjs', import.meta.url)
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
  useRecommendations: true,
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

// One region carrying at least one place of each kind, so all three inserts render.
const recommendations = {
  regions: [
    {
      region: 'Alfama',
      description: 'The oldest district, all tiled alleys and fado houses.',
      hotels: [{ name: 'Memmo Alfama', price: '$$$', rating: '4.6', description: 'Rooftop pool over the rooftops.' }],
      restaurants: [{ name: 'Ramiro', cuisine: 'Seafood', price: '$$', description: 'Garlic prawns and beer.' }],
      excursions: [{ name: 'Tram 28 ride', duration: '1h', price: '$', description: 'The classic hillside loop.' }],
    },
  ],
  itinerary: [],
  events: [],
}

const LINK_TEXT = '★   Recommendations   →'

// WCAG contrast, mirroring styleFactory.ts — the chip's accent color is chosen by
// measurement, so the test measures the result rather than trusting a hardcoded hex.
const lin = (h) => {
  const c = parseInt(h, 16) / 255
  return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4)
}
const lum = (argb) =>
  0.2126 * lin(argb.slice(2, 4)) + 0.7152 * lin(argb.slice(4, 6)) + 0.0722 * lin(argb.slice(6, 8))
const contrast = (a, b) => {
  const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x)
  return (hi + 0.05) / (lo + 0.05)
}

// [tab name, guide section-band row, button anchor cell]
// Rows follow each builder's `totRow + 2`; the button is merged over the last two cells
// of that sheet's row-2 band, anchored on the first of them.
const SHEETS = [
  ['ACCOMMODATION', 28, 'F2'],
  ['DINING', 58, 'G2'],
  ['EXCURSIONS', 38, 'I2'],
]

// NOTE: we deliberately assert only the TEXT at the ExcelJS level, never the link target.
// ExcelJS's reader binds a hyperlink to its cell only when the element carries an r:id
// (worksheet-xform.js: `if (hyperlink.rId)`), so a native internal link — Excel's own
// shape, r:id-less by definition — is parsed but never surfaced on the cell. That's a
// reader limitation which applies to genuine Excel-authored files too. The targets are
// verified against the raw XML further down, which is the real contract anyway.
// What this DOES prove: the label survives as the cell's own value, so a consumer that
// ignores hyperlinks entirely still shows the words rather than a blank cell.
function textOf(cell) {
  const v = cell.value
  return typeof v === 'object' && v ? v.text : v
}

async function generate(...args) {
  const blob = await generateWorkbook(...args)
  return Buffer.from(await blob.arrayBuffer())
}

async function loadWorkbook(buffer) {
  const wb = new ExcelJS.Workbook()
  // Read back through ExcelJS so we exercise the real xlsx parse, not a string grep.
  await wb.xlsx.load(buffer)
  return wb
}

// ── With recommendations: the link is present, styled and correctly targeted ──
const buffer = await generate(state, recommendations)
const wb = await loadWorkbook(buffer)

for (const [name, recRow, linkRef] of SHEETS) {
  const ws = wb.getWorksheet(name)
  assert(!!ws, `${name} sheet exists`)
  if (!ws) continue

  const link = ws.getCell(linkRef)
  assert(textOf(link) === LINK_TEXT, `${name}!${linkRef} label is "${LINK_TEXT}"`)

  // Button chip: it must CONTRAST with the band rather than match it. Compared against
  // A2 (the description cell) rather than hardcoded hex, so this holds for every theme.
  const band = ws.getCell('A2')
  assert(
    !!link.fill?.fgColor?.argb && link.fill.fgColor.argb !== band.fill?.fgColor?.argb,
    `${name}!${linkRef} fill contrasts with the band`
  )
  assert(
    !!link.font?.color?.argb && link.font.color.argb !== band.font?.color?.argb,
    `${name}!${linkRef} text color contrasts with the band`
  )
  assert(link.font?.bold === true, `${name}!${linkRef} is bold`)
  assert(!link.font?.underline, `${name}!${linkRef} is not underlined`)
  assert(link.alignment?.horizontal === 'center', `${name}!${linkRef} is centered`)
  for (const side of ['top', 'bottom', 'left', 'right']) {
    assert(link.border?.[side]?.style === 'thin', `${name}!${linkRef} has a ${side} frame edge`)
    assert(
      link.border?.[side]?.color?.argb === link.font?.color?.argb,
      `${name}!${linkRef} ${side} edge matches the accent`
    )
  }

  // The guide's section band is one full-width merge again (no back-to-top link).
  const title = ws.getCell(`A${recRow}`)
  assert(
    typeof title.value === 'string' && title.value.includes('RECOMMENDED'),
    `${name}!A${recRow} holds the guide title`
  )

  // Row 4 was the old link home; it must be clean again.
  assert(ws.getCell('A4').value == null, `${name}!A4 is empty (link moved to row 2)`)
}

// No back-to-top links survive anywhere in the workbook.
const strays = []
wb.eachSheet((ws) => {
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      const t = typeof cell.value === 'object' && cell.value ? cell.value.text : cell.value
      if (typeof t === 'string' && t.includes('Back to top')) strays.push(`${ws.name}!${cell.address}`)
    })
  })
})
assert(strays.length === 0, `no "Back to top" text anywhere${strays.length ? ` (found: ${strays})` : ''}`)

// ── The chip's accent must stay legible on EVERY theme ───────────────────────
// The accent is picked by measuring contrast against the chip fill, so a new theme with
// a low-contrast secondary silently degrades to unreadable text without this check.
// inkwell is the shipping theme and is asserted to get the intended gold specifically.
const THEME_SECONDARY = {
  sakura: 'FFB8A9D4',
  ocean: 'FF5E9EC8',
  forest: 'FF5E9E78',
  desert: 'FFD4A050',
  inkwell: 'FFC9A96E',
  parchment: 'FFC9A96E',
}

for (const theme of Object.keys(THEME_SECONDARY)) {
  const themed = await loadWorkbook(await generate({ ...state, theme }, recommendations))
  const chip = themed.getWorksheet('ACCOMMODATION').getCell('F2')
  const fg = chip.font?.color?.argb
  const bg = chip.fill?.fgColor?.argb
  const ratio = contrast(fg, bg)
  assert(
    ratio >= 3,
    `${theme}: chip text ${fg} on ${bg} is legible (${ratio.toFixed(2)}:1)`
  )
  if (theme === 'inkwell') {
    assert(
      fg === THEME_SECONDARY.inkwell,
      `inkwell: chip uses the spec's gold accent (${fg}, ${ratio.toFixed(2)}:1)`
    )
  }
}

// ── The emitted XML must match Excel's native internal-hyperlink shape ────────
// This is the part that makes the links work in Google Sheets. ExcelJS on its own emits
//   <hyperlink ref="G2" r:id="rId1" location="'ACCOMMODATION'!A28"/>
// plus a TargetMode="External" relationship. Excel emits (and nativeHyperlinks.ts
// rewrites to)
//   <hyperlink ref="G2" location="'ACCOMMODATION'!A28" display="⭐  Recommendations"/>
// Without `display`, Google Sheets shows the link text as a literal "#gid=xxxx".
// (Confirmed against a real Drive import — see CLAUDE.md.)
const zip = await JSZip.loadAsync(buffer)
const sheetPaths = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))

const hyperlinks = []
for (const path of sheetPaths) {
  const xml = await zip.file(path).async('string')
  for (const m of xml.matchAll(/<hyperlink\b[^>]*\/>/g)) hyperlinks.push({ path, tag: m[0] })
}

// Count only the tracker links by their target — OVERVIEW's quick-nav strip also emits
// native hyperlinks now, and its count varies with which sheets are toggled on.
const guideTargets = SHEETS.map(([name, recRow]) => `'${name}'!A${recRow}`)
const trackerLinks = hyperlinks.filter((h) =>
  guideTargets.some((t) => h.tag.includes(`location="${t.replace(/'/g, '&apos;')}"`))
)
assert(
  trackerLinks.length === 3,
  `3 tracker guide links emitted (one per sheet), got ${trackerLinks.length}`
)

for (const { tag } of hyperlinks) {
  assert(!/r:id=/.test(tag), `hyperlink carries no r:id — ${tag}`)
  assert(/\blocation="/.test(tag), `hyperlink carries a location — ${tag}`)
  assert(/\bdisplay="/.test(tag), `hyperlink carries a display attribute — ${tag}`)
}

// Every expected target/label pair is present, with the label in `display`.
// ExcelJS XML-escapes the quotes around the sheet name.
const q = (s) => s.replace(/'/g, '&apos;')
for (const [name, recRow, linkRef] of SHEETS) {
  const tag = hyperlinks.find((h) => h.tag.includes(`location="${q(`'${name}'!A${recRow}`)}"`))
  assert(!!tag, `native hyperlink targets '${name}'!A${recRow}`)
  assert(!!tag && tag.tag.includes(`ref="${linkRef}"`), `${name} link sits at ${linkRef}`)
  assert(
    !!tag && tag.tag.includes(`display="${LINK_TEXT}"`),
    `${name} hyperlink display is "${LINK_TEXT}"`
  )
}

// The bogus TargetMode="External" relationships ExcelJS paired with them must be gone —
// otherwise Excel tries to resolve "'ACCOMMODATION'!A28" as an external file.
for (const path of sheetPaths) {
  const relsPath = path.replace(/worksheets\/(sheet\d+)\.xml$/, 'worksheets/_rels/$1.xml.rels')
  const rels = zip.file(relsPath)
  if (!rels) continue
  const xml = await rels.async('string')
  assert(
    !/relationships\/hyperlink/.test(xml),
    `${relsPath} has no leftover hyperlink relationship`
  )
}

// Deleting relationships risks leaving a dangling r:id behind — the classic cause of
// Excel's "we found a problem, do you want us to repair" prompt. Every r:id still
// referenced by a sheet must resolve in that sheet's rels. (OVERVIEW's <drawing r:id>
// is the one that would break if this pass ever renumbered instead of just deleting.)
for (const path of sheetPaths) {
  const xml = await zip.file(path).async('string')
  const referenced = [...xml.matchAll(/r:id="(rId\d+)"/g)].map((m) => m[1])
  if (!referenced.length) continue
  const relsPath = path.replace(/worksheets\/(sheet\d+)\.xml$/, 'worksheets/_rels/$1.xml.rels')
  const rels = (await zip.file(relsPath)?.async('string')) ?? ''
  for (const rid of new Set(referenced)) {
    assert(rels.includes(`Id="${rid}"`), `${path} r:id ${rid} resolves in ${relsPath}`)
  }
}

// ── Without recommendations: no guide, so no links anywhere ───────────────────
const bare = await loadWorkbook(await generate({ ...state, useRecommendations: false }))

for (const [name, recRow, linkRef] of SHEETS) {
  const ws = bare.getWorksheet(name)
  const slot = ws.getCell(linkRef)
  assert(slot.value == null, `${name}!${linkRef} has no link when there is no guide`)
  assert(ws.getCell(`A${recRow}`).value == null, `${name}!A${recRow} is empty when there is no guide`)

  // ...but the reserved slot still carries the band fill, so row 2 reads as one
  // continuous strip rather than a coloured bar with a white notch bitten out of it.
  assert(
    slot.fill?.fgColor?.argb === ws.getCell('A2').fill?.fgColor?.argb,
    `${name}!${linkRef} still matches row 2's fill when there is no guide`
  )
}

console.log(process.exitCode ? '\nFAILED' : '\nAll recommendation jump-link checks passed')
