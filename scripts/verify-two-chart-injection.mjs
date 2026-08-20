// Verifies injectChart correctly appends a SECOND chart (the readiness doughnut) into a
// drawing part that already holds a FIRST chart (the budget bar), both merged into the
// same <xdr:wsDr>. This is the code path the OVERVIEW sheet now exercises: two sequential
// injectChart calls on one buffer. Regression-tests the cNvPr frame-id parameterization
// in chartInjection.ts — before that fix, both charts wrote <xdr:cNvPr id="2">, which
// Excel treats as corrupt content and offers to "repair" on open.
// Run: node scripts/verify-two-chart-injection.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { writeFileSync, rmSync } from 'node:fs'

const tmpModPath = new URL('./.chartInjection.tmp3.mjs', import.meta.url)
await build({
  entryPoints: ['src/lib/excel/chartInjection.ts'],
  bundle: true,
  format: 'esm',
  outfile: new URL(tmpModPath).pathname,
  platform: 'node',
  external: ['jszip'],
})
const mod = await import(tmpModPath.href)
const { injectChart, CHART_COLORS } = mod
rmSync(new URL(tmpModPath).pathname)

function assert(cond, msg) {
  if (!cond) {
    console.error('❌ FAIL:', msg)
    process.exitCode = 1
  } else {
    console.log('✓', msg)
  }
}

// Mirrors the OVERVIEW budget table + chart helper columns (M/N/O/P) and the
// trip-readiness helper cells (S1:S3/T1:T2) from overview.ts.
function buildOverviewLikeSheet(wb) {
  const ws = wb.addWorksheet('OVERVIEW')
  const cats = ['Transportation', 'Accommodation', 'Food', 'Activities', 'Shopping', 'Miscellaneous']
  cats.forEach((c, i) => {
    const row = 18 + i
    ws.getCell(`F${row}`).value = c
    ws.getCell(`G${row}`).value = (i + 1) * 200
    ws.getCell(`H${row}`).value = i === 0 ? 250 : 0
    ws.getCell(`N${row}`).value = { formula: `MIN(H${row},G${row})` }
    ws.getCell(`O${row}`).value = { formula: `MAX(G${row}-H${row},0)` }
    ws.getCell(`P${row}`).value = { formula: `MAX(H${row}-G${row},0)` }
  })
  ws.getColumn('N').hidden = true
  ws.getColumn('O').hidden = true
  ws.getColumn('P').hidden = true

  ws.getCell('T1').value = 'Ready'
  ws.getCell('T2').value = 'To do'
  ws.getCell('S1').value = { formula: 'IFERROR(COUNTIF(\'PACKING LIST\'!$C$5:$C$500,"✓"),0)', result: 0 }
  ws.getCell('S3').value = { formula: 'IFERROR(COUNTA(\'PACKING LIST\'!$B$5:$B$500),0)', result: 0 }
  ws.getCell('S2').value = { formula: 'MAX(S3-S1,0)', result: 1 }
  ws.getColumn('S').hidden = true
  ws.getColumn('T').hidden = true

  wb.addWorksheet('PACKING LIST')
  return ws
}

const BAR_OPTS = {
  kind: 'bar',
  sheetName: 'OVERVIEW',
  categoryRange: "'OVERVIEW'!$F$18:$F$23",
  valueRange: "'OVERVIEW'!$N$18:$N$23",
  remainderRange: "'OVERVIEW'!$O$18:$O$23",
  overRange: "'OVERVIEW'!$P$18:$P$23",
  anchor: { fromCol: 5, fromRow: 26, toCol: 11, toRow: 44 },
  colors: CHART_COLORS,
  title: 'Test Budget Breakdown',
}

// Mirrors the production OVERVIEW readiness-ring call (index.ts): no title/legend/
// per-slice labels, so the doughnut's hole stays centered for the in-cell % overlay.
const DOUGHNUT_OPTS = {
  kind: 'doughnut',
  sheetName: 'OVERVIEW',
  categoryRange: "'OVERVIEW'!$T$1:$T$2",
  categoryLabels: ['Ready', 'To do'],
  valueRange: "'OVERVIEW'!$S$1:$S$2",
  values: [0, 1],
  anchor: { fromCol: 0, fromRow: 23, toCol: 4, toRow: 39 },
  colors: [CHART_COLORS[0], CHART_COLORS[1]],
  legend: false,
  dataLabels: false,
}

async function resolveOverviewSheetXml(zip) {
  const wbXml = await zip.file('xl/workbook.xml').async('string')
  const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
  const rid = /<sheet[^>]*name="OVERVIEW"[^>]*r:id="(rId\d+)"/.exec(wbXml)[1]
  const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(wbRels)[1]
  const sheetPath = 'xl/' + target.replace(/^\//, '').replace(/^xl\//, '')
  return { sheetPath, sheetXml: await zip.file(sheetPath).async('string') }
}

// ── Case A: two sequential injectChart calls (bar, then doughnut) on one buffer ─────
console.log('── Case A: bar chart + doughnut chart, merged into one drawing part ──')
{
  const wb = new ExcelJS.Workbook()
  buildOverviewLikeSheet(wb)
  const baseBuffer = await wb.xlsx.writeBuffer()

  let buffer = await injectChart(baseBuffer, BAR_OPTS)
  buffer = await injectChart(buffer, DOUGHNUT_OPTS)
  writeFileSync('/tmp/two-chart-test.xlsx', Buffer.from(buffer))

  const zip = await JSZip.loadAsync(buffer)

  assert(zip.file('xl/charts/chart1.xml'), 'chart1.xml exists (bar)')
  assert(zip.file('xl/charts/chart2.xml'), 'chart2.xml exists (doughnut)')
  assert(zip.file('xl/drawings/drawing1.xml'), 'drawing1.xml exists')
  assert(!zip.file('xl/drawings/drawing2.xml'), 'no drawing2.xml was created — both charts share one drawing part')

  const drawingXml = await zip.file('xl/drawings/drawing1.xml').async('string')
  const anchorCount = (drawingXml.match(/<xdr:twoCellAnchor[ >]/g) || []).length
  assert(anchorCount === 2, `drawing1.xml has 2 <xdr:twoCellAnchor> elements (got ${anchorCount})`)

  // Regression test for the id-collision fix: each chart's cNvPr id must be distinct.
  const frameIds = [...drawingXml.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map(m => m[1])
  assert(frameIds.length === 2, `found 2 <xdr:cNvPr> elements (got ${frameIds.length})`)
  assert(new Set(frameIds).size === 2, `cNvPr ids are distinct, not colliding (got [${frameIds.join(', ')}])`)
  assert(
    frameIds[0] === '2' && frameIds[1] === '3',
    `cNvPr ids are 2 (first chart) then 3 (second chart) (got [${frameIds.join(', ')}])`
  )

  const drawingRels = await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string')
  assert(
    drawingRels.includes('Id="rId1"') && drawingRels.includes('../charts/chart1.xml'),
    'drawing1.xml.rels has rId1 → chart1.xml'
  )
  assert(
    drawingRels.includes('Id="rId2"') && drawingRels.includes('../charts/chart2.xml'),
    'drawing1.xml.rels has rId2 → chart2.xml'
  )

  const { sheetXml } = await resolveOverviewSheetXml(zip)
  const drawingMatches = sheetXml.match(/<drawing r:id="rId\d+"\/>/g) || []
  assert(drawingMatches.length === 1, `worksheet has exactly one <drawing> element (got ${drawingMatches.length})`)

  const ct = await zip.file('[Content_Types].xml').async('string')
  const drawingOverrides = (ct.match(/PartName="\/xl\/drawings\/drawing1\.xml"/g) || []).length
  const chart1Overrides = (ct.match(/PartName="\/xl\/charts\/chart1\.xml"/g) || []).length
  const chart2Overrides = (ct.match(/PartName="\/xl\/charts\/chart2\.xml"/g) || []).length
  assert(drawingOverrides === 1, `[Content_Types].xml has exactly one drawing1.xml Override (got ${drawingOverrides})`)
  assert(chart1Overrides === 1, `[Content_Types].xml has exactly one chart1.xml Override (got ${chart1Overrides})`)
  assert(chart2Overrides === 1, `[Content_Types].xml has exactly one chart2.xml Override (got ${chart2Overrides})`)

  // Readiness-ring specific: no legend/per-slice labels, transparent chart space (so an
  // in-cell % label floated behind the doughnut's hole isn't hidden by a white box).
  const chart2Xml = await zip.file('xl/charts/chart2.xml').async('string')
  assert(!chart2Xml.includes('<c:legend>'), 'chart2.xml (doughnut) has no <c:legend>')
  assert(chart2Xml.includes('<c:showCatName val="0"/>'), 'chart2.xml (doughnut) data labels are off (showCatName=0)')
  assert(chart2Xml.includes('<c:roundedCorners val="0"/>'), 'chart2.xml (doughnut) chart space has roundedCorners=0')
  assert(
    /<\/c:chart><c:spPr><a:noFill\/>/.test(chart2Xml),
    'chart2.xml (doughnut) chart space has a transparent (noFill) background'
  )

  try {
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(buffer)
    const ws2 = wb2.getWorksheet('OVERVIEW')
    assert(ws2.getCell('F18').value === 'Transportation', 'ExcelJS re-opens the twice-injected file intact')
  } catch (e) {
    assert(false, 'ExcelJS re-open threw: ' + e.message)
  }
}

console.log('\nDone. Wrote /tmp/two-chart-test.xlsx for manual inspection.')
