// Verifies injectChart correctly merges a chart into a worksheet that already has an
// ExcelJS-native image (the OVERVIEW cover photo, anchored via <xdr:pic>) into the SAME
// drawing part, and that the no-image path still works as a regression check.
// Run: node scripts/verify-image-chart-injection.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { writeFileSync, rmSync } from 'node:fs'

// Transpile the TS injection module to a temp file inside the project so node can
// resolve its bare `jszip` import against node_modules.
const tmpModPath = new URL('./.chartInjection.tmp2.mjs', import.meta.url)
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

// Minimal (structurally valid, content-empty) JPEG: SOI + EOI markers. ExcelJS stores
// image buffers as opaque media parts — it never decodes them — so this is sufficient
// to exercise the <xdr:pic> drawing path without a real photo.
const FAKE_JPEG = Buffer.from([0xff, 0xd8, 0xff, 0xd9])

// Mirrors the OVERVIEW budget table + chart helper columns (M/N/O/P) from overview.ts.
function buildBudgetTable(ws) {
  const cats = ['Transportation', 'Accommodation', 'Food', 'Activities', 'Shopping', 'Miscellaneous']
  cats.forEach((c, i) => {
    const row = 18 + i
    ws.getCell(`F${row}`).value = c
    ws.getCell(`G${row}`).value = (i + 1) * 200
    ws.getCell(`H${row}`).value = i === 0 ? 250 : 0
    ws.getCell(`M${row}`).value = { formula: `IF(H${row}>0,H${row},G${row})` }
    ws.getCell(`N${row}`).value = { formula: `MIN(H${row},G${row})` }
    ws.getCell(`O${row}`).value = { formula: `MAX(G${row}-H${row},0)` }
    ws.getCell(`P${row}`).value = { formula: `MAX(H${row}-G${row},0)` }
  })
  ws.getColumn('M').hidden = true
  ws.getColumn('N').hidden = true
  ws.getColumn('O').hidden = true
  ws.getColumn('P').hidden = true
}

const BASE_OPTS = {
  sheetName: 'OVERVIEW',
  categoryRange: "'OVERVIEW'!$F$18:$F$23",
  anchor: { fromCol: 5, fromRow: 26, toCol: 11, toRow: 44 },
  colors: CHART_COLORS,
  title: 'Test Budget Breakdown',
}

function injectOptsFor(kind) {
  if (kind === 'bar') {
    return {
      ...BASE_OPTS,
      kind,
      valueRange: "'OVERVIEW'!$N$18:$N$23",
      remainderRange: "'OVERVIEW'!$O$18:$O$23",
      overRange: "'OVERVIEW'!$P$18:$P$23",
    }
  }
  return { ...BASE_OPTS, kind, valueRange: "'OVERVIEW'!$M$18:$M$23" }
}

// ── Case A: cover photo (F3:K14) + injected chart, merged into one drawing part ─────
console.log('── Case A: image + chart ──')
{
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('OVERVIEW')
  for (let r = 3; r <= 14; r++) ws.getRow(r).height = 20
  buildBudgetTable(ws)
  wb.addWorksheet('OTHER EXPENSES')

  const imageId = wb.addImage({ buffer: FAKE_JPEG, extension: 'jpeg' })
  ws.addImage(imageId, 'F3:K14')

  const baseBuffer = await wb.xlsx.writeBuffer()

  for (const kind of ['bar', 'pie', 'doughnut']) {
    console.log(`\n  ${kind}:`)
    const injected = await injectChart(baseBuffer, injectOptsFor(kind))
    writeFileSync(`/tmp/image-chart-${kind}-test.xlsx`, Buffer.from(injected))

    const zip = await JSZip.loadAsync(injected)

    assert(zip.file('xl/drawings/drawing1.xml'), '  drawing1.xml exists (merged, not a new drawing2.xml)')
    assert(!zip.file('xl/drawings/drawing2.xml'), '  no drawing2.xml was created')

    const drawingXml = await zip.file('xl/drawings/drawing1.xml').async('string')
    // ExcelJS writes the image anchor as <xdr:twoCellAnchor editAs="oneCell">, so match
    // with optional attributes (not just the bare "<xdr:twoCellAnchor>").
    const anchorCount = (drawingXml.match(/<xdr:twoCellAnchor[ >]/g) || []).length
    assert(anchorCount === 2, `  drawing1.xml has 2 <xdr:twoCellAnchor> elements (got ${anchorCount})`)
    assert(/<xdr:pic[ >]/.test(drawingXml), '  drawing1.xml retains the <xdr:pic> (cover photo)')
    assert(drawingXml.includes('<xdr:graphicFrame'), '  drawing1.xml contains the chart <xdr:graphicFrame>')

    const drawingRels = await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string')
    assert(drawingRels.includes('Id="rId1"'), '  drawing1.xml.rels retains rId1 (image)')
    assert(
      drawingRels.includes('Id="rId2"') && drawingRels.includes('../charts/chart1.xml'),
      '  drawing1.xml.rels has new rId2 → chart1.xml'
    )

    // Worksheet has exactly one <drawing> element.
    const wbXml = await zip.file('xl/workbook.xml').async('string')
    const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
    const rid = /<sheet[^>]*name="OVERVIEW"[^>]*r:id="(rId\d+)"/.exec(wbXml)[1]
    const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(wbRels)[1]
    const sheetPath = 'xl/' + target.replace(/^\//, '').replace(/^xl\//, '')
    const sheetXml = await zip.file(sheetPath).async('string')
    const drawingMatches = sheetXml.match(/<drawing r:id="rId\d+"\/>/g) || []
    assert(drawingMatches.length === 1, `  worksheet has exactly one <drawing> element (got ${drawingMatches.length})`)

    // Content types — exactly one Override each for drawing1.xml and chart1.xml.
    const ct = await zip.file('[Content_Types].xml').async('string')
    const drawingOverrides = (ct.match(/PartName="\/xl\/drawings\/drawing1\.xml"/g) || []).length
    const chartOverrides = (ct.match(/PartName="\/xl\/charts\/chart1\.xml"/g) || []).length
    assert(drawingOverrides === 1, `  [Content_Types].xml has exactly one drawing1.xml Override (got ${drawingOverrides})`)
    assert(chartOverrides === 1, `  [Content_Types].xml has exactly one chart1.xml Override (got ${chartOverrides})`)

    // Round-trip through ExcelJS.
    try {
      const wb2 = new ExcelJS.Workbook()
      await wb2.xlsx.load(injected)
      const ws2 = wb2.getWorksheet('OVERVIEW')
      assert(ws2.getCell('F18').value === 'Transportation', '  ExcelJS re-opens injected file intact')
      assert(ws2.getImages().length === 1, `  ws2.getImages() returns 1 image (got ${ws2.getImages().length})`)
    } catch (e) {
      assert(false, '  ExcelJS re-open threw: ' + e.message)
    }
  }
}

// ── Case B: chart-only regression (no cover photo) ─────────────────────────────────
console.log('\n── Case B: chart only (no image) ──')
{
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('OVERVIEW')
  buildBudgetTable(ws)
  wb.addWorksheet('OTHER EXPENSES')

  const baseBuffer = await wb.xlsx.writeBuffer()
  const injected = await injectChart(baseBuffer, injectOptsFor('pie'))
  writeFileSync('/tmp/image-chart-none-test.xlsx', Buffer.from(injected))

  const zip = await JSZip.loadAsync(injected)
  const drawingXml = await zip.file('xl/drawings/drawing1.xml').async('string')
  const anchorCount = (drawingXml.match(/<xdr:twoCellAnchor[ >]/g) || []).length
  assert(anchorCount === 1, `drawing1.xml has 1 <xdr:twoCellAnchor> (got ${anchorCount})`)

  const wbXml = await zip.file('xl/workbook.xml').async('string')
  const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
  const rid = /<sheet[^>]*name="OVERVIEW"[^>]*r:id="(rId\d+)"/.exec(wbXml)[1]
  const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(wbRels)[1]
  const sheetPath = 'xl/' + target.replace(/^\//, '').replace(/^xl\//, '')
  const sheetXml = await zip.file(sheetPath).async('string')
  const drawingMatches = sheetXml.match(/<drawing r:id="rId\d+"\/>/g) || []
  assert(drawingMatches.length === 1, `worksheet has exactly one <drawing> element (got ${drawingMatches.length})`)

  const ct = await zip.file('[Content_Types].xml').async('string')
  assert(ct.includes('/xl/drawings/drawing1.xml'), 'content-types declares drawing1.xml')
  assert(ct.includes('/xl/charts/chart1.xml'), 'content-types declares chart1.xml')

  try {
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(injected)
    const ws2 = wb2.getWorksheet('OVERVIEW')
    assert(ws2.getCell('F18').value === 'Transportation', 'ExcelJS re-opens injected file intact')
    assert(ws2.getImages().length === 0, 'ws2.getImages() returns 0 images')
  } catch (e) {
    assert(false, 'ExcelJS re-open threw: ' + e.message)
  }
}

console.log('\nDone. Wrote /tmp/image-chart-{bar,pie,doughnut,none}-test.xlsx for manual inspection.')
