// Verifies injectChart produces a structurally valid, re-openable .xlsx for every
// chart kind (bar / pie / doughnut).
// Run: node scripts/verify-chart-injection.mjs
import { build } from 'esbuild'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { writeFileSync, rmSync } from 'node:fs'

// Transpile the TS injection module to a temp file inside the project so node can
// resolve its bare `jszip` import against node_modules.
const tmpModPath = new URL('./.chartInjection.tmp.mjs', import.meta.url)
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

// Build a representative workbook resembling the real OVERVIEW layout.
const wb = new ExcelJS.Workbook()
const ws = wb.addWorksheet('OVERVIEW')
const cats = ['Transportation', 'Accommodation', 'Food', 'Activities', 'Shopping', 'Miscellaneous']
cats.forEach((c, i) => {
  const row = 18 + i
  ws.getCell(`F${row}`).value = c
  ws.getCell(`G${row}`).value = (i + 1) * 200  // budget estimate
  ws.getCell(`H${row}`).value = i === 0 ? 250 : 0  // first category over budget; rest no actuals
  // Col M: pie/donut helper — actual when > 0, else estimate (mirrors overview.ts).
  ws.getCell(`M${row}`).value = { formula: `IF(H${row}>0,H${row},G${row})` }
  // Col N: bar chart helper — MIN(actual, budget), colored "spent" segment (mirrors overview.ts).
  ws.getCell(`N${row}`).value = { formula: `MIN(H${row},G${row})` }
  // Col O: bar chart helper — MAX(budget - actual, 0), light "remaining" segment.
  ws.getCell(`O${row}`).value = { formula: `MAX(G${row}-H${row},0)` }
})
ws.getColumn('M').hidden = true
ws.getColumn('N').hidden = true
ws.getColumn('O').hidden = true
// A second sheet to ensure sheet-path resolution picks the right one.
wb.addWorksheet('OTHER EXPENSES')

const baseBuffer = await wb.xlsx.writeBuffer()

// Per-kind config: expected XML wrapper + extra assertions + injectChart options override.
const KINDS = [
  {
    kind: 'bar',
    element: '<c:barChart>',
    extra: [
      '<c:barDir val="bar"/>',
      '<c:grouping val="stacked"/>',     // stacked, not clustered (Google Sheets ignores overlap)
      '<c:overlap val="100"/>',
      '<a:noFill/>',                    // no borders on any segment (equal bar height)
      '<a:pattFill prst="ltDnDiag">',  // hatched "remaining budget" track
      '<a:srgbClr val="E53935"/>',      // solid red "over budget" overage segment
      '<c:catAx>', '<c:valAx>',
      '<c:orientation val="maxMin"/>',  // first category at top, matching the table order
      '<c:crosses val="max"/>',         // value scale stays along the bottom after flip
    ],
    opts: {
      valueRange: "'OVERVIEW'!$N$18:$N$23",
      remainderRange: "'OVERVIEW'!$O$18:$O$23",
      overRange: "'OVERVIEW'!$P$18:$P$23",
    },
    // bar has 3 series; dPts only on Series 0 (spent, per-category theme colors)
    expectedDPts: 6,
    valueRangeCheck: "'OVERVIEW'!$N$18:$N$23",
  },
  {
    kind: 'pie',
    element: '<c:pieChart>',
    extra: [],
    opts: { valueRange: "'OVERVIEW'!$M$18:$M$23" },
    expectedDPts: 6,
    valueRangeCheck: "'OVERVIEW'!$M$18:$M$23",
  },
  {
    kind: 'doughnut',
    element: '<c:doughnutChart>',
    extra: ['<c:holeSize val="50"/>'],
    opts: { valueRange: "'OVERVIEW'!$M$18:$M$23" },
    expectedDPts: 6,
    valueRangeCheck: "'OVERVIEW'!$M$18:$M$23",
  },
]

for (const { kind, element, extra, opts: kindOpts, expectedDPts, valueRangeCheck } of KINDS) {
  console.log(`\n── ${kind} ──`)
  const injected = await injectChart(baseBuffer, {
    kind,
    sheetName: 'OVERVIEW',
    categoryRange: "'OVERVIEW'!$F$18:$F$23",
    anchor: { fromCol: 5, fromRow: 26, toCol: 11, toRow: 44 },
    colors: CHART_COLORS,
    title: 'Test Budget Breakdown',
    ...kindOpts,
  })

  const outPath = `/tmp/chart-${kind}-test.xlsx`
  writeFileSync(outPath, Buffer.from(injected))

  // 1. Inspect the zip parts.
  const zip = await JSZip.loadAsync(injected)
  assert(zip.file('xl/charts/chart1.xml'), 'chart1.xml exists')
  assert(zip.file('xl/drawings/drawing1.xml'), 'drawing1.xml exists')
  assert(zip.file('xl/drawings/_rels/drawing1.xml.rels'), 'drawing1.xml.rels exists')

  const ct = await zip.file('[Content_Types].xml').async('string')
  assert(ct.includes('/xl/charts/chart1.xml'), 'content-types declares chart')
  assert(ct.includes('/xl/drawings/drawing1.xml'), 'content-types declares drawing')

  const chart = await zip.file('xl/charts/chart1.xml').async('string')
  assert(chart.includes(element), `chart is a ${kind}`)
  for (const e of extra) assert(chart.includes(e), `chart contains ${e}`)
  assert(chart.includes(valueRangeCheck), `chart binds value range (${valueRangeCheck})`)
  assert(chart.includes("'OVERVIEW'!$F$18:$F$23"), 'chart binds category range')
  assert(chart.includes('<c:plotVisOnly val="0"/>'), 'plotVisOnly=0 so hidden helper cols still plot')
  assert((chart.match(/<c:dPt>/g) || []).length === expectedDPts, `chart has ${expectedDPts} colored data points`)
  assert(chart.includes('<a:srgbClr val="4ECDC4"'), 'ARGB alpha stripped to RGB')

  // Find the OVERVIEW sheet part and confirm the <drawing> element + rel wiring.
  const wbXml = await zip.file('xl/workbook.xml').async('string')
  const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
  const rid = /<sheet[^>]*name="OVERVIEW"[^>]*r:id="(rId\d+)"/.exec(wbXml)[1]
  const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(wbRels)[1]
  const sheetPath = 'xl/' + target.replace(/^\//, '').replace(/^xl\//, '')
  const sheetXml = await zip.file(sheetPath).async('string')
  const drawMatch = /<drawing r:id="(rId\d+)"\/>/.exec(sheetXml)
  assert(drawMatch, '<drawing> element present in OVERVIEW sheet')
  assert(
    sheetXml.indexOf('<drawing') > sheetXml.indexOf('</sheetData>'),
    '<drawing> comes after </sheetData> (valid child order)'
  )
  const sheetRelsPath = sheetPath.replace(/worksheets\//, 'worksheets/_rels/') + '.rels'
  const sheetRels = await zip.file(sheetRelsPath).async('string')
  assert(
    sheetRels.includes(`Id="${drawMatch[1]}"`) && sheetRels.includes('drawing1.xml'),
    'sheet rels maps the drawing rId to drawing1.xml'
  )

  // 2. The real test: ExcelJS re-opens it without throwing (proves it isn't corrupt).
  try {
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(injected)
    const ws2 = wb2.getWorksheet('OVERVIEW')
    assert(ws2 && ws2.getCell('F18').value === 'Transportation', 'ExcelJS re-opens injected file intact')
  } catch (e) {
    assert(false, 'ExcelJS re-open threw: ' + e.message)
  }

  console.log(`Wrote ${outPath} for manual inspection in Excel/Google Sheets.`)
}
