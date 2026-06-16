import JSZip from 'jszip'
import type { ThemeId } from '../../types/wizard'

// ExcelJS has no API to write native chart objects. To get a real, dynamic doughnut
// chart that updates as the user types — and that survives .xlsx import into Google
// Sheets — we let ExcelJS write the workbook, then surgically inject the OOXML chart
// parts (chart + drawing + rels + content-type overrides) into the .zip afterwards.

/** Native chart kinds we can inject. All three share the same JSZip plumbing. */
export type ChartKind = 'doughnut' | 'pie' | 'bar'

export interface ChartAnchor {
  /** 0-based column index of the chart's top-left corner (col F = 5). */
  fromCol: number
  /** 0-based row index of the chart's top-left corner (row 27 = 26). */
  fromRow: number
  /** 0-based column index of the chart's bottom-right corner. */
  toCol: number
  /** 0-based row index of the chart's bottom-right corner. */
  toRow: number
}

export interface InjectChartOptions {
  /** Which native chart to render. */
  kind: ChartKind
  /** Worksheet name the chart is anchored to, e.g. 'OVERVIEW'. */
  sheetName: string
  /** Category labels range, e.g. "'OVERVIEW'!$F$18:$F$23". */
  categoryRange: string
  /**
   * Numeric values range.
   * - pie/donut: helper col M — actual spend falling back to estimate (chart is never empty).
   * - bar: helper col N — MIN(actual, budget), the colored "actual spent" segment
   *   (base of the stacked bar).
   */
  valueRange: string
  /**
   * Bar chart only — remaining-budget range (col O = MAX(budget - actual, 0)).
   * Stacked after valueRange as the light diagonal-pattern "remaining" segment, so one
   * bar = full budget. Stacked bars render identically in Excel and Google Sheets
   * (clustered + overlap does not — Google Sheets ignores overlap).
   */
  remainderRange?: string
  /**
   * Bar chart only — over-budget overage range (col P = MAX(actual - budget, 0)).
   * Stacked after remainderRange as a solid red segment. Zero-width when on/under budget;
   * extends the bar past the budget length in red when a category is overspent.
   */
  overRange?: string
  /** Where the chart frame sits, in cell anchors. */
  anchor: ChartAnchor
  /** Slice/bar colors — ARGB or RGB hex strings, one per category. */
  colors: string[]
  /** Optional chart title. */
  title?: string
}

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

const REL_TYPE_DRAWING = `${R_NS}/drawing`
const REL_TYPE_CHART = `${R_NS}/chart`

const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

// Escapes XML element-text content. Only &, <, > are significant in element text;
// quotes are left literal so chart formula refs like 'OVERVIEW'!$F$1 match what Excel
// itself writes (it never entity-escapes the apostrophes that quote sheet names).
// Safe here because no user-supplied value is ever placed in an XML attribute.
function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Normalize an ARGB/RGB hex string to the 6-digit RRGGBB form OOXML's srgbClr wants. */
function toRgb(hex: string): string {
  const clean = hex.replace(/^#/, '')
  // Strip a leading alpha byte if an 8-char ARGB value was passed.
  return (clean.length === 8 ? clean.slice(2) : clean).toUpperCase()
}

/** Highest existing rId number in a .rels file (0 if none / file absent). */
function maxRelId(relsXml: string | null): number {
  if (!relsXml) return 0
  let max = 0
  const re = /Id="rId(\d+)"/g
  let m: RegExpExecArray | null
  while ((m = re.exec(relsXml))) {
    max = Math.max(max, parseInt(m[1], 10))
  }
  return max
}

// Axis IDs for the bar chart's category/value axes — arbitrary but must be stable
// and cross-reference each other.
const BAR_CAT_AX_ID = '111111111'
const BAR_VAL_AX_ID = '222222222'

function buildChartXml(opts: InjectChartOptions): string {
  // Per-category colors so each slice/bar picks up the theme palette.
  const dPts = opts.colors
    .map(
      (c, i) =>
        `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>` +
        `<c:spPr><a:solidFill><a:srgbClr val="${toRgb(c)}"/></a:solidFill></c:spPr></c:dPt>`
    )
    .join('')

  const title = opts.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a:r><a:t>${escapeXml(opts.title)}</a:t></a:r></a:p></c:rich></c:tx>` +
      `<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : `<c:autoTitleDeleted val="1"/>`

  const catRef = `<c:cat><c:strRef><c:f>${escapeXml(opts.categoryRange)}</c:f></c:strRef></c:cat>`
  const valRef = `<c:val><c:numRef><c:f>${escapeXml(opts.valueRange)}</c:f></c:numRef></c:val>`

  // Pie/donut: category name + percentage; all three bar series: no labels
  // (the table below the chart already shows the numbers).
  const pieLbls =
    `<c:dLbls>` +
    `<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/>` +
    `<c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>` +
    `</c:dLbls>`
  const noLbls =
    `<c:dLbls>` +
    `<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>` +
    `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>` +
    `</c:dLbls>`

  let plotBody: string
  let legend: string

  if (opts.kind === 'bar') {
    // 3-series horizontal STACKED bar chart (progress-bar technique):
    //
    //   idx=0 (base, left) : actual spend MIN(actual, budget) — per-category theme colors
    //   idx=1 (stacked)    : remaining budget MAX(budget - actual, 0) — light diagonal track
    //   idx=2 (stacked)    : over-budget overage MAX(actual - budget, 0) — solid red
    //
    // idx=0 + idx=1 sum to the budget, so a within-budget category is ONE bar whose length =
    // budget, with a colored "spent" portion and a light "remaining" portion.
    //
    // When H=0:   colored segment = 0, gray fills the whole bar; no red.
    // When H<G:   colored "spent" + gray "remaining" = budget; no red.
    // When H=G:   colored fills the budget; no gray, no red.
    // When H>G:   colored caps at budget, gray = 0, red overage extends the bar past
    //             the budget length — the visual "over budget" cue.
    //
    // Stacked is used instead of clustered + overlap="100" because Google Sheets ignores
    // overlap and renders clustered series as two separate bars; stacked renders the same
    // single bar in both Excel and Google Sheets.

    const remainderFill =
      `<c:spPr>` +
      `<a:pattFill prst="ltDnDiag">` +
      `<a:fgClr><a:srgbClr val="B0BEC5"/></a:fgClr>` +
      `<a:bgClr><a:srgbClr val="ECEFF1"/></a:bgClr>` +
      `</a:pattFill>` +
      `<a:ln><a:solidFill><a:srgbClr val="B0BEC5"/></a:solidFill></a:ln>` +
      `</c:spPr>`

    // Solid red overage segment, with a slightly darker red outline.
    const overFill =
      `<c:spPr>` +
      `<a:solidFill><a:srgbClr val="E53935"/></a:solidFill>` +
      `<a:ln><a:solidFill><a:srgbClr val="C62828"/></a:solidFill></a:ln>` +
      `</c:spPr>`

    const remainderValRef = `<c:val><c:numRef><c:f>${escapeXml(opts.remainderRange ?? opts.valueRange)}</c:f></c:numRef></c:val>`
    const overValRef = `<c:val><c:numRef><c:f>${escapeXml(opts.overRange ?? opts.valueRange)}</c:f></c:numRef></c:val>`

    plotBody =
      `<c:barChart>` +
      `<c:barDir val="bar"/>` +
      `<c:grouping val="stacked"/>` +
      `<c:varyColors val="0"/>` +
      // idx=0: actual spend (base of the stack, theme colors)
      `<c:ser><c:idx val="0"/><c:order val="0"/>` +
      dPts +
      noLbls +
      catRef +
      valRef +
      `</c:ser>` +
      // idx=1: remaining budget (stacked after, light diagonal track)
      `<c:ser><c:idx val="1"/><c:order val="1"/>` +
      remainderFill +
      noLbls +
      catRef +
      remainderValRef +
      `</c:ser>` +
      // idx=2: over-budget overage (stacked last, solid red)
      `<c:ser><c:idx val="2"/><c:order val="2"/>` +
      overFill +
      noLbls +
      catRef +
      overValRef +
      `</c:ser>` +
      `<c:gapWidth val="100"/>` +
      `<c:overlap val="100"/>` +
      `<c:axId val="${BAR_CAT_AX_ID}"/>` +
      `<c:axId val="${BAR_VAL_AX_ID}"/>` +
      `</c:barChart>` +
      // catAx orientation = maxMin so the FIRST category renders at the TOP of the bar
      // chart, matching the OVERVIEW table's top-to-bottom row order. (A default minMax
      // horizontal bar chart plots category 1 at the bottom, reversing the table.)
      `<c:catAx>` +
      `<c:axId val="${BAR_CAT_AX_ID}"/>` +
      `<c:scaling><c:orientation val="maxMin"/></c:scaling>` +
      `<c:delete val="0"/><c:axPos val="l"/>` +
      `<c:crossAx val="${BAR_VAL_AX_ID}"/>` +
      `</c:catAx>` +
      // valAx crosses at max so the value scale stays along the BOTTOM after the category
      // axis is flipped (otherwise it jumps to the top).
      `<c:valAx>` +
      `<c:axId val="${BAR_VAL_AX_ID}"/>` +
      `<c:scaling><c:orientation val="minMax"/></c:scaling>` +
      `<c:delete val="0"/><c:axPos val="b"/>` +
      `<c:crossAx val="${BAR_CAT_AX_ID}"/>` +
      `<c:crosses val="max"/>` +
      `</c:valAx>`
    legend = ''
  } else {
    // pie and doughnut share everything but the wrapper element + holeSize.
    const sliceTail =
      opts.kind === 'doughnut'
        ? `<c:firstSliceAng val="0"/><c:holeSize val="50"/>`
        : `<c:firstSliceAng val="0"/>`
    const wrapper = opts.kind === 'doughnut' ? 'doughnutChart' : 'pieChart'
    plotBody =
      `<c:${wrapper}>` +
      `<c:varyColors val="1"/>` +
      `<c:ser><c:idx val="0"/><c:order val="0"/>` +
      dPts +
      pieLbls +
      catRef +
      valRef +
      `</c:ser>` +
      sliceTail +
      `</c:${wrapper}>`
    legend = `<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>`
  }

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    `<c:chart>` +
    title +
    `<c:plotArea><c:layout/>` +
    plotBody +
    `</c:plotArea>` +
    legend +
    // plotVisOnly=0 so the chart still plots its hidden helper column (col M).
    `<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/>` +
    `</c:chart>` +
    `</c:chartSpace>`
  )
}

/**
 * Builds just the <xdr:twoCellAnchor>...</xdr:twoCellAnchor> fragment for the chart
 * graphicFrame. Reused both when creating a brand-new drawing part (wrapped in
 * <xdr:wsDr> by buildDrawingXml below) and when merging into an existing drawing part
 * (inserted as a sibling before its closing </xdr:wsDr> — see injectChart's merge
 * branch).
 *
 * `cNvPr id="2"` only needs to be unique within the drawing part. The cover photo's
 * <xdr:pic> (written by ExcelJS, if present) uses id="1", so "2" is safe — revisit if a
 * second image is ever added to the same drawing.
 */
function buildChartAnchorXml(anchor: ChartAnchor, chartRelId: string): string {
  return (
    `<xdr:twoCellAnchor>` +
    `<xdr:from><xdr:col>${anchor.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff>` +
    `<xdr:row>${anchor.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${anchor.toCol}</xdr:col><xdr:colOff>0</xdr:colOff>` +
    `<xdr:row>${anchor.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame macro="">` +
    `<xdr:nvGraphicFramePr>` +
    `<xdr:cNvPr id="2" name="Budget Chart"/><xdr:cNvGraphicFramePr/>` +
    `</xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${C_NS}">` +
    `<c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${chartRelId}"/>` +
    `</a:graphicData></a:graphic>` +
    `</xdr:graphicFrame>` +
    `<xdr:clientData/>` +
    `</xdr:twoCellAnchor>`
  )
}

/**
 * Wraps a single chart anchor fragment in a brand-new <xdr:wsDr> document. Used only
 * when the worksheet has no existing drawing part (i.e. no cover photo was embedded).
 */
function buildDrawingXml(anchor: ChartAnchor, chartRelId: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">` +
    buildChartAnchorXml(anchor, chartRelId) +
    `</xdr:wsDr>`
  )
}

function buildDrawingRelsXml(chartRelId: string, chartTarget: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="${chartRelId}" Type="${REL_TYPE_CHART}" Target="${chartTarget}"/>` +
    `</Relationships>`
  )
}

/** Resolve the OVERVIEW worksheet's part path via workbook.xml + its rels. */
async function resolveSheetPath(zip: JSZip, sheetName: string): Promise<string | null> {
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string')
  if (!workbookXml || !relsXml) return null

  // <sheet name="OVERVIEW" sheetId="1" r:id="rId1"/> — attribute order varies.
  const sheetTag = new RegExp(
    `<sheet\\b[^>]*\\bname="${sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`,
    'i'
  ).exec(workbookXml)
  if (!sheetTag) return null
  const ridMatch = /r:id="(rId\d+)"/.exec(sheetTag[0])
  if (!ridMatch) return null
  const rid = ridMatch[1]

  const relMatch = new RegExp(
    `<Relationship\\b[^>]*\\bId="${rid}"[^>]*\\bTarget="([^"]+)"[^>]*/>`,
    'i'
  ).exec(relsXml)
  if (!relMatch) {
    // Try Target before Id (attribute order can swap).
    const alt = new RegExp(
      `<Relationship\\b[^>]*\\bTarget="([^"]+)"[^>]*\\bId="${rid}"[^>]*/>`,
      'i'
    ).exec(relsXml)
    if (!alt) return null
    return normalizeSheetTarget(alt[1])
  }
  return normalizeSheetTarget(relMatch[1])
}

function normalizeSheetTarget(target: string): string {
  const t = target.replace(/^\//, '').replace(/^xl\//, '')
  return `xl/${t}`
}

/** Pick the smallest N such that xl/{dir}/{base}N.xml is unused. */
function nextFreeIndex(zip: JSZip, dir: string, base: string): number {
  let n = 1
  while (zip.file(`xl/${dir}/${base}${n}.xml`)) n++
  return n
}

/**
 * ExcelJS always emits a self-closing `<drawing r:id="rIdN"/>` as the last child before
 * `</worksheet>` when the worksheet has any image/drawing (e.g. the cover photo).
 * Returns the referenced rId, or null if the worksheet has no drawing yet.
 */
function findExistingDrawingRelId(sheetXml: string): string | null {
  const m = /<drawing r:id="(rId\d+)"\/>/.exec(sheetXml)
  return m ? m[1] : null
}

/**
 * Resolves a worksheet `<drawing r:id="rIdN"/>` reference to its drawing part's path
 * via the worksheet's `_rels` file, e.g. "xl/drawings/drawing1.xml".
 */
function resolveDrawingPath(sheetRelsXml: string, relId: string): string | null {
  let m = new RegExp(
    `<Relationship\\b[^>]*\\bId="${relId}"[^>]*\\bTarget="([^"]+)"[^>]*/>`,
    'i'
  ).exec(sheetRelsXml)
  if (!m) {
    // Try Target before Id (attribute order can swap).
    m = new RegExp(
      `<Relationship\\b[^>]*\\bTarget="([^"]+)"[^>]*\\bId="${relId}"[^>]*/>`,
      'i'
    ).exec(sheetRelsXml)
  }
  if (!m) return null
  // Target is relative to xl/worksheets/_rels/, e.g. "../drawings/drawing1.xml".
  const fileName = m[1].split('/').pop() as string
  return `xl/drawings/${fileName}`
}

export async function injectChart(
  buffer: ArrayBuffer | Uint8Array,
  opts: InjectChartOptions
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer)

  const sheetPath = await resolveSheetPath(zip, opts.sheetName)
  if (!sheetPath) {
    throw new Error(`injectChart: could not resolve sheet "${opts.sheetName}"`)
  }

  const sheetXml = await zip.file(sheetPath)?.async('string')
  if (!sheetXml) {
    throw new Error(`injectChart: sheet part missing at ${sheetPath}`)
  }
  if (!/\sxmlns:r=/.test(sheetXml)) {
    throw new Error('injectChart: worksheet is missing the r: namespace declaration')
  }

  const chartIdx = nextFreeIndex(zip, 'charts', 'chart')
  const chartFile = `xl/charts/chart${chartIdx}.xml`

  // 1. Chart part (same in both branches).
  zip.file(chartFile, buildChartXml(opts))

  const ctPath = '[Content_Types].xml'
  const ctXml = await zip.file(ctPath)?.async('string')
  if (!ctXml) throw new Error('injectChart: [Content_Types].xml missing')

  const sheetBase = sheetPath.replace(/^xl\//, '') // e.g. worksheets/sheet1.xml
  const sheetFile = sheetBase.split('/').pop() as string
  const sheetRelsPath = `xl/worksheets/_rels/${sheetFile}.rels`

  const existingDrawingRelId = findExistingDrawingRelId(sheetXml)

  if (existingDrawingRelId) {
    // ── MERGE branch: the worksheet's single <drawing> slot is already taken (e.g. by
    //    the cover photo ExcelJS embedded). Add the chart as a second anchor inside
    //    that SAME drawing part instead of creating a second <drawing> element, which
    //    CT_Worksheet does not allow.
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async('string')
    if (!sheetRelsXml) {
      throw new Error(
        `injectChart: expected sheet rels at ${sheetRelsPath} for existing drawing ${existingDrawingRelId}`
      )
    }
    const drawingPath = resolveDrawingPath(sheetRelsXml, existingDrawingRelId)
    if (!drawingPath) {
      throw new Error(`injectChart: could not resolve drawing path for ${existingDrawingRelId}`)
    }
    const drawingXml = await zip.file(drawingPath)?.async('string')
    if (!drawingXml) {
      throw new Error(`injectChart: drawing part missing at ${drawingPath}`)
    }
    if (!drawingXml.includes('</xdr:wsDr>')) {
      throw new Error(`injectChart: ${drawingPath} is missing </xdr:wsDr> — cannot merge chart anchor`)
    }

    const drawingFileName = drawingPath.replace(/^xl\/drawings\//, '')
    const drawingRelsPath = `xl/drawings/_rels/${drawingFileName}.rels`
    const existingDrawingRels = (await zip.file(drawingRelsPath)?.async('string')) ?? null
    const chartRelId = `rId${maxRelId(existingDrawingRels) + 1}`

    // 2. Insert the chart's <xdr:twoCellAnchor> as a new sibling before </xdr:wsDr>.
    const chartAnchorXml = buildChartAnchorXml(opts.anchor, chartRelId)
    zip.file(drawingPath, drawingXml.replace('</xdr:wsDr>', `${chartAnchorXml}</xdr:wsDr>`))

    // 3. Append the chart relationship into the existing drawing's rels (create if absent).
    const chartRelEntry = `<Relationship Id="${chartRelId}" Type="${REL_TYPE_CHART}" Target="../charts/chart${chartIdx}.xml"/>`
    if (existingDrawingRels) {
      zip.file(drawingRelsPath, existingDrawingRels.replace('</Relationships>', `${chartRelEntry}</Relationships>`))
    } else {
      zip.file(
        drawingRelsPath,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<Relationships xmlns="${PKG_REL_NS}">${chartRelEntry}</Relationships>`
      )
    }

    // 4. Content types — only the chart part is new (the drawing's Override was already
    //    added by ExcelJS when it wrote the cover photo).
    zip.file(ctPath, ctXml.replace('</Types>', `<Override PartName="/${chartFile}" ContentType="${CT_CHART}"/></Types>`))
  } else {
    // ── CREATE branch: no existing drawing on this sheet — original behavior.
    const drawingIdx = nextFreeIndex(zip, 'drawings', 'drawing')
    const drawingFile = `xl/drawings/drawing${drawingIdx}.xml`
    const drawingRelsFile = `xl/drawings/_rels/drawing${drawingIdx}.xml.rels`

    // 2. Drawing part (chart referenced via rId1, local to the drawing's rels).
    zip.file(drawingFile, buildDrawingXml(opts.anchor, 'rId1'))

    // 3. Drawing rels → chart.
    zip.file(drawingRelsFile, buildDrawingRelsXml('rId1', `../charts/chart${chartIdx}.xml`))

    // 4. Worksheet rels → drawing (create or extend).
    const existingRels = (await zip.file(sheetRelsPath)?.async('string')) ?? null
    const drawingRelId = `rId${maxRelId(existingRels) + 1}`
    const drawingRelEntry =
      `<Relationship Id="${drawingRelId}" Type="${REL_TYPE_DRAWING}" ` +
      `Target="../drawings/drawing${drawingIdx}.xml"/>`

    if (existingRels) {
      zip.file(sheetRelsPath, existingRels.replace('</Relationships>', `${drawingRelEntry}</Relationships>`))
    } else {
      zip.file(
        sheetRelsPath,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<Relationships xmlns="${PKG_REL_NS}">${drawingRelEntry}</Relationships>`
      )
    }

    // 5. Worksheet XML → <drawing> element. Must be the last child before </worksheet>
    //    (it precedes only legacyDrawing/picture/oleObjects, which ExcelJS doesn't emit here).
    const drawingEl = `<drawing r:id="${drawingRelId}"/>`
    zip.file(sheetPath, sheetXml.replace('</worksheet>', `${drawingEl}</worksheet>`))

    // 6. Content types → declare the new parts.
    const overrides =
      `<Override PartName="/${drawingFile}" ContentType="${CT_DRAWING}"/>` +
      `<Override PartName="/${chartFile}" ContentType="${CT_CHART}"/>`
    zip.file(ctPath, ctXml.replace('</Types>', `${overrides}</Types>`))
  }

  return zip.generateAsync({ type: 'arraybuffer' })
}

/**
 * Default/fallback palette for the budget chart. Order matches the OVERVIEW budget
 * categories (Transportation, Accommodation, Food, Activities, Shopping, Miscellaneous).
 * Per-theme palettes live in CHART_PALETTES below; this fixed set is the neutral default
 * and is used by scripts/verify-chart-injection.mjs.
 */
export const CHART_COLORS = [
  'FF4EA8DE', // Transportation — sky blue
  'FF4ECDC4', // Accommodation — teal
  'FFFFE66D', // Food — amber
  'FFA06CD5', // Activities — purple
  'FFF0865C', // Shopping — coral
  'FFB0B0B0', // Miscellaneous — gray
]

/**
 * Theme-harmonized palettes — one 6-color set per wizard theme, in category order
 * (Transportation, Accommodation, Food, Activities, Shopping, Miscellaneous). Keyed by
 * ThemeId so TypeScript keeps it in sync with the theme list.
 */
export const CHART_PALETTES: Record<ThemeId, string[]> = {
  sakura: ['FF7B5EA8', 'FFEFA8C4', 'FFA87BD4', 'FFC9A0E0', 'FFD98CB3', 'FFB0A4C8'], // lavenders → plum + rose
  ocean: ['FF4ECDC4', 'FFF0A868', 'FF5E9EC8', 'FF7FD4D0', 'FFE08A6A', 'FFA8C8E8'],  // teals + warm accents + blues
  forest: ['FF5E9E78', 'FFE0B062', 'FF8AB87A', 'FFA8D4B8', 'FFCBA15A', 'FFB7C4A0'], // greens + sage + golds
  desert: ['FFD4A050', 'FFC0613A', 'FFE0C078', 'FF9E8C5A', 'FFD98C5A', 'FFE8CFA8'], // warm earth tones
  inkwell: ['FFC9A96E', 'FFA0784A', 'FFE0C68A', 'FF8A7860', 'FFD4915C', 'FFBFAE90'], // golds + warm browns, dark ink palette
  parchment: ['FF7A5A18', 'FFAD8C56', 'FFC9A96E', 'FF6A5030', 'FFBFAE90', 'FF8A7860'], // golds + warm browns, light ink palette
}
