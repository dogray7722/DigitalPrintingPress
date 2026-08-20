import JSZip from 'jszip'

// Rewrites ExcelJS's internal (same-workbook) hyperlinks into the shape Excel itself
// writes, which is also the only shape Google Sheets imports correctly.
//
// ExcelJS DOES detect an internal link — `cell.value = { text, hyperlink: "'SHEET'!A1" }`
// hits its `isInternalLink` branch and emits a `location` attribute — but it gets two
// things wrong:
//
//   1. It ALSO emits an `r:id` and pushes a matching `TargetMode="External"` relationship
//      whose Target is the sheet reference. A hyperlink is either internal (`location`)
//      or external (`r:id`), never both; Excel tries to resolve "SHEET!A1" as a file.
//   2. It never emits `display`. That one is why this matters for Google Sheets: without
//      `display`, Sheets renders the link's text as a literal "#gid=xxxx" instead of the
//      label (PHPOffice/PhpSpreadsheet#807 — same bug class, different library).
//
// ExcelJS writes:
//   <hyperlink ref="A4" r:id="rId1" location="'ACCOMMODATION'!A28"/>
//   <Relationship Id="rId1" Type=".../hyperlink" Target="'ACCOMMODATION'!A28" TargetMode="External"/>
//
// Excel writes (and this is what we produce):
//   <hyperlink ref="A4" location="'ACCOMMODATION'!A28" display="↓  JUMP TO RECOMMENDATIONS"/>
//
// The `display` text is recovered from the cell's own value, so this pass is fully
// self-contained: any internal hyperlink anywhere in the workbook is fixed up, with no
// registry threaded through the sheet builders.
//
// ORDERING: must run AFTER `injectChart`. Chart injection allocates its worksheet rel as
// `maxRelId(rels) + 1`; if we deleted the hyperlink rels first, that max would drop and
// the new rel could collide with the `<drawing r:id="rIdN"/>` reference ExcelJS already
// wrote into the sheet XML. Deleting them afterwards is safe — rIds needn't be
// contiguous, and nothing is renumbered.

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

/**
 * Parse `xl/sharedStrings.xml` into an index-addressed array. Values are left in their
 * XML-escaped form (`&amp;`, `&lt;`) because they are headed straight back into an
 * attribute — re-escaping would double them up.
 */
function parseSharedStrings(xml: string | null | undefined): string[] {
  if (!xml) return []
  const out: string[] = []
  // Each <si> is one string; rich text splits it across several <r><t> runs.
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g
  let si: RegExpExecArray | null
  while ((si = siRe.exec(xml))) {
    const body = si[1] ?? ''
    let text = ''
    const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g
    let t: RegExpExecArray | null
    while ((t = tRe.exec(body))) text += t[1] ?? ''
    out.push(text)
  }
  return out
}

/**
 * Recover the display text for a hyperlink from the cell it sits on. Returns null when
 * the cell has no readable text (an empty cell, or a numeric one) — in that case we emit
 * no `display` attribute rather than inventing one.
 */
function cellDisplayText(sheetXml: string, ref: string, sharedStrings: string[]): string | null {
  // Attribute order varies, and the cell may be self-closing (no value).
  const cell = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`).exec(sheetXml)
  if (!cell) return null
  const attrs = cell[1] ?? ''
  const body = cell[2]
  if (!body) return null

  const type = /\bt="([^"]*)"/.exec(attrs)?.[1]

  if (type === 's') {
    const idx = Number(/<v>(\d+)<\/v>/.exec(body)?.[1])
    return Number.isInteger(idx) ? sharedStrings[idx] ?? null : null
  }
  if (type === 'inlineStr') {
    let text = ''
    const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g
    let t: RegExpExecArray | null
    while ((t = tRe.exec(body))) text += t[1]
    return text || null
  }
  // t="str" (formula string result) and untyped/numeric cells both keep it in <v>.
  return /<v>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? null
}

/** Escape for an XML attribute. Input is already element-escaped, so only `"` is left. */
function escapeAttr(text: string): string {
  return text.replace(/"/g, '&quot;')
}

/** Escape a literal for embedding in a RegExp (the rel Type is a URL full of dots). */
function escapeRe(literal: string): string {
  return literal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Post-process a written .xlsx buffer, converting every internal hyperlink to Excel's
 * native `location` + `display` form and dropping the bogus external relationships
 * ExcelJS paired with them. Returns the rewritten buffer.
 */
export async function injectNativeHyperlinks(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer)
  const sharedStrings = parseSharedStrings(await zip.file('xl/sharedStrings.xml')?.async('string'))

  const sheetPaths = Object.keys(zip.files).filter((p) =>
    /^xl\/worksheets\/sheet\d+\.xml$/.test(p)
  )

  for (const sheetPath of sheetPaths) {
    const original = await zip.file(sheetPath)?.async('string')
    if (!original || !original.includes('<hyperlink ')) continue

    const droppedRelIds: string[] = []

    const rewritten = original.replace(/<hyperlink\b([^>]*?)\/>/g, (whole, attrs: string) => {
      const location = /\blocation="([^"]*)"/.exec(attrs)?.[1]
      // No `location` means a genuine external link — leave it exactly as it is.
      if (!location) return whole

      const ref = /\bref="([^"]*)"/.exec(attrs)?.[1]
      if (!ref) return whole

      const relId = /\br:id="([^"]*)"/.exec(attrs)?.[1]
      if (relId) droppedRelIds.push(relId)

      // ExcelJS passes the target through verbatim, so a caller who wrote a "#SHEET!A1"
      // style target would land a stray "#" in `location`, which Excel won't resolve.
      const cleanLocation = location.replace(/^#/, '')
      const tooltip = /\btooltip="([^"]*)"/.exec(attrs)?.[1]
      const display = cellDisplayText(original, ref, sharedStrings)

      return (
        `<hyperlink ref="${ref}" location="${cleanLocation}"` +
        (display !== null ? ` display="${escapeAttr(display)}"` : '') +
        (tooltip !== undefined ? ` tooltip="${tooltip}"` : '') +
        `/>`
      )
    })

    if (rewritten === original) continue
    zip.file(sheetPath, rewritten)

    if (!droppedRelIds.length) continue

    // Drop the paired external relationships. Match on Id AND the hyperlink Type so a
    // drawing/comment rel can never be caught by an id collision.
    const relsPath = sheetPath.replace(
      /worksheets\/(sheet\d+)\.xml$/,
      'worksheets/_rels/$1.xml.rels'
    )
    const rels = await zip.file(relsPath)?.async('string')
    if (!rels) continue

    let cleanedRels = rels
    for (const relId of droppedRelIds) {
      cleanedRels = cleanedRels.replace(
        new RegExp(
          `<Relationship\\b(?=[^>]*\\bId="${escapeRe(relId)}")` +
            `(?=[^>]*Type="${escapeRe(HYPERLINK_REL_TYPE)}")[^>]*?/>`,
          'g'
        ),
        ''
      )
    }
    zip.file(relsPath, cleanedRels)
  }

  return zip.generateAsync({ type: 'arraybuffer' })
}
