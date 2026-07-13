import JSZip from 'jszip'

// ExcelJS never writes xl/calcChain.xml (the calculation-chain part every
// Excel-native workbook carries). Its absence breaks chart repainting in Excel
// for Mac: the injected chart redraws one calculation BEHIND the cells — each
// typed entry shows the previous entry's values — because Excel orders its
// chart-refresh hooks relative to the loaded chain. Diffing our file against
// the same workbook re-saved by Excel showed calcChain.xml (plus a normalized
// <calcPr>) as the only calc-relevant difference, and adding just these two
// fixed the lag (verified by typing tests in Excel for Mac 16.110).
//
// Like chartInjection, this post-processes the written .xlsx with JSZip.

// The calc-engine version Excel 16.x writes. With calcId present (and >= the
// opening app's engine) Excel trusts the cached formula results we serialize,
// so every formula cell MUST carry a correct cached `result:` value — which is
// also why fullCalcOnLoad is no longer set (a full rebuild on load re-broke
// the chart repaint ordering; see workbookConfig.ts).
const CALC_ID = '191029'

/** Map each worksheet part path to its sheetId (calcChain's `i` attribute). */
async function resolveSheetIds(zip: JSZip): Promise<Map<string, number>> {
  const workbookXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? ''
  const relsXml = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? ''

  const relTargets = new Map<string, string>()
  for (const m of relsXml.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relTargets.set(m[1], m[2])
  }

  const paths = new Map<string, number>()
  for (const m of workbookXml.matchAll(/<sheet [^>]*sheetId="(\d+)"[^>]*r:id="([^"]+)"/g)) {
    const target = relTargets.get(m[2])
    if (!target) continue
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`
    paths.set(path, parseInt(m[1], 10))
  }
  return paths
}

/**
 * Adds xl/calcChain.xml (listing every formula cell) and pins <calcPr> to a
 * native calcId. Call AFTER injectChart — both steps only touch distinct parts,
 * but this keeps the final calcPr authoritative.
 */
export async function injectCalcChain(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer)

  // Enumerate formula cells straight from the sheet XML (ExcelJS writes <f> as
  // the first child of a formula <c>), so the chain exactly matches the file —
  // no risk of drift from merged-cell or shared-formula API quirks.
  const entries: string[] = []
  for (const [path, sheetId] of await resolveSheetIds(zip)) {
    const sheetXml = await zip.file(path)?.async('string')
    if (!sheetXml) continue
    for (const m of sheetXml.matchAll(/<c r="([A-Z]+\d+)"[^>]*><f[\s>/]/g)) {
      entries.push(`<c r="${m[1]}" i="${sheetId}"/>`)
    }
  }

  if (entries.length > 0) {
    // Reverse document order — the pattern Excel itself writes (roughly
    // dependents-first). Excel self-corrects ordering at runtime; it only
    // needs the part to exist and be well-formed.
    entries.reverse()
    zip.file(
      'xl/calcChain.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        entries.join('') +
        `</calcChain>`
    )

    const ctPath = '[Content_Types].xml'
    const ctXml = (await zip.file(ctPath)?.async('string')) ?? ''
    if (!ctXml.includes('/xl/calcChain.xml')) {
      zip.file(
        ctPath,
        ctXml.replace(
          '</Types>',
          `<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`
        )
      )
    }

    const relsPath = 'xl/_rels/workbook.xml.rels'
    const relsXml = (await zip.file(relsPath)?.async('string')) ?? ''
    if (!relsXml.includes('relationships/calcChain')) {
      let maxId = 0
      for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) {
        maxId = Math.max(maxId, parseInt(m[1], 10))
      }
      zip.file(
        relsPath,
        relsXml.replace(
          '</Relationships>',
          `<Relationship Id="rId${maxId + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>`
        )
      )
    }
  }

  // Pin calcPr to the native engine id (ExcelJS hardcodes calcId="171027" and
  // there is no API to change it) and ensure fullCalcOnLoad stays absent.
  const wbXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? ''
  zip.file('xl/workbook.xml', wbXml.replace(/<calcPr[^/]*\/>/, `<calcPr calcId="${CALC_ID}"/>`))

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
}
