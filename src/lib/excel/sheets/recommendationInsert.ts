import type ExcelJS from 'exceljs'
import type { ThemeStyles } from '../styleFactory'
import type { RegionRec, PlaceRec } from '../../recommendations/types'
import { styleSectionHeader, styleColumnHeader, styleDataRow, truncate, wrappedLineCount, rowHeightForLines, contrastRatio, AA_CONTRAST } from '../styleFactory'

// Enumerate single-letter columns from `from` to `to` inclusive (e.g. 'E'..'H').
function colRange(from: string, to: string): string[] {
  const out: string[] = []
  for (let c = from.charCodeAt(0); c <= to.charCodeAt(0); c++) out.push(String.fromCharCode(c))
  return out
}

// ★ and → are plain BMP characters, not emoji, so they take the cell's font color and
// render as part of the cream label. (An emoji ⭐ would render in its own gold and ignore
// the font color.) They stay inside the single label string: a cell can carry one
// hyperlink over its whole value, so splitting into rich-text runs would lose the link.
const LINK_TEXT = '★   Recommendations   →'

export type InsertKind = 'hotels' | 'restaurants' | 'excursions'

interface KindConfig {
  title: string
  lastCol: string // description spans E..lastCol; region banner spans A..lastCol
  nameHeader: string
  pick: (r: RegionRec) => PlaceRec[]
  // Two attribute columns (C and D), each pulling a field off the PlaceRec
  attr1: { header: string; field: keyof PlaceRec }
  attr2: { header: string; field: keyof PlaceRec }
}

const CONFIG: Record<InsertKind, KindConfig> = {
  hotels: {
    title: 'RECOMMENDED ACCOMMODATION BY REGION',
    lastCol: 'G',
    nameHeader: 'Accommodation',
    pick: (r) => r.hotels,
    attr1: { header: 'Price', field: 'price' },
    attr2: { header: 'Stars', field: 'rating' },
  },
  restaurants: {
    title: 'RECOMMENDED DINING BY REGION',
    lastCol: 'H',
    nameHeader: 'Restaurant',
    pick: (r) => r.restaurants,
    attr1: { header: 'Cuisine', field: 'cuisine' },
    attr2: { header: 'Price', field: 'price' },
  },
  excursions: {
    title: 'RECOMMENDED EXCURSIONS BY REGION',
    lastCol: 'H',
    nameHeader: 'Excursion',
    pick: (r) => r.excursions,
    attr1: { header: 'Duration', field: 'duration' },
    attr2: { header: 'Price', field: 'price' },
  },
}

/**
 * Renders an informational, AI-generated guide table below the working tracker on a
 * tab. Places are grouped under a region/city banner (e.g. "Cayo District"). This is
 * reference content only — it never feeds budget formulas and the user keeps logging
 * their own bookings in the tracker above. Mirrors the donut "insert" idea: a self
 * contained graphic dropped onto an existing sheet.
 *
 * @param startRow first row to write to (caller leaves a spacer above it)
 * @returns the section-band row (=== startRow) when the guide rendered, else null.
 *   The guide no-ops on two conditions, so callers can't just test `regions` to decide
 *   whether to link down to it — they must gate on this return value or ship a link
 *   pointing at an empty row.
 */
export function buildRecommendationInsert(
  ws: ExcelJS.Worksheet,
  startRow: number,
  ts: ThemeStyles,
  regions: RegionRec[] | undefined,
  kind: InsertKind
): number | null {
  if (!regions?.length) return null
  const cfg = CONFIG[kind]

  // Keep only regions that actually have at least one place of this kind.
  const relevant = regions
    .map((r) => ({ region: r.region, description: r.description, places: cfg.pick(r) }))
    .filter((r) => r.places.length > 0)
  if (!relevant.length) return null

  // Merged-cell widths (in chars) so we can size row heights — Excel won't auto-grow
  // merged rows. Widths were already set by the caller, so read them back.
  const colW = (c: string) => ws.getColumn(c).width ?? 10
  const sumW = (cols: string[]) => cols.reduce((s, c) => s + colW(c), 0)
  const nameWidth = sumW(['A', 'B'])
  const descWidth = sumW(colRange('E', cfg.lastCol))
  const fullWidth = sumW(colRange('A', cfg.lastCol))

  let row = startRow

  // ── Section header band ─────────────────────────────────────────────────────
  ws.getCell(`A${row}`).value = `${cfg.title}  ·  Travel guide`
  ws.mergeCells(`A${row}:${cfg.lastCol}${row}`)
  styleSectionHeader(ws.getRow(row), ts)
  row++

  // ── Subtitle / disclaimer ───────────────────────────────────────────────────
  const noteText =
    'Informational suggestions only — use the tracker above to log your actual bookings.'
  const noteCell = ws.getCell(`A${row}`)
  noteCell.value = noteText
  ws.mergeCells(`A${row}:${cfg.lastCol}${row}`)
  noteCell.font = { name: ts.fontName, size: ts.sizes.data - 1, italic: true, color: { argb: 'FF888888' } }
  noteCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }
  ws.getRow(row).height = rowHeightForLines(wrappedLineCount(noteText, fullWidth), ts, { maxLines: 2 })
  row++

  // ── Column headers ──────────────────────────────────────────────────────────
  ws.getCell(`A${row}`).value = cfg.nameHeader
  ws.mergeCells(`A${row}:B${row}`)
  ws.getCell(`C${row}`).value = cfg.attr1.header
  ws.getCell(`D${row}`).value = cfg.attr2.header
  ws.getCell(`E${row}`).value = 'Description'
  ws.mergeCells(`E${row}:${cfg.lastCol}${row}`)
  styleColumnHeader(ws.getRow(row), ts)
  row++

  // ── Region groups ───────────────────────────────────────────────────────────
  let stripe = 0
  relevant.forEach(({ region, description, places }) => {
    // Region banner row (spans the full table width). Wrap + size it so long region
    // blurbs aren't clipped at the merge boundary.
    const bannerText = description
      ? `📍  ${region}  —  ${truncate(description, 150)}`
      : `📍  ${region}`
    const bannerCell = ws.getCell(`A${row}`)
    bannerCell.value = bannerText
    ws.mergeCells(`A${row}:${cfg.lastCol}${row}`)
    bannerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ts.palette.mediumBg } } as ExcelJS.Fill
    bannerCell.font = { name: ts.fontName, size: ts.sizes.header, bold: true }
    bannerCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }
    ws.getRow(row).height = rowHeightForLines(wrappedLineCount(bannerText, fullWidth), ts, { min: 22, maxLines: 3 })
    row++

    places.forEach((place) => {
      const r = row
      const name = truncate(place.name, 80)
      const desc = truncate(place.description, 240)
      ws.getCell(`A${r}`).value = name
      ws.mergeCells(`A${r}:B${r}`)
      ws.getCell(`C${r}`).value = (place[cfg.attr1.field] as string | undefined) ?? ''
      ws.getCell(`D${r}`).value = (place[cfg.attr2.field] as string | undefined) ?? ''
      const descCell = ws.getCell(`E${r}`)
      descCell.value = desc
      ws.mergeCells(`E${r}:${cfg.lastCol}${r}`)

      styleDataRow(ws.getRow(r), ts, stripe % 2 === 0)
      stripe++

      // Name bold + wraps; description wraps. Size row to the taller of the two
      // merged cells (Excel won't auto-grow merged rows).
      ws.getCell(`A${r}`).font = { name: ts.fontName, size: ts.sizes.data, bold: true }
      ws.getCell(`A${r}`).alignment = { vertical: 'top', horizontal: 'left', wrapText: true }
      descCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true }
      const lines = Math.max(wrappedLineCount(name, nameWidth), wrappedLineCount(desc, descWidth))
      ws.getRow(r).height = rowHeightForLines(lines, ts, { min: 22, maxLines: 8 })
      row++
    })
  })

  return startRow
}

/**
 * Reserve the trailing cells of a tracker's row-2 description band for the
 * recommendations button, merging them and painting the band's own fill so the strip
 * still reads as one continuous bar when there is no guide to link to.
 *
 * Call this unconditionally, right after `styleSectionHeader(ws.getRow(2), ts)` — that
 * helper skips empty cells, so without this the reserved range would be a white notch at
 * the end of the band whenever recommendations are off.
 *
 * @param range two-cell range at the end of the band, e.g. 'F2:G2'
 */
export function reserveRecommendationsSlot(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  range: string
): void {
  ws.mergeCells(range)
  ws.getCell(range.split(':')[0]).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: ts.palette.secondary },
  } as ExcelJS.Fill
}

/**
 * An internal link at the end of the tracker's row-2 band, jumping to the guide's
 * section band so it's reachable without scrolling past the (mostly empty) template
 * rows. Row 2 is inside every tracker's frozen pane (ySplit 5), so it stays on screen
 * at any scroll depth.
 *
 * Styled as a **button chip** that overrides the band underneath it: dark `primary`
 * fill against the band's mid-tone `secondary`, bold accent text with a matching thin
 * frame on all four sides, centered so it reads as a composed button rather than
 * edge-to-edge text. No underline — the fill, frame and ★/→ already say "clickable",
 * and the underline read as a dated hyperlink squiggle.
 * Must be called AFTER `styleSectionHeader` has run on the row so these overrides win.
 *
 * **The accent color is chosen by contrast, not fixed.** Tinting the text with the
 * band's own `secondary` ties the chip to the band and looks great on `inkwell`
 * (gold on charcoal, 6.7:1) — but that pairing is illegible on every other theme
 * (1.6–2.8:1, below even the 3:1 large-text floor; `desert` is the worst). So we
 * measure and fall back to `primaryText`, which clears AA everywhere. Any theme added
 * later is handled automatically instead of silently shipping unreadable text.
 *
 * The frame follows the same decision rather than staying `secondary` unconditionally:
 * a frame at 1.6:1 isn't a frame, and a cream chip with an invisible gold edge just
 * looks unframed.
 *
 * Only call this when `buildRecommendationInsert` returned a row; a link to a blank row
 * is worse than no link.
 *
 * Uses a NATIVE internal hyperlink (`{ text, hyperlink }`), NOT a `HYPERLINK()` formula.
 * This is load-bearing and must not be "simplified" back to a formula: the formula's
 * `#'SHEET'!A1` fragment navigates in Excel but is inert in Google Sheets, which was
 * confirmed against a real Drive import. The native form works in both once
 * `nativeHyperlinks.ts` rewrites it into Excel's `location` + `display` shape, and it
 * sidesteps the Recalculation Contract entirely — no formula, no cached `result:`.
 */
export function buildRecommendationsLink(
  ws: ExcelJS.Worksheet,
  ts: ThemeStyles,
  ref: string,
  targetRow: number
): void {
  const cell = ws.getCell(ref)
  cell.value = { text: LINK_TEXT, hyperlink: `'${ws.name}'!A${targetRow}` }
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: ts.palette.primary },
  } as ExcelJS.Fill

  // Prefer the band's own color as the accent; fall back to primaryText where that
  // pairing can't be read against the chip fill. On inkwell this resolves to the
  // intended gold-on-charcoal.
  const accent =
    contrastRatio(ts.palette.secondary, ts.palette.primary) >= AA_CONTRAST
      ? ts.palette.secondary
      : ts.palette.primaryText

  cell.font = {
    name: ts.fontName,
    size: ts.sizes.sectionHeader,
    bold: true,
    color: { argb: accent },
  }
  cell.alignment = { vertical: 'middle', horizontal: 'center' }
  const edge = { style: 'thin' as const, color: { argb: accent } }
  cell.border = { top: edge, bottom: edge, left: edge, right: edge } as ExcelJS.Borders
}
