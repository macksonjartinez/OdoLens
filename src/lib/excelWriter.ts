/**
 * excelWriter.ts — JSZip-based binary-faithful output
 *
 * Strategy: read the template file as a raw ZIP, update ONLY the <v> element
 * content inside worksheet XMLs, keep every other file (styles.xml, drawings/,
 * media/, calcChain.xml, printerSettings, etc.) completely untouched, then
 * re-pack and download.  This preserves 100% of the original formatting.
 */

import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import type {
  ConsolidatedOutput,
  InformeEconomicoSheet,
  LabeledSheet,
  LabeledSheetKey,
  MonitoreoSheet,
} from '../types'
import { normalizeName, todayISO } from './normalize'
import { TEMPLATE_V2026_BASE64 } from '../assets/templateV2026'
import { resolveRowId } from '../schema/catalog'
import {
  ECONOMICO_SHEET,
  INPUT_COLS,
  ODOLENS_OUTPUT_MARKER,
  findBlock,
  findEconomicoSignatureRow,
} from '../schema/reportSchemas'
import type { DataBlock } from '../schema/reportSchemas'

// ── Column / cell reference helpers ─────────────────────────────────────────

function colLetters(idx: number): string {
  let s = ''
  let n = idx + 1
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function colIndex(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

function encodeCellRef(col: number, row: number): string {
  return `${colLetters(col)}${row + 1}`
}

function parseCellRef(ref: string): { col: number; row: number } | null {
  const m = ref.match(/^([A-Z]+)(\d+)$/)
  return m ? { col: colIndex(m[1]), row: parseInt(m[2]) - 1 } : null
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

// ── Shared strings ───────────────────────────────────────────────────────────

function parseSharedStrings(xml: string): string[] {
  const result: string[] = []
  const siRe = /<si(?:\s[^>]*)?>([^]*?)<\/si>/g
  let m: RegExpExecArray | null
  while ((m = siRe.exec(xml)) !== null) {
    const si = m[1]
    let text = ''
    const tRe = /<t(?:\s[^>]*)?>([^<]*)<\/t>/g
    let tm: RegExpExecArray | null
    while ((tm = tRe.exec(si)) !== null) text += tm[1]
    result.push(text)
  }
  return result
}

// ── Workbook sheet map  (sheet name → worksheet file path inside ZIP) ────────

interface SheetEntry {
  name: string
  path: string // e.g. "xl/worksheets/sheet1.xml"
}

async function getWorkbookSheets(zip: JSZip): Promise<SheetEntry[]> {
  const wbFile = zip.file('xl/workbook.xml')
  const relsFile = zip.file('xl/_rels/workbook.xml.rels')
  if (!wbFile || !relsFile) return []

  const wbXml = await wbFile.async('text')
  const relsXml = await relsFile.async('text')

  const sheets: SheetEntry[] = []
  const sheetRe = /<sheet\s[^>]*>/g
  let m: RegExpExecArray | null

  while ((m = sheetRe.exec(wbXml)) !== null) {
    const tag = m[0]
    const nameM = tag.match(/\bname="([^"]*)"/)
    const ridM = tag.match(/\br:id="([^"]*)"/)
    if (!nameM || !ridM) continue

    const relRe = new RegExp(`Id="${ridM[1]}"[^>]*Target="([^"]*)"`)
    const relM = relRe.exec(relsXml)
    if (!relM) continue

    const rawTarget = relM[1]
    const path = rawTarget.startsWith('/xl/')
      ? rawTarget.slice(1)
      : rawTarget.startsWith('worksheets/')
      ? `xl/${rawTarget}`
      : rawTarget

    sheets.push({ name: nameM[1], path })
  }
  return sheets
}

// ── Cell parsing ─────────────────────────────────────────────────────────────

interface ParsedCell {
  ref: string
  col: number
  row: number
  styleAttr: string   // raw s="N" attribute string (preserved verbatim)
  type: string        // t attribute value ('s', 'n', 'str', 'inlineStr', 'b', '')
  hasFormula: boolean
  formulaAttrs: string
  formulaText: string
  rawValue: string    // text inside <v>
  textValue: string   // resolved string value
}

function parseWorksheetCells(wsXml: string, sharedStrings: string[]): ParsedCell[] {
  const cells: ParsedCell[] = []

  // Track refs already added to avoid double-counting
  const seen = new Set<string>()

  function pushCell(attrs: string, body: string | null) {
    const refM = attrs.match(/\br="([A-Z]+\d+)"/)
    if (!refM) return
    const ref = refM[1]
    if (seen.has(ref)) return
    seen.add(ref)
    const pos = parseCellRef(ref)
    if (!pos) return

    const sM = attrs.match(/\bs="\d+"/)
    const tM = attrs.match(/\bt="([^"]*)"/)
    const vM = body ? body.match(/<v>([^<]*)<\/v>/) : null
    const formulaWithBody = body ? body.match(/<f([^>]*)>([^<]*)<\/f>/) : null
    const formulaSelfClosing = body && !formulaWithBody ? body.match(/<f([^>]*)\/>/) : null
    const hasFormula = Boolean(formulaWithBody || formulaSelfClosing)
    const formulaAttrs = formulaWithBody?.[1] ?? formulaSelfClosing?.[1] ?? ''
    const formulaText = formulaWithBody?.[2] ?? ''

    const type = tM ? tM[1] : ''
    const rawValue = vM ? vM[1] : ''
    let textValue = rawValue

    if (type === 's') {
      const idx = parseInt(rawValue)
      if (!isNaN(idx)) textValue = sharedStrings[idx] ?? ''
    } else if (type === 'inlineStr') {
      const tInline = body ? body.match(/<is>\s*<t(?:\s[^>]*)?>([^<]*)<\/t>\s*<\/is>/) : null
      if (tInline) textValue = tInline[1]
    }

    cells.push({
      ref,
      col: pos.col,
      row: pos.row,
      styleAttr: sM ? sM[0] : '',
      type,
      hasFormula,
      formulaAttrs,
      formulaText,
      rawValue,
      textValue,
    })
  }

  // 1. Self-closing cells: <c r="B11" s="3"/> — no body, no formula, no value
  const scRe = /<c\s([^>]*?)\/>/g
  let m: RegExpExecArray | null
  while ((m = scRe.exec(wsXml)) !== null) pushCell(m[1], null)

  // 2. Cells with content: <c r="I10" s="5">...</c>
  //    Use negative lookbehind so we don't match the > in />
  const cRe = /<c\s([^>]*[^/])>([\s\S]*?)<\/c>/g
  while ((m = cRe.exec(wsXml)) !== null) pushCell(m[1], m[2])

  return cells
}

// ── Target-update helper ─────────────────────────────────────────────────────

/**
 * Replace the <v> content for a specific cell reference in raw worksheet XML.
 * For formula cells this updates the cached result while leaving <f> intact.
 * Uses a tight regex that matches the exact r="REF" attribute.
 */
function setCellStringValue(xml: string, ref: string, value: string): string {
  const escapedValue = escapeXml(value)
  const inlineStringBody = `<is><t xml:space="preserve">${escapedValue}</t></is>`

  const applyInlineStringType = (openTag: string): string => {
    if (/\bt="[^"]*"/.test(openTag)) {
      return openTag.replace(/\bt="[^"]*"/, 't="inlineStr"')
    }
    return openTag.replace(/>$/, ' t="inlineStr">')
  }

  const selfClosing = new RegExp(`(<c(?=\\s[^>]*\\br="${ref}")[^>]*?)\\s*/>`)
  const expanded = xml.replace(selfClosing, (_, openTag: string) => {
    const normalizedTag = applyInlineStringType(`${openTag}>`)
    return `${normalizedTag}${inlineStringBody}</c>`
  })
  if (expanded !== xml) return expanded

  const cellBlock = new RegExp(
    `(<c(?=\\s[^>]*\\br="${ref}")(?![^>]*\\/>)\\s[^>]*>)([\\s\\S]*?)(</c>)`,
  )

  return xml.replace(cellBlock, (_, openTag: string, _body: string, closeTag: string) => {
    return `${applyInlineStringType(openTag)}${inlineStringBody}${closeTag}`
  })
}

function stripCellTypeAttr(openTag: string): string {
  return openTag.replace(/\s+t="[^"]*"/, '')
}

function setCellBlankValue(xml: string, ref: string): string {
  const selfClosing = new RegExp(`(<c(?=\\s[^>]*\\br="${ref}")[^>]*?)\\s*/>`)
  const normalizedSelfClosing = xml.replace(selfClosing, (_, openTag: string) => {
    return `${stripCellTypeAttr(openTag)}/>`
  })
  if (normalizedSelfClosing !== xml) return normalizedSelfClosing

  const cellBlock = new RegExp(
    `(<c(?=\\s[^>]*\\br="${ref}")(?![^>]*\\/>)\\s[^>]*>)([\\s\\S]*?)(</c>)`,
  )

  return xml.replace(cellBlock, (_, openTag: string, body: string) => {
    const normalizedOpenTag = stripCellTypeAttr(openTag)
    const formulaTag = body.match(/<f(?:\s[^>]*)?>[^<]*<\/f>|<f(?:\s[^>]*)?\/>/)?.[0]

    if (formulaTag) {
      return `${normalizedOpenTag}${formulaTag}</c>`
    }

    return `${normalizedOpenTag.slice(0, -1)}/>`
  })
}

function setCellNumberValue(xml: string, ref: string, value: number): string {
  const escapedValue = typeof value === 'string' ? escapeXml(value) : String(value)

  // Case 1: self-closing cell <c r="REF" s="N"/> → expand and inject value
  const selfClosing = new RegExp(`(<c(?=\\s[^>]*\\br="${ref}")[^>]*?)\\s*/>`)
  const expanded = xml.replace(selfClosing, `$1><v>${escapedValue}</v></c>`)
  if (expanded !== xml) return expanded

  // Case 2: existing cell block (formula or plain numeric)
  const cellBlock = new RegExp(
    `(<c(?=\\s[^>]*\\br="${ref}")(?![^>]*\\/>)\\s[^>]*>)([\\s\\S]*?)(</c>)`,
  )

  return xml.replace(cellBlock, (_, openTag: string, body: string, closeTag: string) => {
    if (/<v>[^<]*<\/v>/.test(body)) {
      return `${openTag}${body.replace(/<v>[^<]*<\/v>/, `<v>${escapedValue}</v>`)}${closeTag}`
    }

    return `${openTag}${body}<v>${escapedValue}</v>${closeTag}`
  })
}

function setCellValue(xml: string, ref: string, value: number | string): string {
  if (typeof value === 'string') {
    return setCellStringValue(xml, ref, value)
  }
  if (value === 0) {
    return setCellBlankValue(xml, ref)
  }
  return setCellNumberValue(xml, ref, value)
}

function setCellFormula(xml: string, ref: string, formula: string): string {
  const escapedFormula = escapeXml(formula)

  const selfClosing = new RegExp(`(<c(?=\\s[^>]*\\br="${ref}")[^>]*?)\\s*/>`)
  const expanded = xml.replace(selfClosing, `$1><f>${escapedFormula}</f></c>`)
  if (expanded !== xml) return expanded

  const cellBlock = new RegExp(
    `(<c(?=\\s[^>]*\\br="${ref}")(?![^>]*\\/>)\\s[^>]*>)([\\s\\S]*?)(</c>)`,
  )

  return xml.replace(cellBlock, (_, openTag: string, body: string, closeTag: string) => {
    if (/<f(?:\s[^>]*)?>([^<]*)<\/f>/.test(body)) {
      return `${openTag}${body.replace(/<f(?:\s[^>]*)?>[^<]*<\/f>/, `<f>${escapedFormula}</f>`)}${closeTag}`
    }

    if (/<f(?:\s[^>]*)?\/>/.test(body)) {
      return `${openTag}${body.replace(/<f(?:\s[^>]*)?\/>/, `<f>${escapedFormula}</f>`)}${closeTag}`
    }

    return `${openTag}<f>${escapedFormula}</f>${body}${closeTag}`
  })
}

interface SharedFormula {
  baseRef: string
  formulaText: string
}

function shiftFormulaRefs(formula: string, rowDelta: number, colDelta: number): string {
  return formula.replace(/\b([A-Z]+)(\d+)\b/g, (_, letters, rowNum) => {
    const col = colIndex(letters) + colDelta
    const row = parseInt(rowNum, 10) - 1 + rowDelta
    if (col < 0 || row < 0) return `${letters}${rowNum}`
    return encodeCellRef(col, row)
  })
}

function resolveFormulaText(cell: ParsedCell, sharedFormulas: Map<string, SharedFormula>): string {
  if (!cell.hasFormula) return ''
  if (cell.formulaText.trim()) return cell.formulaText.trim()

  const isShared = /\bt="shared"/.test(cell.formulaAttrs)
  const si = cell.formulaAttrs.match(/\bsi="([^"]+)"/)?.[1]
  if (!isShared || !si) return ''

  const shared = sharedFormulas.get(si)
  if (!shared) return ''

  const basePos = parseCellRef(shared.baseRef)
  const cellPos = parseCellRef(cell.ref)
  if (!basePos || !cellPos) return ''

  return shiftFormulaRefs(
    shared.formulaText,
    cellPos.row - basePos.row,
    cellPos.col - basePos.col,
  )
}

function buildSharedFormulaMap(cells: ParsedCell[]): Map<string, SharedFormula> {
  const shared = new Map<string, SharedFormula>()

  for (const cell of cells) {
    if (!cell.hasFormula || !cell.formulaText.trim()) continue
    if (!/\bt="shared"/.test(cell.formulaAttrs)) continue

    const si = cell.formulaAttrs.match(/\bsi="([^"]+)"/)?.[1]
    if (!si) continue

    shared.set(si, {
      baseRef: cell.ref,
      formulaText: cell.formulaText.trim(),
    })
  }

  return shared
}

function getCellNumericValue(cell: ParsedCell | undefined): number {
  if (!cell) return 0
  return Number.parseFloat(cell.rawValue) || 0
}

function iterateRange(startRef: string, endRef: string, visit: (ref: string) => void): void {
  const start = parseCellRef(startRef)
  const end = parseCellRef(endRef)
  if (!start || !end) return

  const rowStart = Math.min(start.row, end.row)
  const rowEnd = Math.max(start.row, end.row)
  const colStart = Math.min(start.col, end.col)
  const colEnd = Math.max(start.col, end.col)

  for (let row = rowStart; row <= rowEnd; row++) {
    for (let col = colStart; col <= colEnd; col++) {
      visit(encodeCellRef(col, row))
    }
  }
}

/** Cells of other worksheets, keyed by normalized sheet name, for cross-sheet references */
type SheetCellIndex = Map<string, Map<string, ParsedCell>>

const CROSS_SHEET_REF_RE = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]+)\$?(\d+)$/

function evaluateCrossSheetRef(formula: string, otherSheets?: SheetCellIndex): number | null {
  const match = formula.trim().match(CROSS_SHEET_REF_RE)
  if (!match || !otherSheets) return null
  const sheetName = (match[1] ?? match[2]).replace(/''/g, "'")
  const cells = otherSheets.get(normalizeName(sheetName))
  if (!cells) return null
  return getCellNumericValue(cells.get(`${match[3]}${match[4]}`))
}

function evaluateSimpleFormula(
  formula: string,
  cellsByRef: Map<string, ParsedCell>,
  otherSheets?: SheetCellIndex,
): number | null {
  const crossSheet = evaluateCrossSheetRef(formula.replace(/&apos;/g, "'"), otherSheets)
  if (crossSheet !== null) return crossSheet

  const normalized = formula.trim().replace(/\s+/g, '')

  const sumMatch = normalized.match(/^SUM\((.+)\)$/i)
  if (sumMatch) {
    let total = 0
    const parts = sumMatch[1].split(',')

    for (const part of parts) {
      const rangeMatch = part.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)
      if (rangeMatch) {
        iterateRange(rangeMatch[1], rangeMatch[2], (ref) => {
          total += getCellNumericValue(cellsByRef.get(ref))
        })
        continue
      }

      const cellMatch = part.match(/^([A-Z]+\d+)$/)
      if (cellMatch) {
        total += getCellNumericValue(cellsByRef.get(cellMatch[1]))
        continue
      }

      return null
    }

    return total
  }

  const arithmeticMatch = normalized.match(/^([A-Z]+\d+)([+-])([A-Z]+\d+)$/)
  if (arithmeticMatch) {
    const left = getCellNumericValue(cellsByRef.get(arithmeticMatch[1]))
    const right = getCellNumericValue(cellsByRef.get(arithmeticMatch[3]))
    return arithmeticMatch[2] === '+' ? left + right : left - right
  }

  return null
}

function recalculateSimpleFormulas(
  wsXml: string,
  sharedStrings: string[],
  otherSheets?: SheetCellIndex,
): string {
  let updatedXml = wsXml

  // A few passes are enough for the template's dependency chains
  for (let pass = 0; pass < 4; pass++) {
    const cells = parseWorksheetCells(updatedXml, sharedStrings)
    const sharedFormulas = buildSharedFormulaMap(cells)
    const cellsByRef = new Map(cells.map((cell) => [cell.ref, cell]))

    let passXml = updatedXml

    for (const cell of cells) {
      if (!cell.hasFormula) continue

      const formula = resolveFormulaText(cell, sharedFormulas)
      if (!formula) continue

      const value = evaluateSimpleFormula(formula, cellsByRef, otherSheets)
      if (value === null) continue

      passXml = setCellValue(passXml, cell.ref, value)
    }

    if (passXml === updatedXml) break
    updatedXml = passXml
  }

  return updatedXml
}

function normalizeSheetFormulas(wsXml: string, sharedStrings: string[]): string {
  const cells = parseWorksheetCells(wsXml, sharedStrings)
  const sharedFormulas = buildSharedFormulaMap(cells)
  let updatedXml = wsXml

  for (const cell of cells) {
    if (!cell.hasFormula) continue

    const formula = resolveFormulaText(cell, sharedFormulas)
    if (!formula) continue

    updatedXml = setCellFormula(updatedXml, cell.ref, formula)
  }

  return updatedXml
}

function processInformeEconomicoSheet(
  wsXml: string,
  sharedStrings: string[],
  consolidated: InformeEconomicoSheet,
  signatureRow: number | null,
): string {
  for (const [ref, value] of Object.entries(consolidated.sums)) {
    wsXml = setCellValue(wsXml, ref, value)
  }

  if (consolidated.concept) wsXml = setCellValue(wsXml, 'I13', consolidated.concept)
  if (consolidated.supplier) wsXml = setCellValue(wsXml, 'J13', consolidated.supplier)

  if (signatureRow) {
    if (consolidated.signer) wsXml = setCellValue(wsXml, `D${signatureRow}`, consolidated.signer)
    if (consolidated.sentDate !== null) wsXml = setCellValue(wsXml, `J${signatureRow}`, consolidated.sentDate)
  }

  return recalculateSimpleFormulas(wsXml, sharedStrings)
}

// ── Sheets 1–4: update numeric cell values by row-label match ────────────────

/**
 * Write the canonical values of one label-driven sheet into the template.
 * Only rows inside the anchored data block are touched, each value goes to the
 * same column letter it was read from, and formula cells are never overwritten.
 * Capture cells without data are written as blank so no template value survives.
 */
function processLabeledSheet(
  wsXml: string,
  sharedStrings: string[],
  consolidated: LabeledSheet,
  key: LabeledSheetKey,
  block: DataBlock | null,
  otherSheets?: SheetCellIndex,
): string {
  if (!block) return wsXml
  const cellsByRef = new Map(parseWorksheetCells(wsXml, sharedStrings).map((cell) => [cell.ref, cell]))

  for (let row = block.firstRow; row <= block.lastRow; row++) {
    const label = cellsByRef.get(`A${row}`)?.textValue.trim()
    if (!label) continue

    const resolved = resolveRowId(key, label)
    const values = 'id' in resolved ? consolidated.values.get(resolved.id) : undefined

    for (const col of INPUT_COLS) {
      const target = cellsByRef.get(`${col}${row}`)
      if (!target || target.hasFormula) continue
      wsXml = setCellValue(wsXml, target.ref, values?.[col] ?? 0)
    }
  }

  return recalculateSimpleFormulas(wsXml, sharedStrings, otherSheets)
}

// ── Sheet 5 (Monitoreo): doctor columns, formula roles and column expansion ──

/** Shift same-sheet A1 references whose column is ≥ fromCol (like Excel's column insert). */
function shiftFormulaColumns(formula: string, fromCol: number, count: number): string {
  return formula.replace(
    /(^|[^A-Za-z0-9_!'])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g,
    (match, pre: string, colAbs: string, letters: string, rowAbs: string, rowNum: string) => {
      const col = colIndex(letters)
      if (col < fromCol) return match
      return `${pre}${colAbs}${colLetters(col + count)}${rowAbs}${rowNum}`
    },
  )
}

/**
 * Shift every cell reference at column ≥ fromCol rightward by `count` columns:
 * r="..." attributes, the row spans="start:end" attribute and formula texts.
 */
function shiftCellsRight(xml: string, fromCol: number, count: number): string {
  // Shift individual cell r attributes
  xml = xml.replace(/\br="([A-Z]+)(\d+)"/g, (match, letters, rowNum) => {
    const col = colIndex(letters)
    if (col >= fromCol) {
      return `r="${encodeCellRef(col + count, parseInt(rowNum) - 1)}"`
    }
    return match
  })

  // Update spans="start:end" on every <row> element
  xml = xml.replace(/(<row\s[^>]*\bspans=")(\d+):(\d+)(")/g, (_, pre, start, end, post) => {
    const s = parseInt(start)
    const e = parseInt(end)
    const insertCol = fromCol + 1
    const newEnd = e >= insertCol ? e + count : e
    const newStart = s >= insertCol ? s + count : s
    return `${pre}${newStart}:${newEnd}${post}`
  })

  // Formulas keep pointing at the same (moved) cells
  return xml.replace(/<f\b([^>]*)>([^<]*)<\/f>/g, (_, attrs: string, text: string) => {
    return `<f${attrs}>${shiftFormulaColumns(text, fromCol, count)}</f>`
  })
}

function setColRange(colTag: string, min: number, max: number): string {
  return colTag
    .replace(/\bmin="\d+"/, `min="${min}"`)
    .replace(/\bmax="\d+"/, `max="${max}"`)
}

function shiftWorksheetColumnsRight(xml: string, fromCol: number, count: number): string {
  const maxExcelColumn = 16384
  const insertCol = fromCol + 1
  const colsMatch = xml.match(/<cols>[\s\S]*?<\/cols>/)
  if (!colsMatch) return xml

  const colsXml = colsMatch[0]
  const colTags = [...colsXml.matchAll(/<col\b[^>]*\/>/g)].map((match) => match[0])
  if (colTags.length === 0) return xml

  const parseRange = (tag: string): { min: number; max: number } | null => {
    const minMatch = tag.match(/\bmin="(\d+)"/)
    const maxMatch = tag.match(/\bmax="(\d+)"/)
    if (!minMatch || !maxMatch) return null
    return { min: parseInt(minMatch[1], 10), max: parseInt(maxMatch[1], 10) }
  }

  const sourceCol = colTags.find((tag) => {
    const range = parseRange(tag)
    return range ? range.min <= insertCol - 1 && range.max >= insertCol - 1 : false
  }) ?? colTags[0]

  const updatedTags: string[] = []
  let inserted = false

  for (const tag of colTags) {
    const range = parseRange(tag)
    if (!range) {
      updatedTags.push(tag)
      continue
    }

    if (!inserted && range.min >= insertCol) {
      updatedTags.push(setColRange(sourceCol, insertCol, insertCol))
      inserted = true
    }

    if (range.min >= insertCol) {
      const shiftedMin = range.min + count
      const shiftedMax = Math.min(range.max + count, maxExcelColumn)
      if (shiftedMin <= maxExcelColumn && shiftedMin <= shiftedMax) {
        updatedTags.push(setColRange(tag, shiftedMin, shiftedMax))
      }
    } else {
      updatedTags.push(tag)
    }
  }

  if (!inserted) {
    updatedTags.push(setColRange(sourceCol, insertCol, insertCol))
  }

  return xml.replace(colsXml, `<cols>${updatedTags.join('')}</cols>`)
}

function shiftWorksheetMergeRefsRight(xml: string, fromCol: number, count: number): string {
  return xml.replace(/<mergeCell ref="([A-Z]+\d+):([A-Z]+\d+)"\/>/g, (match, startRef, endRef) => {
    const start = parseCellRef(startRef)
    const end = parseCellRef(endRef)
    if (!start || !end) return match

    const nextStartCol = start.col >= fromCol ? start.col + count : start.col
    const nextEndCol = end.col >= fromCol ? end.col + count : end.col

    return `<mergeCell ref="${encodeCellRef(nextStartCol, start.row)}:${encodeCellRef(nextEndCol, end.row)}"/>`
  })
}

/**
 * Shift one range ("B23:AF117" or "AC119"). Ranges that end on the last doctor
 * column grow with the inserted column so new doctors share its validation.
 */
function shiftRangeRef(range: string, fromCol: number, count: number, lastDoctorCol: number): string {
  const [startRef, endRef = startRef] = range.split(':')
  const start = parseCellRef(startRef)
  const end = parseCellRef(endRef)
  if (!start || !end) return range

  const startCol = start.col >= fromCol ? start.col + count : start.col
  const endCol = end.col >= fromCol || end.col === lastDoctorCol ? end.col + count : end.col
  const nextStart = encodeCellRef(startCol, start.row)
  const nextEnd = encodeCellRef(endCol, end.row)
  return range.includes(':') || nextStart !== nextEnd ? `${nextStart}:${nextEnd}` : nextStart
}

/** Data validations, conditional formats (incl. x14 extensions) and <dimension>. */
function shiftSqrefsRight(xml: string, fromCol: number, count: number, lastDoctorCol: number): string {
  const shiftList = (list: string) => list
    .split(/\s+/)
    .filter(Boolean)
    .map((range) => shiftRangeRef(range, fromCol, count, lastDoctorCol))
    .join(' ')

  return xml
    .replace(
      /(<(?:dataValidation|conditionalFormatting)\b[^>]*\bsqref=")([^"]+)(")/g,
      (_, open: string, list: string, close: string) => `${open}${shiftList(list)}${close}`,
    )
    .replace(
      /(<xm:sqref>)([^<]+)(<\/xm:sqref>)/g,
      (_, open: string, list: string, close: string) => `${open}${shiftList(list)}${close}`,
    )
    .replace(
      /(<dimension ref=")([^"]+)(")/,
      (_, open: string, ref: string, close: string) => `${open}${shiftRangeRef(ref, fromCol, count, -1)}${close}`,
    )
}

/** Insert a prebuilt <c> element into a row block, keeping column order. */
function insertCellXml(rowXml: string, newRef: string, cellXml: string): string {
  const pos = parseCellRef(newRef)
  if (!pos) return rowXml

  // Find the first cell with a column index >= newCol and insert before it.
  let inserted = false
  const insertBefore = rowXml.replace(/<c\s[^>]*>/g, (tag) => {
    if (inserted) return tag

    const refM = tag.match(/\br="([A-Z]+\d+)"/)
    if (!refM) return tag
    const p = parseCellRef(refM[1])
    if (p && p.col >= pos.col) {
      inserted = true
      return `${cellXml}${tag}`
    }
    return tag
  })

  // If nothing was replaced (new column is the last one), append before </row>
  if (!inserted) {
    return rowXml.replace('</row>', `${cellXml}</row>`)
  }
  return insertBefore
}

/** Add a cell to an existing (non-empty) row of the sheet. */
function insertCellIntoSheet(xml: string, ref: string, cellXml: string): string {
  const pos = parseCellRef(ref)
  if (!pos) return xml
  const rowRe = new RegExp(`<row\\b[^>]*?\\br="${pos.row + 1}"[^>]*>[\\s\\S]*?</row>`)
  return xml.replace(rowRe, (rowXml) => insertCellXml(rowXml, ref, cellXml))
}

/**
 * Insert one doctor column at `insertAt` (the current total column). Every row
 * of the doctor table gets a cell styled like the previous doctor column; rows
 * where that column holds a per-doctor formula get the same formula shifted.
 */
function insertDoctorColumn(
  xml: string,
  sharedStrings: string[],
  insertAt: number,
  block: DataBlock & { headerRow: number },
): string {
  const lastDoctorCol = insertAt - 1

  xml = shiftCellsRight(xml, insertAt, 1)
  xml = shiftWorksheetColumnsRight(xml, insertAt, 1)
  xml = shiftWorksheetMergeRefsRight(xml, insertAt, 1)
  xml = shiftSqrefsRight(xml, insertAt, 1, lastDoctorCol)

  const cellsByRef = new Map(parseWorksheetCells(xml, sharedStrings).map((cell) => [cell.ref, cell]))

  return xml.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (rowXml) => {
    const rowNum = Number(rowXml.match(/^<row\b[^>]*?\br="(\d+)"/)?.[1])
    if (!rowNum || rowNum < block.headerRow || rowNum > block.lastRow) return rowXml

    const previous = cellsByRef.get(encodeCellRef(lastDoctorCol, rowNum - 1))
    if (!previous) return rowXml

    const newRef = encodeCellRef(insertAt, rowNum - 1)
    const sAttr = previous.styleAttr ? ` ${previous.styleAttr}` : ''
    const formula = previous.hasFormula ? previous.formulaText.trim() : ''
    const cellXml = formula
      ? `<c r="${newRef}"${sAttr}><f>${shiftFormulaRefs(formula, 0, 1)}</f></c>`
      : `<c r="${newRef}"${sAttr}/>`
    return insertCellXml(rowXml, newRef, cellXml)
  })
}

const HORIZONTAL_SUM_RE = /^SUM\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)$/

/**
 * Write doctors into the Monitoreo template.
 *
 * Formula roles are kept: per-doctor formulas (TOTAL ASEGURADOS, TOTAL DE
 * PROCEDIMIENTOS...) are never overwritten, vertical subtotals of the total
 * column (AF28=SUM(AF29:AF31)) are preserved, and only same-row totals
 * (SUM(B12:AE12)) are regenerated over the final doctor range.
 */
function processMonitoreoSheet(
  wsXml: string,
  sharedStrings: string[],
  consolidated: MonitoreoSheet,
  block: DataBlock | null,
): string {
  if (!block?.headerRow) return wsXml
  const table = { ...block, headerRow: block.headerRow }
  const headerRowIdx = table.headerRow - 1

  // Explicit formulas first: shared-formula groups cannot survive column moves
  wsXml = normalizeSheetFormulas(wsXml, sharedStrings)

  let cellsByRef = new Map(parseWorksheetCells(wsXml, sharedStrings).map((cell) => [cell.ref, cell]))
  const maxCol = Math.max(...[...cellsByRef.values()].map((cell) => cell.col))

  let totalCol = 1
  while (totalCol <= maxCol && !cellsByRef.get(encodeCellRef(totalCol, table.firstRow - 1))?.hasFormula) totalCol++

  const doctorCols: number[] = []
  for (let col = 1; col < totalCol; col++) {
    if (cellsByRef.has(encodeCellRef(col, headerRowIdx))) doctorCols.push(col)
  }

  const orderedDoctors = [...consolidated.doctors.values()]

  while (orderedDoctors.length > doctorCols.length) {
    wsXml = insertDoctorColumn(wsXml, sharedStrings, totalCol, table)
    doctorCols.push(totalCol)
    totalCol++
  }

  cellsByRef = new Map(parseWorksheetCells(wsXml, sharedStrings).map((cell) => [cell.ref, cell]))
  const rowIds = new Map<number, string>()
  for (let row = table.firstRow; row <= table.lastRow; row++) {
    const label = cellsByRef.get(`A${row}`)?.textValue.trim()
    if (!label) continue
    const resolved = resolveRowId('monitoreo', label)
    if ('id' in resolved) rowIds.set(row, resolved.id)
  }

  doctorCols.forEach((col, index) => {
    const entry = orderedDoctors[index]
    wsXml = setCellValue(wsXml, encodeCellRef(col, headerRowIdx), entry?.displayName ?? '')

    for (const [row, rowId] of rowIds) {
      const ref = encodeCellRef(col, row - 1)
      const target = cellsByRef.get(ref)
      if (target?.hasFormula) continue

      const value = entry?.values.get(rowId) ?? 0
      if (!target) {
        if (value !== 0) wsXml = insertCellIntoSheet(wsXml, ref, `<c r="${ref}"><v>${value}</v></c>`)
        continue
      }
      // Sparse writes; a template value is cleared even when the new value is 0
      if (value !== 0 || target.rawValue !== '') wsXml = setCellValue(wsXml, ref, value)
    }
  })

  wsXml = setCellValue(wsXml, encodeCellRef(totalCol, headerRowIdx), 'TOTAL')

  const firstDoctorCol = doctorCols[0]
  const lastDoctorCol = doctorCols[doctorCols.length - 1]
  if (firstDoctorCol !== undefined && lastDoctorCol !== undefined) {
    cellsByRef = new Map(parseWorksheetCells(wsXml, sharedStrings).map((cell) => [cell.ref, cell]))
    for (let row = table.firstRow; row <= table.lastRow; row++) {
      const totalCell = cellsByRef.get(encodeCellRef(totalCol, row - 1))
      if (!totalCell?.hasFormula) continue

      const match = totalCell.formulaText.trim().match(HORIZONTAL_SUM_RE)
      if (!match || Number(match[2]) !== row || Number(match[4]) !== row) continue

      wsXml = setCellFormula(
        wsXml,
        totalCell.ref,
        `SUM(${encodeCellRef(firstDoctorCol, row - 1)}:${encodeCellRef(lastDoctorCol, row - 1)})`,
      )
    }
  }

  return recalculateSimpleFormulas(wsXml, sharedStrings)
}

async function removeCalcChain(zip: JSZip): Promise<void> {
  zip.remove('xl/calcChain.xml')

  const contentTypes = zip.file('[Content_Types].xml')
  if (contentTypes) {
    const xml = await contentTypes.async('text')
    zip.file(
      '[Content_Types].xml',
      xml.replace(
        /<Override PartName="\/xl\/calcChain\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.calcChain\+xml"\/>/,
        '',
      ),
    )
  }

  const workbookRels = zip.file('xl/_rels/workbook.xml.rels')
  if (workbookRels) {
    const xml = await workbookRels.async('text')
    zip.file(
      'xl/_rels/workbook.xml.rels',
      xml.replace(
        /<Relationship Id="[^"]+" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain" Target="calcChain\.xml"\/>/,
        '',
      ),
    )
  }
}

// ── Main export ──────────────────────────────────────────────────────────────

export async function writeConsolidatedWorkbook(output: ConsolidatedOutput): Promise<void> {
  const bytes = await buildConsolidatedWorkbook(output)
  const outBlob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  saveAs(outBlob, `OdoLens_Consolidado_${todayISO()}.xlsx`)
}

/**
 * Zero results must read as blank cells (user decision 5). With fullCalcOnLoad
 * Excel recomputes every formula, so blank caches are not enough: the sheet
 * views hide zero values instead.
 */
function hideZeroValues(wsXml: string): string {
  return wsXml.replace(/<sheetView\b[^>]*?(?=\/?>)/g, (tag) => (
    /\bshowZeros="/.test(tag)
      ? tag.replace(/\bshowZeros="[^"]*"/, 'showZeros="0"')
      : tag.replace(/^<sheetView\b/, '<sheetView showZeros="0"')
  ))
}

/** Hidden defined name that lets OdoLens recognize its own output if it is fed back. */
async function stampOutputMarker(zip: JSZip): Promise<void> {
  const file = zip.file('xl/workbook.xml')
  if (!file) return
  const xml = await file.async('text')
  if (xml.includes(`name="${ODOLENS_OUTPUT_MARKER}"`)) return

  const definedName = `<definedName name="${ODOLENS_OUTPUT_MARKER}" hidden="1">"OdoLens"</definedName>`
  const updated = xml.includes('</definedNames>')
    ? xml.replace('</definedNames>', `${definedName}</definedNames>`)
    : xml.replace('</sheets>', `</sheets><definedNames>${definedName}</definedNames>`)
  zip.file('xl/workbook.xml', updated)
}

/** The template's author metadata must not travel into every consolidated file. */
async function scrubDocumentProperties(zip: JSZip): Promise<void> {
  const file = zip.file('docProps/core.xml')
  if (!file) return
  const xml = await file.async('text')
  zip.file(
    'docProps/core.xml',
    xml
      .replace(/<dc:creator>[^<]*<\/dc:creator>/, '<dc:creator>OdoLens</dc:creator>')
      .replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, '<cp:lastModifiedBy>OdoLens</cp:lastModifiedBy>'),
  )
}

/** Ask Excel to recalculate every formula when the file is opened. */
async function ensureFullCalcOnLoad(zip: JSZip): Promise<void> {
  const file = zip.file('xl/workbook.xml')
  if (!file) return
  const xml = await file.async('text')

  let updated: string
  if (/<calcPr\b[^>]*\bfullCalcOnLoad="/.test(xml)) {
    updated = xml.replace(/(<calcPr\b[^>]*\bfullCalcOnLoad=")[^"]*"/, '$11"')
  } else if (/<calcPr\b/.test(xml)) {
    updated = xml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"')
  } else {
    const anchor = xml.includes('</definedNames>') ? '</definedNames>' : '</sheets>'
    updated = xml.replace(anchor, `${anchor}<calcPr fullCalcOnLoad="1"/>`)
  }
  zip.file('xl/workbook.xml', updated)
}

export async function buildConsolidatedWorkbook(output: ConsolidatedOutput): Promise<Uint8Array> {
  // Load the embedded template as a ZIP — keeps every internal file intact
  const templateBytes = Uint8Array.from(atob(TEMPLATE_V2026_BASE64), (c) => c.charCodeAt(0))
  const zip = await JSZip.loadAsync(templateBytes)
  // Same anchors as the reader: blocks and signature row come from the template itself
  const templateRaw = XLSX.read(templateBytes, { type: 'array' })

  // Parse shared strings (needed to resolve text cell values)
  const ssFile = zip.file('xl/sharedStrings.xml')
  const sharedStrings = ssFile ? parseSharedStrings(await ssFile.async('text')) : []

  // Get the sheet name → file path mapping
  const sheetEntries = await getWorkbookSheets(zip)
  // Already-written sheets, for cross-sheet formulas (Consolidado → Aseguramiento)
  const writtenSheets: SheetCellIndex = new Map()

  const labeled = (key: LabeledSheetKey) => (xml: string) =>
    processLabeledSheet(xml, sharedStrings, output[key], key, findBlock(templateRaw, key), writtenSheets)

  // Order matters: Aseguramiento is written before Consolidado reads it
  const sheetProcessors: Array<[string, (xml: string) => string]> = [
    ['Nacionalidad', labeled('nacionalidad')],
    ['Aseguramiento', labeled('aseguramiento')],
    ['Consolidado', labeled('consolidado')],
    [ECONOMICO_SHEET, (xml) => processInformeEconomicoSheet(
      xml, sharedStrings, output.informeEconomico, findEconomicoSignatureRow(templateRaw),
    )],
    ['Monitoreo', (xml) => processMonitoreoSheet(xml, sharedStrings, output.monitoreo, findBlock(templateRaw, 'monitoreo'))],
  ]

  for (const [sheetName, processor] of sheetProcessors) {
    // Match by normalized name (handles accent/case/space differences)
    const entry = sheetEntries.find(
      (e) => normalizeName(e.name) === normalizeName(sheetName),
    )
    if (!entry) continue

    const wsFile = zip.file(entry.path)
    if (!wsFile) continue

    const originalXml = await wsFile.async('text')
    const updatedXml = hideZeroValues(processor(originalXml))
    zip.file(entry.path, updatedXml)
    writtenSheets.set(
      normalizeName(entry.name),
      new Map(parseWorksheetCells(updatedXml, sharedStrings).map((cell) => [cell.ref, cell])),
    )
  }

  await removeCalcChain(zip)
  await stampOutputMarker(zip)
  await ensureFullCalcOnLoad(zip)
  await scrubDocumentProperties(zip)

  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}
