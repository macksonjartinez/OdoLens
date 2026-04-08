/**
 * excelWriter.ts — JSZip-based binary-faithful output
 *
 * Strategy: read the template file as a raw ZIP, update ONLY the <v> element
 * content inside worksheet XMLs, keep every other file (styles.xml, drawings/,
 * media/, calcChain.xml, printerSettings, etc.) completely untouched, then
 * re-pack and download.  This preserves 100% of the original formatting.
 */

import JSZip from 'jszip'
import { saveAs } from 'file-saver'
import type {
  ConsolidatedOutput,
  InformeEconomicoSheet,
  LabeledSheet,
  MonitoreoSheet,
} from '../types'
import { normalizeName, todayISO } from './normalize'
import { TEMPLATE_BASE64 } from '../assets/template'

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

function evaluateSimpleFormula(formula: string, cellsByRef: Map<string, ParsedCell>): number | null {
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

function recalculateSimpleFormulas(wsXml: string, sharedStrings: string[]): string {
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

      const value = evaluateSimpleFormula(formula, cellsByRef)
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
): string {
  for (const [ref, value] of consolidated.sumCells) {
    wsXml = setCellValue(wsXml, ref, value)
  }

  for (const [ref, value] of consolidated.copiedCells) {
    if (value === null || value === '') continue
    wsXml = setCellValue(wsXml, ref, typeof value === 'number' ? value : String(value))
  }

  return recalculateSimpleFormulas(wsXml, sharedStrings)
}

const MONITOREO_HEADER_LABEL = normalizeName('Nombre del Dr')

function getSortedRowCells(byRow: Map<number, ParsedCell[]>, rowIdx: number): ParsedCell[] {
  return (byRow.get(rowIdx) ?? []).slice().sort((a, b) => a.col - b.col)
}

function getFirstCellInRow(byRow: Map<number, ParsedCell[]>, rowIdx: number): ParsedCell | undefined {
  return getSortedRowCells(byRow, rowIdx)[0]
}

function findMonitoreoHeaderRowIdx(
  byRow: Map<number, ParsedCell[]>,
  sortedRows: number[],
): number {
  return sortedRows.find((rowIdx) => {
    const firstCell = getFirstCellInRow(byRow, rowIdx)
    return normalizeName(firstCell?.textValue ?? '') === MONITOREO_HEADER_LABEL
  }) ?? -1
}

function findMonitoreoTotalCol(
  byRow: Map<number, ParsedCell[]>,
  sortedRows: number[],
  headerRowIdx: number,
  labelCol: number,
): number {
  for (const rowIdx of sortedRows) {
    if (rowIdx <= headerRowIdx) continue

    const rowCells = getSortedRowCells(byRow, rowIdx)
    const labelCell = rowCells.find((cell) => cell.col === labelCol)
    if (!labelCell?.textValue.trim()) continue

    const formulaCell = rowCells.find((cell) => cell.hasFormula)
    if (formulaCell) return formulaCell.col
  }

  const headerCells = getSortedRowCells(byRow, headerRowIdx)
  return Math.max(...headerCells.map((cell) => cell.col)) + 1
}

function rebuildRowsMap(cells: ParsedCell[]): Map<number, ParsedCell[]> {
  const byRow = new Map<number, ParsedCell[]>()
  for (const cell of cells) {
    if (!byRow.has(cell.row)) byRow.set(cell.row, [])
    byRow.get(cell.row)!.push(cell)
  }
  return byRow
}

// ── Sheets 1–4: update numeric cell values by row-label match ────────────────

function processLabeledSheet(
  wsXml: string,
  sharedStrings: string[],
  consolidated: LabeledSheet,
): string {
  const cells = parseWorksheetCells(wsXml, sharedStrings)

  // Build normalized-label → values lookup
  const normData = new Map<string, number[]>()
  for (const [label, values] of consolidated.data) {
    normData.set(normalizeName(label), values)
  }

  // Group cells by row, sort columns within each row
  const byRow = new Map<number, ParsedCell[]>()
  for (const cell of cells) {
    if (!byRow.has(cell.row)) byRow.set(cell.row, [])
    byRow.get(cell.row)!.push(cell)
  }

  const sortedRows = [...byRow.keys()].sort((a, b) => a - b)
  const headerRow = sortedRows[0] // skip the first row (headers)

  for (const rowIdx of sortedRows) {
    if (rowIdx === headerRow) continue

    const rowCells = byRow.get(rowIdx)!.sort((a, b) => a.col - b.col)
    const labelCell = rowCells[0]
    if (!labelCell?.textValue.trim()) continue

    const values = normData.get(normalizeName(labelCell.textValue.trim()))
    if (!values) continue

    // Preserve the original write behavior: fill the non-formula numeric inputs
    // in order, then recalculate the visible total formulas afterwards.
    const dataCells = rowCells.filter(
      (c) => c.col > labelCell.col && !c.hasFormula && (c.type === '' || c.type === 'n'),
    )

    for (let i = 0; i < dataCells.length; i++) {
      if (i < values.length) {
        wsXml = setCellValue(wsXml, dataCells[i].ref, values[i])
      }
    }
  }

  return recalculateSimpleFormulas(wsXml, sharedStrings)
}

// ── Sheet 5 (Monitoreo): update existing doctors + insert new ones ───────────

/**
 * Shift every cell reference at column ≥ fromCol rightward by `count` columns.
 * Updates r="..." attributes and the row's spans="start:end" attribute.
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

  return xml
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
 * Insert a new <c> element at the given position (col/row) into a row block.
 * `styleAttr` is the verbatim s="N" attribute to apply (copy from neighbour).
 */
function insertCellInRow(
  rowXml: string,
  newRef: string,
  value: number | string,
  styleAttr: string,
): string {
  const isNum = typeof value === 'number'
  const sAttr = styleAttr ? ` ${styleAttr}` : ''
  const newCell = isNum
    ? value === 0
      ? `<c r="${newRef}"${sAttr}/>`
      : `<c r="${newRef}"${sAttr}><v>${value}</v></c>`
    : `<c r="${newRef}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(String(value))}</t></is></c>`

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
      return `${newCell}${tag}`
    }
    return tag
  })

  // If nothing was replaced (new column is the last one), append before </row>
  if (!inserted) {
    return rowXml.replace('</row>', `${newCell}</row>`)
  }
  return insertBefore
}

function processMonitoreoSheet(
  wsXml: string,
  sharedStrings: string[],
  consolidated: MonitoreoSheet,
): string {
  let cells = parseWorksheetCells(wsXml, sharedStrings)
  let byRow = rebuildRowsMap(cells)
  let sortedRows = [...byRow.keys()].sort((a, b) => a - b)

  const headerRowIdx = findMonitoreoHeaderRowIdx(byRow, sortedRows)
  if (headerRowIdx < 0) return wsXml

  const headerCells = getSortedRowCells(byRow, headerRowIdx)
  const labelCol = headerCells[0]?.col ?? 0
  let totalCol = findMonitoreoTotalCol(byRow, sortedRows, headerRowIdx, labelCol)

  let doctorCols = headerCells
    .filter((cell) => cell.col > labelCol && cell.col < totalCol)
    .map((cell) => ({ col: cell.col, styleAttr: cell.styleAttr }))

  const orderedDoctors = [...consolidated.doctors.values()]

  while (orderedDoctors.length > doctorCols.length) {
    const insertAt = totalCol
    const headerStyleAttr = doctorCols.length > 0
      ? doctorCols[doctorCols.length - 1].styleAttr
      : ''

    wsXml = shiftCellsRight(wsXml, insertAt, 1)
    wsXml = shiftWorksheetColumnsRight(wsXml, insertAt, 1)
    wsXml = shiftWorksheetMergeRefsRight(wsXml, insertAt, 1)
    totalCol++

    cells = parseWorksheetCells(wsXml, sharedStrings)
    byRow = rebuildRowsMap(cells)
    sortedRows = [...byRow.keys()].sort((a, b) => a - b)

    for (const rowIdx of sortedRows) {
      const rowXmlPattern = new RegExp(`<row\\s[^>]*\\br="${rowIdx + 1}"[^>]*>[\\s\\S]*?</row>`)
      const newRef = encodeCellRef(insertAt, rowIdx)

      wsXml = wsXml.replace(rowXmlPattern, (rowBlock) => {
        if (rowIdx < headerRowIdx) {
          return rowBlock
        }

        if (rowIdx === headerRowIdx) {
          return insertCellInRow(rowBlock, newRef, '', headerStyleAttr)
        }

        const rowCells = getSortedRowCells(byRow, rowIdx)
        const firstCell = rowCells[0]
        if (!firstCell?.textValue.trim()) return rowBlock

        const sampleDataCell = rowCells.find(
          (cell) => cell.col > labelCol && cell.col < totalCol && !cell.hasFormula,
        )
        return insertCellInRow(
          rowBlock,
          newRef,
          0,
          sampleDataCell?.styleAttr ?? headerStyleAttr,
        )
      })
    }

    doctorCols.push({ col: insertAt, styleAttr: headerStyleAttr })

    cells = parseWorksheetCells(wsXml, sharedStrings)
    byRow = rebuildRowsMap(cells)
    sortedRows = [...byRow.keys()].sort((a, b) => a - b)
  }

  const doctorData = orderedDoctors.map((entry) => {
    const values = new Map<string, number>()
    for (const [label, value] of entry.values) {
      values.set(normalizeName(label), value)
    }
    return { displayName: entry.displayName, values }
  })

  for (let index = 0; index < doctorCols.length; index++) {
    const col = doctorCols[index].col
    const entry = doctorData[index]
    const headerRef = encodeCellRef(col, headerRowIdx)
    wsXml = setCellValue(wsXml, headerRef, entry?.displayName ?? '')

    if (!entry) continue

    for (const rowIdx of sortedRows) {
      if (rowIdx <= headerRowIdx) continue

      const rowCells = getSortedRowCells(byRow, rowIdx)
      const labelCell = rowCells.find((cell) => cell.col === labelCol)
      const rawLabel = labelCell?.textValue.trim()
      if (!rawLabel) continue

      const value = entry.values.get(normalizeName(rawLabel)) ?? 0
      if (value === 0) continue

      wsXml = setCellValue(wsXml, encodeCellRef(col, rowIdx), value)
    }
  }

  cells = parseWorksheetCells(wsXml, sharedStrings)
  byRow = rebuildRowsMap(cells)
  sortedRows = [...byRow.keys()].sort((a, b) => a - b)

  for (const rowIdx of sortedRows) {
    if (rowIdx <= headerRowIdx) continue

    const rowCells = getSortedRowCells(byRow, rowIdx)
    const labelCell = rowCells.find((cell) => cell.col === labelCol)
    const rawLabel = labelCell?.textValue.trim()
    if (!rawLabel) continue

    const rowLabel = normalizeName(rawLabel)
    const totalCell = rowCells.find((cell) => cell.col === totalCol)
    const total = doctorCols.reduce((sum, doctorCol) => {
      return sum + getCellNumericValue(rowCells.find((cell) => cell.col === doctorCol.col))
    }, 0)

    if (totalCell?.hasFormula || rowLabel.startsWith('total')) {
      const firstDoctorCol = doctorCols[0]?.col
      const lastDoctorCol = doctorCols[doctorCols.length - 1]?.col
      if (firstDoctorCol !== undefined && lastDoctorCol !== undefined && totalCell?.hasFormula) {
        wsXml = setCellFormula(
          wsXml,
          encodeCellRef(totalCol, rowIdx),
          `SUM(${encodeCellRef(firstDoctorCol, rowIdx)}:${encodeCellRef(lastDoctorCol, rowIdx)})`,
        )
      }
      wsXml = setCellValue(wsXml, encodeCellRef(totalCol, rowIdx), total)
    }
  }

  return normalizeSheetFormulas(wsXml, sharedStrings)
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
  // Load the embedded template as a ZIP — keeps every internal file intact
  const templateBytes = Uint8Array.from(atob(TEMPLATE_BASE64), (c) => c.charCodeAt(0))
  const zip = await JSZip.loadAsync(templateBytes)

  // Parse shared strings (needed to resolve text cell values)
  const ssFile = zip.file('xl/sharedStrings.xml')
  const sharedStrings = ssFile ? parseSharedStrings(await ssFile.async('text')) : []

  // Get the sheet name → file path mapping
  const sheetEntries = await getWorkbookSheets(zip)

  const sheetProcessors: Array<[string, (xml: string) => string]> = [
    ['Nacionalidad',     (xml) => processLabeledSheet(xml, sharedStrings, output.nacionalidad)],
    ['Aseguramiento',   (xml) => processLabeledSheet(xml, sharedStrings, output.aseguramiento)],
    ['Consolidado',     (xml) => processLabeledSheet(xml, sharedStrings, output.consolidado)],
    ['Informe Económico',(xml) => processInformeEconomicoSheet(xml, sharedStrings, output.informeEconomico)],
    ['Monitoreo',       (xml) => processMonitoreoSheet(xml, sharedStrings, output.monitoreo)],
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
    const updatedXml = processor(originalXml)
    zip.file(entry.path, updatedXml)
  }

  await removeCalcChain(zip)

  // Generate the output ZIP as a Blob — same compression as original
  const outBlob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })

  saveAs(outBlob, `OdoLens_Consolidado_${todayISO()}.xlsx`)
}
