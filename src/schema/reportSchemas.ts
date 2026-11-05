/**
 * Report versions and anchor-based data blocks.
 *
 * Both report versions share the same anchor labels, so blocks are located by
 * label instead of fixed row numbers: a center that inserts a row keeps working.
 */
import type { WorkBook, WorkSheet } from 'xlsx'
import { findSheetKey, normalizeName } from '../lib/normalize'

export type SheetKey = 'nacionalidad' | 'aseguramiento' | 'consolidado' | 'monitoreo'

export const SHEET_NAMES: Record<SheetKey, string> = {
  nacionalidad: 'Nacionalidad',
  aseguramiento: 'Aseguramiento',
  consolidado: 'Consolidado',
  monitoreo: 'Monitoreo',
}

export type SchemaId = 'v2025' | 'v2026'

export interface DataBlock {
  /** 1-based row of the doctor header (Monitoreo only) */
  headerRow?: number
  /** 1-based first data row */
  firstRow: number
  /** 1-based last data row (inclusive) */
  lastRow: number
}

interface BlockAnchors {
  header?: string
  start: string
  end: string
}

export const BLOCK_ANCHORS: Record<SheetKey, BlockAnchors> = {
  nacionalidad: { start: 'argentina', end: 'total' },
  aseguramiento: { start: 'no asegurado', end: 'total de pacientes' },
  consolidado: { start: 'no asegurado', end: 'pacientes de alta' },
  monitoreo: { header: 'nombre del dr', start: 'no asegurado', end: 'pacientes de alta' },
}

/** Data columns of Nacionalidad, Aseguramiento and Consolidado */
export const INPUT_COLS = ['B', 'C', 'D', 'E', 'F', 'G', 'H'] as const
export type InputCol = (typeof INPUT_COLS)[number]
export const LABELED_TOTAL_COL = 'I'

export const ECONOMICO_SHEET = 'Informe Económico'
export const ECONOMICO_SIGNATURE_LABEL = 'odontologo responsable'

export function getSheet(raw: WorkBook, name: string): WorkSheet | undefined {
  const key = findSheetKey(raw.Sheets, name)
  return key ? raw.Sheets[key] : undefined
}

function lastRowOf(ws: WorkSheet): number {
  const ref = ws['!ref']
  if (!ref) return 0
  const match = ref.match(/(\d+)$/)
  return match ? parseInt(match[1], 10) : 0
}

/** Normalized text of a cell (A1 notation) or '' */
export function cellText(ws: WorkSheet, ref: string): string {
  const value = ws[ref]?.v
  return value === undefined || value === null ? '' : normalizeName(String(value))
}

function findLabelRow(ws: WorkSheet, col: string, label: string, fromRow: number): number {
  const last = lastRowOf(ws)
  for (let row = fromRow; row <= last; row++) {
    if (cellText(ws, `${col}${row}`) === label) return row
  }
  return -1
}

export function findBlock(raw: WorkBook, sheet: SheetKey): DataBlock | null {
  const ws = getSheet(raw, SHEET_NAMES[sheet])
  if (!ws) return null

  const anchors = BLOCK_ANCHORS[sheet]
  let searchFrom = 1
  let headerRow: number | undefined

  if (anchors.header) {
    headerRow = findLabelRow(ws, 'A', anchors.header, 1)
    if (headerRow < 0) return null
    searchFrom = headerRow + 1
  }

  const firstRow = findLabelRow(ws, 'A', anchors.start, searchFrom)
  if (firstRow < 0) return null
  const lastRow = findLabelRow(ws, 'A', anchors.end, firstRow)
  if (lastRow < 0) return null

  return headerRow === undefined ? { firstRow, lastRow } : { headerRow, firstRow, lastRow }
}

const HEADER_TEXT_COLS = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']

/**
 * First row to read in a label-driven sheet: rows a center inserted between the
 * column headers (text in B:I) and the start anchor are read too, so they are
 * mapped or reported instead of silently skipped.
 */
export function findScanStart(raw: WorkBook, sheet: SheetKey, block: DataBlock): number {
  if (block.headerRow !== undefined) return block.firstRow
  const ws = getSheet(raw, SHEET_NAMES[sheet])
  if (!ws) return block.firstRow

  let row = block.firstRow
  while (row > 1 && !HEADER_TEXT_COLS.some((col) => typeof ws[`${col}${row - 1}`]?.v === 'string')) row--
  return row
}

function blockHasLabel(raw: WorkBook, sheet: SheetKey, label: string): boolean {
  const block = findBlock(raw, sheet)
  const ws = getSheet(raw, SHEET_NAMES[sheet])
  if (!block || !ws) return false
  for (let row = block.firstRow; row <= block.lastRow; row++) {
    if (cellText(ws, `A${row}`) === label) return true
  }
  return false
}

/**
 * v2026 introduced the Odontopediatria section; v2025 has Preventivos but not
 * Odontopediatria. Anything else is an unknown layout.
 */
export function detectSchema(raw: WorkBook): SchemaId | null {
  const sheets: SheetKey[] = ['nacionalidad', 'aseguramiento', 'consolidado', 'monitoreo']
  if (sheets.some((sheet) => !findBlock(raw, sheet))) return null

  if (blockHasLabel(raw, 'consolidado', 'odontopediatria') && blockHasLabel(raw, 'monitoreo', 'odontopediatria')) {
    return 'v2026'
  }
  if (blockHasLabel(raw, 'consolidado', 'preventivos')) return 'v2025'
  return null
}

/** Hidden defined name stamped by the writer on every consolidated output */
export const ODOLENS_OUTPUT_MARKER = 'OdoLens_Consolidado'

/** True when the workbook is itself an OdoLens consolidated output. */
export function isOdoLensOutput(raw: WorkBook): boolean {
  if (raw.Workbook?.Names?.some((name) => name.Name === ODOLENS_OUTPUT_MARKER)) return true
  // Outputs saved by older versions: the CELL("filename") cache after Excel saved them
  const ws = getSheet(raw, ECONOMICO_SHEET)
  const value = ws?.D13?.v
  return typeof value === 'string' && normalizeName(value).startsWith('odolens_consolidado')
}

/** 1-based row of the signer/date inputs: the row right above the signature caption. */
export function findEconomicoSignatureRow(raw: WorkBook): number | null {
  const ws = getSheet(raw, ECONOMICO_SHEET)
  if (!ws) return null
  const captionRow = findLabelRow(ws, 'D', ECONOMICO_SIGNATURE_LABEL, 1)
  return captionRow > 1 ? captionRow - 1 : null
}
