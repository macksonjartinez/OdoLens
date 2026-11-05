import type { WorkSheet } from 'xlsx'
import { MONITOREO_NON_DOCTOR_HEADERS } from '../constants/sheetSchema'
import type {
  ParsedWorkbook,
  LabeledSheet,
  LabeledSheetKey,
  InformeEconomicoSheet,
  MonitoreoSheet,
  ConsolidatedOutput,
  DoctorEntry,
  InputCol,
  RowId,
  UnmappedRow,
} from '../types'
import { normalizeName, safeNum } from './normalize'
import { SHEET_NAMES, resolveRowId } from '../schema/catalog'
import {
  ECONOMICO_SHEET,
  INPUT_COLS,
  findBlock,
  findEconomicoSignatureRow,
  findScanStart,
  getSheet,
  isOdoLensOutput,
} from '../schema/reportSchemas'

const MONITOREO_PLACEHOLDER_DOCTOR_RE = /^dra?\s*\d+$/
const ECONOMICO_SUM_CELLS = ['E13', 'F13', 'G13', 'H13'] as const

// ── Cell helpers (1-based rows, A1 notation) ────────────────────────────────

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

function lastColIndex(ws: WorkSheet): number {
  const match = ws['!ref']?.match(/:([A-Z]+)\d+$/)
  if (!match) return 0
  let n = 0
  for (const ch of match[1]) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

/** Numeric value of a non-formula cell; formula cells are derived data and read as 0 */
function inputNumber(ws: WorkSheet, ref: string): number {
  const cell = ws[ref]
  if (!cell || cell.f) return 0
  return safeNum(cell.v)
}

function cellString(ws: WorkSheet, ref: string): string {
  const value = ws[ref]?.v
  return value === undefined || value === null ? '' : String(value).trim()
}

function emptyInputs(): Record<InputCol, number> {
  return { B: 0, C: 0, D: 0, E: 0, F: 0, G: 0, H: 0 }
}

// ── Doctor identity ──────────────────────────────────────────────────────────

function normalizeDoctorName(raw: string): string {
  return normalizeName(raw)
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
}

function isPlaceholderDoctorName(raw: string): boolean {
  return MONITOREO_PLACEHOLDER_DOCTOR_RE.test(normalizeDoctorName(raw))
}

function getPlaceholderDoctorKey(fileName: string, colIdx: number): string {
  const baseName = fileName.replace(/\.[^.]+$/, '').trim() || 'archivo'
  return `__placeholder__:${normalizeDoctorName(baseName)}:${colIdx}`
}

// ── Sheets 1–3: address-based capture blocks ────────────────────────────────

function consolidateLabeledSheet(
  sheet: LabeledSheetKey,
  workbooks: ParsedWorkbook[],
  unmapped: UnmappedRow[],
): LabeledSheet {
  const values = new Map<RowId, Record<InputCol, number>>()

  for (const wb of workbooks) {
    if (!wb.raw) continue
    const ws = getSheet(wb.raw, SHEET_NAMES[sheet])
    const block = findBlock(wb.raw, sheet)
    if (!ws || !block) continue

    const seenInFile = new Set<RowId>()
    const takeDuplicatesOnce = isOdoLensOutput(wb.raw)

    for (let row = findScanStart(wb.raw, sheet, block); row <= block.lastRow; row++) {
      const label = cellString(ws, `A${row}`)
      if (!label) continue

      const rowValues = emptyInputs()
      let total = 0
      for (const col of INPUT_COLS) {
        rowValues[col] = inputNumber(ws, `${col}${row}`)
        total += rowValues[col]
      }

      const resolved = resolveRowId(sheet, label)
      if ('unmapped' in resolved) {
        if (total !== 0) {
          unmapped.push({
            file: wb.fileName, sheet, row, label, total,
            ...(resolved.suggestion ? { suggestion: resolved.suggestion } : {}),
          })
        }
        continue
      }

      if (seenInFile.has(resolved.id) && takeDuplicatesOnce) continue
      seenInFile.add(resolved.id)

      const sums = values.get(resolved.id) ?? emptyInputs()
      for (const col of INPUT_COLS) sums[col] += rowValues[col]
      values.set(resolved.id, sums)
    }
  }

  return { values }
}

// ── Sheet 5: Monitoreo doctor-column merge ───────────────────────────────────

function consolidateMonitoreo(workbooks: ParsedWorkbook[], unmapped: UnmappedRow[]): MonitoreoSheet {
  const doctors = new Map<string, DoctorEntry>()
  let placeholderDoctorCount = 0

  for (const wb of workbooks) {
    if (!wb.raw) continue
    const ws = getSheet(wb.raw, SHEET_NAMES.monitoreo)
    const block = findBlock(wb.raw, 'monitoreo')
    if (!ws || !block?.headerRow) continue

    // Doctor columns run from B up to the total column: the first column whose
    // first data row holds a formula.
    const lastCol = lastColIndex(ws)
    const doctorColumns: Array<{ idx: number; letters: string; rawName: string; placeholder: boolean }> = []
    for (let idx = 1; idx <= lastCol; idx++) {
      const letters = colLetters(idx)
      if (ws[`${letters}${block.firstRow}`]?.f) break
      const rawName = cellString(ws, `${letters}${block.headerRow}`)
      if (!rawName || MONITOREO_NON_DOCTOR_HEADERS.has(normalizeName(rawName))) continue
      doctorColumns.push({ idx, letters, rawName, placeholder: isPlaceholderDoctorName(rawName) })
    }

    const rows: Array<{ row: number; label: string; id: RowId | null; suggestion?: RowId }> = []
    for (let row = block.firstRow; row <= block.lastRow; row++) {
      const label = cellString(ws, `A${row}`)
      if (!label) continue
      const resolved = resolveRowId('monitoreo', label)
      rows.push('unmapped' in resolved
        ? { row, label, id: null, suggestion: resolved.suggestion }
        : { row, label, id: resolved.id })
    }

    const unmappedTotals = new Map<number, number>()

    for (const column of doctorColumns) {
      const columnValues = rows.map((r) => inputNumber(ws, `${column.letters}${r.row}`))
      if (!columnValues.some((value) => value !== 0)) continue

      const doctorKey = column.placeholder
        ? getPlaceholderDoctorKey(wb.fileName, column.idx)
        : normalizeDoctorName(column.rawName)

      if (!doctors.has(doctorKey)) {
        doctors.set(doctorKey, {
          displayName: column.placeholder ? `Dr. ${++placeholderDoctorCount}` : column.rawName,
          values: new Map<RowId, number>(),
        })
      }
      const doctor = doctors.get(doctorKey)!

      rows.forEach((r, i) => {
        const value = columnValues[i]
        if (value === 0) return
        if (r.id === null) {
          unmappedTotals.set(i, (unmappedTotals.get(i) ?? 0) + value)
          return
        }
        doctor.values.set(r.id, (doctor.values.get(r.id) ?? 0) + value)
      })
    }

    for (const [i, total] of unmappedTotals) {
      const r = rows[i]
      unmapped.push({
        file: wb.fileName, sheet: 'monitoreo', row: r.row, label: r.label, total,
        ...(r.suggestion ? { suggestion: r.suggestion } : {}),
      })
    }
  }

  return { doctors }
}

// ── Sheet 4: Informe económico ───────────────────────────────────────────────

function consolidateInformeEconomico(workbooks: ParsedWorkbook[]): InformeEconomicoSheet {
  const sums = { E13: 0, F13: 0, G13: 0, H13: 0 }
  const concepts: string[] = []
  const suppliers: string[] = []
  let signer: string | null = null
  let sentDate: number | null = null

  const addDistinct = (list: string[], value: string) => {
    if (value && !list.some((existing) => normalizeName(existing) === normalizeName(value))) list.push(value)
  }

  for (const wb of workbooks) {
    if (!wb.raw) continue
    const ws = getSheet(wb.raw, ECONOMICO_SHEET)
    if (!ws) continue

    for (const ref of ECONOMICO_SUM_CELLS) sums[ref] += inputNumber(ws, ref)
    addDistinct(concepts, cellString(ws, 'I13'))
    addDistinct(suppliers, cellString(ws, 'J13'))

    const signatureRow = findEconomicoSignatureRow(wb.raw)
    if (signatureRow) {
      signer ??= cellString(ws, `D${signatureRow}`) || null
      const date = ws[`J${signatureRow}`]?.v
      if (sentDate === null && typeof date === 'number') sentDate = date
    }
  }

  return { sums, concept: concepts.join('; '), supplier: suppliers.join('; '), signer, sentDate }
}

// ── Main entry point ─────────────────────────────────────────────────────────

export function consolidate(workbooks: ParsedWorkbook[]): ConsolidatedOutput {
  const unmapped: UnmappedRow[] = []
  const nacionalidad = consolidateLabeledSheet('nacionalidad', workbooks, unmapped)
  const aseguramiento = consolidateLabeledSheet('aseguramiento', workbooks, unmapped)
  const consolidado = consolidateLabeledSheet('consolidado', workbooks, unmapped)
  const informeEconomico = consolidateInformeEconomico(workbooks)
  const monitoreo = consolidateMonitoreo(workbooks, unmapped)

  const warnings = workbooks
    .filter((wb) => wb.raw && isOdoLensOutput(wb.raw))
    .map((wb) => `${wb.fileName}: es un consolidado de OdoLens; sus filas duplicadas se tomaron una sola vez`)

  const byContent = new Map<string, string[]>()
  for (const wb of workbooks) {
    if (!wb.contentHash) continue
    byContent.set(wb.contentHash, [...(byContent.get(wb.contentHash) ?? []), wb.fileName])
  }
  for (const names of byContent.values()) {
    if (names.length > 1) {
      warnings.push(`${names.join(', ')} tienen el mismo contenido; sus cifras se sumaron ${names.length} veces`)
    }
  }

  return {
    nacionalidad,
    aseguramiento,
    consolidado,
    informeEconomico,
    monitoreo,
    fileCount: workbooks.length,
    doctorCount: monitoreo.doctors.size,
    unmapped,
    warnings,
  }
}
