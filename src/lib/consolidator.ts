import { MONITOREO_NON_DOCTOR_HEADERS } from '../constants/sheetSchema'
import type {
  ParsedWorkbook,
  LabeledSheet,
  InformeEconomicoSheet,
  MonitoreoSheet,
  ConsolidatedOutput,
  CellValue,
  DoctorEntry,
} from '../types'
import { normalizeName, safeNum, findSheetKey } from './normalize'

const MONITOREO_HEADER_LABEL = 'nombre del dr'
const MONITOREO_PLACEHOLDER_DOCTOR_RE = /^dra?\s*\d+$/
const INFORME_ECONOMICO_SUM_CELLS = ['E13', 'F13', 'G13', 'H13'] as const
const INFORME_ECONOMICO_COPY_CELLS = ['I13', 'J13', 'D15', 'J15'] as const

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

function hasDoctorColumnData(rows: CellValue[][], colIdx: number): boolean {
  return rows.some((row) => safeNum(row[colIdx]) !== 0)
}

function findMonitoreoHeaderRowIndex(rows: CellValue[][]): number {
  return rows.findIndex(
    (row) => normalizeName(String(row[0] ?? '')) === MONITOREO_HEADER_LABEL,
  )
}

function findMonitoreoDataEndRowIndex(rows: CellValue[][], headerRowIdx: number): number {
  const endLabel = 'pacientes de alta'

  for (let rowIdx = rows.length - 1; rowIdx > headerRowIdx; rowIdx--) {
    if (normalizeName(String(rows[rowIdx]?.[0] ?? '')) === endLabel) {
      return rowIdx
    }
  }

  return rows.length - 1
}

function hasVisibleValue(value: CellValue): boolean {
  return value !== null && value !== ''
}

function getRawCellValue(
  workbook: ParsedWorkbook,
  sheetName: string,
  ref: string,
): CellValue {
  const key = findSheetKey(workbook.sheets, sheetName)
  if (!key || !workbook.raw) return null

  const sheet = workbook.raw.Sheets[key]
  const cell = sheet?.[ref]
  if (!cell) return null

  return cell.v ?? null
}

// ── Sheets 1–4: row-label summation ─────────────────────────────────────────

function consolidateLabeledSheet(
  sheetName: string,
  workbooks: ParsedWorkbook[],
): LabeledSheet {
  const headers: CellValue[] = []
  const rowOrder: string[] = []
  const data = new Map<string, number[]>()
  // Maps normalizedLabel → display label (first occurrence)
  const labelDisplay = new Map<string, string>()

  for (const wb of workbooks) {
    const key = findSheetKey(wb.sheets, sheetName)
    const sheet = key ? wb.sheets[key] : undefined
    if (!sheet || sheet.rows.length === 0) continue

    const [headerRow, ...dataRows] = sheet.rows

    // Capture headers from the first file that has this sheet
    if (headers.length === 0 && headerRow) {
      headers.push(...headerRow)
    }

    const colCount = Math.max(headers.length, headerRow?.length ?? 0) - 1

    for (const row of dataRows) {
      const rawLabel = String(row[0] ?? '').trim()
      if (!rawLabel) continue

      const key = normalizeName(rawLabel)

      if (!data.has(key)) {
        data.set(key, new Array<number>(colCount).fill(0))
        rowOrder.push(key)
        labelDisplay.set(key, rawLabel)
      }

      const sums = data.get(key)!
      for (let col = 1; col <= colCount; col++) {
        sums[col - 1] += safeNum(row[col])
      }
    }
  }

  // Restore display labels in rowOrder
  const displayRowOrder = rowOrder.map((key) => labelDisplay.get(key) ?? key)

  return { headers, rowOrder: displayRowOrder, data: remapToDisplay(data, labelDisplay) }
}

/** Remap the internal normalized-key map back to display keys for output */
function remapToDisplay(
  data: Map<string, number[]>,
  labelDisplay: Map<string, string>,
): Map<string, number[]> {
  const out = new Map<string, number[]>()
  for (const [key, vals] of data) {
    out.set(labelDisplay.get(key) ?? key, vals)
  }
  return out
}

// ── Sheet 5: Monitoreo doctor-column merge ───────────────────────────────────

function consolidateMonitoreo(workbooks: ParsedWorkbook[]): MonitoreoSheet {
  const rowOrder: string[] = []
  const rowDisplay = new Map<string, string>() // normalized → display
  const doctors = new Map<string, DoctorEntry>()
  let fixedHeaders: CellValue[] = []
  let placeholderDoctorCount = 0

  for (const wb of workbooks) {
    const monitoreoKey = findSheetKey(wb.sheets, 'Monitoreo')
    const sheet = monitoreoKey ? wb.sheets[monitoreoKey] : undefined
    if (!sheet || sheet.rows.length === 0) continue

    const headerRowIdx = findMonitoreoHeaderRowIndex(sheet.rows)
    if (headerRowIdx < 0) continue

    const headerRow = sheet.rows[headerRowIdx]
    const dataEndRowIdx = findMonitoreoDataEndRowIndex(sheet.rows, headerRowIdx)
    const dataRows = sheet.rows.slice(headerRowIdx + 1, dataEndRowIdx + 1)

    // Identify doctor columns (everything that isn't a known non-doctor header)
    const doctorColumns: Array<{
      idx: number
      rawName: string
      placeholder: boolean
      hasData: boolean
    }> = []
    headerRow.forEach((cell, idx) => {
      const rawName = String(cell ?? '').trim()
      const normalized = normalizeName(rawName)
      if (idx > 0 && hasVisibleValue(cell) && !MONITOREO_NON_DOCTOR_HEADERS.has(normalized)) {
        doctorColumns.push({
          idx,
          rawName,
          placeholder: isPlaceholderDoctorName(rawName),
          hasData: hasDoctorColumnData(dataRows, idx),
        })
      }
    })

    // Fixed headers = all headers that are NOT doctor columns
    if (fixedHeaders.length === 0) {
      const doctorColIndices = new Set(doctorColumns.map((column) => column.idx))
      fixedHeaders = headerRow.filter((_, idx) => !doctorColIndices.has(idx))
    }

    for (const row of dataRows) {
      const rawLabel = String(row[0] ?? '').trim()
      if (!rawLabel) continue

      const labelKey = normalizeName(rawLabel)
      if (!rowDisplay.has(labelKey)) {
        rowOrder.push(labelKey)
        rowDisplay.set(labelKey, rawLabel)
      }

      for (const column of doctorColumns) {
        if (!column.hasData) continue

        const doctorKey = column.placeholder
          ? getPlaceholderDoctorKey(wb.fileName, column.idx)
          : normalizeDoctorName(column.rawName)

        if (!doctors.has(doctorKey)) {
          const doctorDisplayName = column.placeholder
            ? `Dr. ${++placeholderDoctorCount}`
            : column.rawName

          doctors.set(doctorKey, {
            displayName: doctorDisplayName,
            values: new Map<string, number>(),
          })
        }

        const doctor = doctors.get(doctorKey)!
        const procedureLabel = rowDisplay.get(labelKey) ?? rawLabel
        const existing = doctor.values.get(procedureLabel) ?? 0
        doctor.values.set(procedureLabel, existing + safeNum(row[column.idx]))
      }
    }
  }

  const displayRowOrder = rowOrder.map((key) => rowDisplay.get(key) ?? key)

  return { fixedHeaders, rowOrder: displayRowOrder, doctors }
}

function consolidateInformeEconomico(workbooks: ParsedWorkbook[]): InformeEconomicoSheet {
  const sumCells = new Map<string, number>()
  const copiedCells = new Map<string, CellValue>()

  for (const ref of INFORME_ECONOMICO_SUM_CELLS) {
    sumCells.set(ref, 0)
  }

  for (const wb of workbooks) {
    for (const ref of INFORME_ECONOMICO_SUM_CELLS) {
      const current = sumCells.get(ref) ?? 0
      sumCells.set(ref, current + safeNum(getRawCellValue(wb, 'Informe Económico', ref)))
    }

    for (const ref of INFORME_ECONOMICO_COPY_CELLS) {
      if (copiedCells.has(ref) && hasVisibleValue(copiedCells.get(ref) ?? null)) continue

      const value = getRawCellValue(wb, 'Informe Económico', ref)
      if (hasVisibleValue(value)) {
        copiedCells.set(ref, value)
      }
    }
  }

  return { sumCells, copiedCells }
}

// ── Main entry point ─────────────────────────────────────────────────────────

export function consolidate(workbooks: ParsedWorkbook[]): ConsolidatedOutput {
  const nacionalidad = consolidateLabeledSheet('Nacionalidad', workbooks)
  const aseguramiento = consolidateLabeledSheet('Aseguramiento', workbooks)
  const consolidado = consolidateLabeledSheet('Consolidado', workbooks)
  const informeEconomico = consolidateInformeEconomico(workbooks)
  const monitoreo = consolidateMonitoreo(workbooks)

  return {
    nacionalidad,
    aseguramiento,
    consolidado,
    informeEconomico,
    monitoreo,
    fileCount: workbooks.length,
    doctorCount: monitoreo.doctors.size,
    templateWorkbook: workbooks[0].raw!,
  }
}
