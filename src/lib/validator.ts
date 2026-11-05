import { REQUIRED_SHEETS, MONITOREO_NON_DOCTOR_HEADERS } from '../constants/sheetSchema'
import type { ParsedWorkbook, ValidationError } from '../types'
import { normalizeName, findSheetKey } from './normalize'
import { detectSchema, isOdoLensOutput } from '../schema/reportSchemas'

export interface ValidationResult {
  valid: boolean
  errors: ValidationError[]
}

function findMonitoreoHeaderRowIndex(rows: Array<Array<string | number | null>>): number {
  return rows.findIndex((row) => normalizeName(String(row[0] ?? '')) === 'nombre del dr')
}

function result(errors: ValidationError[]): ValidationResult {
  return { valid: !errors.some((error) => error.severity !== 'warning'), errors }
}

export function validateWorkbook(wb: ParsedWorkbook): ValidationResult {
  const errors: ValidationError[] = []

  // 1. Check all required sheets exist — case-insensitive, accent-insensitive, trimmed
  for (const required of REQUIRED_SHEETS) {
    if (!findSheetKey(wb.sheets, required)) {
      errors.push({
        sheet: 'general',
        message: `Falta la hoja requerida: "${required}"`,
      })
    }
  }

  if (errors.length > 0) return result(errors)

  // 2. Validate Monitoreo has at least one doctor column
  const monitoreoKey = findSheetKey(wb.sheets, 'Monitoreo')
  const monitoreo = monitoreoKey ? wb.sheets[monitoreoKey] : undefined
  if (monitoreo && monitoreo.rows.length > 0) {
    const headerRowIdx = findMonitoreoHeaderRowIndex(monitoreo.rows)
    if (headerRowIdx < 0) {
      errors.push({
        sheet: 'Monitoreo',
        message: 'No se encontró la fila "Nombre del Dr" en la hoja Monitoreo',
      })
      return result(errors)
    }

    const headerRow = monitoreo.rows[headerRowIdx]
    const doctorCols = headerRow.filter((cell) => {
      const normalized = normalizeName(String(cell ?? ''))
      return normalized !== 'nombre del dr' &&
        normalized !== '' &&
        !MONITOREO_NON_DOCTOR_HEADERS.has(normalized)
    })

    if (doctorCols.length === 0) {
      errors.push({
        sheet: 'Monitoreo',
        message: 'No se encontraron columnas de médicos en la hoja Monitoreo',
      })
    }
  } else {
    errors.push({
      sheet: 'Monitoreo',
      message: 'La hoja Monitoreo está vacía',
    })
  }

  // 3. Known report layout (2025 or 2026)
  if (wb.raw && !detectSchema(wb.raw)) {
    errors.push({ sheet: 'general', message: 'Formato de informe no reconocido' })
  }

  // 4. A consolidated output fed back as a source double-counts its centers
  if (wb.raw && isOdoLensOutput(wb.raw)) {
    errors.push({
      sheet: 'general',
      severity: 'warning',
      message: 'Este archivo parece un consolidado de OdoLens; incluirlo junto con sus fuentes duplica las cifras',
    })
  }

  return result(errors)
}
