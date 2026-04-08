import type { WorkBook } from 'xlsx'

// ── File state ──────────────────────────────────────────────────────────────

export type FileStatus = 'pending' | 'valid' | 'invalid'

export interface ValidationError {
  sheet: string
  message: string
}

export interface FileEntry {
  id: string
  file: File
  name: string
  sizeKB: number
  status: FileStatus
  errors: ValidationError[]
  workbook?: ParsedWorkbook
}

// ── Parsed workbook ─────────────────────────────────────────────────────────

export type CellValue = string | number | null

export interface ParsedSheet {
  name: string
  rows: CellValue[][]
}

export interface ParsedWorkbook {
  fileName: string
  sheets: Record<string, ParsedSheet>
  /** Raw SheetJS workbook (kept for potential future use) */
  raw?: WorkBook
}

// ── Consolidation output ────────────────────────────────────────────────────

export interface LabeledSheet {
  /** Column headers (row 0 of the original sheet) */
  headers: CellValue[]
  /**
   * Ordered list of row labels as they appear in the first-seen file.
   * Preserves original display text.
   */
  rowOrder: string[]
  /**
   * Map from row-label → array of summed numeric values
   * (one entry per column after the label column)
   */
  data: Map<string, number[]>
}

export interface DoctorEntry {
  /** Original casing from first occurrence */
  displayName: string
  /** Map from procedure-label → summed value */
  values: Map<string, number>
}

export interface InformeEconomicoSheet {
  /** Numeric cells that must be summed across all files */
  sumCells: Map<string, number>
  /** Non-numeric cells copied from the first non-empty source */
  copiedCells: Map<string, CellValue>
}

export interface MonitoreoSheet {
  /** Fixed column headers before doctor columns (e.g. procedure name col) */
  fixedHeaders: CellValue[]
  /** Ordered procedure row labels */
  rowOrder: string[]
  /** Map from normalizedDoctorName → DoctorEntry */
  doctors: Map<string, DoctorEntry>
}

export interface ConsolidatedOutput {
  nacionalidad: LabeledSheet
  aseguramiento: LabeledSheet
  consolidado: LabeledSheet
  informeEconomico: InformeEconomicoSheet
  monitoreo: MonitoreoSheet
  fileCount: number
  doctorCount: number
  /** Raw SheetJS workbook of the first valid file */
  templateWorkbook: WorkBook
}

export type ReportTab = 'estadistico' | 'ejecutivo'
