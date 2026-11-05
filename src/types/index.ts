import type { WorkBook } from 'xlsx'

// ── File state ──────────────────────────────────────────────────────────────

export type FileStatus = 'pending' | 'valid' | 'invalid'

export interface ValidationError {
  sheet: string
  message: string
  /** Warnings are shown but do not invalidate the file (default: error) */
  severity?: 'error' | 'warning'
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
  /** SHA-256 of the file bytes, to detect the same report loaded twice */
  contentHash?: string
}

// ── Consolidation output ────────────────────────────────────────────────────

/** Canonical row id: normalized label of the 2026 report template */
export type RowId = string
export type InputCol = 'B' | 'C' | 'D' | 'E' | 'F' | 'G' | 'H'
export type LabeledSheetKey = 'nacionalidad' | 'aseguramiento' | 'consolidado'

export interface LabeledSheet {
  /** Summed capture values per canonical row, by column letter */
  values: Map<RowId, Record<InputCol, number>>
}

export interface DoctorEntry {
  /** Original casing from first occurrence */
  displayName: string
  /** Map from canonical procedure row → summed value */
  values: Map<RowId, number>
}

export interface InformeEconomicoSheet {
  /** Amounts summed across all files */
  sums: Record<'E13' | 'F13' | 'G13' | 'H13', number>
  /** Distinct "otros gastos" concepts (I13), joined with "; " */
  concept: string
  /** Distinct suppliers (J13), joined with "; " */
  supplier: string
  /** First non-empty signer name */
  signer: string | null
  /** First sent date (Excel serial) */
  sentDate: number | null
}

export interface MonitoreoSheet {
  /** Map from doctor key → DoctorEntry, in first-seen order */
  doctors: Map<string, DoctorEntry>
}

/** A source row with data that has no approved equivalent in the 2026 template */
export interface UnmappedRow {
  file: string
  sheet: LabeledSheetKey | 'monitoreo'
  /** 1-based row in the source sheet */
  row: number
  label: string
  total: number
  /** Proposed 2026 row, pending approval */
  suggestion?: RowId
}

export interface ConsolidatedOutput {
  nacionalidad: LabeledSheet
  aseguramiento: LabeledSheet
  consolidado: LabeledSheet
  informeEconomico: InformeEconomicoSheet
  monitoreo: MonitoreoSheet
  fileCount: number
  doctorCount: number
  unmapped: UnmappedRow[]
  warnings: string[]
}

export type ReportTab = 'estadistico' | 'ejecutivo'
