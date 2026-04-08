export const REQUIRED_SHEETS = [
  'Nacionalidad',
  'Aseguramiento',
  'Consolidado',
  'Informe Económico',
  'Monitoreo',
] as const

export type SheetName = (typeof REQUIRED_SHEETS)[number]

// Column index of the row label (first column) in sheets 1–4
export const LABEL_COL_INDEX = 0

// In the Monitoreo sheet, these header values are NOT doctor names
export const MONITOREO_NON_DOCTOR_HEADERS = new Set([
  'total',
  'totales',
  'procedimiento',
  'procedimientos',
  'código',
  'codigo',
  'actividad',
  'actividades',
  'descripcion',
  'descripción',
  'indicador',
  'indicadores',
  '',
])

export const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024 // 50 MB
