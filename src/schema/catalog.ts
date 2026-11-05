/**
 * Canonical row identifiers for the four label-driven sheets.
 *
 * A RowId is the normalized label of the 2026 report template. Labels of older
 * report versions resolve to a RowId only through an approved alias; pending or
 * missing equivalences are reported as unmapped so their values never vanish
 * silently.
 */
import * as XLSX from 'xlsx'
import { TEMPLATE_V2026_BASE64 } from '../assets/templateV2026'
import { normalizeName } from '../lib/normalize'
import { SHEET_NAMES, findBlock, getSheet } from './reportSchemas'
import type { SheetKey } from './reportSchemas'

export type { SheetKey } from './reportSchemas'
export { SHEET_NAMES } from './reportSchemas'
export type RowId = string

export interface LabelAlias {
  /** Normalized retired label */
  from: string
  /** Normalized 2026 label, or null when there is no equivalent */
  to: string | null
  status: 'approved' | 'pending'
}

/**
 * Equivalences between retired labels and the 2026 catalog.
 * Only spelling fixes are approved; every other entry waits for the user's
 * decision (see Hermes: odolens-plan-migracion-informes, decision 2).
 */
const PROCEDURE_ALIASES: LabelAlias[] = [
  { from: 'recubrimiento directo', to: 'recubrimiento pulpar directo', status: 'pending' },
  { from: 'recubrimiento indirecto', to: 'recubrimiento pulpar indirecto', status: 'pending' },
  { from: 'endodoncias anteriores', to: 'endodoncia anteriores', status: 'pending' },
  { from: 'endodoncias posteriores', to: 'endodoncia posteriores', status: 'pending' },
  { from: 'exod. multiples con alveoloplastia / c', to: 'exod. multiples con alveoloplastia', status: 'pending' },
  { from: 'escision de lesion odontogenica', to: 'quiste y/o lesion odontogenica', status: 'pending' },
  { from: 'sellantes auto curado', to: 'sellantes fosas y fisuras', status: 'pending' },
  { from: 'sellantes foto curado', to: 'sellantes fosas y fisuras', status: 'pending' },
  { from: 'aplicacion resina preventiva', to: null, status: 'pending' },
  { from: 'mantenedor de espacio', to: null, status: 'pending' },
  { from: 'cirugia periodontal', to: null, status: 'pending' },
  { from: 'pulpectomias e instrumentacion', to: null, status: 'pending' },
  { from: 'obturacion en amalgama', to: null, status: 'pending' },
]

export const LABEL_ALIASES: Record<SheetKey, LabelAlias[]> = {
  nacionalidad: [
    { from: 'norteamericana', to: 'estadounidense', status: 'pending' },
    { from: 'checa', to: null, status: 'pending' },
    { from: 'dominicense', to: null, status: 'pending' },
    { from: 'grenadina', to: null, status: 'pending' },
    { from: 'jordana', to: null, status: 'pending' },
    { from: 'ruandesa', to: null, status: 'pending' },
  ],
  aseguramiento: [
    { from: 'ars plan saud banco central', to: 'ars plan salud banco central', status: 'approved' },
    { from: 'ars palic', to: 'ars mapfre', status: 'pending' },
  ],
  consolidado: PROCEDURE_ALIASES,
  monitoreo: PROCEDURE_ALIASES,
}

export type RowResolution =
  | { id: RowId }
  | { unmapped: true; normalized: string; suggestion?: RowId }

let catalogCache: Record<SheetKey, Set<RowId>> | null = null

/**
 * Normalized column-A labels of the data block of each sheet of the embedded
 * 2026 template — exactly the rows the writer can fill. Lists, signatures and
 * headers outside the blocks are not catalog rows.
 */
export function templateCatalog(): Record<SheetKey, Set<RowId>> {
  if (catalogCache) return catalogCache

  const bytes = Uint8Array.from(atob(TEMPLATE_V2026_BASE64), (c) => c.charCodeAt(0))
  const wb = XLSX.read(bytes, { type: 'array' })
  const catalog = {} as Record<SheetKey, Set<RowId>>

  for (const key of Object.keys(SHEET_NAMES) as SheetKey[]) {
    const ws = getSheet(wb, SHEET_NAMES[key])
    const block = findBlock(wb, key)
    const labels = new Set<RowId>()
    if (ws && block) {
      for (let row = block.firstRow; row <= block.lastRow; row++) {
        const label = normalizeName(String(ws[`A${row}`]?.v ?? ''))
        if (label) labels.add(label)
      }
    }
    catalog[key] = labels
  }

  catalogCache = catalog
  return catalog
}

export function resolveRowId(sheet: SheetKey, rawLabel: string): RowResolution {
  const normalized = normalizeName(rawLabel)
  if (templateCatalog()[sheet].has(normalized)) return { id: normalized }

  const alias = LABEL_ALIASES[sheet].find((entry) => entry.from === normalized)
  if (alias?.status === 'approved' && alias.to) return { id: alias.to }
  if (alias?.to) return { unmapped: true, normalized, suggestion: alias.to }
  return { unmapped: true, normalized }
}
