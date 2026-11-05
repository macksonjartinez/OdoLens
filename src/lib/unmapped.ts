import type { ConsolidatedOutput, UnmappedRow } from '../types'
import { normalizeName } from './normalize'

export interface UnmappedSummary {
  sheet: UnmappedRow['sheet']
  /** Label as first seen (trimmed) */
  label: string
  total: number
  /** Number of distinct source files with this row */
  files: number
  suggestion?: string
}

const SHEET_ORDER: Array<UnmappedRow['sheet']> = ['nacionalidad', 'aseguramiento', 'consolidado', 'monitoreo']

/** Group unmapped source rows by sheet and label for display before export. */
export function summarizeUnmapped(rows: UnmappedRow[]): UnmappedSummary[] {
  const groups = new Map<string, UnmappedSummary & { fileSet: Set<string> }>()

  for (const row of rows) {
    const key = `${row.sheet}|${normalizeName(row.label)}`
    const group = groups.get(key) ?? {
      sheet: row.sheet,
      label: row.label.trim(),
      total: 0,
      files: 0,
      fileSet: new Set<string>(),
      ...(row.suggestion ? { suggestion: row.suggestion } : {}),
    }
    group.total += row.total
    group.fileSet.add(row.file)
    groups.set(key, group)
  }

  return [...groups.values()]
    .map(({ fileSet, ...summary }) => ({ ...summary, files: fileSet.size }))
    .sort((a, b) => SHEET_ORDER.indexOf(a.sheet) - SHEET_ORDER.indexOf(b.sheet) || b.total - a.total)
}

/**
 * Export waits for the user's confirmation while some source row has no
 * equivalent or a warning (duplicated content, re-ingested output) was raised.
 */
export function requiresExportConfirmation(output: Pick<ConsolidatedOutput, 'unmapped' | 'warnings'>): boolean {
  return output.unmapped.length > 0 || output.warnings.length > 0
}
