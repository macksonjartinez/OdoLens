import { describe, expect, it } from 'vitest'
import { requiresExportConfirmation, summarizeUnmapped } from '../src/lib/unmapped'
import type { UnmappedRow } from '../src/types'

const row = (over: Partial<UnmappedRow>): UnmappedRow => ({
  file: 'a.xlsx', sheet: 'consolidado', row: 44, label: 'Recubrimiento Directo', total: 1, ...over,
})

describe('summarizeUnmapped', () => {
  it('groups by sheet and label, sums totals and counts files', () => {
    const summary = summarizeUnmapped([
      row({ total: 26, suggestion: 'recubrimiento pulpar directo' }),
      row({ file: 'b.xlsx', label: 'Recubrimiento Directo ', total: 4, suggestion: 'recubrimiento pulpar directo' }),
      row({ sheet: 'monitoreo', row: 48, total: 24, suggestion: 'recubrimiento pulpar directo' }),
      row({ sheet: 'nacionalidad', row: 12, label: 'Checa', total: 2 }),
    ])

    expect(summary).toEqual([
      { sheet: 'nacionalidad', label: 'Checa', total: 2, files: 1 },
      { sheet: 'consolidado', label: 'Recubrimiento Directo', total: 30, files: 2, suggestion: 'recubrimiento pulpar directo' },
      { sheet: 'monitoreo', label: 'Recubrimiento Directo', total: 24, files: 1, suggestion: 'recubrimiento pulpar directo' },
    ])
  })

  it('orders rows of the same sheet by total, largest first', () => {
    const summary = summarizeUnmapped([row({ label: 'A', total: 1 }), row({ label: 'B', total: 9 })])
    expect(summary.map((s) => s.label)).toEqual(['B', 'A'])
  })
})

describe('requiresExportConfirmation', () => {
  it('is true when some row has no equivalent or a warning was raised', () => {
    expect(requiresExportConfirmation({ unmapped: [], warnings: [] })).toBe(false)
    expect(requiresExportConfirmation({ unmapped: [row({})], warnings: [] })).toBe(true)
    expect(requiresExportConfirmation({ unmapped: [], warnings: ['a.xlsx y b.xlsx tienen el mismo contenido'] })).toBe(true)
  })
})
