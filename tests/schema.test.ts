import { describe, expect, it } from 'vitest'
import { detectSchema, findBlock, isOdoLensOutput } from '../src/schema/reportSchemas'
import { validateWorkbook } from '../src/lib/validator'
import { makeReport, readReport } from './helpers/fixtures'

const load = async (base: 'v2025' | 'v2026', patch: Parameters<typeof makeReport>[1] = []) =>
  readReport(`${base}.xlsx`, await makeReport(base, patch))

describe('detectSchema', () => {
  it('recognizes both report versions', async () => {
    expect(detectSchema((await load('v2025')).raw!)).toBe('v2025')
    expect(detectSchema((await load('v2026')).raw!)).toBe('v2026')
  })

  it('returns null and invalidates a report without known section anchors', async () => {
    const wb = await load('v2026', [
      { sheet: 3, ref: 'A23', value: 'Seccion X' },
      { sheet: 3, ref: 'A27', value: 'Seccion Y' },
    ])
    expect(detectSchema(wb.raw!)).toBeNull()
    const result = validateWorkbook(wb)
    expect(result.valid).toBe(false)
    expect(result.errors.map((e) => e.message)).toContain('Formato de informe no reconocido')
  })
})

describe('findBlock', () => {
  it('locates data blocks of the 2025 report by anchors', async () => {
    const raw = (await load('v2025')).raw!
    expect(findBlock(raw, 'nacionalidad')).toEqual({ firstRow: 10, lastRow: 33 })
    expect(findBlock(raw, 'aseguramiento')).toEqual({ firstRow: 9, lastRow: 35 })
    expect(findBlock(raw, 'consolidado')).toEqual({ firstRow: 9, lastRow: 81 })
    expect(findBlock(raw, 'monitoreo')).toEqual({ headerRow: 11, firstRow: 12, lastRow: 85 })
  })

  it('locates data blocks of the 2026 report by anchors', async () => {
    const raw = (await load('v2026')).raw!
    expect(findBlock(raw, 'nacionalidad')).toEqual({ firstRow: 10, lastRow: 33 })
    expect(findBlock(raw, 'aseguramiento')).toEqual({ firstRow: 10, lastRow: 37 })
    expect(findBlock(raw, 'consolidado')).toEqual({ firstRow: 9, lastRow: 112 })
    expect(findBlock(raw, 'monitoreo')).toEqual({ headerRow: 11, firstRow: 12, lastRow: 117 })
  })
})

describe('re-ingested OdoLens output', () => {
  it('is valid but carries a warning', async () => {
    const wb = await load('v2025', [{ sheet: 4, ref: 'D13', value: 'OdoLens_Consolidado_2026-07-27 (1)' }])
    expect(isOdoLensOutput(wb.raw!)).toBe(true)
    const result = validateWorkbook(wb)
    expect(result.valid).toBe(true)
    expect(result.errors).toHaveLength(1)
    expect(result.errors[0].severity).toBe('warning')
  })

  it('a regular report is not flagged', async () => {
    expect(isOdoLensOutput((await load('v2026')).raw!)).toBe(false)
    expect(validateWorkbook(await load('v2026')).errors).toEqual([])
  })
})
