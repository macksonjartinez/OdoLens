import { beforeAll, describe, expect, it } from 'vitest'
import { cell, makeReport, runPipeline, sheetXml } from './helpers/fixtures'
import type { CellPatch, PipelineResult } from './helpers/fixtures'

const DOCTOR_COLS_2026 = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE']
const mon = (ref: string, value: number | string): CellPatch => ({ sheet: 5, ref, value })

describe('monitoreo: two centers with one doctor each', () => {
  let result: PipelineResult

  beforeAll(async () => {
    const patch = [mon('B12', 3), mon('B29', 7), mon('B30', 2), mon('B117', 1)]
    result = await runPipeline([
      { name: 'centro A.xlsx', bytes: await makeReport('v2026', patch) },
      { name: 'centro B.xlsx', bytes: await makeReport('v2026', patch) },
    ])
  })

  it('vertical subtotal survives', () => {
    expect(cell(result.wb, 'Monitoreo', 'AF28')).toEqual({ v: 18, f: 'SUM(AF29:AF31)' })
  })

  it('placeholders of different files stay in separate columns', () => {
    expect(cell(result.wb, 'Monitoreo', 'B11')?.v).toBe('Dr. 1')
    expect(cell(result.wb, 'Monitoreo', 'C11')?.v).toBe('Dr. 2')
    expect(cell(result.wb, 'Monitoreo', 'D11')?.v ?? '').toBe('')
    expect(cell(result.wb, 'Monitoreo', 'C29')?.v).toBe(7)
  })

  it('per-doctor total cached', () => {
    expect(cell(result.wb, 'Monitoreo', 'B116')).toEqual({ v: 9, f: 'SUM(B23:B115)' })
    expect(cell(result.wb, 'Monitoreo', 'B15')?.f).toBe('SUM(B13:B14)')
    expect(cell(result.wb, 'Monitoreo', 'AF116')).toEqual({ v: 18, f: 'SUM(B116:AE116)' })
  })

  it('total column header reads TOTAL', () => {
    expect(cell(result.wb, 'Monitoreo', 'AF11')?.v).toBe('TOTAL')
  })

  it('no shared formula left', async () => {
    expect(await sheetXml(result.zip, 5)).not.toMatch(/t="shared"/)
  })
})

describe('monitoreo: 35 doctors expand past the template capacity', () => {
  let result: PipelineResult

  beforeAll(async () => {
    const full = DOCTOR_COLS_2026.map((col) => mon(`${col}29`, 1))
    const five = DOCTOR_COLS_2026.slice(0, 5).map((col) => mon(`${col}29`, 1))
    result = await runPipeline([
      { name: 'centro A.xlsx', bytes: await makeReport('v2026', full) },
      { name: 'centro B.xlsx', bytes: await makeReport('v2026', five) },
    ])
  })

  it('35 doctors expand to AK', () => {
    expect(result.output.doctorCount).toBe(35)
    expect(cell(result.wb, 'Monitoreo', 'AJ11')?.v).toBe('Dr. 35')
    expect(cell(result.wb, 'Monitoreo', 'AK11')?.v).toBe('TOTAL')
    expect(cell(result.wb, 'Monitoreo', 'AK12')?.f).toBe('SUM(B12:AJ12)')
    expect(cell(result.wb, 'Monitoreo', 'AK29')).toEqual({ v: 35, f: 'SUM(B29:AJ29)' })
    expect(cell(result.wb, 'Monitoreo', 'AK28')).toEqual({ v: 35, f: 'SUM(AK29:AK31)' })
  })

  it('inserted columns receive the per-doctor formulas', () => {
    expect(cell(result.wb, 'Monitoreo', 'AG15')?.f).toBe('SUM(AG13:AG14)')
    expect(cell(result.wb, 'Monitoreo', 'AG21')?.f).toBe('SUM(AG18:AG20)')
    expect(cell(result.wb, 'Monitoreo', 'AG116')).toEqual({ v: 1, f: 'SUM(AG23:AG115)' })
  })

  it('validations follow total column', async () => {
    const xml = await sheetXml(result.zip, 5)
    const sqrefs = [...xml.matchAll(/<dataValidation\b[^>]*\bsqref="([^"]+)"/g)].map((m) => m[1]).join(' ')
    expect(sqrefs.split(' ')).toContain('B23:AK117')
    expect(sqrefs.split(' ')).toContain('B22:AJ22')
    expect(sqrefs.split(' ')).toContain('AC119')
    expect(xml).toMatch(/<dimension ref="A1:AO153"\/>/)
    expect(xml).not.toMatch(/t="shared"/)
  })
})

describe('monitoreo: identified doctors', () => {
  it('a doctor repeated across files accumulates in one column', async () => {
    const patch = [mon('B11', 'Dra. Ana Perez'), mon('B29', 7)]
    const { wb, output } = await runPipeline([
      { name: 'centro A.xlsx', bytes: await makeReport('v2026', patch) },
      { name: 'centro B.xlsx', bytes: await makeReport('v2026', patch) },
    ])
    expect(output.doctorCount).toBe(1)
    expect(cell(wb, 'Monitoreo', 'B11')?.v).toBe('Dra. Ana Perez')
    expect(cell(wb, 'Monitoreo', 'B29')?.v).toBe(14)
  })
})
