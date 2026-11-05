import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import { consolidate } from '../src/lib/consolidator'
import { findSheetKey } from '../src/lib/normalize'
import { makeReport, readReport } from './helpers/fixtures'
import type { CellPatch, BaseTemplate } from './helpers/fixtures'

const load = async (name: string, base: BaseTemplate, patch: CellPatch[] = []) =>
  readReport(name, await makeReport(base, patch))

const ODOLENS_D13: CellPatch = { sheet: 4, ref: 'D13', value: 'OdoLens_Consolidado_2026-07-27 (1)' }

describe('labeled sheets', () => {
  it('sums Canadiense from two v2026 files', async () => {
    const a = await load('a.xlsx', 'v2026', [{ sheet: 1, ref: 'B12', value: 5 }])
    const b = await load('b.xlsx', 'v2026', [{ sheet: 1, ref: 'B12', value: 5 }])
    expect(consolidate([a, b]).nacionalidad.values.get('canadiense')?.B).toBe(10)
  })

  it('header and signature rows are not data', async () => {
    const old = await load('old.xlsx', 'v2025', [
      { sheet: 1, ref: 'A35', value: 'Dra. Firma' },
      { sheet: 1, ref: 'G35', value: 46023 },
    ])
    const out = consolidate([old, await load('new.xlsx', 'v2026')])
    const keys = [...out.aseguramiento.values.keys(), ...out.consolidado.values.keys(), ...out.nacionalidad.values.keys()]
    expect(keys).not.toContain('mes')
    expect(keys).not.toContain('dra. firma')
    expect(out.unmapped.map((u) => u.label)).not.toContain('Dra. Firma')
  })

  it('pending alias is reported unmapped', async () => {
    const old = await load('old.xlsx', 'v2025', [
      { sheet: 3, ref: 'D44', value: 6 },
      { sheet: 3, ref: 'G44', value: 20 },
    ])
    expect(consolidate([old]).unmapped).toContainEqual({
      file: 'old.xlsx', sheet: 'consolidado', row: 44, label: 'Recubrimiento Directo', total: 26,
      suggestion: 'recubrimiento pulpar directo',
    })
  })

  it('mixed batch reports unmapped totals', async () => {
    const old = await load('old.xlsx', 'v2025', [
      { sheet: 3, ref: 'D52', value: 3 },
      { sheet: 3, ref: 'D44', value: 6 },
    ])
    const fresh = await load('new.xlsx', 'v2026', [{ sheet: 3, ref: 'D68', value: 4 }])
    const out = consolidate([old, fresh])
    expect(out.consolidado.values.get('restauracion en resina')?.D).toBe(7)
    expect(out.unmapped.filter((u) => u.sheet === 'consolidado').map((u) => u.total)).toEqual([6])
  })

  it('row inserted inside block still maps', async () => {
    const wb = XLSX.read(await makeReport('v2026', [{ sheet: 1, ref: 'B12', value: 5 }]), { type: 'array', cellFormula: true })
    const key = findSheetKey(wb.Sheets, 'Nacionalidad')!
    const rows = XLSX.utils.sheet_to_json<Array<string | number | null>>(wb.Sheets[key], { header: 1, defval: null })
    rows.splice(11, 0, ['Nueva Nacionalidad', 1, null, null, null, null, null, null, null])
    wb.Sheets[key] = XLSX.utils.aoa_to_sheet(rows)
    const edited = await readReport('edited.xlsx', new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' })))

    const out = consolidate([edited])
    expect(out.nacionalidad.values.get('canadiense')?.B).toBe(5)
    expect(out.unmapped).toContainEqual({ file: 'edited.xlsx', sheet: 'nacionalidad', row: 12, label: 'Nueva Nacionalidad', total: 1 })
  })

  it('row inserted above the start anchor is reported', async () => {
    const wb = XLSX.read(await makeReport('v2026', [{ sheet: 1, ref: 'B12', value: 5 }]), { type: 'array', cellFormula: true })
    const key = findSheetKey(wb.Sheets, 'Nacionalidad')!
    const rows = XLSX.utils.sheet_to_json<Array<string | number | null>>(wb.Sheets[key], { header: 1, defval: null })
    rows.splice(9, 0, ['Alemana', 3, null, null, null, null, null, null, null])
    wb.Sheets[key] = XLSX.utils.aoa_to_sheet(rows)
    const edited = await readReport('edited.xlsx', new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' })))

    const out = consolidate([edited])
    expect(out.nacionalidad.values.get('canadiense')?.B).toBe(5)
    expect(out.unmapped).toContainEqual({ file: 'edited.xlsx', sheet: 'nacionalidad', row: 10, label: 'Alemana', total: 3 })
  })

  it('identical files under different names produce a warning', async () => {
    const bytes = await makeReport('v2026', [{ sheet: 1, ref: 'B12', value: 5 }])
    const out = consolidate([await readReport('centro.xlsx', bytes), await readReport('centro - copia.xlsx', bytes)])
    expect(out.warnings).toHaveLength(1)
    expect(out.warnings[0]).toContain('centro.xlsx')
    expect(out.warnings[0]).toContain('centro - copia.xlsx')
  })

  it('duplicated ARS Yunen rows: summed for raw reports, once for OdoLens outputs', async () => {
    const patch: CellPatch[] = [{ sheet: 2, ref: 'C31', value: 2 }, { sheet: 2, ref: 'C32', value: 2 }]
    const raw = await load('raw.xlsx', 'v2025', patch)
    const output = await load('out.xlsx', 'v2025', [...patch, ODOLENS_D13])
    expect(consolidate([raw]).aseguramiento.values.get('ars yunen')?.C).toBe(4)
    expect(consolidate([output]).aseguramiento.values.get('ars yunen')?.C).toBe(2)
  })
})

describe('informe economico', () => {
  it('signer/date read from D15/J15 (v2025) and D16/J16 (v2026)', async () => {
    const old = await load('old.xlsx', 'v2025', [{ sheet: 4, ref: 'D15', value: 'Dra. A' }, { sheet: 4, ref: 'J15', value: 46023 }])
    const fresh = await load('new.xlsx', 'v2026', [{ sheet: 4, ref: 'D16', value: 'Dra. B' }, { sheet: 4, ref: 'J16', value: 46100 }])
    expect(consolidate([old]).informeEconomico).toMatchObject({ signer: 'Dra. A', sentDate: 46023 })
    expect(consolidate([fresh]).informeEconomico).toMatchObject({ signer: 'Dra. B', sentDate: 46100 })
  })

  it('sums amounts and joins distinct concepts and suppliers', async () => {
    const a = await load('a.xlsx', 'v2026', [
      { sheet: 4, ref: 'E13', value: 100 }, { sheet: 4, ref: 'I13', value: 'Agua' }, { sheet: 4, ref: 'J13', value: 'Lab A' },
    ])
    const b = await load('b.xlsx', 'v2026', [
      { sheet: 4, ref: 'E13', value: 50 }, { sheet: 4, ref: 'I13', value: 'Gas' }, { sheet: 4, ref: 'J13', value: 'Lab A' },
    ])
    expect(consolidate([a, b]).informeEconomico).toMatchObject({
      sums: { E13: 150, F13: 0, G13: 0, H13: 0 }, concept: 'Agua; Gas', supplier: 'Lab A',
    })
  })
})

describe('monitoreo', () => {
  it('keys doctor values by row id and skips per-doctor formula rows', async () => {
    const a = await load('centro A.xlsx', 'v2026', [{ sheet: 5, ref: 'B29', value: 7 }])
    const b = await load('centro B.xlsx', 'v2026', [{ sheet: 5, ref: 'B29', value: 7 }])
    const out = consolidate([a, b])
    const doctors = [...out.monitoreo.doctors.values()]
    expect(doctors.map((d) => d.displayName)).toEqual(['Dr. 1', 'Dr. 2'])
    expect(doctors.map((d) => d.values.get('consulta odontopediatria'))).toEqual([7, 7])
    expect(doctors[0].values.has('total asegurados')).toBe(false)
  })

  it('a Monitoreo label that only exists outside the template table is reported', async () => {
    const report = await load('a.xlsx', 'v2026', [{ sheet: 5, ref: 'A29', value: 'Endodoncia' }, { sheet: 5, ref: 'B29', value: 7 }])
    expect(consolidate([report]).unmapped).toContainEqual({ file: 'a.xlsx', sheet: 'monitoreo', row: 29, label: 'Endodoncia', total: 7 })
  })

  it('reports retired procedures of 2025 reports as unmapped', async () => {
    const old = await load('old.xlsx', 'v2025', [{ sheet: 5, ref: 'B48', value: 3 }])
    expect(consolidate([old]).unmapped).toContainEqual({
      file: 'old.xlsx', sheet: 'monitoreo', row: 48, label: 'Recubrimiento Directo', total: 3,
      suggestion: 'recubrimiento pulpar directo',
    })
  })
})
