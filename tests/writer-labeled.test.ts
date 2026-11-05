import { beforeAll, describe, expect, it } from 'vitest'
import { cell, makeReport, runPipeline, sheetXml } from './helpers/fixtures'
import type { PipelineResult } from './helpers/fixtures'

const cellXml = (xml: string, ref: string) => xml.match(new RegExp(`<c\\b(?=[^>]*\\br="${ref}")[^>]*?(?:/>|>[\\s\\S]*?</c>)`))?.[0]

describe('labeled sheets on the 2026 template', () => {
  let result: PipelineResult

  beforeAll(async () => {
    const patch = [
      { sheet: 2 as const, ref: 'B10', value: 3 },
      { sheet: 3 as const, ref: 'B24', value: 7 },
      { sheet: 1 as const, ref: 'B12', value: 6000 },
    ]
    result = await runPipeline([
      { name: 'centro A.xlsx', bytes: await makeReport('v2026', patch) },
      { name: 'centro B.xlsx', bytes: await makeReport('v2026', patch) },
    ])
  })

  it('year cells survive', () => {
    expect(cell(result.wb, 'Aseguramiento', 'G6')?.v).toBe(2026)
    expect(cell(result.wb, 'Consolidado', 'H5')?.v).toBe(2026)
    expect(cell(result.wb, 'Consolidado', 'I5')?.v ?? null).toBeNull()
  })

  it('cross-sheet caches', () => {
    expect(cell(result.wb, 'Aseguramiento', 'B10')?.v).toBe(6)
    expect(cell(result.wb, 'Consolidado', 'B9')).toEqual({ v: 6, f: "'Aseguramiento '!B10" })
  })

  it('subtotal in Consolidado', () => {
    expect(cell(result.wb, 'Consolidado', 'B24')?.v).toBe(14)
    expect(cell(result.wb, 'Consolidado', 'I23')).toEqual({ v: 14, f: 'SUM(I24:I26)' })
  })

  it('aggregated value above validation limit is kept', () => {
    expect(cell(result.wb, 'Nacionalidad', 'B12')?.v).toBe(12000)
  })

  it('zero sums render blank', async () => {
    const xml = await sheetXml(result.zip, 1)
    expect(cellXml(xml, 'C12')).not.toMatch(/<v>/)
    expect(cellXml(xml, 'C33')).not.toMatch(/<v>/)
  })

  it('zero values are hidden on every sheet after recalculation', async () => {
    for (const position of [1, 2, 3, 4, 5]) {
      const xml = await sheetXml(result.zip, position)
      expect(xml.match(/<sheetView\b[^>]*>/g)?.every((tag) => /\bshowZeros="0"/.test(tag)), `sheet ${position}`).toBe(true)
    }
  })

  it('output carries no personal metadata from the template', async () => {
    const core = await result.zip.file('docProps/core.xml')!.async('text')
    expect(core).toMatch(/<dc:creator>OdoLens<\/dc:creator>/)
    expect(core).toMatch(/<cp:lastModifiedBy>OdoLens<\/cp:lastModifiedBy>/)
  })

  it('calcPr has fullCalcOnLoad', async () => {
    const workbookXml = await result.zip.file('xl/workbook.xml')!.async('text')
    expect(workbookXml).toMatch(/<calcPr\b[^>]*\bfullCalcOnLoad="1"/)
  })
})

describe('informe economico on the 2026 template', () => {
  it('economico date lands on J16', async () => {
    const old = await makeReport('v2025', [
      { sheet: 4, ref: 'E13', value: 1000 },
      { sheet: 4, ref: 'F13', value: 200 },
      { sheet: 4, ref: 'I13', value: 'Agua' },
      { sheet: 4, ref: 'J13', value: 'Lab A' },
      { sheet: 4, ref: 'D15', value: 'Dra. A' },
      { sheet: 4, ref: 'J15', value: 46023 },
    ])
    const { wb } = await runPipeline([{ name: 'old.xlsx', bytes: old }])

    expect(cell(wb, 'Informe economico', 'J16')?.v).toBe(46023)
    expect(cell(wb, 'Informe economico', 'D16')?.v).toBe('Dra. A')
    expect(cell(wb, 'Informe economico', 'J15')?.v ?? null).toBeNull()
    expect(cell(wb, 'Informe economico', 'I13')?.v).toBe('Agua')
    expect(cell(wb, 'Informe economico', 'J13')?.v).toBe('Lab A')
    expect(cell(wb, 'Informe economico', 'E13')?.v).toBe(1000)
    expect(cell(wb, 'Informe economico', 'L13')).toEqual({ v: 800, f: 'E13-K13' })
  })
})
