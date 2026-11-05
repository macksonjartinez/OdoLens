import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { cell, makeReport, readReport, runPipeline, templateBytes } from './helpers/fixtures'
import { isOdoLensOutput } from '../src/schema/reportSchemas'
import { validateWorkbook } from '../src/lib/validator'

describe('pipeline', () => {
  it('pipeline smoke: one v2026 report round-trips', async () => {
    const bytes = await makeReport('v2026', [{ sheet: 1, ref: 'B12', value: 5 }])
    const { wb } = await runPipeline([{ name: 'centro A.xlsx', bytes }])

    expect(cell(wb, 'Nacionalidad', 'A12')?.v).toBe('Canadiense')
  })

  it('the embedded 2026 template carries no personal metadata', async () => {
    const zip = await JSZip.loadAsync(templateBytes('v2026'))
    const core = await zip.file('docProps/core.xml')!.async('text')
    expect(core.match(/<dc:creator>([^<]*)<\/dc:creator>/)?.[1]).toBe('OdoLens')
    expect(core.match(/<cp:lastModifiedBy>([^<]*)<\/cp:lastModifiedBy>/)?.[1]).toBe('OdoLens')
  })

  it('a fresh OdoLens output fed back as a source is flagged', async () => {
    const source = await makeReport('v2026', [{ sheet: 5, ref: 'B29', value: 7 }])
    const { bytes } = await runPipeline([{ name: 'centro A.xlsx', bytes: source }])
    const reread = await readReport('consolidado renombrado.xlsx', bytes)

    expect(isOdoLensOutput(reread.raw!)).toBe(true)
    expect(validateWorkbook(reread).errors.map((e) => e.severity)).toEqual(['warning'])
  })
})
