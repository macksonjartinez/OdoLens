import { describe, expect, it } from 'vitest'
import { normalizeName } from '../src/lib/normalize'
import { LABEL_ALIASES, resolveRowId, templateCatalog } from '../src/schema/catalog'
import type { SheetKey } from '../src/schema/catalog'

describe('normalizeName', () => {
  it('strips invisible format characters and non-breaking spaces', () => {
    expect(normalizeName('⁠Ajuste Oclusal ')).toBe('ajuste oclusal')
    expect(normalizeName('Prueba de  rodetes')).toBe('prueba de rodetes')
  })
})

describe('resolveRowId', () => {
  it('maps the approved spelling fix for ARS Plan Salud', () => {
    expect(resolveRowId('aseguramiento', 'ARS Plan Saud Banco Central')).toEqual({ id: 'ars plan salud banco central' })
  })

  it('reports pending aliases as unmapped with a suggestion', () => {
    expect(resolveRowId('consolidado', 'Sellantes Auto Curado')).toMatchObject({
      unmapped: true,
      suggestion: 'sellantes fosas y fisuras',
    })
  })

  it('reports labels without equivalent as unmapped', () => {
    expect(resolveRowId('consolidado', 'Mantenedor de Espacio')).toEqual({ unmapped: true, normalized: 'mantenedor de espacio' })
  })

  it('matches accent and U+2060 variants of 2026 labels', () => {
    expect(resolveRowId('monitoreo', 'Endodonticos')).toEqual({ id: 'endodonticos' })
    expect(resolveRowId('consolidado', 'Ajuste Oclusal')).toEqual({ id: 'ajuste oclusal' })
  })

  it('labels outside the template data blocks are not catalog rows', () => {
    expect(resolveRowId('monitoreo', 'Endodoncia')).toMatchObject({ unmapped: true })
    expect(resolveRowId('monitoreo', '01 - ENERO')).toMatchObject({ unmapped: true })
    expect(resolveRowId('aseguramiento', 'Mes')).toMatchObject({ unmapped: true })
  })

  it('every alias target exists in the 2026 catalog of its sheet', () => {
    const catalog = templateCatalog()
    for (const [sheet, aliases] of Object.entries(LABEL_ALIASES) as Array<[SheetKey, typeof LABEL_ALIASES[SheetKey]]>) {
      for (const alias of aliases) {
        if (alias.to) expect(catalog[sheet].has(alias.to), `${sheet}: ${alias.to}`).toBe(true)
        expect(catalog[sheet].has(alias.from), `${sheet}: ${alias.from} must be a retired label`).toBe(false)
      }
    }
  })
})
