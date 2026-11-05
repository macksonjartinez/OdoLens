import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { TEMPLATE_BASE64 } from '../../src/assets/template'
import { TEMPLATE_V2026_BASE64 } from '../../src/assets/templateV2026'
import { readWorkbook } from '../../src/lib/excelReader'
import { validateWorkbook } from '../../src/lib/validator'
import { consolidate } from '../../src/lib/consolidator'
import { buildConsolidatedWorkbook } from '../../src/lib/excelWriter'
import type { ConsolidatedOutput, ParsedWorkbook } from '../../src/types'

export type BaseTemplate = 'v2025' | 'v2026'
export interface CellPatch {
  /** 1-based sheet position: 1 Nacionalidad, 2 Aseguramiento, 3 Consolidado, 4 Informe economico, 5 Monitoreo */
  sheet: 1 | 2 | 3 | 4 | 5
  ref: string
  value: number | string
}

const fromBase64 = (b64: string): Uint8Array => Uint8Array.from(Buffer.from(b64, 'base64'))

export function templateBytes(base: BaseTemplate): Uint8Array {
  return fromBase64(base === 'v2025' ? TEMPLATE_BASE64 : TEMPLATE_V2026_BASE64)
}

async function sheetPaths(zip: JSZip): Promise<string[]> {
  const wb = await zip.file('xl/workbook.xml')!.async('text')
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('text')
  return [...wb.matchAll(/<sheet\b[^>]*\br:id="([^"]+)"/g)].map(([, rid]) => {
    const target = rels.match(new RegExp(`<Relationship\\b[^>]*Id="${rid}"[^>]*Target="([^"]+)"`))![1]
    return target.startsWith('/') ? target.slice(1) : `xl/${target}`
  })
}

function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Replace a cell's content (formula included) with a literal value, keeping its style. */
function setXmlCell(xml: string, ref: string, value: number | string): string {
  const re = new RegExp(`<c\\b(?=[^>]*\\br="${ref}")([^>]*?)(?:/>|>[\\s\\S]*?</c>)`)
  if (!re.test(xml)) throw new Error(`fixture: cell ${ref} missing in template XML`)
  return xml.replace(re, (_m, attrs: string) => {
    const base = attrs.replace(/\s+t="[^"]*"/, '').replace(/\/$/, '')
    return typeof value === 'number'
      ? `<c${base}><v>${value}</v></c>`
      : `<c${base} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`
  })
}

/** Build a report file from one of the two known templates with literal cell overrides. */
export async function makeReport(base: BaseTemplate, patch: CellPatch[] = []): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(templateBytes(base))
  const paths = await sheetPaths(zip)
  for (const { sheet, ref, value } of patch) {
    const path = paths[sheet - 1]
    zip.file(path, setXmlCell(await zip.file(path)!.async('text'), ref, value))
  }
  return zip.generateAsync({ type: 'uint8array' })
}

export function asFile(name: string, bytes: Uint8Array): File {
  return { name, arrayBuffer: async () => bytes.slice().buffer } as unknown as File
}

export async function readReport(name: string, bytes: Uint8Array): Promise<ParsedWorkbook> {
  return readWorkbook(asFile(name, bytes))
}

export interface PipelineResult {
  output: ConsolidatedOutput
  bytes: Uint8Array
  wb: XLSX.WorkBook
  zip: JSZip
}

export async function runPipeline(files: Array<{ name: string; bytes: Uint8Array }>): Promise<PipelineResult> {
  const books: ParsedWorkbook[] = []
  for (const f of files) {
    const parsed = await readReport(f.name, f.bytes)
    const validation = validateWorkbook(parsed)
    if (!validation.valid) throw new Error(`${f.name} invalid: ${JSON.stringify(validation.errors)}`)
    books.push(parsed)
  }
  const output = consolidate(books)
  const bytes = await buildConsolidatedWorkbook(output)
  // sheetStubs keeps formula cells whose cached value was blanked (zeros render empty)
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, sheetStubs: true })
  return { output, bytes, wb, zip: await JSZip.loadAsync(bytes) }
}

const norm = (s: string) => s.trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')

export function sheetOf(wb: XLSX.WorkBook, name: string): XLSX.WorkSheet {
  const key = wb.SheetNames.find((n) => norm(n) === norm(name))
  if (!key) throw new Error(`sheet ${name} not found in ${wb.SheetNames.join(', ')}`)
  return wb.Sheets[key]
}

export function cell(wb: XLSX.WorkBook, sheet: string, ref: string): { v: unknown; f: string | null } | null {
  const c = sheetOf(wb, sheet)[ref]
  return c ? { v: c.v ?? null, f: c.f ?? null } : null
}

export async function sheetXml(zip: JSZip, position: number): Promise<string> {
  const paths = await sheetPaths(zip)
  return zip.file(paths[position - 1])!.async('text')
}
