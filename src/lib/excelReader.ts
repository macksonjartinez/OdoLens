import * as XLSX from 'xlsx'
import type { ParsedWorkbook, CellValue } from '../types'

export async function readWorkbook(file: File): Promise<ParsedWorkbook> {
  const buffer = await file.arrayBuffer()
  const wb = XLSX.read(buffer, { type: 'array', cellDates: false })

  const sheets: ParsedWorkbook['sheets'] = {}

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name]
    const rows = XLSX.utils.sheet_to_json<CellValue[]>(ws, {
      header: 1,
      defval: null,
      blankrows: false,
    })
    // Trim sheet names to handle trailing/leading spaces in files
    const trimmedName = name.trim()
    sheets[trimmedName] = { name: trimmedName, rows }
  }

  return { fileName: file.name, sheets, raw: wb, contentHash: await sha256Hex(buffer) }
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', buffer)
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}
