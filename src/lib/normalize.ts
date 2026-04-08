/** Normalize a string for deduplication: trim, lowercase, strip accents, collapse whitespace */
export function normalizeName(raw: string): string {
  return raw
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // strip accent marks
    .replace(/\s+/g, ' ')
}

/** Parse a cell value as a number; returns 0 for null/empty/non-numeric */
export function safeNum(val: unknown): number {
  if (val === null || val === undefined || val === '') return 0
  // Handle numbers stored as strings with commas e.g. "1,234"
  const cleaned = String(val).replace(/,/g, '').trim()
  const n = Number(cleaned)
  return isFinite(n) ? n : 0
}

/** Format today's date as YYYY-MM-DD */
export function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

/**
 * Find a sheet in a workbook by normalized name (case-insensitive, accent-insensitive, trimmed).
 * Returns the actual key used in the sheets map, or undefined if not found.
 */
export function findSheetKey(
  sheets: Record<string, unknown>,
  requiredName: string,
): string | undefined {
  const normalizedRequired = normalizeName(requiredName)
  return Object.keys(sheets).find((key) => normalizeName(key) === normalizedRequired)
}
