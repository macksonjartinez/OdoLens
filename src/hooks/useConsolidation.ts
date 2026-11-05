import { useState, useCallback } from 'react'
import type { FileEntry, ConsolidatedOutput } from '../types'
import { readWorkbook } from '../lib/excelReader'
import { validateWorkbook } from '../lib/validator'
import { consolidate } from '../lib/consolidator'
import { writeConsolidatedWorkbook } from '../lib/excelWriter'
import { requiresExportConfirmation } from '../lib/unmapped'

interface UseConsolidationReturn {
  isConsolidating: boolean
  result: ConsolidatedOutput | null
  processFiles: (files: File[]) => Promise<void>
  runConsolidation: (entries: FileEntry[]) => Promise<ConsolidatedOutput | null>
  reset: () => void
}

export function useConsolidation(
  setEntries: React.Dispatch<React.SetStateAction<FileEntry[]>>,
): UseConsolidationReturn {
  const [isConsolidating, setIsConsolidating] = useState(false)
  const [result, setResult] = useState<ConsolidatedOutput | null>(null)

  const processFiles = useCallback(
    async (files: File[]) => {
      // Create pending entries
      const pendingEntries: FileEntry[] = files.map((file) => ({
        id: crypto.randomUUID(),
        file,
        name: file.name,
        sizeKB: Math.round(file.size / 1024),
        status: 'pending',
        errors: [],
      }))

      setEntries((prev) => [...prev, ...pendingEntries])

      // Process each file concurrently
      await Promise.all(
        pendingEntries.map(async (entry) => {
          try {
            const parsed = await readWorkbook(entry.file)
            const validation = validateWorkbook(parsed)

            setEntries((prev) =>
              prev.map((e) =>
                e.id === entry.id
                  ? {
                      ...e,
                      status: validation.valid ? 'valid' : 'invalid',
                      errors: validation.errors,
                      workbook: parsed,
                    }
                  : e,
              ),
            )
          } catch {
            setEntries((prev) =>
              prev.map((e) =>
                e.id === entry.id
                  ? {
                      ...e,
                      status: 'invalid',
                      errors: [{ sheet: 'general', message: 'Archivo dañado o formato incorrecto' }],
                    }
                  : e,
              ),
            )
          }
        }),
      )
    },
    [setEntries],
  )

  const runConsolidation = useCallback(
    async (entries: FileEntry[]): Promise<ConsolidatedOutput | null> => {
      const validEntries = entries.filter((e) => e.status === 'valid' && e.workbook)
      if (validEntries.length === 0) return null

      setIsConsolidating(true)
      try {
        const workbooks = validEntries.map((e) => e.workbook!)
        const output = consolidate(workbooks)
        // Rows without an equivalent in the 2026 template must be acknowledged
        // on the result screen before anything is downloaded.
        if (!requiresExportConfirmation(output)) {
          await writeConsolidatedWorkbook(output)
        }
        setResult(output)
        return output
      } finally {
        setIsConsolidating(false)
      }
    },
    [],
  )

  const reset = useCallback(() => {
    setResult(null)
  }, [])

  return { isConsolidating, result, processFiles, runConsolidation, reset }
}
