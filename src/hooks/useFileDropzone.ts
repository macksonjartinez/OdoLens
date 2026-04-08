import { useCallback } from 'react'
import { MAX_FILE_SIZE_BYTES } from '../constants/sheetSchema'
import type { FileEntry } from '../types'

interface UseFileDropzoneProps {
  fileEntries: FileEntry[]
  onFilesAdded: (files: File[]) => void
}

export function useFileDropzone({ fileEntries, onFilesAdded }: UseFileDropzoneProps) {
  const filterFiles = useCallback(
    (rawFiles: FileList | File[]): File[] => {
      const existingKeys = new Set(fileEntries.map((e) => `${e.name}__${e.sizeKB}`))
      const result: File[] = []

      for (const file of Array.from(rawFiles)) {
        // Extension check
        if (!file.name.toLowerCase().endsWith('.xlsx')) continue
        // Size check
        if (file.size > MAX_FILE_SIZE_BYTES) continue
        // Duplicate check
        const key = `${file.name}__${Math.round(file.size / 1024)}`
        if (existingKeys.has(key)) continue

        result.push(file)
        existingKeys.add(key) // prevent duplicates within the same drop
      }

      return result
    },
    [fileEntries],
  )

  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault()
      const files = filterFiles(e.dataTransfer.files)
      if (files.length > 0) onFilesAdded(files)
    },
    [filterFiles, onFilesAdded],
  )

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
  }, [])

  const handleInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      if (!e.target.files) return
      const files = filterFiles(e.target.files)
      if (files.length > 0) onFilesAdded(files)
      // Reset input so same file can be re-selected after removal
      e.target.value = ''
    },
    [filterFiles, onFilesAdded],
  )

  return { handleDrop, handleDragOver, handleInputChange }
}
