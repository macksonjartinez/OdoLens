import { useRef, useState } from 'react'
import { Upload } from 'lucide-react'
import { useFileDropzone } from '../../hooks/useFileDropzone'
import type { FileEntry } from '../../types'

interface DropZoneProps {
  fileEntries: FileEntry[]
  onFilesAdded: (files: File[]) => void
}

export function DropZone({ fileEntries, onFilesAdded }: DropZoneProps) {
  const [isDragging, setIsDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

  const { handleDrop, handleDragOver, handleInputChange } = useFileDropzone({
    fileEntries,
    onFilesAdded,
  })

  return (
    <div
      onDragOver={(e) => { handleDragOver(e); setIsDragging(true) }}
      onDragEnter={() => setIsDragging(true)}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(e) => { handleDrop(e); setIsDragging(false) }}
      className={`relative flex flex-col items-center justify-center rounded-xl border-2 border-dashed py-14 px-8 transition-all ${
        isDragging
          ? 'border-olive-600 bg-olive-50'
          : 'border-olive-300 bg-white hover:border-olive-500 hover:bg-olive-50/40'
      }`}
    >
      {/* Upload icon */}
      <div className="mb-4 flex h-16 w-16 items-center justify-center rounded-full bg-olive-100">
        <Upload className="h-7 w-7 text-olive-600" strokeWidth={1.8} />
      </div>

      <p className="mb-1 text-lg font-semibold text-gray-800">Arrastre sus archivos aquí</p>
      <p className="mb-6 text-center text-sm text-gray-500">
        Soporta formato .xlsx de registros hospitalarios.
        <br />
        Tamaño máximo por archivo: 50MB.
      </p>

      <button
        onClick={() => inputRef.current?.click()}
        className="rounded-lg bg-olive-600 px-6 py-2.5 text-sm font-semibold text-white transition-colors hover:bg-olive-700 active:bg-olive-800 cursor-pointer"
      >
        Seleccionar Archivos
      </button>

      <input
        ref={inputRef}
        type="file"
        multiple
        accept=".xlsx"
        className="hidden"
        onChange={handleInputChange}
      />
    </div>
  )
}
