import { useRef } from 'react'
import { FileCheck2, Upload, X } from 'lucide-react'

interface TemplateUploadProps {
  templateFile: File | null
  onTemplateSet: (file: File | null) => void
}

export function TemplateUpload({ templateFile, onTemplateSet }: TemplateUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null)

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file && file.name.toLowerCase().endsWith('.xlsx')) {
      onTemplateSet(file)
    }
    e.target.value = ''
  }

  return (
    <div className="flex items-center gap-3 rounded-lg border border-dashed border-gray-200 bg-white px-4 py-3">
      <div className={`flex h-8 w-8 flex-shrink-0 items-center justify-center rounded-md ${
        templateFile ? 'bg-olive-100 text-olive-600' : 'bg-gray-100 text-gray-400'
      }`}>
        {templateFile ? <FileCheck2 className="h-4 w-4" /> : <Upload className="h-4 w-4" />}
      </div>

      <div className="flex-1 min-w-0">
        <p className="text-[10px] font-semibold uppercase tracking-widest text-gray-400">
          Plantilla de Formato
        </p>
        {templateFile ? (
          <p className="truncate text-sm font-medium text-olive-700" title={templateFile.name}>
            {templateFile.name}
          </p>
        ) : (
          <p className="text-sm text-gray-400">
            Sin plantilla — se usará el primer archivo válido
          </p>
        )}
      </div>

      {templateFile ? (
        <button
          onClick={() => onTemplateSet(null)}
          className="flex-shrink-0 rounded p-1 text-gray-300 hover:text-red-400 transition-colors cursor-pointer"
          title="Quitar plantilla"
        >
          <X className="h-4 w-4" />
        </button>
      ) : (
        <button
          onClick={() => inputRef.current?.click()}
          className="flex-shrink-0 rounded-md border border-gray-200 px-3 py-1 text-xs font-medium text-gray-500 hover:border-olive-400 hover:text-olive-700 transition-colors cursor-pointer whitespace-nowrap"
        >
          Seleccionar
        </button>
      )}

      <input
        ref={inputRef}
        type="file"
        accept=".xlsx"
        className="hidden"
        onChange={handleChange}
      />
    </div>
  )
}
