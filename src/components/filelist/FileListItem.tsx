import { FileSpreadsheet, X } from 'lucide-react'
import type { FileEntry } from '../../types'
import { StatusBadge } from './StatusBadge'

interface FileListItemProps {
  entry: FileEntry
  onRemove: (id: string) => void
}

export function FileListItem({ entry, onRemove }: FileListItemProps) {
  const errorMsg = entry.errors.find((error) => error.severity !== 'warning')?.message
  const warnings = entry.errors.filter((error) => error.severity === 'warning')

  return (
    <div className={`rounded-lg border p-3 ${
      entry.status === 'invalid' ? 'border-red-100 bg-red-50/40' : 'border-gray-100 bg-white'
    }`}>
      <div className="flex items-start gap-3">
        {/* Icon */}
        <div className={`mt-0.5 flex h-8 w-8 flex-shrink-0 items-center justify-center rounded-md ${
          entry.status === 'valid'
            ? 'bg-emerald-50 text-emerald-600'
            : entry.status === 'invalid'
            ? 'bg-red-50 text-red-500'
            : 'bg-gray-100 text-gray-400'
        }`}>
          <FileSpreadsheet className="h-4 w-4" />
        </div>

        {/* Info */}
        <div className="flex-1 min-w-0">
          <p className="truncate text-sm font-medium text-gray-800" title={entry.name}>
            {entry.name}
          </p>
          <p className="text-xs text-gray-400">{entry.sizeKB} KB</p>
          <div className="mt-1">
            <StatusBadge status={entry.status} errors={entry.errors} />
          </div>
          {entry.status === 'invalid' && errorMsg && (
            <p className="mt-1 text-xs italic text-gray-400 leading-snug">{errorMsg}</p>
          )}
          {entry.status === 'valid' && warnings.map((warning) => (
            <p key={warning.message} className="mt-1 text-xs text-amber-600 leading-snug">{warning.message}</p>
          ))}
        </div>

        {/* Remove */}
        <button
          onClick={() => onRemove(entry.id)}
          className="flex-shrink-0 rounded p-0.5 text-gray-300 hover:text-red-500 transition-colors cursor-pointer"
          title="Eliminar archivo"
        >
          <X className="h-4 w-4" />
        </button>
      </div>
    </div>
  )
}
