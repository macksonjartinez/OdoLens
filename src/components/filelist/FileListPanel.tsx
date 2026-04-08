import { Info } from 'lucide-react'
import type { FileEntry } from '../../types'
import { FileListItem } from './FileListItem'

interface FileListPanelProps {
  entries: FileEntry[]
  onRemove: (id: string) => void
}

export function FileListPanel({ entries, onRemove }: FileListPanelProps) {
  const hasInvalid = entries.some((e) => e.status === 'invalid')

  return (
    <div className="flex h-full flex-col rounded-xl bg-white shadow-sm border border-gray-100">
      {/* Panel header */}
      <div className="flex items-center justify-between border-b border-gray-100 px-4 py-3">
        <h2 className="font-display font-semibold text-gray-800">Archivos Preparados</h2>
        <span className="rounded bg-gray-100 px-2 py-0.5 text-[10px] font-semibold uppercase tracking-widest text-gray-400">
          Portal v1.4
        </span>
      </div>

      {/* File list */}
      <div className="flex-1 overflow-y-auto p-3 space-y-2">
        {entries.length === 0 ? (
          <div className="flex h-full items-center justify-center py-12">
            <p className="text-sm text-gray-400">No hay archivos cargados</p>
          </div>
        ) : (
          entries.map((entry) => (
            <FileListItem key={entry.id} entry={entry} onRemove={onRemove} />
          ))
        )}
      </div>

      {/* Warning note */}
      {hasInvalid && (
        <div className="border-t border-gray-100 p-3">
          <div className="flex gap-2 rounded-lg bg-gray-50 p-3">
            <Info className="h-4 w-4 flex-shrink-0 mt-0.5 text-gray-400" />
            <p className="text-xs text-gray-500 leading-relaxed">
              Los archivos con <strong>formato incorrecto</strong> no serán incluidos en el proceso
              final de consolidación. Por favor corrija o elimine los archivos señalados.
            </p>
          </div>
        </div>
      )}
    </div>
  )
}
