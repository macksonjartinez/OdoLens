import { CheckCircle, Download, FileSpreadsheet, RefreshCw } from 'lucide-react'
import type { ConsolidatedOutput } from '../../types'
import { todayISO } from '../../lib/normalize'
import { writeConsolidatedWorkbook } from '../../lib/excelWriter'

interface SuccessScreenProps {
  result: ConsolidatedOutput
  onReset: () => void
}

export function SuccessScreen({ result, onReset }: SuccessScreenProps) {
  const filename = `OdoLens_Consolidado_${todayISO()}.xlsx`

  const handleDownload = () => {
    void writeConsolidatedWorkbook(result)
  }

  return (
    <div className="flex flex-1 items-center justify-center p-8">
      <div className="w-full max-w-lg rounded-2xl bg-white shadow-md border border-gray-100 overflow-hidden">
        {/* Top section */}
        <div className="flex flex-col items-center px-8 pt-10 pb-6 text-center">
          <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-full bg-olive-600">
            <CheckCircle className="h-9 w-9 text-white" strokeWidth={2} />
          </div>
          <h2 className="font-display text-2xl font-bold text-gray-900">
            ¡Consolidación Completada!
          </h2>
          <p className="mt-2 text-sm text-gray-500">
            Su archivo consolidado ha sido generado exitosamente.
          </p>
        </div>

        {/* Stats */}
        <div className="grid grid-cols-3 divide-x divide-gray-100 border-t border-gray-100">
          <div className="flex flex-col items-center py-5">
            <span className="text-3xl font-bold text-olive-600">{result.fileCount}</span>
            <span className="mt-1 text-xs text-gray-500 text-center">Archivos Procesados</span>
          </div>
          <div className="flex flex-col items-center py-5">
            <span className="text-3xl font-bold text-olive-600">5</span>
            <span className="mt-1 text-xs text-gray-500 text-center">Hojas Consolidadas</span>
          </div>
          <div className="flex flex-col items-center py-5">
            <span className="text-3xl font-bold text-olive-600">{result.doctorCount}</span>
            <span className="mt-1 text-xs text-gray-500 text-center">Médicos Registrados</span>
          </div>
        </div>

        {/* File info + download */}
        <div className="border-t border-gray-100 px-8 py-6 space-y-4">
          <p className="text-[10px] font-semibold uppercase tracking-widest text-gray-400">
            Archivo Generado
          </p>

          <div className="flex items-center gap-3 rounded-lg border border-gray-100 bg-gray-50 px-4 py-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-md bg-emerald-50 text-emerald-600">
              <FileSpreadsheet className="h-5 w-5" />
            </div>
            <div className="flex-1 min-w-0">
              <p className="truncate text-sm font-medium text-gray-800">{filename}</p>
            </div>
            <span className="flex items-center gap-1 text-xs font-semibold text-emerald-600">
              <CheckCircle className="h-3.5 w-3.5" />
              LISTO
            </span>
          </div>

          <button
            onClick={handleDownload}
            className="flex w-full items-center justify-center gap-2 rounded-lg bg-olive-600 py-3 text-sm font-semibold text-white hover:bg-olive-700 active:bg-olive-800 transition-colors cursor-pointer"
          >
            <Download className="h-4 w-4" />
            Descargar Archivo Consolidado
          </button>

          <button
            onClick={onReset}
            className="flex w-full items-center justify-center gap-2 rounded-lg border border-gray-200 py-2.5 text-sm font-medium text-gray-600 hover:border-olive-400 hover:text-olive-700 transition-colors cursor-pointer"
          >
            <RefreshCw className="h-4 w-4" />
            Procesar Nuevos Archivos
          </button>
        </div>
      </div>
    </div>
  )
}
