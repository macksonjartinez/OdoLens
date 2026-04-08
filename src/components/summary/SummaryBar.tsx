import { Database, Loader2 } from 'lucide-react'

interface SummaryBarProps {
  total: number
  ready: number
  isConsolidating: boolean
  onConsolidate: () => void
}

export function SummaryBar({ total, ready, isConsolidating, onConsolidate }: SummaryBarProps) {
  return (
    <div className="flex items-center justify-between rounded-xl bg-white border border-gray-100 shadow-sm px-5 py-3">
      <div>
        <p className="text-[10px] font-semibold uppercase tracking-widest text-gray-400">
          Resumen de Carga
        </p>
        <p className="text-sm text-gray-700 mt-0.5">
          <span className="font-semibold text-gray-900">{total}</span>{' '}
          {total === 1 ? 'archivo detectado' : 'archivos detectados'} —{' '}
          <span className="font-semibold text-olive-700">{ready}</span>{' '}
          {ready === 1 ? 'listo' : 'listos'} para consolidar
        </p>
      </div>

      <button
        onClick={onConsolidate}
        disabled={ready === 0 || isConsolidating}
        className={`flex items-center gap-2 rounded-full px-6 py-2.5 text-sm font-semibold text-white transition-all cursor-pointer ${
          ready > 0 && !isConsolidating
            ? 'bg-olive-600 hover:bg-olive-700 active:bg-olive-800 shadow-sm'
            : 'bg-gray-300 cursor-not-allowed'
        }`}
      >
        {isConsolidating ? (
          <>
            <Loader2 className="h-4 w-4 animate-spin" />
            Procesando…
          </>
        ) : (
          <>
            <Database className="h-4 w-4" />
            Consolidar Archivos
          </>
        )}
      </button>
    </div>
  )
}
