import { CheckCircle, AlertCircle, Loader2 } from 'lucide-react'
import type { FileStatus, ValidationError } from '../../types'

interface StatusBadgeProps {
  status: FileStatus
  errors?: ValidationError[]
}

export function StatusBadge({ status, errors }: StatusBadgeProps) {
  if (status === 'pending') {
    return (
      <span className="flex items-center gap-1 text-xs text-gray-400">
        <Loader2 className="h-3.5 w-3.5 animate-spin" />
        Procesando…
      </span>
    )
  }

  if (status === 'valid') {
    return (
      <span className="flex items-center gap-1 text-xs font-semibold text-emerald-600">
        <CheckCircle className="h-3.5 w-3.5" />
        LISTO
      </span>
    )
  }

  // invalid
  const errorMsg = errors?.[0]?.message ?? 'Formato incorrecto'
  return (
    <span
      className="flex items-center gap-1 text-xs font-semibold text-red-600"
      title={errorMsg}
    >
      <AlertCircle className="h-3.5 w-3.5 flex-shrink-0" />
      FORMATO INCORRECTO
    </span>
  )
}
