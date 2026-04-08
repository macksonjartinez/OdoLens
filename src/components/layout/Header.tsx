import type { ReportTab } from '../../types'

interface HeaderProps {
  activeTab: ReportTab
  onTabChange: (tab: ReportTab) => void
}

export function Header({ activeTab, onTabChange }: HeaderProps) {
  return (
    <header className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-3">
      {/* Logo */}
      <div className="w-8 h-8 rounded-lg bg-olive-600 flex items-center justify-center flex-shrink-0">
        <svg width="18" height="18" viewBox="0 0 18 18" fill="none" xmlns="http://www.w3.org/2000/svg">
          <rect x="2" y="2" width="6" height="6" rx="1" fill="white" fillOpacity="0.9"/>
          <rect x="10" y="2" width="6" height="6" rx="1" fill="white" fillOpacity="0.6"/>
          <rect x="2" y="10" width="6" height="6" rx="1" fill="white" fillOpacity="0.6"/>
          <rect x="10" y="10" width="6" height="6" rx="1" fill="white" fillOpacity="0.9"/>
        </svg>
      </div>
      <span className="font-display font-700 text-lg text-gray-900 tracking-tight">OdoLens</span>

      {/* Spacer */}
      <div className="flex-1" />

      {/* Report type tabs */}
      <div className="flex gap-2">
        <button
          onClick={() => onTabChange('estadistico')}
          className={`px-4 py-1.5 rounded-full text-sm font-medium transition-colors cursor-pointer ${
            activeTab === 'estadistico'
              ? 'bg-olive-600 text-white'
              : 'border border-gray-300 text-gray-600 hover:border-olive-500 hover:text-olive-700'
          }`}
        >
          Informe Estadístico
        </button>
        <button
          onClick={() => onTabChange('ejecutivo')}
          className={`px-4 py-1.5 rounded-full text-sm font-medium transition-colors cursor-pointer ${
            activeTab === 'ejecutivo'
              ? 'bg-olive-600 text-white'
              : 'border border-gray-300 text-gray-600 hover:border-olive-500 hover:text-olive-700'
          }`}
        >
          Informe Ejecutivo
        </button>
      </div>
    </header>
  )
}
