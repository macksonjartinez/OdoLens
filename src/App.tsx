import { useState, useCallback } from 'react'
import type { FileEntry, ReportTab } from './types'
import { useConsolidation } from './hooks/useConsolidation'
import { Header } from './components/layout/Header'
import { Footer } from './components/layout/Footer'
import { DropZone } from './components/dropzone/DropZone'
import { FileListPanel } from './components/filelist/FileListPanel'
import { SummaryBar } from './components/summary/SummaryBar'
import { SuccessScreen } from './components/success/SuccessScreen'

export default function App() {
  const [activeTab, setActiveTab] = useState<ReportTab>('estadistico')
  const [fileEntries, setFileEntries] = useState<FileEntry[]>([])

  const { isConsolidating, result, processFiles, runConsolidation, reset } =
    useConsolidation(setFileEntries)

  const handleFilesAdded = useCallback(
    (files: File[]) => {
      void processFiles(files)
    },
    [processFiles],
  )

  const handleRemove = useCallback((id: string) => {
    setFileEntries((prev) => prev.filter((e) => e.id !== id))
  }, [])

  const handleConsolidate = useCallback(() => {
    void runConsolidation(fileEntries)
  }, [runConsolidation, fileEntries])

  const handleReset = useCallback(() => {
    reset()
    setFileEntries([])
  }, [reset])

  const readyCount = fileEntries.filter((e) => e.status === 'valid').length

  return (
    <div className="flex min-h-screen flex-col">
      <Header activeTab={activeTab} onTabChange={setActiveTab} />

      {result ? (
        <SuccessScreen result={result} onReset={handleReset} />
      ) : (
        <main className="flex flex-1 gap-6 p-6">
          {/* Left column */}
          <div className="flex flex-1 flex-col gap-5 min-w-0">
            {/* Heading */}
            <div>
              <h1 className="font-display text-3xl font-bold text-gray-900">
                Consolidación de Datos Clínicos
              </h1>
              <p className="mt-2 text-sm text-gray-500">
                Unifique sus registros de salud de manera eficiente en el portal administrativo.
                Seleccione el propósito de su análisis para optimizar el proceso.
              </p>
            </div>

            {/* Drop zone */}
            <DropZone fileEntries={fileEntries} onFilesAdded={handleFilesAdded} />

            {/* Summary + consolidate */}
            <SummaryBar
              total={fileEntries.length}
              ready={readyCount}
              isConsolidating={isConsolidating}
              onConsolidate={handleConsolidate}
            />
          </div>

          {/* Right panel */}
          <div className="w-[360px] flex-shrink-0">
            <FileListPanel entries={fileEntries} onRemove={handleRemove} />
          </div>
        </main>
      )}

      <Footer />
    </div>
  )
}
