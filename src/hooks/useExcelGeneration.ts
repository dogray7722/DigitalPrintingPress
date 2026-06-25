import { useState, useCallback } from 'react'
import type { WizardState } from '../types/wizard'
import type { Recommendations, ExchangeRate } from '../lib/recommendations/types'

type Status = 'idle' | 'researching' | 'generating' | 'success' | 'error'

export function useExcelGeneration() {
  const [status, setStatus] = useState<Status>('idle')
  const [error, setError] = useState<string | null>(null)

  const generate = useCallback(async (state: WizardState) => {
    setError(null)

    let recommendations: Recommendations | undefined
    let exchangeRate: ExchangeRate | undefined

    try {
      if (state.useRecommendations && state.destination.trim()) {
        setStatus('researching')
        const { fetchRecommendations, fetchExchangeRate } = await import('../lib/recommendations/client')
        const needsRate = state.destinationCurrency &&
          state.destinationCurrency !== state.currency
        const [recs, rate] = await Promise.all([
          fetchRecommendations({
            destination: state.destination,
            duration: state.duration,
            partySize: state.partySize,
            travelMonth: state.travelMonth,
            tripStyle: state.tripStyle,
          }),
          needsRate
            ? fetchExchangeRate(state.currency, state.destinationCurrency)
            : Promise.resolve(null),
        ])
        recommendations = recs
        exchangeRate = rate ?? undefined
      }

      setStatus('generating')
      const { generateWorkbook } = await import('../lib/excel')
      const blob = await generateWorkbook(state, recommendations, exchangeRate)

      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      const dest = (state.destination || 'Travel').replace(/[^a-z0-9]/gi, '-')
      a.download = `${dest}-Travel-Planner.xlsx`
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)

      setStatus('success')
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Generation failed'
      setError(msg)
      setStatus('error')
    }
  }, [])

  const reset = useCallback(() => {
    setStatus('idle')
    setError(null)
  }, [])

  return {
    status,
    error,
    generate,
    reset,
    isGenerating: status === 'researching' || status === 'generating',
    isResearching: status === 'researching',
  }
}
