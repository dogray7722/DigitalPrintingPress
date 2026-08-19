import { Download, CheckCircle2, AlertCircle, Loader2, Sparkles } from 'lucide-react'
import { format, addDays } from 'date-fns'
import { useWizardStore } from '../../store/wizardStore'
import { useExcelGeneration } from '../../hooks/useExcelGeneration'
import { computeTotalBudget } from '../../lib/utils'
import { getCurrencySymbol } from '../../data/currencies'
import { THEME_LABELS, THEMES } from '../../types/theme'
import { TRIP_STYLE_PRESETS } from '../../data/tripStylePresets'
import { ALL_SHEET_IDS, SHEET_LABELS } from '../../types/wizard'

function hexFromArgb(argb: string) {
  return '#' + argb.slice(2)
}

export function Step4Preview() {
  const state = useWizardStore()
  const { status, error, generate, reset, isGenerating, isResearching } = useExcelGeneration()

  const totalBudget = computeTotalBudget(state.budgets, state.duration)
  const sym = getCurrencySymbol(state.currency)
  const palette = THEMES[state.theme]

  const startFormatted = state.startDate
    ? format(new Date(state.startDate), 'dd MMM yyyy')
    : 'Not set'
  const endFormatted = state.startDate
    ? format(addDays(new Date(state.startDate), state.duration - 1), 'dd MMM yyyy')
    : '—'

  const enabledSheets = ALL_SHEET_IDS.filter((id) => {
    if (!state.sheets[id]) return false
    if (id === 'events' && !state.useRecommendations) return false
    return true
  })

  const fontFamilyLabels = { sans: 'Sans-Serif', serif: 'Serif', mono: 'Monospace' }

  const recommendationsBlocked = state.useRecommendations && !state.destination.trim()
  const canGenerate = !!state.startDate && (!state.useRecommendations || !!state.destination.trim())

  return (
    <div className="space-y-6">
      {/* Summary card */}
      <div className="rounded-2xl overflow-hidden border border-gray-200 shadow-sm">
        {/* Themed header */}
        <div
          className="px-5 py-4 text-center"
          style={{ backgroundColor: hexFromArgb(palette.primary) }}
        >
          <div
            className="font-bold text-lg tracking-wide"
            style={{ color: hexFromArgb(palette.primaryText) }}
          >
            ✈  {state.destination ? state.destination.toUpperCase() : 'YOUR DESTINATION'}
          </div>
          <div
            className="text-sm mt-0.5 opacity-80"
            style={{ color: hexFromArgb(palette.primaryText) }}
          >
            Travel Planner
            {state.useRecommendations && state.destination && (
              <span className="ml-2 inline-flex items-center gap-1">
                · <Sparkles size={11} /> AI-prefilled
              </span>
            )}
          </div>
        </div>

        <div className="bg-white divide-y divide-gray-100">
          <div className="grid grid-cols-2 divide-x divide-gray-100">
            <div className="px-4 py-3">
              <div className="text-xs text-gray-400 font-medium">DEPART</div>
              <div className="text-sm font-semibold text-gray-800 mt-0.5">{startFormatted}</div>
            </div>
            <div className="px-4 py-3">
              <div className="text-xs text-gray-400 font-medium">RETURN</div>
              <div className="text-sm font-semibold text-gray-800 mt-0.5">{endFormatted}</div>
            </div>
          </div>

          <div className="grid grid-cols-3 divide-x divide-gray-100">
            <div className="px-4 py-3">
              <div className="text-xs text-gray-400 font-medium">DURATION</div>
              <div className="text-sm font-semibold text-gray-800 mt-0.5">{state.duration} days</div>
            </div>
            <div className="px-4 py-3">
              <div className="text-xs text-gray-400 font-medium">TRAVELERS</div>
              <div className="text-sm font-semibold text-gray-800 mt-0.5">{state.partySize} people</div>
            </div>
            <div className="px-4 py-3">
              <div className="text-xs text-gray-400 font-medium">STYLE</div>
              <div className="text-sm font-semibold text-gray-800 mt-0.5">
                {TRIP_STYLE_PRESETS[state.tripStyle]?.icon} {state.tripStyle.charAt(0).toUpperCase() + state.tripStyle.slice(1)}
              </div>
            </div>
          </div>

          <div className="px-4 py-3 flex items-center justify-between">
            <div>
              <div className="text-xs text-gray-400 font-medium">TOTAL BUDGET</div>
              <div className="text-xl font-bold text-indigo-700 mt-0.5">
                {sym}{totalBudget.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
                <span className="text-sm font-normal text-gray-400 ml-1">{state.currency}</span>
              </div>
            </div>
            <div className="text-right">
              <div className="text-xs text-gray-400">Per person / day</div>
              <div className="text-sm font-semibold text-gray-700">
                {sym}{state.partySize > 0 && state.duration > 0 ? (totalBudget / state.partySize / state.duration).toFixed(0) : '—'}
              </div>
            </div>
          </div>

          <div className="grid grid-cols-2 divide-x divide-gray-100">
            <div className="px-4 py-3">
              <div className="text-xs text-gray-400 font-medium">THEME</div>
              <div className="flex items-center gap-1.5 mt-0.5">
                <div className="w-3 h-3 rounded-full" style={{ backgroundColor: hexFromArgb(palette.primary) }} />
                <div className="text-sm font-semibold text-gray-800">{THEME_LABELS[state.theme]}</div>
              </div>
            </div>
            <div className="px-4 py-3">
              <div className="text-xs text-gray-400 font-medium">FONT</div>
              <div className="text-sm font-semibold text-gray-800 mt-0.5">
                {fontFamilyLabels[state.fontFamily]} / {state.fontSize}
              </div>
            </div>
          </div>

          <div className="px-4 py-3">
            <div className="text-xs text-gray-400 font-medium mb-2">
              SHEETS INCLUDED ({enabledSheets.length + 2} tabs — OVERVIEW + INSTRUCTIONS always on)
            </div>
            <div className="flex flex-wrap gap-1.5">
              <span className="inline-flex items-center gap-1 text-xs bg-indigo-50 text-indigo-700 font-medium px-2 py-0.5 rounded-full border border-indigo-200">
                📊 Overview
              </span>
              {enabledSheets.map((id) => (
                <span
                  key={id}
                  className={`inline-flex items-center gap-1 text-xs px-2 py-0.5 rounded-full ${
                    id === 'events'
                      ? 'bg-purple-50 text-purple-700 border border-purple-200'
                      : 'bg-gray-100 text-gray-600'
                  }`}
                >
                  {SHEET_LABELS[id]}
                </span>
              ))}
              <span className="inline-flex items-center gap-1 text-xs bg-gray-100 text-gray-600 px-2 py-0.5 rounded-full">
                📖 Instructions
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Status feedback */}
      {status === 'success' && (
        <div className="flex items-center gap-3 px-4 py-3 bg-green-50 border border-green-200 rounded-xl text-green-800">
          <CheckCircle2 size={18} className="shrink-0 text-green-500" />
          <div className="text-sm">
            <div className="font-semibold">Download started!</div>
            <div className="text-green-600 text-xs mt-0.5">Open in Excel or import to Google Sheets.</div>
          </div>
          <button onClick={reset} className="ml-auto text-xs text-green-600 underline hover:no-underline">
            Reset
          </button>
        </div>
      )}

      {status === 'error' && (
        <div className="flex items-start gap-3 px-4 py-3 bg-red-50 border border-red-200 rounded-xl text-red-800">
          <AlertCircle size={18} className="shrink-0 text-red-500 mt-0.5" />
          <div className="text-sm flex-1">
            <div className="font-semibold">Generation failed</div>
            <div className="text-red-600 text-xs mt-0.5 break-words">{error}</div>
            {error?.includes('ANTHROPIC_API_KEY') && (
              <div className="text-red-600 text-xs mt-1.5">
                Copy <code className="bg-red-100 px-1 rounded">.env.example</code> to <code className="bg-red-100 px-1 rounded">.env</code> and add your Anthropic API key, then restart the dev server.
              </div>
            )}
          </div>
          <button onClick={reset} className="text-xs text-red-600 underline hover:no-underline shrink-0">
            Retry
          </button>
        </div>
      )}

      {/* Generate button */}
      <button
        type="button"
        onClick={() => generate(state)}
        disabled={isGenerating || !canGenerate}
        className="w-full flex items-center justify-center gap-3 py-4 px-6 rounded-2xl bg-gradient-to-r from-indigo-600 to-purple-600 text-white font-semibold text-base shadow-lg shadow-indigo-200 hover:from-indigo-500 hover:to-purple-500 active:scale-[0.98] transition-all disabled:opacity-50 disabled:cursor-not-allowed disabled:shadow-none"
      >
        {isResearching ? (
          <>
            <Sparkles size={20} className="animate-pulse" />
            Researching {state.destination} with AI…
          </>
        ) : isGenerating ? (
          <>
            <Loader2 size={20} className="animate-spin" />
            Generating your spreadsheet…
          </>
        ) : (
          <>
            <Download size={20} />
            Generate & Download .xlsx
          </>
        )}
      </button>

      {isResearching && (
        <p className="text-center text-xs text-purple-600">
          Searching the web for hotels, restaurants, excursions, itinerary ideas, and major events. This usually takes 20–60 seconds.
        </p>
      )}

      {recommendationsBlocked && (
        <p className="text-center text-xs text-amber-600">
          ⚠️ AI recommendations need a destination. Go back to Step 1 to enter one (or turn off the recommendations toggle).
        </p>
      )}

      {!state.startDate && (
        <p className="text-center text-xs text-amber-600">
          ⚠️ Please set a start date in Step 1 before generating.
        </p>
      )}

      <p className="text-center text-xs text-gray-400">
        File generates in your browser — no spreadsheet data is uploaded or stored.
      </p>
    </div>
  )
}
