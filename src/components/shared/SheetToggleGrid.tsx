import { cn } from '../../lib/utils'
import type { SheetId } from '../../types/wizard'
import { SHEET_LABELS } from '../../types/wizard'
import { Sparkles } from 'lucide-react'

const SHEET_ICONS: Record<SheetId, string> = {
  budgetTracker: '💰',
  itinerary: '📅',
  hotels: '🏨',
  restaurants: '🍽️',
  excursions: '🎯',
  packingList: '🎒',
  flights: '✈️',
  tasks: '✅',
  events: '🎉',
}

const REGULAR_SHEET_IDS: SheetId[] = [
  'budgetTracker',
  'itinerary',
  'hotels',
  'restaurants',
  'excursions',
  'packingList',
  'flights',
  'tasks',
]

interface Props {
  sheets: Record<SheetId, boolean>
  useRecommendations: boolean
  onChange: (id: SheetId, enabled: boolean) => void
}

export function SheetToggleGrid({ sheets, useRecommendations, onChange }: Props) {
  return (
    <div>
      {/* OVERVIEW — always on */}
      <div className="flex items-center gap-3 p-3 rounded-lg bg-gray-50 border border-gray-200 mb-2">
        <span className="text-xl">📊</span>
        <div className="flex-1">
          <div className="text-sm font-medium text-gray-900">Overview</div>
          <div className="text-xs text-gray-500">Dashboard with budget summary — always included</div>
        </div>
        <div className="w-10 h-6 bg-indigo-500 rounded-full flex items-center justify-end px-1 cursor-not-allowed">
          <div className="w-4 h-4 bg-white rounded-full shadow-sm" />
        </div>
      </div>

      <div className="grid grid-cols-2 gap-2">
        {REGULAR_SHEET_IDS.map((id) => {
          const enabled = sheets[id]
          return (
            <button
              key={id}
              type="button"
              onClick={() => onChange(id, !enabled)}
              className={cn(
                'flex items-center gap-3 p-3 rounded-lg border-2 text-left transition-all',
                enabled
                  ? 'border-indigo-400 bg-indigo-50'
                  : 'border-gray-200 bg-white hover:border-gray-300'
              )}
            >
              <span className="text-xl shrink-0">{SHEET_ICONS[id]}</span>
              <div className="flex-1 min-w-0">
                <div className={cn('text-sm font-medium', enabled ? 'text-indigo-800' : 'text-gray-700')}>
                  {SHEET_LABELS[id]}
                </div>
              </div>
              <div
                className={cn(
                  'w-9 h-5 rounded-full flex items-center px-0.5 transition-colors shrink-0',
                  enabled ? 'bg-indigo-500 justify-end' : 'bg-gray-300 justify-start'
                )}
              >
                <div className="w-4 h-4 bg-white rounded-full shadow-sm" />
              </div>
            </button>
          )
        })}
      </div>

      {/* EVENTS — only visible when AI recommendations are enabled */}
      {useRecommendations && (
        <button
          type="button"
          onClick={() => onChange('events', !sheets.events)}
          className={cn(
            'mt-2 w-full flex items-center gap-3 p-3 rounded-lg border-2 text-left transition-all',
            sheets.events
              ? 'border-purple-400 bg-gradient-to-r from-purple-50 to-indigo-50'
              : 'border-gray-200 bg-white hover:border-gray-300'
          )}
        >
          <span className="text-xl shrink-0">🎉</span>
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-1.5">
              <span className={cn('text-sm font-medium', sheets.events ? 'text-purple-800' : 'text-gray-700')}>
                {SHEET_LABELS.events}
              </span>
              <Sparkles size={11} className="text-purple-500" />
            </div>
            <div className="text-xs text-gray-500">Top 20 festivals & annual events at your destination</div>
          </div>
          <div
            className={cn(
              'w-9 h-5 rounded-full flex items-center px-0.5 transition-colors shrink-0',
              sheets.events ? 'bg-purple-500 justify-end' : 'bg-gray-300 justify-start'
            )}
          >
            <div className="w-4 h-4 bg-white rounded-full shadow-sm" />
          </div>
        </button>
      )}

      {/* INSTRUCTIONS — always on */}
      <div className="flex items-center gap-3 p-3 rounded-lg bg-gray-50 border border-gray-200 mt-2">
        <span className="text-xl">📖</span>
        <div className="flex-1">
          <div className="text-sm font-medium text-gray-900">Instructions</div>
          <div className="text-xs text-gray-500">How-to guide — always last tab</div>
        </div>
        <div className="w-10 h-6 bg-indigo-500 rounded-full flex items-center justify-end px-1 cursor-not-allowed">
          <div className="w-4 h-4 bg-white rounded-full shadow-sm" />
        </div>
      </div>
    </div>
  )
}
