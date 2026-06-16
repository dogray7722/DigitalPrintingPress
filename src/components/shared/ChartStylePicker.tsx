import { cn } from '../../lib/utils'
import type { ChartStyle } from '../../types/wizard'
import { Check } from 'lucide-react'

interface Props {
  value: ChartStyle
  onChange: (style: ChartStyle) => void
}

const OPTIONS: { id: ChartStyle; label: string; description: string; preview: string[] }[] = [
  {
    id: 'bar',
    label: 'Bar Chart',
    description: 'Spent vs budget — dotted track shows limit, turns red when over budget',
    preview: [
      'Flights  ████████▒▒▒▒  $800 of $1k',
      'Hotels   ██████▒▒▒▒▒▒  $600 of $1k',
      'Meals    ████████████▓  OVER! +$50',
    ],
  },
  {
    id: 'pie',
    label: 'Pie Chart',
    description: 'Real pie chart — updates live as you enter actual spending',
    preview: [
      '     ╭─────╮      Flights 38%',
      '    │ ◕   │     Hotels  33%',
      '     ╰─────╯      Food    18%',
    ],
  },
  {
    id: 'donut',
    label: 'Donut Chart',
    description: 'Real circular chart — updates live as you enter actual spending',
    preview: [
      '     ╭─────╮      Flights 38%',
      '    │  ◍  │     Hotels  33%',
      '     ╰─────╯      Food    18%',
    ],
  },
]

export function ChartStylePicker({ value, onChange }: Props) {
  return (
    <div className="flex flex-col gap-2">
      {OPTIONS.map((opt) => (
        <button
          key={opt.id}
          type="button"
          onClick={() => onChange(opt.id)}
          className={cn(
            'flex items-start gap-4 p-4 rounded-xl border-2 text-left transition-all',
            value === opt.id
              ? 'border-indigo-500 bg-indigo-50'
              : 'border-gray-200 bg-white hover:border-gray-300'
          )}
        >
          <div
            className={cn(
              'mt-0.5 w-5 h-5 rounded-full border-2 shrink-0 flex items-center justify-center transition-colors',
              value === opt.id ? 'border-indigo-500 bg-indigo-500' : 'border-gray-300'
            )}
          >
            {value === opt.id && <Check size={11} className="text-white" strokeWidth={3} />}
          </div>
          <div className="flex-1 min-w-0">
            <div className="font-medium text-sm text-gray-900">{opt.label}</div>
            <div className="text-xs text-gray-500 mb-2">{opt.description}</div>
            <div className="font-mono text-xs bg-gray-50 rounded-lg px-3 py-2 border border-gray-100 space-y-0.5">
              {opt.preview.map((line, i) => (
                <div key={i} className="text-gray-600 whitespace-pre">{line}</div>
              ))}
            </div>
          </div>
        </button>
      ))}
    </div>
  )
}
