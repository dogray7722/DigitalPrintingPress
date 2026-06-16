import { cn } from '../../lib/utils'
import type { ChartStyle } from '../../types/wizard'

interface Option {
  id: ChartStyle
  label: string
  description: string
  preview: string[]
}

const OPTIONS: Option[] = [
  {
    id: 'bar',
    label: 'Bar Chart',
    description: 'Spent vs budget per category',
    preview: [
      'Transport ████████▒▒▒▒',
      'Hotel     ██████▒▒▒▒▒▒',
      'Food      ████████████',
      'Activities ████▒▒▒▒▒▒▒',
    ],
  },
  {
    id: 'pie',
    label: 'Pie Chart',
    description: 'Share of spend by category',
    preview: [
      '  ╭─────╮',
      ' ╱╲  ·  ╱╲',
      '│  ╲   ╱  │',
      ' ╲  ╲ ╱  ╱',
      '  ╰─────╯',
    ],
  },
  {
    id: 'donut',
    label: 'Donut Chart',
    description: 'Share of spend by category',
    preview: [
      '  ╭─────╮',
      ' ╱ ╭───╮ ╲',
      '│  │ ◍ │  │',
      ' ╲ ╰───╯ ╱',
      '  ╰─────╯',
    ],
  },
]

interface Props {
  value: ChartStyle
  onChange: (s: ChartStyle) => void
}

export function ChartStylePicker({ value, onChange }: Props) {
  return (
    <div className="grid grid-cols-3 gap-3">
      {OPTIONS.map((opt) => {
        const isSelected = value === opt.id
        return (
          <button
            key={opt.id}
            type="button"
            onClick={() => onChange(opt.id)}
            className={cn(
              'p-3 rounded-xl border-2 text-left transition-all hover:scale-105',
              isSelected
                ? 'border-indigo-500 bg-indigo-50 shadow-lg shadow-indigo-100'
                : 'border-gray-200 bg-white hover:border-gray-300 shadow-sm'
            )}
          >
            <div className="font-mono text-[9px] leading-tight text-gray-500 mb-2 select-none">
              {opt.preview.map((line, i) => (
                <div key={i}>{line}</div>
              ))}
            </div>
            <div className={cn('text-xs font-semibold', isSelected ? 'text-indigo-700' : 'text-gray-700')}>
              {opt.label}
            </div>
            <div className="text-xs text-gray-400 mt-0.5">{opt.description}</div>
          </button>
        )
      })}
    </div>
  )
}
