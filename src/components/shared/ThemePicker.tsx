import { cn } from '../../lib/utils'
import type { ThemeId } from '../../types/wizard'
import { THEMES, THEME_LABELS } from '../../types/theme'
import { Check } from 'lucide-react'

interface Props {
  value: ThemeId
  onChange: (id: ThemeId) => void
}

const THEME_IDS = Object.keys(THEMES) as ThemeId[]

function hexFromArgb(argb: string): string {
  return '#' + argb.slice(2)
}

export function ThemePicker({ value, onChange }: Props) {
  return (
    <div className="grid grid-cols-3 gap-3">
      {THEME_IDS.map((id) => {
        const palette = THEMES[id]
        const isSelected = value === id
        return (
          <button
            key={id}
            type="button"
            onClick={() => onChange(id)}
            className={cn(
              'relative rounded-xl overflow-hidden border-2 transition-all hover:scale-105',
              isSelected ? 'border-indigo-500 shadow-lg shadow-indigo-200' : 'border-transparent shadow-sm hover:border-gray-300'
            )}
          >
            {/* Mini spreadsheet preview */}
            <div className="flex flex-col">
              {/* Header bar */}
              <div
                className="h-8 flex items-center justify-center text-xs font-bold px-2"
                style={{
                  backgroundColor: hexFromArgb(palette.primary),
                  color: hexFromArgb(palette.primaryText),
                }}
              >
                DESTINATION
              </div>
              {/* Sub-header */}
              <div
                className="h-4"
                style={{ backgroundColor: hexFromArgb(palette.secondary) }}
              />
              {/* Data rows preview */}
              {[palette.lightBg, 'FFFFFFFF', palette.lightBg].map((bg, i) => (
                <div
                  key={i}
                  className="h-3 flex items-center gap-1 px-2"
                  style={{ backgroundColor: hexFromArgb(bg) }}
                >
                  <div
                    className="h-1.5 rounded-full"
                    style={{
                      width: `${35 + i * 10}%`,
                      backgroundColor: hexFromArgb(palette.accent),
                    }}
                  />
                </div>
              ))}
              {/* Label */}
              <div
                className="py-1.5 text-center text-xs font-medium"
                style={{ backgroundColor: hexFromArgb(palette.mediumBg) }}
              >
                {THEME_LABELS[id]}
              </div>
            </div>
            {isSelected && (
              <div className="absolute top-1.5 right-1.5 w-5 h-5 bg-indigo-500 rounded-full flex items-center justify-center">
                <Check size={11} className="text-white" strokeWidth={3} />
              </div>
            )}
          </button>
        )
      })}
    </div>
  )
}
