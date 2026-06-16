import { cn } from '../../lib/utils'
import type { FontFamily, FontSize } from '../../types/wizard'
import { FONT_NAMES } from '../../types/theme'

interface Props {
  fontFamily: FontFamily
  fontSize: FontSize
  onFontFamilyChange: (f: FontFamily) => void
  onFontSizeChange: (s: FontSize) => void
}

const FAMILY_OPTIONS: { id: FontFamily; label: string; sample: string }[] = [
  { id: 'sans', label: 'Sans-Serif', sample: 'AaBbCc — Calibri' },
  { id: 'serif', label: 'Serif', sample: 'AaBbCc — Cambria' },
  { id: 'mono', label: 'Monospace', sample: 'AaBbCc — Courier' },
]

const SIZE_OPTIONS: { id: FontSize; label: string; hint: string }[] = [
  { id: 'small', label: 'Small', hint: '10pt data' },
  { id: 'medium', label: 'Medium', hint: '11pt data' },
  { id: 'large', label: 'Large', hint: '12pt data' },
]

export function FontPicker({ fontFamily, fontSize, onFontFamilyChange, onFontSizeChange }: Props) {
  return (
    <div className="space-y-5">
      <div>
        <div className="text-sm font-medium text-gray-700 mb-2">Font Family</div>
        <div className="grid grid-cols-3 gap-2">
          {FAMILY_OPTIONS.map((opt) => (
            <button
              key={opt.id}
              type="button"
              onClick={() => onFontFamilyChange(opt.id)}
              className={cn(
                'p-3 rounded-lg border-2 text-left transition-all',
                fontFamily === opt.id
                  ? 'border-indigo-500 bg-indigo-50'
                  : 'border-gray-200 bg-white hover:border-gray-300'
              )}
            >
              <div
                className="text-base mb-1"
                style={{ fontFamily: FONT_NAMES[opt.id] }}
              >
                {opt.sample.split(' — ')[0]}
              </div>
              <div className="text-xs text-gray-600">{opt.label}</div>
              <div className="text-xs text-gray-400">{opt.sample.split(' — ')[1]}</div>
            </button>
          ))}
        </div>
      </div>

      <div>
        <div className="text-sm font-medium text-gray-700 mb-2">Font Size</div>
        <div className="flex gap-2">
          {SIZE_OPTIONS.map((opt) => (
            <button
              key={opt.id}
              type="button"
              onClick={() => onFontSizeChange(opt.id)}
              className={cn(
                'flex-1 py-2.5 px-3 rounded-lg border-2 text-center transition-all',
                fontSize === opt.id
                  ? 'border-indigo-500 bg-indigo-50'
                  : 'border-gray-200 bg-white hover:border-gray-300'
              )}
            >
              <div className="font-medium text-sm">{opt.label}</div>
              <div className="text-xs text-gray-400 mt-0.5">{opt.hint}</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  )
}
