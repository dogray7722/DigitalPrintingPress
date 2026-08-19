import { Palette, Type, Layers } from 'lucide-react'
import { useWizardStore } from '../../store/wizardStore'
import { ThemePicker } from '../shared/ThemePicker'
import { FontPicker } from '../shared/FontPicker'
import { SheetToggleGrid } from '../shared/SheetToggleGrid'
import type { ThemeId, FontFamily, FontSize, SheetId } from '../../types/wizard'

export function Step3Style() {
  const {
    theme, fontFamily, fontSize, sheets, useRecommendations,
    setField, setSheet,
  } = useWizardStore()

  return (
    <div className="space-y-8">
      {/* Theme */}
      <div>
        <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-3">
          <Palette size={15} className="text-indigo-500" />
          Color Theme
        </label>
        <ThemePicker
          value={theme}
          onChange={(id: ThemeId) => setField('theme', id)}
        />
      </div>

      {/* Font */}
      <div>
        <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-3">
          <Type size={15} className="text-indigo-500" />
          Typography
        </label>
        <FontPicker
          fontFamily={fontFamily}
          fontSize={fontSize}
          onFontFamilyChange={(f: FontFamily) => setField('fontFamily', f)}
          onFontSizeChange={(s: FontSize) => setField('fontSize', s)}
        />
      </div>

      {/* Sheets */}
      <div>
        <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-3">
          <Layers size={15} className="text-indigo-500" />
          Included Sheets
        </label>
        <SheetToggleGrid
          sheets={sheets}
          useRecommendations={useRecommendations}
          onChange={(id: SheetId, enabled: boolean) => setSheet(id, enabled)}
        />
      </div>
    </div>
  )
}
