import { useState, useEffect } from 'react'
import { DollarSign, ArrowLeftRight, Plane, Building2, Utensils, Compass, Bus, ShoppingBag, Wallet } from 'lucide-react'
import { useWizardStore } from '../../store/wizardStore'
import { CurrencySelector } from '../shared/CurrencySelector'
import { TRIP_STYLE_PRESETS, BUDGET_CATEGORIES } from '../../data/tripStylePresets'
import { computeTotalBudget } from '../../lib/utils'
import { getCurrencySymbol } from '../../data/currencies'
import type { TripStyle } from '../../types/wizard'

function BudgetInput({ value, onChange, sym }: { value: number; onChange: (v: number) => void; sym: string }) {
  const [localVal, setLocalVal] = useState(String(value))
  const [focused, setFocused] = useState(false)

  useEffect(() => {
    if (!focused) setLocalVal(String(value))
  }, [value, focused])

  return (
    <div className="relative">
      <span className="absolute left-3 top-1/2 -translate-y-1/2 text-sm text-gray-400 font-mono pointer-events-none">
        {sym}
      </span>
      <input
        type="text"
        inputMode="decimal"
        value={localVal}
        onChange={(e) => setLocalVal(e.target.value)}
        onFocus={(e) => { setFocused(true); e.target.select() }}
        onBlur={() => {
          setFocused(false)
          const parsed = Math.max(0, parseFloat(localVal) || 0)
          setLocalVal(String(parsed))
          onChange(parsed)
        }}
        className="w-full pl-8 pr-3 py-2.5 rounded-xl border border-gray-200 text-sm font-medium focus:outline-none focus:ring-2 focus:ring-indigo-400"
      />
    </div>
  )
}

const CATEGORY_ICONS = {
  flights: Plane,
  hotel: Building2,
  food: Utensils,
  activities: Compass,
  transport: Bus,
  shopping: ShoppingBag,
  misc: Wallet,
}

export function Step2Budget() {
  const { currency, destinationCurrency, tripStyle, budgets, duration, partySize, setField } = useWizardStore()

  const totalBudget = computeTotalBudget(budgets, duration)
  const sym = getCurrencySymbol(currency)

  function applyPreset(style: TripStyle) {
    setField('tripStyle', style)
    setField('budgets', TRIP_STYLE_PRESETS[style].budgets)
  }

  return (
    <div className="space-y-7">
      {/* Currency pair */}
      <div className="space-y-3">
        <div>
          <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-2">
            <DollarSign size={15} className="text-indigo-500" />
            Budget Currency
          </label>
          <CurrencySelector
            value={currency}
            onChange={(code) => setField('currency', code)}
          />
        </div>
        <div>
          <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-1">
            <ArrowLeftRight size={15} className="text-indigo-500" />
            Destination Currency
            <span className="ml-auto text-xs font-normal text-indigo-400 bg-indigo-50 px-2 py-0.5 rounded-full">AI only</span>
          </label>
          <p className="text-xs text-gray-400 mb-2">
            The local currency at your destination — adds an exchange rate widget to your spreadsheet when AI is enabled.
          </p>
          <CurrencySelector
            value={destinationCurrency || currency}
            onChange={(code) => setField('destinationCurrency', code)}
          />
        </div>
      </div>

      {/* Trip style presets */}
      <div>
        <div className="text-sm font-semibold text-gray-700 mb-2">Trip Style Preset</div>
        <div className="grid grid-cols-3 gap-3">
          {(Object.entries(TRIP_STYLE_PRESETS) as [TripStyle, typeof TRIP_STYLE_PRESETS[TripStyle]][]).map(([id, preset]) => (
            <button
              key={id}
              type="button"
              onClick={() => applyPreset(id)}
              className={`flex flex-col items-center gap-2 p-4 rounded-xl border-2 text-center transition-all ${
                tripStyle === id
                  ? 'border-indigo-500 bg-indigo-50'
                  : 'border-gray-200 bg-white hover:border-gray-300'
              }`}
            >
              <span className="text-2xl">{preset.icon}</span>
              <div className={`text-xs font-semibold ${tripStyle === id ? 'text-indigo-700' : 'text-gray-700'}`}>
                {preset.label}
              </div>
              <div className="text-xs text-gray-400 leading-tight">{preset.description}</div>
            </button>
          ))}
        </div>
        <p className="mt-2 text-xs text-gray-400">Selecting a preset fills the fields below — you can edit them after.</p>
      </div>

      {/* Budget inputs */}
      <div>
        <div className="text-sm font-semibold text-gray-700 mb-3">Budget Amounts</div>
        <div className="grid grid-cols-2 gap-3">
          {BUDGET_CATEGORIES.map(({ key, label, description }) => {
            const Icon = CATEGORY_ICONS[key]
            const isPerTrip = key === 'flights' || key === 'shopping'
            return (
              <div key={key} className="flex flex-col gap-1">
                <label className="flex items-center gap-1.5 text-xs font-medium text-gray-600">
                  <Icon size={12} className="text-indigo-400" />
                  {label}
                  <span className="text-gray-400 font-normal ml-auto">
                    {isPerTrip ? 'per trip' : 'per day'}
                  </span>
                </label>
                <BudgetInput
                  sym={sym}
                  value={budgets[key]}
                  onChange={(v) => setField('budgets', { ...budgets, [key]: v })}
                />
                <div className="text-xs text-gray-400">{description}</div>
              </div>
            )
          })}
        </div>
      </div>

      {/* Total summary */}
      <div className="flex items-center justify-between bg-gradient-to-r from-indigo-50 to-purple-50 rounded-xl p-4 border border-indigo-100">
        <div>
          <div className="text-xs text-gray-500 mb-0.5">Estimated Total Budget</div>
          <div className="text-2xl font-bold text-indigo-700">
            {sym}{totalBudget.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
          </div>
          <div className="text-xs text-gray-500 mt-0.5">
            {duration} days · {partySize} traveler{partySize !== 1 ? 's' : ''} · {currency}
          </div>
        </div>
        <div className="text-4xl opacity-20">💰</div>
      </div>
    </div>
  )
}
