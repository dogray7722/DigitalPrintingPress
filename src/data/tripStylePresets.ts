import type { TripStyle, BudgetAmounts } from '../types/wizard'

// hotel / food / activities / transport / misc are per-night/per-day rates;
// flights and shopping are per-trip totals. Default estimates were bumped ~30%
// over prior figures to reflect current (inflationary) prices.
export interface TripStylePreset {
  label: string
  description: string
  icon: string
  budgets: BudgetAmounts
}

export const TRIP_STYLE_PRESETS: Record<TripStyle, TripStylePreset> = {
  budget: {
    label: 'Budget Backpacker',
    description: 'Hostels, street food, free attractions, local transport',
    icon: '🎒',
    budgets: {
      flights: 1000,
      hotel: 52,
      food: 33,
      activities: 26,
      transport: 13,
      shopping: 75,
      misc: 20,
    },
  },
  midrange: {
    label: 'Mid-Range Traveler',
    description: '3-star hotels, local restaurants, mix of paid attractions',
    icon: '✈️',
    budgets: {
      flights: 1040,
      hotel: 156,
      food: 78,
      activities: 65,
      transport: 33,
      shopping: 200,
      misc: 46,
    },
  },
  luxury: {
    label: 'Luxury Explorer',
    description: '5-star hotels, fine dining, private tours, business class',
    icon: '💎',
    budgets: {
      flights: 3250,
      hotel: 585,
      food: 260,
      activities: 234,
      transport: 104,
      shopping: 500,
      misc: 156,
    },
  },
}

export const BUDGET_CATEGORIES = [
  { key: 'flights' as const, label: 'Flights', description: 'Total round-trip airfare' },
  { key: 'hotel' as const, label: 'Hotel / Night', description: 'Accommodation per night' },
  { key: 'food' as const, label: 'Food / Day', description: 'Daily food & dining budget' },
  { key: 'activities' as const, label: 'Activities / Day', description: 'Daily sightseeing & activities' },
  { key: 'transport' as const, label: 'Transport / Day', description: 'Daily local transport' },
  { key: 'shopping' as const, label: 'Shopping', description: 'Total shopping & souvenirs' },
  { key: 'misc' as const, label: 'Misc / Day', description: 'Daily miscellaneous expenses' },
]
