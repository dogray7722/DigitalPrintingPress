export type SheetId =
  | 'budgetTracker'
  | 'itinerary'
  | 'hotels'
  | 'restaurants'
  | 'excursions'
  | 'packingList'
  | 'flights'
  | 'tasks'
  | 'events'

export type PartyType = 'solo' | 'couple' | 'family' | 'group'
export type TripStyle = 'budget' | 'midrange' | 'luxury'
export type FontFamily = 'sans' | 'serif' | 'mono'
export type FontSize = 'small' | 'medium' | 'large'
export type ThemeId = 'sakura' | 'ocean' | 'forest' | 'desert' | 'inkwell' | 'parchment'

export interface BudgetAmounts {
  flights: number
  hotel: number
  food: number
  activities: number
  transport: number
  shopping: number
  misc: number
}

export interface OverviewImage {
  /** Final baked crop, ready for Excel embed — 1400x756 JPEG data URL. */
  dataUrl: string
  /** Original uploaded image (untouched), used to recompute the crop on reposition. */
  sourceDataUrl: string
  sourceWidth: number
  sourceHeight: number
  /** 0-1 vertical focal point within the available crop slack; 0.5 = center (default). 0 = top of source visible, 1 = bottom of source visible. Only meaningful when the source is taller than the ~1.85:1 target ratio. */
  offsetY: number
}

export const ALL_SHEET_IDS: SheetId[] = [
  'itinerary',
  'flights',
  'hotels',
  'restaurants',
  'excursions',
  'budgetTracker',
  'packingList',
  'tasks',
  'events',
]

export const SHEET_LABELS: Record<SheetId, string> = {
  budgetTracker: 'Other Expenses',
  itinerary: 'Itinerary',
  hotels: 'Accommodation',
  restaurants: 'Dining',
  excursions: 'Excursions',
  packingList: 'Packing List',
  flights: 'Transportation',
  tasks: 'Tasks',
  events: 'Annual Events',
}

// Emoji icon per sheet — shared by the Step 3 toggle grid and the OVERVIEW JUMP TO
// strip, so the wizard and the generated workbook can't drift apart. Same family as
// the 📍 pin on the AI recommendation guide tables: plain Unicode emoji, rendered by
// the OS emoji font (no glyph exists in Calibri/Cambria/Courier New).
export const SHEET_ICONS: Record<SheetId, string> = {
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

export interface WizardState {
  // Step 1
  destination: string
  useRecommendations: boolean
  startDate: string
  duration: number
  partyType: PartyType
  partySize: number
  travelMonth: number
  overviewImage: OverviewImage | null

  // Step 2
  currency: string
  destinationCurrency: string
  tripStyle: TripStyle
  budgets: BudgetAmounts

  // Step 3
  theme: ThemeId
  accentColor: string
  fontFamily: FontFamily
  fontSize: FontSize
  sheets: Record<SheetId, boolean>

  currentStep: 1 | 2 | 3 | 4
}

export const DEFAULT_STATE: WizardState = {
  destination: '',
  useRecommendations: false,
  startDate: '',
  duration: 7,
  partyType: 'couple',
  partySize: 2,
  travelMonth: new Date().getMonth() + 1,
  overviewImage: null,

  currency: 'USD',
  destinationCurrency: '',
  tripStyle: 'midrange',
  // Defaults bumped ~30% over prior estimates to reflect current (inflationary) prices.
  budgets: {
    flights: 1040,
    hotel: 195,
    food: 78,
    activities: 65,
    transport: 39,
    shopping: 200,
    misc: 52,
  },

  theme: 'sakura',
  accentColor: '',
  fontFamily: 'sans',
  fontSize: 'medium',
  sheets: {
    budgetTracker: true,
    itinerary: true,
    hotels: true,
    restaurants: true,
    excursions: true,
    packingList: true,
    flights: true,
    tasks: true,
    events: true,
  },

  currentStep: 1,
}
