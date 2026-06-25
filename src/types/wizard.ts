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
export type ChartStyle = 'bar' | 'pie' | 'donut'
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
  /** data URL, e.g. "data:image/jpeg;base64,...." — already center-cropped/resized to ~1400x648 JPEG client-side. */
  dataUrl: string
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
  chartStyle: ChartStyle
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
  chartStyle: 'bar',
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
