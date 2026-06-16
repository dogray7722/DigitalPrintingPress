import type { TripStyle } from '../../types/wizard'

export interface RecommendationRequest {
  destination: string
  duration: number
  partySize: number
  travelMonth: number
  tripStyle: TripStyle
}

// A single recommended place. Fields are deliberately loose strings so the AI can
// return human-friendly values ("$120/night", "★★★★½", "2–3 hrs") that drop straight
// into an informational table cell. Different tabs surface different attributes:
//   hotels      → price + rating (stars)
//   restaurants → cuisine + price
//   excursions  → duration + price
export interface PlaceRec {
  name: string
  price?: string
  rating?: string
  cuisine?: string
  duration?: string
  description?: string
}

// A region / city / tourist area grouping its recommended places. Rendered as an
// informational "insert" table on the HOTELS, RESTAURANTS, and EXCURSIONS tabs.
export interface RegionRec {
  region: string
  description?: string
  hotels: PlaceRec[]
  restaurants: PlaceRec[]
  excursions: PlaceRec[]
}

export interface ItineraryDay {
  day: number
  location?: string
  morning?: string
  afternoon?: string
  evening?: string
  transport?: string
  notes?: string
}

export interface EventRec {
  name: string
  monthOrDate?: string
  type?: string
  location?: string
  description?: string
}

export interface Recommendations {
  regions: RegionRec[]
  itinerary: ItineraryDay[]
  events: EventRec[]
}
