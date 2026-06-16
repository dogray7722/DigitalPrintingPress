import type { Recommendations, RecommendationRequest, RegionRec, PlaceRec } from './types'

function normalizePlaces(raw: unknown): PlaceRec[] {
  if (!Array.isArray(raw)) return []
  return raw
    .filter((p): p is PlaceRec => !!p && typeof (p as PlaceRec).name === 'string')
    .map((p) => ({
      name: p.name,
      price: p.price,
      rating: p.rating,
      cuisine: p.cuisine,
      duration: p.duration,
      description: p.description,
    }))
}

function normalizeRegions(raw: unknown): RegionRec[] {
  if (!Array.isArray(raw)) return []
  return raw
    .filter((r): r is RegionRec => !!r && typeof (r as RegionRec).region === 'string')
    .map((r) => ({
      region: r.region,
      description: r.description,
      hotels: normalizePlaces(r.hotels),
      restaurants: normalizePlaces(r.restaurants),
      excursions: normalizePlaces(r.excursions),
    }))
}

export async function fetchRecommendations(
  input: RecommendationRequest
): Promise<Recommendations> {
  const resp = await fetch('/api/recommendations', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(input),
  })

  if (!resp.ok) {
    const data = await resp.json().catch(() => ({ error: resp.statusText }))
    throw new Error(data.error || `Request failed (${resp.status})`)
  }

  const data = (await resp.json()) as Partial<Recommendations>

  return {
    regions: normalizeRegions(data.regions),
    itinerary: Array.isArray(data.itinerary) ? data.itinerary : [],
    events: Array.isArray(data.events) ? data.events : [],
  }
}
