import type { Plugin, ViteDevServer } from 'vite'
import type { IncomingMessage, ServerResponse } from 'http'
import Anthropic from '@anthropic-ai/sdk'

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
]

interface RecommendationInput {
  destination: string
  duration: number
  partySize: number
  travelMonth: number
  tripStyle: 'budget' | 'midrange' | 'luxury'
}

// Module-level client — instantiated once, reused across all requests.
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY ?? '' })

// ---------------------------------------------------------------------------
// Prompt constants — static content is separated from dynamic content so that
// cache_control breakpoints can cover the maximum possible prefix.
// ---------------------------------------------------------------------------

const SYSTEM_PROMPT = `You are an expert travel research assistant. You use web search to find current, accurate information about destinations, including real hotels, real restaurants, real attractions/excursions, and real annual events/festivals.

When asked for trip recommendations, you research the destination thoroughly using web search, then respond with a single valid JSON object matching the schema provided. Output JSON only — no preamble, no markdown, no commentary. The response must start with "{" and end with "}".`

// All static structure: instructions, JSON schema, and constraints.
// Dynamic trip details (destination, duration, etc.) are appended separately.
const STATIC_PROMPT = `Research a real trip and return recommendations as JSON.

INSTRUCTIONS
1. Use web search to find current information about the destination.
2. Identify the major cities, regions, or tourist areas within the destination (e.g. for Belize: Cayo District, Ambergris Caye, Placencia). Aim for 3–5 of the most popular.
3. For EACH region, find real, currently-operating tier-appropriate hotels, restaurants, and excursions/activities located in or near that region. Use actual names — do not invent.
4. Build a day-by-day itinerary using REAL named attractions, restaurants, and neighborhoods in the destination.
5. List major annual events, festivals, and holidays that happen in the destination (max 20).

RESPONSE FORMAT (return EXACTLY this JSON structure, no other text):
{
  "regions": [
    {
      "region": "string (city / region / tourist area name, e.g. 'Cayo District')",
      "description": "string (one short sentence about what the area is known for — MAX ~120 characters)",
      "hotels": [
        {
          "name": "string (real hotel name)",
          "price": "string (typical nightly rate, e.g. '$120/night')",
          "rating": "string (stars, e.g. '★★★★' or '4.5/5')",
          "description": "string (one short sentence — what it's known for — MAX ~140 characters)"
        }
      ],
      "restaurants": [
        {
          "name": "string (real restaurant name)",
          "cuisine": "string (e.g. 'Belizean / Seafood')",
          "price": "string (price level, e.g. '$$' or '~$25/person')",
          "description": "string (one short sentence — signature dish or vibe — MAX ~140 characters)"
        }
      ],
      "excursions": [
        {
          "name": "string (real tour / attraction / activity name)",
          "duration": "string (e.g. 'Half day' or '2–3 hrs')",
          "price": "string (typical cost, e.g. '$75/person')",
          "description": "string (one short sentence — what you do — MAX ~140 characters)"
        }
      ]
    }
  ],
  "itinerary": [
    {
      "day": number (1, 2, 3, ...),
      "location": "string (city or neighborhood for that day)",
      "morning": "string (real activity / place — brief phrase, MAX ~160 characters)",
      "afternoon": "string (real activity / place — brief phrase, MAX ~160 characters)",
      "evening": "string (real activity / restaurant — brief phrase, MAX ~160 characters)",
      "transport": "string (how to get around that day — brief)",
      "notes": "string (tips, reservations needed, etc. — brief phrase, MAX ~160 characters)"
    }
  ],
  "events": [
    {
      "name": "string (real event/festival name)",
      "monthOrDate": "string (e.g. 'Late March' or 'July 14')",
      "type": "string (Festival / Holiday / Cultural / Religious / Sport / Music / Food)",
      "location": "string (where it happens)",
      "description": "string (one-sentence summary)"
    }
  ]
}

CONSTRAINTS
- regions: 3–5 items, each with 2–4 hotels, 2–4 restaurants, and 2–4 excursions
- Use real, currently-operating names. Do not invent.
- Keep ALL description / activity text concise — one short sentence or phrase each, no run-ons. Respect the character limits noted above so the text fits neatly in spreadsheet cells.
- Output JSON only. No markdown fences, no preamble.`

// ---------------------------------------------------------------------------
// Response-level cache — provides guaranteed savings on repeated lookups of
// the same destination+params within a rolling TTL window. This is the most
// common pattern during development and demo generation.
// ---------------------------------------------------------------------------

interface CacheEntry { data: unknown; expiresAt: number }
const responseCache = new Map<string, CacheEntry>()
const CACHE_TTL_MS = 60 * 60 * 1000 // 1 hour

function getCacheKey(input: RecommendationInput): string {
  return JSON.stringify([
    input.destination.trim().toLowerCase(),
    input.duration,
    input.partySize,
    input.travelMonth,
    input.tripStyle,
  ])
}

// ---------------------------------------------------------------------------
// Prompt builders
// ---------------------------------------------------------------------------

function buildDynamicTripDetails(input: RecommendationInput): string {
  const monthName = MONTH_NAMES[input.travelMonth - 1] ?? 'this year'
  const tierLabel =
    input.tripStyle === 'budget'
      ? 'Budget Backpacker (hostels, budget airlines, street food, free attractions)'
      : input.tripStyle === 'midrange'
        ? 'Mid-Range (3-star hotels, economy flights, local restaurants, paid attractions)'
        : 'Luxury (5-star hotels, business/first class flights, fine dining, private tours)'

  return `TRIP DETAILS
- Destination: ${input.destination}
- Duration: ${input.duration} days
- Travel month: ${monthName}
- Travelers: ${input.partySize}
- Tier: ${tierLabel}

Apply the instructions above to this specific trip. The itinerary must have EXACTLY ${input.duration} day objects (day 1 through day ${input.duration}). Prices must match the ${input.tripStyle} tier.`
}

// ---------------------------------------------------------------------------
// Core generation
// ---------------------------------------------------------------------------

function extractJson(rawText: string): string {
  let trimmed = rawText.trim()
  if (trimmed.startsWith('```')) {
    trimmed = trimmed.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '')
  }
  const start = trimmed.indexOf('{')
  if (start === -1) throw new Error('No JSON object found in response')
  let depth = 0
  let inString = false
  let escape = false
  for (let i = start; i < trimmed.length; i++) {
    const ch = trimmed[i]
    if (escape) { escape = false; continue }
    if (ch === '\\') { escape = true; continue }
    if (ch === '"') { inString = !inString; continue }
    if (inString) continue
    if (ch === '{') depth++
    else if (ch === '}') {
      depth--
      if (depth === 0) return trimmed.slice(start, i + 1)
    }
  }
  throw new Error('Incomplete JSON in response')
}

async function generateRecommendations(input: RecommendationInput) {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error(
      'ANTHROPIC_API_KEY is not set. Copy .env.example to .env and add your key.'
    )
  }

  // Check response-level cache first.
  const cacheKey = getCacheKey(input)
  const cached = responseCache.get(cacheKey)
  if (cached && cached.expiresAt > Date.now()) {
    console.log('[recommendations] response cache hit for:', input.destination)
    return cached.data
  }

  const message = await client.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 16000,
    // cache_control on system covers the tools + system prefix together.
    system: [
      {
        type: 'text',
        text: SYSTEM_PROMPT,
        cache_control: { type: 'ephemeral' },
      },
    ],
    tools: [
      {
        type: 'web_search_20250305',
        name: 'web_search',
        max_uses: 5,
      },
    ],
    messages: [
      {
        role: 'user',
        content: [
          // Static block — cache_control breakpoint lets this prefix be served
          // from cache on subsequent calls. Sonnet 4.6 requires ≥2048 tokens
          // for a cache hit; check `cacheWrite`/`cacheRead` in the usage log
          // below to confirm threshold is being reached.
          {
            type: 'text',
            text: STATIC_PROMPT,
            cache_control: { type: 'ephemeral' },
          },
          // Dynamic block — unique per request, always sent fresh.
          {
            type: 'text',
            text: buildDynamicTripDetails(input),
          },
        ],
      },
    ],
  })

  // Log token usage so cache effectiveness is visible in the Vite console.
  console.log('[recommendations] usage:', {
    input: message.usage.input_tokens,
    output: message.usage.output_tokens,
    cacheWrite: (message.usage as Record<string, number>).cache_creation_input_tokens ?? 0,
    cacheRead: (message.usage as Record<string, number>).cache_read_input_tokens ?? 0,
  })

  const textBlocks = message.content
    .filter((block: { type: string }) => block.type === 'text')
    .map((block: { type: string; text?: string }) => block.text ?? '')
    .join('\n')

  if (!textBlocks.trim()) {
    throw new Error('Empty response from Claude. Try again.')
  }

  const jsonStr = extractJson(textBlocks)
  const result = JSON.parse(jsonStr)

  // Populate response cache.
  responseCache.set(cacheKey, { data: result, expiresAt: Date.now() + CACHE_TTL_MS })

  return result
}

// ---------------------------------------------------------------------------
// HTTP body reader
// ---------------------------------------------------------------------------

function readBody(req: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    req.on('data', (chunk) => chunks.push(chunk))
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    req.on('error', reject)
  })
}

// ---------------------------------------------------------------------------
// Vite plugin
// ---------------------------------------------------------------------------

export function recommendationsPlugin(): Plugin {
  return {
    name: 'travel-recommendations',
    configureServer(server: ViteDevServer) {
      server.middlewares.use('/api/recommendations', async (req: IncomingMessage, res: ServerResponse) => {
        if (req.method !== 'POST') {
          res.statusCode = 405
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'Method not allowed' }))
          return
        }
        try {
          const rawBody = await readBody(req)
          const body = JSON.parse(rawBody) as RecommendationInput
          if (!body.destination?.trim()) {
            res.statusCode = 400
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'destination is required' }))
            return
          }
          const result = await generateRecommendations(body)
          res.statusCode = 200
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify(result))
        } catch (err) {
          const message = err instanceof Error ? err.message : 'Internal error'
          console.error('[recommendations] error:', err)
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: message }))
        }
      })
    },
  }
}
