import type { Plugin, ViteDevServer } from "vite";
import type { IncomingMessage, ServerResponse } from "http";
import Anthropic from "@anthropic-ai/sdk";
import path from "path";
import fs from "fs";

const MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];

interface RecommendationInput {
  destination: string;
  duration: number;
  partySize: number;
  travelMonth: number;
  tripStyle: "budget" | "midrange" | "luxury";
}

// Lazy singleton — deferred until first request so that loadEnv() in vite.config.ts
// has already populated process.env before the SDK reads ANTHROPIC_API_KEY.
let _client: Anthropic | null = null;
function getClient(): Anthropic {
  if (!_client) _client = new Anthropic();
  return _client;
}

// ---------------------------------------------------------------------------
// Prompt constants — static content is separated from dynamic content so that
// cache_control breakpoints can cover the maximum possible prefix.
// ---------------------------------------------------------------------------

const SYSTEM_PROMPT = `You are an expert travel research assistant. You use web search to find current, accurate information about destinations, including real hotels, real restaurants, real attractions/excursions, and real annual events/festivals.

When asked for trip recommendations, you research the destination thoroughly using web search, then respond with a single valid JSON object matching the schema provided. Output JSON only — no preamble, no markdown, no commentary. The response must start with "{" and end with "}".`;

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
- Output JSON only. No markdown fences, no preamble.`;

// ---------------------------------------------------------------------------
// Response-level cache — provides guaranteed savings on repeated lookups of
// the same destination+params within a rolling TTL window. This is the most
// common pattern during development and demo generation.
// ---------------------------------------------------------------------------

interface CacheEntry {
  data: unknown;
  expiresAt: number;
}
const responseCache = new Map<string, CacheEntry>();
// Recommendations are a list of real hotels/restaurants/excursions — that doesn't rot
// on a 30-day timescale, and every miss buys a Sonnet call plus 5 web searches.
const CACHE_TTL_MS = 183 * 24 * 60 * 60 * 1000; // ~6 months

const CACHE_FILE_PATH = path.resolve(
  process.cwd(),
  ".recommendations-cache.json"
);

// ---------------------------------------------------------------------------
// Shared disk-cache plumbing (used by both the response and exchange-rate caches)
// ---------------------------------------------------------------------------

function readCacheFile<T>(filePath: string): Map<string, T> {
  try {
    const raw = fs.readFileSync(filePath, "utf8");
    return new Map(JSON.parse(raw) as Array<[string, T]>);
  } catch {
    // File doesn't exist yet or is corrupt — start empty.
    return new Map();
  }
}

// Merge-on-write. Each process loads the file once at import and would otherwise
// overwrite it wholesale from its own snapshot, so anything a *different* process
// (a second dev server, an overlapping restart) wrote in the meantime was silently
// erased. Re-read first, merge, then write — and adopt the merged view so this
// process picks up the other one's entries too.
function persistCache<T extends { expiresAt: number }>(
  filePath: string,
  memory: Map<string, T>
): void {
  try {
    const merged = readCacheFile<T>(filePath);
    for (const [key, entry] of memory) {
      const existing = merged.get(key);
      // Same key on both sides: keep whichever copy was fetched most recently.
      if (!existing || entry.expiresAt >= existing.expiresAt)
        merged.set(key, entry);
    }
    for (const [key, entry] of merged) memory.set(key, entry);

    // Write-then-rename so a crash mid-write can't leave a truncated file behind.
    const tmpPath = `${filePath}.${process.pid}.tmp`;
    fs.writeFileSync(tmpPath, JSON.stringify([...merged]), "utf8");
    fs.renameSync(tmpPath, filePath);
  } catch {
    // Non-fatal — cache will still work in-memory this session.
  }
}

// Expired entries are deliberately kept, both here and on disk: past its TTL an entry
// stops being served directly, but it stays available as a fallback if the refresh
// call fails. Dropping them meant destroying still-usable data and re-paying for it.
function loadCacheFromDisk(): void {
  for (const [key, entry] of readCacheFile<CacheEntry>(CACHE_FILE_PATH)) {
    responseCache.set(key, entry);
  }
  const now = Date.now();
  const fresh = [...responseCache.values()].filter(
    (e) => e.expiresAt > now
  ).length;
  console.log(
    `[recommendations] loaded ${responseCache.size} cached destination(s) from disk (${fresh} fresh, ${responseCache.size - fresh} stale)`
  );
}

function saveCacheToDisk(): void {
  persistCache(CACHE_FILE_PATH, responseCache);
}

loadCacheFromDisk();

// ---------------------------------------------------------------------------
// Exchange rate cache — same 30-day TTL, keyed by "FROM-TO" pair
// ---------------------------------------------------------------------------

interface ExchangeRateEntry {
  from: string;
  to: string;
  rate: number;
  fetchedAt: string;
}
interface ExchangeRateCacheEntry {
  data: ExchangeRateEntry;
  expiresAt: number;
}
const exchangeRateCache = new Map<string, ExchangeRateCacheEntry>();
// Rates get their own TTL — deliberately left at 30 days, not raised alongside the
// recommendations cache above.
const EXCHANGE_RATE_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 days
const EXCHANGE_RATE_CACHE_FILE = path.resolve(
  process.cwd(),
  ".exchange-rate-cache.json"
);

function loadExchangeRateCache(): void {
  for (const [key, entry] of readCacheFile<ExchangeRateCacheEntry>(
    EXCHANGE_RATE_CACHE_FILE
  )) {
    exchangeRateCache.set(key, entry);
  }
  const now = Date.now();
  const fresh = [...exchangeRateCache.values()].filter(
    (e) => e.expiresAt > now
  ).length;
  console.log(
    `[exchange-rate] loaded ${exchangeRateCache.size} cached pair(s) from disk (${fresh} fresh, ${exchangeRateCache.size - fresh} stale)`
  );
}

function saveExchangeRateCache(): void {
  persistCache(EXCHANGE_RATE_CACHE_FILE, exchangeRateCache);
}

loadExchangeRateCache();

function formatFetchedAt(date: Date): string {
  const months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"
  ];
  return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

async function generateExchangeRate(
  from: string,
  to: string
): Promise<ExchangeRateEntry> {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error("ANTHROPIC_API_KEY is not set.");
  }

  const cacheKey = `${from}-${to}`;
  const cached = exchangeRateCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) {
    console.log(`[exchange-rate] cache hit for ${cacheKey}`);
    return cached.data;
  }
  if (cached) {
    console.log(`[exchange-rate] ${cacheKey} is past its TTL — refreshing`);
  }

  try {
    return await fetchExchangeRateFromApi(from, to, cacheKey);
  } catch (err) {
    // Stale beats nothing: the widget renders the rate with its own `fetchedAt`
    // date, so an old value is visibly old rather than silently wrong.
    if (cached) {
      console.warn(
        `[exchange-rate] refresh for ${cacheKey} failed — serving stale rate from ${cached.data.fetchedAt}:`,
        err instanceof Error ? err.message : err
      );
      return cached.data;
    }
    throw err;
  }
}

async function fetchExchangeRateFromApi(
  from: string,
  to: string,
  cacheKey: string
): Promise<ExchangeRateEntry> {
  const message = await getClient().messages.create({
    model: "claude-haiku-4-5-20251001",
    max_tokens: 256,
    tools: [
      { type: "web_search_20250305" as const, name: "web_search", max_uses: 2 }
    ],
    messages: [
      {
        role: "user",
        content: `Search for the current ${from} to ${to} exchange rate today. Return ONLY a JSON object: {"rate": <number>}. The rate should be how many ${to} equal 1 ${from}. No other text.`
      }
    ]
  });

  const textBlocks = message.content
    .filter((b: { type: string }) => b.type === "text")
    .map((b: { type: string; text?: string }) => b.text ?? "")
    .join("\n");

  let rate: number | null = null;
  try {
    const jsonStr = extractJson(textBlocks);
    const parsed = JSON.parse(jsonStr) as { rate?: unknown };
    if (typeof parsed.rate === "number" && parsed.rate > 0) rate = parsed.rate;
  } catch {
    const match = textBlocks.match(/\d[\d,]*\.?\d*/g);
    if (match) {
      const num = parseFloat(match[0].replace(/,/g, ""));
      if (num > 0) rate = num;
    }
  }

  if (rate === null)
    throw new Error(`Could not parse exchange rate for ${from}→${to}`);

  console.log(`[exchange-rate] ${from}→${to} = ${rate}`);

  const result: ExchangeRateEntry = {
    from,
    to,
    rate,
    fetchedAt: formatFetchedAt(new Date())
  };
  exchangeRateCache.set(cacheKey, {
    data: result,
    expiresAt: Date.now() + EXCHANGE_RATE_TTL_MS
  });
  saveExchangeRateCache();

  return result;
}

// Case, accents, punctuation and internal whitespace are folded away so "Paris, France",
// "paris france" and "Paris  France" share one entry instead of buying three API calls.
// Deliberately NOT fuzzy beyond that — "Paris" and "Paris, France" stay distinct, because
// anything looser starts conflating real places (Georgia the country vs. the state).
// Falls back to the raw lowercased string when normalization strips everything (a
// destination typed in a non-Latin script), so those can't all collide on "".
function getCacheKey(input: RecommendationInput): string {
  const normalized = input.destination
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
  return normalized || input.destination.trim().toLowerCase();
}

// ---------------------------------------------------------------------------
// Prompt builders
// ---------------------------------------------------------------------------

function buildDynamicTripDetails(input: RecommendationInput): string {
  const monthName = MONTH_NAMES[input.travelMonth - 1] ?? "this year";
  const tierLabel =
    input.tripStyle === "budget"
      ? "Budget Backpacker (hostels, budget airlines, street food, free attractions)"
      : input.tripStyle === "midrange"
        ? "Mid-Range (3-star hotels, economy flights, local restaurants, paid attractions)"
        : "Luxury (5-star hotels, business/first class flights, fine dining, private tours)";

  return `TRIP DETAILS
- Destination: ${input.destination}
- Duration: ${input.duration} days
- Travel month: ${monthName}
- Travelers: ${input.partySize}
- Tier: ${tierLabel}

Apply the instructions above to this specific trip. The itinerary must have EXACTLY ${input.duration} day objects (day 1 through day ${input.duration}). Prices must match the ${input.tripStyle} tier.`;
}

// ---------------------------------------------------------------------------
// Core generation
// ---------------------------------------------------------------------------

function extractJson(rawText: string): string {
  let trimmed = rawText.trim();
  if (trimmed.startsWith("```")) {
    trimmed = trimmed.replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/, "");
  }
  const start = trimmed.indexOf("{");
  if (start === -1) throw new Error("No JSON object found in response");
  let depth = 0;
  let inString = false;
  let escape = false;
  for (let i = start; i < trimmed.length; i++) {
    const ch = trimmed[i];
    if (escape) {
      escape = false;
      continue;
    }
    if (ch === "\\") {
      escape = true;
      continue;
    }
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) return trimmed.slice(start, i + 1);
    }
  }
  throw new Error("Incomplete JSON in response");
}

async function generateRecommendations(input: RecommendationInput) {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error(
      "ANTHROPIC_API_KEY is not set. Copy .env.example to .env and add your key."
    );
  }

  // Check response-level cache first.
  const cacheKey = getCacheKey(input);
  const cached = responseCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) {
    console.log("[recommendations] response cache hit for:", input.destination);
    return cached.data;
  }
  if (cached) {
    console.log(
      `[recommendations] cached entry for "${input.destination}" is past its 6-month TTL — refreshing`
    );
  }

  try {
    return await fetchRecommendationsFromApi(input, cacheKey);
  } catch (err) {
    // A stale guide is far better than a failed generation — the workbook still
    // gets real hotels/restaurants, and we don't burn a second call retrying.
    if (cached) {
      console.warn(
        `[recommendations] refresh for "${input.destination}" failed — serving stale cached data:`,
        err instanceof Error ? err.message : err
      );
      return cached.data;
    }
    throw err;
  }
}

async function fetchRecommendationsFromApi(
  input: RecommendationInput,
  cacheKey: string
) {
  const message = await getClient().messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 16000,
    // cache_control on system covers the tools + system prefix together.
    system: [
      {
        type: "text",
        text: SYSTEM_PROMPT,
        cache_control: { type: "ephemeral" }
      }
    ],
    tools: [
      {
        type: "web_search_20250305",
        name: "web_search",
        max_uses: 5
      }
    ],
    messages: [
      {
        role: "user",
        content: [
          // Static block — cache_control breakpoint lets this prefix be served
          // from cache on subsequent calls. Sonnet 4.6 requires ≥2048 tokens
          // for a cache hit; check `cacheWrite`/`cacheRead` in the usage log
          // below to confirm threshold is being reached.
          {
            type: "text",
            text: STATIC_PROMPT,
            cache_control: { type: "ephemeral" }
          },
          // Dynamic block — unique per request, always sent fresh.
          {
            type: "text",
            text: buildDynamicTripDetails(input)
          }
        ]
      }
    ]
  });

  // Log token usage so cache effectiveness is visible in the Vite console.
  console.log("[recommendations] usage:", {
    input: message.usage.input_tokens,
    output: message.usage.output_tokens,
    cacheWrite:
      (message.usage as unknown as Record<string, number>)
        .cache_creation_input_tokens ?? 0,
    cacheRead:
      (message.usage as unknown as Record<string, number>)
        .cache_read_input_tokens ?? 0
  });

  const textBlocks = message.content
    .filter((block: { type: string }) => block.type === "text")
    .map((block: { type: string; text?: string }) => block.text ?? "")
    .join("\n");

  if (!textBlocks.trim()) {
    throw new Error("Empty response from Claude. Try again.");
  }

  const jsonStr = extractJson(textBlocks);
  const result = JSON.parse(jsonStr);

  // Populate response cache and persist to disk.
  responseCache.set(cacheKey, {
    data: result,
    expiresAt: Date.now() + CACHE_TTL_MS
  });
  saveCacheToDisk();

  return result;
}

// ---------------------------------------------------------------------------
// HTTP body reader
// ---------------------------------------------------------------------------

function readBody(req: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

// ---------------------------------------------------------------------------
// Vite plugin
// ---------------------------------------------------------------------------

export function recommendationsPlugin(): Plugin {
  return {
    name: "travel-recommendations",
    configureServer(server: ViteDevServer) {
      server.middlewares.use(
        "/api/exchange-rate",
        async (req: IncomingMessage, res: ServerResponse) => {
          if (req.method !== "POST") {
            res.statusCode = 405;
            res.setHeader("Content-Type", "application/json");
            res.end(JSON.stringify({ error: "Method not allowed" }));
            return;
          }
          try {
            const rawBody = await readBody(req);
            const body = JSON.parse(rawBody) as { from?: string; to?: string };
            if (!body.from?.trim() || !body.to?.trim()) {
              res.statusCode = 400;
              res.setHeader("Content-Type", "application/json");
              res.end(JSON.stringify({ error: "from and to are required" }));
              return;
            }
            const result = await generateExchangeRate(
              body.from.toUpperCase(),
              body.to.toUpperCase()
            );
            res.statusCode = 200;
            res.setHeader("Content-Type", "application/json");
            res.end(JSON.stringify(result));
          } catch (err) {
            const message =
              err instanceof Error ? err.message : "Internal error";
            console.error("[exchange-rate] error:", err);
            res.statusCode = 500;
            res.setHeader("Content-Type", "application/json");
            res.end(JSON.stringify({ error: message }));
          }
        }
      );

      server.middlewares.use(
        "/api/recommendations",
        async (req: IncomingMessage, res: ServerResponse) => {
          if (req.method !== "POST") {
            res.statusCode = 405;
            res.setHeader("Content-Type", "application/json");
            res.end(JSON.stringify({ error: "Method not allowed" }));
            return;
          }
          try {
            const rawBody = await readBody(req);
            const body = JSON.parse(rawBody) as RecommendationInput;
            if (!body.destination?.trim()) {
              res.statusCode = 400;
              res.setHeader("Content-Type", "application/json");
              res.end(JSON.stringify({ error: "destination is required" }));
              return;
            }
            const result = await generateRecommendations(body);
            res.statusCode = 200;
            res.setHeader("Content-Type", "application/json");
            res.end(JSON.stringify(result));
          } catch (err) {
            const message =
              err instanceof Error ? err.message : "Internal error";
            console.error("[recommendations] error:", err);
            res.statusCode = 500;
            res.setHeader("Content-Type", "application/json");
            res.end(JSON.stringify({ error: message }));
          }
        }
      );
    }
  };
}
