// Season is derived from travelMonth
export function getSeason(month: number): 'spring' | 'summer' | 'autumn' | 'winter' {
  if (month >= 3 && month <= 5) return 'spring'
  if (month >= 6 && month <= 8) return 'summer'
  if (month >= 9 && month <= 11) return 'autumn'
  return 'winter'
}

export interface PackingItem {
  category: string
  item: string
}

const ESSENTIALS: PackingItem[] = [
  { category: 'Documents', item: 'Passport / ID' },
  { category: 'Documents', item: 'Travel insurance documents' },
  { category: 'Documents', item: 'Flight tickets / e-tickets' },
  { category: 'Documents', item: 'Hotel confirmations' },
  { category: 'Documents', item: 'Emergency contacts list' },
  { category: 'Documents', item: 'Visa / entry documents (if required)' },
  { category: 'Documents', item: 'Copies of all documents (digital + physical)' },
  { category: 'Money', item: 'Local currency (cash)' },
  { category: 'Money', item: 'Credit / debit cards' },
  { category: 'Money', item: 'Travel money card' },
  { category: 'Electronics', item: 'Phone + charger' },
  { category: 'Electronics', item: 'Universal power adapter' },
  { category: 'Electronics', item: 'Portable battery bank' },
  { category: 'Electronics', item: 'Camera + memory cards' },
  { category: 'Electronics', item: 'Earphones / headphones' },
  { category: 'Electronics', item: 'Laptop / tablet (if needed)' },
  { category: 'Toiletries', item: 'Toothbrush + toothpaste' },
  { category: 'Toiletries', item: 'Shampoo + conditioner' },
  { category: 'Toiletries', item: 'Body wash / soap' },
  { category: 'Toiletries', item: 'Deodorant' },
  { category: 'Toiletries', item: 'Sunscreen SPF 30+' },
  { category: 'Toiletries', item: 'Lip balm' },
  { category: 'Toiletries', item: 'Razor + shaving items' },
  { category: 'Toiletries', item: 'Hair brush / comb' },
  { category: 'Health', item: 'Prescription medications' },
  { category: 'Health', item: 'Pain relievers (ibuprofen / paracetamol)' },
  { category: 'Health', item: 'Antidiarrheal medication' },
  { category: 'Health', item: 'Antihistamines' },
  { category: 'Health', item: 'Antiseptic wipes / hand sanitizer' },
  { category: 'Health', item: 'Plasters / bandages' },
  { category: 'Clothing', item: 'Underwear (1 per day + 2 extra)' },
  { category: 'Clothing', item: 'Socks (1 per day + 2 extra)' },
  { category: 'Clothing', item: 'T-shirts / tops' },
  { category: 'Clothing', item: 'Casual trousers / jeans' },
  { category: 'Clothing', item: 'Smart/smart-casual outfit' },
  { category: 'Clothing', item: 'Comfortable walking shoes' },
  { category: 'Bags', item: 'Main luggage / backpack' },
  { category: 'Bags', item: 'Daypack / small backpack' },
  { category: 'Bags', item: 'Packing cubes' },
  { category: 'Bags', item: 'Laundry bag' },
]

const SUMMER_EXTRAS: PackingItem[] = [
  { category: 'Clothing', item: 'Shorts / light trousers' },
  { category: 'Clothing', item: 'Swimwear (x2)' },
  { category: 'Clothing', item: 'Sandals / flip-flops' },
  { category: 'Clothing', item: 'Sun hat / cap' },
  { category: 'Clothing', item: 'Sunglasses' },
  { category: 'Clothing', item: 'Light linen shirt' },
  { category: 'Toiletries', item: 'After-sun lotion' },
  { category: 'Toiletries', item: 'Insect repellent' },
]

const WINTER_EXTRAS: PackingItem[] = [
  { category: 'Clothing', item: 'Warm coat / parka' },
  { category: 'Clothing', item: 'Thermal base layers (top + bottom)' },
  { category: 'Clothing', item: 'Woollen jumper / fleece' },
  { category: 'Clothing', item: 'Scarf' },
  { category: 'Clothing', item: 'Gloves / mittens' },
  { category: 'Clothing', item: 'Beanie / winter hat' },
  { category: 'Clothing', item: 'Waterproof boots' },
  { category: 'Clothing', item: 'Thick socks (extra)' },
  { category: 'Health', item: 'Lip balm (extra — cold weather)' },
  { category: 'Health', item: 'Hand cream / moisturiser' },
]

const SPRING_AUTUMN_EXTRAS: PackingItem[] = [
  { category: 'Clothing', item: 'Light jacket or cardigan' },
  { category: 'Clothing', item: 'Waterproof layer / poncho' },
  { category: 'Clothing', item: 'Mix of light + warm layers' },
  { category: 'Clothing', item: 'Comfortable trainers / sneakers' },
]

export function getPackingList(travelMonth: number): PackingItem[] {
  const season = getSeason(travelMonth)
  const extras =
    season === 'summer'
      ? SUMMER_EXTRAS
      : season === 'winter'
        ? WINTER_EXTRAS
        : SPRING_AUTUMN_EXTRAS

  return [...ESSENTIALS, ...extras]
}
