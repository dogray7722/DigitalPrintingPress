import { useState, useRef, useEffect } from 'react'
import { ChevronDown, Search } from 'lucide-react'
import { cn } from '../../lib/utils'
import { CURRENCIES, type Currency } from '../../data/currencies'

interface Props {
  value: string
  onChange: (code: string) => void
}

export function CurrencySelector({ value, onChange }: Props) {
  const [open, setOpen] = useState(false)
  const [query, setQuery] = useState('')
  const ref = useRef<HTMLDivElement>(null)

  const selected = CURRENCIES.find((c) => c.code === value) ?? CURRENCIES[0]

  const filtered = query.trim()
    ? CURRENCIES.filter(
        (c) =>
          c.code.toLowerCase().includes(query.toLowerCase()) ||
          c.name.toLowerCase().includes(query.toLowerCase()) ||
          c.symbol.toLowerCase().includes(query.toLowerCase())
      )
    : CURRENCIES

  useEffect(() => {
    function onClickOutside(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false)
        setQuery('')
      }
    }
    document.addEventListener('mousedown', onClickOutside)
    return () => document.removeEventListener('mousedown', onClickOutside)
  }, [])

  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex items-center gap-2 w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-sm hover:border-gray-300 focus:outline-none focus:ring-2 focus:ring-indigo-400 transition-colors"
      >
        <span className="text-lg">{selected.flag}</span>
        <span className="font-medium">{selected.code}</span>
        <span className="text-gray-500">— {selected.name}</span>
        <span className="ml-auto text-gray-400">
          <span className="font-mono text-gray-600 mr-1">{selected.symbol}</span>
          <ChevronDown size={14} className={cn('transition-transform', open && 'rotate-180')} />
        </span>
      </button>

      {open && (
        <div className="absolute z-50 mt-1 w-full bg-white border border-gray-200 rounded-lg shadow-lg max-h-72 overflow-hidden flex flex-col">
          <div className="p-2 border-b border-gray-100">
            <div className="flex items-center gap-2 px-2 py-1.5 bg-gray-50 rounded-md">
              <Search size={14} className="text-gray-400 shrink-0" />
              <input
                autoFocus
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Search currencies..."
                className="flex-1 text-sm bg-transparent outline-none placeholder:text-gray-400"
              />
            </div>
          </div>
          <div className="overflow-y-auto">
            {filtered.map((c: Currency) => (
              <button
                key={c.code}
                type="button"
                onClick={() => {
                  onChange(c.code)
                  setOpen(false)
                  setQuery('')
                }}
                className={cn(
                  'w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-indigo-50 transition-colors text-left',
                  c.code === value && 'bg-indigo-50 text-indigo-700 font-medium'
                )}
              >
                <span className="text-base">{c.flag}</span>
                <span className="font-medium w-10 shrink-0">{c.code}</span>
                <span className="text-gray-600 flex-1">{c.name}</span>
                <span className="font-mono text-gray-500 shrink-0">{c.symbol}</span>
              </button>
            ))}
            {filtered.length === 0 && (
              <div className="px-4 py-6 text-center text-sm text-gray-400">No currencies found</div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}
