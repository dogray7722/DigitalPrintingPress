import { addDays, format } from 'date-fns'
import { MapPin, Calendar, Users, Sparkles, ImagePlus } from 'lucide-react'
import { useWizardStore } from '../../store/wizardStore'
import { PictureUploader } from '../shared/PictureUploader'

const PARTY_OPTIONS = [
  { id: 'solo' as const, label: 'Solo', icon: '🧍', size: 1 },
  { id: 'couple' as const, label: 'Couple', icon: '👫', size: 2 },
  { id: 'family' as const, label: 'Family', icon: '👨‍👩‍👧', size: 3 },
  { id: 'group' as const, label: 'Group', icon: '👥', size: 4 },
]

export function Step1Destination() {
  const {
    destination, useRecommendations,
    startDate, duration, partyType, partySize, overviewImage,
    setField,
  } = useWizardStore()

  const endDate = startDate
    ? format(addDays(new Date(startDate), duration - 1), 'dd MMM yyyy')
    : '—'

  const today = format(new Date(), 'yyyy-MM-dd')

  return (
    <div className="space-y-7">
      {/* Destination */}
      <div>
        <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-2">
          <MapPin size={15} className="text-indigo-500" />
          Destination
        </label>
        <input
          type="text"
          value={destination}
          onChange={(e) => setField('destination', e.target.value)}
          placeholder="e.g. Tokyo, Japan  /  Amalfi Coast, Italy  /  New York, USA"
          className="w-full px-4 py-3 rounded-xl border border-gray-200 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-400 placeholder:text-gray-400 transition-shadow"
        />
        <p className="mt-1.5 text-xs text-gray-400">Enter any destination — leave blank for a generic, location-free planner</p>
      </div>

      {/* AI Recommendations toggle */}
      <div className="rounded-xl border-2 border-dashed border-purple-200 bg-gradient-to-br from-purple-50/50 to-indigo-50/50 p-4">
        <div className="flex items-start gap-3">
          <div className="w-10 h-10 rounded-lg bg-gradient-to-br from-purple-500 to-indigo-500 flex items-center justify-center shrink-0 shadow-md shadow-purple-200">
            <Sparkles size={18} className="text-white" />
          </div>
          <div className="flex-1">
            <div className="flex items-center justify-between gap-3 mb-1">
              <div className="font-semibold text-sm text-gray-900">AI Destination Recommendations</div>
              <button
                type="button"
                onClick={() => setField('useRecommendations', !useRecommendations)}
                className={`w-11 h-6 rounded-full flex items-center px-0.5 transition-colors shrink-0 ${
                  useRecommendations ? 'bg-purple-500 justify-end' : 'bg-gray-300 justify-start'
                }`}
              >
                <div className="w-5 h-5 bg-white rounded-full shadow-sm" />
              </button>
            </div>
            <p className="text-xs text-gray-600 leading-relaxed">
              {useRecommendations ? (
                <>
                  ✨ Claude will research <strong>{destination || 'your destination'}</strong> live with web search
                  and add informational hotel, restaurant, and excursion guides (grouped by region) to those tabs,
                  prefill a day-by-day itinerary, plus a bonus <strong>Annual Events</strong> tab. Requires a destination.
                </>
              ) : (
                <>
                  Off — generates a generic blank planner you can fill in yourself.
                  Turn on to have Claude search the web and prefill recommendations.
                </>
              )}
            </p>
          </div>
        </div>
      </div>

      {/* Dates */}
      <div>
        <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-3">
          <Calendar size={15} className="text-indigo-500" />
          Trip Dates
        </label>
        <div className="grid grid-cols-2 gap-4">
          <div>
            <div className="text-xs text-gray-500 mb-1.5 font-medium">Start Date</div>
            <input
              type="date"
              value={startDate}
              min={today}
              onChange={(e) => {
                setField('startDate', e.target.value)
                const m = new Date(e.target.value).getMonth() + 1
                setField('travelMonth', m)
              }}
              className="w-full px-3 py-2.5 rounded-xl border border-gray-200 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-400"
            />
          </div>
          <div>
            <div className="text-xs text-gray-500 mb-1.5 font-medium">Duration</div>
            <div className="flex items-center gap-3">
              <input
                type="range"
                min={2}
                max={30}
                value={duration}
                onChange={(e) => setField('duration', parseInt(e.target.value))}
                className="flex-1 accent-indigo-500"
              />
              <div className="text-sm font-semibold text-indigo-600 w-16 text-right shrink-0">
                {duration} days
              </div>
            </div>
          </div>
        </div>

        {startDate && (
          <div className="mt-3 flex gap-6 text-sm">
            <div className="flex items-center gap-2 text-gray-600">
              <span className="w-1.5 h-1.5 rounded-full bg-green-400 inline-block" />
              Depart: <span className="font-medium">{format(new Date(startDate), 'dd MMM yyyy')}</span>
            </div>
            <div className="flex items-center gap-2 text-gray-600">
              <span className="w-1.5 h-1.5 rounded-full bg-red-400 inline-block" />
              Return: <span className="font-medium">{endDate}</span>
            </div>
          </div>
        )}
      </div>

      {/* Party */}
      <div>
        <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-3">
          <Users size={15} className="text-indigo-500" />
          Travel Party
        </label>
        <div className="grid grid-cols-4 gap-2 mb-4">
          {PARTY_OPTIONS.map((opt) => (
            <button
              key={opt.id}
              type="button"
              onClick={() => {
                setField('partyType', opt.id)
                setField('partySize', opt.size)
              }}
              className={`flex flex-col items-center gap-1.5 py-3 px-2 rounded-xl border-2 transition-all ${
                partyType === opt.id
                  ? 'border-indigo-500 bg-indigo-50'
                  : 'border-gray-200 bg-white hover:border-gray-300'
              }`}
            >
              <span className="text-2xl">{opt.icon}</span>
              <span className={`text-xs font-medium ${partyType === opt.id ? 'text-indigo-700' : 'text-gray-600'}`}>
                {opt.label}
              </span>
            </button>
          ))}
        </div>

        <div className="flex items-center gap-3">
          <label className="text-sm text-gray-600 whitespace-nowrap">Number of travelers:</label>
          <input
            type="number"
            min={1}
            max={20}
            value={partySize}
            onChange={(e) => setField('partySize', Math.max(1, parseInt(e.target.value) || 1))}
            className="w-20 px-3 py-2 rounded-lg border border-gray-200 text-sm text-center font-semibold focus:outline-none focus:ring-2 focus:ring-indigo-400"
          />
          <span className="text-sm text-gray-400">people</span>
        </div>
      </div>

      {/* Cover Photo */}
      <div>
        <label className="flex items-center gap-2 text-sm font-semibold text-gray-700 mb-2">
          <ImagePlus size={15} className="text-indigo-500" />
          Cover Photo
          <span className="text-xs font-normal text-gray-400">(optional)</span>
        </label>
        <PictureUploader
          value={overviewImage}
          onChange={(img) => setField('overviewImage', img)}
        />
      </div>
    </div>
  )
}
