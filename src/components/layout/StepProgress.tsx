import { cn } from '../../lib/utils'
import { Check } from 'lucide-react'

const STEPS = [
  { num: 1, label: 'Destination' },
  { num: 2, label: 'Budget' },
  { num: 3, label: 'Style' },
  { num: 4, label: 'Generate' },
]

interface Props {
  currentStep: number
  onGoToStep?: (step: number) => void
}

export function StepProgress({ currentStep, onGoToStep }: Props) {
  return (
    <nav className="flex items-center justify-center gap-0">
      {STEPS.map((step, i) => {
        const isDone = currentStep > step.num
        const isActive = currentStep === step.num
        const isClickable = isDone && !!onGoToStep

        return (
          <div key={step.num} className="flex items-center">
            <button
              type="button"
              onClick={isClickable ? () => onGoToStep!(step.num) : undefined}
              disabled={!isClickable && !isActive}
              className={cn(
                'flex flex-col items-center gap-1 px-3 py-1 rounded-lg transition-all group',
                isClickable ? 'cursor-pointer hover:bg-indigo-50' : 'cursor-default'
              )}
            >
              <div
                className={cn(
                  'w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold transition-all',
                  isDone && 'bg-indigo-500 text-white',
                  isActive && 'bg-indigo-600 text-white ring-4 ring-indigo-200',
                  !isDone && !isActive && 'bg-gray-100 text-gray-400'
                )}
              >
                {isDone ? <Check size={14} strokeWidth={3} /> : step.num}
              </div>
              <span
                className={cn(
                  'text-xs font-medium whitespace-nowrap',
                  isActive && 'text-indigo-700',
                  isDone && 'text-indigo-500',
                  !isDone && !isActive && 'text-gray-400'
                )}
              >
                {step.label}
              </span>
            </button>

            {i < STEPS.length - 1 && (
              <div
                className={cn(
                  'w-12 h-0.5 mt-[-10px]',
                  currentStep > step.num ? 'bg-indigo-400' : 'bg-gray-200'
                )}
              />
            )}
          </div>
        )
      })}
    </nav>
  )
}
