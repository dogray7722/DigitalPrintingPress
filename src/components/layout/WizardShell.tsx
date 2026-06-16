import { ChevronLeft, ChevronRight, Printer } from 'lucide-react'
import { StepProgress } from './StepProgress'
import { Step1Destination } from '../steps/Step1Destination'
import { Step2Budget } from '../steps/Step2Budget'
import { Step3Style } from '../steps/Step3Style'
import { Step4Preview } from '../steps/Step4Preview'
import { useWizardStore } from '../../store/wizardStore'
import { useWizardNavigation } from '../../hooks/useWizardNavigation'
import { cn } from '../../lib/utils'

const STEP_TITLES = [
  { title: 'Where are you going?', subtitle: 'Set your destination, dates, and travel party' },
  { title: 'What\'s your budget?', subtitle: 'Configure estimated costs and choose your travel style' },
  { title: 'Customize the look', subtitle: 'Pick a theme, chart style, font, and included sheets' },
  { title: 'Ready to generate', subtitle: 'Review your settings and download your planner' },
]

export function WizardShell() {
  const { goToStep, reset } = useWizardStore()
  const { currentStep, canAdvance, handleNext, handlePrev } = useWizardNavigation()

  const { title, subtitle } = STEP_TITLES[currentStep - 1]

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-indigo-950 to-slate-900 flex flex-col">
      {/* App header */}
      <header className="border-b border-white/10 backdrop-blur-sm">
        <div className="max-w-3xl mx-auto px-4 py-4 flex items-center justify-between">
          <button
            onClick={reset}
            className="flex items-center gap-2.5 text-white hover:text-indigo-200 transition-colors"
          >
            <div className="w-8 h-8 bg-indigo-500 rounded-lg flex items-center justify-center shadow-lg">
              <Printer size={16} className="text-white" />
            </div>
            <div className="text-left">
              <div className="text-sm font-bold tracking-wide">Digital Printing Press</div>
              <div className="text-xs text-indigo-300 leading-none">Travel Planner Generator</div>
            </div>
          </button>
          <div className="text-xs text-indigo-400 hidden sm:block">
            Generates professional .xlsx files
          </div>
        </div>
      </header>

      {/* Main content */}
      <main className="flex-1 flex items-start justify-center px-4 py-8">
        <div className="w-full max-w-2xl">
          {/* Progress */}
          <div className="bg-white/5 backdrop-blur rounded-2xl p-4 mb-5 border border-white/10">
            <StepProgress
              currentStep={currentStep}
              onGoToStep={(s) => goToStep(s as 1 | 2 | 3 | 4)}
            />
          </div>

          {/* Step card */}
          <div className="bg-white rounded-2xl shadow-2xl overflow-hidden">
            {/* Step header */}
            <div className="bg-gradient-to-r from-indigo-600 to-purple-600 px-6 pt-6 pb-5">
              <div className="flex items-center gap-2 mb-1">
                <span className="text-xs font-semibold text-indigo-200 uppercase tracking-widest">
                  Step {currentStep} of 4
                </span>
              </div>
              <h1 className="text-xl font-bold text-white">{title}</h1>
              <p className="text-sm text-indigo-200 mt-1">{subtitle}</p>
            </div>

            {/* Step body */}
            <div className="px-6 py-6 animate-fade-in">
              {currentStep === 1 && <Step1Destination />}
              {currentStep === 2 && <Step2Budget />}
              {currentStep === 3 && <Step3Style />}
              {currentStep === 4 && <Step4Preview />}
            </div>

            {/* Navigation */}
            {currentStep < 4 && (
              <div className="px-6 pb-6 flex items-center justify-between">
                <button
                  type="button"
                  onClick={handlePrev}
                  disabled={currentStep === 1}
                  className={cn(
                    'flex items-center gap-1.5 px-4 py-2.5 rounded-xl text-sm font-medium transition-all',
                    currentStep === 1
                      ? 'text-gray-300 cursor-default'
                      : 'text-gray-600 hover:bg-gray-100'
                  )}
                >
                  <ChevronLeft size={16} />
                  Back
                </button>

                <button
                  type="button"
                  onClick={handleNext}
                  disabled={!canAdvance}
                  className={cn(
                    'flex items-center gap-1.5 px-6 py-2.5 rounded-xl text-sm font-semibold transition-all',
                    canAdvance
                      ? 'bg-indigo-600 text-white hover:bg-indigo-500 shadow-sm shadow-indigo-200'
                      : 'bg-gray-100 text-gray-400 cursor-not-allowed'
                  )}
                >
                  {currentStep === 3 ? 'Review & Generate' : 'Continue'}
                  <ChevronRight size={16} />
                </button>
              </div>
            )}
          </div>

          {/* Footer note */}
          <p className="text-center text-xs text-indigo-400/60 mt-4">
            Files are generated locally in your browser. No data is sent to any server.
          </p>
        </div>
      </main>
    </div>
  )
}
