import { useCallback } from 'react'
import { useWizardStore } from '../store/wizardStore'

export function useWizardNavigation() {
  const {
    currentStep, nextStep, prevStep,
    destination, useRecommendations,
    startDate, duration,
  } = useWizardStore()

  const baseStep1Ok = startDate.length > 0 && duration >= 1
  // Generic mode: destination optional. AI mode: destination required.
  const canAdvanceStep1 = useRecommendations
    ? baseStep1Ok && destination.trim().length > 0
    : baseStep1Ok
  const canAdvanceStep2 = true
  const canAdvanceStep3 = true

  const canAdvance =
    currentStep === 1
      ? canAdvanceStep1
      : currentStep === 2
        ? canAdvanceStep2
        : currentStep === 3
          ? canAdvanceStep3
          : false

  const handleNext = useCallback(() => {
    if (canAdvance) nextStep()
  }, [canAdvance, nextStep])

  const handlePrev = useCallback(() => {
    prevStep()
  }, [prevStep])

  return { currentStep, canAdvance, handleNext, handlePrev }
}
