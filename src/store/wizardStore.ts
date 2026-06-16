import { create } from 'zustand'
import type { WizardState, SheetId } from '../types/wizard'
import { DEFAULT_STATE } from '../types/wizard'

interface WizardActions {
  setField: <K extends keyof WizardState>(key: K, value: WizardState[K]) => void
  setSheet: (id: SheetId, enabled: boolean) => void
  nextStep: () => void
  prevStep: () => void
  goToStep: (step: 1 | 2 | 3 | 4) => void
  reset: () => void
}

export const useWizardStore = create<WizardState & WizardActions>((set) => ({
  ...DEFAULT_STATE,

  setField: (key, value) => set((state) => ({ ...state, [key]: value })),

  setSheet: (id, enabled) =>
    set((state) => ({
      sheets: { ...state.sheets, [id]: enabled },
    })),

  nextStep: () =>
    set((state) => ({
      currentStep: Math.min(state.currentStep + 1, 4) as 1 | 2 | 3 | 4,
    })),

  prevStep: () =>
    set((state) => ({
      currentStep: Math.max(state.currentStep - 1, 1) as 1 | 2 | 3 | 4,
    })),

  goToStep: (step) => set({ currentStep: step }),

  reset: () => set({ ...DEFAULT_STATE }),
}))
