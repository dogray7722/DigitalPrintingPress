import { clsx, type ClassValue } from 'clsx'
import { twMerge } from 'tailwind-merge'

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}

export function formatCurrency(amount: number, symbol: string, decimals = 2): string {
  return `${symbol}${amount.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals })}`
}

export function computeTotalBudget(
  budgets: { flights: number; hotel: number; food: number; activities: number; transport: number; shopping: number; misc: number },
  duration: number
): number {
  const { flights, hotel, food, activities, transport, shopping, misc } = budgets
  // flights and shopping are per-trip totals; the rest are per-day rates.
  return flights + shopping + (hotel + food + activities + transport + misc) * duration
}
