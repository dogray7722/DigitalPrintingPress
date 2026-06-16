import type { ThemeId, FontFamily, FontSize } from './wizard'

export interface ColorPalette {
  primary: string
  primaryText: string
  secondary: string
  secondaryText: string
  accent: string
  accentText: string
  lightBg: string
  mediumBg: string
  border: string
  tabColor: string
}

export const THEMES: Record<ThemeId, ColorPalette> = {
  sakura: {
    primary: 'FF7B6BA8',
    primaryText: 'FFFFFFFF',
    secondary: 'FFB8A9D4',
    secondaryText: 'FF2D1B6B',
    accent: 'FFC8BEE0',
    accentText: 'FF2D1B6B',
    lightBg: 'FFF5F0FA',
    mediumBg: 'FFE8E0F0',
    border: 'FFC8BEE0',
    tabColor: 'FF7B6BA8',
  },
  ocean: {
    primary: 'FF2D6A9F',
    primaryText: 'FFFFFFFF',
    secondary: 'FF5E9EC8',
    secondaryText: 'FF0D2840',
    accent: 'FFA8C8E8',
    accentText: 'FF0D2840',
    lightBg: 'FFEEF5FB',
    mediumBg: 'FFD4E9F5',
    border: 'FFA8C8E8',
    tabColor: 'FF2D6A9F',
  },
  forest: {
    primary: 'FF2D6B4F',
    primaryText: 'FFFFFFFF',
    secondary: 'FF5E9E78',
    secondaryText: 'FF0D2D1B',
    accent: 'FFA8D4B8',
    accentText: 'FF0D2D1B',
    lightBg: 'FFEEF7F2',
    mediumBg: 'FFD4EEE0',
    border: 'FFA8D4B8',
    tabColor: 'FF2D6B4F',
  },
  desert: {
    primary: 'FFB5762A',
    primaryText: 'FFFFFFFF',
    secondary: 'FFD4A050',
    secondaryText: 'FF3D2000',
    accent: 'FFE8CFA8',
    accentText: 'FF3D2000',
    lightBg: 'FFFBF4EE',
    mediumBg: 'FFF0E0C8',
    border: 'FFE8CFA8',
    tabColor: 'FFB5762A',
  },
  inkwell: {
    primary: 'FF2C2620',
    primaryText: 'FFF0E8D5',
    secondary: 'FFC9A96E',
    secondaryText: 'FF140F0A',
    accent: 'FFC4B89A',
    accentText: 'FF140F0A',
    lightBg: 'FFF5EFE4',
    mediumBg: 'FFEDE5D8',
    border: 'FFC4B89A',
    tabColor: 'FF2C2620',
  },
  parchment: {
    primary: 'FF7A5A18',
    primaryText: 'FFF5EFE4',
    secondary: 'FFC9A96E',
    secondaryText: 'FF2E2010',
    accent: 'FFBFAE90',
    accentText: 'FF2E2010',
    lightBg: 'FFF5EFE4',
    mediumBg: 'FFEDE5D8',
    border: 'FFBFAE90',
    tabColor: 'FF7A5A18',
  },
}

export const THEME_LABELS: Record<ThemeId, string> = {
  sakura: 'Sakura',
  ocean: 'Ocean Blue',
  forest: 'Forest Green',
  desert: 'Desert Sand',
  inkwell: 'Inkwell',
  parchment: 'Parchment',
}

export interface FontSizeValues {
  data: number
  header: number
  title: number
  sectionHeader: number
}

export const FONT_SIZES: Record<FontSize, FontSizeValues> = {
  small: { data: 10, header: 11, sectionHeader: 12, title: 18 },
  medium: { data: 11, header: 12, sectionHeader: 13, title: 22 },
  large: { data: 12, header: 13, sectionHeader: 14, title: 26 },
}

export const FONT_NAMES: Record<FontFamily, string> = {
  sans: 'Calibri',
  serif: 'Cambria',
  mono: 'Courier New',
}
