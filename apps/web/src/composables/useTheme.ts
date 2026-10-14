import { ref } from 'vue'

export type ThemePreference = 'system' | 'dark' | 'light'

const STORAGE_KEY = 'office_theme'

const preference = ref<ThemePreference>(readStored())
let media: MediaQueryList | null = null
let listening = false

function readStored(): ThemePreference {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (raw === 'dark' || raw === 'light' || raw === 'system') return raw
  } catch {
    /* ignore */
  }
  return 'system'
}

function resolved(pref: ThemePreference): 'dark' | 'light' {
  if (pref === 'dark' || pref === 'light') return pref
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

function apply(pref: ThemePreference) {
  const mode = resolved(pref)
  document.documentElement.dataset.theme = mode
  document.documentElement.style.colorScheme = mode
}

function onSystemChange() {
  if (preference.value === 'system') apply('system')
}

export function initTheme() {
  apply(preference.value)
  if (listening) return
  media = window.matchMedia('(prefers-color-scheme: dark)')
  media.addEventListener('change', onSystemChange)
  listening = true
}

export function setTheme(pref: ThemePreference) {
  preference.value = pref
  try {
    localStorage.setItem(STORAGE_KEY, pref)
  } catch {
    /* ignore */
  }
  apply(pref)
}

export function useTheme() {
  return { preference, setTheme }
}
