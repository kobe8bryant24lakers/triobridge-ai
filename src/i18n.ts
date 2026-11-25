import { useSyncExternalStore } from 'react';

export type UiLanguage = 'zh-CN' | 'en';
const storageKey = 'triobridge.ui-language.v1';
const listeners = new Set<() => void>();
let language: UiLanguage | undefined;

export function isUiLanguage(value: unknown): value is UiLanguage {
  return value === 'zh-CN' || value === 'en';
}

export function getUiLanguage(): UiLanguage {
  if (language) return language;
  try {
    const saved = window.localStorage.getItem(storageKey);
    if (isUiLanguage(saved)) return (language = saved);
  } catch {
    // Office can restrict browser storage. The installed workbench also stores this preference locally.
  }
  return 'zh-CN';
}

export function setUiLanguage(next: UiLanguage): void {
  if (!isUiLanguage(next)) return;
  const changed = getUiLanguage() !== next;
  language = next;
  if (typeof document !== 'undefined') document.documentElement.lang = next;
  try {
    if (window.localStorage.getItem(storageKey) !== next)
      window.localStorage.setItem(storageKey, next);
  } catch {
    // The current window still works when its browser storage is unavailable.
  }
  if (changed) listeners.forEach((listener) => listener());
}

function onStorage(event: StorageEvent) {
  if (event.key === storageKey && isUiLanguage(event.newValue)) setUiLanguage(event.newValue);
}

function subscribe(listener: () => void) {
  if (!listeners.size) window.addEventListener('storage', onStorage);
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
    if (!listeners.size) window.removeEventListener('storage', onStorage);
  };
}

export function useUiLanguage(): UiLanguage {
  return useSyncExternalStore(subscribe, getUiLanguage, () => 'zh-CN' as const);
}

/** Select only application-authored copy. Document, model and form values pass through unchanged. */
export function t(zh: string, en: string): string {
  return getUiLanguage() === 'en' ? en : zh;
}
