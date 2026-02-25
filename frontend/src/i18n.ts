import { ja } from './locales/ja';
import { en } from './locales/en';

export type Language = 'ja' | 'en';
export type TranslationKey = keyof typeof ja;

const translations = {
  ja,
  en,
} as const;

const STORAGE_KEY = 'app_language';

export function getLanguage(): Language {
  if (typeof window === 'undefined') return 'en';
  
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored === 'ja' || stored === 'en') {
    return stored;
  }
  
  // ブラウザの言語設定を確認
  const browserLang = navigator.language.toLowerCase();
  if (browserLang.startsWith('ja')) {
    return 'ja';
  }
  
  return 'en';
}

export function setLanguage(lang: Language): void {
  if (typeof window === 'undefined') return;
  localStorage.setItem(STORAGE_KEY, lang);
}

export function t(key: TranslationKey, lang?: Language): string {
  const currentLang = lang || getLanguage();
  return translations[currentLang][key] || translations.en[key] || key;
}

export function getTranslations(lang?: Language) {
  const currentLang = lang || getLanguage();
  return translations[currentLang];
}















