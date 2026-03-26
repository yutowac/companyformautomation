import { ja } from './locales/ja';
import { en } from './locales/en';

export type Language = 'ja' | 'en';
export type TranslationKey = keyof typeof ja;

const translations = {
  ja,
  en,
} as const;

export function getLanguage(): Language {
  return 'en';
}

export function setLanguage(_lang: Language): void {
  // English only mode: language state is fixed to 'en'.
}

export function t(key: TranslationKey, lang?: Language): string {
  const currentLang = lang || getLanguage();
  return translations[currentLang][key] || translations.en[key] || key;
}

export function getTranslations(lang?: Language) {
  const currentLang = lang || getLanguage();
  return translations[currentLang];
}















