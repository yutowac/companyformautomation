import type { FormData, Language } from './types';
import {
  submitApplication,
  downloadRegistrationApplication,
  downloadArticleOfIncorporation,
  downloadSealRegistration,
} from './api';
import { getLanguage, setLanguage, t, getTranslations } from './i18n';

type PageId = 'landing' | 'form' | 'thanks';

function showPage(pageId: PageId): void {
  const newPage = document.getElementById(`page-${pageId}`) as HTMLElement | null;
  if (!newPage) return;

  const current = document.querySelector<HTMLElement>('.page.page-visible.page-active');

  // すでに表示中なら何もしない
  if (current === newPage) {
    return;
  }

  // 退場アニメーション（現在のページ）
  if (current) {
    current.classList.remove('page-active');
    current.classList.add('page-leave');

    const handleLeaveEnd = (event: TransitionEvent) => {
      if (event.propertyName !== 'opacity') return;
      current.classList.remove('page-leave', 'page-visible');
    };

    current.addEventListener('transitionend', handleLeaveEnd, { once: true });
  }

  // 入場アニメーション（新しいページ）
  newPage.classList.remove('page-leave');
  newPage.classList.add('page-visible');
  // 再描画を挟んでからactiveを付与してフェードインさせる
  void newPage.offsetWidth;
  newPage.classList.add('page-active');
}

function routeFromHash(hash: string): PageId {
  switch (hash) {
    case '#/form':
      return 'form';
    case '#/thanks':
      return 'thanks';
    case '#/':
    default:
      return 'landing';
  }
}

function renderRoute(): void {
  let hash = window.location.hash;
  if (!hash || hash === '#') {
    // デフォルトはトップページ
    hash = '#/';
    if (window.location.hash !== hash) {
      window.location.hash = hash;
      return;
    }
  }
  const pageId = routeFromHash(hash);
  showPage(pageId);
}

function handleHashChange(): void {
  renderRoute();
}

function updateLanguageToggleUI(lang: Language): void {
  const toggle = document.getElementById('lang-toggle');
  if (!toggle) return;
  toggle.setAttribute('data-lang', lang);
  if (lang === 'ja') {
    toggle.classList.add('is-ja');
    toggle.classList.remove('is-en');
  } else {
    toggle.classList.add('is-en');
    toggle.classList.remove('is-ja');
  }
}

function toggleLanguage(): void {
  const current = getLanguage();
  const next: Language = current === 'ja' ? 'en' : 'ja';
  switchLanguage(next);
  updateLanguageToggleUI(next);
}

function getElementById<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Element with id "${id}" not found`);
  }
  return element as T;
}

function getInputValue(id: string): string {
  const input = getElementById<HTMLInputElement>(id);
  return input.value;
}

function getFormData(): FormData {
  return {
    companyName: getInputValue('companyName'),
    presidentName: getInputValue('presidentName'),
    presidentAddress: getInputValue('presidentAddress'),
    birthyear: parseInt(getInputValue('birthyear'), 10),
    birthmonth: parseInt(getInputValue('birthmonth'), 10),
    birthday: parseInt(getInputValue('birthday'), 10),
    purpose1: getInputValue('purpose1'),
    purpose2: getInputValue('purpose2'),
    purpose3: getInputValue('purpose3'),
    purpose4: getInputValue('purpose4'),
    purpose5: getInputValue('purpose5'),
    email: getInputValue('email'),
  };
}

function validateEmail(email: string): boolean {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function validateFormData(data: FormData): boolean {
  if (
    !data.companyName ||
    !data.presidentName ||
    !data.presidentAddress ||
    !data.purpose1
  ) {
    alert(t('errorRequiredFields'));
    return false;
  }
  
  if (!data.email || !validateEmail(data.email)) {
    alert(t('errorInvalidEmail'));
    return false;
  }
  
  return true;
}

// 生成されたファイル名を保存（ダウンロードボタン用・現在は非表示だが将来用に保持）
let generatedFilenames: {
  registrationApplication?: string;
  articleOfIncorporation?: string;
  sealRegistration?: string;
} = {};

async function submitForm(): Promise<void> {
  const formData = getFormData();

  if (!validateFormData(formData)) {
    return;
  }

  try {
    await submitApplication(formData);
    // URLハッシュでサンクスページへ遷移
    window.location.hash = '#/thanks';
  } catch (error) {
    const message = error instanceof Error ? error.message : t('errorSubmissionFailed');
    alert(`${t('errorSubmissionFailed')}: ${message}`);
  }
}

// 多言語対応: UI要素を更新する関数
function updateUI(lang?: Language): void {
  const currentLang = lang || getLanguage();
  const translations = getTranslations(currentLang);

  // data-i18n属性を持つ要素を更新
  document.querySelectorAll('[data-i18n]').forEach((element) => {
    const key = element.getAttribute('data-i18n');
    if (key && key in translations) {
      // ラベル内の.label-text要素を更新（ツールチップ付きラベルの場合）
      const labelText = element.querySelector('.label-text');
      if (labelText) {
        labelText.textContent = translations[key as keyof typeof translations];
      } else if (!element.querySelector('.tooltip-icon')) {
        // ツールチップアイコンがない場合は要素全体を更新
        element.textContent = translations[key as keyof typeof translations];
      }
    }
  });

  // data-i18n-placeholder属性を持つ要素を更新
  document.querySelectorAll('[data-i18n-placeholder]').forEach((element) => {
    const key = element.getAttribute('data-i18n-placeholder');
    if (key && key in translations && element instanceof HTMLInputElement) {
      element.placeholder = translations[key as keyof typeof translations];
    }
  });

  // data-i18n-tooltip属性を持つ要素を更新（カスタムツールチップ用）
  document.querySelectorAll('[data-i18n-tooltip]').forEach((element) => {
    const key = element.getAttribute('data-i18n-tooltip');
    if (key && key in translations) {
      const tooltipText = translations[key as keyof typeof translations];
      if (tooltipText) {
        element.textContent = tooltipText;
      }
    }
  });

  // HTMLのlang属性を更新
  document.documentElement.lang = currentLang;
  updateLanguageToggleUI(currentLang);
}

// 言語切り替え関数
function switchLanguage(lang: Language): void {
  setLanguage(lang);
  updateUI(lang);
  updateLanguageToggleUI(lang);
}

// ページ読み込み時にUIを更新
function initializeApp(): void {
  const currentLang = getLanguage();
  updateUI(currentLang);
  updateLanguageToggleUI(currentLang);
  renderRoute();
}

// DOMContentLoaded時に初期化
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
    window.addEventListener('hashchange', handleHashChange);
  });
} else {
  initializeApp();
  window.addEventListener('hashchange', handleHashChange);
}

// ダウンロード関数（グローバルスコープに公開）
function downloadRegistrationApplicationFile(): void {
  if (generatedFilenames.registrationApplication) {
    downloadRegistrationApplication(generatedFilenames.registrationApplication);
  }
}

function downloadArticleOfIncorporationFile(): void {
  if (generatedFilenames.articleOfIncorporation) {
    downloadArticleOfIncorporation(generatedFilenames.articleOfIncorporation);
  }
}

function downloadSealRegistrationFile(): void {
  if (generatedFilenames.sealRegistration) {
    downloadSealRegistration(generatedFilenames.sealRegistration);
  }
}

// グローバルスコープに公開（HTMLのonclickから呼び出すため）
(window as any).showPage = showPage;
(window as any).toggleLanguage = toggleLanguage;
(window as any).submitForm = submitForm;
(window as any).downloadWordFile = downloadRegistrationApplicationFile;
(window as any).downloadWordFile2 = downloadArticleOfIncorporationFile;
(window as any).downloadExcelFile = downloadSealRegistrationFile;
(window as any).switchLanguage = switchLanguage;

