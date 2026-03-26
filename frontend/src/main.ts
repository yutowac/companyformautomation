import type { FormData, Language } from './types';
import {
  submitApplication,
  downloadRegistrationApplication,
  downloadArticleOfIncorporation,
  downloadSealRegistration,
} from './api';
import { getLanguage, setLanguage, t, getTranslations } from './i18n';

type PageId = 'landing' | 'form' | 'confirm' | 'thanks';

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

  // ホームに戻ったタイミングでフォームをクリア
  if (pageId === 'landing') {
    resetFormFields();
    generatedFilenames = {};
  }
}

function routeFromHash(hash: string): PageId {
  switch (hash) {
    case '#/form':
      return 'form';
    case '#/confirm':
      return 'confirm';
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
    presidentNameLocal: getInputValue('presidentNameLocal'),
    presidentAddress: getInputValue('presidentAddress'),
    presidentAddressLocal: getInputValue('presidentAddressLocal'),
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

function resetFormFields(): void {
  const ids = [
    'companyName',
    'presidentName',
    'presidentNameLocal',
    'presidentAddress',
    'presidentAddressLocal',
    'birthyear',
    'birthmonth',
    'birthday',
    'purpose1',
    'purpose2',
    'purpose3',
    'purpose4',
    'purpose5',
    'email',
  ];

  ids.forEach((id) => {
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (el) {
      el.value = '';
    }
  });

  // purpose2〜purpose5 は初期状態だと非表示
  const purposeGroups = ['purpose2Group', 'purpose3Group', 'purpose4Group', 'purpose5Group'];
  purposeGroups.forEach((id) => {
    const group = document.getElementById(id) as HTMLElement | null;
    if (group) group.style.display = 'none';
  });

  const addBtn = document.getElementById('addPurposeButton') as HTMLElement | null;
  if (addBtn) addBtn.style.display = '';
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

let pendingFormData: FormData | null = null;

function formatBirthDay(data: FormData): string {
  const y = data.birthyear;
  const m = data.birthmonth;
  const d = data.birthday;
  if ([y, m, d].some((n) => Number.isNaN(n))) return '';
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

function renderConfirmPage(data: FormData): void {
  const setText = (id: string, value: string) => {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  };

  setText('confirm_companyName', data.companyName || '-');
  setText('confirm_presidentName', data.presidentName || '-');
  setText('confirm_presidentNameLocal', data.presidentNameLocal || '-');
  setText('confirm_birthDay', formatBirthDay(data) || '-');
  setText('confirm_presidentAddress', data.presidentAddress || '-');
  setText('confirm_presidentAddressLocal', data.presidentAddressLocal || '-');

  const purposes = [data.purpose1, data.purpose2, data.purpose3, data.purpose4, data.purpose5]
    .map((p) => p?.trim())
    .filter((p) => p);
  setText('confirm_purposes', purposes.length ? purposes.join(', ') : '-');

  setText('confirm_email', data.email || '-');
}

function addPurpose(): void {
  const order = ['purpose2Group', 'purpose3Group', 'purpose4Group', 'purpose5Group'];
  const nextId = order.find((id) => {
    const group = document.getElementById(id) as HTMLElement | null;
    return group && group.style.display === 'none';
  });

  if (!nextId) return;

  const nextGroup = document.getElementById(nextId) as HTMLElement | null;
  if (nextGroup) nextGroup.style.display = 'block';

  // purpose5 まで出たら add ボタンを隠す
  const allVisible = order.every((id) => {
    const group = document.getElementById(id) as HTMLElement | null;
    return group && group.style.display !== 'none';
  });
  const addBtn = document.getElementById('addPurposeButton') as HTMLElement | null;
  if (addBtn) addBtn.style.display = allVisible ? 'none' : '';
}

async function goToConfirm(): Promise<void> {
  const formData = getFormData();
  if (!validateFormData(formData)) return;

  pendingFormData = formData;
  renderConfirmPage(formData);
  window.location.hash = '#/confirm';
}

function backToFormFromConfirm(): void {
  window.location.hash = '#/form';
}

async function confirmAndSubmit(): Promise<void> {
  if (!pendingFormData) {
    window.location.hash = '#/form';
    return;
  }

  try {
    await submitApplication(pendingFormData);
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
(window as any).goToConfirm = goToConfirm;
(window as any).confirmAndSubmit = confirmAndSubmit;
(window as any).backToFormFromConfirm = backToFormFromConfirm;
(window as any).addPurpose = addPurpose;
(window as any).downloadWordFile = downloadRegistrationApplicationFile;
(window as any).downloadWordFile2 = downloadArticleOfIncorporationFile;
(window as any).downloadExcelFile = downloadSealRegistrationFile;
(window as any).switchLanguage = switchLanguage;

