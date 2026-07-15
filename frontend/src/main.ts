import type { FormData, Language, MeResponse } from './types';
import {
  submitApplication,
  downloadRegistrationApplication,
  downloadArticleOfIncorporation,
  downloadSealRegistration,
  loginRequest,
  getMe,
  listApplications,
  clearAccessToken,
  getAccessToken,
} from './api';
import { getLanguage, setLanguage, t, getTranslations } from './i18n';

type ShellPageId = 'page-home' | 'page-form' | 'page-confirm' | 'page-thanks' | 'page-applications';

let hasApplication = false;
let applicationStatus: string | null = null;
let applicationSubmittedAt: string | null = null;

const STATUS_LABELS: Record<string, string> = {
  pending: 'Pending',
  in_review: 'In review',
  completed: 'Completed',
  rejected: 'Rejected',
};

function getStatusLabel(status: string | null | undefined): string {
  if (!status) return 'No application';
  return STATUS_LABELS[status] ?? status;
}

function getRawHash(): string {
  const h = window.location.hash;
  if (!h || h === '#') {
    return '#/';
  }
  return h;
}

function normalizeLegacyAppHash(h: string): string {
  const map: Record<string, string> = {
    '#/form': '#/app/form',
    '#/confirm': '#/app/confirm',
    '#/thanks': '#/app/thanks',
  };
  if (map[h]) {
    return map[h];
  }
  if (h === '#/app' || h === '#/app/') {
    return '#/app/home';
  }
  return h;
}

function showLoginView(): void {
  const login = document.getElementById('page-login');
  const shell = document.getElementById('app-shell');
  shell?.classList.add('hidden');
  login?.classList.remove('hidden');
  login?.classList.add('page-visible', 'page-active');
}

function showAppShellView(): void {
  const login = document.getElementById('page-login');
  const shell = document.getElementById('app-shell');
  login?.classList.add('hidden');
  login?.classList.remove('page-visible', 'page-active');
  shell?.classList.remove('hidden');
}

function showShellPage(pageId: ShellPageId): void {
  document.querySelectorAll<HTMLElement>('.app-main .page').forEach((p) => {
    p.classList.remove('page-visible', 'page-active', 'page-leave');
  });
  const el = document.getElementById(pageId);
  if (!el) {
    return;
  }
  el.classList.add('page-visible', 'page-active');
}

function applyMeState(me: MeResponse): void {
  hasApplication = me.has_application;
  applicationStatus = me.application_status;
  applicationSubmittedAt = me.application_submitted_at;
  updateHomeButtons();
}

async function syncMe(): Promise<void> {
  const me = await getMe();
  applyMeState(me);
}

function updateHomeButtons(): void {
  const btnNew = document.getElementById('btnNewApplication') as HTMLButtonElement | null;
  const btnView = document.getElementById('btnViewApplication') as HTMLButtonElement | null;
  if (btnNew) {
    btnNew.disabled = hasApplication;
    btnNew.classList.toggle('nav-disabled', hasApplication);
  }
  if (btnView) {
    btnView.disabled = !hasApplication;
    btnView.classList.toggle('nav-disabled', !hasApplication);
  }
}

function updateHomeStatusCard(): void {
  const valueEl = document.getElementById('homeStatusValue');
  const metaEl = document.getElementById('homeStatusMeta');
  if (!valueEl) return;

  if (!hasApplication) {
    valueEl.textContent = 'No application';
    if (metaEl) metaEl.textContent = '';
    return;
  }

  valueEl.textContent = getStatusLabel(applicationStatus);
  if (metaEl) {
    metaEl.textContent = applicationSubmittedAt
      ? `Submitted: ${formatDisplayDate(applicationSubmittedAt)}`
      : '';
  }
}

function formatDisplayDate(iso: string): string {
  try {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleString('en-US');
  } catch {
    return iso;
  }
}

function setText(id: string, value: string): void {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function payloadToFormData(payload: Record<string, unknown>): FormData {
  const num = (v: unknown): number => {
    if (typeof v === 'number') return v;
    if (typeof v === 'string') return parseInt(v, 10);
    return NaN;
  };
  const str = (v: unknown): string => (typeof v === 'string' ? v : '');
  return {
    companyName: str(payload.companyName),
    presidentName: str(payload.presidentName),
    presidentNameLocal: str(payload.presidentNameLocal),
    presidentAddress: str(payload.presidentAddress),
    presidentAddressLocal: str(payload.presidentAddressLocal),
    birthyear: num(payload.birthyear),
    birthmonth: num(payload.birthmonth),
    birthday: num(payload.birthday),
    purpose1: str(payload.purpose1),
    purpose2: str(payload.purpose2),
    purpose3: str(payload.purpose3),
    purpose4: str(payload.purpose4),
    purpose5: str(payload.purpose5),
    email: str(payload.email),
  };
}

function formatBirthDay(data: FormData): string {
  const y = data.birthyear;
  const m = data.birthmonth;
  const d = data.birthday;
  if ([y, m, d].some((n) => Number.isNaN(n))) return '';
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

function renderApplicationSummary(data: FormData, prefix: 'confirm' | 'detail'): void {
  setText(`${prefix}_companyName`, data.companyName || '-');
  setText(`${prefix}_presidentName`, data.presidentName || '-');
  setText(`${prefix}_presidentNameLocal`, data.presidentNameLocal || '-');
  setText(`${prefix}_birthDay`, formatBirthDay(data) || '-');
  setText(`${prefix}_presidentAddress`, data.presidentAddress || '-');
  setText(`${prefix}_presidentAddressLocal`, data.presidentAddressLocal || '-');

  const purposes = [data.purpose1, data.purpose2, data.purpose3, data.purpose4, data.purpose5]
    .map((p) => p?.trim())
    .filter((p) => p);
  setText(`${prefix}_purposes`, purposes.length ? purposes.join(', ') : '-');
  setText(`${prefix}_email`, data.email || '-');
}

function renderConfirmPage(data: FormData): void {
  renderApplicationSummary(data, 'confirm');
}

const DEFAULT_CHANGE_REQUEST_URL =
  'https://docs.google.com/forms/d/e/1FAIpQLSfFaHomIpvrEwNFCYvQ6s8u7XwYZaiuC2VitvBGJUXK8Hu6Fw/viewform?usp=publish-editor';

function setChangeRequestLinkHref(): void {
  const el = document.getElementById('changeRequestLink') as HTMLAnchorElement | null;
  if (!el) return;
  const raw = import.meta.env.VITE_APPLICATION_CHANGE_URL;
  const url = typeof raw === 'string' && raw.trim() ? raw.trim() : DEFAULT_CHANGE_REQUEST_URL;
  el.href = url;
}

async function loadApplicationDetail(): Promise<void> {
  const badge = document.getElementById('applicationStatusBadge');
  const meta = document.getElementById('applicationSubmittedAt');

  try {
    const rows = await listApplications();
    if (rows.length === 0) {
      window.location.hash = '#/app/home';
      return;
    }
    const row = rows[0];
    const statusLabel = getStatusLabel(row.status);
    if (badge) {
      badge.textContent = statusLabel;
      badge.className = `status-badge status-${row.status || 'pending'}`;
    }
    if (meta) {
      const when = row.submitted_at || row.created_at;
      meta.textContent = when ? `Submitted: ${formatDisplayDate(when)}` : '';
    }
    renderApplicationSummary(payloadToFormData(row.payload), 'detail');
  } catch {
    if (badge) badge.textContent = 'Error';
    if (meta) meta.textContent = 'Failed to load application.';
  }
}

function renderRoute(): void {
  const raw = getRawHash();
  const mapped = normalizeLegacyAppHash(raw);
  if (mapped !== raw) {
    window.location.hash = mapped;
    return;
  }

  const token = getAccessToken();

  if (!token) {
    showLoginView();
    if (!mapped.startsWith('#/login')) {
      window.location.hash = '#/login';
    }
    return;
  }

  showAppShellView();

  if (mapped === '#/login' || mapped === '#/') {
    window.location.hash = '#/app/home';
    return;
  }

  if (!mapped.startsWith('#/app/')) {
    window.location.hash = '#/app/home';
    return;
  }

  if (mapped === '#/app/home') {
    void (async () => {
      try {
        await syncMe();
        updateHomeStatusCard();
        showShellPage('page-home');
      } catch {
        clearAccessToken();
        hasApplication = false;
        applicationStatus = null;
        applicationSubmittedAt = null;
        updateHomeButtons();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (mapped === '#/app/form') {
    void (async () => {
      try {
        await syncMe();
        if (hasApplication) {
          if (window.location.hash === '#/app/form') {
            window.location.hash = '#/app/home';
          }
          return;
        }
        showShellPage('page-form');
      } catch {
        clearAccessToken();
        hasApplication = false;
        updateHomeButtons();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (mapped === '#/app/confirm') {
    if (!pendingFormData) {
      window.location.hash = '#/app/form';
      return;
    }
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-confirm');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (mapped === '#/app/thanks') {
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-thanks');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (mapped === '#/app/applications') {
    void (async () => {
      try {
        await syncMe();
        if (!hasApplication) {
          window.location.hash = '#/app/home';
          return;
        }
        showShellPage('page-applications');
        await loadApplicationDetail();
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  window.location.hash = '#/app/home';
}

function handleHashChange(): void {
  renderRoute();
}

async function loginSubmit(): Promise<void> {
  const emailEl = document.getElementById('loginEmail') as HTMLInputElement | null;
  const passEl = document.getElementById('loginPassword') as HTMLInputElement | null;
  const errEl = document.getElementById('loginError') as HTMLElement | null;
  const email = emailEl?.value.trim() ?? '';
  const password = passEl?.value ?? '';
  if (!email || !password) {
    if (errEl) {
      errEl.style.display = 'block';
      errEl.textContent = 'Please enter email and password.';
    }
    return;
  }
  try {
    await loginRequest(email, password);
    if (errEl) {
      errEl.style.display = 'none';
      errEl.textContent = '';
    }
    await syncMe();
    window.location.hash = '#/app/home';
  } catch (e) {
    if (errEl) {
      errEl.style.display = 'block';
      errEl.textContent = e instanceof Error ? e.message : 'Login failed';
    }
  }
}

function logout(): void {
  clearAccessToken();
  hasApplication = false;
  applicationStatus = null;
  applicationSubmittedAt = null;
  updateHomeButtons();
  pendingFormData = null;
  resetFormFields();
  generatedFilenames = {};
  window.location.hash = '#/login';
}

function navToHome(): void {
  window.location.hash = '#/app/home';
}

function navToForm(): void {
  if (hasApplication) {
    return;
  }
  window.location.hash = '#/app/form';
}

function navToApplications(): void {
  if (!hasApplication) {
    return;
  }
  window.location.hash = '#/app/applications';
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

let generatedFilenames: {
  registrationApplication?: string;
  articleOfIncorporation?: string;
  sealRegistration?: string;
} = {};

let pendingFormData: FormData | null = null;

function addPurpose(): void {
  const order = ['purpose2Group', 'purpose3Group', 'purpose4Group', 'purpose5Group'];
  const nextId = order.find((id) => {
    const group = document.getElementById(id) as HTMLElement | null;
    return group && group.style.display === 'none';
  });

  if (!nextId) return;

  const nextGroup = document.getElementById(nextId) as HTMLElement | null;
  if (nextGroup) nextGroup.style.display = 'block';

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
  window.location.hash = '#/app/confirm';
}

function backToFormFromConfirm(): void {
  window.location.hash = '#/app/form';
}

async function confirmAndSubmit(): Promise<void> {
  if (!pendingFormData) {
    window.location.hash = '#/app/form';
    return;
  }

  try {
    await submitApplication(pendingFormData);
    await syncMe();
    window.location.hash = '#/app/thanks';
  } catch (error) {
    const message = error instanceof Error ? error.message : t('errorSubmissionFailed');
    alert(`${t('errorSubmissionFailed')}: ${message}`);
  }
}

function updateUI(lang?: Language): void {
  const currentLang = lang || getLanguage();
  const translations = getTranslations(currentLang);

  document.querySelectorAll('[data-i18n]').forEach((element) => {
    const key = element.getAttribute('data-i18n');
    if (key && key in translations) {
      const labelText = element.querySelector('.label-text');
      if (labelText) {
        labelText.textContent = translations[key as keyof typeof translations];
      } else if (!element.querySelector('.tooltip-icon')) {
        element.textContent = translations[key as keyof typeof translations];
      }
    }
  });

  document.querySelectorAll('[data-i18n-placeholder]').forEach((element) => {
    const key = element.getAttribute('data-i18n-placeholder');
    if (key && key in translations && element instanceof HTMLInputElement) {
      element.placeholder = translations[key as keyof typeof translations];
    }
  });

  document.querySelectorAll('[data-i18n-tooltip]').forEach((element) => {
    const key = element.getAttribute('data-i18n-tooltip');
    if (key && key in translations) {
      const tooltipText = translations[key as keyof typeof translations];
      if (tooltipText) {
        element.textContent = tooltipText;
      }
    }
  });

  document.documentElement.lang = currentLang;
  updateLanguageToggleUI(currentLang);
}

function switchLanguage(lang: Language): void {
  setLanguage(lang);
  updateUI(lang);
  updateLanguageToggleUI(lang);
}

function initializeApp(): void {
  const currentLang = getLanguage();
  updateUI(currentLang);
  updateLanguageToggleUI(currentLang);
  setChangeRequestLinkHref();
  renderRoute();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
    window.addEventListener('hashchange', handleHashChange);
  });
} else {
  initializeApp();
  window.addEventListener('hashchange', handleHashChange);
}

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

(window as any).toggleLanguage = toggleLanguage;
(window as any).loginSubmit = loginSubmit;
(window as any).logout = logout;
(window as any).navToHome = navToHome;
(window as any).navToForm = navToForm;
(window as any).navToApplications = navToApplications;
(window as any).goToConfirm = goToConfirm;
(window as any).confirmAndSubmit = confirmAndSubmit;
(window as any).backToFormFromConfirm = backToFormFromConfirm;
(window as any).addPurpose = addPurpose;
(window as any).downloadWordFile = downloadRegistrationApplicationFile;
(window as any).downloadWordFile2 = downloadArticleOfIncorporationFile;
(window as any).downloadExcelFile = downloadSealRegistrationFile;
(window as any).switchLanguage = switchLanguage;
