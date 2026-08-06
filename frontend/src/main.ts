import type { FormData, Language, MeResponse, PaymentFormData, PaymentAccountType, PaymentRequestItem } from './types';
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
  changePassword,
  listPaymentRequests,
  getPaymentRequest,
  createPaymentRequest,
  updatePaymentRequest,
} from './api';
import { getLanguage, setLanguage, t, getTranslations } from './i18n';

type ShellPageId =
  | 'page-hub'
  | 'page-profile'
  | 'page-bank'
  | 'page-home'
  | 'page-form'
  | 'page-confirm'
  | 'page-complete'
  | 'page-applications'
  | 'page-payment-list'
  | 'page-payment-form'
  | 'page-payment-confirm'
  | 'page-payment-complete'
  | 'page-payment-status';

let hasApplication = false;
let applicationStatus: string | null = null;
let applicationSubmittedAt: string | null = null;
let currentUserEmail: string | null = null;

const STATUS_LABELS: Record<string, string> = {
  pending: 'Pending',
  in_review: 'In review',
  completed: 'Completed',
  rejected: 'Rejected',
};

const PAYMENT_STATUS_LABELS: Record<string, string> = {
  pending: 'Pending',
  paid: 'Paid',
};

const ACCOUNT_TYPE_LABELS: Record<string, string> = {
  checking: 'Checking',
  ordinary: 'Ordinary',
  savings: 'Savings',
};

function getStatusLabel(status: string | null | undefined): string {
  if (!status) return 'No application';
  return STATUS_LABELS[status] ?? status;
}

function getPaymentStatusLabel(status: string | null | undefined): string {
  if (!status) return 'Pending';
  return PAYMENT_STATUS_LABELS[status] ?? status;
}

function getRawHash(): string {
  const h = window.location.hash;
  if (!h || h === '#') {
    return '#/';
  }
  return h;
}

function normalizeLegacyAppHash(path: string): string {
  const map: Record<string, string> = {
    '#/app/home': '#/incorporation',
    '#/app/form': '#/incorporation/form',
    '#/app/confirm': '#/incorporation/confirm',
    '#/app/thanks': '#/incorporation/complete',
    '#/app/applications': '#/incorporation/status',
    '#/form': '#/incorporation/form',
    '#/confirm': '#/incorporation/confirm',
    '#/thanks': '#/incorporation/complete',
  };
  if (map[path]) {
    return map[path];
  }
  if (path === '#/app' || path === '#/app/') {
    return '#/incorporation';
  }
  return path;
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
  currentUserEmail = me.email;
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

function formatJpy(value: unknown): string {
  const s = typeof value === 'string' ? value : typeof value === 'number' ? String(value) : '';
  const n = parseInt(s.replace(/[^0-9]/g, ''), 10);
  if (Number.isNaN(n)) return '¥-';
  return `¥${n.toLocaleString('en-US')}`;
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
      window.location.hash = '#/incorporation';
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

/* ------------------------------------------------------------------ */
/* Monthly Payment Requests helpers                                    */
/* ------------------------------------------------------------------ */

let pendingPaymentData: PaymentFormData | null = null;
let pendingPaymentAttachment: File | null = null;
let pendingPaymentEditId: number | null = null;
let pendingPaymentExistingAttachmentName: string | null = null;
let currentPaymentDetail: PaymentRequestItem | null = null;
let suppressPaymentFormReload = false;

function payloadToPaymentFormData(payload: Record<string, unknown>): PaymentFormData {
  const str = (v: unknown): string => (typeof v === 'string' ? v : v != null ? String(v) : '');
  return {
    payeeName: str(payload.payeeName),
    bankName: str(payload.bankName),
    branchName: str(payload.branchName),
    accountType: str(payload.accountType) as PaymentAccountType | '',
    accountNumber: str(payload.accountNumber),
    accountHolderKana: str(payload.accountHolderKana),
    amountJpy: str(payload.amountJpy),
    invoiceNumber: str(payload.invoiceNumber),
  };
}

function renderPaymentSummary(
  data: PaymentFormData,
  prefix: string,
  attachmentUrl: string | null,
  attachmentName: string | null,
): void {
  setText(`${prefix}_payeeName`, data.payeeName || '-');
  setText(`${prefix}_bankName`, data.bankName || '-');
  setText(`${prefix}_branchName`, data.branchName || '-');
  setText(`${prefix}_accountType`, (data.accountType && ACCOUNT_TYPE_LABELS[data.accountType]) || '-');
  setText(`${prefix}_accountNumber`, data.accountNumber || '-');
  setText(`${prefix}_accountHolderKana`, data.accountHolderKana || '-');
  setText(`${prefix}_amountJpy`, data.amountJpy ? formatJpy(data.amountJpy) : '-');
  setText(`${prefix}_invoiceNumber`, data.invoiceNumber || '-');

  const attachEl = document.getElementById(`${prefix}_attachment`);
  if (attachEl) {
    attachEl.textContent = '';
    if (attachmentUrl && attachmentName) {
      const a = document.createElement('a');
      a.href = attachmentUrl;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.className = 'change-request-link';
      a.textContent = attachmentName;
      attachEl.appendChild(a);
    } else if (attachmentName) {
      attachEl.textContent = attachmentName;
    } else {
      attachEl.textContent = 'None';
    }
  }
}

function renderPaymentConfirmPage(
  data: PaymentFormData,
  newAttachment: File | null,
  existingAttachmentName: string | null,
): void {
  const attachmentName = newAttachment ? newAttachment.name : existingAttachmentName;
  renderPaymentSummary(data, 'paymentConfirm', null, attachmentName);
}

function updatePaymentFormTitle(): void {
  const titleEl = document.getElementById('paymentFormTitle');
  if (titleEl) {
    titleEl.textContent = pendingPaymentEditId != null ? 'Edit Payment Request' : 'New Payment Request';
  }
}

function populatePaymentForm(data: PaymentFormData): void {
  (document.getElementById('paymentPayeeName') as HTMLInputElement).value = data.payeeName;
  (document.getElementById('paymentBankName') as HTMLInputElement).value = data.bankName;
  (document.getElementById('paymentBranchName') as HTMLInputElement).value = data.branchName;
  (document.getElementById('paymentAccountType') as HTMLSelectElement).value = data.accountType;
  (document.getElementById('paymentAccountNumber') as HTMLInputElement).value = data.accountNumber;
  (document.getElementById('paymentAccountHolderKana') as HTMLInputElement).value = data.accountHolderKana;
  (document.getElementById('paymentAmountJpy') as HTMLInputElement).value = data.amountJpy;
  (document.getElementById('paymentInvoiceNumber') as HTMLInputElement).value = data.invoiceNumber;
  const fileInput = document.getElementById('paymentAttachment') as HTMLInputElement | null;
  if (fileInput) fileInput.value = '';
}

function resetPaymentFormFields(): void {
  const ids = [
    'paymentPayeeName',
    'paymentBankName',
    'paymentBranchName',
    'paymentAccountNumber',
    'paymentAccountHolderKana',
    'paymentAmountJpy',
    'paymentInvoiceNumber',
  ];
  ids.forEach((id) => {
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (el) el.value = '';
  });
  const select = document.getElementById('paymentAccountType') as HTMLSelectElement | null;
  if (select) select.value = '';
  const fileInput = document.getElementById('paymentAttachment') as HTMLInputElement | null;
  if (fileInput) fileInput.value = '';
}

function showExistingAttachment(name: string | null | undefined): void {
  const el = document.getElementById('paymentExistingAttachment');
  if (!el) return;
  if (name) {
    el.style.display = 'block';
    el.textContent = `Current attachment: ${name} (choose a new file to replace it)`;
  } else {
    el.style.display = 'none';
    el.textContent = '';
  }
}

function getPaymentFormData(): PaymentFormData {
  return {
    payeeName: getInputValue('paymentPayeeName').trim(),
    bankName: getInputValue('paymentBankName').trim(),
    branchName: getInputValue('paymentBranchName').trim(),
    accountType: getInputValue('paymentAccountType') as PaymentAccountType | '',
    accountNumber: getInputValue('paymentAccountNumber').trim(),
    accountHolderKana: getInputValue('paymentAccountHolderKana').trim(),
    amountJpy: getInputValue('paymentAmountJpy').trim(),
    invoiceNumber: getInputValue('paymentInvoiceNumber').trim(),
  };
}

const HALF_WIDTH_KANA_RE = /^[\uFF65-\uFF9F\u0020]+$/;
const ACCOUNT_NUMBER_RE = /^\d{7}$/;
const AMOUNT_RE = /^\d+$/;

function validatePaymentFormData(data: PaymentFormData): boolean {
  if (!data.payeeName || !data.bankName || !data.branchName || !data.accountType) {
    alert('Please fill in all required fields.');
    return false;
  }
  if (!ACCOUNT_NUMBER_RE.test(data.accountNumber)) {
    alert('Account number must be exactly 7 digits.');
    return false;
  }
  if (!data.accountHolderKana || data.accountHolderKana.length > 30 || !HALF_WIDTH_KANA_RE.test(data.accountHolderKana)) {
    alert('Account holder name must be half-width kana (max 30 characters).');
    return false;
  }
  if (!AMOUNT_RE.test(data.amountJpy)) {
    alert('Amount must be half-width digits (JPY).');
    return false;
  }
  return true;
}

function renderPaymentRow(item: PaymentRequestItem): HTMLElement {
  const row = document.createElement('div');
  row.className = 'payment-row';

  const main = document.createElement('div');
  main.className = 'payment-row-main';
  const payeeEl = document.createElement('p');
  payeeEl.className = 'payment-row-payee';
  payeeEl.textContent = String(item.payload.payeeName || '-');
  const amountEl = document.createElement('p');
  amountEl.className = 'payment-row-amount';
  amountEl.textContent = formatJpy(item.payload.amountJpy);
  main.appendChild(payeeEl);
  main.appendChild(amountEl);

  const meta = document.createElement('div');
  meta.className = 'payment-row-meta';
  const badge = document.createElement('span');
  badge.className = `status-badge status-${item.status}`;
  badge.textContent = getPaymentStatusLabel(item.status);
  meta.appendChild(badge);

  const actions = document.createElement('div');
  actions.className = 'payment-row-actions';
  const viewLink = document.createElement('a');
  viewLink.className = 'change-request-link';
  viewLink.href = `#/monthly-payment-requests/status?id=${item.id}`;
  viewLink.textContent = 'View';
  actions.appendChild(viewLink);
  if (item.editable) {
    const editLink = document.createElement('a');
    editLink.className = 'change-request-link';
    editLink.href = `#/monthly-payment-requests/form?id=${item.id}`;
    editLink.textContent = 'Edit';
    actions.appendChild(editLink);
  }

  row.appendChild(main);
  row.appendChild(meta);
  row.appendChild(actions);
  return row;
}

async function loadPaymentList(): Promise<void> {
  const container = document.getElementById('paymentListContainer');
  const freezeNotice = document.getElementById('paymentFreezeNotice');
  const slotsInfo = document.getElementById('paymentSlotsInfo');
  const slotsMeta = document.getElementById('paymentSlotsMeta');
  const newBtn = document.getElementById('btnNewPaymentRequest') as HTMLButtonElement | null;
  if (!container) return;

  container.textContent = '';
  const loading = document.createElement('p');
  loading.className = 'payment-list-loading';
  loading.textContent = 'Loading...';
  container.appendChild(loading);

  try {
    const data = await listPaymentRequests();
    if (freezeNotice) {
      freezeNotice.style.display = data.editable_window ? 'none' : 'block';
    }
    if (slotsInfo) {
      slotsInfo.textContent = `${data.remaining_slots} / ${data.max_active}`;
    }
    if (slotsMeta) {
      slotsMeta.textContent = data.editable_window
        ? 'You can create or edit requests until the 20th of each month (JST).'
        : 'Editing is frozen until the next window (1st–20th JST).';
    }
    const disableNew = data.remaining_slots === 0 || !data.editable_window;
    if (newBtn) {
      newBtn.disabled = disableNew;
      newBtn.classList.toggle('nav-disabled', disableNew);
    }

    container.textContent = '';
    if (data.items.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'payment-list-empty';
      empty.textContent = 'No payment requests yet.';
      container.appendChild(empty);
      return;
    }
    data.items.forEach((item) => {
      container.appendChild(renderPaymentRow(item));
    });
  } catch {
    container.textContent = '';
    const errorEl = document.createElement('p');
    errorEl.className = 'payment-list-empty';
    errorEl.textContent = 'Failed to load payment requests.';
    container.appendChild(errorEl);
  }
}

async function loadPaymentDetail(id: number): Promise<void> {
  const badge = document.getElementById('paymentStatusBadge');
  const meta = document.getElementById('paymentStatusMeta');
  const editBtn = document.getElementById('btnEditPaymentRequest') as HTMLButtonElement | null;
  try {
    const item = await getPaymentRequest(id);
    currentPaymentDetail = item;
    const label = getPaymentStatusLabel(item.status);
    if (badge) {
      badge.textContent = label;
      badge.className = `status-badge status-${item.status || 'pending'}`;
    }
    if (meta) {
      const when = item.submitted_at || item.created_at;
      meta.textContent = when ? `Submitted: ${formatDisplayDate(when)}` : '';
    }
    renderPaymentSummary(payloadToPaymentFormData(item.payload), 'paymentDetail', item.attachment_url, item.attachment_name);
    if (editBtn) {
      editBtn.style.display = item.editable ? '' : 'none';
    }
  } catch {
    currentPaymentDetail = null;
    if (badge) badge.textContent = 'Error';
    if (meta) meta.textContent = 'Failed to load payment request.';
    if (editBtn) editBtn.style.display = 'none';
  }
}

/* ------------------------------------------------------------------ */
/* Routing                                                              */
/* ------------------------------------------------------------------ */

function renderRoute(): void {
  const raw = getRawHash();
  const [rawPath, rawQuery] = raw.split('?');
  const mappedPath = normalizeLegacyAppHash(rawPath);
  if (mappedPath !== rawPath) {
    window.location.hash = rawQuery ? `${mappedPath}?${rawQuery}` : mappedPath;
    return;
  }
  const path = mappedPath;
  const params = new URLSearchParams(rawQuery || '');

  const token = getAccessToken();

  if (!token) {
    showLoginView();
    if (path !== '#/login') {
      window.location.hash = '#/login';
    }
    return;
  }

  showAppShellView();

  if (path === '#/login' || path === '#/') {
    window.location.hash = '#/services';
    return;
  }

  if (path === '#/services') {
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-hub');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (path === '#/profile') {
    void (async () => {
      try {
        await syncMe();
        loadProfile();
        showShellPage('page-profile');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (path === '#/opening-bank-account') {
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-bank');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (path === '#/incorporation') {
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

  if (path === '#/incorporation/form') {
    void (async () => {
      try {
        await syncMe();
        if (hasApplication) {
          if (window.location.hash === '#/incorporation/form') {
            window.location.hash = '#/incorporation';
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

  if (path === '#/incorporation/confirm') {
    if (!pendingFormData) {
      window.location.hash = '#/incorporation/form';
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

  if (path === '#/incorporation/complete') {
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-complete');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (path === '#/incorporation/status') {
    void (async () => {
      try {
        await syncMe();
        if (!hasApplication) {
          window.location.hash = '#/incorporation';
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

  if (path === '#/monthly-payment-requests') {
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-payment-list');
        await loadPaymentList();
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (path === '#/monthly-payment-requests/form') {
    if (suppressPaymentFormReload) {
      suppressPaymentFormReload = false;
      void (async () => {
        try {
          await syncMe();
          updatePaymentFormTitle();
          showShellPage('page-payment-form');
        } catch {
          clearAccessToken();
          window.location.hash = '#/login';
        }
      })();
      return;
    }
    void (async () => {
      try {
        await syncMe();
        const idParam = params.get('id');
        if (idParam) {
          const id = parseInt(idParam, 10);
          const item = await getPaymentRequest(id);
          pendingPaymentEditId = item.id;
          pendingPaymentAttachment = null;
          pendingPaymentExistingAttachmentName = item.attachment_name;
          populatePaymentForm(payloadToPaymentFormData(item.payload));
          showExistingAttachment(item.attachment_name);
        } else {
          pendingPaymentEditId = null;
          pendingPaymentAttachment = null;
          pendingPaymentExistingAttachmentName = null;
          resetPaymentFormFields();
          showExistingAttachment(null);
        }
        updatePaymentFormTitle();
        showShellPage('page-payment-form');
      } catch {
        window.location.hash = '#/monthly-payment-requests';
      }
    })();
    return;
  }

  if (path === '#/monthly-payment-requests/confirm') {
    if (!pendingPaymentData) {
      window.location.hash = '#/monthly-payment-requests/form';
      return;
    }
    void (async () => {
      try {
        await syncMe();
        renderPaymentConfirmPage(pendingPaymentData!, pendingPaymentAttachment, pendingPaymentExistingAttachmentName);
        showShellPage('page-payment-confirm');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (path === '#/monthly-payment-requests/complete') {
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-payment-complete');
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  if (path === '#/monthly-payment-requests/status') {
    const idParam = params.get('id');
    if (!idParam) {
      window.location.hash = '#/monthly-payment-requests';
      return;
    }
    void (async () => {
      try {
        await syncMe();
        showShellPage('page-payment-status');
        await loadPaymentDetail(parseInt(idParam, 10));
      } catch {
        clearAccessToken();
        window.location.hash = '#/login';
      }
    })();
    return;
  }

  window.location.hash = '#/services';
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
    window.location.hash = '#/services';
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
  currentUserEmail = null;
  updateHomeButtons();
  pendingFormData = null;
  resetFormFields();
  generatedFilenames = {};
  pendingPaymentData = null;
  pendingPaymentAttachment = null;
  pendingPaymentEditId = null;
  pendingPaymentExistingAttachmentName = null;
  currentPaymentDetail = null;
  window.location.hash = '#/login';
}

function navToServices(): void {
  window.location.hash = '#/services';
}

function navToIncorporation(): void {
  window.location.hash = '#/incorporation';
}

function navToBankAccount(): void {
  window.location.hash = '#/opening-bank-account';
}

function navToPaymentRequests(): void {
  window.location.hash = '#/monthly-payment-requests';
}

function navToHome(): void {
  window.location.hash = '#/incorporation';
}

function navToForm(): void {
  if (hasApplication) {
    return;
  }
  window.location.hash = '#/incorporation/form';
}

function navToApplications(): void {
  if (!hasApplication) {
    return;
  }
  window.location.hash = '#/incorporation/status';
}

function navToPaymentForm(): void {
  window.location.hash = '#/monthly-payment-requests/form';
}

function navToPaymentList(): void {
  pendingPaymentData = null;
  pendingPaymentAttachment = null;
  pendingPaymentEditId = null;
  pendingPaymentExistingAttachmentName = null;
  window.location.hash = '#/monthly-payment-requests';
}

function navToEditCurrentPayment(): void {
  if (!currentPaymentDetail) return;
  window.location.hash = `#/monthly-payment-requests/form?id=${currentPaymentDetail.id}`;
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
  window.location.hash = '#/incorporation/confirm';
}

function backToFormFromConfirm(): void {
  window.location.hash = '#/incorporation/form';
}

async function confirmAndSubmit(): Promise<void> {
  if (!pendingFormData) {
    window.location.hash = '#/incorporation/form';
    return;
  }

  try {
    await submitApplication(pendingFormData);
    await syncMe();
    window.location.hash = '#/incorporation/complete';
  } catch (error) {
    const message = error instanceof Error ? error.message : t('errorSubmissionFailed');
    alert(`${t('errorSubmissionFailed')}: ${message}`);
  }
}

/* ------------------------------------------------------------------ */
/* Profile                                                              */
/* ------------------------------------------------------------------ */

function loadProfile(): void {
  const emailEl = document.getElementById('profileEmail') as HTMLInputElement | null;
  const msgEl = document.getElementById('profileMessage') as HTMLElement | null;
  if (msgEl) {
    msgEl.style.display = 'none';
    msgEl.textContent = '';
    msgEl.classList.remove('is-error', 'is-success');
  }
  const passIds = ['profileCurrentPassword', 'profileNewPassword', 'profileConfirmPassword'];
  passIds.forEach((id) => {
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (el) el.value = '';
  });
  if (emailEl) emailEl.value = currentUserEmail || '';
}

function showProfileMessage(text: string, success: boolean): void {
  const msgEl = document.getElementById('profileMessage') as HTMLElement | null;
  if (!msgEl) return;
  msgEl.style.display = 'block';
  msgEl.textContent = text;
  msgEl.classList.toggle('is-success', success);
  msgEl.classList.toggle('is-error', !success);
}

async function changePasswordSubmit(): Promise<void> {
  const currentEl = document.getElementById('profileCurrentPassword') as HTMLInputElement | null;
  const newEl = document.getElementById('profileNewPassword') as HTMLInputElement | null;
  const confirmEl = document.getElementById('profileConfirmPassword') as HTMLInputElement | null;

  const current = currentEl?.value ?? '';
  const next = newEl?.value ?? '';
  const confirmValue = confirmEl?.value ?? '';

  if (!current || !next || !confirmValue) {
    showProfileMessage('Please fill in all password fields.', false);
    return;
  }
  if (next.length < 8) {
    showProfileMessage('New password must be at least 8 characters.', false);
    return;
  }
  if (next !== confirmValue) {
    showProfileMessage('New password and confirmation do not match.', false);
    return;
  }

  try {
    await changePassword(current, next);
    showProfileMessage('Password updated successfully.', true);
    if (currentEl) currentEl.value = '';
    if (newEl) newEl.value = '';
    if (confirmEl) confirmEl.value = '';
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Failed to update password';
    showProfileMessage(message, false);
  }
}

/* ------------------------------------------------------------------ */
/* Monthly Payment Requests actions                                     */
/* ------------------------------------------------------------------ */

async function goToPaymentConfirm(): Promise<void> {
  const data = getPaymentFormData();
  if (!validatePaymentFormData(data)) return;

  pendingPaymentData = data;
  const fileInput = document.getElementById('paymentAttachment') as HTMLInputElement | null;
  pendingPaymentAttachment = fileInput?.files?.[0] ?? null;
  window.location.hash = '#/monthly-payment-requests/confirm';
}

function backToPaymentFormFromConfirm(): void {
  suppressPaymentFormReload = true;
  window.location.hash =
    pendingPaymentEditId != null
      ? `#/monthly-payment-requests/form?id=${pendingPaymentEditId}`
      : '#/monthly-payment-requests/form';
}

async function confirmAndSubmitPayment(): Promise<void> {
  if (!pendingPaymentData) {
    window.location.hash = '#/monthly-payment-requests/form';
    return;
  }

  try {
    if (pendingPaymentEditId != null) {
      await updatePaymentRequest(pendingPaymentEditId, pendingPaymentData, pendingPaymentAttachment);
    } else {
      await createPaymentRequest(pendingPaymentData, pendingPaymentAttachment);
    }
    pendingPaymentData = null;
    pendingPaymentAttachment = null;
    pendingPaymentEditId = null;
    pendingPaymentExistingAttachmentName = null;
    window.location.hash = '#/monthly-payment-requests/complete';
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Submission failed';
    alert(`Submission failed: ${message}`);
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
(window as any).navToServices = navToServices;
(window as any).navToIncorporation = navToIncorporation;
(window as any).navToBankAccount = navToBankAccount;
(window as any).navToPaymentRequests = navToPaymentRequests;
(window as any).navToHome = navToHome;
(window as any).navToForm = navToForm;
(window as any).navToApplications = navToApplications;
(window as any).goToConfirm = goToConfirm;
(window as any).confirmAndSubmit = confirmAndSubmit;
(window as any).backToFormFromConfirm = backToFormFromConfirm;
(window as any).addPurpose = addPurpose;
(window as any).changePasswordSubmit = changePasswordSubmit;
(window as any).navToPaymentForm = navToPaymentForm;
(window as any).navToPaymentList = navToPaymentList;
(window as any).navToEditCurrentPayment = navToEditCurrentPayment;
(window as any).goToPaymentConfirm = goToPaymentConfirm;
(window as any).backToPaymentFormFromConfirm = backToPaymentFormFromConfirm;
(window as any).confirmAndSubmitPayment = confirmAndSubmitPayment;
(window as any).downloadWordFile = downloadRegistrationApplicationFile;
(window as any).downloadWordFile2 = downloadArticleOfIncorporationFile;
(window as any).downloadExcelFile = downloadSealRegistrationFile;
(window as any).switchLanguage = switchLanguage;
