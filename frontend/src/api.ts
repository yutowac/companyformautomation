import type {
  FormData,
  ApiResponse,
  ApiError,
  MeResponse,
  ApplicationListItem,
  PaymentFormData,
  PaymentRequestItem,
  PaymentListResponse,
} from './types';

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:10000';

const ACCESS_TOKEN_KEY = 'access_token';

export function getAccessToken(): string | null {
  return sessionStorage.getItem(ACCESS_TOKEN_KEY);
}

export function setAccessToken(token: string): void {
  sessionStorage.setItem(ACCESS_TOKEN_KEY, token);
}

export function clearAccessToken(): void {
  sessionStorage.removeItem(ACCESS_TOKEN_KEY);
}

function parseErrorDetail(err: unknown): string {
  if (err && typeof err === 'object' && 'detail' in err) {
    const d = (err as ApiError).detail;
    if (typeof d === 'string') return d;
  }
  return 'Request failed';
}

async function handleResponse<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const error: ApiError = await response.json();
    throw new Error(error.detail || `HTTP error! status: ${response.status}`);
  }
  return response.json();
}

export async function generateRegistrationApplication(data: FormData): Promise<ApiResponse & { filename?: string }> {
  const response = await fetch(`${API_BASE_URL}/generate-registration-application`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return handleResponse<ApiResponse & { filename?: string }>(response);
}

export async function generateArticleOfIncorporation(data: FormData): Promise<ApiResponse & { filename?: string }> {
  const response = await fetch(`${API_BASE_URL}/generate-article-of-incorporation`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return handleResponse<ApiResponse & { filename?: string }>(response);
}

export async function generateSealRegistration(data: FormData): Promise<ApiResponse & { filename?: string }> {
  const response = await fetch(`${API_BASE_URL}/generate-seal-registration`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return handleResponse<ApiResponse & { filename?: string }>(response);
}

export async function downloadRegistrationApplicationFile(filename: string): Promise<Blob> {
  const response = await fetch(`${API_BASE_URL}/download-registration-application?filename=${encodeURIComponent(filename)}`);
  if (!response.ok) {
    throw new Error('Failed to fetch Registration Application file');
  }
  return response.blob();
}

export async function downloadArticleOfIncorporationFile(filename: string): Promise<Blob> {
  const response = await fetch(`${API_BASE_URL}/download-article-of-incorporation?filename=${encodeURIComponent(filename)}`);
  if (!response.ok) {
    throw new Error('Failed to fetch Article of Incorporation file');
  }
  return response.blob();
}

export async function downloadSealRegistrationFile(filename: string): Promise<Blob> {
  const response = await fetch(`${API_BASE_URL}/download-seal-registration?filename=${encodeURIComponent(filename)}`);
  if (!response.ok) {
    throw new Error('Failed to fetch Seal Registration file');
  }
  return response.blob();
}

function downloadBlob(blob: Blob, filename: string): void {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.URL.revokeObjectURL(url);
}

export async function downloadRegistrationApplication(filename: string): Promise<void> {
  try {
    const blob = await downloadRegistrationApplicationFile(filename);
    downloadBlob(blob, filename);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Download failed';
    alert(`Download failed: ${message}`);
  }
}

export async function downloadArticleOfIncorporation(filename: string): Promise<void> {
  try {
    const blob = await downloadArticleOfIncorporationFile(filename);
    downloadBlob(blob, filename);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Download failed';
    alert(`Download failed: ${message}`);
  }
}

export async function downloadSealRegistration(filename: string): Promise<void> {
  try {
    const blob = await downloadSealRegistrationFile(filename);
    downloadBlob(blob, filename);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Download failed';
    alert(`Download failed: ${message}`);
  }
}

export async function recordToSpreadsheet(data: FormData): Promise<ApiResponse> {
  const response = await fetch(`${API_BASE_URL}/record-to-spreadsheet`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  return handleResponse<ApiResponse>(response);
}

export async function loginRequest(email: string, password: string): Promise<void> {
  const response = await fetch(`${API_BASE_URL}/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: 'Incorrect email or password' }));
    throw new Error(parseErrorDetail(err));
  }
  const data = (await response.json()) as { access_token: string };
  setAccessToken(data.access_token);
}

export async function getMe(): Promise<MeResponse> {
  const token = getAccessToken();
  if (!token) {
    throw new Error('Not authenticated');
  }
  const response = await fetch(`${API_BASE_URL}/me`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: `HTTP ${response.status}` }));
    throw new Error(parseErrorDetail(err));
  }
  return response.json() as Promise<MeResponse>;
}

export async function listApplications(): Promise<ApplicationListItem[]> {
  const token = getAccessToken();
  if (!token) {
    throw new Error('Not authenticated');
  }
  const response = await fetch(`${API_BASE_URL}/applications`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: `HTTP ${response.status}` }));
    throw new Error(parseErrorDetail(err));
  }
  return response.json() as Promise<ApplicationListItem[]>;
}

/** 申請を送信し、即座に200で受理。ファイル生成・Drive・Spreadsheetはバックエンドで非同期実行 */
export async function submitApplication(data: FormData): Promise<{ message: string }> {
  const token = getAccessToken();
  if (!token) {
    throw new Error('You must login.');
  }
  const response = await fetch(`${API_BASE_URL}/submit-application`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`,
    },
    body: JSON.stringify(data),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: `HTTP ${response.status}` }));
    throw new Error(parseErrorDetail(err));
  }
  return response.json();
}

export async function changePassword(currentPassword: string, newPassword: string): Promise<{ message: string }> {
  const token = getAccessToken();
  if (!token) {
    throw new Error('You must login.');
  }
  const response = await fetch(`${API_BASE_URL}/auth/change-password`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`,
    },
    body: JSON.stringify({ current_password: currentPassword, new_password: newPassword }),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: `HTTP ${response.status}` }));
    throw new Error(parseErrorDetail(err));
  }
  return response.json();
}

export async function listPaymentRequests(): Promise<PaymentListResponse> {
  const token = getAccessToken();
  if (!token) {
    throw new Error('Not authenticated');
  }
  const response = await fetch(`${API_BASE_URL}/monthly-payment-requests`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: `HTTP ${response.status}` }));
    throw new Error(parseErrorDetail(err));
  }
  return response.json() as Promise<PaymentListResponse>;
}

export async function getPaymentRequest(id: number): Promise<PaymentRequestItem> {
  const token = getAccessToken();
  if (!token) {
    throw new Error('Not authenticated');
  }
  const response = await fetch(`${API_BASE_URL}/monthly-payment-requests/${id}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: `HTTP ${response.status}` }));
    throw new Error(parseErrorDetail(err));
  }
  return response.json() as Promise<PaymentRequestItem>;
}

function buildPaymentMultipart(data: PaymentFormData, attachment?: File | null): globalThis.FormData {
  const body = new globalThis.FormData();
  body.append('payeeName', data.payeeName);
  body.append('bankName', data.bankName);
  body.append('branchName', data.branchName);
  body.append('accountType', data.accountType);
  body.append('accountNumber', data.accountNumber);
  body.append('accountHolderKana', data.accountHolderKana);
  body.append('amountJpy', data.amountJpy);
  body.append('invoiceNumber', data.invoiceNumber || '');
  if (attachment) {
    body.append('attachment', attachment);
  }
  return body;
}

async function submitPaymentMultipart(
  url: string,
  method: 'POST' | 'PUT',
  data: PaymentFormData,
  attachment?: File | null,
): Promise<PaymentRequestItem> {
  const token = getAccessToken();
  if (!token) {
    throw new Error('You must login.');
  }
  const response = await fetch(`${API_BASE_URL}${url}`, {
    method,
    headers: { Authorization: `Bearer ${token}` },
    body: buildPaymentMultipart(data, attachment),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({ detail: `HTTP ${response.status}` }));
    throw new Error(parseErrorDetail(err));
  }
  return response.json() as Promise<PaymentRequestItem>;
}

export async function createPaymentRequest(
  data: PaymentFormData,
  attachment?: File | null,
): Promise<PaymentRequestItem> {
  return submitPaymentMultipart('/monthly-payment-requests', 'POST', data, attachment);
}

export async function updatePaymentRequest(
  id: number,
  data: PaymentFormData,
  attachment?: File | null,
): Promise<PaymentRequestItem> {
  return submitPaymentMultipart(`/monthly-payment-requests/${id}`, 'PUT', data, attachment);
}












