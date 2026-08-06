export interface FormData {
  companyName: string;
  presidentName: string;
  presidentNameLocal: string;
  presidentAddress: string;
  presidentAddressLocal: string;
  birthyear: number;
  birthmonth: number;
  birthday: number;
  purpose1: string;
  purpose2: string;
  purpose3: string;
  purpose4: string;
  purpose5: string;
  email: string;
}

export interface ApiResponse {
  message: string;
}

export interface ApiError {
  detail: string;
}

export type ApplicationStatusCode = 'pending' | 'in_review' | 'completed' | 'rejected' | string;

export interface MeResponse {
  email: string;
  has_application: boolean;
  application_status: ApplicationStatusCode | null;
  application_submitted_at: string | null;
}

export interface ApplicationListItem {
  id: number;
  created_at: string;
  submitted_at: string | null;
  status: ApplicationStatusCode;
  payload: Record<string, unknown>;
}

export type Language = 'ja' | 'en';

export type PaymentAccountType = 'checking' | 'ordinary' | 'savings';

export interface PaymentFormData {
  payeeName: string;
  bankName: string;
  branchName: string;
  accountType: PaymentAccountType | '';
  accountNumber: string;
  accountHolderKana: string;
  amountJpy: string;
  invoiceNumber: string;
}

export type PaymentRequestStatusCode = 'pending' | 'paid' | string;

export interface PaymentRequestItem {
  id: number;
  status: PaymentRequestStatusCode;
  payload: Record<string, unknown>;
  attachment_url: string | null;
  attachment_name: string | null;
  created_at: string;
  submitted_at: string | null;
  updated_at: string | null;
  editable: boolean;
}

export interface PaymentListResponse {
  items: PaymentRequestItem[];
  editable_window: boolean;
  remaining_slots: number;
  max_active: number;
}
