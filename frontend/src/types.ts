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
