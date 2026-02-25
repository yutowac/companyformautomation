export interface FormData {
  companyName: string;
  presidentName: string;
  presidentAddress: string;
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

export type Language = 'ja' | 'en';

