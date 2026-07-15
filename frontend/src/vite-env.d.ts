/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_API_BASE_URL?: string;
  readonly VITE_APPLICATION_CHANGE_URL?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
