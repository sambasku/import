/**
 * Basis URL API untuk import.sambasku.com.
 *
 * Dev lokal: Vite proxy mengarahkan `/api` ke staging/localhost.
 * Produksi: VITE_API_BASE_URL (https://api.sambasku.com/api/v1).
 * Jika kosong/undefined (mis. dev lokal dengan relative path proxy), fallback ke '/api/v1'.
 */
export const API_BASE_URL = (import.meta.env.VITE_API_BASE_URL ?? '').replace(/\/$/, '') || '/api/v1';
