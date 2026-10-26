/**
 * Basis URL API untuk import.sambasku.com.
 *
 * Dev lokal: Vite proxy mengarahkan `/api` ke staging/localhost.
 * Produksi: VITE_API_BASE_URL (https://api.sambasku.com/api/v1).
 * Jika kosong/undefined (mis. dev lokal dengan relative path proxy), fallback ke '/api/v1'.
 */
export const API_BASE_URL = (import.meta.env.VITE_API_BASE_URL ?? '').replace(/\/$/, '') || '/api/v1';

/**
 * Cek ketersediaan lemma di kamus (publik, tanpa login).
 * Pola sama dengan mobile `LemmaCheckNotifier`: non-blocking, timeout 3s.
 * - `true`  = lemma ada di DB (200)
 * - `false` = lemma tidak ditemukan (404)
 * - `null`  = error jaringan / timeout / dibatalkan (diam, tidak mengganggu UI)
 *
 * @param lemma  lemma yang akan dicek (case-insensitive di server)
 * @param signal optional AbortSignal dari luar ( untuk membatalkan request usang)
 */
export async function checkLemma(lemma: string, signal?: AbortSignal): Promise<boolean | null> {
  const encoded = encodeURIComponent(lemma.trim());
  if (!encoded) return null;
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 3000);
  try {
    const res = await fetch(`${API_BASE_URL}/words/lemma/${encoded}`, {
      signal: signal ?? controller.signal,
    });
    return res.ok; // 200 = exists, 404 = not found
  } catch {
    return null; // network error / timeout / aborted
  } finally {
    clearTimeout(timeoutId);
  }
}
