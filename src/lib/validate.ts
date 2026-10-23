/**
 * Validasi tunggal untuk seluruh alur impor (grid + submit + CSV).
 * Batas di sini WAJIB selaras dengan `api/.../web-import.validator.ts`,
 * biar frontend tidak mengirim data yang pasti ditolak server.
 */
import type { SheetRow } from '../components/ExcelSheet';

export const LIMITS = {
  kata: 100,
  terjemahan: 500,
  penjelasanArti: 1000,
  type: 20,
  contoh: 1000,
  artiContoh: 1000,
  officerNameMin: 2,
  officerNameMax: 100,
  sourceLabel: 200,
  supportName: 200,
  supportAddress: 500,
  supportTitle: 300,
  supportDesc: 1000,
  maxRows: 500,
} as const;

export type Severity = 'error' | 'warn';

export interface Issue {
  row: number; // 1-based, baris pada grid
  field: keyof SheetRow | 'row';
  severity: Severity;
  message: string;
}

// Karakter yang tidak boleh ada di data kamus.
const CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]/;
const ZERO_WIDTH = /[\u200B-\u200F\u202A-\u202E\u2060-\u2064\uFEFF]/;
const EMOJI = /[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}\u{1F1E6}-\u{1F1FF}]/u;
const BAD_TYPE = /[^a-z0-9._-]/;

/**
 * Kode CSV lama (csv/RULES.md) -> kode standar word_classes.
 * Dua sistem kode beredar; yang disimpan/dikirim hanya kode standar
 * (n, v, adj, adv, prep) supaya tidak ada duplikat makna di DB.
 */
export const LEGACY_TYPE_MAP: Record<string, string> = {
  kb: 'n',
  kk: 'v',
  ks: 'adj',
  kt: 'adv',
  kd: 'prep',
};

/** Normalisasi kode kelas kata: lowercase + konversi kode lama ke kode standar. */
export function normalizeType(raw: string): string {
  const code = sanitizeCell(raw).toLowerCase();
  return LEGACY_TYPE_MAP[code] ?? code;
}

/** Buang karakter tak terlihat + normalisasi spasi (dipakai sebelum kirim/simpan). */
export function sanitizeCell(v: string): string {
  return v
    .replace(CONTROL, '')
    .replace(ZERO_WIDTH, '')
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

/** Baris dianggap kosong kalau semua kolom wajib/konten kosong. */
export function isEmptyRow(r: SheetRow): boolean {
  return !r.kata.trim() && !r.terjemahan.trim() && !r.penjelasanArti.trim() && !r.contoh.trim();
}

function hasText(r: SheetRow): boolean {
  return !!(
    r.kata.trim() ||
    r.terjemahan.trim() ||
    r.penjelasanArti.trim() ||
    r.type.trim() ||
    r.contoh.trim() ||
    r.artiContoh.trim()
  );
}

const FIELD_LABEL: Record<string, string> = {
  kata: 'kata',
  terjemahan: 'terjemahan',
  penjelasanArti: 'penjelasan_arti',
  contoh: 'contoh',
  artiContoh: 'arti_contoh',
};

/**
 * Periksa satu baris. Mengembalikan daftar isu (error = blokir submit,
 * warn = boleh lanjut tapi perlu diperiksa).
 */
export function checkRow(row: SheetRow, lineNo: number): Issue[] {
  const issues: Issue[] = [];
  const add = (field: Issue['field'], severity: Severity, message: string) =>
    issues.push({ row: lineNo, field, severity, message });

  if (!hasText(row)) return issues;

  // 1. Karakter terlarang (error) + emoji (warn)
  (Object.keys(FIELD_LABEL) as (keyof SheetRow)[]).forEach((key) => {
    const raw = row[key];
    if (!raw) return;
    if (CONTROL.test(raw) || ZERO_WIDTH.test(raw)) {
      add(key, 'error', `${FIELD_LABEL[key]} mengandung karakter tak terlihat. Hapus lalu ketik ulang.`);
    }
    if (EMOJI.test(raw)) {
      add(key, 'warn', `${FIELD_LABEL[key]} mengandung emoji. Kamus biasanya tanpa emoji.`);
    }
  });

  // 2. Panjang (error, samakan dengan batas API)
  (Object.keys(FIELD_LABEL) as (keyof SheetRow)[]).forEach((key) => {
    const max = LIMITS[key as keyof typeof LIMITS] as number;
    const len = row[key].length;
    if (len > max) {
      add(key, 'error', `${FIELD_LABEL[key]} ${len} karakter, maksimal ${max}.`);
    }
  });

  // 3. Wajib: kata
  if (!row.kata.trim()) {
    add('kata', 'error', 'Kolom kata wajib diisi.');
  }

  // 4. Minimal salah satu: terjemahan atau penjelasan_arti
  if (!row.terjemahan.trim() && !row.penjelasanArti.trim()) {
    add('terjemahan', 'error', 'Isi minimal terjemahan atau penjelasan arti.');
  }

  // 5. Kata = terjemahan (indikasi salah isi)
  if (
    row.kata.trim() &&
    row.terjemahan.trim() &&
    row.kata.trim().toLowerCase() === row.terjemahan.trim().toLowerCase()
  ) {
    add('terjemahan', 'warn', 'Kata sama dengan terjemahannya. Cek lagi.');
  }

  // 6. contoh & arti_contoh harus berpasangan
  const adaContoh = !!row.contoh.trim();
  const adaArtiContoh = !!row.artiContoh.trim();
  if (adaContoh !== adaArtiContoh) {
    add(
      adaContoh ? 'artiContoh' : 'contoh',
      'warn',
      'Contoh dan arti_contoh sebaiknya diisi berpasangan.',
    );
  }

  // 7. Format kode kelas kata
  const type = row.type.trim();
  if (type && (type.length > LIMITS.type || BAD_TYPE.test(type))) {
    add('type', 'warn', 'Kode type tidak dikenal. Pilih dari dropdown.');
  }

  return issues;
}

export interface ValidateResult {
  issues: Issue[];
  /** Baris yang benar-benar akan dikirim (bukan baris kosong). */
  payload: SheetRow[];
  errorCount: number;
  warnCount: number;
}

/** Validasi seluruh grid + deteksi duplikat antar baris. */
export function validateSheet(rows: SheetRow[]): ValidateResult {
  const issues: Issue[] = [];
  const payload: SheetRow[] = [];

  // Duplikat: (kata|terjemahan|penjelasan_arti) identik = error,
  // lemma sama beda arti = warn (multi-makna itu sah).
  const exact = new Map<string, number>();
  const lemma = new Map<string, number>();

  rows.forEach((raw, i) => {
    const lineNo = i + 1;
    const row: SheetRow = {
      kata: sanitizeCell(raw.kata),
      terjemahan: sanitizeCell(raw.terjemahan),
      penjelasanArti: sanitizeCell(raw.penjelasanArti),
      type: normalizeType(raw.type),
      contoh: sanitizeCell(raw.contoh),
      artiContoh: sanitizeCell(raw.artiContoh),
    };
    if (!hasText(row)) return;
    payload.push(row);
    issues.push(...checkRow(row, lineNo));

    const k = row.kata.toLowerCase();
    if (k) {
      const exactKey = `${k}|${row.terjemahan.toLowerCase()}|${row.penjelasanArti.toLowerCase()}`;
      const prevExact = exact.get(exactKey);
      if (prevExact !== undefined) {
        issues.push({
          row: lineNo,
          field: 'kata',
          severity: 'error',
          message: `Duplikat persis dengan baris ${prevExact}. Hapus salah satu.`,
        });
      } else {
        exact.set(exactKey, lineNo);
      }
      const prevLemma = lemma.get(k);
      if (prevLemma !== undefined) {
        issues.push({
          row: lineNo,
          field: 'kata',
          severity: 'warn',
          message: `Lemma "${row.kata}" sudah ada di baris ${prevLemma} (beda arti). Pastikan memang makna berbeda.`,
        });
      } else {
        lemma.set(k, lineNo);
      }
    }
  });

  if (payload.length > LIMITS.maxRows) {
    issues.push({
      row: LIMITS.maxRows + 1,
      field: 'row',
      severity: 'error',
      message: `${payload.length} baris terisi, maksimal ${LIMITS.maxRows} per sekali kirim. Bagi jadi beberapa pengajuan.`,
    });
  }

  const errorCount = issues.filter((x) => x.severity === 'error').length;
  const warnCount = issues.filter((x) => x.severity === 'warn').length;
  return { issues, payload, errorCount, warnCount };
}

/** Peta baris -> pesan untuk badge kuning di grid (error menang atas warn). */
export function issuesByRow(issues: Issue[]): Map<number, string> {
  const map = new Map<number, string>();
  for (const it of issues) {
    const lineNo = it.row;
    const prev = map.get(lineNo);
    if (it.severity === 'error' || prev === undefined) {
      map.set(lineNo, it.message);
    }
  }
  return map;
}
