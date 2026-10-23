import { useRef, useState, type ChangeEvent, type ClipboardEvent, type CSSProperties } from 'react';

export interface SheetRow {
  kata: string;
  terjemahan: string;
  penjelasanArti: string;
  type: string;
  contoh: string;
  artiContoh: string;
}

export function emptyRow(): SheetRow {
  return { kata: '', terjemahan: '', penjelasanArti: '', type: '', contoh: '', artiContoh: '' };
}

const COLS: { key: keyof SheetRow; label: string; w: number }[] = [
  { key: 'kata', label: 'Kata', w: 140 },
  { key: 'terjemahan', label: 'Terjemahan', w: 150 },
  { key: 'penjelasanArti', label: 'Penjelasan arti', w: 220 },
  { key: 'type', label: 'Tipe', w: 70 },
  { key: 'contoh', label: 'Contoh', w: 180 },
  { key: 'artiContoh', label: 'Arti contoh', w: 170 },
];

/** Parse CSV sederhana (dukung kutip ganda & baris multiline). */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"' && text[i + 1] === '"') {
        field += '"';
        i += 1;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += ch;
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

const HEADER_KEYS: (keyof SheetRow)[] = ['kata', 'terjemahan', 'penjelasanArti', 'type', 'contoh', 'artiContoh'];

/** Deteksi baris header: kalau kolom pertama = "kata" (case-insensitive) buang. */
function mapToRows(raw: string[][]): SheetRow[] {
  const [first, ...rest] = raw;
  const body = first && first[0]?.trim().toLowerCase() === 'kata' ? rest : raw;
  return body
    .map((cells) => {
      const row = emptyRow();
      HEADER_KEYS.forEach((key, i) => {
        row[key] = (cells[i] ?? '').trim();
      });
      return row;
    })
    .filter((r) => r.kata !== '' || r.terjemahan !== '' || r.penjelasanArti !== '');
}

export function rowFingerprint(r: SheetRow): string {
  return `${r.kata.trim().toLowerCase()}|${r.terjemahan.trim().toLowerCase()}|${r.penjelasanArti.trim().toLowerCase()}`;
}

export const VALID_TYPES = ['kb', 'kk', 'ks', 'kt', 'kd'];

export function validateRow(r: SheetRow): string[] {
  const errs: string[] = [];
  if (!r.kata.trim()) errs.push('Kata wajib diisi');
  if (!r.terjemahan.trim() && !r.penjelasanArti.trim()) {
    errs.push('Isi terjemahan atau penjelasan arti');
  }
  if (r.type.trim() && !VALID_TYPES.includes(r.type.trim().toLowerCase())) {
    errs.push(`Tipe "${r.type}" tidak valid (${VALID_TYPES.join('/')})`);
  }
  return errs;
}

export function ImportSheet({
  rows,
  onAddRow,
  onRowChange,
  onRemoveRow,
}: {
  rows: SheetRow[];
  onAddRow: (row: SheetRow) => void;
  onRowChange: (index: number, updates: Partial<SheetRow>) => void;
  onRemoveRow: (index: number) => void;
}) {
  const [active, setActive] = useState({ row: 0, col: 0 });
  const fileRef = useRef<HTMLInputElement>(null);

  const dupCounts = new Map<string, number>();
  rows.forEach((r) => {
    if (!r.kata.trim()) return;
    const fp = rowFingerprint(r);
    dupCounts.set(fp, (dupCounts.get(fp) ?? 0) + 1);
  });

  const importText = (text: string) => {
    const parsed = mapToRows(parseCsv(text));
    parsed.forEach((r) => onAddRow(r));
  };

  const onPaste = (e: ClipboardEvent) => {
    const text = e.clipboardData.getData('text/plain');
    if (!text.includes('\n') && !text.includes(',')) return;
    const parsed = mapToRows(parseCsv(text));
    if (parsed.length === 0) return;
    e.preventDefault();
    // Satu baris ke baris aktif yang masih kosong = isi baris itu (bukan tambah baru)
    const target = rows[active.row];
    if (parsed.length === 1 && target && !target.kata && !target.terjemahan && !target.penjelasanArti) {
      onRowChange(active.row, parsed[0]);
      return;
    }
    parsed.forEach((r) => onAddRow(r));
  };

  const onFile = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    file.text().then(importText);
    e.target.value = '';
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 8, flexWrap: 'wrap' }}>
        <button type="button" onClick={() => onAddRow(emptyRow())}>
          + Baris baru
        </button>
        <button type="button" onClick={() => fileRef.current?.click()}>
          Upload CSV
        </button>
        <input ref={fileRef} type="file" accept=".csv,text/csv" onChange={onFile} hidden />
        <span style={{ color: '#666', fontSize: 13, alignSelf: 'center' }}>
          Tempel (Ctrl+V) dari Excel / Google Sheets - kolom: kata, terjemahan, penjelasan arti, type, contoh, arti contoh
        </span>
      </div>
      <div style={{ overflow: 'auto', border: '1px solid #d9d9d9', maxHeight: '60vh' }}>
        <table style={{ borderCollapse: 'collapse', width: '100%', minWidth: 960 }}>
          <thead>
            <tr>
              <th style={thStyle(40)}>#</th>
              {COLS.map((c) => (
                <th key={c.key} style={{ ...thStyle(c.w), textAlign: 'left' }}>
                  {c.label}
                </th>
              ))}
              <th style={thStyle(44)}>Aksi</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={8} style={{ padding: 16, color: '#999' }}>
                  Belum ada baris. Klik "+ Baris baru", upload CSV, atau tempel dari spreadsheet.
                </td>
              </tr>
            ) : null}
            {rows.map((r, ri) => {
              const errs = validateRow(r);
              const dup = (dupCounts.get(rowFingerprint(r)) ?? 0) > 1;
              const hasIssue = errs.length > 0 || dup;
              const hint = [...errs, dup ? 'Baris duplikat di file ini' : ''].filter(Boolean).join('; ');
              return (
                <tr key={ri} style={{ background: hasIssue ? '#fff7e6' : ri % 2 ? '#fafafa' : '#fff' }} title={hint}>
                  <td style={tdStyle(40, { textAlign: 'center', color: '#666' })}>{ri + 1}</td>
                  {COLS.map((c, ci) => (
                    <td
                      key={c.key}
                      style={{
                        ...tdStyle(c.w),
                        boxShadow:
                          active.row === ri && active.col === ci
                            ? 'inset 0 0 0 2px #1677ff'
                            : undefined,
                      }}
                      onMouseDown={() => setActive({ row: ri, col: ci })}
                    >
                      <input
                        value={r[c.key]}
                        onChange={(e) => onRowChange(ri, { [c.key]: e.target.value })}
                        onPaste={ci === 0 ? onPaste : undefined}
                        style={{
                          width: '100%',
                          border: 'none',
                          background: 'transparent',
                          font: 'inherit',
                          padding: 2,
                          outline: 'none',
                        }}
                      />
                    </td>
                  ))}
                  <td style={tdStyle(44, { textAlign: 'center' })}>
                    <button
                      type="button"
                      onClick={() => onRemoveRow(ri)}
                      style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: '#cf1322' }}
                      title="Hapus baris"
                    >
                      x
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: 8, fontSize: 13, color: '#666' }}>
        {rows.length} baris
        {rows.some((r) => validateRow(r).length > 0) && (
          <span style={{ color: '#d46b08' }}>
            {' '}
            - ada baris belum valid (kata wajib, terjemahan/penjelasan minimal satu)
          </span>
        )}
      </div>
    </div>
  );
}

function thStyle(w: number): CSSProperties {
  return {
    width: w,
    padding: '6px 8px',
    background: '#f0f0f0',
    border: '1px solid #d9d9d9',
    fontWeight: 600,
    fontSize: 13,
    whiteSpace: 'nowrap',
  };
}

function tdStyle(w: number, extra?: CSSProperties): CSSProperties {
  return {
    width: w,
    padding: '4px 6px',
    border: '1px solid #eee',
    fontSize: 13,
    ...extra,
  };
}
