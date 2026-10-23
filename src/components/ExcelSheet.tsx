import {
  useState,
  useRef,
  useEffect,
  useCallback,
  useMemo,
  type KeyboardEvent,
  type MouseEvent,
} from 'react';
import { LIMITS, normalizeType, sanitizeCell, validateSheet, type Severity } from '../lib/validate';
import { API_BASE_URL } from '../lib/api';

export interface WordClassOption {
  id: string;
  code: string;
  name: string;
  alias?: string | null;
  parent_id: string | null;
}

export async function fetchWordClasses(): Promise<WordClassOption[]> {
  try {
    const res = await fetch(`${API_BASE_URL}/word-classes`);
    if (!res.ok) return [];
    const json = await res.json();
    const items = json?.data?.items ?? json?.data ?? [];
    return Array.isArray(items) ? items : [];
  } catch {
    return [];
  }
}

export interface SheetRow {
  kata: string;
  terjemahan: string;
  penjelasanArti: string;
  type: string;
  contoh: string;
  artiContoh: string;
}

export const COLUMNS: { key: keyof SheetRow; colLetter: string; label: string; width: number; hint: string }[] = [
  { key: 'kata', colLetter: 'A', label: 'kata *', width: 160, hint: 'Lemma bahasa Sambas (wajib)' },
  { key: 'terjemahan', colLetter: 'B', label: 'terjemahan', width: 170, hint: 'Padanan kata bahasa Indonesia' },
  { key: 'penjelasanArti', colLetter: 'C', label: 'penjelasan_arti', width: 260, hint: 'Definisi atau keterangan makna' },
  { key: 'type', colLetter: 'D', label: 'type', width: 75, hint: 'Kelas kata - pilih dari dropdown (kb, kk, ks, kt, kd, atau kode referensi API)' },
  { key: 'contoh', colLetter: 'E', label: 'contoh', width: 220, hint: 'Contoh kalimat bahasa Sambas' },
  { key: 'artiContoh', colLetter: 'F', label: 'arti_contoh', width: 220, hint: 'Terjemahan contoh kalimat' },
];

export function emptyRow(): SheetRow {
  return { kata: '', terjemahan: '', penjelasanArti: '', type: '', contoh: '', artiContoh: '' };
}

export function parseTsvOrCsv(text: string): string[][] {
  const isTsv = text.includes('\t');
  if (isTsv) {
    return text
      .split(/\r?\n/)
      .filter((line) => line.trim().length > 0)
      .map((line) => line.split('\t'));
  }

  // Parse CSV standar (dengan penanganan quotes)
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
  return rows.filter((r) => r.some((c) => c.trim().length > 0));
}

interface CellPos {
  row: number;
  col: number;
}

interface ContextMenuState {
  x: number;
  y: number;
  row: number;
  col: number;
}

export function ExcelSheet({
  rows,
  onChangeRows,
}: {
  rows: SheetRow[];
  onChangeRows: (rows: SheetRow[]) => void;
}) {
  const [selected, setSelected] = useState<CellPos>({ row: 0, col: 0 });
  const [rangeEnd, setRangeEnd] = useState<CellPos | null>(null);
  const [isMouseDown, setIsMouseDown] = useState(false);
  const [editing, setEditing] = useState<{ row: number; col: number; value: string } | null>(null);
  const [colWidths, setColWidths] = useState<number[]>(COLUMNS.map((c) => c.width));
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [history, setHistory] = useState<SheetRow[][]>([]);
const [copiedRange, setCopiedRange] = useState<{ r1: number; r2: number; c1: number; c2: number } | null>(null);
const [typeDropdown, setTypeDropdown] = useState<{ row: number; col: number; top: number; left: number; width: number } | null>(null);
const [typeHoverIdx, setTypeHoverIdx] = useState(0);
  const [wordClasses, setWordClasses] = useState<WordClassOption[]>([]);

  useEffect(() => {
    fetchWordClasses().then(setWordClasses);
  }, []);

const gridRef = useRef<HTMLDivElement>(null);
const cellInputRef = useRef<HTMLInputElement>(null);
const formulaInputRef = useRef<HTMLInputElement>(null);
const copiedTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const typeDropdownRef = useRef<HTMLDivElement>(null);
  const cellRefs = useRef<Record<string, HTMLTableCellElement | null>>({});

  // Daftar opsi kelas kata untuk dropdown. HANYA kode standar word_classes
  // (n, v, adj, adv, prep, ...) supaya tidak ada duplikat makna. Kode CSV lama
  // (kb, kk, ks, kt, kd) tidak ditampilkan sebagai opsi, tapi di-auto-map saat
  // paste/upload (lihat normalizeType).
  const typeOptions = useMemo(() => {
    const map = new Map<string, string>();
    for (const wc of wordClasses) {
      if (!wc.code) continue;
      const alias = wc.alias ? ` (${wc.alias})` : '';
      map.set(wc.code.toLowerCase(), `${wc.name}${alias}`);
    }
    // Fallback statis bila API word-classes gagal dimuat (offline) - daftar kode
    // standar tetap bisa dipilih, tidak mengunci user.
    if (map.size === 0) {
      const fallback: [string, string][] = [
        ['n', 'Nomina (Kata Benda)'],
        ['v', 'Verba (Kata Kerja)'],
        ['adj', 'Adjektiva (Kata Sifat)'],
        ['adv', 'Adverbia (Kata Keterangan)'],
        ['prep', 'Preposisi (Kata Depan)'],
      ];
      for (const [code, name] of fallback) map.set(code, name);
    }
    return Array.from(map, ([code, name]) => ({ code, name })).sort((a, b) =>
      a.code.localeCompare(b.code),
    );
  }, [wordClasses]);

  // Label tampilan kode (mis. "adj" -> "Adjektiva (Kata Sifat)").
  const typeLabel = useCallback(
    (code: string) => typeOptions.find((o) => o.code === code.toLowerCase())?.name ?? code,
    [typeOptions],
  );

  const pushHistory = useCallback(
    (newRows: SheetRow[]) => {
      setHistory((prev) => [...prev.slice(-30), rows]);
      onChangeRows(newRows);
    },
    [rows, onChangeRows],
  );

  const undo = useCallback(() => {
    if (history.length === 0) return;
    const prev = history[history.length - 1];
    setHistory((h) => h.slice(0, -1));
    onChangeRows(prev);
  }, [history, onChangeRows]);

  // Jamin minimal ada 30 baris seperti Excel kosong
  useEffect(() => {
    if (rows.length < 25) {
      const diff = 25 - rows.length;
      const extra = Array.from({ length: diff }, emptyRow);
      onChangeRows([...rows, ...extra]);
    }
  }, [rows, onChangeRows]);

  // Focus input hanya saat edit DIMULAI (bukan tiap keystroke - kalau tiap
  // keystroke, .select() menandai teks lalu huruf berikutnya menimpanya).
  const editingSigRef = useRef<string | null>(null);
  useEffect(() => {
    if (editing) {
      const sig = `${editing.row}:${editing.col}`;
      if (editingSigRef.current !== sig) {
        editingSigRef.current = sig;
        cellInputRef.current?.focus();
        const el = cellInputRef.current;
        if (el) {
          const end = el.value.length;
          el.setSelectionRange(end, end);
        }
      }
    } else {
      editingSigRef.current = null;
      gridRef.current?.focus();
    }
  }, [editing]);

// Tutup context menu saat klik luar
useEffect(() => {
const handleClose = () => setContextMenu(null);
window.addEventListener('click', handleClose);
return () => window.removeEventListener('click', handleClose);
}, []);

// Tutup dropdown type saat klik luar / scroll / resize
useEffect(() => {
if (!typeDropdown) return;
const close = (e: Event) => {
  if (typeDropdownRef.current?.contains(e.target as Node)) return;
  setTypeDropdown(null);
};
    // Scroll di grid = tutup; scroll di DALAM daftar dropdown jangan tutup
    // (kalau semua scroll ditangkap, daftar tak bisa discroll).
    const closeOnScroll = (e: Event) => {
      if (typeDropdownRef.current?.contains(e.target as Node)) return;
      setTypeDropdown(null);
    };
window.addEventListener('mousedown', close);
window.addEventListener('resize', closeOnScroll);
window.addEventListener('scroll', closeOnScroll, true);
return () => {
  window.removeEventListener('mousedown', close);
  window.removeEventListener('resize', closeOnScroll);
  window.removeEventListener('scroll', closeOnScroll, true);
};
}, [typeDropdown]);

  const getColLetter = (c: number) => COLUMNS[c]?.colLetter ?? 'A';
  const cellAddress = `${getColLetter(selected.col)}${selected.row + 1}`;

  const currentCellValue =
    editing && editing.row === selected.row && editing.col === selected.col
      ? editing.value
      : rows[selected.row]?.[COLUMNS[selected.col]?.key] ?? '';

  const commitEdit = (nextMove?: { rowDelta: number; colDelta: number }) => {
    if (!editing) return;
    const { row, col, value } = editing;
    const key = COLUMNS[col].key;
    const nextRows = rows.map((r, i) => (i === row ? { ...r, [key]: value.trim() } : r));
    pushHistory(nextRows);
    setEditing(null);

    if (nextMove) {
      const nextRow = Math.max(0, Math.min(rows.length - 1, row + nextMove.rowDelta));
      const nextCol = Math.max(0, Math.min(COLUMNS.length - 1, col + nextMove.colDelta));
      setSelected({ row: nextRow, col: nextCol });
      setRangeEnd(null);
    }
  };

  const cancelEdit = () => {
    setEditing(null);
  };

  const startEdit = (row: number, col: number, initialValue?: string) => {
    const existing = rows[row]?.[COLUMNS[col].key] ?? '';
    setEditing({ row, col, value: initialValue !== undefined ? initialValue : existing });
  };

  const openTypeDropdown = useCallback(
    (row: number, col: number, el?: HTMLElement | null, initialSearchChar?: string) => {
      commitEdit();
      setSelected({ row, col });
      setRangeEnd(null);
      const targetEl = el ?? cellRefs.current[`${row}:${col}`];
      const rect = targetEl ? targetEl.getBoundingClientRect() : { top: 100, bottom: 124, left: 100, width: 180 };
      const listHeight = Math.min(320, typeOptions.length * 26 + 8);
      const spaceBelow = window.innerHeight - rect.bottom;
      const top = spaceBelow < listHeight + 8 ? rect.top - listHeight - 2 : rect.bottom + 2;
      setTypeDropdown({ row, col, top: Math.max(4, top), left: rect.left, width: Math.max(rect.width, 180) });

      let targetIdx = -1;
      if (initialSearchChar) {
        const char = initialSearchChar.toLowerCase();
        targetIdx = typeOptions.findIndex(
          (o) => o.code.toLowerCase().startsWith(char) || o.name.toLowerCase().startsWith(char),
        );
      }
      if (targetIdx < 0) {
        targetIdx = typeOptions.findIndex(
          (o) => o.code === (rows[row]?.[COLUMNS[col].key] ?? '').toLowerCase(),
        );
      }
      setTypeHoverIdx(targetIdx >= 0 ? targetIdx : 0);
    },
    [commitEdit, rows, typeOptions],
  );

  const pickType = useCallback(
    (code: string, moveNext?: { rowDelta: number; colDelta: number }) => {
      if (!typeDropdown) return;
      const targetRow = typeDropdown.row;
      const targetCol = typeDropdown.col;
      const key = COLUMNS[targetCol].key;
      const nextRows = rows.map((r, i) => (i === targetRow ? { ...r, [key]: code } : r));
      pushHistory(nextRows);
      setTypeDropdown(null);
      gridRef.current?.focus();

      if (moveNext) {
        const nextRow = Math.max(0, Math.min(rows.length - 1, targetRow + moveNext.rowDelta));
        const nextCol = Math.max(0, Math.min(COLUMNS.length - 1, targetCol + moveNext.colDelta));
        setSelected({ row: nextRow, col: nextCol });
        setRangeEnd(null);
      }
    },
    [typeDropdown, rows, pushHistory],
  );

  // Navigasi Keyboard
  const handleKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    if (editing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit({ rowDelta: e.shiftKey ? -1 : 1, colDelta: 0 });
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit({ rowDelta: 0, colDelta: e.shiftKey ? -1 : 1 });
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
      }
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      undo();
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      // Copy TSV
      e.preventDefault();
      copySelection();
      return;
    }
    // Dropdown kelas kata: Esc tutup, Arrow naik/turun, Enter/Tab/Space pilih
    if (typeDropdown) {
      if (e.key === 'Escape') {
        e.preventDefault();
        setTypeDropdown(null);
        return;
      }
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        setTypeHoverIdx((i) =>
          Math.max(0, Math.min(typeOptions.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1))),
        );
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        const opt = typeOptions[typeHoverIdx];
        if (opt) pickType(opt.code, { rowDelta: e.shiftKey ? -1 : 1, colDelta: 0 });
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        const opt = typeOptions[typeHoverIdx];
        if (opt) pickType(opt.code, { rowDelta: 0, colDelta: e.shiftKey ? -1 : 1 });
        return;
      }
      if (e.key === ' ') {
        e.preventDefault();
        const opt = typeOptions[typeHoverIdx];
        if (opt) pickType(opt.code);
        return;
      }
      // Jika ketik karakter saat dropdown terbuka, cari opsi yang berawalan huruf tersebut
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        const char = e.key.toLowerCase();
        const found = typeOptions.findIndex(
          (o) => o.code.toLowerCase().startsWith(char) || o.name.toLowerCase().startsWith(char),
        );
        if (found >= 0) setTypeHoverIdx(found);
        return;
      }
    }

    const isTypeCol = COLUMNS[selected.col]?.key === 'type';

    // Buka dropdown di kolom type dengan Alt+Down, F4, Space, atau Enter
    if (isTypeCol && (e.altKey && e.key === 'ArrowDown' || e.key === 'F4' || e.key === ' ')) {
      e.preventDefault();
      openTypeDropdown(selected.row, selected.col);
      return;
    }

    // Enter di kolom type saat belum dropdown terbuka: buka dropdown
    if (isTypeCol && e.key === 'Enter') {
      e.preventDefault();
      openTypeDropdown(selected.row, selected.col);
      return;
    }

    // Ketik huruf langsung di kolom type (misal 'v', 'n', 'a', 'p') -> langsung buka dropdown & sorot
    if (isTypeCol && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      openTypeDropdown(selected.row, selected.col, null, e.key);
      return;
    }

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      moveSelect(-1, 0, e.shiftKey);
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      moveSelect(1, 0, e.shiftKey);
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      moveSelect(0, -1, e.shiftKey);
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      moveSelect(0, 1, e.shiftKey);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      moveSelect(0, e.shiftKey ? -1 : 1, false);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      moveSelect(e.shiftKey ? -1 : 1, 0, false);
    } else if (e.key === 'F2') {
      e.preventDefault();
      startEdit(selected.row, selected.col);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      clearSelection();
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      startEdit(selected.row, selected.col, e.key);
    }
  };

  const moveSelect = (rDelta: number, cDelta: number, extendRange: boolean) => {
    if (extendRange) {
      // Anchor = selected (sel awal). Range end bergerak dari situ.
      const anchorRow = selected.row;
      const anchorCol = selected.col;
      const nextRow = Math.max(0, Math.min(rows.length - 1, anchorRow + rDelta));
      const nextCol = Math.max(0, Math.min(COLUMNS.length - 1, anchorCol + cDelta));
      setRangeEnd({ row: nextRow, col: nextCol });
      return;
    }
    const nextRow = Math.max(0, Math.min(rows.length - 1, selected.row + rDelta));
    const nextCol = Math.max(0, Math.min(COLUMNS.length - 1, selected.col + cDelta));
    setSelected({ row: nextRow, col: nextCol });
    setRangeEnd(null);
  };

  // Copy range ke Clipboard dalam format TSV
  const copySelection = () => {
    const rStart = rangeEnd ? Math.min(selected.row, rangeEnd.row) : selected.row;
    const rEnd = rangeEnd ? Math.max(selected.row, rangeEnd.row) : selected.row;
    const cStart = rangeEnd ? Math.min(selected.col, rangeEnd.col) : selected.col;
    const cEnd = rangeEnd ? Math.max(selected.col, rangeEnd.col) : selected.col;

    const lines: string[] = [];
    for (let r = rStart; r <= rEnd; r += 1) {
      const line: string[] = [];
      for (let c = cStart; c <= cEnd; c += 1) {
        line.push(rows[r]?.[COLUMNS[c].key] ?? '');
      }
      lines.push(line.join('\t'));
    }
    navigator.clipboard.writeText(lines.join('\n'));
    // Tandai rentang yang baru saja disalin (marching ants)
    const copyBox = {
      r1: rStart,
      r2: rEnd,
      c1: cStart,
      c2: cEnd,
    };
    setCopiedRange(copyBox);
    if (copiedTimerRef.current) window.clearTimeout(copiedTimerRef.current);
    copiedTimerRef.current = window.setTimeout(() => setCopiedRange(null), 4000);
  };

  // Paste dari clipboard mulai dari sel aktif
  const handlePaste = (e: React.ClipboardEvent) => {
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    if (!text) return;
    const matrix = parseTsvOrCsv(text);
    if (matrix.length === 0) return;

    // Batasi 500 baris (API limit)
    const maxPasteRows = Math.min(matrix.length, LIMITS.maxRows);

    const startRow = selected.row;
    const startCol = selected.col;

    // Pastikan baris cukup (maks 500 total)
    const neededRows = startRow + maxPasteRows;
    let nextRows = [...rows];
    if (neededRows > nextRows.length) {
      const diff = Math.min(neededRows - nextRows.length, LIMITS.maxRows - nextRows.length);
      if (diff > 0) nextRows = [...nextRows, ...Array.from({ length: diff }, emptyRow)];
    }

    for (let r = 0; r < maxPasteRows; r += 1) {
      const targetRowIdx = startRow + r;
      const targetRow = { ...nextRows[targetRowIdx] };
      for (let c = 0; c < matrix[r].length; c += 1) {
        const targetColIdx = startCol + c;
        if (targetColIdx < COLUMNS.length) {
          const key = COLUMNS[targetColIdx].key;
          // Sanitize + normalisasi type sama seperti CSV upload
          if (key === 'type') {
            targetRow[key] = normalizeType(matrix[r][c]);
          } else {
            targetRow[key] = sanitizeCell(matrix[r][c]);
          }
        }
      }
      nextRows[targetRowIdx] = targetRow;
    }

    pushHistory(nextRows);
    setRangeEnd({
      row: startRow + maxPasteRows - 1,
      col: Math.min(COLUMNS.length - 1, startCol + (matrix[0]?.length ?? 1) - 1),
    });
  };

  const clearSelection = () => {
    const rStart = rangeEnd ? Math.min(selected.row, rangeEnd.row) : selected.row;
    const rEnd = rangeEnd ? Math.max(selected.row, rangeEnd.row) : selected.row;
    const cStart = rangeEnd ? Math.min(selected.col, rangeEnd.col) : selected.col;
    const cEnd = rangeEnd ? Math.max(selected.col, rangeEnd.col) : selected.col;

    const nextRows = rows.map((r, ri) => {
      if (ri < rStart || ri > rEnd) return r;
      const copy = { ...r };
      for (let ci = cStart; ci <= cEnd; ci += 1) {
        copy[COLUMNS[ci].key] = '';
      }
      return copy;
    });
    pushHistory(nextRows);
  };

  // Resize kolom
  const handleResizeStart = (colIndex: number, e: MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = colWidths[colIndex];

    const onMouseMove = (moveEvent: globalThis.MouseEvent) => {
      const newWidth = Math.max(50, startW + (moveEvent.clientX - startX));
      setColWidths((w) => w.map((curr, idx) => (idx === colIndex ? newWidth : curr)));
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  const minR = rangeEnd ? Math.min(selected.row, rangeEnd.row) : selected.row;
  const maxR = rangeEnd ? Math.max(selected.row, rangeEnd.row) : selected.row;
  const minC = rangeEnd ? Math.min(selected.col, rangeEnd.col) : selected.col;
  const maxC = rangeEnd ? Math.max(selected.col, rangeEnd.col) : selected.col;
  const isFullRowSelected = Boolean(rangeEnd && minC === 0 && maxC === COLUMNS.length - 1);
  // Berapa baris yang kena aksi (hapus/sisip) dari context menu. Hanya hitung
  // multi-baris kalau seleksinya memang seleksi baris penuh (klik nomor baris),
  // bukan seleksi range sel biasa.
  const multiRowCount = isFullRowSelected ? maxR - minR + 1 : 1;

  // Context Menu Actions
  const handleContextMenu = (e: MouseEvent, row: number, col: number) => {
    e.preventDefault();
    // Kalau baris ini di luar seleksi range yang sudah ada, set ulang seleksi ke baris ini
    const isRowInCurrentRange = rangeEnd && row >= minR && row <= maxR;
    if (!isRowInCurrentRange && (selected.row !== row || rangeEnd !== null)) {
      setSelected({ row, col });
      setRangeEnd(null);
    }
    setContextMenu({ x: e.clientX, y: e.clientY, row, col });
  };

  const insertRowAbove = (idx: number) => {
    const targetIdx = rangeEnd ? minR : idx;
    const count = rangeEnd && isFullRowSelected ? maxR - minR + 1 : 1;
    const newRows = Array.from({ length: count }, emptyRow);
    const next = [...rows.slice(0, targetIdx), ...newRows, ...rows.slice(targetIdx)];
    pushHistory(next);
  };

  const insertRowBelow = (idx: number) => {
    const targetIdx = (rangeEnd ? maxR : idx) + 1;
    const count = rangeEnd && isFullRowSelected ? maxR - minR + 1 : 1;
    const newRows = Array.from({ length: count }, emptyRow);
    const next = [...rows.slice(0, targetIdx), ...newRows, ...rows.slice(targetIdx)];
    pushHistory(next);
  };

  const deleteRow = (idx: number) => {
    if (rows.length <= 1) return;
    let start = rangeEnd ? minR : idx;
    let end = rangeEnd ? maxR : idx;
    if (start > end) [start, end] = [end, start];
    const next = rows.filter((_, i) => i < start || i > end);
    pushHistory(next.length === 0 ? [emptyRow()] : next);
    setSelected({ row: Math.min(start, Math.max(0, (next.length === 0 ? 1 : next.length) - 1)), col: selected.col });
    setRangeEnd(null);
  };

  // Peringatan per baris (ala Excel data validation). Sumber tunggal:
  // lib/validate.ts, biar badge di grid = aturan yang dipakai saat submit.
  // error = merah, warn = kuning. Dua-duanya memakai latar tint Excel.
  const rowIssues = useMemo(() => {
    const { issues } = validateSheet(rows);
    const map = new Map<number, { message: string; severity: Severity }>();
    for (const it of issues) {
      const lineNo = it.row; // 1-based
      const idx = lineNo - 1;
      const prev = map.get(idx);
      if (it.severity === 'error' || prev === undefined) {
        map.set(idx, { message: `Baris ${lineNo}: ${it.message}`, severity: it.severity });
      }
    }
    return map;
  }, [rows]);

  return (
    <div
      ref={gridRef}
      tabIndex={0}
      onKeyDown={handleKeyDown}
      onPaste={handlePaste}
      style={{
        display: 'flex',
        flexDirection: 'column',
        flex: 1,
        height: '100%',
        outline: 'none',
        userSelect: 'none',
        fontFamily: 'Calibri, Aptos, "Segoe UI", sans-serif',
        fontSize: 13,
        border: 'none',
        background: '#fff',
      }}
    >
      {/* 1. EXCEL FORMULA BAR */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          height: 28,
          background: '#f8f9fa',
          borderBottom: '1px solid #d4d4d4',
          padding: '0 4px',
          gap: 6,
        }}
      >
        {/* Name Box (e.g. A1, B3) */}
        <div
          style={{
            width: 58,
            height: 22,
            background: '#fff',
            border: '1px solid #d4d4d4',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 600,
            fontSize: 12,
            color: '#333',
          }}
        >
          {cellAddress}
        </div>

        {/* fx symbol */}
        <div
          style={{
            fontStyle: 'italic',
            fontWeight: 'bold',
            color: '#707070',
            fontSize: 13,
            padding: '0 4px',
            userSelect: 'none',
          }}
        >
          fx
        </div>

        {/* Formula Input */}
        <input
          ref={formulaInputRef}
          type="text"
          value={currentCellValue}
          onChange={(e) => {
            const val = e.target.value;
            const key = COLUMNS[selected.col].key;
            onChangeRows(
              rows.map((r, i) => (i === selected.row ? { ...r, [key]: val } : r)),
            );
          }}
          placeholder="Isi sel aktif..."
          style={{
            flex: 1,
            height: 22,
            border: '1px solid #d4d4d4',
            padding: '2px 4px',
            fontFamily: 'inherit',
            fontSize: 13,
            outline: 'none',
            background: '#fff',
          }}
        />
      </div>

      {/* 2. SPREADSHEET TABLE */}
      <div
        style={{
          overflow: 'auto',
          flex: 1,
          height: '100%',
          background: '#fff',
          position: 'relative',
        }}
        onMouseUp={() => setIsMouseDown(false)}
      >
        <table
          style={{
            borderCollapse: 'collapse',
            tableLayout: 'fixed',
            width: 'max-content',
          }}
        >
          {/* HEADER ROW */}
          <thead>
            <tr style={{ height: 26, background: '#f3f3f3' }}>
              {/* Top-Left Select All Corner */}
              <th
                style={{
                  width: 44,
                  minWidth: 44,
                  position: 'sticky',
                  top: 0,
                  left: 0,
                  zIndex: 3,
                  background: '#e9ecef',
                  borderRight: '1px solid #d4d4d4',
                  borderBottom: '1px solid #d4d4d4',
                  textAlign: 'center',
                  cursor: 'pointer',
                }}
                onClick={() => {
                  setSelected({ row: 0, col: 0 });
                  setRangeEnd({ row: rows.length - 1, col: COLUMNS.length - 1 });
                }}
                title="Pilih Semua (Ctrl+A)"
              >
                ◢
              </th>

              {/* Column Headers (A, B, C...) */}
              {COLUMNS.map((col, cIdx) => {
                const isColActive = selected.col === cIdx || (rangeEnd && cIdx >= minC && cIdx <= maxC);
                return (
                  <th
                    key={col.key}
                    style={{
                      width: colWidths[cIdx],
                      minWidth: 50,
                      position: 'sticky',
                      top: 0,
                      zIndex: 2,
                      background: isColActive ? '#d9ead3' : '#f3f3f3',
                      borderRight: '1px solid #d4d4d4',
                      borderBottom: '1px solid #d4d4d4',
                      padding: '0 4px',
                      fontSize: 12,
                      fontWeight: 600,
                      color: isColActive ? '#107c41' : '#495057',
                      textAlign: 'center',
                    }}
                    title={col.hint}
                  >
                    <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: 4 }}>
                      <span style={{ fontSize: 13 }}>{col.colLetter}</span>
                      <span style={{ fontWeight: 400, color: '#666', fontSize: 11 }}>({col.label})</span>
                    </div>

                    {/* Resize Handle */}
                    <div
                      onMouseDown={(e) => handleResizeStart(cIdx, e)}
                      style={{
                        position: 'absolute',
                        right: 0,
                        top: 0,
                        bottom: 0,
                        width: 5,
                        cursor: 'col-resize',
                        zIndex: 4,
                      }}
                    />
                  </th>
                );
              })}
            </tr>
          </thead>

          {/* TABLE BODY */}
          <tbody>
            {rows.map((row, rIdx) => {
              const isRowActive = selected.row === rIdx || (rangeEnd && rIdx >= minR && rIdx <= maxR);
              const issue = rowIssues.get(rIdx);
              const warn = issue?.message;
              const isErrorRow = issue?.severity === 'error';
              return (
                <tr key={rIdx} style={{ height: 24 }}>
                  {/* Row Number Gutter (1, 2, 3...) */}
                  <td
                    style={{
                      width: 44,
                      minWidth: 44,
                      position: 'sticky',
                      left: 0,
                      zIndex: 1,
                      background: isRowActive
                        ? '#d9ead3'
                        : warn
                          ? isErrorRow
                            ? '#ffe0e6'
                            : '#ffe699'
                          : '#f8f9fa',
                      borderRight: '1px solid #d4d4d4',
                      borderBottom: '1px solid #e9ecef',
                      textAlign: 'center',
                      fontSize: 11,
                      fontWeight: isRowActive ? 700 : 400,
                      color: isRowActive ? '#107c41' : '#6c757d',
                      cursor: 'pointer',
                    }}
              onMouseDown={(e) => {
                // Klik nomor baris = seleksi seluruh baris (ala Excel).
                // Shift+klik = perluas seleksi dari anchor.
                // Drag (mousedown lalu geser) = perluas seleksi.
                e.preventDefault();
                // Hanya tombol kiri yang memulai seleksi. Tanpa ini, klik kanan
                // ikut me-reset seleksi jadi 1 baris sebelum context menu muncul.
                if (e.button !== 0) return;
                setIsMouseDown(true);
                if (e.shiftKey) {
                  // Shift+klik: pertahankan selected (anchor), update rangeEnd
                  setRangeEnd({ row: rIdx, col: COLUMNS.length - 1 });
                } else {
                  // Klik biasa: reset anchor dan clear range
                  setSelected({ row: rIdx, col: 0 });
                  setRangeEnd(null);
                }
              }}
              onMouseEnter={() => {
                // Saat drag (mouse ditekan), perluas seleksi ke baris ini
                if (isMouseDown) {
                  setRangeEnd({ row: rIdx, col: COLUMNS.length - 1 });
                }
              }}
              onContextMenu={(e) => handleContextMenu(e, rIdx, 0)}
              title="Klik untuk pilih baris. Shift+klik atau geser untuk beberapa baris."
            >
              {rIdx + 1}
            </td>

                  {/* Row Cells */}
                  {COLUMNS.map((col, cIdx) => {
                    const isSelectedCell = selected.row === rIdx && selected.col === cIdx;
                    const isInRange = rangeEnd && rIdx >= minR && rIdx <= maxR && cIdx >= minC && cIdx <= maxC;
                    const isEditingCell = editing && editing.row === rIdx && editing.col === cIdx;
                    const val = row[col.key] ?? '';

                    // Cek batas rentang seleksi aktif
                    const isTopEdge = rangeEnd ? rIdx === minR && cIdx >= minC && cIdx <= maxC : isSelectedCell;
                    const isBottomEdge = rangeEnd ? rIdx === maxR && cIdx >= minC && cIdx <= maxC : isSelectedCell;
                    const isLeftEdge = rangeEnd ? cIdx === minC && rIdx >= minR && rIdx <= maxR : isSelectedCell;
                    const isRightEdge = rangeEnd ? cIdx === maxC && rIdx >= minR && rIdx <= maxR : isSelectedCell;

                    // Cek batas rentang yang dicopy (marching ants hanya di sisi terluar kotak!)
                    const isCopyTop = copiedRange && rIdx === copiedRange.r1 && cIdx >= copiedRange.c1 && cIdx <= copiedRange.c2;
                    const isCopyBottom = copiedRange && rIdx === copiedRange.r2 && cIdx >= copiedRange.c1 && cIdx <= copiedRange.c2;
                    const isCopyLeft = copiedRange && cIdx === copiedRange.c1 && rIdx >= copiedRange.r1 && rIdx <= copiedRange.r2;
                    const isCopyRight = copiedRange && cIdx === copiedRange.c2 && rIdx >= copiedRange.r1 && rIdx <= copiedRange.r2;

                    return (
<td
                key={col.key}
                ref={(el) => {
                  cellRefs.current[`${rIdx}:${cIdx}`] = el;
                }}
                onMouseDown={(e) => {
                          // Hanya tombol kiri yang memulai seleksi. Klik kanan jangan
                          // reset seleksi sebelum context menu muncul.
                          if (e.button !== 0) return;
                          setIsMouseDown(true);
                          if (isEditingCell) return;
                          commitEdit();
                          // Shift+klik = perluas rentang dari anchor (sel aktif), ala Excel.
                          if (e.shiftKey) {
                            setRangeEnd({ row: rIdx, col: cIdx });
                          } else {
                            setSelected({ row: rIdx, col: cIdx });
                            setRangeEnd(null);
                          }
                        }}
                        onMouseEnter={() => {
                          if (isMouseDown && (selected.row !== rIdx || selected.col !== cIdx)) {
                            setRangeEnd({ row: rIdx, col: cIdx });
                          }
                        }}
                onDoubleClick={() => (col.key === 'type' ? undefined : startEdit(rIdx, cIdx))}
                onContextMenu={(e) => handleContextMenu(e, rIdx, cIdx)}
                title={
                  col.key === 'type' && val
                    ? typeLabel(val)
                    : warn && cIdx <= 1
                      ? warn
                      : undefined
                }
                style={{
                          width: colWidths[cIdx],
                          maxWidth: colWidths[cIdx],
                          height: 24,
                          padding: isEditingCell ? 0 : '1px 3px',
                          borderRight: '1px solid #e1dfdd',
                          borderBottom: '1px solid #e1dfdd',
                          background: isInRange
                            ? 'rgba(16, 124, 65, 0.12)'
                            : isErrorRow
                              ? '#fde8e8'
                              : warn
                                ? '#fff2cc'
                                : '#fff',
                          // Border seleksi hijau disembunyikan saat sel ini kena
                          // copy (Ctrl+C): outline digambar DI ATAS isi sel, jadi
                          // kalau tetap ada, marching ants tertutup dan user tidak
                          // tahu copy sudah aktif.
                          outline:
                            isSelectedCell && !rangeEnd && !isCopyTop && !isCopyBottom && !isCopyLeft && !isCopyRight
                              ? '2px solid #107c41'
                              : undefined,
                          outlineOffset: -2,
                          boxShadow: rangeEnd && isInRange ? [
                            isTopEdge ? 'inset 0 2px 0 0 #107c41' : '',
                            isBottomEdge ? 'inset 0 -2px 0 0 #107c41' : '',
                            isLeftEdge ? 'inset 2px 0 0 0 #107c41' : '',
                            isRightEdge ? 'inset -2px 0 0 0 #107c41' : '',
                          ].filter(Boolean).join(', ') || undefined : undefined,
                          position: 'relative',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          cursor: 'cell',
                        }}
                      >
                        {/* Indikator Marching Ants saat Copy (hanya garis tepi luar range!) */}
                        {isCopyTop && <div className="marching-ants-top" />}
                        {isCopyBottom && <div className="marching-ants-bottom" />}
                        {isCopyLeft && <div className="marching-ants-left" />}
                        {isCopyRight && <div className="marching-ants-right" />}
                        {isEditingCell ? (
                          <input
                            ref={cellInputRef}
                            type="text"
                            list={col.key === 'type' ? 'word-class-list' : undefined}
                            value={editing.value}
                            onChange={(e) => setEditing({ ...editing, value: e.target.value })}
                            onBlur={() => commitEdit()}
                            style={{
                              width: '100%',
                              height: '100%',
                              border: 'none',
                              outline: 'none',
                              padding: '1px 3px',
                              fontFamily: 'inherit',
                              fontSize: 13,
                              background: '#fff',
                              boxSizing: 'border-box',
                            }}
                          />
                        ) : col.key === 'type' ? (
                          <div
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              height: '100%',
                              width: '100%',
                              padding: '1px 2px',
                              background: 'transparent',
                              cursor: 'pointer',
                            }}
                            onClick={(e) => {
                              e.stopPropagation();
                              const target = e.currentTarget as HTMLElement;
                              openTypeDropdown(rIdx, cIdx, target);
                            }}
                          >
                            <span
                              style={{
                                flex: 1,
                                color: val ? '#1f1f1f' : '#aaa',
                                textOverflow: 'ellipsis',
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                              }}
                            >
                              {val}
                            </span>
                            <span
                              style={{
                                color: '#707070',
                                fontSize: 10,
                                marginLeft: 4,
                                flexShrink: 0,
                              }}
                            >
                              ▼
                            </span>
                          </div>
                        ) : (
                          <>
                            <span
                              style={{
                                color: val ? '#1f1f1f' : '#aaa',
                                textDecoration: warn && cIdx <= 1 && !isErrorRow
                              ? 'underline wavy #bf8f00'
                              : undefined,
                              }}
                            >
                              {val}
                            </span>
                            {warn && cIdx <= 1 && (
                              <span
                                style={{
                                  position: 'absolute',
                                  right: 1,
                                  top: 1,
                                  width: 0,
                                  height: 0,
                                  borderLeft: '5px solid transparent',
                                  borderRight: '5px solid transparent',
                                  borderBottom: isErrorRow
                              ? '9px solid #d93025'
                              : warn && cIdx <= 1
                                ? '9px solid #bf8f00'
                                : '9px solid #bf8f00',
                                  cursor: 'help',
                                }}
                              />
                            )}
                            {/* Excel Fill Handle (kotak hijau di sudut kanan bawah keseluruhan seleksi) */}
                            {((isSelectedCell && !rangeEnd) || (rangeEnd && rIdx === maxR && cIdx === maxC)) && (
                              <div
                                style={{
                                  position: 'absolute',
                                  right: 0,
                                  bottom: 0,
                                  width: 6,
                                  height: 6,
                                  background: '#107c41',
                                  cursor: 'crosshair',
                                  zIndex: 4,
                                }}
                              />
                            )}
                          </>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
      </table>
    </div>

    {/* Custom Dropdown panel untuk kolom type (kelas kata) */}
    {typeDropdown && (
      <div
        ref={typeDropdownRef}
        style={{
          position: 'fixed',
          top: typeDropdown.top,
          left: typeDropdown.left,
          width: typeDropdown.width,
          maxHeight: 320,
          background: '#fff',
          border: '1px solid #b0b0b0',
          borderRadius: 3,
          boxShadow: '0 4px 12px rgba(0,0,0,0.24)',
          zIndex: 10000,
          overflow: 'auto',
        }}
      >
        <div
          style={{
            maxHeight: 304,
            overflowY: 'auto',
            padding: 4,
          }}
        >
          {typeOptions.map((opt, idx) => {
            const isCurrent =
              rows[typeDropdown.row]?.[COLUMNS[typeDropdown.col].key]?.toLowerCase() === opt.code;
            const isHover = idx === typeHoverIdx;
            return (
              <div
                key={opt.code}
                ref={(node) => {
                  if (isHover && node) {
                    node.scrollIntoView({ block: 'nearest' });
                  }
                }}
                onMouseDown={(e) => {
                  // mousedown (bukan click) agar tidak trigger close-outer sebelum pick
                  e.preventDefault();
                  e.stopPropagation();
                  pickType(opt.code);
                }}
                onMouseEnter={() => setTypeHoverIdx(idx)}
                style={{
                  padding: '4px 8px',
                  background: isCurrent
                    ? 'rgba(16, 124, 65, 0.18)'
                    : isHover
                      ? '#f0f0f0'
                      : '#fff',
                  color: isCurrent ? '#107c41' : '#1f1f1f',
                  fontSize: 12,
                  fontWeight: isCurrent ? 700 : 400,
                  cursor: 'pointer',
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                }}
                title={opt.name}
              >
                <span style={{ fontFamily: 'monospace', marginRight: 6 }}>{opt.code}</span>
                <span style={{ color: '#666' }}>{opt.name}</span>
              </div>
            );
          })}
        </div>
      </div>
    )}

      {/* 3. EXCEL BOTTOM BAR (Sheet Tabs & Status) */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          height: 28,
          background: '#f8f9fa',
          borderTop: '1px solid #d4d4d4',
          padding: '0 8px',
          fontSize: 12,
        }}
      >
        {/* Sheet Tabs */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
          <div
            style={{
              padding: '2px 14px',
              background: '#fff',
              borderTop: '2px solid #107c41',
              borderRight: '1px solid #d4d4d4',
              borderLeft: '1px solid #d4d4d4',
              fontWeight: 600,
              color: '#107c41',
              cursor: 'pointer',
            }}
          >
            Lemma Sambas (Sheet1)
          </div>
          <button
            type="button"
            onClick={() => {
              const countToAdd = Math.max(1, rows.length);
              onChangeRows([...rows, ...Array.from({ length: countToAdd }, emptyRow)]);
            }}
            style={{
              border: '1px solid #d4d4d4',
              background: '#fff',
              fontSize: 11,
              fontWeight: 500,
              color: '#333',
              cursor: 'pointer',
              padding: '2px 8px',
              borderRadius: 2,
              display: 'inline-flex',
              alignItems: 'center',
              gap: 4,
            }}
            title={`Gandakan baris: tambah ${rows.length} baris kosong lagi`}
          >
            <span style={{ fontWeight: 'bold', color: '#107c41', fontSize: 13 }}>+</span>
            <span>Tambah {rows.length} Baris</span>
          </button>
        </div>

        {/* Status Bar */}
        <div style={{ display: 'flex', gap: 16, color: '#555', alignItems: 'center' }}>
          <span>{editing ? 'EDIT' : 'SIAP (READY)'}</span>
          <span>
            {rows.filter((r) => r.kata.trim().length > 0).length} lemma terisi dari {rows.length} baris
          </span>
          {rowIssues.size > 0 ? (
            <>
              <span style={{ color: '#d93025', fontWeight: 600 }}>
                ✕ {Array.from(rowIssues.values()).filter((x) => x.severity === 'error').length} error
              </span>
              <span style={{ color: '#bf8f00', fontWeight: 600 }}>
                ⚠ {Array.from(rowIssues.values()).filter((x) => x.severity === 'warn').length} peringatan
              </span>
            </>
          ) : null}
        </div>
      </div>

      {/* Opsi kelas kata (referensi GET /api/v1/word-classes) - HANYA kode standar */}
      <datalist id="word-class-list">
        {typeOptions.map((opt) => (
          <option key={opt.code} value={opt.code} label={opt.name} />
        ))}
      </datalist>

      {/* CONTEXT MENU (Klik Kanan) */}
      {contextMenu && (
        <div
          style={{
            position: 'fixed',
            top: contextMenu.y,
            left: contextMenu.x,
            background: '#fff',
            border: '1px solid #ccc',
            boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
            zIndex: 9999,
            padding: '4px 0',
            fontSize: 12,
            minWidth: 160,
          }}
          onClick={(e) => e.stopPropagation()}
        >
          <div
            onClick={() => {
              copySelection();
              setContextMenu(null);
            }}
            style={contextMenuItemStyle}
          >
            Salin (Copy)
          </div>
          <div
            onClick={() => {
              clearSelection();
              setContextMenu(null);
            }}
            style={contextMenuItemStyle}
          >
            Bersihkan Sel (Clear)
          </div>
          <div style={{ height: 1, background: '#eee', margin: '4px 0' }} />
          <div
            onClick={() => {
              insertRowAbove(contextMenu.row);
              setContextMenu(null);
            }}
            style={contextMenuItemStyle}
          >
            {multiRowCount > 1
              ? `Sisipkan ${multiRowCount} Baris di Atas`
              : 'Sisipkan Baris di Atas'}
          </div>
          <div
            onClick={() => {
              insertRowBelow(contextMenu.row);
              setContextMenu(null);
            }}
            style={contextMenuItemStyle}
          >
            {multiRowCount > 1
              ? `Sisipkan ${multiRowCount} Baris di Bawah`
              : 'Sisipkan Baris di Bawah'}
          </div>
          <div
            onClick={() => {
              deleteRow(contextMenu.row);
              setContextMenu(null);
            }}
            style={{ ...contextMenuItemStyle, color: '#d93025' }}
          >
            {multiRowCount > 1 ? `Hapus ${multiRowCount} Baris` : 'Hapus Baris'}
          </div>
        </div>
      )}
    </div>
  );
}

const contextMenuItemStyle: React.CSSProperties = {
  padding: '6px 14px',
  cursor: 'pointer',
  userSelect: 'none',
};
