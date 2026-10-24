import { useState, useRef, type ChangeEvent } from 'react';
import { ExcelSheet, emptyRow, parseTsvOrCsv, type SheetRow } from './components/ExcelSheet';
import { LIMITS, normalizeType, sanitizeCell, validateSheet } from './lib/validate';
import { API_BASE_URL } from './lib/api';

interface FormData {
  officerName: string;
  sourceLabel: string;
  supportName: string;
  supportType: string;
  supportAddress: string;
  supportTitle: string;
  supportDesc: string;
}

const EMPTY_FORM: FormData = {
  officerName: '',
  sourceLabel: '',
  supportName: '',
  supportType: '',
  supportAddress: '',
  supportTitle: '',
  supportDesc: '',
};

export function App() {
  const [formData, setFormData] = useState<FormData>(EMPTY_FORM);
  const [rows, setRows] = useState<SheetRow[]>(Array.from({ length: 30 }, emptyRow));
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [showMetadata, setShowMetadata] = useState(false);
  const [alertInfo, setAlertInfo] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const setField = (key: keyof FormData) => (e: ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) =>
    setFormData((f) => ({ ...f, [key]: e.target.value }));

  // Upload file CSV langsung mengisi grid
  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    file.text().then((text) => {
      // BOM bisa menyatu ke karakter pertama kolom `kata` bila CSV di-export dari Excel.
      const clean = text.replace(/^\uFEFF/, '');
      const parsed = parseTsvOrCsv(clean);
      if (parsed.length === 0) {
        setAlertInfo({ type: 'error', message: 'File CSV kosong atau tidak ada baris data.' });
        return;
      }

      // Cek apakah baris pertama header
      const [first, ...rest] = parsed;
      const hasHeader = first && first[0]?.trim().toLowerCase() === 'kata';
      let body = hasHeader ? rest : parsed;
      // Buang baris yang seluruh selnya kosong (artifact CSV dari baris kosong di akhir file).
      body = body.filter((cells) => cells.some((c) => (c ?? '').trim().length > 0));
      if (body.length === 0) {
        setAlertInfo({ type: 'error', message: 'Tidak ada baris data yang bisa dimuat dari file.' });
        return;
      }
      if (body.length > LIMITS.maxRows) {
        setAlertInfo({
          type: 'error',
          message: `File berisi ${body.length} baris, melebihi batas ${LIMITS.maxRows}. Bagi jadi beberapa file.`,
        });
        return;
      }

      const newRows: SheetRow[] = body.map((cells) => ({
        kata: sanitizeCell(cells[0]),
        terjemahan: sanitizeCell(cells[1]),
        penjelasanArti: sanitizeCell(cells[2]),
        type: normalizeType(cells[3]),
        contoh: sanitizeCell(cells[4]),
        artiContoh: sanitizeCell(cells[5]),
      }));

      // Pad minimal 30 baris
      const finalRows = newRows.length < 30
        ? [...newRows, ...Array.from({ length: 30 - newRows.length }, emptyRow)]
        : newRows;

      setRows(finalRows);
      setAlertInfo({ type: 'success', message: `${newRows.length} baris berhasil dimuat dari file ${file.name}` });
    });

    e.target.value = '';
  };

  // Unduh CSV saat ini
  const handleDownloadCsv = () => {
    const nonEmpty = rows.filter((r) => r.kata.trim() || r.terjemahan.trim() || r.penjelasanArti.trim());
    const header = 'kata,terjemahan,penjelasan_arti,type,contoh,arti_contoh\n';
    const csvContent =
      header +
      nonEmpty
        .map((r) =>
          [
            `"${r.kata.replace(/"/g, '""')}"`,
            `"${r.terjemahan.replace(/"/g, '""')}"`,
            `"${r.penjelasanArti.replace(/"/g, '""')}"`,
            `"${r.type.replace(/"/g, '""')}"`,
            `"${r.contoh.replace(/"/g, '""')}"`,
            `"${r.artiContoh.replace(/"/g, '""')}"`,
          ].join(','),
        )
        .join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `lemma_sambas_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Kosongkan seluruh isi tabel (mulai lembar kerja baru)
  const handleClearTable = () => {
    if (validWordCount > 0) {
      const ok = window.confirm(
        `Kosongkan lembar kerja? ${validWordCount} lemma yang belum dikirim akan terhapus.`,
      );
      if (!ok) return;
    }
    setRows(Array.from({ length: 30 }, emptyRow));
    setAlertInfo({ type: 'info', message: 'Lembar kerja telah dikosongkan. Siap untuk data baru.' });
  };

  const handleSubmit = async () => {
    setAlertInfo(null);

    const officerName = sanitizeCell(formData.officerName);
    if (officerName.length < LIMITS.officerNameMin || officerName.length > LIMITS.officerNameMax) {
      setAlertInfo({ type: 'error', message: 'Nama petugas wajib diisi (2-100 karakter).' });
      setShowMetadata(true);
      return;
    }

    // Validator bersama (satu aturan dengan badge di grid, aturan sama dengan API).
    const { payload, errorCount, warnCount, issues } = validateSheet(rows);
    if (payload.length === 0) {
      setAlertInfo({ type: 'error', message: 'Tabel masih kosong. Silakan isi setidaknya satu kata.' });
      return;
    }
    if (errorCount > 0) {
      const first = issues.find((x) => x.severity === 'error');
      setAlertInfo({
        type: 'error',
        message: `${errorCount} error. ${first ? `Contoh: baris ${first.row} - ${first.message} ` : ''}Perbaiki dulu sebelum mengirim.`,
      });
      return;
    }
    if (warnCount > 0) {
      const ok = window.confirm(
        `Ada ${warnCount} peringatan yang perlu diperiksa (mis. kata = terjemahan, duplikat lemma, contoh tanpa arti). Lanjut kirim?`,
      );
      if (!ok) return;
    }

    setIsSubmitting(true);
    try {
      const response = await fetch(`${API_BASE_URL}/web-imports`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          officer_name: officerName,
          source_label: sanitizeCell(formData.sourceLabel) || undefined,
          support_name: sanitizeCell(formData.supportName) || undefined,
          support_type: formData.supportType || undefined,
          support_address: sanitizeCell(formData.supportAddress) || undefined,
          support_title: sanitizeCell(formData.supportTitle) || undefined,
          support_desc: sanitizeCell(formData.supportDesc) || undefined,
          rows: payload.map((r) => ({
            kata: r.kata,
            terjemahan: r.terjemahan || null,
            penjelasan_arti: r.penjelasanArti || null,
            type: r.type || null,
            contoh: r.contoh || null,
            arti_contoh: r.artiContoh || null,
          })),
        }),
      });

      const data = await response.json();
      if (response.ok && data.success) {
        setAlertInfo({
          type: 'success',
          message: `Data berhasil dikirim (${data.data.total_rows} baris). Status: Menunggu review admin. ID: ${data.data.id}`,
        });
        setRows(Array.from({ length: 30 }, emptyRow));
        setFormData(EMPTY_FORM);
      } else {
        setAlertInfo({ type: 'error', message: data.error?.message ?? data.message ?? 'Gagal mengirim data.' });
      }
    } catch {
      setAlertInfo({ type: 'error', message: 'Gagal terhubung ke server API. Coba lagi dalam beberapa saat.' });
    } finally {
      setIsSubmitting(false);
    }
  };

  const validWordCount = rows.filter((r) => r.kata.trim().length > 0).length;

  return (
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        height: '100vh',
        fontFamily: 'Calibri, Aptos, "Segoe UI", sans-serif',
        background: '#f3f3f3',
        overflow: 'hidden',
      }}
    >
      {/* 1. EXCEL TITLE BAR (Green Ribbon #107c41) - compact */}
      <div
        style={{
          background: '#107c41',
          color: '#fff',
          height: 40,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '0 12px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.2)',
          userSelect: 'none',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          {/* SambasKu Logo Icon */}
          <img
            src="/favicon-192.png"
            alt="SambasKu"
            style={{
              width: 24,
              height: 24,
              borderRadius: 4,
              display: 'block',
              objectFit: 'contain',
              background: '#fff',
              padding: 1,
            }}
          />
          <div>
            <div style={{ fontWeight: 600, fontSize: 12 }}>SambasKu - Impor Lemma.xlsx</div>
            <div style={{ fontSize: 10, opacity: 0.85 }}>Web Importer Publik (Tanpa Login)</div>
          </div>
        </div>

        {/* Action Buttons in Header - compact */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <button
            type="button"
            onClick={() => setShowMetadata(!showMetadata)}
            style={{
              background: showMetadata ? '#0b5329' : 'rgba(255,255,255,0.18)',
              color: '#fff',
              border: 'none',
              padding: '4px 10px',
              borderRadius: 3,
              cursor: 'pointer',
              fontSize: 11,
              fontWeight: 600,
              display: 'flex',
              alignItems: 'center',
              gap: 5,
            }}
          >
            <span>{showMetadata ? '▼' : '▶'} Data Petugas & Sumber</span>
            {!formData.officerName.trim() && (
              <span style={{ background: '#d93025', color: '#fff', borderRadius: '50%', width: 7, height: 7 }} />
            )}
          </button>

          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            style={{
              background: 'rgba(255,255,255,0.18)',
              color: '#fff',
              border: 'none',
              padding: '4px 10px',
              borderRadius: 3,
              cursor: 'pointer',
              fontSize: 11,
              fontWeight: 500,
            }}
          >
            Buka CSV...
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,text/csv,text/plain"
            onChange={handleFileUpload}
            style={{ display: 'none' }}
          />

          <button
            type="button"
            onClick={handleDownloadCsv}
            disabled={validWordCount === 0}
            style={{
              background: 'rgba(255,255,255,0.18)',
              color: '#fff',
              border: 'none',
              padding: '4px 10px',
              borderRadius: 3,
              cursor: validWordCount === 0 ? 'not-allowed' : 'pointer',
              opacity: validWordCount === 0 ? 0.6 : 1,
              fontSize: 11,
              fontWeight: 500,
            }}
          >
            Unduh CSV
          </button>

          <button
            type="button"
            onClick={handleClearTable}
            disabled={validWordCount === 0}
            style={{
              background: 'rgba(255,255,255,0.18)',
              color: '#fff',
              border: 'none',
              padding: '4px 10px',
              borderRadius: 3,
              cursor: validWordCount === 0 ? 'not-allowed' : 'pointer',
              opacity: validWordCount === 0 ? 0.45 : 1,
              fontSize: 11,
              fontWeight: 500,
            }}
            title="Hapus seluruh isi baris dan mulai lembar kerja baru"
          >
            Kosongkan Lembar
          </button>

          <button
            type="button"
            onClick={handleSubmit}
            disabled={isSubmitting}
            style={{
              background: '#fff',
              color: '#107c41',
              border: 'none',
              padding: '4px 14px',
              borderRadius: 3,
              cursor: isSubmitting ? 'not-allowed' : 'pointer',
              fontSize: 12,
              fontWeight: 700,
              boxShadow: '0 2px 4px rgba(0,0,0,0.15)',
            }}
          >
            {isSubmitting ? 'Mengirim...' : `Kirim (${validWordCount})`}
          </button>
        </div>
      </div>

      {/* ALERT NOTIFICATION */}
      {alertInfo && (
        <div
          style={{
            padding: '8px 16px',
            fontSize: 13,
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            background:
              alertInfo.type === 'success'
                ? '#edf7ed'
                : alertInfo.type === 'info'
                  ? '#e8f4fd'
                  : '#fdeded',
            color:
              alertInfo.type === 'success'
                ? '#1e4620'
                : alertInfo.type === 'info'
                  ? '#0c5460'
                  : '#5f2120',
            borderBottom: `1px solid ${
              alertInfo.type === 'success'
                ? '#c8e6c9'
                : alertInfo.type === 'info'
                  ? '#bee5eb'
                  : '#ffcdd2'
            }`,
          }}
        >
          <span>{alertInfo.message}</span>
          <button
            type="button"
            onClick={() => setAlertInfo(null)}
            style={{
              border: 'none',
              background: 'transparent',
              cursor: 'pointer',
              fontWeight: 'bold',
              color: 'inherit',
            }}
          >
            ✕
          </button>
        </div>
      )}

      {/* 2. COLLAPSIBLE METADATA PANEL (Properties Ribbon) */}
      {showMetadata && (
        <div
          style={{
            background: '#fff',
            borderBottom: '1px solid #d4d4d4',
            padding: '12px 16px',
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
            gap: 12,
            boxShadow: 'inset 0 1px 3px rgba(0,0,0,0.04)',
          }}
        >
          <div>
            <label style={{ fontSize: 11, fontWeight: 700, color: '#333', display: 'block', marginBottom: 2 }}>
              NAMA PETUGAS <span style={{ color: '#d93025' }}>*</span>
            </label>
            <input
              type="text"
              value={formData.officerName}
              onChange={setField('officerName')}
              placeholder="Nama pengumpul data..."
              style={metaInputStyle}
            />
          </div>

          <div>
            <label style={{ fontSize: 11, fontWeight: 700, color: '#555', display: 'block', marginBottom: 2 }}>
              LABEL SUMBER
            </label>
            <input
              type="text"
              value={formData.sourceLabel}
              onChange={setField('sourceLabel')}
              placeholder="Contoh: Kamus Sambas hlm 12"
              style={metaInputStyle}
            />
          </div>

          <div>
            <label style={{ fontSize: 11, fontWeight: 700, color: '#555', display: 'block', marginBottom: 2 }}>
              TIPE SUMBER
            </label>
            <select value={formData.supportType} onChange={setField('supportType')} style={metaInputStyle}>
              <option value="">Pilih Tipe</option>
              <option value="book">Buku (Book)</option>
              <option value="web">Web</option>
              <option value="article">Artikel (Article)</option>
              <option value="other">Lainnya (Other)</option>
            </select>
          </div>

          <div>
            <label style={{ fontSize: 11, fontWeight: 700, color: '#555', display: 'block', marginBottom: 2 }}>
              JUDUL / NAMA SUMBER
            </label>
            <input
              type="text"
              value={formData.supportName}
              onChange={setField('supportName')}
              placeholder="Contoh: Kamus Melayu Sambas-Indonesia"
              style={metaInputStyle}
            />
          </div>

          <div>
            <label style={{ fontSize: 11, fontWeight: 700, color: '#555', display: 'block', marginBottom: 2 }}>
              ALAMAT / URL SUMBER
            </label>
            <input
              type="text"
              value={formData.supportAddress}
              onChange={setField('supportAddress')}
              placeholder="Perpustakaan / URL"
              style={metaInputStyle}
            />
          </div>

          <div>
            <label style={{ fontSize: 11, fontWeight: 700, color: '#555', display: 'block', marginBottom: 2 }}>
              CATATAN PETUGAS
            </label>
            <input
              type="text"
              value={formData.supportDesc}
              onChange={setField('supportDesc')}
              placeholder="Catatan tambahan bila ada..."
              style={metaInputStyle}
            />
          </div>
        </div>
      )}

      {/* 3. EXCEL SPREADSHEET COMPONENT */}
      <div style={{ flex: 1, overflow: 'hidden', padding: 0 }}>
        <ExcelSheet rows={rows} onChangeRows={setRows} />
      </div>
    </div>
  );
}

const metaInputStyle: React.CSSProperties = {
  width: '100%',
  height: 28,
  padding: '0 8px',
  border: '1px solid #d4d4d4',
  borderRadius: 3,
  fontSize: 12,
  fontFamily: 'inherit',
  boxSizing: 'border-box',
  outline: 'none',
};
