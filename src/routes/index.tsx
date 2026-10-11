import { createFileRoute } from '@tanstack/react-router';
import { useMemo, useRef, useState } from 'react';
import { ArrowRight, ArrowLeft, Upload, FileSpreadsheet, FileCode2, ShieldCheck, SlidersHorizontal, CircleHelp, ChevronRight, Search, Download, Check, X, CheckCircle2, LockKeyhole, LoaderCircle, Files, RotateCcw } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Dialog, DialogTrigger, DialogContent, DialogTitle, DialogDescription } from '@/components/ui/dialog';
import { generateLdif, isValidRow, requiredColumns, validBaseDn, type DataRow } from '@/lib/ldif';

export const Route = createFileRoute('/')({
  head: () => ({ meta: [
    { title: 'Condif — Konverter CSV & Excel ke LDIF' },
    { name: 'description', content: 'Konversi CSV dan Excel ke LDIF OpenLDAP. Pilih file, filter pengguna, dan unduh hasil langsung dengan Condif.' },
    { property: 'og:title', content: 'Condif — Konverter CSV & Excel ke LDIF' },
    { property: 'og:description', content: 'Konverter LDIF dengan drag-and-drop, filter data, dan pengaturan LDAP. Semua diproses di perangkat Anda.' },
    { property: 'og:type', content: 'website' }, { name: 'twitter:card', content: 'summary_large_image' },
  ] }), component: Index,
});

function Index() {
  const inputRef = useRef<HTMLInputElement>(null);
  const dragDepth = useRef(0);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [step, setStep] = useState(0);
  const [fileName, setFileName] = useState('');
  const [rows, setRows] = useState<DataRow[]>([]);
  const [sheets, setSheets] = useState<string[]>([]);
  const [sheet, setSheet] = useState('');
  const [fileData, setFileData] = useState<ArrayBuffer | null>(null);
  const [selected, setSelected] = useState<Set<number>>(new Set());
  const [query, setQuery] = useState('');
  const [filterColumn, setFilterColumn] = useState('all');
  const [baseDn, setBaseDn] = useState('dc=sekolah,dc=ac,dc=id');
  const [ou, setOu] = useState('users');
  const [passwordMode, setPasswordMode] = useState('suffix');
  const [customPassword, setCustomPassword] = useState('');
  const [outputName, setOutputName] = useState('');
  const [downloaded, setDownloaded] = useState(false);
  const [hiddenCols, setHiddenCols] = useState<Set<string>>(new Set());
  const allCols = Object.keys(rows[0] || {});
  const cols = allCols.filter(col => !hiddenCols.has(col));
  const visible = useMemo(() => rows.map((row, index) => ({ row, index })).filter(({ row }) => (filterColumn === 'all' ? Object.values(row).join(' ') : row[filterColumn] || '').toLowerCase().includes(query.toLowerCase())), [rows, query, filterColumn]);
  const exportRows = visible.filter(({ index, row }) => selected.has(index) && isValidRow(row)).map(({ row }) => row);
  const uniqueCount = new Set(exportRows.map(row => (row.username || '').trim().toLowerCase())).size;
  const ldif = generateLdif(exportRows, { baseDn, ou, passwordMode, customPassword });
  const configValid = validBaseDn(baseDn) && (passwordMode !== 'custom' || Boolean(customPassword));

  async function parseData(data: ArrayBuffer, sheetName?: string) {
    const XLSX = await import('xlsx');
    const workbook = XLSX.read(data, { type: 'array', codepage: 65001 });
    const name = sheetName || workbook.SheetNames[0];
    if (!name || !workbook.Sheets[name]) throw new Error('File tidak memiliki lembar data.');
    const raw = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[name], { defval: '', raw: false });
    if (!raw.length) throw new Error('File kosong. Tambahkan data pengguna terlebih dahulu.');
    const clean: DataRow[] = raw.map(row => Object.fromEntries(Object.entries(row).map(([key, value]) => [key.replace(/^\uFEFF/, '').trim().toLowerCase(), String(value).trim()])));
    const firstRow = clean[0];
    if (!firstRow) throw new Error('File kosong.');
    const missing = requiredColumns.filter(key => !(key in firstRow));
    if (missing.length) throw new Error(`Kolom wajib tidak ditemukan: ${missing.join(', ')}.`);
    setRows(clean); setSelected(new Set(clean.map((_, i) => i))); setSheets(workbook.SheetNames); setSheet(name); setQuery(''); setStep(1); setDownloaded(false); setHiddenCols(new Set());
  }
  async function loadFile(file?: File) {
    setDragging(false); dragDepth.current = 0;
    if (!file) return;
    setError('');
    if (!/\.(csv|xlsx|xls)$/i.test(file.name)) { setError('Format tidak didukung. Pilih file .csv, .xlsx, atau .xls.'); return; }
    if (file.size > 20 * 1024 * 1024) { setError('Ukuran file melebihi 20 MB. Pilih file yang lebih kecil.'); return; }
    setLoading(true);
    try { const data = await file.arrayBuffer(); await parseData(data); setFileData(data); setFileName(file.name); setOutputName(file.name.replace(/\.[^.]+$/, '') + '.ldif'); }
    catch (e) { setError(e instanceof Error ? e.message : 'File tidak dapat dibaca. Periksa kembali isi file.'); }
    finally { setLoading(false); if (inputRef.current) inputRef.current.value = ''; }
  }
  function downloadSample() {
    download('username,firstname,lastname,email\nahmad.fauzi,Ahmad,Fauzi,ahmad@sekolah.ac.id\nsiti.rahma,Siti,Rahma,siti@sekolah.ac.id\nbudi.santoso,Budi,Santoso,budi@sekolah.ac.id\n', 'contoh_pengguna.csv', 'text/csv;charset=utf-8');
  }
  function download(content: string, name: string, type: string) {
    const url = URL.createObjectURL(new Blob([content], { type })); const link = document.createElement('a'); link.href = url; link.download = name; link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function toggleRow(index: number) { setSelected(prev => { const next = new Set(prev); if (next.has(index)) next.delete(index); else next.add(index); return next; }); setDownloaded(false); }
  function reset() { setStep(0); setRows([]); setFileName(''); setFileData(null); setError(''); setDownloaded(false); setQuery(''); }

  return <>
    <header className="app-header"><div className="header-inner">
      <div className="brand"><span className="brand-symbol" aria-label="n2=n3">n²=n³</span><span className="brand-name">condif<span className="text-success">.</span></span></div>
      <div className="flex items-center gap-6"><span className="header-caption text-xs text-muted-foreground">CSV / Excel <span className="mx-2">→</span> LDIF</span>
        <Dialog><DialogTrigger asChild><Button variant="ghost" size="sm"><CircleHelp /> Bantuan</Button></DialogTrigger><DialogContent className="help-dialog"><DialogTitle>Format file Condif</DialogTitle><DialogDescription>Kolom wajib: username, firstname, lastname, email. CSV dengan pemisah koma atau titik koma serta Excel .xlsx dan .xls didukung hingga 20 MB.</DialogDescription><p className="text-sm text-muted-foreground">Baris tanpa username atau lastname tidak diekspor. Username duplikat hanya diekspor sekali. Hasil menggunakan objectClass inetOrgPerson.</p><Button variant="outline" onClick={downloadSample}><Download /> Unduh contoh CSV</Button></DialogContent></Dialog>
      </div>
    </div></header>
    <main className="workspace">
      <div className="eyebrow"><span className="status-dot" /> WORKSPACE / KONVERTER</div>
      <h1 className="page-title">Data rapi. Konversi mudah<span className="text-success">.</span></h1>
      <p className="page-description">Dari spreadsheet ke direktori LDAP, tanpa ribet.<br className="sm:hidden" /> Unggah, sesuaikan, dan siap digunakan.</p>
      <nav className="workflow" aria-label="Tahapan konversi">{['Unggah file', 'Filter & sesuaikan', 'Konversi & unduh'].map((label, i) => <div key={label} className="contents">{i > 0 && <span className="workflow-line" />}<div className={`workflow-step ${step >= i ? 'active' : ''}`} aria-current={step === i ? 'step' : undefined}><span className="step-number">{step > i ? <Check size={12} /> : `0${i + 1}`}</span>{label}</div></div>)}</nav>
      <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" aria-label="Unggah CSV atau Excel" onChange={e => void loadFile(e.target.files?.[0])} />
      <div className="work-grid">
        <div>
          <section className="upload-panel">
            <div className="panel-heading"><h2 className="panel-title">{step === 0 ? <Upload size={17} /> : step === 1 ? <Files size={17} /> : <FileCode2 size={17} />}{step === 0 ? 'Unggah data pengguna' : step === 1 ? 'Data pengguna' : 'Pratinjau LDIF'}</h2><span className="small-label">{step === 0 ? 'LANGKAH 01' : step === 1 ? `${rows.length} BARIS` : `${uniqueCount} PENGGUNA`}</span></div>
            {error && <div role="alert" className="error-note">{error}</div>}
            {step === 0 ? <>
              <div className={`dropzone ${dragging ? 'dragging' : ''}`} onDragEnter={e => { e.preventDefault(); dragDepth.current++; setDragging(true); }} onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; }} onDragLeave={e => { e.preventDefault(); dragDepth.current--; if (dragDepth.current <= 0) setDragging(false); }} onDrop={e => { e.preventDefault(); if (e.dataTransfer.files.length > 1) { setError('Unggah satu file terlebih dahulu.'); setDragging(false); dragDepth.current = 0; return; } void loadFile(e.dataTransfer.files[0]); }}>
                <div className="file-art" aria-hidden="true"><div className="file-sheet csv"><FileSpreadsheet size={24} /><span>CSV</span></div><div className="file-sheet ldif"><FileCode2 size={24} /><span>LDIF</span></div><div className="file-transfer"><ArrowRight size={13} /></div></div>
                <div className="drop-title">{loading ? 'Membaca file Anda…' : dragging ? 'Lepaskan file di sini' : 'Tarik file Anda ke sini'}</div>
                <div className="drop-subtitle">{dragging ? 'Siap untuk dikonversi' : 'atau pilih langsung dari perangkat Anda'}</div>
                <Button onClick={() => inputRef.current?.click()} disabled={loading}>{loading ? <LoaderCircle className="animate-spin" /> : <Upload />} Pilih file <ChevronRight className="ml-4" /></Button>
                <div className="file-types"><span>.CSV</span><span>.XLSX</span><span>.XLS</span></div>
              </div>
              <div className="upload-bottom"><span className="upload-note"><ShieldCheck size={13} /> Maks. 20 MB · File tidak disimpan</span><Button variant="link" size="sm" onClick={downloadSample}>Unduh contoh CSV <Download /></Button></div>
            </> : step === 1 ? <>
              <div className="flex items-center justify-between gap-2 mb-5"><span className="text-xs text-muted-foreground flex items-center gap-2 min-w-0"><FileSpreadsheet size={16} className="text-success shrink-0" /><span className="truncate">{fileName}</span></span><Button variant="ghost" size="icon" title="Ganti file" aria-label="Ganti file" onClick={() => inputRef.current?.click()}><RotateCcw /></Button></div>
              {sheets.length > 1 && <label className="field"><span className="field-label">Lembar Excel</span><select value={sheet} onChange={async e => { if (!fileData) return; setError(''); try { await parseData(fileData, e.target.value); } catch (err) { setError(err instanceof Error ? err.message : 'Lembar tidak dapat dibaca.'); } }}>{sheets.map(name => <option key={name}>{name}</option>)}</select></label>}
<div className="data-toolbar"><div className="search-wrap"><Search size={15} /><input className="filter-input" aria-label="Cari data pengguna" placeholder="Cari pengguna…" value={query} onChange={e => { setQuery(e.target.value); setDownloaded(false); }} /></div><select className="filter-input max-w-32" aria-label="Kolom filter" value={filterColumn} onChange={e => setFilterColumn(e.target.value)}><option value="all">Semua kolom</option>{allCols.map(col => <option key={col}>{col}</option>)}</select><details className="col-picker"><summary className="filter-input" aria-label="Pilih kolom yang tampil">Kolom tampil{hiddenCols.size ? ` (${cols.length}/${allCols.length})` : ''}</summary><div className="col-picker-menu">{allCols.map(col => <label key={col} className="col-picker-item"><input type="checkbox" checked={!hiddenCols.has(col)} onChange={() => setHiddenCols(prev => { const next = new Set(prev); if (next.has(col)) next.delete(col); else next.add(col); return next; })} /> {col}</label>)}</div></details></div>
              <div className="table-wrap"><table className="data-table"><thead><tr><th><input type="checkbox" aria-label="Pilih semua data yang terlihat" checked={visible.length > 0 && visible.every(({ index }) => selected.has(index))} onChange={e => { const checked = e.target.checked; setSelected(prev => { const next = new Set(prev); visible.forEach(({ index }) => checked ? next.add(index) : next.delete(index)); return next; }); }} /></th>{cols.map(col => <th key={col}>{col}</th>)}<th>Status</th></tr></thead><tbody>{visible.map(({ row, index }) => <tr key={index}><td><input type="checkbox" aria-label={`Pilih ${row.username || 'baris ' + (index + 1)}`} checked={selected.has(index)} onChange={() => toggleRow(index)} /></td>{cols.map(col => <td key={col} className={col === 'username' ? 'font-mono' : 'text-muted-foreground'}>{row[col] || '—'}</td>)}<td><span className={`data-status ${isValidRow(row) ? '' : 'invalid'}`}>{isValidRow(row) ? 'Valid' : 'Tidak lengkap'}</span></td></tr>)}{!visible.length && <tr><td colSpan={cols.length + 2} className="text-center text-muted-foreground">Tidak ada pengguna yang cocok.</td></tr>}</tbody></table></div>
              <p className="field-hint mt-3">{visible.length} dari {rows.length} baris ditampilkan · {rows.filter(row => !isValidRow(row)).length} baris tidak lengkap · {exportRows.length - uniqueCount} duplikat tidak diekspor</p>
              <div className="action-bar"><span className="text-xs text-muted-foreground">{uniqueCount} pengguna dipilih</span><Button disabled={!uniqueCount || !configValid} onClick={() => { setStep(2); setError(''); }}>Pratinjau LDIF <ArrowRight /></Button></div>
            </> : <>
              <div className="summary-strip"><div><strong>{uniqueCount}</strong>Pengguna siap</div><div><strong>{rows.length - uniqueCount}</strong>Tidak diekspor</div><div><strong>LDIF</strong>Format output</div></div>
              <pre className="code-preview" aria-label="Hasil LDIF">{ldif.length > 18000 ? ldif.slice(0, 18000) + '\n# Pratinjau dipotong; unduhan berisi seluruh data.' : ldif}</pre>
              <div className="mt-5"><label className="field"><span className="field-label">Nama file output</span><input value={outputName} onChange={e => setOutputName(e.target.value)} placeholder="pengguna.ldif" /></label></div>
              {downloaded && <p role="status" className="text-success text-xs flex items-center gap-2 mb-4"><CheckCircle2 size={15} /> File LDIF berhasil diunduh.</p>}
              <div className="action-bar"><Button variant="ghost" size="sm" onClick={() => setStep(1)}><ArrowLeft /> Kembali</Button><Button disabled={!configValid || !uniqueCount || !outputName.trim()} onClick={() => { download(ldif, outputName.trim().replace(/(?:\.ldif)?$/i, '.ldif'), 'text/plain;charset=utf-8'); setDownloaded(true); }}><Download /> Unduh LDIF</Button></div>
            </>}
          </section>
          <div className="privacy-band"><LockKeyhole size={13} /><span>Privasi terjaga. Semua proses berlangsung di browser Anda.</span></div>
          {step > 0 && <div className="text-center"><Button variant="ghost" size="sm" onClick={reset}><X /> Mulai konversi baru</Button></div>}
        </div>
        <aside className="settings"><div className="settings-head"><h2 className="panel-title"><SlidersHorizontal size={17} /> Pengaturan LDAP</h2><p className="field-hint">Tujuan direktori untuk data Anda.</p></div><div className="settings-body">
          <label className="field"><span className="field-label">Base DN <span className="text-success">*</span></span><input className="font-mono" value={baseDn} onChange={e => { setBaseDn(e.target.value); setDownloaded(false); }} placeholder="dc=sekolah,dc=ac,dc=id" /><span className="field-hint block">Root direktori LDAP Anda.</span>{!validBaseDn(baseDn) && <span className="text-destructive text-xs block mt-1">Base DN wajib diisi dengan format valid.</span>}</label>
          <label className="field"><span className="field-label">Organizational Unit <span className="text-muted-foreground font-normal text-[10px]">Opsional</span></span><input value={ou} onChange={e => { setOu(e.target.value); setDownloaded(false); }} placeholder="users" /><span className="field-hint block">Kosongkan untuk menggunakan Base DN.</span></label>
          <div className="settings-divider" />
          <label className="field"><span className="field-label">Password pengguna</span><select value={passwordMode} onChange={e => { setPasswordMode(e.target.value); setDownloaded(false); }}><option value="suffix">Username + 123</option><option value="username">Sama dengan username</option><option value="custom">Password custom</option><option value="none">Tanpa password</option></select><span className="field-hint block">{passwordMode === 'suffix' ? 'Contoh: ahmad → ahmad123' : passwordMode === 'username' ? 'Contoh: ahmad → ahmad' : passwordMode === 'none' ? 'Atribut userPassword tidak diekspor.' : 'Password yang sama untuk semua pengguna.'}</span></label>
          {passwordMode === 'custom' && <label className="field"><span className="field-label">Password custom</span><input type="password" autoComplete="new-password" value={customPassword} onChange={e => { setCustomPassword(e.target.value); setDownloaded(false); }} placeholder="Masukkan password" /></label>}
          {passwordMode !== 'none' && step > 0 && <div className="warning-note">Password diekspor sebagai teks biasa. Gunakan hanya untuk provisioning awal, lalu wajibkan perubahan password.</div>}
          <div className="target"><span className="flex items-center gap-2 text-muted-foreground"><FileCode2 size={12} /> PRATINJAU DIREKTORI</span><code>{ou.trim() ? `ou=${ou},` : ''}{baseDn || '…'}</code></div>
        </div></aside>
      </div>
      <footer className="workspace-footer"><span>© {new Date().getUTCFullYear()} Condif <span className="mx-2">/</span> Less manual. More done.</span><span className="flex items-center gap-2"><span className="status-dot" /> Dibuat untuk OpenLDAP</span></footer>
    </main>
  </>;
}
