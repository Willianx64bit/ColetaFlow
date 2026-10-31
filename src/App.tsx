import { useEffect, useMemo, useRef, useState } from 'react';
import { BrowserMultiFormatReader, IScannerControls } from '@zxing/browser';
import * as XLSX from 'xlsx';
import {
  ArrowDownToLine, Barcode, Camera, Check, FileSpreadsheet, FolderUp,
  Keyboard, Package, RotateCcw, ScanLine, Search, ShieldCheck, Smartphone, X, RefreshCw
} from 'lucide-react';

type ColumnMap = { code: number; quantity: number };
const STORAGE_KEY = 'coletaflow.blind-session.v2';
const normalize = (value: string) => value.normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim();
const cellText = (cell: XLSX.CellObject | undefined) => {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  return cell.w !== undefined ? String(cell.w).trim() : String(cell.v).trim();
};
const hasTerm = (value: string, terms: string[]) => terms.some(term => normalize(value).includes(term));

export default function App() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState('');
  const [sheetName, setSheetName] = useState('');
  const [headerRow, setHeaderRow] = useState(0);
  const [columns, setColumns] = useState<ColumnMap>({ code: 0, quantity: 1 });
  const [counts, setCounts] = useState<Record<string, number>>({});
  const [scanValue, setScanValue] = useState('');
  const [quantityValue, setQuantityValue] = useState('1');
  const [pendingCode, setPendingCode] = useState('');
  const [search, setSearch] = useState('');
  const [notice, setNotice] = useState('Importe o modelo vazio da empresa para começar.');
  const [cameraOpen, setCameraOpen] = useState(false);
  const [cameraUsed, setCameraUsed] = useState(false);
  const [cameraStatus, setCameraStatus] = useState('Aguardando câmera');
  const videoRef = useRef<HTMLVideoElement>(null);
  const scannerRef = useRef<IScannerControls | null>(null);
  const lastCameraScan = useRef<{ code: string; at: number } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const quantityRef = useRef<HTMLInputElement>(null);

  const sheet = workbook && sheetName ? workbook.Sheets[sheetName] : null;
  const range = sheet ? (sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }) : null;
  const headers = useMemo(() => {
    if (!sheet || !range) return [];
    const result: string[] = [];
    for (let c = range.s.c; c <= Math.max(range.e.c, range.s.c + 1); c++) {
      result.push(cellText(sheet[XLSX.utils.encode_cell({ r: headerRow, c })]) || `Coluna ${XLSX.utils.encode_col(c)}`);
    }
    return result;
  }, [sheet, range, headerRow]);

  const entries = Object.entries(counts).filter(([code]) => !search || normalize(code).includes(normalize(search)));
  const totalUnits = Object.values(counts).reduce((sum, qty) => sum + qty, 0);
  const totalCodes = Object.keys(counts).length;

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (!saved) return;
      const session = JSON.parse(saved) as { counts?: Record<string, number>; fileName?: string };
      if (session.counts) setCounts(session.counts);
      if (session.fileName) {
        setFileName(session.fileName);
        setNotice(`Contagem local recuperada. Reimporte “${session.fileName}” para exportar.`);
      }
    } catch { /* ignore invalid local session */ }
  }, []);

  useEffect(() => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ counts, fileName })); } catch { /* storage may be full */ }
  }, [counts, fileName]);

  useEffect(() => {
    if (!cameraOpen || !videoRef.current || pendingCode) return;
    let cancelled = false;
    let controls: IScannerControls | null = null;
    const video = videoRef.current;
    setCameraStatus('Solicitando câmera traseira…');

    const reader = new BrowserMultiFormatReader();
    reader.decodeFromConstraints({
      audio: false,
      video: {
        facingMode: { ideal: 'environment' },
        width: { ideal: 1920 },
        height: { ideal: 1080 },
        frameRate: { ideal: 30 }
      }
    }, video, result => {
      if (!result || cancelled) return;
      const code = result.getText().trim();
      if (!code) return;
      setCameraStatus('Código identificado');
      receiveCode(code, true);
      setCameraOpen(false);
    }).then(activeControls => {
      controls = activeControls;
      if (cancelled) activeControls.stop();
      else {
        scannerRef.current = activeControls;
        setCameraUsed(true);
        setCameraStatus('Câmera ativa · aponte para o código de barras');
      }
    }).catch(() => {
      if (!cancelled) {
        setCameraOpen(false);
        setCameraStatus('Não foi possível abrir. Confira a permissão da câmera.');
        setNotice('Não foi possível abrir a câmera. Autorize o acesso e tente novamente.');
      }
    });

    return () => {
      cancelled = true;
      controls?.stop();
      scannerRef.current?.stop();
      scannerRef.current = null;
      const stream = video.srcObject as MediaStream | null;
      stream?.getTracks().forEach(track => track.stop());
      video.srcObject = null;
    };
  }, [cameraOpen, pendingCode]);

  const loadWorkbook = async (file: File) => {
    try {
      const data = await file.arrayBuffer();
      const parsed = XLSX.read(data, { type: 'array', cellText: true, cellNF: true });
      const firstSheet = parsed.SheetNames[0];
      if (!firstSheet) throw new Error('O arquivo não possui uma aba legível.');
      const ws = parsed.Sheets[firstSheet];
      const wsRange = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } };
      let detectedHeader = wsRange.s.r;
      let bestScore = -1;
      for (let r = wsRange.s.r; r <= Math.min(wsRange.s.r + 12, wsRange.e.r); r++) {
        let score = 0;
        for (let c = wsRange.s.c; c <= Math.max(wsRange.e.c, wsRange.s.c + 1); c++) {
          const value = cellText(ws[XLSX.utils.encode_cell({ r, c })]);
          if (hasTerm(value, ['codigo', 'ean', 'gtin', 'barras', 'barcode', 'quantidade', 'qtd'])) score++;
        }
        if (score > bestScore) { bestScore = score; detectedHeader = r; }
      }
      const rowHeaders: string[] = [];
      for (let c = wsRange.s.c; c <= Math.max(wsRange.e.c, wsRange.s.c + 1); c++) {
        rowHeaders.push(cellText(ws[XLSX.utils.encode_cell({ r: detectedHeader, c })]));
      }
      const findColumn = (terms: string[], fallback: number) => {
        const found = rowHeaders.findIndex(value => hasTerm(value, terms));
        return found >= 0 ? found : Math.min(fallback, Math.max(0, rowHeaders.length - 1));
      };
      const previous = (() => {
        try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') as { fileName?: string; counts?: Record<string, number> }; }
        catch { return {}; }
      })();
      const sameTemplate = previous.fileName === file.name;
      setWorkbook(parsed);
      setFileName(file.name);
      setSheetName(firstSheet);
      setHeaderRow(detectedHeader);
      setColumns({
        code: findColumn(['ean', 'gtin', 'codigo de barras', 'cod barras', 'barcode', 'codigo', 'cod.'], 0),
        quantity: findColumn(['quantidade', 'qtd', 'contagem', 'unidades'], 1)
      });
      if (!sameTemplate) setCounts({});
      setPendingCode('');
      setNotice(sameTemplate ? 'Modelo importado. A contagem anterior deste arquivo foi recuperada.' : 'Modelo vazio importado. Agora é só bipar o código e informar a quantidade.');
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Não foi possível abrir o modelo.');
    }
  };

  const receiveCode = (raw: string, fromCamera = false) => {
    const code = raw.trim();
    if (!code) return;
    if (fromCamera) {
      const last = lastCameraScan.current;
      if (last?.code === code && Date.now() - last.at < 1600) return;
      lastCameraScan.current = { code, at: Date.now() };
    }
    setPendingCode(code);
    setScanValue('');
    setNotice(`Código lido: ${code}. Informe a quantidade para registrar.`);
    requestAnimationFrame(() => quantityRef.current?.focus());
  };

  const confirmCount = () => {
    const code = pendingCode.trim();
    const qty = Number(quantityValue);
    if (!code) { setNotice('Bipe ou digite um código primeiro.'); inputRef.current?.focus(); return; }
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) { setNotice('Informe uma quantidade inteira maior que zero.'); quantityRef.current?.focus(); return; }
    setCounts(previous => ({ ...previous, [code]: (previous[code] ?? 0) + qty }));
    setNotice(`Registrado: código ${code} · quantidade adicionada: ${qty}.`);
    setPendingCode('');
    setQuantityValue('1');
    setScanValue('');
    requestAnimationFrame(() => inputRef.current?.focus());
  };

  const startCamera = () => {
    if (!workbook || pendingCode) return;
    setCameraUsed(true);
    setCameraStatus('Preparando leitor de código de barras…');
    setCameraOpen(true);
  };

  const stopCamera = () => {
    setCameraOpen(false);
    setCameraStatus('Câmera pausada');
  };

  const exportWorkbook = () => {
    if (!workbook || !sheet || !range) return;
    const output = XLSX.read(XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }), { type: 'array', cellStyles: true });
    const outputSheet = output.Sheets[sheetName];
    const hasHeader = Array.from({ length: range.e.c - range.s.c + 1 }, (_, i) => cellText(sheet[XLSX.utils.encode_cell({ r: headerRow, c: range.s.c + i })])).some(Boolean);
    let lastUsedRow = hasHeader ? headerRow : range.s.r - 1;
    for (let r = hasHeader ? headerRow + 1 : range.s.r; r <= range.e.r; r++) {
      const rowHasValue = Array.from({ length: range.e.c - range.s.c + 1 }, (_, i) => cellText(sheet[XLSX.utils.encode_cell({ r, c: range.s.c + i })])).some(Boolean);
      if (rowHasValue) lastUsedRow = r;
    }
    let row = lastUsedRow + 1;
    const startColumn = range.s.c;
    for (const [code, qty] of Object.entries(counts)) {
      const codeCell = XLSX.utils.encode_cell({ r: row, c: startColumn + columns.code });
      const qtyCell = XLSX.utils.encode_cell({ r: row, c: startColumn + columns.quantity });
      outputSheet[codeCell] = { t: 's', v: code, z: '@' };
      outputSheet[qtyCell] = { t: 'n', v: qty };
      row++;
    }
    outputSheet['!ref'] = XLSX.utils.encode_range({
      s: range.s,
      e: { r: Math.max(range.e.r, row - 1), c: Math.max(range.e.c, startColumn + columns.code, startColumn + columns.quantity) }
    });
    const outputName = fileName.replace(/\.(xlsx|xls|csv)$/i, '') + '-coletado.xlsx';
    XLSX.writeFile(output, outputName);
    setNotice(`Planilha preenchida: ${totalCodes} códigos e ${totalUnits} unidades. O arquivo modelo original foi preservado.`);
  };

  const resetSession = () => {
    if (!window.confirm('Limpar toda a contagem desta sessão?')) return;
    setCounts({});
    setPendingCode('');
    setNotice('Contagem reiniciada. O modelo continua carregado.');
  };

  const updateCount = (code: string, delta: number) => {
    setCounts(previous => {
      const next = Math.max(0, (previous[code] ?? 0) + delta);
      const updated = { ...previous };
      if (next === 0) delete updated[code]; else updated[code] = next;
      return updated;
    });
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#" aria-label="ColetaFlow início">
          <span className="brand-mark"><ScanLine size={25} /></span>
          <span><strong>ColetaFlow</strong><small>COLETOR DE CÓDIGOS E QUANTIDADES</small></span>
        </a>
        <div className="topbar-right"><span className="offline-pill"><span /> CONTAGEM LOCAL</span><span className="version">MVP 0.2</span></div>
      </header>

      <section className="hero">
        <div className="hero-copy">
          <span className="eyebrow">COLETA SEM CADASTRO PRÉVIO</span>
          <h1>Bipe o código.<br /><em>Informe a quantidade.</em></h1>
          <p>Use o modelo vazio da empresa. O ColetaFlow registra somente o código de barras e a quantidade contada.</p>
        </div>
        <div className="hero-art"><div className="scan-orbit orbit-one" /><div className="scan-orbit orbit-two" /><Barcode size={76} strokeWidth={1.25} /></div>
      </section>

      <section className="workspace">
        <div className="section-heading"><div><span className="step-label">ETAPA 01</span><h2>Modelo da empresa</h2><p>Importe a planilha vazia que deverá receber os códigos e quantidades.</p></div>
          <label className="primary-button upload-button"><FolderUp size={18} /> Importar modelo<input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /></label>
        </div>
        {workbook && sheet ? (
          <div className="file-card">
            <div className="file-icon"><FileSpreadsheet size={24} /></div>
            <div className="file-info"><strong>{fileName}</strong><span>Modelo carregado · {workbook.SheetNames.length} aba(s) · sem necessidade de cadastro de produtos</span></div>
            <span className="file-status"><Check size={14} /> Pronto</span>
            <button className="icon-button" title="Remover modelo" onClick={() => { setWorkbook(null); setFileName(''); setSheetName(''); setCounts({}); setPendingCode(''); }}><X size={18} /></button>
          </div>
        ) : <label className="dropzone"><input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /><span className="drop-icon"><FolderUp size={25} /></span><strong>Selecione o modelo vazio da empresa</strong><span>Excel (.xlsx, .xls) ou CSV</span></label>}

        {workbook && sheet && (
          <div className="mapping-panel">
            <div className="mapping-title"><div><ShieldCheck size={18} /><strong>Defina onde gravar os dados</strong></div><span>MAPEAMENTO DO MODELO</span></div>
            <div className="mapping-grid">
              <label>Aba<select value={sheetName} onChange={event => { setSheetName(event.target.value); setCounts({}); setPendingCode(''); }}>{workbook.SheetNames.map(name => <option key={name} value={name}>{name}</option>)}</select></label>
              <label>Coluna do código<select value={columns.code} onChange={event => setColumns(v => ({ ...v, code: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
              <label>Coluna da quantidade<select value={columns.quantity} onChange={event => setColumns(v => ({ ...v, quantity: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
            </div>
            <div className="mapping-note">O sistema não procura produtos na planilha. Essas duas colunas serão preenchidas somente na exportação.</div>
          </div>
        )}
      </section>

      <section className="workspace collector-section">
        <div className="section-heading"><div><span className="step-label">ETAPA 02</span><h2>Coletor</h2><p>Bipe qualquer código e informe quantas unidades encontrou.</p></div><span className={workbook ? 'ready-chip' : 'waiting-chip'}>{workbook ? 'MODELO PRONTO' : 'AGUARDANDO MODELO'}</span></div>
        <div className="stats-grid">
          <div className="stat-card"><span>Unidades contadas</span><strong>{totalUnits}</strong><small>Soma das quantidades informadas</small></div>
          <div className="stat-card"><span>Códigos diferentes</span><strong>{totalCodes}</strong><small>Identificadores registrados</small></div>
          <div className="stat-card warning-stat"><span>Aguardando quantidade</span><strong>{pendingCode ? '1' : '0'}</strong><small>{pendingCode ? 'Código lido, falta confirmar' : 'Nenhum código pendente'}</small></div>
        </div>
        <div className="scanner-panel">
          <div className="scanner-heading"><div className="scanner-icon"><Barcode size={22} /></div><div><strong>Bipagem contínua</strong><span>Sem catálogo, descrição ou saldo de estoque</span></div><span className="live-indicator"><i /> AO VIVO</span></div>
          <form className="scan-form" onSubmit={event => { event.preventDefault(); receiveCode(scanValue); }}>
            <div className="scan-input-wrap"><Keyboard size={20} /><input ref={inputRef} value={scanValue} onChange={event => setScanValue(event.target.value)} placeholder={!workbook ? 'Importe o modelo primeiro' : pendingCode ? 'Informe a quantidade do código lido' : 'Bipe ou digite o código...'} disabled={!workbook || !!pendingCode} autoComplete="off" autoFocus /><button type="submit" disabled={!workbook || !scanValue.trim()}><ScanLine size={19} /><span>Ler código</span></button></div>
          </form>
          {pendingCode && <div className="pending-code-card"><div><small>CÓDIGO LIDO</small><strong>{pendingCode}</strong></div><label>Quantidade<input ref={quantityRef} type="number" min="1" step="1" value={quantityValue} onChange={event => setQuantityValue(event.target.value)} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); confirmCount(); } }} /></label><button className="primary-button" onClick={confirmCount}><Check size={17} /> Confirmar</button><button className="icon-button" title="Cancelar código" onClick={() => { setPendingCode(''); setQuantityValue('1'); inputRef.current?.focus(); }}><X size={18} /></button></div>}
          <div className="scanner-actions">
            <button className="secondary-button camera-launch" onClick={startCamera} disabled={!workbook || cameraOpen || !!pendingCode}>
              {cameraUsed ? <RefreshCw size={17} /> : <Camera size={17} />}
              {cameraUsed ? 'Reabrir câmera' : 'Ler pela câmera'}
            </button>
            <span>Leitor USB/Bluetooth funciona como teclado (HID).</span>
          </div>
          {cameraOpen && (
            <div className="camera-box">
              <div className="camera-status"><span className="camera-live-dot" />{cameraStatus}</div>
              <video ref={videoRef} muted playsInline autoPlay />
              <div className="camera-guide" aria-hidden="true"><span /><i /></div>
              <div className="camera-help">Enquadre o código de barras dentro da área central, com boa iluminação e sem inclinar o celular.</div>
              <button className="camera-close" onClick={stopCamera}><X size={18} /> Fechar câmera</button>
            </div>
          )}
          {!cameraOpen && cameraUsed && <div className="camera-reopen-hint">A câmera está pausada. Toque em <strong>Reabrir câmera</strong> para iniciar outra leitura.</div>}
          <div className="scan-feedback">{notice}</div>
        </div>
      </section>

      <section className="workspace results-section">
        <div className="section-heading"><div><span className="step-label">ETAPA 03</span><h2>Resumo da coleta</h2><p>Confira somente os códigos e as quantidades registradas.</p></div><button className="text-button" onClick={resetSession} disabled={!workbook}><RotateCcw size={16} /> Reiniciar contagem</button></div>
        <div className="results-toolbar">
          <div className="blind-mode-note"><ShieldCheck size={17} /> SOMENTE CÓDIGO + QUANTIDADE</div>
          <label className="search-box"><Search size={17} /><input value={search} onChange={event => setSearch(event.target.value)} placeholder="Buscar código registrado" /></label>
        </div>
        <div className="product-list">
          {!workbook ? <div className="empty-state"><Package size={28} /><strong>Aguardando modelo</strong><span>Importe a planilha vazia para iniciar a coleta.</span></div> :
            entries.slice(0,100).map(([code, qty], index) => <div className="product-row" key={code}><div className="product-number">{index + 1}</div><div className="product-details"><strong>{code}</strong><span>CÓDIGO</span></div><div className="quantity-control"><button onClick={() => updateCount(code, -1)} disabled={qty <= 1}>−</button><strong>{qty}</strong><button onClick={() => updateCount(code, 1)}>+</button></div></div>)}
          {workbook && entries.length === 0 && <div className="empty-state"><Search size={24} /><strong>Nenhum código registrado</strong><span>Depois da leitura, os códigos e quantidades aparecerão aqui.</span></div>}
          {entries.length > 100 && <div className="list-footnote">Mostrando os primeiros 100 códigos. Use a busca para localizar outro.</div>}
        </div>
      </section>

      <section className="export-bar">
        <div><span className="step-label">ETAPA 04</span><strong>Exportar modelo preenchido</strong><small>O arquivo original fica intacto; será criada uma cópia com código e quantidade.</small></div>
        <div className="export-actions"><button className="primary-button" onClick={exportWorkbook} disabled={!workbook || totalCodes === 0 || !!pendingCode}><ArrowDownToLine size={18} /> Exportar XLSX</button></div>
      </section>
      <footer><span>ColetaFlow <b>•</b> Bipe. Informe. Exporte.</span><span><Smartphone size={14} /> Coleta sem cadastro prévio</span></footer>
    </main>
  );
}
