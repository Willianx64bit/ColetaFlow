import { useEffect, useMemo, useRef, useState } from 'react';
import { BrowserMultiFormatReader, IScannerControls } from '@zxing/browser';
import * as XLSX from 'xlsx';
import {
  ArrowDownToLine, Camera, Check, FileSpreadsheet, FolderUp,
  Keyboard, Package, RotateCcw, Search, ShieldCheck, Smartphone, X, RefreshCw
} from 'lucide-react';

type ColumnMap = { code: number; quantity: number; description: number };
type SavedSession = {
  workbookBytes?: ArrayBuffer | Uint8Array | number[] | null;
  fileName?: string;
  sheetName?: string;
  headerRow?: number;
  columns?: ColumnMap;
  counts?: Record<string, number>;
  currentStep?: number;
  pendingCode?: string;
  quantityValue?: string;
  autoReopenCamera?: boolean;
  cameraZoom?: number;
};
const STORAGE_KEY = 'coletaflow.blind-session.v2';
const SESSION_DB = 'coletaflow-session-db';
const SESSION_STORE = 'sessions';
const openSessionDb = () => new Promise<IDBDatabase>((resolve, reject) => {
  const request = indexedDB.open(SESSION_DB, 1);
  request.onupgradeneeded = () => {
    if (!request.result.objectStoreNames.contains(SESSION_STORE)) request.result.createObjectStore(SESSION_STORE);
  };
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(request.error);
});
const readSavedSession = async (): Promise<SavedSession | null> => {
  const db = await openSessionDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(SESSION_STORE, 'readonly').objectStore(SESSION_STORE).get('active');
    request.onsuccess = () => { db.close(); resolve((request.result as SavedSession) || null); };
    request.onerror = () => { db.close(); reject(request.error); };
  });
};
const writeSavedSession = async (session: SavedSession) => {
  const db = await openSessionDb();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(SESSION_STORE, 'readwrite');
    transaction.objectStore(SESSION_STORE).put(session, 'active');
    transaction.oncomplete = () => { db.close(); resolve(); };
    transaction.onerror = () => { db.close(); reject(transaction.error); };
    transaction.onabort = () => { db.close(); reject(transaction.error); };
  });
};

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
  const [columns, setColumns] = useState<ColumnMap>({ code: 0, quantity: 1, description: -1 });
  const [counts, setCounts] = useState<Record<string, number>>({});
  const [scanValue, setScanValue] = useState('');
  const [quantityValue, setQuantityValue] = useState('');
  const [quantityShortcuts, setQuantityShortcuts] = useState<number[]>([1, 5, 10, 20, 50]);
  const [editingShortcuts, setEditingShortcuts] = useState(false);
  const [pendingCode, setPendingCode] = useState('');
  const [search, setSearch] = useState('');
  const [notice, setNotice] = useState('Importe o modelo vazio da empresa para começar.');
  const [cameraOpen, setCameraOpen] = useState(false);
  const [cameraUsed, setCameraUsed] = useState(false);
  const [autoReopenCamera, setAutoReopenCamera] = useState(false);
  const [currentStep, setCurrentStep] = useState(1);
  const [cameraStatus, setCameraStatus] = useState('Aguardando câmera');
  const [cameraZoom, setCameraZoom] = useState(1);
  const videoRef = useRef<HTMLVideoElement>(null);
  const scannerRef = useRef<IScannerControls | null>(null);
  const lastCameraScan = useRef<{ code: string; at: number } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const quantityRef = useRef<HTMLInputElement>(null);
  const shortcutHoldTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const shortcutLongPressHandled = useRef(false);

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
      const savedShortcuts = localStorage.getItem('coletaflow_quantity_shortcuts');
      if (savedShortcuts) {
        const parsed = JSON.parse(savedShortcuts);
        if (Array.isArray(parsed) && parsed.length === 5 && parsed.every(value => Number.isInteger(value) && value > 0)) setQuantityShortcuts(parsed);
      }
    } catch { /* ignore invalid shortcut preferences */ }
  }, []);

  useEffect(() => {
    try { localStorage.setItem('coletaflow_quantity_shortcuts', JSON.stringify(quantityShortcuts)); } catch { /* storage may be unavailable */ }
  }, [quantityShortcuts]);

  const [sessionRestored, setSessionRestored] = useState(false);
  const workbookBytesRef = useRef<ArrayBuffer | null>(null);

  useEffect(() => {
    let active = true;
    const restore = async () => {
      try {
        const session = await readSavedSession();
        if (!active) return;
        if (session?.workbookBytes) {
          const bytes = session.workbookBytes instanceof ArrayBuffer
            ? session.workbookBytes
            : new Uint8Array(session.workbookBytes as unknown as number[]).buffer;
          workbookBytesRef.current = bytes;
          const restoredWorkbook = XLSX.read(bytes, { type: 'array', cellText: true, cellNF: true });
          setWorkbook(restoredWorkbook);
          setFileName(session.fileName || '');
          setSheetName(session.sheetName && restoredWorkbook.Sheets[session.sheetName]
            ? session.sheetName
            : restoredWorkbook.SheetNames[0] || '');
          setHeaderRow(Math.max(0, Number(session.headerRow ?? 0) || 0));
          setColumns(session.columns || { code: 0, quantity: 1, description: -1 });
          setCounts(session.counts || {});
          setCurrentStep(Math.min(4, Math.max(1, Number(session.currentStep ?? 1))));
          setPendingCode(session.pendingCode || '');
          setQuantityValue(session.quantityValue || '');
          setAutoReopenCamera(!!session.autoReopenCamera);
          setCameraZoom(Number(session.cameraZoom ?? 1));
          setNotice('Sessão recuperada automaticamente. Você pode continuar de onde parou.');
        } else {
          try {
            const legacy = JSON.parse(localStorage.getItem('coletaflow.blind-session.v2') || '{}') as { counts?: Record<string, number>; fileName?: string };
            if (legacy.counts) setCounts(legacy.counts);
            if (legacy.fileName) setFileName(legacy.fileName);
          } catch { /* ignore legacy session */ }
        }
      } catch {
        // Uma sessão incompatível ou corrompida não impede iniciar uma nova coleta.
      } finally {
        if (active) setSessionRestored(true);
      }
    };
    void restore();
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!sessionRestored) return;
    const timer = window.setTimeout(() => {
      try {
        const snapshot = {
          workbookBytes: workbookBytesRef.current,
          fileName,
          sheetName,
          headerRow,
          columns,
          counts,
          currentStep,
          pendingCode,
          quantityValue,
          autoReopenCamera,
          cameraZoom
        };
        void writeSavedSession(snapshot).catch(() => setNotice('Não foi possível salvar a sessão neste dispositivo. Verifique o espaço disponível.'));
      } catch {
        setNotice('Não foi possível salvar a sessão neste dispositivo. Verifique o espaço disponível.');
      }
    }, 250);
    return () => window.clearTimeout(timer);
  }, [sessionRestored, fileName, sheetName, headerRow, columns, counts, currentStep, pendingCode, quantityValue, autoReopenCamera, cameraZoom]);

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
        const stream = video.srcObject as MediaStream | null;
        const track = stream?.getVideoTracks()[0];
        if (track) {
          const capabilities = track.getCapabilities() as MediaTrackCapabilities & { zoom?: { min: number; max: number; step?: number } };
          if (capabilities.zoom) {
            const targetZoom = Math.min(capabilities.zoom.max, Math.max(capabilities.zoom.min, cameraZoom));
            track.applyConstraints({ advanced: [{ zoom: targetZoom } as MediaTrackConstraintSet] }).catch(() => {});
          }
          const focusCapabilities = capabilities as MediaTrackCapabilities & { focusMode?: string[] };
          if (focusCapabilities.focusMode?.includes('continuous')) {
            track.applyConstraints({ advanced: [{ focusMode: 'continuous' } as MediaTrackConstraintSet] }).catch(() => {});
          }
        }
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

  useEffect(() => {
    if (!cameraOpen || !videoRef.current) return;
    const stream = videoRef.current.srcObject as MediaStream | null;
    const track = stream?.getVideoTracks()[0];
    if (!track) return;
    const capabilities = track.getCapabilities() as MediaTrackCapabilities & { zoom?: { min: number; max: number } };
    if (!capabilities.zoom) return;
    const targetZoom = Math.min(capabilities.zoom.max, Math.max(capabilities.zoom.min, cameraZoom));
    track.applyConstraints({ advanced: [{ zoom: targetZoom } as MediaTrackConstraintSet] }).catch(() => {
      setCameraStatus('Zoom não disponível nesta câmera');
    });
  }, [cameraZoom, cameraOpen]);

  const loadWorkbook = async (file: File) => {
    try {
      const data = await file.arrayBuffer();
      workbookBytesRef.current = data.slice(0);
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
      const detectedCodeColumn = findColumn(['ean', 'gtin', 'codigo de barras', 'cod barras', 'barcode', 'codigo', 'cod.'], 0);
      const detectedDescriptionColumn = rowHeaders.findIndex(value => hasTerm(value, ['descricao', 'produto', 'nome', 'mercadoria', 'item']));
      setColumns({
        code: detectedCodeColumn,
        quantity: findColumn(['quantidade', 'qtd', 'contagem', 'unidades'], 1),
        description: detectedDescriptionColumn >= 0 && detectedDescriptionColumn !== detectedCodeColumn ? detectedDescriptionColumn : -1
      });
      if (!sameTemplate) setCounts({});
      setPendingCode('');
      setNotice(sameTemplate ? '' : 'Modelo vazio importado. Agora é só bipar o código e informar a quantidade.');
      setCurrentStep(1);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Não foi possível abrir o modelo.');
    }
  };

  const findProductRow = (code: string) => {
    if (!sheet || !range) return -1;
    const codeColumn = range.s.c + columns.code;
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const rowCode = cellText(sheet[XLSX.utils.encode_cell({ r, c: codeColumn })]);
      if (rowCode && rowCode.trim() === code.trim()) return r;
    }
    return -1;
  };

  const findProductDescription = (code: string) => {
    if (!sheet || !range || columns.description < 0) return '';
    const productRow = findProductRow(code);
    return productRow >= 0
      ? cellText(sheet[XLSX.utils.encode_cell({ r: productRow, c: range.s.c + columns.description })])
      : '';
  };

  const receiveCode = (raw: string, fromCamera = false) => {
    const code = raw.trim();
    if (!code) return;
    if (fromCamera) {
      const last = lastCameraScan.current;
      if (last?.code === code && Date.now() - last.at < 1600) return;
      lastCameraScan.current = { code, at: Date.now() };
    }
    const productDescription = findProductDescription(code);
    const catalogCodes = sheet && range ? Array.from({ length: Math.max(0, range.e.r - headerRow) }, (_, index) => {
      const r = headerRow + 1 + index;
      return cellText(sheet[XLSX.utils.encode_cell({ r, c: range.s.c + columns.code })]).trim();
    }).filter(Boolean) : [];
    const codeExists = catalogCodes.includes(code);
    setPendingCode(code);
    setScanValue('');
    setNotice(productDescription
      ? `${productDescription} · informe a quantidade.`
      : catalogCodes.length > 0 && !codeExists
        ? `Código ${code} não encontrado na lista. Você pode continuar a coleta.`
        : `Código ${code} lido. Informe a quantidade.`);
    requestAnimationFrame(() => quantityRef.current?.focus());
  };

  const confirmCount = (shortcutQuantity?: number) => {
    const code = pendingCode.trim();
    const qty = shortcutQuantity ?? Number(quantityValue);
    if (!code) { setNotice('Bipe ou digite um código primeiro.'); inputRef.current?.focus(); return; }
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) { setNotice('Informe uma quantidade inteira maior que zero.'); quantityRef.current?.focus(); return; }
    setCounts(previous => ({ ...previous, [code]: (previous[code] ?? 0) + qty }));
    setNotice(`Registrado: código ${code} · quantidade adicionada: ${qty}.`);
    setPendingCode('');
    setQuantityValue('');
    setScanValue('');
    if (autoReopenCamera && cameraUsed) {
      setCameraStatus('Preparando próxima leitura…');
      setCameraOpen(true);
    } else {
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        inputRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      });
    }
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
      let matchingRow = -1;
      for (let r = headerRow + 1; r <= lastUsedRow; r++) {
        const existingCode = cellText(sheet[XLSX.utils.encode_cell({ r, c: startColumn + columns.code })]).trim();
        if (existingCode === code) { matchingRow = r; break; }
      }
      if (matchingRow >= 0) {
        const qtyCell = XLSX.utils.encode_cell({ r: matchingRow, c: startColumn + columns.quantity });
        outputSheet[qtyCell] = { t: 'n', v: qty };
      } else {
        const codeCell = XLSX.utils.encode_cell({ r: row, c: startColumn + columns.code });
        const qtyCell = XLSX.utils.encode_cell({ r: row, c: startColumn + columns.quantity });
        outputSheet[codeCell] = { t: 's', v: code, z: '@' };
        outputSheet[qtyCell] = { t: 'n', v: qty };
        row++;
      }
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
          <span className="brand-mark" aria-hidden="true"><FileSpreadsheet size={34} strokeWidth={2.2} /><RefreshCw size={17} strokeWidth={2.8} className="brand-sync" /></span>
          <span className="brand-wordmark"><span>Coleta</span><strong>Flow</strong></span>
        </a>
        <div className="topbar-right"><span className="offline-pill"><span /> CONTAGEM LOCAL</span><span className="version">MVP 0.2</span></div>
      </header>


      <nav className="wizard-progress" aria-label="Etapas da coleta">{[1,2,3,4].map((step) => <div key={step} className={`wizard-step ${currentStep === step ? "current" : currentStep > step ? "completed" : ""}`}><span>{currentStep > step ? "✓" : step}</span><small>{["Modelo","Coleta","Conferência","Finalizar"][step - 1]}</small></div>)}</nav>

      {currentStep === 1 && <section className="workspace wizard-screen">
        <div className="section-heading"><div><span className="step-label">ETAPA 01 DE 04</span><h2>Modelo da empresa</h2><p>Importe a planilha vazia que deverá receber os códigos e quantidades.</p></div>
          <label className="primary-button upload-button"><FolderUp size={18} /> Importar modelo<input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /></label>
        </div>
        {workbook && sheet ? (
          <div className="file-card">
            <div className="file-icon"><FileSpreadsheet size={24} /></div>
            <div className="file-info"><strong>{fileName}</strong><span>Modelo carregado · {workbook.SheetNames.length} aba(s) · sem necessidade de cadastro de produtos</span></div>
            <span className="file-status"><Check size={14} /> Pronto</span>
            <button className="icon-button" title="Remover modelo" onClick={() => { setWorkbook(null); workbookBytesRef.current = null; setFileName(''); setSheetName(''); setCounts({}); setPendingCode(''); setCurrentStep(1); }}><X size={18} /></button>
          </div>
        ) : <label className="dropzone"><input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /><span className="drop-icon"><FolderUp size={25} /></span><strong>Selecione o modelo vazio da empresa</strong><span>Excel (.xlsx, .xls) ou CSV</span></label>}

        {workbook && sheet && (
          <div className="mapping-panel">
            <div className="mapping-title"><div><ShieldCheck size={18} /><strong>Defina onde gravar os dados</strong></div><span>MAPEAMENTO DO MODELO</span></div>
            <div className="mapping-grid">
              <label>Aba<select value={sheetName} onChange={event => { setSheetName(event.target.value); setCounts({}); setPendingCode(''); }}>{workbook.SheetNames.map(name => <option key={name} value={name}>{name}</option>)}</select></label>
              <label>Coluna do código<select value={columns.code} onChange={event => setColumns(v => ({ ...v, code: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
              <label>Coluna da quantidade<select value={columns.quantity} onChange={event => setColumns(v => ({ ...v, quantity: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
              <label>Coluna da descrição<select value={columns.description} onChange={event => setColumns(v => ({ ...v, description: Number(event.target.value) }))}><option value={-1}>Não possui descrição</option>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
            </div>
            <div className="mapping-note">Se houver códigos e descrições na planilha, o ColetaFlow identifica o produto ao bipar. Sem descrição ou sem código cadastrado, a coleta continua normalmente.</div>
          </div>
        )}
        <div className="wizard-navigation wizard-next"><span>Confira o modelo e ajuste as colunas antes de continuar.</span><button className="primary-button" disabled={!workbook || !sheet} onClick={() => setCurrentStep(2)}>Continuar <ArrowDownToLine size={17} /></button></div>
      </section>}

      {currentStep === 2 && <section className="workspace collector-section wizard-screen">
        <div className="section-heading"><div><span className="step-label">ETAPA 02 DE 04</span><h2>Coletor</h2><p>Bipe qualquer código e informe quantas unidades encontrou.</p></div><span className={workbook ? 'ready-chip' : 'waiting-chip'}>{workbook ? 'MODELO PRONTO' : 'AGUARDANDO MODELO'}</span></div>
          {cameraOpen && (
            <div className="camera-box">
              <div className="camera-status"><span className="camera-live-dot" />{cameraStatus}</div>
              <video ref={videoRef} muted playsInline autoPlay />
              <div className="camera-guide" aria-hidden="true"><span /><i /></div>
              <div className="camera-zoom-controls" role="group" aria-label="Zoom da câmera">
                {[1, 2, 3].map(level => <button type="button" key={level} className={cameraZoom === level ? 'active' : ''} onClick={() => setCameraZoom(level)} aria-pressed={cameraZoom === level} aria-label={`Zoom ${level}x`} title={`Zoom ${level}x`}><span /></button>)}
              </div>
              <button className="camera-close" onClick={stopCamera}><X size={18} /> Fechar câmera</button>
            </div>
          )}
          {pendingCode && <div className="pending-code-card"><div><small>CÓDIGO LIDO</small><strong>{pendingCode}</strong>{findProductDescription(pendingCode) ? <span className="pending-product-name">{findProductDescription(pendingCode)}</span> : <span className={findProductRow(pendingCode) < 0 && !!sheet && !!range && Array.from({ length: Math.max(0, range.e.r - headerRow) }, (_, index) => cellText(sheet[XLSX.utils.encode_cell({ r: headerRow + 1 + index, c: range.s.c + columns.code })])).some(Boolean) ? 'pending-product-name pending-product-missing' : 'pending-product-name'}>{findProductRow(pendingCode) >= 0 || !sheet || !range || !Array.from({ length: Math.max(0, range.e.r - headerRow) }, (_, index) => cellText(sheet[XLSX.utils.encode_cell({ r: headerRow + 1 + index, c: range.s.c + columns.code })])).some(Boolean) ? 'Sem descrição cadastrada' : 'Produto não encontrado na lista'}</span>}</div><div className="quantity-shortcuts"><div className="quantity-shortcut-list">{quantityShortcuts.map((amount, index) => <button type="button" className="quantity-shortcut" key={index} onPointerDown={event => { event.preventDefault(); shortcutLongPressHandled.current = false; if (shortcutHoldTimer.current) clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = setTimeout(() => { shortcutLongPressHandled.current = true; confirmCount(amount); shortcutHoldTimer.current = null; }, 500); }} onPointerUp={() => { if (shortcutHoldTimer.current) { clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = null; } }} onPointerLeave={() => { if (!shortcutLongPressHandled.current && shortcutHoldTimer.current) { clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = null; } }} onPointerCancel={() => { if (!shortcutLongPressHandled.current && shortcutHoldTimer.current) { clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = null; } }} onContextMenu={event => event.preventDefault()} onClick={event => { event.preventDefault(); if (shortcutLongPressHandled.current) { shortcutLongPressHandled.current = false; return; } setQuantityValue(String((Number(quantityValue) || 0) + amount)); } }>{amount}</button>)}</div><button type="button" className="edit-shortcuts-button" onClick={() => setEditingShortcuts(value => !value)}>{editingShortcuts ? 'Concluir edição' : 'Editar atalhos'}</button>{editingShortcuts && <div className="edit-shortcuts-grid">{quantityShortcuts.map((amount, index) => <label key={index}>Atalho {index + 1}<input type="number" min="1" step="1" value={amount} onChange={event => setQuantityShortcuts(values => values.map((value, i) => i === index ? Math.max(1, Number(event.target.value) || 1) : value))} /></label>)}</div>}</div><label>Quantidade<input ref={quantityRef} type="number" min="1" step="1" value={quantityValue} onChange={event => setQuantityValue(event.target.value)} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); confirmCount(); } }} /></label><button className="primary-button" onClick={() => confirmCount()}><Check size={17} /> Confirmar</button><button className="icon-button" title="Cancelar código" onClick={() => { setPendingCode(''); setQuantityValue(''); inputRef.current?.focus(); }}><X size={18} /></button></div>}
        <div className="stats-grid">
          <div className="stat-card"><span>Unidades contadas</span><strong>{totalUnits}</strong><small>Soma das quantidades informadas</small></div>
          <div className="stat-card"><span>Códigos diferentes</span><strong>{totalCodes}</strong><small>Identificadores registrados</small></div>
          <div className="stat-card warning-stat"><span>Aguardando quantidade</span><strong>{pendingCode ? '1' : '0'}</strong><small>{pendingCode ? 'Código lido, falta confirmar' : 'Nenhum código pendente'}</small></div>
        </div>
        <div className="scanner-panel">

          <form className="scan-form" onSubmit={event => { event.preventDefault(); receiveCode(scanValue); }}>
            <div className="scan-input-wrap"><Keyboard size={20} /><input ref={inputRef} value={scanValue} onChange={event => setScanValue(event.target.value)} placeholder={!workbook ? 'Importe o modelo primeiro' : pendingCode ? 'Informe a quantidade do código lido' : 'Bipe ou digite o código...'} disabled={!workbook || !!pendingCode} autoComplete="off" autoFocus /><button type="button" className="scan-camera-inline" onClick={startCamera} disabled={!workbook || cameraOpen || !!pendingCode} aria-label="Abrir câmera" title="Ler código com a câmera"><Camera size={21} /></button></div>
          </form>
          <div className="scanner-heading scanner-controls"><button type="button" className={autoReopenCamera ? 'secondary-button auto-reopen active' : 'secondary-button auto-reopen'} onClick={() => setAutoReopenCamera(value => !value)} aria-pressed={autoReopenCamera} disabled={!workbook || !!pendingCode} title="Ao confirmar a quantidade, inicia a câmera para a próxima leitura."><RefreshCw size={16} /> {autoReopenCamera ? 'Reabertura automática: ligada' : 'Reabertura automática: desligada'}</button></div>

          <div className="scan-feedback">{notice}</div>
        </div>
        <div className="wizard-navigation"><button className="secondary-button" onClick={() => { stopCamera(); setCurrentStep(1); }}>Voltar</button><button className="primary-button" disabled={!!pendingCode || totalCodes === 0} onClick={() => { stopCamera(); setCurrentStep(3); }}>Continuar para conferência <ArrowDownToLine size={17} /></button></div>
      </section>}

      {currentStep === 3 && <section className="workspace results-section wizard-screen">
        <div className="section-heading"><div><span className="step-label">ETAPA 03 DE 04</span><h2>Resumo da coleta</h2><p>Confira somente os códigos e as quantidades registradas.</p></div><button className="text-button" onClick={resetSession} disabled={!workbook}><RotateCcw size={16} /> Reiniciar contagem</button></div>
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
        <div className="wizard-navigation"><button className="secondary-button" onClick={() => setCurrentStep(2)}>Voltar à coleta</button><button className="primary-button" disabled={totalCodes === 0} onClick={() => setCurrentStep(4)}>Continuar para finalizar <ArrowDownToLine size={17} /></button></div>
      </section>}

      {currentStep === 4 && <section className="export-bar wizard-screen">
        <div><span className="step-label">ETAPA 04 DE 04</span><strong>Exportar modelo preenchido</strong><small>O arquivo original fica intacto; será criada uma cópia com código e quantidade.</small></div>
        <div className="export-actions"><button className="secondary-button" onClick={() => setCurrentStep(3)}>Voltar</button><button className="primary-button" onClick={exportWorkbook} disabled={!workbook || totalCodes === 0 || !!pendingCode}><ArrowDownToLine size={18} /> Finalizar e exportar XLSX</button></div>
      </section>}
      <footer><span>ColetaFlow <b>•</b> Bipe. Informe. Exporte.</span><span><Smartphone size={14} /> Coleta sem cadastro prévio</span></footer>
    </main>
  );
}
