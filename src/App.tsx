import { useEffect, useMemo, useRef, useState } from 'react';
import { BrowserMultiFormatReader, IScannerControls } from '@zxing/browser';
import * as XLSX from 'xlsx';
import {
  AlertCircle, ArrowDownToLine, Barcode, Camera, Check,
  FileSpreadsheet, FolderUp, Keyboard, Package, RotateCcw,
  ScanLine, Search, ShieldCheck, Smartphone, X
} from 'lucide-react';

type ColumnMap = { code: number; description: number; quantity: number };
type ProductRow = { code: string; description: string; row: number; currentQty: number };
type UnknownCode = { code: string; quantity: number };

const STORAGE_KEY = 'coletaflow.session.v1';
const normalize = (value: string) => value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
const codeText = (cell: XLSX.CellObject | undefined) => {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  if (cell.w !== undefined) return String(cell.w).trim();
  return String(cell.v).trim();
};
const headerScore = (value: string, terms: string[]) => terms.some(term => normalize(value).includes(term));

export default function App() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState('');
  const [sheetName, setSheetName] = useState('');
  const [headerRow, setHeaderRow] = useState(0);
  const [columns, setColumns] = useState<ColumnMap>({ code: 0, description: 1, quantity: 2 });
  const [counts, setCounts] = useState<Record<string, number>>({});
  const [unknown, setUnknown] = useState<Record<string, number>>({});
  const [scanValue, setScanValue] = useState('');
  const [search, setSearch] = useState('');
  const [notice, setNotice] = useState('Importe uma planilha para começar.');
  const [cameraOpen, setCameraOpen] = useState(false);
  const [quantityMode, setQuantityMode] = useState<'replace' | 'add'>('replace');
  const [activeTab, setActiveTab] = useState<'all' | 'collected' | 'pending'>('all');
  const videoRef = useRef<HTMLVideoElement>(null);
  const scannerRef = useRef<IScannerControls | null>(null);
  const lastCameraScan = useRef<{ code: string; at: number } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const sheet = workbook && sheetName ? workbook.Sheets[sheetName] : null;
  const range = sheet?.['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
  const headers = useMemo(() => {
    if (!sheet || !range) return [];
    const result: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      result.push(codeText(sheet[XLSX.utils.encode_cell({ r: headerRow, c })]) || `Coluna ${XLSX.utils.encode_col(c)}`);
    }
    return result;
  }, [sheet, range, headerRow]);

  const products = useMemo<ProductRow[]>(() => {
    if (!sheet || !range || !headers.length) return [];
    const list: ProductRow[] = [];
    for (let r = Math.max(headerRow + 1, range.s.r); r <= range.e.r; r++) {
      const code = codeText(sheet[XLSX.utils.encode_cell({ r, c: range.s.c + columns.code })]);
      const description = codeText(sheet[XLSX.utils.encode_cell({ r, c: range.s.c + columns.description })]);
      const quantityCell = sheet[XLSX.utils.encode_cell({ r, c: range.s.c + columns.quantity })] as XLSX.CellObject | undefined;
      const currentQty = Number(quantityCell?.v ?? 0) || 0;
      if (code) list.push({ code, description, row: r, currentQty });
    }
    return list;
  }, [sheet, range, headers, headerRow, columns]);

  const productIndex = useMemo(() => {
    const index = new Map<string, ProductRow[]>();
    products.forEach(product => {
      const key = product.code.trim();
      index.set(key, [...(index.get(key) ?? []), product]);
    });
    return index;
  }, [products]);

  const collectedUnits = Object.values(counts).reduce((sum, qty) => sum + qty, 0);
  const uniqueCollected = Object.keys(counts).length;
  const unknownUnits = Object.values(unknown).reduce((sum, qty) => sum + qty, 0);
  const visibleProducts = products.filter(product => {
    const qty = counts[product.code] ?? 0;
    const match = !search || normalize(product.code + ' ' + product.description).includes(normalize(search));
    const tabMatch = activeTab === 'all' || (activeTab === 'collected' ? qty > 0 : qty === 0);
    return match && tabMatch;
  });

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const session = JSON.parse(saved) as { counts?: Record<string, number>; unknown?: Record<string, number>; fileName?: string };
        setCounts(session.counts ?? {});
        setUnknown(session.unknown ?? {});
        if (session.fileName) {
          setFileName(session.fileName);
          setNotice(`Sessão recuperada. Reimporte “${session.fileName}” para continuar e exportar.`);
        }
      }
    } catch { /* ignore invalid local session */ }
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ counts, unknown, fileName }));
  }, [counts, unknown, fileName]);

  useEffect(() => () => scannerRef.current?.stop(), []);

  const loadWorkbook = async (file: File) => {
    try {
      const data = await file.arrayBuffer();
      const parsed = XLSX.read(data, { type: 'array', cellText: true, cellNF: true });
      const firstSheet = parsed.SheetNames[0];
      if (!firstSheet) throw new Error('A planilha não possui abas legíveis.');
      const ws = parsed.Sheets[firstSheet];
      const wsRange = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
      if (!wsRange) throw new Error('Não encontrei dados nessa planilha.');
      let detectedHeader = wsRange.s.r;
      let bestScore = -1;
      for (let r = wsRange.s.r; r <= Math.min(wsRange.s.r + 12, wsRange.e.r); r++) {
        let score = 0;
        for (let c = wsRange.s.c; c <= wsRange.e.c; c++) {
          const text = codeText(ws[XLSX.utils.encode_cell({ r, c })]);
          if (headerScore(text, ['codigo', 'ean', 'gtin', 'barras', 'barcode', 'produto', 'descricao', 'quantidade', 'qtd', 'estoque'])) score++;
        }
        if (score > bestScore) { bestScore = score; detectedHeader = r; }
      }
      const rowHeaders: string[] = [];
      for (let c = wsRange.s.c; c <= wsRange.e.c; c++) rowHeaders.push(codeText(ws[XLSX.utils.encode_cell({ r: detectedHeader, c })]));
      const findColumn = (terms: string[], fallback: number) => {
        const found = rowHeaders.findIndex(value => headerScore(value, terms));
        return found >= 0 ? found : Math.min(fallback, Math.max(0, rowHeaders.length - 1));
      };
      setWorkbook(parsed);
      setFileName(file.name);
      setSheetName(firstSheet);
      setHeaderRow(detectedHeader);
      setColumns({
        code: findColumn(['ean', 'gtin', 'codigo de barras', 'cod barras', 'barcode', 'codigo', 'cod.'], 0),
        description: findColumn(['descricao', 'produto', 'nome'], Math.min(1, rowHeaders.length - 1)),
        quantity: findColumn(['quantidade', 'qtd', 'contagem', 'estoque'], Math.min(2, rowHeaders.length - 1))
      });
      let restored = false;
      try {
        const saved = localStorage.getItem(STORAGE_KEY);
        const previous = saved ? JSON.parse(saved) as { fileName?: string } : null;
        restored = previous?.fileName === file.name;
      } catch { /* start a clean session */ }
      if (!restored) { setCounts({}); setUnknown({}); }
      setNotice(restored ? `“${file.name}” importada. Contagem anterior recuperada.` : `“${file.name}” importada. Confira as colunas antes de bipar.`);
      setActiveTab('all');
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Não foi possível abrir o arquivo.');
    }
  };

  const registerScan = (raw: string, fromCamera = false) => {
    const code = raw.trim();
    if (!code) return;
    if (fromCamera) {
      const last = lastCameraScan.current;
      if (last?.code === code && Date.now() - last.at < 1400) return;
      lastCameraScan.current = { code, at: Date.now() };
    }
    const matches = productIndex.get(code) ?? [];
    if (matches.length === 1) {
      setCounts(previous => ({ ...previous, [code]: (previous[code] ?? 0) + 1 }));
      setNotice(`✓ Leitura registrada: ${matches[0].description || code}`);
    } else {
      setUnknown(previous => ({ ...previous, [code]: (previous[code] ?? 0) + 1 }));
      setNotice(matches.length > 1 ? `Código duplicado na planilha: ${code}. Precisa de conferência.` : `Código não encontrado: ${code}. Guardado nas pendências.`);
    }
    setScanValue('');
    requestAnimationFrame(() => inputRef.current?.focus());
  };

  const startCamera = async () => {
    if (!videoRef.current) return;
    try {
      const reader = new BrowserMultiFormatReader();
      scannerRef.current = await reader.decodeFromVideoDevice(undefined, videoRef.current, (result) => {
        if (result) {
          const code = result.getText();
          registerScan(code, true);
        }
      });
      setCameraOpen(true);
    } catch {
      setNotice('Não foi possível abrir a câmera. Verifique a permissão e use HTTPS.');
    }
  };

  const stopCamera = () => {
    scannerRef.current?.stop();
    scannerRef.current = null;
    setCameraOpen(false);
  };

  const exportWorkbook = () => {
    if (!workbook || !sheet || !range) return;
    const output = XLSX.read(XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }), { type: 'array', cellStyles: true });
    const outputSheet = output.Sheets[sheetName];
    let updated = 0;
    let ambiguous = 0;
    products.forEach(product => {
      const qty = counts[product.code] ?? 0;
      if (quantityMode === 'add' && !qty) return;
      if ((productIndex.get(product.code) ?? []).length !== 1) { ambiguous++; return; }
      const cellAddress = XLSX.utils.encode_cell({ r: product.row, c: range.s.c + columns.quantity });
      const existing = Number(outputSheet[cellAddress]?.v ?? 0) || 0;
      const value = quantityMode === 'add' ? existing + qty : qty;
      outputSheet[cellAddress] = { t: 'n', v: value };
      updated++;
    });
    // Códigos ausentes no cadastro também retornam no arquivo, em novas linhas.
    let appended = 0;
    let nextRow = range.e.r + 1;
    Object.entries(unknown).forEach(([code, qty]) => {
      if (productIndex.has(code)) { ambiguous++; return; }
      outputSheet[XLSX.utils.encode_cell({ r: nextRow, c: range.s.c + columns.code })] = { t: 's', v: code };
      outputSheet[XLSX.utils.encode_cell({ r: nextRow, c: range.s.c + columns.description })] = { t: 's', v: 'NÃO CADASTRADO NA PLANILHA' };
      outputSheet[XLSX.utils.encode_cell({ r: nextRow, c: range.s.c + columns.quantity })] = { t: 'n', v: qty };
      nextRow++;
      appended++;
    });
    outputSheet['!ref'] = XLSX.utils.encode_range({
      s: range.s,
      e: { r: Math.max(range.e.r, nextRow - 1), c: Math.max(range.e.c, range.s.c + columns.code, range.s.c + columns.description, range.s.c + columns.quantity) }
    });
    const outputName = fileName.replace(/\\.(xlsx|xls|csv)$/i, '') + '-coletado.xlsx';
    XLSX.writeFile(output, outputName);
    setNotice(`Arquivo gerado: ${updated} linha(s) atualizada(s) e ${appended} código(s) novo(s) acrescentado(s). ${ambiguous ? ambiguous + ' código(s) duplicado(s) ficaram sem alteração.' : ''}`);
  };

  const resetSession = () => {
    if (!window.confirm('Limpar todas as quantidades desta sessão?')) return;
    setCounts({});
    setUnknown({});
    setNotice('Contagem reiniciada. A planilha continua carregada.');
  };

  const updateCount = (code: string, delta: number) => {
    setCounts(previous => {
      const next = Math.max(0, (previous[code] ?? 0) + delta);
      const updated = { ...previous };
      if (next === 0) delete updated[code]; else updated[code] = next;
      return updated;
    });
  };

  const changeSheet = (name: string) => {
    setSheetName(name);
    setCounts({});
    setUnknown({});
    setNotice(`Aba “${name}” selecionada. Confira o mapeamento das colunas.`);
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#" aria-label="ColetaFlow início">
          <span className="brand-mark"><ScanLine size={25} /></span>
          <span><strong>ColetaFlow</strong><small>COLETOR INTELIGENTE</small></span>
        </a>
        <div className="topbar-right"><span className="offline-pill"><span /> CONTAGEM LOCAL</span><span className="version">MVP 0.1</span></div>
      </header>

      <section className="hero">
        <div className="hero-copy">
          <span className="eyebrow">COLETA UNIVERSAL DE PRODUTOS</span>
          <h1>Bipe rápido.<br /><em>Confira com precisão.</em></h1>
          <p>Importe a planilha da empresa, faça a coleta e exporte os dados no modelo original.</p>
        </div>
        <div className="hero-art"><div className="scan-orbit orbit-one" /><div className="scan-orbit orbit-two" /><Barcode size={76} strokeWidth={1.25} /></div>
      </section>

      <section className="workspace">
        <div className="section-heading"><div><span className="step-label">ETAPA 01</span><h2>Planilha de origem</h2><p>Comece com o arquivo que sua empresa já utiliza.</p></div>
          <label className="primary-button upload-button"><FolderUp size={18} /> Importar planilha<input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /></label>
        </div>

        {workbook && sheet ? (
          <div className="file-card">
            <div className="file-icon"><FileSpreadsheet size={24} /></div>
            <div className="file-info"><strong>{fileName}</strong><span>{products.length} produtos identificados · {workbook.SheetNames.length} aba(s)</span></div>
            <span className="file-status"><Check size={14} /> Importada</span>
            <button className="icon-button" title="Remover planilha" onClick={() => { setWorkbook(null); setFileName(''); setSheetName(''); setCounts({}); setUnknown({}); }}><X size={18} /></button>
          </div>
        ) : <label className="dropzone"><input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /><span className="drop-icon"><FolderUp size={25} /></span><strong>Selecione ou arraste sua planilha</strong><span>Excel (.xlsx, .xls) ou CSV</span></label>}

        {workbook && sheet && (
          <div className="mapping-panel">
            <div className="mapping-title"><div><ShieldCheck size={18} /><strong>Confira o reconhecimento das colunas</strong></div><span>MAPEAMENTO EDITÁVEL</span></div>
            <div className="mapping-grid">
              <label>Aba da planilha<select value={sheetName} onChange={event => changeSheet(event.target.value)}>{workbook.SheetNames.map(name => <option key={name}>{name}</option>)}</select></label>
              <label>Código de barras<select value={columns.code} onChange={event => setColumns(v => ({ ...v, code: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h || `Coluna ${i+1}`}</option>)}</select></label>
              <label>Descrição do produto<select value={columns.description} onChange={event => setColumns(v => ({ ...v, description: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h || `Coluna ${i+1}`}</option>)}</select></label>
              <label>Quantidade de destino<select value={columns.quantity} onChange={event => setColumns(v => ({ ...v, quantity: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h || `Coluna ${i+1}`}</option>)}</select></label>
            </div>
            <div className="mapping-note"><AlertCircle size={15} /> A identificação é uma sugestão automática. Confira as colunas antes de iniciar a coleta.</div>
          </div>
        )}
      </section>

      <section className="workspace collector-section">
        <div className="section-heading"><div><span className="step-label">ETAPA 02</span><h2>Coletor de produtos</h2><p>Use um leitor físico ou a câmera do dispositivo.</p></div><span className={workbook ? 'ready-chip' : 'waiting-chip'}>{workbook ? 'PLANILHA PRONTA' : 'AGUARDANDO PLANILHA'}</span></div>
        <div className="stats-grid">
          <div className="stat-card"><span>Unidades coletadas</span><strong>{collectedUnits}</strong><small>Total de bipagens válidas</small></div>
          <div className="stat-card"><span>Produtos coletados</span><strong>{uniqueCollected}</strong><small>Itens diferentes</small></div>
          <div className="stat-card warning-stat"><span>Códigos pendentes</span><strong>{unknownUnits}</strong><small>Não encontrados ou ambíguos</small></div>
        </div>
        <div className="scanner-panel">
          <div className="scanner-heading"><div className="scanner-icon"><Barcode size={22} /></div><div><strong>Leitura contínua</strong><span>Pronto para receber o próximo código</span></div><span className="live-indicator"><i /> AO VIVO</span></div>
          <form className="scan-form" onSubmit={event => { event.preventDefault(); registerScan(scanValue); }}>
            <div className="scan-input-wrap"><Keyboard size={20} /><input ref={inputRef} value={scanValue} onChange={event => setScanValue(event.target.value)} placeholder={workbook ? 'Bipe ou digite o código de barras...' : 'Importe uma planilha primeiro'} disabled={!workbook} autoComplete="off" autoFocus /><button type="submit" disabled={!workbook || !scanValue.trim()}><ScanLine size={19} /><span>Registrar</span></button></div>
          </form>
          <div className="scanner-actions"><button className="secondary-button" onClick={() => void startCamera()} disabled={!workbook || cameraOpen}><Camera size={17} /> Abrir câmera</button><span>Leitor Bluetooth/USB compatível com modo teclado (HID)</span></div>
          {cameraOpen && <div className="camera-box"><video ref={videoRef} muted playsInline /><button className="camera-close" onClick={stopCamera}><X size={18} /> Fechar câmera</button></div>}
          <div className="scan-feedback">{notice}</div>
        </div>
      </section>

      <section className="workspace results-section">
        <div className="section-heading"><div><span className="step-label">ETAPA 03</span><h2>Conferência da coleta</h2><p>Revise os produtos e corrija quantidades quando necessário.</p></div><button className="text-button" onClick={resetSession} disabled={!workbook}><RotateCcw size={16} /> Reiniciar contagem</button></div>
        <div className="results-toolbar">
          <div className="tabs"><button className={activeTab === 'all' ? 'active' : ''} onClick={() => setActiveTab('all')}>Todos ({products.length})</button><button className={activeTab === 'collected' ? 'active' : ''} onClick={() => setActiveTab('collected')}>Coletados ({uniqueCollected})</button><button className={activeTab === 'pending' ? 'active' : ''} onClick={() => setActiveTab('pending')}>Não coletados ({products.length - uniqueCollected})</button></div>
          <label className="search-box"><Search size={17} /><input value={search} onChange={event => setSearch(event.target.value)} placeholder="Buscar produto ou código" /></label>
        </div>
        <div className="product-list">
          {!workbook ? <div className="empty-state"><Package size={28} /><strong>Sua lista aparecerá aqui</strong><span>Importe uma planilha para visualizar os produtos.</span></div> :
            visibleProducts.slice(0,100).map((product, index) => <div className="product-row" key={product.row}><div className="product-number">{index + 1}</div><div className="product-details"><strong>{product.description || 'Produto sem descrição'}</strong><span>{product.code}{(productIndex.get(product.code)?.length ?? 0) > 1 && <em className="duplicate-tag">Código duplicado</em>}</span></div><div className="quantity-control"><button onClick={() => updateCount(product.code, -1)} disabled={!counts[product.code]}>−</button><strong>{counts[product.code] ?? 0}</strong><button onClick={() => updateCount(product.code, 1)}>+</button></div></div>)}
          {workbook && visibleProducts.length === 0 && <div className="empty-state"><Search size={24} /><strong>Nenhum produto encontrado</strong><span>Altere a busca ou o filtro.</span></div>}
          {visibleProducts.length > 100 && <div className="list-footnote">Mostrando os primeiros 100 resultados. Use a busca para localizar outros produtos.</div>}
        </div>
        {Object.keys(unknown).length > 0 && <div className="unknown-panel"><div><AlertCircle size={18} /><strong>Pendências de leitura</strong></div>{Object.entries(unknown).map(([code,qty]) => <span key={code}><code>{code}</code><b>{qty}x</b></span>)}</div>}
      </section>

      <section className="export-bar">
        <div><span className="step-label">ETAPA 04</span><strong>Planilha pronta para voltar à empresa?</strong><small>O arquivo original não será sobrescrito.</small></div>
        <div className="export-actions"><label>Regra de quantidade<select value={quantityMode} onChange={event => setQuantityMode(event.target.value as 'replace' | 'add')}><option value="replace">Substituir pela coleta</option><option value="add">Somar à quantidade original</option></select></label><button className="primary-button" onClick={exportWorkbook} disabled={!workbook || collectedUnits === 0}><ArrowDownToLine size={18} /> Exportar XLSX</button></div>
      </section>
      <footer><span>ColetaFlow <b>•</b> Bipe. Confira. Exporte.</span><span><Smartphone size={14} /> Desenvolvido para coleta rápida</span></footer>
    </main>
  );
}