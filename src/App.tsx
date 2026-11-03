import { useEffect, useMemo, useRef, useState, type FormEvent } from 'react';
import { createPortal } from 'react-dom';
import { BrowserMultiFormatReader, IScannerControls } from '@zxing/browser';
import * as XLSX from 'xlsx';
import {
  ArrowDownToLine, Camera, Check, FileSpreadsheet, FolderUp,
  Keyboard, Package, RotateCcw, Search, ShieldCheck, Smartphone, X, RefreshCw, Zap, Pencil, Volume2, VolumeX, Sun, Moon
} from 'lucide-react';

type ColumnMap = { code: number[]; quantity: number; description: number };
type ScanEvent = { eventId: string; barcode: string; quantity: number; eventType: 'scan' | 'adjustment'; scannedAt: string; deviceId: string; description?: string };
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
  quickCollectEnabled?: boolean;
  quickCollectQuantity?: number;
  scanEvents?: ScanEvent[];
  hasProductList?: boolean;
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
const readSavedSession = async (sessionKey = 'active'): Promise<SavedSession | null> => {
  const db = await openSessionDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(SESSION_STORE, 'readonly').objectStore(SESSION_STORE).get(sessionKey);
    request.onsuccess = () => { db.close(); resolve((request.result as SavedSession) || null); };
    request.onerror = () => { db.close(); reject(request.error); };
  });
};
const writeSavedSession = async (session: SavedSession, sessionKey = 'active') => {
  const db = await openSessionDb();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(SESSION_STORE, 'readwrite');
    transaction.objectStore(SESSION_STORE).put(session, sessionKey);
    transaction.oncomplete = () => { db.close(); resolve(); };
    transaction.onerror = () => { db.close(); reject(transaction.error); };
    transaction.onabort = () => { db.close(); reject(transaction.error); };
  });
};
const readSavedEvents = async (sessionKey: string): Promise<ScanEvent[] | null> => {
  const db = await openSessionDb();
  return new Promise((resolve, reject) => {
    const request = db.transaction(SESSION_STORE, 'readonly').objectStore(SESSION_STORE).get(`${sessionKey}:events`);
    request.onsuccess = () => { db.close(); resolve(request.result === undefined ? null : (Array.isArray(request.result) ? request.result as ScanEvent[] : [])); };
    request.onerror = () => { db.close(); reject(request.error); };
  });
};
const writeSavedEvents = async (events: ScanEvent[], sessionKey: string) => {
  const db = await openSessionDb();
  return new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(SESSION_STORE, 'readwrite');
    transaction.objectStore(SESSION_STORE).put(events, `${sessionKey}:events`);
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

function CollectorApp({ sharedCollection, onFinished }: { sharedCollection?: { id: string; token: string; userId: string; accountToken?: string; isOwner?: boolean; mappingLocked?: boolean; template?: SharedTemplate; onMappingSaved?: (mapping: SharedMapping) => void }; onFinished?: () => void } = {}) {
  const sessionKey = sharedCollection ? `shared:${sharedCollection.id}:${sharedCollection.userId}` : 'active';
  const [theme, setTheme] = useState<'light' | 'dark'>(() => { try { return localStorage.getItem('coletaflow.theme') === 'dark' ? 'dark' : 'light'; } catch { return 'light'; } });
  useEffect(() => { document.documentElement.dataset.theme = theme; document.body.dataset.theme = theme; try { localStorage.setItem('coletaflow.theme', theme); } catch { /* armazenamento pode estar indisponível */ } }, [theme]);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState('');
  const [sheetName, setSheetName] = useState('');
  const [headerRow, setHeaderRow] = useState(0);
  const [columns, setColumns] = useState<ColumnMap>({ code: [0], quantity: 1, description: -1 });
  const [hasProductList, setHasProductList] = useState(true);
  const [pendingDescription, setPendingDescription] = useState('');
  const [counts, setCounts] = useState<Record<string, number>>({});
  const [scanEvents, setScanEvents] = useState<ScanEvent[]>([]);
  const scanEventsRef = useRef<ScanEvent[]>([]);
  const persistEventsRef = useRef<Promise<void>>(Promise.resolve());
  const persistEvents = (events: ScanEvent[]) => {
    persistEventsRef.current = persistEventsRef.current.catch(() => {}).then(() => writeSavedEvents(events, sessionKey));
    return persistEventsRef.current;
  };
  const [finishBusy, setFinishBusy] = useState(false);
  const [mappingBusy, setMappingBusy] = useState(false);
  const [scanValue, setScanValue] = useState('');
  const [quantityValue, setQuantityValue] = useState('');
  const [removeMode, setRemoveMode] = useState(false);
  const [quantityShortcuts, setQuantityShortcuts] = useState<number[]>([1, 5, 10, 20, 50]);
  const [editingShortcuts, setEditingShortcuts] = useState(false);
  const [pendingCode, setPendingCode] = useState('');
  const [search, setSearch] = useState('');
  const [editingReviewCode, setEditingReviewCode] = useState<string | null>(null);
  const [editingReviewQuantity, setEditingReviewQuantity] = useState('');
  const [notice, setNotice] = useState('Importe o modelo vazio da empresa para começar.');
  const [cameraOpen, setCameraOpen] = useState(false);
  const [cameraForSearch, setCameraForSearch] = useState(false);
  const [cameraUsed, setCameraUsed] = useState(false);
  const [autoReopenCamera, setAutoReopenCamera] = useState(false);
  const [currentStep, setCurrentStep] = useState(1);
  const [cameraStatus, setCameraStatus] = useState('Aguardando câmera');
  const [cameraZoom, setCameraZoom] = useState(1);
  const [flashEnabled, setFlashEnabled] = useState(false);
  const [scanSoundEnabled, setScanSoundEnabled] = useState(true);
  const [quickCollectEnabled, setQuickCollectEnabled] = useState(false);
  const [quickCollectQuantity, setQuickCollectQuantity] = useState(1);
  const [editingQuickCollectQuantity, setEditingQuickCollectQuantity] = useState(false);
  const [scanPulse, setScanPulse] = useState(0);
  const audioContextRef = useRef<AudioContext | null>(null);
  const videoRef = useRef<HTMLVideoElement>(null);
  const scannerRef = useRef<IScannerControls | null>(null);
  const lastCameraScan = useRef<{ code: string; at: number } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const quantityRef = useRef<HTMLInputElement>(null);
  const shortcutHoldTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const shortcutLongPressHandled = useRef(false);
  const removeAllHoldTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const removeAllLongPressHandled = useRef(false);
  const syncInFlightRef = useRef(false);

  const recordEvent = (barcode: string, quantity: number, eventType: 'scan' | 'adjustment' = 'scan', description = '') => {
    if (!sharedCollection || !barcode || !Number.isFinite(quantity) || quantity === 0) return;
    let deviceId = 'device-unknown';
    try {
      deviceId = localStorage.getItem('coletaflow.device-id') || crypto.randomUUID();
      localStorage.setItem('coletaflow.device-id', deviceId);
    } catch { /* o identificador do dispositivo é apenas auxiliar */ }
    const event={ eventId: crypto.randomUUID(), barcode, quantity, eventType, scannedAt: new Date().toISOString(), deviceId, ...(description.trim()?{description:description.trim()}: {}) };
    const next=[...scanEventsRef.current,event];
    scanEventsRef.current=next;
    setScanEvents(next);
    void persistEvents(next).catch(()=>setNotice('Não foi possível gravar a fila offline. Verifique o armazenamento deste dispositivo antes de continuar.'));
  };
  const finishParticipation = async () => {
    if (!sharedCollection || sharedCollection.isOwner || syncInFlightRef.current) return;
    if (!navigator.onLine) { setNotice('Conecte-se à internet para sincronizar e finalizar sua participação.'); return; }
    syncInFlightRef.current = true;
    setFinishBusy(true);
    setNotice('Preparando as bipagens salvas neste aparelho…');
    const readResponse = async (response: Response, action: string) => {
      const raw = await response.text();
      let data: any;
      try { data = raw ? JSON.parse(raw) : {}; }
      catch { throw new Error(`${action}: servidor respondeu HTTP ${response.status} em formato inesperado.`); }
      if (!response.ok) throw new Error(data.error || `${action}: erro HTTP ${response.status}.`);
      return data;
    };
    try {
      await persistEventsRef.current;
      const pending = [...scanEventsRef.current];
      for (let offset = 0; offset < pending.length; offset += 500) {
        const batch = pending.slice(offset, offset + 500);
        setNotice(`Enviando bipagens ${Math.min(offset + batch.length, pending.length)} de ${pending.length}…`);
        const syncResponse = await fetch('/api/shared/sync', {
          method: 'POST',
          headers: { Authorization: `Bearer ${sharedCollection.token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ collectionId: sharedCollection.id, events: batch })
        });
        const syncData = await readResponse(syncResponse, 'Falha ao enviar bipagens');
        const received = Number(syncData.received);
        const accounted = Number(syncData.inserted) + Number(syncData.duplicates);
        if (received !== batch.length || accounted !== batch.length) {
          throw new Error(`O servidor confirmou ${Number.isFinite(received) ? received : 0} de ${batch.length} bipagem(ns). Os dados continuam salvos neste aparelho; tente finalizar novamente.`);
        }
      }
      setNotice('Bipagens recebidas. Confirmando sua participação…');
      const response = await fetch('/api/shared/finish', {
        method: 'POST',
        headers: { Authorization: `Bearer ${sharedCollection.token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ collectionId: sharedCollection.id })
      });
      const data = await readResponse(response, 'Falha ao finalizar participação');
      if (data.participant?.status !== 'finished') throw new Error('O servidor não confirmou a finalização. Suas bipagens continuam salvas neste aparelho.');
      await persistEvents([]);
      scanEventsRef.current = [];
      setScanEvents([]);
      setNotice('Participação finalizada e confirmada pelo servidor.');
      onFinished?.();
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Falha ao finalizar participação. Seus dados continuam salvos neste aparelho.');
    } finally {
      syncInFlightRef.current = false;
      setFinishBusy(false);
    }
  };
  const clearCurrentCounts = () => {
    Object.entries(counts).forEach(([code, quantity]) => recordEvent(code, -quantity, 'adjustment'));
    setCounts({});
  };
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

  const entries = Object.entries(counts).filter(([code]) => !search || normalize(code).includes(normalize(search)) || normalize(findProductDescription(code)).includes(normalize(search)));
  const totalCodes = Object.keys(counts).length;
  const catalogCodeSet = new Set<string>();
  if (sheet && range) {
    for (let row = headerRow + 1; row <= range.e.r; row++) {
      for (const column of columns.code) {
        const catalogCode = cellText(sheet[XLSX.utils.encode_cell({ r: row, c: range.s.c + column })]).trim();
        if (catalogCode) catalogCodeSet.add(catalogCode);
      }
    }
  }
  const notFoundProducts = Object.keys(counts).filter(code => !catalogCodeSet.has(code)).length;

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
        const session = await readSavedSession(sessionKey);
        if (!active) return;
        const pendingEvents=await readSavedEvents(sessionKey).catch(()=>null);
        const restoredEvents=pendingEvents??(session?.scanEvents||[]);
        scanEventsRef.current=restoredEvents;
        setScanEvents(restoredEvents);
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
          setColumns(session.columns ? { ...session.columns, code: Array.isArray(session.columns.code) ? session.columns.code : [session.columns.code as unknown as number] } : { code: [0], quantity: 1, description: -1 });
          setCounts(session.counts || {});
          setHasProductList(session.hasProductList !== false);
          setCurrentStep(Math.min(4, Math.max(1, Number(session.currentStep ?? 1))));
          setPendingCode(session.pendingCode || '');
          setQuantityValue(session.quantityValue || '');
          setAutoReopenCamera(!!session.autoReopenCamera);
          setCameraZoom(Number(session.cameraZoom ?? 1));
          setQuickCollectEnabled(!!session.quickCollectEnabled);
          setQuickCollectQuantity(Math.max(1, Math.floor(Number(session.quickCollectQuantity ?? 1) || 1)));
          setNotice('Sessão recuperada automaticamente. Você pode continuar de onde parou.');
        } else if (sessionKey === 'active') {
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
  }, [sessionKey]);

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
          scanEvents,
          hasProductList,
          currentStep,
          pendingCode,
          quantityValue,
          autoReopenCamera,
          cameraZoom,
          quickCollectEnabled,
          quickCollectQuantity
        };
        void writeSavedSession(snapshot, sessionKey).catch(() => setNotice('Não foi possível salvar a sessão neste dispositivo. Verifique o espaço disponível.'));
      } catch {
        setNotice('Não foi possível salvar a sessão neste dispositivo. Verifique o espaço disponível.');
      }
    }, 250);
    return () => window.clearTimeout(timer);
  }, [sessionRestored, sessionKey, fileName, sheetName, headerRow, columns, counts, scanEvents, currentStep, pendingCode, quantityValue, autoReopenCamera, cameraZoom, quickCollectEnabled, quickCollectQuantity, hasProductList]);

  // As bipagens ficam somente neste aparelho durante a coleta.
  // O envio acontece em lote exclusivamente em finishParticipation.

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
      if (cameraForSearch) {
        setSearch(code);
        setNotice(`Busca pela câmera: código ${code}.`);
        setCameraOpen(false);
        setCameraForSearch(false);
      } else {
        receiveCode(code, true);
        if (!quickCollectEnabled) setCameraOpen(false);
      }
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
  }, [cameraOpen, pendingCode, quickCollectEnabled, quickCollectQuantity, cameraForSearch]);

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
      setHasProductList(true);
      setPendingDescription('');
      setColumns({
        code: [detectedCodeColumn],
        quantity: findColumn(['quantidade', 'qtd', 'contagem', 'unidades'], 1),
        description: detectedDescriptionColumn >= 0 && detectedDescriptionColumn !== detectedCodeColumn ? detectedDescriptionColumn : -1
      });
      if (!sameTemplate) clearCurrentCounts();
      setPendingCode('');
      setNotice(sameTemplate ? '' : 'Modelo vazio importado. Agora é só bipar o código e informar a quantidade.');
      setCurrentStep(1);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'Não foi possível abrir o modelo.');
    }
  };

  const parsedSheetExists = (bytes: Uint8Array, name: string) => {
    try { return !!XLSX.read(bytes, { type: 'array' }).Sheets[name]; } catch { return false; }
  };
  const continueFromMapping = async () => {
    if(sharedCollection?.isOwner){
      // Com participantes vinculados, o mapeamento já está bloqueado.
      // O dono deve conseguir retomar a coleta sem tentar gravá-lo novamente.
      if(sharedCollection.mappingLocked && sharedCollection.template?.mapping){
        setCurrentStep(2);
        return;
      }
      if(!sharedCollection.accountToken){setNotice('Sessão do dono inválida. Entre novamente.');return;}
      setMappingBusy(true);
      try{
        const response=await fetch('/api/shared/mapping',{method:'PUT',headers:{Authorization:`Bearer ${sharedCollection.accountToken}`,'Content-Type':'application/json'},body:JSON.stringify({collectionId:sharedCollection.id,mapping:{sheetName,headerRow,codeColumns:columns.code,quantityColumn:columns.quantity,descriptionColumn:columns.description,hasProductList}})});
        const data=await response.json();if(!response.ok)throw new Error(data.error||'Não foi possível salvar o mapeamento');
        setNotice('Mapeamento confirmado. Agora os usuários podem entrar nesta coleta.');
        sharedCollection.onMappingSaved?.(data.mapping);
      }catch(error){setNotice(error instanceof Error?error.message:'Falha ao salvar o mapeamento.');return;}
      finally{setMappingBusy(false);}
    }
    setCurrentStep(2);
  };

  useEffect(() => {
    const template = sharedCollection?.template;
    if (!sharedCollection || !template?.base64 || !sessionRestored || workbook) return;
    try {
      const binary = atob(template.base64);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
      const file = new File([bytes], template.fileName, { type: template.mimeType || 'application/octet-stream' });
      void (async()=>{
        await loadWorkbook(file);
        if(template.mapping&&parsedSheetExists(bytes,template.mapping.sheetName)){
          setSheetName(template.mapping.sheetName);
          setHeaderRow(template.mapping.headerRow);
          setColumns({code:template.mapping.codeColumns,quantity:template.mapping.quantityColumn,description:template.mapping.descriptionColumn});
          setHasProductList(template.mapping.hasProductList !== false);
          // O participante recebe o modelo já configurado pelo dono e começa diretamente na coleta.
          if(!sharedCollection.isOwner)setCurrentStep(2);
        } else if(!sharedCollection.isOwner) {
          setNotice('Aguardando o dono configurar e liberar o modelo. Assim que ele concluir a etapa 1, entre novamente na coleta.');
        }
      })();
    } catch {
      setNotice('Não foi possível abrir o modelo compartilhado. Peça ao dono para reenviar a planilha.');
    }
  }, [sharedCollection?.id, sharedCollection?.template?.sha256, sessionRestored, workbook]);

  const findProductRow = (code: string) => {
    if (!sheet || !range) return -1;
    const codeColumns = columns.code.map(column => range.s.c + column);
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const rowCodes = codeColumns.map(column => cellText(sheet[XLSX.utils.encode_cell({ r, c: column })]).trim());
      if (rowCodes.includes(code.trim())) return r;
    }
    return -1;
  };

  const findProductDescription = (code: string) => {
    if (sheet && range && columns.description >= 0) {
      const productRow = findProductRow(code);
      if (productRow >= 0) return cellText(sheet[XLSX.utils.encode_cell({ r: productRow, c: range.s.c + columns.description })]);
    }
    return [...scanEvents].reverse().find(event => event.barcode === code && event.description?.trim())?.description || '';
  };

  const playScanFeedback = () => {
    setScanPulse(value => value + 1);
    if (!scanSoundEnabled) return;
    try {
      const AudioContextClass = window.AudioContext;
      if (!AudioContextClass) return;
      const context = audioContextRef.current ?? new AudioContextClass();
      audioContextRef.current = context;
      if (context.state === 'suspended') void context.resume();
      const oscillator = context.createOscillator();
      const gain = context.createGain();
      oscillator.type = 'sine'; oscillator.frequency.setValueAtTime(1450, context.currentTime);
      gain.gain.setValueAtTime(0.0001, context.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.16, context.currentTime + 0.012);
      gain.gain.exponentialRampToValueAtTime(0.0001, context.currentTime + 0.105);
      oscillator.connect(gain); gain.connect(context.destination);
      oscillator.start(); oscillator.stop(context.currentTime + 0.11);
    } catch { /* áudio pode estar bloqueado pelo navegador */ }
  };

  const receiveCode = (raw: string, fromCamera = false) => {
    const code = raw.trim();
    if (!code) return;
    if (fromCamera) {
      const last = lastCameraScan.current;
      if (last?.code === code && Date.now() - last.at < 1600) return;
      lastCameraScan.current = { code, at: Date.now() };
    }
    const knownProduct = findProductRow(code) >= 0;
    const needsDescription = hasProductList && !knownProduct;
    if (fromCamera && quickCollectEnabled && !needsDescription) {
      playScanFeedback();
      recordEvent(code, quickCollectQuantity, 'scan');
      setCounts(previous => ({ ...previous, [code]: (previous[code] ?? 0) + quickCollectQuantity }));
      setScanValue('');
      setNotice(`Coleta relâmpago: código ${code} · +${quickCollectQuantity} unidade(s). Próxima leitura...`);
      setCameraStatus('Coleta registrada · aponte para o próximo código');
      return;
    }
    const productDescription = findProductDescription(code);
    setPendingDescription(needsDescription ? '' : productDescription);
    const catalogCodes = sheet && range ? Array.from({ length: Math.max(0, range.e.r - headerRow) }, (_, index) => {
      const r = headerRow + 1 + index;
      return columns.code.map(column => cellText(sheet[XLSX.utils.encode_cell({ r, c: range.s.c + column })]).trim()).filter(Boolean);
    }).flat() : [];
    const codeExists = catalogCodes.includes(code);
    setPendingCode(code);
    setScanValue('');
    if (fromCamera) playScanFeedback();
    setNotice(productDescription
      ? `${productDescription} · informe a quantidade.`
      : hasProductList && catalogCodes.length > 0 && !codeExists
        ? `Produto não encontrado. Informe a descrição para continuar.`
        : `Código ${code} lido. Informe a quantidade.`);
    requestAnimationFrame(() => quantityRef.current?.focus());
  };

  const confirmCount = (shortcutQuantity?: number) => {
    const code = pendingCode.trim();
    const qty = shortcutQuantity ?? Number(quantityValue);
    if (!code) { setNotice('Bipe ou digite um código primeiro.'); inputRef.current?.focus(); return; }
    const isUnlistedProduct = hasProductList && findProductRow(code) < 0;
    if (isUnlistedProduct && !pendingDescription.trim()) { setNotice('Informe a descrição do produto não encontrado para continuar.'); return; }
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) { setNotice('Informe uma quantidade inteira maior que zero.'); quantityRef.current?.focus(); return; }
    const alreadyCollected = counts[code] ?? 0;
    if (removeMode) {
      if (alreadyCollected <= 0) { setNotice('Ainda não existe quantidade coletada deste produto para retirar.'); quantityRef.current?.focus(); return; }
      if (qty > alreadyCollected) { setNotice(`Você coletou ${alreadyCollected} unidade(s). Não é possível retirar ${qty}.`); quantityRef.current?.focus(); return; }
      const remaining = alreadyCollected - qty;
      recordEvent(code, -qty, 'adjustment');
      setCounts(previous => {
        const next = { ...previous };
        if (remaining === 0) delete next[code];
        else next[code] = remaining;
        return next;
      });
      setNotice(`Retirada registrada: código ${code} · quantidade retirada: ${qty}.`);
    } else {
      recordEvent(code, qty, 'scan', isUnlistedProduct ? pendingDescription : '');
      setCounts(previous => ({ ...previous, [code]: (previous[code] ?? 0) + qty }));
      setNotice(`Registrado: código ${code} · quantidade adicionada: ${qty}.`);
    }
    setRemoveMode(false);
    setPendingCode('');
    setPendingDescription('');
    setQuantityValue('');
    setScanValue('');
    if (autoReopenCamera && cameraUsed) {
      setFlashEnabled(false);
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
    setCameraForSearch(false);
    setFlashEnabled(false);
    setCameraUsed(true);
    setCameraStatus('Preparando leitor de código de barras…');
    setCameraOpen(true);
  };

  const startReviewSearchCamera = () => {
    if (!workbook || pendingCode) return;
    setCameraForSearch(true);
    setFlashEnabled(false);
    setCameraUsed(true);
    setCameraStatus('Preparando leitor para buscar produto…');
    setCameraOpen(true);
  };

  const toggleCameraFlash = async () => {
    const video = videoRef.current;
    const stream = video?.srcObject as MediaStream | null;
    const track = stream?.getVideoTracks()[0];
    if (!track) {
      setNotice('A câmera ainda está iniciando. Tente novamente em instantes.');
      return;
    }
    const capabilities = track.getCapabilities() as MediaTrackCapabilities & { torch?: boolean };
    if (!capabilities.torch) {
      setNotice('O flash não está disponível nesta câmera ou navegador.');
      return;
    }
    const next = !flashEnabled;
    try {
      await track.applyConstraints({ advanced: [{ torch: next } as MediaTrackConstraintSet] });
      setFlashEnabled(next);
      setCameraStatus(next ? 'Flash ligado · aponte para o código de barras' : 'Flash desligado · aponte para o código de barras');
    } catch {
      setNotice('Não foi possível alterar o flash. Verifique se a câmera oferece suporte.');
    }
  };

  const stopCamera = () => {
    setFlashEnabled(false);
    setCameraForSearch(false);
    setCameraOpen(false);
    setCameraStatus('Câmera pausada');
  };

  const getExportGroups = () => {
    if (!workbook || !sheet || !range) return { matched: [] as [string, number][], unmatched: [] as [string, number][] };
    const matched: [string, number][] = [];
    const unmatched: [string, number][] = [];
    for (const entry of Object.entries(counts) as [string, number][]) {
      const [code, qty] = entry;
      let exists = false;
      for (let r = headerRow + 1; r <= range.e.r; r++) {
        const rowCodes = columns.code.map(column => cellText(sheet[XLSX.utils.encode_cell({ r, c: range.s.c + column })]).trim());
        if (rowCodes.includes(code)) { exists = true; break; }
      }
      (exists ? matched : unmatched).push([code, qty]);
    }
    return { matched, unmatched };
  };

  const exportMatchedWorkbook = () => {
    if (!workbook || !sheet || !range) return;
    const { matched } = getExportGroups();
    if (matched.length === 0) { setNotice('Nenhum produto da planilha recebeu contagem para exportar.'); return; }
    const output = XLSX.read(XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }), { type: 'array', cellStyles: true });
    const outputSheet = output.Sheets[sheetName];
    // Zera os valores antigos da coluna de quantidade no arquivo exportado.
    // Mantém a planilha original em memória intacta e preserva a formatação das células.
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const qtyCell = XLSX.utils.encode_cell({ r, c: range.s.c + columns.quantity });
      const previousCell = outputSheet[qtyCell] as XLSX.CellObject | undefined;
      outputSheet[qtyCell] = { t: 's', v: '', ...(previousCell?.s ? { s: previousCell.s } : {}) };
    }
    for (const [code, qty] of matched) {
      for (let r = headerRow + 1; r <= range.e.r; r++) {
        const rowCodes = columns.code.map(column => cellText(sheet[XLSX.utils.encode_cell({ r, c: range.s.c + column })]).trim());
        if (rowCodes.includes(code)) {
          const qtyCell = XLSX.utils.encode_cell({ r, c: range.s.c + columns.quantity });
          outputSheet[qtyCell] = { t: 'n', v: qty };
          break;
        }
      }
    }
    XLSX.writeFile(output, fileName.replace(/\\.(xlsx|xls|csv)$/i, '') + '-coletado.xlsx');
    setNotice(`Planilha principal exportada: ${matched.length} produtos encontrados e ${matched.reduce((sum, [, qty]) => sum + qty, 0)} unidades. Produtos não encontrados ficaram fora deste arquivo.`);
  };

  const exportNotFoundReport = () => {
    const { unmatched } = getExportGroups();
    if (unmatched.length === 0) { setNotice('Nenhum produto não encontrado para gerar relatório.'); return; }
    const report = XLSX.utils.book_new();
    const rows = [['Código lido', 'Quantidade coletada'], ...unmatched.map(([code, qty]) => [code, qty])];
    const reportSheet = XLSX.utils.aoa_to_sheet(rows);
    reportSheet['!cols'] = [{ wch: 24 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(report, reportSheet, 'Não encontrados');
    XLSX.writeFile(report, fileName.replace(/\\.(xlsx|xls|csv)$/i, '') + '-nao-encontrados.xlsx');
    setNotice(`Relatório de não encontrados exportado: ${unmatched.length} códigos.`);
  };

  const resetSession = () => {
    if (!window.confirm('Limpar toda a contagem desta sessão?')) return;
    clearCurrentCounts();
    setPendingCode('');
    setNotice('Contagem reiniciada. O modelo continua carregado.');
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#" aria-label="ColetaFlow início">
          <span className="brand-mark" aria-hidden="true"><FileSpreadsheet size={34} strokeWidth={2.2} /><RefreshCw size={17} strokeWidth={2.8} className="brand-sync" /></span>
          <span className="brand-wordmark"><span>Coleta</span><strong>Flow</strong></span>
        </a>
        <button type="button" className="theme-toggle" onClick={() => setTheme(value => value === "light" ? "dark" : "light")} aria-label={theme === "light" ? "Ativar tema escuro" : "Ativar tema claro"} title={theme === "light" ? "Ativar tema escuro" : "Ativar tema claro"}><span>{theme === "light" ? <Moon size={17} /> : <Sun size={17} />}</span><small>{theme === "light" ? "Escuro" : "Claro"}</small></button>
        <div className="topbar-right"><span className="offline-pill"><span /> CONTAGEM LOCAL</span><span className="version">MVP 0.2</span></div>
      </header>


      <nav className="wizard-progress" aria-label="Etapas da coleta">{[1,2,3,4].map((step) => <div key={step} className={`wizard-step ${currentStep === step ? "current" : currentStep > step ? "completed" : ""}`}><span>{currentStep > step ? "✓" : step}</span><small>{["Modelo","Coleta","Conferência","Finalizar"][step - 1]}</small></div>)}</nav>

      {currentStep === 1 && <section className="workspace wizard-screen">
        <div className="section-heading"><div><span className="step-label">ETAPA 01 DE 04</span><h2>Modelo da empresa</h2><p>Importe a planilha vazia que deverá receber os códigos e quantidades.</p></div>
          {!sharedCollection&&<label className="primary-button upload-button"><FolderUp size={18} /> Importar modelo<input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /></label>}
        </div>
        {workbook && sheet ? (
          <div className="file-card">
            <div className="file-icon"><FileSpreadsheet size={24} /></div>
            <div className="file-info"><strong>{fileName}</strong><span>Modelo carregado · {workbook.SheetNames.length} aba(s) · sem necessidade de cadastro de produtos</span></div>
            <span className="file-status"><Check size={14} /> Pronto</span>
            {!sharedCollection&&<button className="icon-button" title="Remover modelo" onClick={() => { setWorkbook(null); workbookBytesRef.current = null; setFileName(''); setSheetName(''); clearCurrentCounts(); setPendingCode(''); setCurrentStep(1); }}><X size={18} /></button>}
          </div>
        ) : sharedCollection ? <div className="dropzone"><span className="drop-icon"><FolderUp size={25} /></span><strong>Carregando o modelo compartilhado</strong><span>O arquivo original é enviado pelo dono da coleta.</span></div> : <label className="dropzone"><input type="file" accept=".xlsx,.xls,.csv" onChange={event => { const file = event.target.files?.[0]; if (file) void loadWorkbook(file); event.currentTarget.value = ''; }} /><span className="drop-icon"><FolderUp size={25} /></span><strong>Selecione o modelo vazio da empresa</strong><span>Excel (.xlsx, .xls) ou CSV</span></label>}

        {workbook && sheet && (
          <div className="mapping-panel">
            <div className="mapping-title"><div><ShieldCheck size={18} /><strong>Defina onde gravar os dados</strong></div><span>MAPEAMENTO DO MODELO</span></div>
            <div className="mapping-grid">
              <label>Cadastro de produtos<select disabled={!!sharedCollection && (!sharedCollection.isOwner || !!sharedCollection.mappingLocked)} value={hasProductList ? 'yes' : 'no'} onChange={event=>setHasProductList(event.target.value==='yes')}><option value="yes">Planilha contém produtos cadastrados</option><option value="no">Planilha sem cadastro de produtos</option></select></label>
               <label>Aba<select disabled={!!sharedCollection && (!sharedCollection.isOwner || !!sharedCollection.mappingLocked)} value={sheetName} onChange={event => { setSheetName(event.target.value); clearCurrentCounts(); setPendingCode(''); }}>{workbook.SheetNames.map(name => <option key={name} value={name}>{name}</option>)}</select></label>
              <label>Código principal<select disabled={!!sharedCollection && (!sharedCollection.isOwner || !!sharedCollection.mappingLocked)} value={columns.code[0] ?? 0} onChange={event => setColumns(v => ({ ...v, code: [Number(event.target.value), ...(v.code[1] !== undefined && v.code[1] !== Number(event.target.value) ? [v.code[1]] : [])] }))}>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
              <label>Código alternativo (opcional)<select disabled={!!sharedCollection && (!sharedCollection.isOwner || !!sharedCollection.mappingLocked)} value={columns.code[1] ?? -1} onChange={event => setColumns(v => ({ ...v, code: event.target.value === '-1' ? [v.code[0] ?? 0] : [v.code[0] ?? 0, Number(event.target.value)] }))}><option value={-1}>Não usar</option>{headers.map((h,i) => <option value={i} key={i} disabled={i === columns.code[0]}>{h}</option>)}</select></label>
              <label>Coluna da quantidade<select disabled={!!sharedCollection && (!sharedCollection.isOwner || !!sharedCollection.mappingLocked)} value={columns.quantity} onChange={event => setColumns(v => ({ ...v, quantity: Number(event.target.value) }))}>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
              <label>Coluna da descrição<select disabled={!!sharedCollection && (!sharedCollection.isOwner || !!sharedCollection.mappingLocked)} value={columns.description} onChange={event => setColumns(v => ({ ...v, description: Number(event.target.value) }))}><option value={-1}>Não possui descrição</option>{headers.map((h,i) => <option value={i} key={i}>{h}</option>)}</select></label>
            </div>
            <div className="mapping-note">Com cadastro: códigos fora da lista exigem descrição antes de salvar. Sem cadastro: a coleta funciona em modo aberto, sem comparação com catálogo.</div>
          </div>
        )}
        <div className="wizard-navigation wizard-next"><span>Confira o modelo e ajuste as colunas antes de continuar.</span><button className="primary-button" disabled={!workbook || !sheet || mappingBusy} onClick={() => void continueFromMapping()}>{mappingBusy?'Salvando mapeamento…':'Continuar'} <ArrowDownToLine size={17} /></button></div>
      </section>}

      {currentStep === 2 && <section className="workspace collector-section wizard-screen">
        <div className="section-heading"><div><span className="step-label">ETAPA 02 DE 04</span><h2>Coletor</h2></div><span className={workbook ? 'ready-chip' : 'waiting-chip'}>{workbook ? 'MODELO PRONTO' : 'AGUARDANDO MODELO'}</span></div>
          {cameraOpen && createPortal(
            <div className="camera-box">
              <div className="camera-status"><span className="camera-live-dot" />{cameraStatus}</div>
              <video ref={videoRef} muted playsInline autoPlay />
              <div key={scanPulse} className={`camera-guide ${scanPulse ? `scan-guide-pulse-${scanPulse % 2}` : ""}`} aria-hidden="true" />
              <div className="camera-zoom-controls" role="group" aria-label="Zoom da câmera">
                {[1, 2, 3].map(level => <button type="button" key={level} className={cameraZoom === level ? 'active' : ''} onClick={() => setCameraZoom(level)} aria-pressed={cameraZoom === level} aria-label={`Zoom ${level}x`} title={`Zoom ${level}x`}><span>{level}x</span></button>)}
              </div>
              <div className="camera-bottom-controls">
                <button type="button" className={`camera-round-control ${flashEnabled ? 'flash-active' : ''}`} onClick={() => void toggleCameraFlash()} aria-label={flashEnabled ? 'Desligar flash' : 'Ligar flash'} aria-pressed={flashEnabled} title={flashEnabled ? 'Desligar flash' : 'Ligar flash'}><Zap size={21} fill={flashEnabled ? 'currentColor' : 'none'} /></button>
                <button type="button" className={`camera-round-control ${scanSoundEnabled ? 'sound-active' : ''}`} onClick={() => setScanSoundEnabled(value => !value)} aria-label={scanSoundEnabled ? 'Desligar som do bip' : 'Ligar som do bip'} aria-pressed={scanSoundEnabled} title={scanSoundEnabled ? 'Desligar som do bip' : 'Ligar som do bip'}>{scanSoundEnabled ? <Volume2 size={20} /> : <VolumeX size={20} />}</button>
              </div>
              <button className="camera-close" onClick={stopCamera}><X size={18} /> Fechar câmera</button>
            </div>,
            document.body
          )}
          {pendingCode && <div className="quantity-modal-backdrop" role="presentation">
            <section className="quantity-modal" role="dialog" aria-modal="true" aria-labelledby="quantity-modal-title">
              <div className="quantity-modal-heading">
                <div className="quantity-product-info">
                  <div className="quantity-modal-topline"><span className="quantity-modal-eyebrow">{hasProductList && findProductRow(pendingCode) < 0 ? 'PRODUTO NÃO ENCONTRADO' : 'PRODUTO IDENTIFICADO'}</span><span className="quantity-collected">Coletado: <strong>{counts[pendingCode] ?? 0}</strong></span></div>
                  <h3 id="quantity-modal-title">{findProductDescription(pendingCode) || (findProductRow(pendingCode) >= 0 ? 'Produto sem descrição' : 'Produto não encontrado')}</h3>
                  <span className="quantity-product-code">Código: <strong>{pendingCode}</strong></span>
                </div>
                <button className="icon-button quantity-modal-close" title="Cancelar leitura" onClick={() => { setPendingCode(''); setQuantityValue(''); setRemoveMode(false); inputRef.current?.focus(); }}><X size={20} /></button>
              </div>
              <div className="quantity-modal-content">
                <div className="quantity-shortcuts">
                  <button type="button" className="edit-shortcuts-button" onClick={() => setEditingShortcuts(value => !value)}>{editingShortcuts ? 'Concluir edição' : 'Editar atalhos'}</button>
                  <div className="quantity-shortcut-list">{quantityShortcuts.map((amount, index) => <button type="button" className="quantity-shortcut" key={index} onPointerDown={event => { event.preventDefault(); shortcutLongPressHandled.current = false; if (shortcutHoldTimer.current) clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = setTimeout(() => { shortcutLongPressHandled.current = true; confirmCount(amount); shortcutHoldTimer.current = null; }, 500); }} onPointerUp={() => { if (shortcutHoldTimer.current) { clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = null; } }} onPointerLeave={() => { if (!shortcutLongPressHandled.current && shortcutHoldTimer.current) { clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = null; } }} onPointerCancel={() => { if (!shortcutLongPressHandled.current && shortcutHoldTimer.current) { clearTimeout(shortcutHoldTimer.current); shortcutHoldTimer.current = null; } }} onContextMenu={event => event.preventDefault()} onClick={event => { event.preventDefault(); if (shortcutLongPressHandled.current) { shortcutLongPressHandled.current = false; return; } setQuantityValue(String((Number(quantityValue) || 0) + amount)); }}>{amount}</button>)}</div>
                  {editingShortcuts && <div className="edit-shortcuts-grid">{quantityShortcuts.map((amount, index) => <label key={index}>Atalho {index + 1}<input type="number" min="1" step="1" value={amount} onChange={event => setQuantityShortcuts(values => values.map((value, i) => i === index ? Math.max(1, Number(event.target.value) || 1) : value))} /></label>)}</div>}
                </div>
                {hasProductList && findProductRow(pendingCode) < 0 && <label className="quantity-modal-label" style={{display:'block',marginBottom:12}}>Descrição do produto <input type="text" value={pendingDescription} onChange={event=>setPendingDescription(event.target.value)} placeholder="Digite o nome/descrição do produto" maxLength={180} required autoComplete="off" /></label>}
              <div className="quantity-entry-row">
                  <label className="quantity-modal-label">{removeMode ? 'Quantidade a retirar' : 'Quantidade encontrada'}<input ref={quantityRef} type="number" min="1" step="1" value={quantityValue} onChange={event => setQuantityValue(event.target.value)} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); confirmCount(); } }} /></label>
                  <button type="button" className={removeMode ? 'remove-mode-button active' : 'remove-mode-button'} aria-pressed={removeMode} title="Toque para retirar uma quantidade; segure para zerar tudo que foi coletado deste produto" onPointerDown={() => { removeAllLongPressHandled.current = false; if (removeAllHoldTimer.current) clearTimeout(removeAllHoldTimer.current); removeAllHoldTimer.current = setTimeout(() => { removeAllLongPressHandled.current = true; const collected = counts[pendingCode] ?? 0; if (collected > 0) { recordEvent(pendingCode, -collected, 'adjustment'); setCounts(previous => { const next = { ...previous }; delete next[pendingCode]; return next; }); setNotice(`Coleta zerada: ${collected} unidade(s) removida(s) do código ${pendingCode}.`); } else { setNotice('Este produto ainda não possui quantidade coletada.'); } setRemoveMode(false); setQuantityValue(''); removeAllHoldTimer.current = null; }, 700); }} onPointerUp={() => { if (removeAllHoldTimer.current) { clearTimeout(removeAllHoldTimer.current); removeAllHoldTimer.current = null; } }} onPointerLeave={() => { if (removeAllHoldTimer.current) { clearTimeout(removeAllHoldTimer.current); removeAllHoldTimer.current = null; } }} onPointerCancel={() => { if (removeAllHoldTimer.current) { clearTimeout(removeAllHoldTimer.current); removeAllHoldTimer.current = null; } }} onClick={() => { if (removeAllLongPressHandled.current) { removeAllLongPressHandled.current = false; return; } setRemoveMode(value => !value); setQuantityValue(''); requestAnimationFrame(() => quantityRef.current?.focus()); }}>{removeMode ? 'Retirar: ligado' : 'Retirar'}</button>
                </div>
                <div className="quantity-modal-actions">
                  <button className="secondary-button" onClick={() => { setPendingCode(''); setQuantityValue(''); inputRef.current?.focus(); }}>Cancelar</button>
                  <button className={removeMode ? "primary-button remove-confirm-button" : "primary-button"} onClick={() => confirmCount()}>{removeMode ? <><RotateCcw size={17} /> Retirar quantidade</> : <><Check size={17} /> Confirmar quantidade</>}</button>
                </div>
              </div>
            </section>
          </div>}

        <div className={`scanner-panel scan-pulse-${scanPulse % 2}`}>

          <form className="scan-form" onSubmit={event => { event.preventDefault(); receiveCode(scanValue); }}>
            <div className="scan-input-wrap"><Keyboard size={20} /><input ref={inputRef} value={scanValue} onChange={event => setScanValue(event.target.value)} placeholder={!workbook ? 'Importe o modelo primeiro' : pendingCode ? 'Informe a quantidade do código lido' : 'Bipe ou digite o código...'} disabled={!workbook || !!pendingCode} autoComplete="off" autoFocus /><button type="button" className="scan-camera-inline" onClick={startCamera} disabled={!workbook || cameraOpen || !!pendingCode} aria-label="Abrir câmera" title="Ler código com a câmera"><Camera size={21} /></button></div>
          </form>
          <div className="scanner-heading scanner-controls" style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 10 }}>
            <button type="button" className={autoReopenCamera ? 'secondary-button auto-reopen active' : 'secondary-button auto-reopen'} onClick={() => setAutoReopenCamera(value => !value)} aria-pressed={autoReopenCamera} disabled={!workbook || !!pendingCode} title="Ao confirmar a quantidade, inicia a câmera para a próxima leitura."><RefreshCw size={16} /> {autoReopenCamera ? 'Reabertura automática: ligada' : 'Reabertura automática: desligada'}</button>
            <div className="quick-collect-control">
              <button type="button" className={quickCollectEnabled ? 'secondary-button auto-reopen active' : 'secondary-button auto-reopen'} onClick={() => { setQuickCollectEnabled(value => !value); setEditingQuickCollectQuantity(false); }} aria-pressed={quickCollectEnabled} disabled={!workbook || !!pendingCode} title="Registra automaticamente a quantidade definida e mantém a câmera pronta para a próxima leitura."><Zap size={16} /> {quickCollectEnabled ? 'Coleta relâmpago: ligada' : 'Coleta relâmpago: desligada'}</button>
              <div className="quick-collect-quantity"><span>Qtd. <strong>{quickCollectQuantity}</strong></span><button type="button" className="icon-button" onClick={() => setEditingQuickCollectQuantity(value => !value)} aria-label="Editar quantidade da coleta relâmpago" title="Editar quantidade"><Pencil size={15} /></button></div>
            </div>
            {editingQuickCollectQuantity && <div className="quick-collect-editor">
              <label style={{ fontSize: 13 }}>Quantidade automática <input type="number" min="1" step="1" value={quickCollectQuantity} onChange={event => setQuickCollectQuantity(Math.max(1, Math.floor(Number(event.target.value) || 1)))} style={{ width: 88, marginLeft: 6 }} /></label>
              <button type="button" className="secondary-button" onClick={() => setEditingQuickCollectQuantity(false)}>Concluir</button>
            </div>}
          </div>

          <div key={scanPulse} className={`scan-feedback ${scanPulse ? 'scan-feedback-active' : ''}`}>{notice}</div>
        </div>
        <div className="stats-grid collector-stats">
          <div className="stat-card"><span>Produtos coletados</span><strong>{totalCodes}</strong><small>Produtos com quantidade registrada</small></div>
          <div className="stat-card warning-stat"><span>Produtos não encontrados</span><strong>{notFoundProducts}</strong><small>{notFoundProducts ? 'Códigos fora da planilha' : 'Nenhum produto fora da planilha'}</small></div>
        </div>
        <div className="wizard-navigation"><button className="secondary-button" onClick={() => { stopCamera(); setCurrentStep(1); }}>Voltar</button><button className="primary-button" disabled={!!pendingCode || totalCodes === 0} onClick={() => { stopCamera(); setCurrentStep(3); }}>Continuar para conferência <ArrowDownToLine size={17} /></button></div>
      </section>}

      {currentStep === 3 && <section className="workspace results-section wizard-screen">
          {cameraOpen && createPortal(
            <div className="camera-box">
              <div className="camera-status"><span className="camera-live-dot" />{cameraStatus}</div>
              <video ref={videoRef} muted playsInline autoPlay />
              <div key={scanPulse} className={`camera-guide ${scanPulse ? `scan-guide-pulse-${scanPulse % 2}` : ""}`} aria-hidden="true" />
              <div className="camera-zoom-controls" role="group" aria-label="Zoom da câmera">
                {[1, 2, 3].map(level => <button type="button" key={level} className={cameraZoom === level ? 'active' : ''} onClick={() => setCameraZoom(level)} aria-pressed={cameraZoom === level} aria-label={`Zoom ${level}x`} title={`Zoom ${level}x`}><span>{level}x</span></button>)}
              </div>
              <div className="camera-bottom-controls">
                <button type="button" className={`camera-round-control ${flashEnabled ? 'flash-active' : ''}`} onClick={() => void toggleCameraFlash()} aria-label={flashEnabled ? 'Desligar flash' : 'Ligar flash'} aria-pressed={flashEnabled} title={flashEnabled ? 'Desligar flash' : 'Ligar flash'}><Zap size={21} fill={flashEnabled ? 'currentColor' : 'none'} /></button>
                <button type="button" className={`camera-round-control ${scanSoundEnabled ? 'sound-active' : ''}`} onClick={() => setScanSoundEnabled(value => !value)} aria-label={scanSoundEnabled ? 'Desligar som do bip' : 'Ligar som do bip'} aria-pressed={scanSoundEnabled} title={scanSoundEnabled ? 'Desligar som do bip' : 'Ligar som do bip'}>{scanSoundEnabled ? <Volume2 size={20} /> : <VolumeX size={20} />}</button>
              </div>
              <button className="camera-close" onClick={stopCamera}><X size={18} /> Fechar câmera</button>
            </div>,
            document.body
          )}
        <div className="section-heading"><div><span className="step-label">ETAPA 03 DE 04</span><h2>Resumo da coleta</h2><p>Confira somente os códigos e as quantidades registradas.</p></div><button className="text-button" onClick={resetSession} disabled={!workbook}><RotateCcw size={16} /> Reiniciar contagem</button></div>
        <div className="results-toolbar">
          <div className="blind-mode-note"><ShieldCheck size={17} /> SOMENTE CÓDIGO + QUANTIDADE</div>
          <div className="search-box review-search-box"><Search size={17} /><input value={search} onChange={event => setSearch(event.target.value)} placeholder="Buscar por código ou descrição" /><button type="button" className="scan-camera-inline" onClick={startReviewSearchCamera} disabled={!workbook || cameraOpen || !!pendingCode} aria-label="Buscar código com a câmera" title="Buscar código com a câmera"><Camera size={20}/></button></div>
        </div>
        <div className="product-list">
          {!workbook ? <div className="empty-state"><Package size={28} /><strong>Aguardando modelo</strong><span>Importe a planilha vazia para iniciar a coleta.</span></div> :
            entries.slice(0,100).map(([code, qty], index) => <div className="product-row" key={code}><div className="product-number">{index + 1}</div><div className="product-details"><strong>{findProductDescription(code) || 'Descrição não identificada'}</strong><span>CÓDIGO: {code}</span></div><div className="review-quantity"><strong>{qty}</strong><button type="button" className="icon-button" aria-label={`Editar quantidade de ${code}`} title="Editar quantidade" onClick={() => { setEditingReviewCode(code); setEditingReviewQuantity(String(qty)); }}><Pencil size={17}/></button></div>{editingReviewCode === code && <div className="review-quantity-editor"><label>Nova quantidade <input type="number" min="0" step="1" value={editingReviewQuantity} onChange={event => setEditingReviewQuantity(event.target.value)} /></label><button type="button" className="primary-button" onClick={() => { const value = Math.max(0, Math.floor(Number(editingReviewQuantity) || 0)); const previous = counts[code] ?? 0; if (value !== previous) recordEvent(code, value - previous, 'adjustment'); setCounts(current => { const next = { ...current }; if (value === 0) delete next[code]; else next[code] = value; return next; }); setEditingReviewCode(null); }}>Salvar</button><button type="button" className="text-button" onClick={() => setEditingReviewCode(null)}>Cancelar</button></div>}</div>)}
          {workbook && entries.length === 0 && <div className="empty-state"><Search size={24} /><strong>Nenhum código registrado</strong><span>Depois da leitura, os códigos e quantidades aparecerão aqui.</span></div>}
          {entries.length > 100 && <div className="list-footnote">Mostrando os primeiros 100 códigos. Use a busca para localizar outro.</div>}
        </div>
        <div className="wizard-navigation"><button className="secondary-button" onClick={() => setCurrentStep(2)}>Voltar à coleta</button><button className="primary-button" disabled={totalCodes === 0} onClick={() => setCurrentStep(4)}>Continuar para finalizar <ArrowDownToLine size={17} /></button></div>
      </section>}

      {currentStep === 4 && <section className="export-bar wizard-screen">
        <div><span className="step-label">ETAPA 04 DE 04</span><strong>Exportar relatórios da coleta</strong><small>Os produtos encontrados entram na planilha original preenchida. Os códigos não encontrados ficam em um relatório separado.</small></div>
        <div className="export-actions export-actions-split">
          <button className="secondary-button" onClick={() => setCurrentStep(3)}>Voltar</button>
          <button className="primary-button" onClick={exportMatchedWorkbook} disabled={!workbook || getExportGroups().matched.length === 0 || !!pendingCode}><ArrowDownToLine size={18} /> Baixar planilha coletada ({getExportGroups().matched.length})</button>
          <button className="secondary-button not-found-download" onClick={exportNotFoundReport} disabled={getExportGroups().unmatched.length === 0 || !!pendingCode}><ArrowDownToLine size={18} /> Baixar não encontrados ({getExportGroups().unmatched.length})</button>
        </div>
      </section>}
      {sharedCollection && !sharedCollection.isOwner && <div className="mapping-note" style={{marginTop:14}}>{scanEvents.length > 0 ? `${scanEvents.length} evento(s) salvos neste aparelho. Serão enviados quando você finalizar.` : navigator.onLine ? 'Bipagens enviadas. Você pode finalizar sua participação.' : 'Sem internet. Suas bipagens permanecem salvas neste aparelho.'}</div>}
      {sharedCollection && !sharedCollection.isOwner && <div className="wizard-navigation"><button className="primary-button" disabled={finishBusy} onClick={() => void finishParticipation()}>{finishBusy ? 'Sincronizando e finalizando…' : 'Finalizar minha participação'}</button></div>}
      <footer><span>ColetaFlow <b>•</b> Bipe. Informe. Exporte.</span><span><Smartphone size={14} /> Coleta sem cadastro prévio</span></footer>
    </main>
  );
}


type SharedCollectionState = { id:string; title:string; accessCode:string; role:string; userId:string };
type SharedMapping = { sheetName:string; headerRow:number; codeColumns:number[]; quantityColumn:number; descriptionColumn:number; hasProductList?:boolean };
type SharedTemplate = { fileName:string; mimeType:string; sha256:string; base64:string; mapping?:SharedMapping|null };
function SharedCollections({token,isOwner,userId}:{token:string;isOwner:boolean;userId:string}) {
  const [title,setTitle]=useState('Inventário compartilhado');
  const [accessCode,setAccessCode]=useState('');
  const [templateFile,setTemplateFile]=useState<File|null>(null);
  const [templateInputKey,setTemplateInputKey]=useState(0);
  const [sharedTemplate,setSharedTemplate]=useState<SharedTemplate|null>(null);
  const [active,setActive]=useState<SharedCollectionState|null>(()=>{try{const saved=JSON.parse(localStorage.getItem('coletaflow.shared.active')||'null');return saved?.userId===userId?{id:saved.id,title:saved.title,accessCode:saved.accessCode||'',role:saved.role||'user',userId}:null;}catch{return null;}});
  const [progress,setProgress]=useState<any>(null);
  const [ownerCollections,setOwnerCollections]=useState<any[]>([]);
  const [busy,setBusy]=useState(false);
  const [exportBusy,setExportBusy]=useState(false);
  const [closeBusy,setCloseBusy]=useState(false);
  const [error,setError]=useState('');
  const [notice,setNotice]=useState('');
  const [users,setUsers]=useState<any[]>([]);
  const [newUser,setNewUser]=useState({name:'',company:'',email:'',password:''});
  const [userBusy,setUserBusy]=useState(false);
  const [userMessage,setUserMessage]=useState('');
  useEffect(()=>{try{const saved=JSON.parse(localStorage.getItem('coletaflow.shared.active')||'null');if(saved&&saved.userId!==userId)localStorage.removeItem('coletaflow.shared.active');}catch{try{localStorage.removeItem('coletaflow.shared.active');}catch{}}},[userId]);
  const auth={Authorization:`Bearer ${token}`,'Content-Type':'application/json'};
  const activate=(data:any)=>{
    const next={id:data.collection.id,title:data.collection.title,accessCode:data.collection.access_code||'',role:data.participant?.role||'owner',userId};
    localStorage.setItem('coletaflow.shared.active',JSON.stringify(next));setActive(next);setProgress(null);setSharedTemplate(null);setError('');setNotice('Você entrou na coleta. O código pode ser compartilhado com os demais participantes.');
  };
  const create=async()=>{
    if(!templateFile){setError('Selecione a planilha original antes de criar a coleta.');return;}
    if(templateFile.size>3*1024*1024){setError('A planilha deve ter no máximo 3 MB.');return;}
    setBusy(true);setError('');setNotice('');
    try{
      const bytes=new Uint8Array(await templateFile.arrayBuffer());
      let binary='';
      for(let offset=0;offset<bytes.length;offset+=8192)binary+=String.fromCharCode(...bytes.subarray(offset,offset+8192));
      const templateBase64=btoa(binary);
      const payload={title,templateFileName:templateFile.name,templateMimeType:templateFile.type||'application/octet-stream',templateBase64};
      const r=await fetch('/api/shared/create',{method:'POST',headers:auth,body:JSON.stringify(payload)});
      const d=await r.json();if(!r.ok)throw new Error(d.error||'Falha ao criar coleta');
      setTemplateFile(null);activate({...d,participant:{role:'owner'}});
    }catch(e){setError(e instanceof Error?e.message:'Erro de conexão');}finally{setBusy(false);}
  };
  const join=async()=>{
    setBusy(true);setError('');setNotice('');
    try{const r=await fetch('/api/shared/join',{method:'POST',headers:auth,body:JSON.stringify({accessCode})});const d=await r.json();if(!r.ok)throw new Error(d.error||'Falha ao entrar na coleta');activate(d);}
    catch(e){setError(e instanceof Error?e.message:'Erro de conexão');}finally{setBusy(false);}
  };
  useEffect(()=>{if(!isOwner)return;let alive=true;const loadUsers=async()=>{try{const r=await fetch('/api/shared/users',{headers:auth,cache:'no-store'});const d=await r.json();if(r.ok&&alive)setUsers(d.users||[]);}catch{}};void loadUsers();return()=>{alive=false;};},[token,isOwner]);
  const loadOwnerCollections=async()=>{
    if(!isOwner)return;
    try{const r=await fetch('/api/shared/collections',{headers:auth,cache:'no-store'});const d=await r.json();if(r.ok)setOwnerCollections(d.collections||[]);}
    catch{/* A lista será atualizada na próxima tentativa. */}
  };
  useEffect(()=>{if(isOwner)void loadOwnerCollections();},[token,isOwner]);

  const addUser=async()=>{setUserBusy(true);setUserMessage('');try{const r=await fetch('/api/shared/users',{method:'POST',headers:auth,body:JSON.stringify(newUser)});const d=await r.json();if(!r.ok)throw new Error(d.error||'Não foi possível cadastrar');setUsers(v=>[d.user,...v]);setNewUser({name:'',company:'',email:'',password:''});setUserMessage('Usuário cadastrado.');}catch(e){setUserMessage(e instanceof Error?e.message:'Erro de conexão');}finally{setUserBusy(false);}};
  useEffect(()=>{
    if(!active){setSharedTemplate(null);return;}
    let alive=true;
    const loadTemplate=async()=>{
      try{
        const r=await fetch(`/api/shared/template?collectionId=${encodeURIComponent(active.id)}`,{headers:{Authorization:`Bearer ${token}`},cache:'no-store'});
        const d=await r.json();if(!r.ok)throw new Error(d.error||'Não foi possível baixar o modelo da coleta');
        if(alive)setSharedTemplate(d);
      }catch(e){if(alive)setError(e instanceof Error?e.message:'Falha ao baixar o modelo');}
    };
    void loadTemplate();
    return()=>{alive=false;};
  },[active?.id,token]);
  useEffect(()=>{
    if(!active)return;
    let alive=true;
    const load=async()=>{try{const r=await fetch(`/api/shared/progress?collectionId=${encodeURIComponent(active.id)}`,{headers:{Authorization:`Bearer ${token}`},cache:'no-store'});const d=await r.json();if(!r.ok)throw new Error(d.error||'Falha ao atualizar progresso');if(alive)setProgress(d);}
      catch(e){if(alive)setError(e instanceof Error?e.message:'Falha de conexão');}};
    void load();const timer=window.setInterval(load,3000);
    return()=>{alive=false;window.clearInterval(timer);};
  },[active?.id,token]);
  const leave=()=>{localStorage.removeItem('coletaflow.shared.active');setActive(null);setProgress(null);setError('');setNotice('Coleta desconectada deste dispositivo.');};
  const closeCollection=async()=>{
    if(!active)return;
    const activeUsers=(progress?.participants||[]).filter((p:any)=>p.role==='user'&&p.status==='active').length;
    if(activeUsers>0){setError(`Ainda existem ${activeUsers} usuário(s) trabalhando.`);return;}
    setCloseBusy(true);setError('');
    try{
      const r=await fetch('/api/shared/close',{method:'POST',headers:auth,body:JSON.stringify({collectionId:active.id})});
      const d=await r.json();if(!r.ok)throw new Error(d.error||'Não foi possível encerrar');
      setProgress((current:any)=>current?{...current,collection:{...current.collection,...d.collection}}:current);
      setNotice('Coleta encerrada. Agora você pode baixar a planilha consolidada.');
    }catch(e){setError(e instanceof Error?e.message:'Falha ao encerrar a coleta.');}
    finally{setCloseBusy(false);}
  };
  const exportUnmatched=async()=>{
    if(!active)return;
    setExportBusy(true);setError('');
    try{
      const r=await fetch(`/api/shared/export-unmatched?collectionId=${encodeURIComponent(active.id)}`,{headers:auth,cache:'no-store'});
      const d=await r.json();if(!r.ok)throw new Error(d.error||'Não foi possível exportar os códigos não encontrados');
      const binary=atob(d.base64),bytes=new Uint8Array(binary.length);
      for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
      const blob=new Blob([bytes],{type:d.mimeType||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const url=URL.createObjectURL(blob),anchor=document.createElement('a');
      anchor.href=url;anchor.download=d.fileName||'codigos-nao-encontrados.xlsx';document.body.appendChild(anchor);anchor.click();anchor.remove();
      window.setTimeout(()=>URL.revokeObjectURL(url),1500);
      setNotice(`Relatório gerado: ${d.unmatchedCodes} códigos não encontrados, ${d.totalUnits} unidades.`);
    }catch(e){setError(e instanceof Error?e.message:'Falha ao gerar relatório de não encontrados.');}
    finally{setExportBusy(false);}
  };

  const exportConsolidated=async()=>{
    if(!active)return;
    setExportBusy(true);setError('');
    try{
      const r=await fetch(`/api/shared/export?collectionId=${encodeURIComponent(active.id)}`,{headers:auth,cache:'no-store'});
      const d=await r.json();if(!r.ok)throw new Error(d.error||'Não foi possível exportar');
      const binary=atob(d.base64),bytes=new Uint8Array(binary.length);
      for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
      const blob=new Blob([bytes],{type:d.mimeType||'application/octet-stream'});
      const url=URL.createObjectURL(blob),anchor=document.createElement('a');
      anchor.href=url;anchor.download=d.fileName||'coletaflow-coletado.xlsx';document.body.appendChild(anchor);anchor.click();anchor.remove();
      window.setTimeout(()=>URL.revokeObjectURL(url),1500);
      setNotice(`Planilha consolidada gerada: ${d.matchedCodes} códigos encontrados, ${d.unmatchedCodes} não encontrados e ${d.totalUnits} unidades.`);
    }catch(e){setError(e instanceof Error?e.message:'Falha ao gerar a planilha consolidada.');}
    finally{setExportBusy(false);}
  };

  return <section className="workspace" style={{maxWidth:980,margin:'10px auto'}}>
    {isOwner&&<div className="mapping-panel" style={{marginBottom:18}}><div className="mapping-title"><strong>Cadastrar usuário</strong><span>CONTROLE DO DONO</span></div><div className="mapping-grid"><label>Nome<input value={newUser.name} onChange={e=>setNewUser(v=>({...v,name:e.target.value}))}/></label><label>Empresa<input value={newUser.company} onChange={e=>setNewUser(v=>({...v,company:e.target.value}))}/></label><label>E-mail<input type="email" value={newUser.email} onChange={e=>setNewUser(v=>({...v,email:e.target.value}))}/></label><label>Senha inicial<input type="password" minLength={8} value={newUser.password} onChange={e=>setNewUser(v=>({...v,password:e.target.value}))}/></label></div><button className="primary-button" style={{marginTop:12}} disabled={userBusy||!newUser.name||!newUser.company||!newUser.email||newUser.password.length<8} onClick={()=>void addUser()}>{userBusy?'Salvando…':'Cadastrar usuário'}</button>{userMessage&&<p className="mapping-note">{userMessage}</p>}<h3>Usuários cadastrados</h3><div className="product-list">{users.map(u=><div className="product-row" key={u.id}><div className="product-details"><strong>{u.display_name}</strong><span>{u.email} · {u.company_name} · {u.status==='active'?'ATIVO':'BLOQUEADO'}</span></div><button className="secondary-button" onClick={async()=>{const status=u.status==='active'?'disabled':'active';const r=await fetch('/api/shared/users',{method:'PATCH',headers:auth,body:JSON.stringify({id:u.id,status})});if(r.ok)setUsers(v=>v.map(x=>x.id===u.id?{...x,status}:x));}}> {u.status==='active'?'Bloquear':'Liberar'} </button></div>)}</div></div>}
    {isOwner&&!active&&ownerCollections.length>0&&<div className="mapping-panel" style={{marginBottom:18}}><div className="mapping-title"><strong>Coletas anteriores</strong><span>{ownerCollections.length} REGISTRO(S)</span></div><div className="product-list">{ownerCollections.map((collection:any)=><div className="product-row" key={collection.id}><div className="product-details"><strong>{collection.title}</strong><span>{collection.status==='open'?'ABERTA':collection.status==='completed'?'CONCLUÍDA':collection.status.toUpperCase()} · {collection.participant_count} usuário(s) · {collection.finished_count} finalizado(s) · {collection.event_count} bipagem(ns)</span><small>Código: {collection.access_code} · {collection.last_scan_at?new Date(collection.last_scan_at).toLocaleString('pt-BR'):'Sem bipagens'}</small></div><button className="secondary-button" onClick={()=>activate({collection,participant:{role:'owner'}})}>Abrir</button></div>)}</div></div>}
    {active && <CollectorApp sharedCollection={{id:active.id,token,userId,accountToken:token,isOwner,mappingLocked:(progress?.participants||[]).some((p:any)=>p.role==='user'),template:sharedTemplate||undefined,onMappingSaved:(mapping)=>setSharedTemplate(current=>current?{...current,mapping}:current)}} onFinished={()=>{localStorage.removeItem('coletaflow.shared.active');setActive(null);setProgress(null);setNotice('Participação finalizada. O dono já pode acompanhar sua conclusão.');}} />}
     {active ? <div className="mapping-panel">
      <div className="mapping-title"><div><strong>{progress?.collection?.title||active.title}</strong></div><span>{progress?.collection?.status==='open'?'COLETA ABERTA':(progress?.collection?.status||'CONECTANDO').toUpperCase()}</span></div>
      {active.accessCode&&<div className="stat-card" style={{margin:'14px 0'}}><span>Código para convidar participantes</span><strong style={{letterSpacing:3}}>{active.accessCode}</strong><small>Compartilhe este código com quem vai participar.</small><button className="secondary-button" style={{marginTop:10}} onClick={()=>{void navigator.clipboard?.writeText(active.accessCode);setNotice('Código copiado.');}}>Copiar código</button></div>}
      {isOwner&&<div className="stats-grid"><div className="stat-card"><span>Usuários na coleta</span><strong>{progress?.participants?.filter((p:any)=>p.role==='user').length??'—'}</strong><small>Entraram usando o código</small></div><div className="stat-card"><span>Finalizados</span><strong>{progress?.participants?`${progress.participants.filter((p:any)=>p.role==='user'&&p.status==='finished').length}/${progress.participants.filter((p:any)=>p.role==='user').length}`:'—'}</strong><small>Contagens enviadas ao concluir</small></div></div>}
       {isOwner&&<div className="mapping-panel" style={{marginTop:16}}><div className="mapping-title"><strong>Planilha consolidada</strong><span>{sharedTemplate?.mapping?'MAPEAMENTO CONFIRMADO':'AGUARDANDO MAPEAMENTO'}</span></div><p className="mapping-note">{sharedTemplate?.mapping?'O modelo original e as colunas confirmadas estão prontos para gerar o arquivo final.':'Abra a etapa Modelo no coletor, confira as colunas e aperte Continuar para liberar a entrada dos usuários.'}</p>{progress?.collection?.status==='open'&&<><button className="secondary-button" style={{marginBottom:10}} disabled={closeBusy||!progress||(progress?.participants||[]).filter((p:any)=>p.role==='user').length===0||(progress?.participants||[]).some((p:any)=>p.role==='user'&&p.status!=='finished')} onClick={()=>void closeCollection()}>{closeBusy?'Encerrando…':'Encerrar coleta e liberar relatórios'}</button><small style={{display:'block',marginBottom:10}}>{(progress?.participants||[]).filter((p:any)=>p.role==='user').length===0?'Aguardando usuários entrarem pelo código.':(progress?.participants||[]).some((p:any)=>p.role==='user'&&p.status!=='finished')?'Aguardando todos os usuários enviarem suas contagens.':'Todos finalizaram. Encerre a coleta para liberar os dois relatórios.'}</small></>}{progress?.collection?.status==='completed'&&<p className="mapping-note"><strong>Coleta finalizada.</strong> Todos os usuários enviaram suas contagens. Os relatórios estão liberados.</p>}<button className="primary-button" disabled={exportBusy||!sharedTemplate?.mapping||progress?.collection?.status!=='completed'} onClick={()=>void exportConsolidated()}>{exportBusy?'Gerando planilha…':'Baixar planilha coletada'}</button><button className="secondary-button" style={{marginTop:10}} disabled={exportBusy||!sharedTemplate?.mapping||progress?.collection?.status!=='completed'} onClick={()=>void exportUnmatched()}>{exportBusy?'Gerando relatório…':'Baixar códigos não encontrados'}</button><small style={{display:'block',marginTop:7}}>Relatório separado com código, descrição (quando disponível no modelo) e quantidade coletada.</small></div>}
      {isOwner&&<><h3 style={{marginTop:22}}>Equipe</h3><div className="product-list">{(progress?.participants||[]).map((p:any)=><div className="product-row" key={p.id}><div className="product-details"><strong>{p.display_name}</strong><span>{p.role==='owner'?'DONO':'USUÁRIO'} · {p.status==='active'?'TRABALHANDO':p.status==='finished'?'FINALIZOU':'BLOQUEADO'} · {p.event_count||0} evento(s) · {Number(p.total_units||0)} unidade(s) · {p.distinct_codes||0} código(s)</span><small>{p.last_scan_at?new Date(p.last_scan_at).toLocaleString('pt-BR'):'Ainda sem bipagens'}</small></div></div>)}{!progress&&<p>Carregando dados da coleta…</p>}</div>
<p className="mapping-note">O painel atualiza os participantes e o status de finalização. As bipagens ficam no aparelho de cada usuário e são enviadas somente quando ele finaliza.</p></>}
      <button className="secondary-button" onClick={leave}>Desconectar desta coleta</button>
    </div> : <div className="mapping-grid">
      {isOwner&&<div className="mapping-panel"><div className="mapping-title"><strong>Criar nova coleta</strong><span>DONO</span></div><label>Nome da coleta<input value={title} onChange={e=>setTitle(e.target.value)} maxLength={120}/></label><label>Planilha original da empresa<input key={templateInputKey} type="file" accept=".xlsx,.xls,.csv" onChange={e=>setTemplateFile(e.target.files?.[0]||null)}/></label>{templateFile&&<div className="mapping-note" style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:10}}><span style={{minWidth:0,overflowWrap:'anywhere'}}>{templateFile.name} · {(templateFile.size/1024/1024).toFixed(2)} MB</span><button type="button" className="secondary-button" style={{padding:'7px 10px',flexShrink:0}} disabled={busy} onClick={()=>{setTemplateFile(null);setTemplateInputKey(value=>value+1);}}>Remover planilha</button></div>}<button className="primary-button" style={{marginTop:14}} disabled={busy||!title.trim()||!templateFile} onClick={()=>void create()}>Criar coleta e compartilhar modelo</button></div>}
      <div className="mapping-panel"><div className="mapping-title"><strong>Entrar em uma coleta</strong><span>PARTICIPANTE</span></div><label>Código de acesso<input value={accessCode} onChange={e=>setAccessCode(e.target.value.toUpperCase().replace(/[^A-F0-9]/g,'').slice(0,8))} placeholder="Ex.: A1B2C3D4" maxLength={8}/></label><button className="secondary-button" style={{marginTop:14}} disabled={busy||accessCode.length!==8} onClick={()=>void join()}>Entrar na coleta</button></div>
    </div>}
    {busy&&<p>Aguarde…</p>}{error&&<div className="account-error">{error}</div>}{notice&&<p className="mapping-note">{notice}</p>}
  </section>;
}


function AccountGate() {

  const [email,setEmail]=useState('');
  const [password,setPassword]=useState('');
  const [busy,setBusy]=useState(false);
  const [error,setError]=useState('');
  const [user,setUser]=useState<{id:string;name:string;email:string;company:string;role:string;isOwner:boolean}|null>(null);
  const [token,setToken]=useState(()=>{try{return localStorage.getItem('coletaflow.auth.token')||'';}catch{return ''; }});
  const [localMode,setLocalMode]=useState(false);
  const [showOwnerSetup,setShowOwnerSetup]=useState(false);
  const [setupName,setSetupName]=useState('');
  const [setupEmail,setSetupEmail]=useState('');
  const [setupPassword,setSetupPassword]=useState('');
  const [setupKey,setSetupKey]=useState('');
  const [setupBusy,setSetupBusy]=useState(false);
  const [setupError,setSetupError]=useState('');
  useEffect(()=>{
    if(!token)return;
    let alive=true;
    fetch('/api/shared/me',{headers:{Authorization:`Bearer ${token}`}})
      .then(async r=>{const d=await r.json();if(!r.ok)throw new Error(d.error||'Sessão expirada');return d;})
      .then(d=>{if(alive)setUser({id:d.user.id,name:d.user.display_name,email:d.user.email,company:d.user.company_name,role:d.user.role,isOwner:!!d.user.is_owner});})
      .catch(()=>{try{localStorage.removeItem('coletaflow.auth.token');}catch{} if(alive){setToken('');setUser(null);}});
    return()=>{alive=false;};
  },[token]);
  const submit=async(e:FormEvent)=>{
    e.preventDefault();setBusy(true);setError('');
    try{
      const endpoint='/api/shared/login';
      const payload={email,password};
      const response=await fetch(endpoint,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
      const data=await response.json();
      if(!response.ok)throw new Error(data.error||'Não foi possível continuar');
      localStorage.setItem('coletaflow.auth.token',data.token);
      setUser({id:data.user.id,name:data.user.name,email:data.user.email,company:data.user.company,role:data.user.role,isOwner:!!data.user.is_owner});
      setToken(data.token);
    }catch(err){setError(err instanceof Error?err.message:'Erro de conexão');}
    finally{setBusy(false);}
  };
  const setupOwner=async(e:FormEvent)=>{
    e.preventDefault();setSetupBusy(true);setSetupError('');
    try{
      const response=await fetch('/api/shared/bootstrap-owner',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({setupKey,email:setupEmail,name:setupName,password:setupPassword})});
      const data=await response.json();
      if(!response.ok){
        const diagnostic=data.diagnostic;
        const details=diagnostic ? ` Diagnóstico: variáveis ausentes: ${diagnostic.missing?.length ? diagnostic.missing.join(', ') : 'nenhuma'}. Hyperdrive disponível: ${diagnostic.hasHyperdrive ? 'sim' : 'não'}.` : '';
        throw new Error((data.error||'Não foi possível configurar o dono')+details);
      }
      localStorage.setItem('coletaflow.auth.token',data.token);
      setUser({id:data.user.id,name:data.user.display_name,email:data.user.email,company:data.user.company_name,role:'owner',isOwner:true});
      setToken(data.token);
    }catch(err){setSetupError(err instanceof Error?err.message:'Falha de conexão');}
    finally{setSetupBusy(false);}
  };
  if(localMode)return <div className="account-shell"><header className="account-top"><strong>ColetaFlow · modo local</strong><span>Funciona sem conta e mantém a coleta neste dispositivo.</span><button onClick={()=>setLocalMode(false)}>Acessar conta</button></header><CollectorApp /></div>;
  if(user)return <div className="account-shell"><header className="account-top"><strong>ColetaFlow</strong><span>{user.name} · {user.company}</span><button onClick={async()=>{try{await fetch('/api/shared/logout',{method:'POST',headers:{Authorization:`Bearer ${token}`}});}catch{} try{localStorage.removeItem('coletaflow.auth.token');}catch{} setToken('');setUser(null);}}>Sair</button></header><SharedCollections token={token} isOwner={!!user.isOwner} userId={user.id}/></div>;
  return <main className="account-page"><section className="account-card">
    <div className="account-brand"><span className="account-mark">CF</span><div><strong>ColetaFlow</strong><small>Coleta colaborativa de inventário</small></div></div>
    <div className="account-intro"><span>ACESSO SEGURO</span><h1>Entre na sua conta</h1><p>Acesse as coletas liberadas para você.</p></div>
    <form className="account-form" onSubmit={submit}>
      <label>E-mail<input required type="email" value={email} onChange={e=>setEmail(e.target.value)} autoComplete="email" placeholder="voce@empresa.com"/></label>
      <label>Senha<input required type="password" minLength={8} value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password" placeholder="Mínimo de 8 caracteres"/></label>
      {error&&<div className="account-error">{error}</div>}
      <button className="primary-button" type="submit" disabled={busy}>{busy?'Aguarde...':'Entrar'}</button>
    </form>
    {!showOwnerSetup ? <button type="button" className="secondary-button" style={{width:'100%',marginTop:12}} onClick={()=>setShowOwnerSetup(true)}>Primeiro acesso do dono</button> : <form className="account-form" onSubmit={setupOwner} style={{marginTop:18}}>
      <div className="account-intro"><span>CONFIGURAÇÃO ÚNICA</span><h1>Conta do dono</h1><p>Disponível somente enquanto a conta principal ainda não foi criada.</p></div>
      <label>Chave de configuração<input required type="password" value={setupKey} onChange={e=>setSetupKey(e.target.value)} autoComplete="off"/></label>
      <label>Nome<input required value={setupName} onChange={e=>setSetupName(e.target.value)} maxLength={80}/></label>
      <label>E-mail autorizado<input required type="email" value={setupEmail} onChange={e=>setSetupEmail(e.target.value)} autoComplete="email"/></label>
      <label>Senha inicial<input required type="password" minLength={10} maxLength={128} value={setupPassword} onChange={e=>setSetupPassword(e.target.value)} autoComplete="new-password"/></label>
      {setupError&&<div className="account-error">{setupError}</div>}
      <button className="primary-button" type="submit" disabled={setupBusy}>{setupBusy?'Configurando…':'Criar conta do dono'}</button>
      <button type="button" className="text-button" onClick={()=>{setShowOwnerSetup(false);setSetupError('');}}>Voltar ao login</button>
    </form>}
    <button type="button" className="secondary-button" style={{width:'100%',marginTop:12}} onClick={()=>setLocalMode(true)}>Continuar no modo local (sem conta)</button>
    <small className="account-security">Suas credenciais são protegidas e a senha não é armazenada em texto aberto.</small>
  </section></main>;
}
export default function App() {
  return <AccountGate />;
}
