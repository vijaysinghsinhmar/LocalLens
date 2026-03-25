import React, {
  useState, useCallback, useEffect, useRef, useMemo
} from 'react';
import {
  Search, FolderOpen, Loader2, Database,
  Download, Zap, ChevronDown, FileSpreadsheet,
  FileCode, FileText, X, SlidersHorizontal,
  ArrowUpDown, CheckCircle2, AlertCircle, Cpu,
  BarChart3
} from 'lucide-react';
import { SearchResult, FileMetadata, SortKey, SortDir } from './types';
import { workerScript } from './worker';

// ── Constants ────────────────────────────────────────────────────────────────
const MAX_WORKERS   = Math.min(navigator.hardwareConcurrency || 4, 16);
const ROW_HEIGHT    = 72;
const BUFFER_ROWS   = 10;
const DEBOUNCE_MS   = 150;
const FLUSH_MS      = 80;
const SUPPORTED_EXT = ['xlsx', 'xls', 'csv', 'txt'] as const;

// ── Helpers ──────────────────────────────────────────────────────────────────
const fmt = (n: number) =>
  n >= 1_000_000 ? (n / 1_000_000).toFixed(1) + 'M'
  : n >= 1_000   ? (n / 1_000).toFixed(1) + 'K'
  : String(n);

const fmtBytes = (b: number) =>
  b >= 1_073_741_824 ? (b / 1_073_741_824).toFixed(1) + ' GB'
  : b >= 1_048_576   ? (b / 1_048_576).toFixed(1) + ' MB'
  : b >= 1_024       ? (b / 1_024).toFixed(1) + ' KB'
  : b + ' B';

const extIcon = (name: string) => {
  const e = name.split('.').pop()?.toLowerCase();
  if (e === 'csv') return <FileCode className="w-4 h-4 text-emerald-400" />;
  if (e === 'txt') return <FileText className="w-4 h-4 text-amber-400" />;
  return <FileSpreadsheet className="w-4 h-4 text-sky-400" />;
};

// ── Main Component ────────────────────────────────────────────────────────────
const App: React.FC = () => {
  // Core state
  const [query,      setQuery]      = useState('');
  const [files,      setFiles]      = useState<FileMetadata[]>([]);
  const [allResults, setAllResults] = useState<SearchResult[]>([]);
  const [isSearching,setIsSearching]= useState(false);
  const [isExporting,setIsExporting]= useState(false);

  // Search options
  const [exactMatch,  setExactMatch]  = useState(false);
  const [fuzzySearch, setFuzzySearch] = useState(true);
  const [activeTypes, setActiveTypes] = useState<string[]>([...SUPPORTED_EXT]);
  const [sortKey,     setSortKey]     = useState<SortKey>('relevance');
  const [sortDir,     setSortDir]     = useState<SortDir>('desc');
  const [showFilters, setShowFilters] = useState(false);

  // Progress
  const [progress, setProgress] = useState({ scanned: 0, total: 0, errors: [] as string[], totalRows: 0 });
  const [elapsed,  setElapsed]  = useState(0);

  // Expanded detail
  const [expandedId,   setExpandedId]   = useState<string | null>(null);
  const [expandedData, setExpandedData] = useState<Record<string, unknown> | null>(null);
  const [loadingDetail, setLoadingDetail] = useState(false);

  // Virtual scroll
  const [scrollTop,       setScrollTop]       = useState(0);
  const [containerHeight, setContainerHeight] = useState(600);
  const scrollRef = useRef<HTMLDivElement>(null);

  // Internal refs
  const workers        = useRef<Worker[]>([]);
  const searchTimeout  = useRef<number | null>(null);
  const activeSearchId = useRef<number>(0);
  const startTime      = useRef<number>(0);
  const resultsBuffer  = useRef<SearchResult[]>([]);
  const flushTimeout   = useRef<number | null>(null);
  const timerInterval  = useRef<number | null>(null);

  // ── Container height ──────────────────────────────────────────────────────
  useEffect(() => {
    const obs = new ResizeObserver(() => {
      if (scrollRef.current) setContainerHeight(scrollRef.current.clientHeight);
    });
    if (scrollRef.current) obs.observe(scrollRef.current);
    return () => obs.disconnect();
  }, []);

  // ── Spawn workers ─────────────────────────────────────────────────────────
  useEffect(() => {
    const blob = new Blob([workerScript], { type: 'application/javascript' });
    const url  = URL.createObjectURL(blob);
    workers.current = Array.from({ length: MAX_WORKERS }, () => new Worker(url));
    return () => {
      workers.current.forEach(w => w.terminate());
      URL.revokeObjectURL(url);
    };
  }, []);

  // ── Flush buffer → state ──────────────────────────────────────────────────
  const flushResults = useCallback(() => {
    if (resultsBuffer.current.length) {
      const items = [...resultsBuffer.current];
      resultsBuffer.current = [];
      setAllResults(prev => [...prev, ...items]);
    }
    flushTimeout.current = null;
  }, []);

  const queueResults = useCallback((items: SearchResult[]) => {
    resultsBuffer.current.push(...items);
    if (!flushTimeout.current)
      flushTimeout.current = window.setTimeout(flushResults, FLUSH_MS);
  }, [flushResults]);

  // ── Search ─────────────────────────────────────────────────────────────────
  const performSearch = useCallback((
    term: string, exact: boolean, fuzzy: boolean, types: string[]
  ) => {
    const sid = ++activeSearchId.current;

    // Clear any pending flush
    if (flushTimeout.current) { clearTimeout(flushTimeout.current); flushTimeout.current = null; }
    resultsBuffer.current = [];

    setAllResults([]);
    setExpandedId(null);
    setExpandedData(null);
    setElapsed(0);

    const filtered = files.filter(f => types.includes(f.type));
    if (!filtered.length) return;

    setIsSearching(true);
    startTime.current = Date.now();
    setProgress({ scanned: 0, total: filtered.length, errors: [], totalRows: 0 });

    if (timerInterval.current) clearInterval(timerInterval.current);
    timerInterval.current = window.setInterval(
      () => setElapsed(Date.now() - startTime.current), 50
    );

    let idx = 0, done = 0;

    const dispatch = (w: Worker) => {
      if (idx >= filtered.length || sid !== activeSearchId.current) return;
      const f = filtered[idx++];

      const prev = w.onmessage;
      w.onmessage = (e) => {
        const { action, payload } = e.data;
        if (action === 'MATCH_CHUNK') {
          queueResults(payload.matches);
        } else if (action === 'FILE_COMPLETE') {
          done++;
          setProgress(p => ({ ...p, scanned: done, totalRows: p.totalRows + payload.totalRows }));
          if (done === filtered.length) {
            flushResults();
            setIsSearching(false);
            if (timerInterval.current) { clearInterval(timerInterval.current); timerInterval.current = null; }
            setElapsed(Date.now() - startTime.current);
          } else {
            dispatch(w);
          }
        } else if (action === 'FILE_ERROR') {
          done++;
          setProgress(p => ({ ...p, scanned: done, errors: [...p.errors, `${f.name}: ${payload.error}`] }));
          if (done === filtered.length) {
            flushResults();
            setIsSearching(false);
            if (timerInterval.current) { clearInterval(timerInterval.current); timerInterval.current = null; }
          } else {
            dispatch(w);
          }
        } else if (action === 'ROW_DETAIL_RESULT') {
          setExpandedData(payload.rowData);
        } else if (action === 'EXPORT_READY') {
          const a = document.createElement('a');
          a.href = URL.createObjectURL(payload.blob);
          a.download = `locallens_export_${Date.now()}.csv`;
          a.click();
          setIsExporting(false);
        }
        if (prev) prev.call(w, e);
      };

      w.postMessage({
        action: 'PROCESS_FILE',
        payload: { fileId: f.id, blob: f.blob, type: f.type, name: f.name, path: f.path, query: term, exactMatch: exact, fuzzy }
      });
    };

    workers.current.forEach(dispatch);
  }, [files, flushResults, queueResults]);

  const triggerSearch = useCallback((val: string) => {
    if (searchTimeout.current) clearTimeout(searchTimeout.current);
    searchTimeout.current = window.setTimeout(
      () => performSearch(val, exactMatch, fuzzySearch, activeTypes),
      DEBOUNCE_MS
    );
  }, [performSearch, exactMatch, fuzzySearch, activeTypes]);

  // Re-run when options change
  useEffect(() => {
    if (files.length && query) triggerSearch(query);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [exactMatch, fuzzySearch, activeTypes]);

  // ── Folder input ──────────────────────────────────────────────────────────
  const handleFolder = (e: React.ChangeEvent<HTMLInputElement>) => {
    const raw = e.target.files;
    if (!raw) return;
    const list: FileMetadata[] = [];
    for (let i = 0; i < raw.length; i++) {
      const f   = raw[i];
      const ext = f.name.split('.').pop()?.toLowerCase() ?? '';
      if (SUPPORTED_EXT.includes(ext as typeof SUPPORTED_EXT[number])) {
        list.push({
          id:   `${f.name}-${f.size}-${f.lastModified}`,
          name: f.name,
          path: (f as File & { webkitRelativePath?: string }).webkitRelativePath || f.name,
          type: ext,
          blob: f,
          size: f.size,
        });
      }
    }
    setFiles(list);
    setAllResults([]);
    setProgress({ scanned: 0, total: 0, errors: [], totalRows: 0 });
  };

  // ── Expand row detail ─────────────────────────────────────────────────────
  const handleExpand = (res: SearchResult) => {
    if (expandedId === res.id) {
      setExpandedId(null);
      setExpandedData(null);
      return;
    }
    setExpandedId(res.id);
    setExpandedData(null);
    // No loadingDetail spinner needed — worker answers from parsed cache synchronously
    workers.current[0].postMessage({
      action: 'FETCH_ROW_DETAIL',
      payload: { fileId: res.fileId, rowNumber: res.rowNumber, sheetName: res.sheetName }
    });
  };

  // ── Sort + virtual slice ──────────────────────────────────────────────────
  const sorted = useMemo(() => {
    const arr = [...allResults];
    const mul = sortDir === 'asc' ? 1 : -1;
    if (sortKey === 'relevance') arr.sort((a, b) => mul * ((b.score ?? 0) - (a.score ?? 0)));
    else if (sortKey === 'fileName') arr.sort((a, b) => mul * a.fileName.localeCompare(b.fileName));
    else arr.sort((a, b) => mul * (a.rowNumber - b.rowNumber));
    return arr;
  }, [allResults, sortKey, sortDir]);

  const visible = useMemo(() => {
    const start = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - BUFFER_ROWS);
    const end   = Math.min(sorted.length, Math.ceil((scrollTop + containerHeight) / ROW_HEIGHT) + BUFFER_ROWS);
    return { items: sorted.slice(start, end), startIdx: start };
  }, [sorted, scrollTop, containerHeight]);

  const pct = progress.total ? Math.round((progress.scanned / progress.total) * 100) : 0;

  // ── Type toggle ───────────────────────────────────────────────────────────
  const toggleType = (t: string) =>
    setActiveTypes(prev => prev.includes(t) ? prev.filter(x => x !== t) : [...prev, t]);

  const cycleSort = (key: SortKey) => {
    if (sortKey === key) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortKey(key); setSortDir('desc'); }
  };

  // ── File size total ───────────────────────────────────────────────────────
  const totalSize = files.reduce((s, f) => s + f.size, 0);

  return (
    <div className="flex flex-col h-screen overflow-hidden" style={{
      background: '#08090b',
      color: '#e2e8f0',
      fontFamily: "'DM Mono', 'Fira Code', 'Cascadia Code', monospace"
    }}>

      {/* ── Header ── */}
      <header style={{
        background: 'linear-gradient(180deg, #0f1117 0%, #0b0d13 100%)',
        borderBottom: '1px solid #1e2433',
        padding: '12px 20px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: 12,
        flexShrink: 0
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{
            width: 32, height: 32, borderRadius: 8,
            background: 'linear-gradient(135deg, #3b82f6 0%, #6366f1 100%)',
            display: 'flex', alignItems: 'center', justifyContent: 'center'
          }}>
            <Zap size={16} color="#fff" />
          </div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, letterSpacing: '0.08em', color: '#f1f5f9' }}>
              LOCAL<span style={{ color: '#3b82f6' }}>LENS</span>
            </div>
            <div style={{ fontSize: 9, color: '#475569', letterSpacing: '0.15em', marginTop: -2 }}>
              BLAZING-FAST LOCAL FILE SEARCH
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          {/* Stats pills */}
          {files.length > 0 && (
            <div style={{ display: 'flex', gap: 6 }}>
              <Pill icon={<Database size={10} />} value={`${files.length} files`} />
              <Pill icon={<Cpu size={10} />} value={`${MAX_WORKERS} workers`} />
              <Pill icon={<BarChart3 size={10} />} value={fmtBytes(totalSize)} />
            </div>
          )}

          {/* Load folder */}
          <label style={{
            display: 'flex', alignItems: 'center', gap: 6,
            padding: '6px 14px', borderRadius: 8, cursor: 'pointer',
            background: '#161b27', border: '1px solid #2a3347',
            fontSize: 11, fontWeight: 600, color: '#94a3b8'
          }}>
            <FolderOpen size={13} />
            {files.length ? 'Change Folder' : 'Open Folder'}
            <input type="file" className="hidden"
              // @ts-ignore
              webkitdirectory="" directory="" multiple
              onChange={handleFolder}
            />
          </label>

          {/* Export */}
          {allResults.length > 0 && (
            <button
              disabled={isExporting}
              onClick={() => { setIsExporting(true); workers.current[0].postMessage({ action: 'GENERATE_EXPORT', payload: {} }); }}
              style={{
                display: 'flex', alignItems: 'center', gap: 6,
                padding: '6px 14px', borderRadius: 8,
                background: isExporting ? '#1e3a5f' : 'linear-gradient(135deg, #2563eb, #4f46e5)',
                border: 'none', cursor: isExporting ? 'not-allowed' : 'pointer',
                fontSize: 11, fontWeight: 600, color: '#fff'
              }}
            >
              {isExporting ? <Loader2 size={12} className="animate-spin" /> : <Download size={12} />}
              Export CSV
            </button>
          )}
        </div>
      </header>

      {/* ── Search bar ── */}
      <div style={{
        background: '#0b0d13',
        borderBottom: '1px solid #1e2433',
        padding: '14px 20px',
        flexShrink: 0
      }}>
        {/* Input row */}
        <div style={{ position: 'relative', marginBottom: 10 }}>
          <Search size={16} style={{
            position: 'absolute', left: 14, top: '50%', transform: 'translateY(-50%)',
            color: isSearching ? '#3b82f6' : '#475569',
            transition: 'color 0.2s'
          }} />
          <input
            value={query}
            disabled={files.length === 0}
            onChange={e => { setQuery(e.target.value); triggerSearch(e.target.value); }}
            placeholder={files.length ? 'Search across all files…' : 'Open a folder to start searching'}
            style={{
              width: '100%', boxSizing: 'border-box',
              padding: '10px 14px 10px 40px',
              background: '#161b27', border: '1px solid #2a3347',
              borderRadius: 10, color: '#e2e8f0',
              fontSize: 14, outline: 'none',
              fontFamily: 'inherit',
              transition: 'border-color 0.15s'
            }}
            onFocus={e => e.target.style.borderColor = '#3b82f6'}
            onBlur={e => e.target.style.borderColor = '#2a3347'}
          />
          {query && (
            <button onClick={() => { setQuery(''); setAllResults([]); }}
              style={{ position: 'absolute', right: 14, top: '50%', transform: 'translateY(-50%)',
                background: 'none', border: 'none', cursor: 'pointer', color: '#475569' }}>
              <X size={14} />
            </button>
          )}
          {isSearching && (
            <div style={{
              position: 'absolute', right: query ? 36 : 14, top: '50%', transform: 'translateY(-50%)',
              display: 'flex', alignItems: 'center', gap: 6
            }}>
              <span style={{ fontSize: 10, color: '#3b82f6', fontWeight: 700 }}>{pct}%</span>
              <Loader2 size={13} color="#3b82f6" style={{ animation: 'spin 1s linear infinite' }} />
            </div>
          )}
        </div>

        {/* Controls row */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
          {/* Mode toggles */}
          <ToggleBtn active={fuzzySearch} onClick={() => setFuzzySearch(!fuzzySearch)} label="FUZZY" />
          <ToggleBtn active={exactMatch}  onClick={() => setExactMatch(!exactMatch)}   label="EXACT" />

          <div style={{ width: 1, height: 14, background: '#2a3347' }} />

          {/* File type filters */}
          {SUPPORTED_EXT.map(t => (
            <ToggleBtn key={t} active={activeTypes.includes(t)}
              onClick={() => toggleType(t)} label={t.toUpperCase()} small />
          ))}

          <div style={{ width: 1, height: 14, background: '#2a3347' }} />

          {/* Sort */}
          <button onClick={() => setShowFilters(!showFilters)} style={{
            display: 'flex', alignItems: 'center', gap: 5,
            background: showFilters ? '#1e2433' : 'none',
            border: '1px solid ' + (showFilters ? '#3b82f6' : 'transparent'),
            borderRadius: 6, padding: '3px 8px', cursor: 'pointer',
            color: '#64748b', fontSize: 10, fontWeight: 700
          }}>
            <SlidersHorizontal size={11} /> SORT
          </button>

          {showFilters && (
            <div style={{ display: 'flex', gap: 6 }}>
              {(['relevance', 'fileName', 'rowNumber'] as SortKey[]).map(k => (
                <button key={k} onClick={() => cycleSort(k)} style={{
                  display: 'flex', alignItems: 'center', gap: 4,
                  padding: '3px 8px', borderRadius: 6, cursor: 'pointer',
                  background: sortKey === k ? '#162040' : 'none',
                  border: '1px solid ' + (sortKey === k ? '#3b82f6' : '#2a3347'),
                  color: sortKey === k ? '#60a5fa' : '#64748b',
                  fontSize: 10, fontWeight: 700
                }}>
                  <ArrowUpDown size={10} />
                  {k === 'relevance' ? 'SCORE' : k === 'fileName' ? 'FILE' : 'ROW'}
                  {sortKey === k && (sortDir === 'asc' ? ' ↑' : ' ↓')}
                </button>
              ))}
            </div>
          )}

          {/* Live stats */}
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 12, alignItems: 'center' }}>
            {allResults.length > 0 && (
              <span style={{ fontSize: 10, color: '#3b82f6', fontWeight: 700 }}>
                {fmt(allResults.length)} matches
              </span>
            )}
            {elapsed > 0 && !isSearching && (
              <span style={{ fontSize: 10, color: '#475569' }}>
                {(elapsed / 1000).toFixed(2)}s
              </span>
            )}
            {isSearching && elapsed > 0 && (
              <span style={{ fontSize: 10, color: '#f59e0b', fontWeight: 700 }}>
                {(elapsed / 1000).toFixed(1)}s…
              </span>
            )}
          </div>
        </div>
      </div>

      {/* ── Results list ── */}
      <div
        ref={scrollRef}
        onScroll={e => setScrollTop(e.currentTarget.scrollTop)}
        style={{ flex: 1, overflowY: 'auto', background: '#08090b' }}
      >
        {files.length === 0 ? (
          <EmptyState />
        ) : sorted.length === 0 && !isSearching && query ? (
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 12, opacity: 0.4 }}>
            <Search size={40} color="#475569" />
            <div style={{ fontSize: 13, color: '#475569' }}>No matches for "{query}"</div>
          </div>
        ) : sorted.length === 0 && !isSearching ? (
          <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 8, opacity: 0.35 }}>
            <Database size={48} color="#3b82f6" />
            <div style={{ fontSize: 14, fontWeight: 700, letterSpacing: '0.1em', color: '#94a3b8' }}>
              {files.length} FILE{files.length !== 1 ? 'S' : ''} LOADED
            </div>
            <div style={{ fontSize: 11, color: '#475569' }}>Type to search</div>
          </div>
        ) : (
          <div style={{ height: sorted.length * ROW_HEIGHT, position: 'relative' }}>
            <div style={{ transform: `translateY(${visible.startIdx * ROW_HEIGHT}px)` }}>
              {visible.items.map((res) => (
                <ResultRow
                  key={res.id}
                  res={res}
                  isExpanded={expandedId === res.id}
                  expandedData={expandedData}
                  loadingDetail={loadingDetail}
                  onExpand={() => handleExpand(res)}
                  query={query}
                />
              ))}
            </div>
          </div>
        )}
      </div>

      {/* ── Footer ── */}
      <footer style={{
        background: '#0b0d13',
        borderTop: '1px solid #1e2433',
        padding: '6px 20px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexShrink: 0
      }}>
        <div style={{ display: 'flex', gap: 16, alignItems: 'center' }}>
          <FootStat label="WORKERS" value={MAX_WORKERS} />
          <FootStat label="FILES" value={files.length} />
          {progress.totalRows > 0 && <FootStat label="ROWS SCANNED" value={fmt(progress.totalRows)} />}
          {progress.errors.length > 0 && (
            <span style={{ fontSize: 9, color: '#ef4444', fontWeight: 700, letterSpacing: '0.1em' }}>
              ⚠ {progress.errors.length} ERR
            </span>
          )}
        </div>
        <div style={{ fontSize: 9, color: '#1e3a5f', letterSpacing: '0.1em', fontWeight: 700 }}>
          LOCALLENS v2.0 · ALL PROCESSING LOCAL
        </div>
      </footer>
    </div>
  );
};

// ── Sub-components ────────────────────────────────────────────────────────────

const Pill: React.FC<{ icon: React.ReactNode; value: string }> = ({ icon, value }) => (
  <div style={{
    display: 'flex', alignItems: 'center', gap: 4,
    padding: '3px 8px', borderRadius: 20,
    background: '#0f1520', border: '1px solid #1e2e4a',
    fontSize: 10, color: '#64748b', fontWeight: 600
  }}>
    {icon} {value}
  </div>
);

const ToggleBtn: React.FC<{ active: boolean; onClick: () => void; label: string; small?: boolean }> = ({ active, onClick, label, small }) => (
  <button onClick={onClick} style={{
    padding: small ? '2px 7px' : '3px 10px',
    borderRadius: 6, cursor: 'pointer',
    background: active ? '#162040' : 'none',
    border: '1px solid ' + (active ? '#3b82f6' : '#2a3347'),
    color: active ? '#60a5fa' : '#475569',
    fontSize: 9, fontWeight: 800, letterSpacing: '0.12em',
    transition: 'all 0.15s'
  }}>
    {active && <span style={{ marginRight: 3 }}>✓</span>}{label}
  </button>
);

const FootStat: React.FC<{ label: string; value: number | string }> = ({ label, value }) => (
  <span style={{ fontSize: 9, color: '#334155', fontWeight: 700, letterSpacing: '0.1em' }}>
    {label}: <span style={{ color: '#475569' }}>{value}</span>
  </span>
);

const EmptyState: React.FC = () => (
  <div style={{
    height: '100%', display: 'flex', flexDirection: 'column',
    alignItems: 'center', justifyContent: 'center', gap: 20, padding: 40
  }}>
    <div style={{ position: 'relative' }}>
      <div style={{
        width: 80, height: 80, borderRadius: 20,
        background: 'linear-gradient(135deg, #0f1520, #161b27)',
        border: '1px solid #1e2e4a',
        display: 'flex', alignItems: 'center', justifyContent: 'center'
      }}>
        <FolderOpen size={32} color="#2a3a5f" />
      </div>
      <div style={{
        position: 'absolute', bottom: -6, right: -6,
        width: 24, height: 24, borderRadius: 8,
        background: 'linear-gradient(135deg, #3b82f6, #6366f1)',
        display: 'flex', alignItems: 'center', justifyContent: 'center'
      }}>
        <Zap size={12} color="#fff" />
      </div>
    </div>
    <div style={{ textAlign: 'center' }}>
      <div style={{ fontSize: 16, fontWeight: 700, color: '#334155', letterSpacing: '0.06em', marginBottom: 6 }}>
        OPEN A FOLDER TO BEGIN
      </div>
      <div style={{ fontSize: 11, color: '#1e2e4a', maxWidth: 320, lineHeight: 1.7 }}>
        Supports <span style={{ color: '#3b82f6' }}>XLSX · XLS · CSV · TXT</span> — all processing
        happens locally in parallel across {MAX_WORKERS} web workers
      </div>
    </div>
    <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', justifyContent: 'center' }}>
      {['⚡ Sub-second search', '🔒 100% local', `🧵 ${navigator.hardwareConcurrency || 4} cores`, '📊 Export CSV'].map(f => (
        <span key={f} style={{
          padding: '4px 10px', borderRadius: 20,
          background: '#0f1520', border: '1px solid #1a2540',
          fontSize: 10, color: '#475569'
        }}>{f}</span>
      ))}
    </div>
  </div>
);

// ── Highlight query terms in text ─────────────────────────────────────────────
const Highlight: React.FC<{ text: string; query: string }> = ({ text, query }) => {
  if (!query.trim()) return <>{text}</>;
  const terms = query.trim().split(/\s+/).filter(t => t.length > 1);
  if (!terms.length) return <>{text}</>;
  const regex = new RegExp(`(${terms.map(t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})`, 'gi');
  const parts = text.split(regex);
  return <>{parts.map((p, i) => regex.test(p) ? <mark key={i} style={{ background: '#1e3a5f', color: '#60a5fa', borderRadius: 2, padding: '0 1px' }}>{p}</mark> : p)}</>;
};

// ── Result row ────────────────────────────────────────────────────────────────
interface ResultRowProps {
  res: SearchResult;
  isExpanded: boolean;
  expandedData: Record<string, unknown> | null;
  loadingDetail: boolean;
  onExpand: () => void;
  query: string;
}

const ResultRow: React.FC<ResultRowProps> = ({ res, isExpanded, expandedData, loadingDetail, onExpand, query }) => (
  <div style={{
    height: isExpanded ? 'auto' : ROW_HEIGHT,
    minHeight: ROW_HEIGHT,
    borderBottom: '1px solid #111520',
    background: isExpanded ? '#0f1520' : 'transparent',
    transition: 'background 0.1s'
  }}>
    {/* Summary row */}
    <div
      onClick={onExpand}
      style={{
        height: ROW_HEIGHT, display: 'flex', alignItems: 'center',
        padding: '0 16px', cursor: 'pointer', gap: 12,
        transition: 'background 0.1s'
      }}
      onMouseEnter={e => (e.currentTarget.style.background = '#0f1520')}
      onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}
    >
      <div style={{
        width: 32, height: 32, borderRadius: 8, flexShrink: 0,
        background: '#111825', border: '1px solid #1e2e44',
        display: 'flex', alignItems: 'center', justifyContent: 'center'
      }}>
        {extIcon(res.fileName)}
      </div>

      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, marginBottom: 2 }}>
          <span style={{ fontSize: 11, fontWeight: 700, color: '#94a3b8', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 200 }}>
            {res.fileName}
          </span>
          {res.sheetName && (
            <span style={{ fontSize: 9, color: '#3b82f6', background: '#0f1e38', border: '1px solid #1e3a5f', borderRadius: 4, padding: '0 4px' }}>
              {res.sheetName}
            </span>
          )}
          <span style={{ fontSize: 9, color: '#334155', flexShrink: 0 }}>ROW {res.rowNumber}</span>
        </div>
        <div style={{ fontSize: 10, color: '#475569', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          <Highlight text={res.searchString} query={query} />
        </div>
      </div>

      {res.score !== undefined && res.score > 0 && (
        <div style={{
          flexShrink: 0, padding: '2px 6px', borderRadius: 6,
          background: '#0f1e38', border: '1px solid #1e3a5f',
          fontSize: 9, color: '#3b82f6', fontWeight: 700
        }}>
          {res.score}
        </div>
      )}

      <ChevronDown size={13} color="#334155"
        style={{ flexShrink: 0, transform: isExpanded ? 'rotate(180deg)' : 'none', transition: 'transform 0.2s' }}
      />
    </div>

    {/* Expanded detail */}
    {isExpanded && (
      <div style={{ padding: '0 16px 14px', borderTop: '1px solid #111520' }}>
        {loadingDetail ? (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '12px 0', color: '#475569', fontSize: 11 }}>
            <Loader2 size={13} style={{ animation: 'spin 1s linear infinite' }} /> Loading row data…
          </div>
        ) : expandedData ? (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: 6, paddingTop: 10 }}>
            {Object.entries(expandedData).map(([k, v]) => (
              <div key={k} style={{
                padding: '6px 8px', borderRadius: 6,
                background: '#111825', border: '1px solid #1e2433'
              }}>
                <div style={{ fontSize: 8, color: '#334155', fontWeight: 700, letterSpacing: '0.12em', marginBottom: 2 }}>{k.toUpperCase()}</div>
                <div style={{ fontSize: 11, color: '#94a3b8', wordBreak: 'break-word' }}>
                  <Highlight text={String(v ?? '—')} query={query} />
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '12px 0', color: '#334155', fontSize: 11 }}>
            <AlertCircle size={13} /> No data available
          </div>
        )}
      </div>
    )}
  </div>
);

export default App;

// CSS for spinner (injected once)
const style = document.createElement('style');
style.textContent = `@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`;
document.head.appendChild(style);
