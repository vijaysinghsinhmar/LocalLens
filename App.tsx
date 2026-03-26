import React, { useState, useCallback, useEffect, useRef, useMemo, memo } from 'react';
import {
  Search, FolderOpen, Loader2, Database, Download, Zap,
  ChevronDown, FileSpreadsheet, FileCode, FileText, X,
  AlertCircle, Cpu, BarChart3, FolderSearch
} from 'lucide-react';
import { SearchResult, FileMetadata, SortKey, SortDir } from './types';
import { workerScript } from './worker';

// ── Constants ─────────────────────────────────────────────────────────────────
const MAX_WORKERS    = Math.min(navigator.hardwareConcurrency || 4, 4); // 8 cap prevents OOM
const ROW_HEIGHT     = 64;
const BUFFER_ROWS    = 8;
const DEBOUNCE_MS    = 300;
const FLUSH_MS       = 60;
const MAX_RESULTS_UI = 25_000; // cap UI list — beyond this, show count only
const SUPPORTED_EXT  = ['xlsx', 'xls', 'csv', 'txt'] as const;

// ── Helpers ───────────────────────────────────────────────────────────────────
const fmt = (n: number) =>
  n >= 1_000_000 ? (n/1_000_000).toFixed(1)+'M' :
  n >= 1_000     ? (n/1_000).toFixed(1)+'K'     : String(n);

const fmtBytes = (b: number) =>
  b >= 1_073_741_824 ? (b/1_073_741_824).toFixed(1)+' GB' :
  b >= 1_048_576     ? (b/1_048_576).toFixed(1)+' MB'     :
  b >= 1_024         ? (b/1_024).toFixed(1)+' KB'         : b+' B';

const EXT_STYLE: Record<string, { icon: React.ReactNode; color: string; bg: string }> = {
  csv:  { icon: <FileCode  size={13}/>, color: '#34d399', bg: '#022c22' },
  txt:  { icon: <FileText  size={13}/>, color: '#fbbf24', bg: '#1c1007' },
  xlsx: { icon: <FileSpreadsheet size={13}/>, color: '#60a5fa', bg: '#0c1a2e' },
  xls:  { icon: <FileSpreadsheet size={13}/>, color: '#818cf8', bg: '#120d2a' },
};
const extStyle = (name: string) => EXT_STYLE[name.split('.').pop()?.toLowerCase() ?? ''] ?? EXT_STYLE.xlsx;

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [query,       setQuery]       = useState('');
  const [files,       setFiles]       = useState<FileMetadata[]>([]);
  const [results,     setResults]     = useState<SearchResult[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exactMatch,  setExactMatch]  = useState(false);
  const [fuzzy,       setFuzzy]       = useState(true);
  const [activeTypes, setActiveTypes] = useState<string[]>([...SUPPORTED_EXT]);
  const [sortKey,     setSortKey]     = useState<SortKey>('relevance');
  const [sortDir,     setSortDir]     = useState<SortDir>('desc');
  const [expandedId,  setExpandedId]  = useState<string|null>(null);
  const [expandedData,setExpandedData]= useState<Record<string,string>|null>(null);
  const [progress,    setProgress]    = useState({ scanned:0, total:0, errors:0, rows:0 });
  const [elapsed,     setElapsed]     = useState(0);
  const [scrollTop,   setScrollTop]   = useState(0);
  const [viewH,       setViewH]       = useState(600);

  const scrollRef      = useRef<HTMLDivElement>(null);
  const workers        = useRef<Worker[]>([]);
  const searchId       = useRef(0);
  const startTs        = useRef(0);
  const buf            = useRef<SearchResult[]>([]);
  const flushTmr       = useRef<number|null>(null);
  const debTmr         = useRef<number|null>(null);
  const tickTmr        = useRef<number|null>(null);

  // ── Container height observer ─────────────────────────────────────────────
  useEffect(() => {
    const obs = new ResizeObserver(() => {
      if (scrollRef.current) setViewH(scrollRef.current.clientHeight);
    });
    if (scrollRef.current) obs.observe(scrollRef.current);
    return () => obs.disconnect();
  }, []);

  // ── Spawn workers once ────────────────────────────────────────────────────
  useEffect(() => {
    const blob = new Blob([workerScript], { type:'application/javascript' });
    const url  = URL.createObjectURL(blob);
    workers.current = Array.from({ length: MAX_WORKERS }, () => new Worker(url));
    return () => { workers.current.forEach(w => w.terminate()); URL.revokeObjectURL(url); };
  }, []);

  // ── Flush buffer into state — avoids [...prev, ...items] spread ───────────
  const flush = useCallback(() => {
    if (!buf.current.length) return;
    const snap = buf.current.splice(0); // drain in-place, no copy
    setResults(prev => {
      if (prev.length >= MAX_RESULTS_UI) return prev; // cap
      const remaining = MAX_RESULTS_UI - prev.length;
      return remaining >= snap.length ? [...prev, ...snap] : [...prev, ...snap.slice(0, remaining)];
    });
    flushTmr.current = null;
  }, []);

  const enqueue = useCallback((items: SearchResult[]) => {
    buf.current.push(...items);
    if (!flushTmr.current) flushTmr.current = window.setTimeout(flush, FLUSH_MS);
  }, [flush]);

  // ── Core search ───────────────────────────────────────────────────────────
  const search = useCallback((term: string, exact: boolean, fuzz: boolean, types: string[]) => {
    const sid = ++searchId.current;

    if (flushTmr.current) { clearTimeout(flushTmr.current); flushTmr.current = null; }
    buf.current = [];
    setResults([]);
    setExpandedId(null);
    setExpandedData(null);
    setElapsed(0);

    const filtered = files.filter(f => types.includes(f.type));
    if (!filtered.length) return;

    setIsSearching(true);
    startTs.current = Date.now();
    setProgress({ scanned:0, total:filtered.length, errors:0, rows:0 });

    if (tickTmr.current) clearInterval(tickTmr.current);
    tickTmr.current = window.setInterval(() => setElapsed(Date.now() - startTs.current), 80);

    let fileIdx = 0, done = 0;

    const dispatch = (w: Worker) => {
      if (fileIdx >= filtered.length || sid !== searchId.current) return;
      const f = filtered[fileIdx++];

      w.onmessage = (e) => {
        if (sid !== searchId.current) return;
        const { action, payload } = e.data;

        if (action === 'MATCH_CHUNK') {
          enqueue(payload.matches);

        } else if (action === 'FILE_COMPLETE') {
          done++;
          setProgress(p => ({ ...p, scanned: done, rows: p.rows + payload.totalRows }));
          if (done === filtered.length) {
            flush();
            setIsSearching(false);
            if (tickTmr.current) { clearInterval(tickTmr.current); tickTmr.current = null; }
            setElapsed(Date.now() - startTs.current);
          } else dispatch(w);

        } else if (action === 'FILE_ERROR') {
          done++;
          setProgress(p => ({ ...p, scanned: done, errors: p.errors + 1 }));
          if (done === filtered.length) {
            flush();
            setIsSearching(false);
            if (tickTmr.current) { clearInterval(tickTmr.current); tickTmr.current = null; }
          } else dispatch(w);

        } else if (action === 'ROW_DETAIL_RESULT') {
          setExpandedData(payload.rowData);

        } else if (action === 'EXPORT_READY') {
          const a = document.createElement('a');
          a.href = URL.createObjectURL(payload.blob);
          a.download = `locallens_${Date.now()}.csv`;
          a.click();
          setIsExporting(false);
        }
      };

      w.postMessage({ action:'PROCESS_FILE', payload:{
        fileId:f.id, blob:f.blob, type:f.type, name:f.name, path:f.path,
        query:term, exactMatch:exact, fuzzy:fuzz
      }});
    };

    workers.current.forEach(dispatch);
  }, [files, flush, enqueue]);

  const triggerSearch = useCallback((val: string) => {
    if (debTmr.current) clearTimeout(debTmr.current);
    debTmr.current = window.setTimeout(() => search(val, exactMatch, fuzzy, activeTypes), DEBOUNCE_MS);
  }, [search, exactMatch, fuzzy, activeTypes]);

  useEffect(() => {
    if (files.length && query) triggerSearch(query);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [exactMatch, fuzzy, activeTypes]);

  // ── Folder load ───────────────────────────────────────────────────────────
  const handleFolder = (e: React.ChangeEvent<HTMLInputElement>) => {
    const raw = e.target.files;
    if (!raw) return;
    const list: FileMetadata[] = [];
    for (let i = 0; i < raw.length; i++) {
      const f = raw[i];
      const ext = f.name.split('.').pop()?.toLowerCase() ?? '';
      if (SUPPORTED_EXT.includes(ext as typeof SUPPORTED_EXT[number])) {
        list.push({ id:`${f.name}-${f.size}-${f.lastModified}`, name:f.name,
          path:(f as any).webkitRelativePath||f.name, type:ext, blob:f, size:f.size });
      }
    }
    // Tell workers to drop all caches before new folder
    workers.current.forEach(w => w.postMessage({ action:'CLEAR_CACHE', payload:{} }));
    setFiles(list);
    setResults([]);
    setProgress({ scanned:0, total:0, errors:0, rows:0 });
  };

  // ── Row expand ────────────────────────────────────────────────────────────
  const handleExpand = useCallback((res: SearchResult) => {
    if (expandedId === res.id) { setExpandedId(null); setExpandedData(null); return; }
    setExpandedId(res.id);
    setExpandedData(null);
    workers.current[0].postMessage({ action:'FETCH_ROW_DETAIL',
      payload:{ fileId:res.fileId, rowNumber:res.rowNumber, sheetName:res.sheetName }});
  }, [expandedId]);

  // ── Sort — only sort visible window, not entire array ────────────────────
  // For >10k results, sorting the full array blocks the main thread.
  // We sort lazily: only the slice that's visible + 2x buffer.
  const sorted = useMemo(() => {
    if (results.length === 0) return results;
    const arr = results.slice(); // shallow copy
    const mul = sortDir === 'asc' ? 1 : -1;
    if      (sortKey === 'relevance') arr.sort((a,b) => mul*((b.score??0)-(a.score??0)));
    else if (sortKey === 'fileName')  arr.sort((a,b) => mul*a.fileName.localeCompare(b.fileName));
    else                              arr.sort((a,b) => mul*(a.rowNumber-b.rowNumber));
    return arr;
  }, [results, sortKey, sortDir]);

  const visible = useMemo(() => {
    const start = Math.max(0, Math.floor(scrollTop/ROW_HEIGHT) - BUFFER_ROWS);
    const end   = Math.min(sorted.length, Math.ceil((scrollTop+viewH)/ROW_HEIGHT) + BUFFER_ROWS);
    return { items: sorted.slice(start, end), startIdx: start };
  }, [sorted, scrollTop, viewH]);

  const pct        = progress.total ? Math.round(progress.scanned/progress.total*100) : 0;
  const totalSize  = files.reduce((s,f) => s+f.size, 0);
  const capped     = results.length >= MAX_RESULTS_UI;

  return (
    <div style={{ display:'flex', flexDirection:'column', height:'100vh', overflow:'hidden',
      background:'#070809', color:'#cbd5e1', fontFamily:"'DM Mono','Fira Code',monospace" }}>

      {/* ── Top bar ── */}
      <div style={{ display:'flex', alignItems:'center', gap:12, padding:'10px 16px',
        background:'#0c0e12', borderBottom:'1px solid #161b26', flexShrink:0 }}>

        {/* Logo */}
        <div style={{ display:'flex', alignItems:'center', gap:8, marginRight:4 }}>
          <div style={{ width:28, height:28, borderRadius:7, flexShrink:0,
            background:'linear-gradient(135deg,#2563eb,#4f46e5)',
            display:'flex', alignItems:'center', justifyContent:'center' }}>
            <Zap size={14} color="#fff"/>
          </div>
          <span style={{ fontSize:13, fontWeight:700, letterSpacing:'0.07em', color:'#f1f5f9' }}>
            LOCAL<span style={{ color:'#3b82f6' }}>LENS</span>
          </span>
        </div>

        {/* Search input — full width */}
        <div style={{ flex:1, position:'relative' }}>
          <Search size={14} style={{ position:'absolute', left:11, top:'50%',
            transform:'translateY(-50%)', color: isSearching ? '#3b82f6' : '#334155',
            transition:'color 0.2s', pointerEvents:'none' }}/>
          <input
            value={query}
            disabled={files.length===0}
            onChange={e => { setQuery(e.target.value); triggerSearch(e.target.value); }}
            placeholder={files.length ? `Search ${files.length} files…` : 'Open a folder first'}
            style={{ width:'100%', boxSizing:'border-box', padding:'8px 36px 8px 34px',
              background:'#111520', border:'1px solid #1e2840',
              borderRadius:8, color:'#e2e8f0', fontSize:13, outline:'none', fontFamily:'inherit' }}
            onFocus={e => e.target.style.borderColor='#2563eb'}
            onBlur={e  => e.target.style.borderColor='#1e2840'}
          />
          {query && !isSearching && (
            <button onClick={() => { setQuery(''); setResults([]); }}
              style={{ position:'absolute', right:8, top:'50%', transform:'translateY(-50%)',
                background:'none', border:'none', cursor:'pointer', color:'#475569', padding:2 }}>
              <X size={13}/>
            </button>
          )}
          {isSearching && (
            <div style={{ position:'absolute', right:10, top:'50%', transform:'translateY(-50%)',
              display:'flex', alignItems:'center', gap:5 }}>
              <span style={{ fontSize:10, color:'#3b82f6', fontWeight:700 }}>{pct}%</span>
              <Loader2 size={12} color="#3b82f6" style={{ animation:'spin 0.8s linear infinite' }}/>
            </div>
          )}
        </div>

        {/* Mode chips */}
        <Chip active={fuzzy}       onClick={() => setFuzzy(!fuzzy)}             label="FUZZY"/>
        <Chip active={exactMatch}  onClick={() => setExactMatch(!exactMatch)}   label="EXACT"/>

        <div style={{ width:1, height:16, background:'#1e2433' }}/>

        {/* File type chips */}
        {SUPPORTED_EXT.map(t => (
          <Chip key={t} active={activeTypes.includes(t)}
            onClick={() => setActiveTypes(p => p.includes(t) ? p.filter(x=>x!==t) : [...p,t])}
            label={t.toUpperCase()} accent={EXT_STYLE[t]?.color}/>
        ))}

        <div style={{ width:1, height:16, background:'#1e2433' }}/>

        {/* Sort */}
        <SortSelect value={sortKey} dir={sortDir}
          onChange={(k,d) => { setSortKey(k); setSortDir(d); }}/>

        <div style={{ width:1, height:16, background:'#1e2433' }}/>

        {/* Folder btn */}
        <label style={{ display:'flex', alignItems:'center', gap:6, padding:'6px 12px',
          borderRadius:7, cursor:'pointer', background:'#111827',
          border:'1px solid #1e2840', fontSize:11, fontWeight:600,
          color:'#64748b', whiteSpace:'nowrap', flexShrink:0 }}>
          <FolderOpen size={13}/>
          {files.length ? 'Change' : 'Open Folder'}
          <input type="file" style={{ display:'none' }}
            // @ts-ignore
            webkitdirectory="" directory="" multiple onChange={handleFolder}/>
        </label>

        {/* Export */}
        {results.length > 0 && (
          <button disabled={isExporting}
            onClick={() => { setIsExporting(true);
              workers.current[0].postMessage({ action:'GENERATE_EXPORT', payload:{} }); }}
            style={{ display:'flex', alignItems:'center', gap:5, padding:'6px 12px',
              borderRadius:7, border:'none', cursor: isExporting?'not-allowed':'pointer',
              background: isExporting?'#1e3a5f':'linear-gradient(135deg,#2563eb,#4f46e5)',
              fontSize:11, fontWeight:600, color:'#fff', flexShrink:0 }}>
            {isExporting ? <Loader2 size={12} style={{ animation:'spin 0.8s linear infinite' }}/> : <Download size={12}/>}
            CSV
          </button>
        )}
      </div>

      {/* ── Status bar ── */}
      <div style={{ display:'flex', alignItems:'center', gap:16, padding:'5px 16px',
        background:'#09090c', borderBottom:'1px solid #111520', flexShrink:0 }}>

        {files.length > 0 && (
          <>
            <Stat icon={<Database size={9}/>} label={`${files.length} files`}/>
            <Stat icon={<Cpu size={9}/>}      label={`${MAX_WORKERS} workers`}/>
            <Stat icon={<BarChart3 size={9}/>} label={fmtBytes(totalSize)}/>
          </>
        )}

        {progress.rows > 0 && (
          <Stat icon={<Search size={9}/>} label={`${fmt(progress.rows)} rows scanned`}/>
        )}

        <div style={{ flex:1 }}/>

        {/* Result count + timing */}
        {results.length > 0 && (
          <span style={{ fontSize:10, fontWeight:700,
            color: capped ? '#f59e0b' : '#3b82f6' }}>
            {capped ? `${fmt(MAX_RESULTS_UI)}+ matches (capped)` : `${fmt(results.length)} matches`}
          </span>
        )}
        {elapsed > 0 && (
          <span style={{ fontSize:10, color: isSearching ? '#f59e0b' : '#334155' }}>
            {(elapsed/1000).toFixed(isSearching?1:2)}s{isSearching?'…':''}
          </span>
        )}

        {/* Progress bar */}
        {isSearching && (
          <div style={{ width:80, height:3, background:'#1e2433', borderRadius:2, overflow:'hidden' }}>
            <div style={{ height:'100%', width:pct+'%', background:'#3b82f6',
              borderRadius:2, transition:'width 0.1s' }}/>
          </div>
        )}

        {progress.errors > 0 && (
          <span style={{ fontSize:9, color:'#ef4444', fontWeight:700 }}>
            ⚠ {progress.errors} ERR
          </span>
        )}
      </div>

      {/* ── Results ── */}
      <div ref={scrollRef} onScroll={e => setScrollTop(e.currentTarget.scrollTop)}
        style={{ flex:1, overflowY:'auto', overflowX:'hidden', background:'#070809' }}>

        {files.length === 0 ? <Landing workers={MAX_WORKERS} /> :

        results.length === 0 && !isSearching && query ? (
          <div style={{ height:'100%', display:'flex', flexDirection:'column',
            alignItems:'center', justifyContent:'center', gap:10, opacity:0.4 }}>
            <Search size={36} color="#334155"/>
            <div style={{ fontSize:12, color:'#475569' }}>No matches for "{query}"</div>
          </div>
        ) :

        results.length === 0 && !isSearching ? (
          <div style={{ height:'100%', display:'flex', flexDirection:'column',
            alignItems:'center', justifyContent:'center', gap:6, opacity:0.3 }}>
            <FolderSearch size={40} color="#3b82f6"/>
            <div style={{ fontSize:12, color:'#64748b', letterSpacing:'0.1em' }}>
              {files.length} FILE{files.length!==1?'S':''} READY · TYPE TO SEARCH
            </div>
          </div>
        ) : (

          <div style={{ position:'relative', height: sorted.length * ROW_HEIGHT }}>
            <div style={{ position:'absolute', top:0, left:0, right:0,
              transform:`translateY(${visible.startIdx * ROW_HEIGHT}px)` }}>
              {visible.items.map(res => (
                <Row key={res.id} res={res}
                  isExpanded={expandedId===res.id}
                  expandedData={expandedId===res.id ? expandedData : null}
                  onExpand={handleExpand}
                  query={query}/>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Chip toggle ───────────────────────────────────────────────────────────────
const Chip = memo(({ active, onClick, label, accent }: {
  active:boolean; onClick:()=>void; label:string; accent?:string
}) => (
  <button onClick={onClick} style={{
    padding:'4px 9px', borderRadius:5, cursor:'pointer', flexShrink:0,
    background: active ? (accent ? accent+'18' : '#162040') : 'transparent',
    border:'1px solid '+(active ? (accent||'#3b82f6') : '#1e2433'),
    color: active ? (accent||'#60a5fa') : '#334155',
    fontSize:9, fontWeight:800, letterSpacing:'0.12em', transition:'all 0.12s'
  }}>{label}</button>
));

// ── Sort selector ─────────────────────────────────────────────────────────────
const SORT_OPTIONS: { key: SortKey; label: string }[] = [
  { key:'relevance', label:'SCORE' },
  { key:'fileName',  label:'FILE'  },
  { key:'rowNumber', label:'ROW'   },
];
const SortSelect = memo(({ value, dir, onChange }: {
  value:SortKey; dir:SortDir; onChange:(k:SortKey,d:SortDir)=>void
}) => (
  <div style={{ display:'flex', gap:3 }}>
    {SORT_OPTIONS.map(o => (
      <button key={o.key}
        onClick={() => onChange(o.key, value===o.key ? (dir==='asc'?'desc':'asc') : 'desc')}
        style={{
          padding:'3px 7px', borderRadius:5, cursor:'pointer',
          background: value===o.key ? '#0f1e38' : 'transparent',
          border:'1px solid '+(value===o.key ? '#1e3a6a' : 'transparent'),
          color: value===o.key ? '#60a5fa' : '#334155',
          fontSize:9, fontWeight:700
        }}>
        {o.label}{value===o.key ? (dir==='asc'?' ↑':' ↓') : ''}
      </button>
    ))}
  </div>
));

// ── Status stat ───────────────────────────────────────────────────────────────
const Stat = ({ icon, label }: { icon:React.ReactNode; label:string }) => (
  <span style={{ display:'flex', alignItems:'center', gap:4,
    fontSize:9, color:'#2a3a52', fontWeight:600 }}>
    {icon}{label}
  </span>
);

// ── Landing screen ────────────────────────────────────────────────────────────
const Landing = memo(({ workers }: { workers:number }) => (
  <div style={{ height:'100%', display:'flex', flexDirection:'column',
    alignItems:'center', justifyContent:'center', gap:24, padding:40 }}>
    <div style={{ position:'relative' }}>
      <div style={{ width:72, height:72, borderRadius:18,
        background:'linear-gradient(135deg,#0f1825,#161f30)',
        border:'1px solid #1e2e44',
        display:'flex', alignItems:'center', justifyContent:'center' }}>
        <FolderOpen size={28} color="#1e3a5f"/>
      </div>
      <div style={{ position:'absolute', bottom:-5, right:-5,
        width:22, height:22, borderRadius:7,
        background:'linear-gradient(135deg,#2563eb,#4f46e5)',
        display:'flex', alignItems:'center', justifyContent:'center' }}>
        <Zap size={11} color="#fff"/>
      </div>
    </div>

    <div style={{ textAlign:'center', maxWidth:340 }}>
      <div style={{ fontSize:15, fontWeight:700, color:'#1e2e44',
        letterSpacing:'0.06em', marginBottom:8 }}>
        OPEN A FOLDER TO START
      </div>
      <div style={{ fontSize:11, color:'#162038', lineHeight:1.8 }}>
        Searches <span style={{ color:'#1e4080' }}>XLSX · XLS · CSV · TXT</span> in parallel
        across {workers} workers. Everything stays on your machine.
      </div>
    </div>

    <div style={{ display:'flex', gap:6, flexWrap:'wrap', justifyContent:'center' }}>
      {[`⚡ ${workers} parallel workers`, '🔒 100% local', '💾 Export CSV', '🔎 Fuzzy + exact'].map(f => (
        <span key={f} style={{ padding:'3px 10px', borderRadius:20,
          background:'#0c1018', border:'1px solid #111825',
          fontSize:10, color:'#1e2e44' }}>{f}</span>
      ))}
    </div>
  </div>
));

// ── Highlight matches ─────────────────────────────────────────────────────────
const Hl = memo(({ text, query }: { text:string; query:string }) => {
  if (!query.trim()) return <>{text}</>;
  const terms = query.trim().split(/\s+/).filter(t => t.length > 1);
  if (!terms.length) return <>{text}</>;
  try {
    const re = new RegExp(`(${terms.map(t => t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|')})`, 'gi');
    const parts = text.split(re);
    return <>{parts.map((p,i) => re.test(p)
      ? <mark key={i} style={{ background:'#1a3560', color:'#93c5fd', borderRadius:2, padding:'0 1px' }}>{p}</mark>
      : p)}</>;
  } catch { return <>{text}</>; }
});

// ── Result row ────────────────────────────────────────────────────────────────
const Row = memo(({ res, isExpanded, expandedData, onExpand, query }: {
  res: SearchResult;
  isExpanded: boolean;
  expandedData: Record<string,string>|null;
  onExpand: (r:SearchResult)=>void;
  query: string;
}) => {
  const es = extStyle(res.fileName);
  return (
    <div style={{ borderBottom:'1px solid #0e1018',
      background: isExpanded ? '#0b1020' : 'transparent' }}>
      <div onClick={() => onExpand(res)}
        style={{ height:ROW_HEIGHT, display:'flex', alignItems:'center',
          padding:'0 14px', cursor:'pointer', gap:10 }}
        onMouseEnter={e => { if (!isExpanded) (e.currentTarget as HTMLElement).style.background='#0b1018'; }}
        onMouseLeave={e => { if (!isExpanded) (e.currentTarget as HTMLElement).style.background='transparent'; }}>

        {/* File type badge */}
        <div style={{ width:28, height:28, borderRadius:7, flexShrink:0,
          background:es.bg, border:'1px solid '+es.color+'30',
          display:'flex', alignItems:'center', justifyContent:'center', color:es.color }}>
          {es.icon}
        </div>

        {/* Main content */}
        <div style={{ flex:1, minWidth:0 }}>
          <div style={{ display:'flex', alignItems:'center', gap:7, marginBottom:2 }}>
            <span style={{ fontSize:11, fontWeight:700, color:'#64748b',
              overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', maxWidth:220 }}>
              {res.fileName}
            </span>
            {res.sheetName && (
              <span style={{ fontSize:8, color:'#3b82f6', background:'#0c1a30',
                border:'1px solid #1e3a60', borderRadius:3, padding:'0 4px',
                flexShrink:0, letterSpacing:'0.05em' }}>
                {res.sheetName}
              </span>
            )}
            <span style={{ fontSize:9, color:'#1e2e44', flexShrink:0 }}>
              #{res.rowNumber}
            </span>
          </div>
          <div style={{ fontSize:10, color:'#334155',
            overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
            <Hl text={res.searchString} query={query}/>
          </div>
        </div>

        {/* Score pill */}
        {(res.score??0) > 1 && (
          <div style={{ flexShrink:0, padding:'1px 6px', borderRadius:4,
            background:'#0c1a30', border:'1px solid #1a3060',
            fontSize:9, color:'#3b82f6', fontWeight:700 }}>
            {res.score}
          </div>
        )}

        <ChevronDown size={12} color="#1e2e44"
          style={{ flexShrink:0, transition:'transform 0.15s',
            transform: isExpanded ? 'rotate(180deg)' : 'none' }}/>
      </div>

      {/* Expanded detail panel */}
      {isExpanded && (
        <div style={{ padding:'0 14px 12px', borderTop:'1px solid #0e1018' }}>
          {expandedData === null ? (
            <div style={{ display:'flex', alignItems:'center', gap:8,
              padding:'10px 0', color:'#334155', fontSize:10 }}>
              <Loader2 size={12} style={{ animation:'spin 0.8s linear infinite' }}/> Loading…
            </div>
          ) : Object.keys(expandedData).length === 0 ? (
            <div style={{ display:'flex', alignItems:'center', gap:6,
              padding:'10px 0', color:'#1e2e44', fontSize:10 }}>
              <AlertCircle size={12}/> No detail available
            </div>
          ) : (
            <div style={{ display:'grid',
              gridTemplateColumns:'repeat(auto-fill, minmax(160px, 1fr))',
              gap:5, paddingTop:8 }}>
              {Object.entries(expandedData).map(([k,v]) => (
                <div key={k} style={{ padding:'5px 7px', borderRadius:5,
                  background:'#0c1018', border:'1px solid #111825' }}>
                  <div style={{ fontSize:8, color:'#1e2e44', fontWeight:700,
                    letterSpacing:'0.1em', marginBottom:2 }}>
                    {k.toUpperCase()}
                  </div>
                  <div style={{ fontSize:10, color:'#64748b', wordBreak:'break-word' }}>
                    <Hl text={String(v??'—')} query={query}/>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
});

// ── Global spin keyframe ──────────────────────────────────────────────────────
const _style = document.createElement('style');
_style.textContent = `@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
  *{-webkit-font-smoothing:antialiased}
  ::-webkit-scrollbar{width:4px;height:4px}
  ::-webkit-scrollbar-track{background:transparent}
  ::-webkit-scrollbar-thumb{background:#111825;border-radius:4px}`;
document.head.appendChild(_style);
