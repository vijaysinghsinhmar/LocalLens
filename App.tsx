import React, { useState, useCallback, useEffect, useRef, useMemo, memo } from 'react';
import {
  Search, FolderOpen, Loader2, Download, Zap, ChevronDown,
  FileSpreadsheet, FileCode, FileText, X, AlertCircle,
  CheckCircle2, Clock, Database
} from 'lucide-react';
import { SearchResult, FileEntry, SortKey, SortDir } from './types';
import { workerScript } from './worker';

// ── Config ────────────────────────────────────────────────────────────────────
const NUM_WORKERS    = 2;
const ROW_H          = 60;
const OVER           = 10;
const DEBOUNCE       = 250;
const FLUSH_INTERVAL = 60;
const MAX_RENDER     = 20_000;
const SUPPORTED      = new Set(['xlsx','xls','csv','txt']);

const fmt = (n: number) =>
  n >= 1e6 ? (n/1e6).toFixed(1)+'M' : n >= 1e3 ? (n/1e3).toFixed(1)+'K' : ''+n;
const fmtBytes = (b: number) =>
  b >= 1<<30 ? (b/(1<<30)).toFixed(1)+' GB' :
  b >= 1<<20 ? (b/(1<<20)).toFixed(1)+' MB' :
  b >= 1<<10 ? (b/(1<<10)).toFixed(1)+' KB' : b+' B';
const fmtTime = (ms: number) =>
  ms >= 60000 ? (ms/60000).toFixed(1)+'m' :
  ms >= 1000  ? (ms/1000).toFixed(2)+'s'  : ms+'ms';

type ExtKey = 'xlsx'|'xls'|'csv'|'txt';
const EXT: Record<ExtKey,{color:string;dim:string;Icon:any}> = {
  xlsx:{color:'#60a5fa',dim:'#1e3a5f',Icon:FileSpreadsheet},
  xls: {color:'#818cf8',dim:'#251b50',Icon:FileSpreadsheet},
  csv: {color:'#34d399',dim:'#032917',Icon:FileCode},
  txt: {color:'#fbbf24',dim:'#1c1205',Icon:FileText},
};
const extOf = (name:string):ExtKey => (name.split('.').pop()?.toLowerCase()||'') as ExtKey;

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [files,      setFiles]      = useState<FileEntry[]>([]);
  const [totalRows,  setTotalRows]  = useState(0);
  const [indexing,   setIndexing]   = useState(false);

  const [query,      setQuery]      = useState('');
  const [results,    setResults]    = useState<SearchResult[]>([]);
  const [searching,  setSearching]  = useState(false);
  const [elapsed,    setElapsed]    = useState<number|null>(null);

  const [fuzzy,      setFuzzy]      = useState(true);
  const [exact,      setExact]      = useState(false);
  const [types,      setTypes]      = useState<string[]>(['xlsx','xls','csv','txt']);
  const [sortKey,    setSortKey]    = useState<SortKey>('relevance');
  const [sortDir,    setSortDir]    = useState<SortDir>('desc');

  const [expandId,   setExpandId]   = useState<string|null>(null);
  const [expandData, setExpandData] = useState<Record<string,string>|null|'loading'>(null);

  const [scrollTop,  setScrollTop]  = useState(0);
  const [viewH,      setViewH]      = useState(600);
  const scrollRef = useRef<HTMLDivElement>(null);

  // ── Stable refs (never cause re-renders) ─────────────────────────────────
  const workers       = useRef<Worker[]>([]);
  const searchId      = useRef(0);
  const startTs       = useRef(0);
  const buf           = useRef<SearchResult[]>([]);
  const flushTmr      = useRef<number|null>(null);
  const debTmr        = useRef<number|null>(null);
  const tickTmr       = useRef<number|null>(null);
  const allResults    = useRef<SearchResult[]>([]);
  const pendingDone   = useRef(0);
  const pendingTotal  = useRef(0);
  // Index queue state — refs so worker handler can read them without stale closures
  const indexQueue    = useRef<FileEntry[]>([]);
  const indexQueuePos = useRef(0);         // next file to dispatch
  const indexDone     = useRef(0);         // files finished indexing
  const indexTotal    = useRef(0);
  // Live files ref so worker handler always sees current files without stale closure
  const filesRef      = useRef<FileEntry[]>([]);

  // ── Resize observer ──────────────────────────────────────────────────────
  useEffect(() => {
    const obs = new ResizeObserver(() => {
      if (scrollRef.current) setViewH(scrollRef.current.clientHeight);
    });
    if (scrollRef.current) obs.observe(scrollRef.current);
    return () => obs.disconnect();
  }, []);

  // ── Single stable worker message router ──────────────────────────────────
  // All logic uses refs — no stale closure risk. Set once on mount.
  useEffect(() => {
    const blob = new Blob([workerScript], {type:'application/javascript'});
    const url  = URL.createObjectURL(blob);

    const dispatch = (workerIdx: number) => {
      const pos = indexQueuePos.current;
      if (pos >= indexQueue.current.length) return;
      indexQueuePos.current++;
      const f = indexQueue.current[pos];
      workers.current[workerIdx].postMessage({
        action: 'INDEX_FILE',
        payload: { fileId:f.id, blob:f.blob, type:f.type, name:f.name }
      });
    };

    const handleMsg = (workerIdx: number) => (e: MessageEvent) => {
      const { action, payload } = e.data;

      // ── Indexing responses ──────────────────────────────────────────────
      if (action === 'INDEX_DONE') {
        indexDone.current++;
        // Update file state
        setFiles(prev => {
          const next = prev.map(f =>
            f.id === payload.fileId
              ? {...f, indexed:true, rowCount:payload.rowCount}
              : f
          );
          filesRef.current = next;
          return next;
        });
        setTotalRows(prev => prev + payload.rowCount);
        // Dispatch next queued file on this worker
        dispatch(workerIdx);
        // Check all done
        if (indexDone.current >= indexTotal.current) setIndexing(false);

      } else if (action === 'INDEX_ERROR') {
        indexDone.current++;
        setFiles(prev => {
          const next = prev.map(f =>
            f.id === payload.fileId
              ? {...f, indexed:false, error:payload.error}
              : f
          );
          filesRef.current = next;
          return next;
        });
        dispatch(workerIdx);
        if (indexDone.current >= indexTotal.current) setIndexing(false);

      // ── Search responses ────────────────────────────────────────────────
      } else if (action === 'SEARCH_CHUNK') {
        if (payload.searchId !== searchId.current) return;
        buf.current.push(...payload.matches);
        if (!flushTmr.current) {
          flushTmr.current = window.setTimeout(() => {
            const snap = buf.current.splice(0);
            allResults.current.push(...snap);
            setResults(allResults.current.slice(0, MAX_RENDER));
            flushTmr.current = null;
          }, FLUSH_INTERVAL);
        }

      } else if (action === 'SEARCH_DONE') {
        if (payload.searchId !== searchId.current) return;
        pendingDone.current++;
        if (pendingDone.current >= pendingTotal.current) {
          if (flushTmr.current) { clearTimeout(flushTmr.current); flushTmr.current = null; }
          const snap = buf.current.splice(0);
          allResults.current.push(...snap);
          setResults(allResults.current.slice(0, MAX_RENDER));
          setSearching(false);
          if (tickTmr.current) { clearInterval(tickTmr.current); tickTmr.current = null; }
          setElapsed(Date.now() - startTs.current);
        }

      } else if (action === 'SEARCH_ERROR') {
        if (payload.searchId !== searchId.current) return;
        pendingDone.current++;
        if (pendingDone.current >= pendingTotal.current) {
          setSearching(false);
          if (tickTmr.current) { clearInterval(tickTmr.current); tickTmr.current = null; }
          setElapsed(Date.now() - startTs.current);
        }

      // ── Row detail ──────────────────────────────────────────────────────
      } else if (action === 'ROW_DATA') {
        setExpandData(payload.data ?? null);

      // ── Export ──────────────────────────────────────────────────────────
      } else if (action === 'EXPORT_READY') {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(payload.blob);
        a.download = `locallens_${Date.now()}.csv`;
        a.click();
      }
    };

    workers.current = Array.from({length:NUM_WORKERS}, (_, i) => {
      const w = new Worker(url);
      w.onmessage = handleMsg(i);
      return w;
    });

    return () => {
      workers.current.forEach(w => w.terminate());
      URL.revokeObjectURL(url);
    };
  // Intentionally empty deps — runs once on mount, uses only refs
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Index a file list ─────────────────────────────────────────────────────
  const startIndexing = useCallback((fileList: FileEntry[]) => {
    // Reset all caches in workers
    workers.current.forEach(w => w.postMessage({action:'CLEAR_ALL',payload:{}}));

    indexQueue.current    = fileList;
    indexQueuePos.current = 0;
    indexDone.current     = 0;
    indexTotal.current    = fileList.length;

    setIndexing(true);
    setTotalRows(0);

    // Seed each worker with one file
    workers.current.forEach((_, i) => {
      if (i < fileList.length) {
        indexQueuePos.current = i + 1;
        workers.current[i].postMessage({
          action:'INDEX_FILE',
          payload:{
            fileId: fileList[i].id,
            blob:   fileList[i].blob,
            type:   fileList[i].type,
            name:   fileList[i].name
          }
        });
      }
    });
  }, []);

  // ── Folder open ───────────────────────────────────────────────────────────
  const handleFolder = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const raw = e.target.files;
    if (!raw) return;
    const list: FileEntry[] = [];
    for (let i = 0; i < raw.length; i++) {
      const f = raw[i];
      const ext = f.name.split('.').pop()?.toLowerCase() ?? '';
      if (SUPPORTED.has(ext)) {
        list.push({
          id:`${f.name}-${f.size}-${f.lastModified}`,
          name:f.name, path:(f as any).webkitRelativePath||f.name,
          type:ext, blob:f, size:f.size, indexed:false, rowCount:0
        });
      }
    }
    filesRef.current = list;
    setFiles(list);
    setResults([]);
    setQuery('');
    setElapsed(null);
    setTotalRows(0);
    allResults.current = [];
    startIndexing(list);
  }, [startIndexing]);

  // ── Search ────────────────────────────────────────────────────────────────
  const runSearch = useCallback((q: string, fuzz: boolean, ex: boolean, activeTypes: string[]) => {
    if (!q.trim()) return;
    const sid = ++searchId.current;

    buf.current = [];
    allResults.current = [];
    if (flushTmr.current) { clearTimeout(flushTmr.current); flushTmr.current = null; }
    setResults([]);
    setElapsed(null);
    setExpandId(null);
    setExpandData(null);

    const targets = filesRef.current.filter(f => f.indexed && activeTypes.includes(f.type));
    if (!targets.length) return;

    setSearching(true);
    startTs.current      = Date.now();
    pendingDone.current  = 0;
    pendingTotal.current = targets.length;

    if (tickTmr.current) clearInterval(tickTmr.current);
    tickTmr.current = window.setInterval(
      () => setElapsed(Date.now() - startTs.current), 100
    );

    targets.forEach((f, i) => {
      workers.current[i % NUM_WORKERS].postMessage({
        action: 'SEARCH_FILE',
        payload: { fileId:f.id, query:q, exactMatch:ex, fuzzy:fuzz, searchId:sid }
      });
    });
  }, []);

  const triggerSearch = useCallback((q: string) => {
    if (debTmr.current) clearTimeout(debTmr.current);
    if (!q.trim()) {
      setResults([]); allResults.current = []; setElapsed(null);
      return;
    }
    debTmr.current = window.setTimeout(
      () => runSearch(q, fuzzy, exact, types), DEBOUNCE
    );
  }, [runSearch, fuzzy, exact, types]);

  // Re-search on option change
  useEffect(() => {
    if (query.trim() && filesRef.current.some(f => f.indexed)) {
      runSearch(query, fuzzy, exact, types);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fuzzy, exact, types]);

  // ── Row expand ────────────────────────────────────────────────────────────
  const handleExpand = useCallback((r: SearchResult) => {
    if (expandId === r.id) { setExpandId(null); setExpandData(null); return; }
    setExpandId(r.id);
    setExpandData('loading');
    workers.current[0].postMessage({
      action:'FETCH_ROW', payload:{fileId:r.fileId, ri:r._ri}
    });
  }, [expandId]);

  // ── Sort ──────────────────────────────────────────────────────────────────
  const sorted = useMemo(() => {
    const arr = results.slice();
    const m   = sortDir==='asc' ? 1 : -1;
    if      (sortKey==='relevance') arr.sort((a,b) => m*((b.score||0)-(a.score||0)));
    else if (sortKey==='fileName')  arr.sort((a,b) => m*a.fileName.localeCompare(b.fileName));
    else                            arr.sort((a,b) => m*(a.rowNumber-b.rowNumber));
    return arr;
  }, [results, sortKey, sortDir]);

  // ── Virtual scroll ────────────────────────────────────────────────────────
  const visible = useMemo(() => {
    const start = Math.max(0, Math.floor(scrollTop/ROW_H) - OVER);
    const end   = Math.min(sorted.length, Math.ceil((scrollTop+viewH)/ROW_H) + OVER);
    return {items:sorted.slice(start,end), start};
  }, [sorted, scrollTop, viewH]);

  // ── Derived ───────────────────────────────────────────────────────────────
  const indexedCount  = files.filter(f => f.indexed).length;
  const indexingCount = files.filter(f => !f.indexed && !f.error).length;
  const totalSize     = files.reduce((s,f)=>s+f.size,0);
  const allIndexed    = files.length > 0 && !indexing && indexedCount === files.length;
  const capped        = allResults.current.length >= MAX_RENDER;

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div style={{display:'flex',flexDirection:'column',height:'100vh',overflow:'hidden',
      background:'#06070a',color:'#94a3b8',
      fontFamily:"'DM Mono','Fira Code','Cascadia Code',monospace"}}>

      {/* ═══ HEADER ═══ */}
      <div style={{flexShrink:0,background:'#08090d',borderBottom:'1px solid #0f1520'}}>

        {/* Top row */}
        <div style={{display:'flex',alignItems:'center',gap:10,padding:'10px 16px'}}>

          {/* Logo */}
          <div style={{display:'flex',alignItems:'center',gap:8,flexShrink:0}}>
            <div style={{width:30,height:30,borderRadius:8,flexShrink:0,
              background:'linear-gradient(135deg,#1d4ed8,#4338ca)',
              display:'flex',alignItems:'center',justifyContent:'center',
              boxShadow:'0 0 14px #1d4ed840'}}>
              <Zap size={15} color="#fff"/>
            </div>
            <span style={{fontSize:13,fontWeight:700,letterSpacing:'0.08em',
              color:'#e2e8f0',userSelect:'none'}}>
              LOCAL<span style={{color:'#3b82f6'}}>LENS</span>
            </span>
          </div>

          {/* Search input */}
          <div style={{flex:1,position:'relative',minWidth:0}}>
            <Search size={14} style={{position:'absolute',left:11,top:'50%',
              transform:'translateY(-50%)',pointerEvents:'none',
              color:searching?'#3b82f6':'#1e3050',transition:'color 0.2s'}}/>
            <input
              value={query}
              disabled={!allIndexed}
              onChange={e=>{setQuery(e.target.value);triggerSearch(e.target.value);}}
              placeholder={
                indexing     ? `Indexing ${indexingCount} files…` :
                !files.length? 'Open a folder to begin' :
                !allIndexed  ? 'Waiting for index…' :
                `Search ${fmt(totalRows)} rows across ${files.length} files`
              }
              style={{width:'100%',boxSizing:'border-box',
                padding:'9px 36px 9px 34px',
                background:'#0d1117',border:'1px solid #161e2e',
                borderRadius:8,color:'#e2e8f0',fontSize:13,
                outline:'none',fontFamily:'inherit',
                opacity:allIndexed?1:0.5,
                transition:'border-color 0.15s,opacity 0.3s'}}
              onFocus={e=>e.target.style.borderColor='#1d4ed8'}
              onBlur={e =>e.target.style.borderColor='#161e2e'}
            />
            {query && (
              <button onClick={()=>{setQuery('');setResults([]);allResults.current=[];setElapsed(null);}}
                style={{position:'absolute',right:9,top:'50%',transform:'translateY(-50%)',
                  background:'none',border:'none',cursor:'pointer',color:'#334155',padding:2,lineHeight:0}}>
                <X size={13}/>
              </button>
            )}
            {searching && (
              <Loader2 size={13} color="#3b82f6"
                style={{position:'absolute',right:query?30:9,top:'50%',
                  transform:'translateY(-50%)',animation:'spin 0.7s linear infinite'}}/>
            )}
          </div>

          {/* Toggles */}
          <Tog active={fuzzy} onClick={()=>setFuzzy(!fuzzy)} label="FUZZY"/>
          <Tog active={exact} onClick={()=>setExact(!exact)} label="EXACT"/>
          <Sep/>
          {(['xlsx','xls','csv','txt'] as ExtKey[]).map(t=>(
            <Tog key={t}
              active={types.includes(t)}
              onClick={()=>setTypes(p=>p.includes(t)?p.filter(x=>x!==t):[...p,t])}
              label={t.toUpperCase()} color={EXT[t]?.color}/>
          ))}
          <Sep/>
          {(['relevance','fileName','rowNumber'] as SortKey[]).map(k=>(
            <SortB key={k} label={k==='relevance'?'SCORE':k==='fileName'?'FILE':'ROW'}
              active={sortKey===k} dir={sortDir}
              onClick={()=>{
                if(sortKey===k) setSortDir(d=>d==='asc'?'desc':'asc');
                else {setSortKey(k);setSortDir('desc');}
              }}/>
          ))}
          <Sep/>

          {/* Folder */}
          <label style={{display:'flex',alignItems:'center',gap:6,
            padding:'6px 11px',borderRadius:7,cursor:'pointer',
            background:'#0d1117',border:'1px solid #161e2e',
            fontSize:10,fontWeight:700,color:'#334155',
            whiteSpace:'nowrap',flexShrink:0,letterSpacing:'0.06em'}}>
            <FolderOpen size={12}/>
            {files.length?'CHANGE':'OPEN FOLDER'}
            <input type="file" style={{display:'none'}}
              // @ts-ignore
              webkitdirectory="" directory="" multiple onChange={handleFolder}/>
          </label>

          {/* Export */}
          {allResults.current.length>0 && (
            <button
              onClick={()=>workers.current[0].postMessage({
                action:'EXPORT_MATCHES',payload:{matches:allResults.current}})}
              style={{display:'flex',alignItems:'center',gap:5,
                padding:'6px 11px',borderRadius:7,border:'none',cursor:'pointer',
                background:'linear-gradient(135deg,#1d4ed8,#4338ca)',
                fontSize:10,fontWeight:700,color:'#fff',flexShrink:0}}>
              <Download size={12}/> CSV
            </button>
          )}
        </div>

        {/* Status bar */}
        <div style={{display:'flex',alignItems:'center',gap:12,
          padding:'4px 16px 6px',borderTop:'1px solid #0a0e17'}}>
          {files.length>0 && (
            <>
              <SP
                icon={indexing
                  ?<Loader2 size={9} style={{animation:'spin 0.7s linear infinite'}}/>
                  :<CheckCircle2 size={9}/>}
                label={indexing?`Indexing ${indexedCount}/${files.length}`:`${indexedCount} files`}
                color={indexing?'#f59e0b':'#22c55e'}/>
              <SP icon={<Database size={9}/>} label={fmt(totalRows)+' rows'} color="#334155"/>
              <SP icon={<Clock size={9}/>}    label={fmtBytes(totalSize)}    color="#334155"/>
            </>
          )}
          <div style={{flex:1}}/>
          {results.length>0 && (
            <span style={{fontSize:10,fontWeight:700,
              color:capped?'#f59e0b':'#3b82f6'}}>
              {capped?`${fmt(MAX_RENDER)}+ results`:`${fmt(results.length)} results`}
            </span>
          )}
          {elapsed!==null && (
            <span style={{fontSize:10,
              color:searching?'#f59e0b':'#1e3a5f',
              fontWeight:searching?700:400}}>
              {fmtTime(elapsed)}{searching?'…':''}
            </span>
          )}
          {searching && (
            <div style={{width:50,height:2,background:'#0f1520',borderRadius:1}}>
              <div style={{height:'100%',background:'#3b82f6',borderRadius:1,
                animation:'pulse 1s ease-in-out infinite'}}/>
            </div>
          )}
        </div>
      </div>

      {/* ═══ BODY ═══ */}
      <div ref={scrollRef}
        onScroll={e=>setScrollTop(e.currentTarget.scrollTop)}
        style={{flex:1,overflowY:'auto',overflowX:'hidden'}}>

        {files.length===0 && <Landing/>}

        {files.length>0 && indexing && results.length===0 && !query && (
          <IndexingView files={files}/>
        )}

        {allIndexed && !query && results.length===0 && (
          <Ready count={files.length} rows={totalRows}/>
        )}

        {allIndexed && query && results.length===0 && !searching && (
          <div style={{height:'100%',display:'flex',flexDirection:'column',
            alignItems:'center',justifyContent:'center',gap:8,opacity:0.4}}>
            <Search size={32} color="#1e3a5f"/>
            <div style={{fontSize:12,color:'#334155'}}>No results for "{query}"</div>
          </div>
        )}

        {sorted.length>0 && (
          <div style={{position:'relative',height:sorted.length*ROW_H}}>
            <div style={{position:'absolute',top:0,left:0,right:0,
              transform:`translateY(${visible.start*ROW_H}px)`}}>
              {visible.items.map(r=>(
                <ResultRow key={r.id} r={r}
                  expanded={expandId===r.id}
                  expandData={expandId===r.id?expandData:undefined}
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

// ═══ Shared sub-components ════════════════════════════════════════════════════
const Sep = ()=><div style={{width:1,height:16,background:'#0f1520',flexShrink:0}}/>;

const Tog = memo(({active,onClick,label,color}:{active:boolean;onClick:()=>void;label:string;color?:string})=>(
  <button onClick={onClick} style={{
    padding:'4px 9px',borderRadius:5,cursor:'pointer',flexShrink:0,
    background:active?(color?color+'15':'#0f2040'):'transparent',
    border:'1px solid '+(active?(color||'#1d4ed8'):'#0f1520'),
    color:active?(color||'#60a5fa'):'#1e3050',
    fontSize:9,fontWeight:800,letterSpacing:'0.1em',transition:'all 0.1s'
  }}>{label}</button>
));

const SortB = memo(({label,active,dir,onClick}:{label:string;active:boolean;dir:SortDir;onClick:()=>void})=>(
  <button onClick={onClick} style={{
    padding:'3px 8px',borderRadius:5,cursor:'pointer',flexShrink:0,
    background:active?'#0a1628':'transparent',
    border:'1px solid '+(active?'#0f2040':'transparent'),
    color:active?'#3b82f6':'#1e3050',fontSize:9,fontWeight:700
  }}>
    {label}{active?(dir==='asc'?' ↑':' ↓'):''}
  </button>
));

const SP = ({icon,label,color}:{icon:React.ReactNode;label:string;color:string})=>(
  <span style={{display:'flex',alignItems:'center',gap:4,fontSize:9,color,fontWeight:600}}>
    {icon}{label}
  </span>
);

// ── Landing ───────────────────────────────────────────────────────────────────
const Landing = memo(()=>(
  <div style={{height:'100%',display:'flex',flexDirection:'column',
    alignItems:'center',justifyContent:'center',gap:20,padding:40}}>
    <div style={{width:64,height:64,borderRadius:16,
      background:'linear-gradient(135deg,#0a1628,#111827)',
      border:'1px solid #0f1e30',
      display:'flex',alignItems:'center',justifyContent:'center'}}>
      <FolderOpen size={26} color="#1d4ed840"/>
    </div>
    <div style={{textAlign:'center'}}>
      <div style={{fontSize:14,fontWeight:700,color:'#0f2040',
        letterSpacing:'0.08em',marginBottom:8}}>OPEN A FOLDER TO START</div>
      <div style={{fontSize:11,color:'#0a1628',maxWidth:320,lineHeight:1.9}}>
        Files are indexed <span style={{color:'#1d4ed8'}}>once</span>.
        All searches run at native JS speed — no re-parsing ever.
      </div>
    </div>
    <div style={{display:'flex',gap:6}}>
      {['XLSX','XLS','CSV','TXT'].map(f=>(
        <span key={f} style={{padding:'3px 10px',borderRadius:20,
          background:'#08090d',border:'1px solid #0d1220',
          fontSize:9,color:'#0f2040',fontWeight:700}}>{f}</span>
      ))}
    </div>
  </div>
));

// ── Indexing screen ───────────────────────────────────────────────────────────
const IndexingView = memo(({files}:{files:FileEntry[]})=>{
  const done  = files.filter(f=>f.indexed||!!f.error).length;
  const total = files.length;
  const pct   = total?Math.round(done/total*100):0;
  return (
    <div style={{height:'100%',display:'flex',flexDirection:'column',
      alignItems:'center',justifyContent:'center',gap:16,padding:40}}>
      <Loader2 size={28} color="#1d4ed8" style={{animation:'spin 0.8s linear infinite'}}/>
      <div style={{textAlign:'center'}}>
        <div style={{fontSize:12,fontWeight:700,color:'#1e3050',
          letterSpacing:'0.08em',marginBottom:6}}>BUILDING SEARCH INDEX</div>
        <div style={{fontSize:10,color:'#0f2040'}}>{done} / {total} files — {pct}%</div>
      </div>
      <div style={{width:200,height:3,background:'#0a0e17',borderRadius:2}}>
        <div style={{height:'100%',width:pct+'%',borderRadius:2,
          background:'linear-gradient(90deg,#1d4ed8,#4338ca)',transition:'width 0.3s'}}/>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:4,width:260}}>
        {files.slice(0,8).map(f=>(
          <div key={f.id} style={{display:'flex',alignItems:'center',gap:8,fontSize:9}}>
            {f.indexed
              ?<CheckCircle2 size={10} color="#22c55e"/>
              :f.error
              ?<AlertCircle  size={10} color="#ef4444"/>
              :<Loader2 size={10} color="#3b82f6" style={{animation:'spin 0.8s linear infinite'}}/>}
            <span style={{color:f.indexed?'#1e3050':f.error?'#ef4444':'#0f2040',
              overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',flex:1}}>
              {f.name}
            </span>
            {f.indexed&&<span style={{color:'#0f2040',flexShrink:0}}>{fmt(f.rowCount)} rows</span>}
          </div>
        ))}
        {files.length>8&&<div style={{fontSize:9,color:'#0a1628',textAlign:'center'}}>
          +{files.length-8} more
        </div>}
      </div>
    </div>
  );
});

// ── Ready ─────────────────────────────────────────────────────────────────────
const Ready = memo(({count,rows}:{count:number;rows:number})=>(
  <div style={{height:'100%',display:'flex',flexDirection:'column',
    alignItems:'center',justifyContent:'center',gap:10,opacity:0.4}}>
    <Search size={32} color="#1d4ed8"/>
    <div style={{fontSize:11,color:'#1e3050',letterSpacing:'0.1em'}}>
      {count} FILES · {fmt(rows)} ROWS · READY
    </div>
  </div>
));

// ── Highlight ─────────────────────────────────────────────────────────────────
const Hl = memo(({text,q}:{text:string;q:string})=>{
  if(!q.trim()||!text) return <>{text}</>;
  const terms=q.trim().split(/\s+/).filter(t=>t.length>1);
  if(!terms.length) return <>{text}</>;
  try{
    const re=new RegExp(`(${terms.map(t=>t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|')})`, 'gi');
    return <>{text.split(re).map((p,i)=>
      re.test(p)?<mark key={i} style={{background:'#1a3560',color:'#93c5fd',borderRadius:2,padding:'0 1px'}}>{p}</mark>:p
    )}</>;
  }catch{return <>{text}</>;}
});

// ── Result row ────────────────────────────────────────────────────────────────
const ResultRow = memo(({r,expanded,expandData,onExpand,query}:{
  r:SearchResult; expanded:boolean;
  expandData?:Record<string,string>|null|'loading';
  onExpand:(r:SearchResult)=>void; query:string;
})=>{
  const es = EXT[extOf(r.fileName)]||EXT.xlsx;
  return (
    <div style={{borderBottom:'1px solid #08090f',
      background:expanded?'#08111e':'transparent'}}>
      <div onClick={()=>onExpand(r)}
        style={{height:ROW_H,display:'flex',alignItems:'center',
          padding:'0 14px',cursor:'pointer',gap:10}}
        onMouseEnter={e=>{if(!expanded)(e.currentTarget as HTMLDivElement).style.background='#080d14';}}
        onMouseLeave={e=>{if(!expanded)(e.currentTarget as HTMLDivElement).style.background='transparent';}}>

        <div style={{width:30,height:30,borderRadius:7,flexShrink:0,
          background:es.dim,border:'1px solid '+es.color+'25',
          display:'flex',alignItems:'center',justifyContent:'center',color:es.color}}>
          <es.Icon size={13}/>
        </div>

        <div style={{flex:1,minWidth:0}}>
          <div style={{display:'flex',alignItems:'center',gap:7,marginBottom:2}}>
            <span style={{fontSize:11,fontWeight:700,color:'#334155',
              overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:240}}>
              {r.fileName}
            </span>
            {r.sheetName&&(
              <span style={{fontSize:8,color:'#2563eb',background:'#0a1628',
                border:'1px solid #0f2040',borderRadius:3,padding:'1px 5px',flexShrink:0}}>
                {r.sheetName}
              </span>
            )}
            <span style={{fontSize:9,color:'#0f1e30',flexShrink:0}}>#{r.rowNumber}</span>
          </div>
          <div style={{fontSize:10,color:'#1e3050',
            overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
            <Hl text={r.searchString} q={query}/>
          </div>
        </div>

        {r.score>1&&(
          <div style={{flexShrink:0,padding:'1px 5px',borderRadius:4,
            background:'#081528',border:'1px solid #0f2040',
            fontSize:9,color:'#1d4ed8',fontWeight:700}}>
            {r.score}
          </div>
        )}

        <ChevronDown size={11} color="#0f1e30"
          style={{flexShrink:0,transition:'transform 0.15s',
            transform:expanded?'rotate(180deg)':'none'}}/>
      </div>

      {expanded&&(
        <div style={{padding:'2px 14px 12px',borderTop:'1px solid #08090f'}}>
          {expandData==='loading'?(
            <div style={{display:'flex',alignItems:'center',gap:8,
              padding:'10px 0',color:'#1e3050',fontSize:10}}>
              <Loader2 size={11} style={{animation:'spin 0.7s linear infinite'}}/> Loading…
            </div>
          ):!expandData||Object.keys(expandData).length===0?(
            <div style={{display:'flex',alignItems:'center',gap:6,
              padding:'10px 0',color:'#0f1e30',fontSize:10}}>
              <AlertCircle size={11}/> No data
            </div>
          ):(
            <div style={{display:'grid',
              gridTemplateColumns:'repeat(auto-fill,minmax(150px,1fr))',
              gap:5,paddingTop:8}}>
              {Object.entries(expandData).map(([k,v])=>(
                <div key={k} style={{padding:'5px 7px',borderRadius:5,
                  background:'#080d14',border:'1px solid #0a1220'}}>
                  <div style={{fontSize:8,color:'#0f1e30',fontWeight:700,
                    letterSpacing:'0.1em',marginBottom:2}}>{k.toUpperCase()}</div>
                  <div style={{fontSize:10,color:'#1e3050',wordBreak:'break-word'}}>
                    <Hl text={String(v??'—')} q={query}/>
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

// ── Global CSS ────────────────────────────────────────────────────────────────
const _s=document.createElement('style');
_s.textContent=`
  @keyframes spin{to{transform:rotate(360deg)}}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}
  *{-webkit-font-smoothing:antialiased;box-sizing:border-box}
  ::-webkit-scrollbar{width:3px}
  ::-webkit-scrollbar-track{background:transparent}
  ::-webkit-scrollbar-thumb{background:#0a0e17;border-radius:3px}
  input::placeholder{color:#0f1e30}
`;
document.head.appendChild(_s);
