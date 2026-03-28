import React, {
  useEffect, useRef, useState, useMemo,
  useCallback, memo, startTransition
} from 'react';
import {
  Zap, FolderOpen, Search, X, Download, Loader2,
  CheckCircle2, AlertCircle, ChevronDown,
  FileSpreadsheet, FileCode, FileText, ScanSearch
} from 'lucide-react';
import workerScript from './engine.js?raw';

// ── Types ─────────────────────────────────────────────────────────────────────
interface FM {
  id: string; name: string; type: string; blob: File; size: number;
  st: 'q' | 'ing' | 'ok' | 'err'; rows: number; err?: string;
}
interface Hit {
  id: string; fid: string; fn: string; sn: string | null;
  rn: number; ss: string; sc: number; ri: number;
}

// ── Constants ─────────────────────────────────────────────────────────────────
const ROW_H  = 64;
const OVER   = 12;
const DEB    = 180;
const MAX_H  = 20000;
const SUPP   = new Set(['xlsx','xls','csv','txt']);

const fmtN = (n: number) =>
  n >= 1e6 ? (n/1e6).toFixed(1)+'M' : n >= 1e3 ? (n/1e3).toFixed(1)+'K' : ''+n;
const fmtB = (b: number) =>
  b >= 1<<20 ? (b/(1<<20)).toFixed(1)+' MB' : b >= 1<<10 ? (b/(1<<10)).toFixed(1)+' KB' : b+' B';
const fmtT = (ms: number) =>
  ms >= 60000 ? (ms/60000).toFixed(1)+'m' : ms >= 1000 ? (ms/1000).toFixed(2)+'s' : ms+'ms';

type Ext = 'xlsx'|'xls'|'csv'|'txt';
const EXT: Record<Ext,{c:string;dim:string;I:any}> = {
  xlsx: { c:'#38bdf8', dim:'rgba(56,189,248,0.08)',   I: FileSpreadsheet },
  xls:  { c:'#818cf8', dim:'rgba(129,140,248,0.08)',  I: FileSpreadsheet },
  csv:  { c:'#4ade80', dim:'rgba(74,222,128,0.08)',   I: FileCode },
  txt:  { c:'#fb923c', dim:'rgba(251,146,60,0.08)',   I: FileText },
};
const eOf = (n: string) => (n.split('.').pop()?.toLowerCase() || 'xlsx') as Ext;

const C = {
  bg:'#0a0b0e', surf:'#0f1117', surf2:'#141720',
  brd:'#1c2030', brd2:'#252b3b',
  txt:'#e2e8f0', muted:'#4a5568', dim:'#2d3748',
  acc:'#f59e0b', blue:'#38bdf8', green:'#4ade80', red:'#f87171',
};

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  // ── State ──────────────────────────────────────────────────────────────────
  const [files,      setFiles]      = useState<FM[]>([]);
  const [totRows,    setTotRows]    = useState(0);
  const [query,      setQuery]      = useState('');
  const [hits,       setHits]       = useState<Hit[]>([]);
  const [sortedHits, setSortedHits] = useState<Hit[]>([]);
  const [busy,       setBusy]       = useState(false);
  const [elapsed,    setElapsed]    = useState<number|null>(null);
  const [fuzzy,      setFuzzy]      = useState(true);
  const [exact,      setExact]      = useState(false);
  const [exts,       setExts]       = useState(['xlsx','xls','csv','txt']);
  const [sortK,      setSortK]      = useState<'sc'|'fn'|'rn'>('sc');
  const [sortD,      setSortD]      = useState<1|-1>(-1);
  const [expId,      setExpId]      = useState<string|null>(null);
  const [expD,       setExpD]       = useState<Record<string,string>|null>(null);
  const [expLoad,    setExpLoad]    = useState(false);
  const [hitCount,   setHitCount]   = useState(0);
  const [scrollTop,  setScrollTop]  = useState(0);
  const [viewH,      setViewH]      = useState(600);

  const bodyEl = useRef<HTMLDivElement>(null);

  // ── Refs (no re-renders) ──────────────────────────────────────────────────
  const W        = useRef<Worker|null>(null);       // single worker
  const iQueue   = useRef<FM[]>([]);                // index queue
  const indexed  = useRef<Map<string,FM>>(new Map()); // indexed files
  const sid      = useRef(0);
  const sDone    = useRef(0);
  const sTotal   = useRef(0);
  const sStart   = useRef(0);
  const hitBuf   = useRef<Hit[]>([]);
  const allHits  = useRef<Hit[]>([]);
  const debTmr   = useRef(0);
  const tickTmr  = useRef(0);
  const qRef     = useRef('');
  const fuzzyRef = useRef(true);
  const exactRef = useRef(false);
  const extsRef  = useRef(['xlsx','xls','csv','txt']);
  const expIdRef = useRef<string|null>(null);

  // Keep option refs in sync with state
  useEffect(() => { fuzzyRef.current = fuzzy; }, [fuzzy]);
  useEffect(() => { exactRef.current = exact; }, [exact]);
  useEffect(() => { extsRef.current  = exts;  }, [exts]);

  // ── Resize observer ────────────────────────────────────────────────────────
  useEffect(() => {
    const ro = new ResizeObserver(() => {
      if (bodyEl.current) setViewH(bodyEl.current.clientHeight);
    });
    if (bodyEl.current) ro.observe(bodyEl.current);
    return () => ro.disconnect();
  }, []);

  // ── Sort hits via startTransition ──────────────────────────────────────────
  useEffect(() => {
    startTransition(() => {
      const a = hits.slice();
      if      (sortK === 'sc') a.sort((x,y) => sortD*(y.sc-x.sc));
      else if (sortK === 'fn') a.sort((x,y) => sortD*x.fn.localeCompare(y.fn));
      else                     a.sort((x,y) => sortD*(x.rn-y.rn));
      setSortedHits(a);
    });
  }, [hits, sortK, sortD]);

  // ── Flush hit buffer → state ───────────────────────────────────────────────
  const flushHits = useCallback(() => {
    if (!hitBuf.current.length) return;
    const snap = hitBuf.current.splice(0);
    Array.prototype.push.apply(allHits.current, snap);
    if (allHits.current.length <= MAX_H) {
      setHits(allHits.current.slice());
    }
    setHitCount(allHits.current.length);
  }, []);

  // ── Dispatch next file from index queue ────────────────────────────────────
  const dispatchNext = useCallback(() => {
    if (!iQueue.current.length || !W.current) return;
    const f = iQueue.current.shift()!;
    setFiles(prev => prev.map(x => x.id === f.id ? {...x, st:'ing' as const} : x));
    W.current.postMessage({ t:'idx', id:f.id, blob:f.blob, ft:f.type, name:f.name });
  }, []);

  // ── Boot single worker ─────────────────────────────────────────────────────
  useEffect(() => {
    const blob = new Blob([workerScript], { type:'application/javascript' });
    const url  = URL.createObjectURL(blob);
    const w    = new Worker(url);
    W.current  = w;

    w.onmessage = (ev) => {
      const d = ev.data;

      if (d.t === 'ok') {
        // File indexed — add to map synchronously, update UI
        setFiles(prev => {
          const nx = prev.map(x => x.id === d.id ? {...x, st:'ok' as const, rows:d.n} : x);
          const entry = nx.find(x => x.id === d.id);
          if (entry) indexed.current.set(d.id, entry);
          return nx;
        });
        setTotRows(r => r + d.n);
        dispatchNext();

      } else if (d.t === 'err') {
        setFiles(prev => prev.map(x => x.id === d.id ? {...x, st:'err' as const, err:d.msg} : x));
        dispatchNext();

      } else if (d.t === 'hits') {
        if (d.sid !== sid.current) return;
        Array.prototype.push.apply(hitBuf.current, d.hits);
        if (hitBuf.current.length >= 500) flushHits();

      } else if (d.t === 'done') {
        if (d.sid !== sid.current) return;
        sDone.current++;
        if (sDone.current >= sTotal.current) {
          flushHits();
          setBusy(false);
          clearInterval(tickTmr.current);
          setElapsed(Date.now() - sStart.current);
        }

      } else if (d.t === 'row') {
        setExpD(d.d);
        setExpLoad(false);

      } else if (d.t === 'csv') {
        const a  = document.createElement('a');
        a.href   = URL.createObjectURL(d.blob);
        a.download = 'locallens_' + Date.now() + '.csv';
        a.click();
      }
    };

    return () => {
      w.terminate();
      URL.revokeObjectURL(url);
      clearInterval(tickTmr.current);
      clearTimeout(debTmr.current);
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Load folder ────────────────────────────────────────────────────────────
  const loadFolder = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const raw = e.target.files;
    if (!raw || !raw.length) return;
    e.target.value = ''; // allow re-selecting same folder

    // Clear worker DB and all state
    W.current?.postMessage({ t:'clear' });
    indexed.current.clear();
    iQueue.current  = [];
    allHits.current = [];
    hitBuf.current  = [];

    setFiles([]);
    setHits([]);
    setSortedHits([]);
    setQuery('');       qRef.current = '';
    setElapsed(null);
    setTotRows(0);
    setExpId(null);     expIdRef.current = null;
    setExpD(null);
    setHitCount(0);
    setBusy(false);
    clearInterval(tickTmr.current);

    // Collect supported files
    const list: FM[] = [];
    for (let i = 0; i < raw.length; i++) {
      const f   = raw[i];
      const ext = (f.name.split('.').pop() || '').toLowerCase();
      if (!SUPP.has(ext)) continue;
      list.push({
        id:   `${f.name}-${f.size}-${f.lastModified}`,
        name: f.name, type: ext, blob: f, size: f.size, st: 'q', rows: 0
      });
    }
    if (!list.length) return;

    setFiles(list);
    iQueue.current = list.slice();
    dispatchNext(); // start indexing first file
  }, [dispatchNext]);

  // ── Search ─────────────────────────────────────────────────────────────────
  const runSearch = useCallback((q: string) => {
    if (!q.trim()) {
      setHits([]); setSortedHits([]);
      allHits.current = []; hitBuf.current = [];
      setElapsed(null); setHitCount(0);
      return;
    }

    const targets = Array.from(indexed.current.values())
      .filter(f => extsRef.current.includes(f.type));
    if (!targets.length) return;

    const newSid = ++sid.current;
    hitBuf.current  = [];
    allHits.current = [];
    sDone.current   = 0;
    sTotal.current  = targets.length;
    sStart.current  = Date.now();

    setHits([]); setSortedHits([]);
    setElapsed(null); setHitCount(0);
    setBusy(true);
    setExpId(null); expIdRef.current = null;
    setExpD(null);

    clearInterval(tickTmr.current);
    tickTmr.current = setInterval(
      () => setElapsed(Date.now() - sStart.current), 100
    ) as unknown as number;

    // Send all search messages to the single worker (queued automatically)
    targets.forEach(f => {
      W.current?.postMessage({
        t: 'search', sid: newSid, id: f.id,
        q, fuzzy: fuzzyRef.current, exact: exactRef.current
      });
    });
  }, []);

  const trigSearch = useCallback((q: string) => {
    clearTimeout(debTmr.current);
    qRef.current = q;
    debTmr.current = setTimeout(() => runSearch(q), DEB) as unknown as number;
  }, [runSearch]);

  // Re-run search when mode/type filters change
  useEffect(() => {
    if (qRef.current.trim()) runSearch(qRef.current);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fuzzy, exact, exts]);

  // ── Row expand ─────────────────────────────────────────────────────────────
  const expand = useCallback((h: Hit) => {
    if (expIdRef.current === h.id) {
      expIdRef.current = null;
      setExpId(null); setExpD(null);
      return;
    }
    expIdRef.current = h.id;
    setExpId(h.id); setExpD(null); setExpLoad(true);
    W.current?.postMessage({ t:'row', id:h.fid, ri:h.ri });
  }, []);

  // ── Virtual scroll window ──────────────────────────────────────────────────
  const {vis, vs} = useMemo(() => {
    const s = Math.max(0, Math.floor(scrollTop/ROW_H) - OVER);
    const e = Math.min(sortedHits.length, Math.ceil((scrollTop+viewH)/ROW_H) + OVER);
    return { vis: sortedHits.slice(s, e), vs: s };
  }, [sortedHits, scrollTop, viewH]);

  // ── Derived ────────────────────────────────────────────────────────────────
  const nOk    = files.filter(f => f.st === 'ok').length;
  const isIdx  = files.some(f => f.st === 'ing' || f.st === 'q');
  const ready  = !isIdx && nOk > 0;
  const totSz  = files.reduce((s,f) => s+f.size, 0);
  const pct    = files.length ? Math.round(nOk/files.length*100) : 0;
  const capped = hitCount >= MAX_H;

  // ── Render ─────────────────────────────────────────────────────────────────
  return (
    <div className="ll-root">
      <header className="ll-header">
        <div className="ll-topbar">

          {/* Logo */}
          <div className="ll-logo">
            <div className="ll-logo-icon"><Zap size={15} color={C.bg}/></div>
            <span className="ll-logo-text">LOCAL<span className="ll-acc">LENS</span></span>
          </div>

          {/* Search */}
          <div className="ll-search-wrap">
            <span className="ll-search-ico">
              {busy
                ? <Loader2 size={14} className="spin" style={{color:C.acc}}/>
                : <Search  size={14} style={{color: query ? C.acc : C.muted}}/>}
            </span>
            <input className="ll-input" value={query} disabled={!ready}
              autoComplete="off" spellCheck={false}
              placeholder={
                isIdx  ? `Indexing ${nOk}/${files.length}…` :
                !files.length ? 'Open a folder to start' :
                !ready ? 'Building index…' :
                `Search ${fmtN(totRows)} rows in ${files.length} files`
              }
              onChange={e => { setQuery(e.target.value); trigSearch(e.target.value); }}
            />
            {query && (
              <button className="ll-clear" onClick={() => { setQuery(''); runSearch(''); }}>
                <X size={13}/>
              </button>
            )}
          </div>

          {/* Mode */}
          <Chip on={fuzzy} click={() => setFuzzy(v=>!v)} label="FUZZY" title="All words must appear"/>
          <Chip on={exact} click={() => setExact(v=>!v)} label="EXACT" title="Exact cell match"/>
          <Sep/>

          {/* Types */}
          {(['xlsx','xls','csv','txt'] as Ext[]).map(t => (
            <TypeChip key={t} t={t} on={exts.includes(t)}
              click={() => setExts(p => p.includes(t) ? p.filter(x=>x!==t) : [...p,t])}/>
          ))}
          <Sep/>

          {/* Sort */}
          {(['sc','fn','rn'] as const).map((k,i) => (
            <SortChip key={k} label={['SCORE','FILE','ROW'][i]}
              active={sortK===k} dir={sortD}
              click={() => { if(sortK===k) setSortD(d=>d===1?-1:1); else {setSortK(k);setSortD(-1);} }}/>
          ))}
          <Sep/>

          {/* Folder */}
          <label className="ll-btn ll-btn-ghost">
            <FolderOpen size={13}/>
            {files.length ? 'CHANGE' : 'OPEN FOLDER'}
            <input type="file" style={{display:'none'}}
              // @ts-ignore
              webkitdirectory="" directory="" multiple onChange={loadFolder}/>
          </label>

          {/* Export */}
          {hitCount > 0 && (
            <button className="ll-btn ll-btn-acc"
              onClick={() => W.current?.postMessage({t:'export', hits:allHits.current})}>
              <Download size={13}/> EXPORT
            </button>
          )}
        </div>

        {/* Status bar */}
        <div className="ll-status">
          <div className="ll-status-l">
            {files.length > 0 && <>
              <Dot ok={!isIdx}/>
              <span className="ll-status-lbl" style={{color: isIdx ? C.acc : C.green}}>
                {isIdx ? `INDEXING ${nOk}/${files.length}` : `${nOk} FILES READY`}
              </span>
              <div className="ll-sdiv"/>
              <span className="ll-stat">{fmtN(totRows)} ROWS</span>
              <div className="ll-sdiv"/>
              <span className="ll-stat">{fmtB(totSz)}</span>
              {isIdx && <>
                <div className="ll-sdiv"/>
                <div className="ll-pbar"><div className="ll-pbar-fill" style={{width:pct+'%'}}/></div>
                <span className="ll-pct">{pct}%</span>
              </>}
            </>}
          </div>
          <div className="ll-status-r">
            {hitCount > 0 && (
              <span className="ll-res" style={{color: capped ? C.acc : C.blue}}>
                {capped ? fmtN(MAX_H)+'+' : fmtN(hitCount)} RESULTS
              </span>
            )}
            {elapsed !== null && (
              <span className="ll-time" style={{color: busy ? C.acc : C.muted}}>
                {fmtT(elapsed)}{busy ? '…' : ''}
              </span>
            )}
          </div>
        </div>
      </header>

      {/* Body */}
      <div className="ll-body" ref={bodyEl}
        onScroll={e => setScrollTop(e.currentTarget.scrollTop)}>

        {files.length === 0 && <Welcome/>}
        {files.length > 0 && isIdx && !query && sortedHits.length === 0 && <Indexing files={files}/>}
        {ready && !query && sortedHits.length === 0 && <Ready n={files.length} rows={totRows}/>}
        {ready && query && sortedHits.length === 0 && !busy && <NoResults q={query}/>}

        {sortedHits.length > 0 && (
          <div style={{position:'relative', height: sortedHits.length * ROW_H}}>
            <div style={{position:'absolute', top:0, left:0, right:0,
              transform:`translateY(${vs*ROW_H}px)`}}>
              {vis.map(h => (
                <HitRow key={h.id} h={h}
                  open={expId===h.id}
                  data={expId===h.id ? expD : undefined}
                  loading={expId===h.id && expLoad}
                  onOpen={expand} q={query}/>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Sub-components ────────────────────────────────────────────────────────────
const Sep  = () => <div className="ll-sep"/>;
const Dot  = ({ok}:{ok:boolean}) => (
  <span className="ll-dot" style={{background:ok?C.green:C.acc, boxShadow:`0 0 6px ${ok?C.green:C.acc}`}}/>
);

const Chip = memo(({on,click,label,title}:{on:boolean;click:()=>void;label:string;title?:string}) => (
  <button className={`ll-chip${on?' ll-chip-on':''}`} onClick={click} title={title}>
    {on && <span className="ll-chip-dot"/>}{label}
  </button>
));

const TypeChip = memo(({t,on,click}:{t:Ext;on:boolean;click:()=>void}) => {
  const e = EXT[t];
  return (
    <button className="ll-typechip" onClick={click} style={{
      color: on ? e.c : C.muted,
      background: on ? e.dim : 'transparent',
      borderColor: on ? e.c+'50' : C.brd,
    }}>{t.toUpperCase()}</button>
  );
});

const SortChip = memo(({label,active,dir,click}:{label:string;active:boolean;dir:1|-1;click:()=>void}) => (
  <button className={`ll-sortchip${active?' ll-sortchip-on':''}`} onClick={click}>
    {label}{active ? (dir===-1?' ↓':' ↑') : ''}
  </button>
));

// ── Screens ───────────────────────────────────────────────────────────────────
const Welcome = memo(() => (
  <div className="ll-center">
    <div className="ll-welcome-icon"><ScanSearch size={30} strokeWidth={1.5} style={{color:C.muted}}/></div>
    <div className="ll-welcome-title">OPEN A FOLDER TO BEGIN</div>
    <div className="ll-welcome-sub">
      Index once, search instantly. All processing is local — nothing leaves your machine.
    </div>
    <div className="ll-badges">
      {(['xlsx','xls','csv','txt'] as Ext[]).map(t => (
        <span key={t} className="ll-badge" style={{color:EXT[t].c,borderColor:EXT[t].c+'40',background:EXT[t].dim}}>
          {t.toUpperCase()}
        </span>
      ))}
    </div>
    <div className="ll-features">
      {['⚡ Parse once, search at JS speed','🔒 100% local','🔍 Fuzzy + exact','📤 Export CSV'].map(s=>(
        <div key={s} className="ll-feature">{s}</div>
      ))}
    </div>
  </div>
));

const Indexing = memo(({files}:{files:FM[]}) => {
  const done = files.filter(f=>f.st==='ok').length;
  const pct  = files.length ? Math.round(done/files.length*100) : 0;
  return (
    <div className="ll-center">
      <Loader2 size={26} className="spin" style={{color:C.acc,marginBottom:16}}/>
      <div className="ll-idx-title">BUILDING INDEX</div>
      <div className="ll-idx-sub">{done} / {files.length} files · {pct}%</div>
      <div className="ll-idx-bar"><div className="ll-idx-fill" style={{width:pct+'%'}}/></div>
      <div className="ll-idx-list">
        {files.slice(0,12).map(f => (
          <div key={f.id} className="ll-idx-row">
            <span className="ll-idx-ico">
              {f.st==='ok'  ? <CheckCircle2 size={10} style={{color:C.green}}/>
              :f.st==='err' ? <AlertCircle  size={10} style={{color:C.red}}/>
              :f.st==='ing' ? <Loader2 size={10} className="spin" style={{color:C.blue}}/>
              : <span className="ll-idx-dot"/>}
            </span>
            <span className="ll-idx-name" style={{color:f.st==='ok'?C.txt:f.st==='err'?C.red:C.muted}}>
              {f.name}
            </span>
            {f.st==='ok'  && <span className="ll-idx-n">{fmtN(f.rows)}</span>}
            {f.st==='err' && <span className="ll-idx-err">ERR</span>}
          </div>
        ))}
        {files.length > 12 && <div className="ll-idx-more">+{files.length-12} more queued</div>}
      </div>
    </div>
  );
});

const Ready = memo(({n,rows}:{n:number;rows:number}) => (
  <div className="ll-center" style={{opacity:0.35}}>
    <Search size={34} strokeWidth={1} style={{color:C.acc}}/>
    <div style={{fontSize:11,color:C.muted,letterSpacing:'0.12em',marginTop:12}}>
      {n} FILES · {fmtN(rows)} ROWS · TYPE TO SEARCH
    </div>
  </div>
));

const NoResults = memo(({q}:{q:string}) => (
  <div className="ll-center" style={{opacity:0.5}}>
    <Search size={30} strokeWidth={1} style={{color:C.muted}}/>
    <div style={{fontSize:12,color:C.muted,marginTop:12}}>
      No matches for <span style={{color:C.txt}}>"{q}"</span>
    </div>
    <div style={{fontSize:10,color:C.dim,marginTop:4}}>Try fuzzy mode or check type filters</div>
  </div>
));

// ── Highlight ─────────────────────────────────────────────────────────────────
const Hl = memo(({text,q}:{text:string;q:string}) => {
  if (!q.trim() || !text) return <>{text}</>;
  const terms = q.trim().split(/\s+/).filter(t => t.length > 0);
  if (!terms.length) return <>{text}</>;
  try {
    const re = new RegExp(
      `(${terms.map(t=>t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|')})`, 'gi'
    );
    return <>{text.split(re).map((p,i) =>
      re.test(p) ? <mark key={i} className="ll-hl">{p}</mark> : p
    )}</>;
  } catch { return <>{text}</>; }
});

// ── Hit row ───────────────────────────────────────────────────────────────────
const HitRow = memo(({h,open,data,loading,onOpen,q}:{
  h:Hit; open:boolean; data?:Record<string,string>|null;
  loading?:boolean; onOpen:(h:Hit)=>void; q:string;
}) => {
  const es = EXT[eOf(h.fn)] || EXT.xlsx;
  return (
    <div className={`ll-hit${open?' ll-hit-open':''}`}>
      <div className="ll-hit-row" onClick={() => onOpen(h)}>
        <div className="ll-hit-ico" style={{background:es.dim, borderColor:es.c+'35', color:es.c}}>
          <es.I size={13} strokeWidth={1.5}/>
        </div>
        <div className="ll-hit-body">
          <div className="ll-hit-meta">
            <span className="ll-hit-fn">{h.fn}</span>
            {h.sn && <span className="ll-hit-sn">{h.sn}</span>}
            <span className="ll-hit-rn">ROW {h.rn}</span>
          </div>
          <div className="ll-hit-prev"><Hl text={h.ss} q={q}/></div>
        </div>
        <div className="ll-hit-right">
          {h.sc > 1 && <span className="ll-hit-sc">{h.sc}</span>}
          <ChevronDown size={13} className={`ll-chev${open?' ll-chev-open':''}`}/>
        </div>
      </div>

      {open && (
        <div className="ll-detail">
          {loading ? (
            <div className="ll-detail-load">
              <Loader2 size={11} className="spin"/> Loading row…
            </div>
          ) : !data || !Object.keys(data).length ? (
            <div className="ll-detail-empty"><AlertCircle size={11}/> No data</div>
          ) : (
            <div className="ll-detail-grid">
              {Object.entries(data).map(([k,v]) => (
                <div key={k} className="ll-detail-cell">
                  <div className="ll-detail-k">{k}</div>
                  <div className="ll-detail-v"><Hl text={String(v||'—')} q={q}/></div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
});

// ── CSS ───────────────────────────────────────────────────────────────────────
const _css = `
  *, *::before, *::after { box-sizing:border-box; margin:0; padding:0; }
  :root {
    --bg:#0a0b0e; --surf:#0f1117; --surf2:#141720;
    --brd:#1c2030; --brd2:#252b3b;
    --txt:#e2e8f0; --muted:#4a5568; --dim:#2d3748;
    --acc:#f59e0b; --blue:#38bdf8; --green:#4ade80; --red:#f87171;
  }
  html,body,#root { height:100%; overflow:hidden; }
  body { background:var(--bg); }
  @keyframes spin { to { transform:rotate(360deg) } }
  @keyframes fadeIn { from { opacity:0;transform:translateY(3px) } to { opacity:1;transform:none } }
  .spin { animation:spin 0.8s linear infinite; }

  .ll-root { display:flex;flex-direction:column;height:100vh;overflow:hidden;
    font-family:'IBM Plex Mono','Fira Code','Cascadia Code',monospace;
    background:var(--bg);color:var(--txt); }

  /* Header */
  .ll-header { flex-shrink:0;background:var(--surf);border-bottom:1px solid var(--brd); }
  .ll-topbar { display:flex;align-items:center;gap:7px;padding:9px 14px;flex-wrap:wrap; }

  /* Logo */
  .ll-logo { display:flex;align-items:center;gap:8px;flex-shrink:0;margin-right:4px; }
  .ll-logo-icon { width:28px;height:28px;border-radius:7px;background:var(--acc);flex-shrink:0;
    display:flex;align-items:center;justify-content:center;box-shadow:0 0 14px rgba(245,158,11,.4); }
  .ll-logo-text { font-size:13px;font-weight:700;letter-spacing:.08em;color:var(--txt);user-select:none; }
  .ll-acc { color:var(--acc); }

  /* Search */
  .ll-search-wrap { flex:1;min-width:160px;position:relative;display:flex;align-items:center; }
  .ll-search-ico { position:absolute;left:11px;top:50%;transform:translateY(-50%);
    pointer-events:none;line-height:0; }
  .ll-input { width:100%;padding:8px 32px 8px 34px;
    background:var(--surf2);border:1px solid var(--brd2);border-radius:7px;
    color:var(--txt);font-size:13px;font-family:inherit;outline:none;
    transition:border-color .15s,box-shadow .15s; }
  .ll-input:focus { border-color:var(--acc);box-shadow:0 0 0 3px rgba(245,158,11,.1); }
  .ll-input::placeholder { color:var(--muted); }
  .ll-input:disabled { opacity:.35;cursor:not-allowed; }
  .ll-clear { position:absolute;right:8px;top:50%;transform:translateY(-50%);
    background:none;border:none;cursor:pointer;color:var(--muted);
    line-height:0;padding:2px;border-radius:3px;transition:color .1s; }
  .ll-clear:hover { color:var(--txt); }

  /* Chips */
  .ll-sep { width:1px;height:16px;background:var(--brd);flex-shrink:0; }
  .ll-chip { display:flex;align-items:center;gap:4px;padding:3px 9px;border-radius:5px;
    cursor:pointer;flex-shrink:0;background:transparent;border:1px solid var(--brd);
    color:var(--muted);font-size:9px;font-weight:700;letter-spacing:.1em;
    font-family:inherit;transition:all .1s; }
  .ll-chip:hover { color:var(--txt); }
  .ll-chip-on { background:rgba(245,158,11,.08);border-color:rgba(245,158,11,.5);color:var(--acc); }
  .ll-chip-dot { width:5px;height:5px;border-radius:50%;background:var(--acc);
    flex-shrink:0;box-shadow:0 0 5px var(--acc); }
  .ll-typechip { padding:3px 8px;border-radius:5px;cursor:pointer;flex-shrink:0;
    border:1px solid;font-size:9px;font-weight:700;letter-spacing:.08em;
    font-family:inherit;transition:all .1s; }
  .ll-sortchip { padding:3px 8px;border-radius:5px;cursor:pointer;flex-shrink:0;
    background:transparent;border:1px solid transparent;color:var(--muted);
    font-size:9px;font-weight:600;font-family:inherit;transition:all .1s; }
  .ll-sortchip:hover { color:var(--txt); }
  .ll-sortchip-on { background:rgba(56,189,248,.08);border-color:rgba(56,189,248,.3);color:var(--blue); }

  /* Buttons */
  .ll-btn { display:flex;align-items:center;gap:5px;padding:6px 11px;border-radius:6px;
    cursor:pointer;flex-shrink:0;font-size:10px;font-weight:600;letter-spacing:.07em;
    font-family:inherit;white-space:nowrap;transition:all .15s; }
  .ll-btn-ghost { background:var(--surf2);border:1px solid var(--brd2);color:var(--muted); }
  .ll-btn-ghost:hover { color:var(--txt); }
  .ll-btn-acc { background:var(--acc);border:none;color:#0a0b0e;font-weight:700;
    box-shadow:0 0 12px rgba(245,158,11,.3); }
  .ll-btn-acc:hover { box-shadow:0 0 20px rgba(245,158,11,.45); }

  /* Status bar */
  .ll-status { display:flex;align-items:center;justify-content:space-between;
    padding:3px 14px 5px;border-top:1px solid var(--brd);font-size:9px;letter-spacing:.08em; }
  .ll-status-l,.ll-status-r { display:flex;align-items:center;gap:8px; }
  .ll-status-lbl { font-weight:700; }
  .ll-stat { color:var(--muted); }
  .ll-res { font-weight:700; }
  .ll-time { font-family:inherit; }
  .ll-pct { color:var(--blue);font-weight:700; }
  .ll-sdiv { width:1px;height:10px;background:var(--brd); }
  .ll-dot { display:inline-block;width:6px;height:6px;border-radius:50%;flex-shrink:0; }
  .ll-pbar { width:70px;height:2px;background:var(--brd);border-radius:2px;overflow:hidden; }
  .ll-pbar-fill { height:100%;background:var(--acc);border-radius:2px;transition:width .3s; }

  /* Body */
  .ll-body { flex:1;overflow-y:auto;overflow-x:hidden; }
  .ll-body::-webkit-scrollbar { width:4px; }
  .ll-body::-webkit-scrollbar-track { background:transparent; }
  .ll-body::-webkit-scrollbar-thumb { background:var(--brd2);border-radius:4px; }

  /* Center screens */
  .ll-center { height:100%;display:flex;flex-direction:column;
    align-items:center;justify-content:center;padding:32px; }
  .ll-welcome-icon { width:64px;height:64px;border-radius:16px;background:var(--surf);
    border:1px solid var(--brd2);display:flex;align-items:center;justify-content:center;
    margin-bottom:18px; }
  .ll-welcome-title { font-size:13px;font-weight:700;color:var(--muted);
    letter-spacing:.1em;margin-bottom:10px; }
  .ll-welcome-sub { font-size:11px;color:var(--dim);text-align:center;
    line-height:1.8;max-width:320px;margin-bottom:18px; }
  .ll-badges { display:flex;gap:7px;margin-bottom:18px;flex-wrap:wrap;justify-content:center; }
  .ll-badge { padding:2px 9px;border-radius:20px;border:1px solid;
    font-size:8px;font-weight:700;letter-spacing:.1em; }
  .ll-features { display:flex;flex-direction:column;gap:5px;align-items:flex-start; }
  .ll-feature { font-size:11px;color:var(--dim); }

  /* Indexing */
  .ll-idx-title { font-size:11px;font-weight:700;color:var(--muted);
    letter-spacing:.1em;margin-bottom:5px; }
  .ll-idx-sub { font-size:10px;color:var(--dim);margin-bottom:12px; }
  .ll-idx-bar { width:180px;height:2px;background:var(--brd);border-radius:2px;
    overflow:hidden;margin-bottom:18px; }
  .ll-idx-fill { height:100%;background:var(--acc);border-radius:2px;transition:width .4s; }
  .ll-idx-list { display:flex;flex-direction:column;gap:4px;width:260px; }
  .ll-idx-row { display:flex;align-items:center;gap:7px;padding:4px 8px;
    border-radius:5px;background:var(--surf);border:1px solid var(--brd); }
  .ll-idx-ico { line-height:0;flex-shrink:0; }
  .ll-idx-dot { display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--dim); }
  .ll-idx-name { flex:1;font-size:9px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
  .ll-idx-n { font-size:8px;color:var(--muted);flex-shrink:0; }
  .ll-idx-err { font-size:8px;color:var(--red);font-weight:700;
    background:rgba(248,113,113,.1);padding:1px 4px;border-radius:3px; }
  .ll-idx-more { font-size:9px;color:var(--dim);text-align:center;padding:4px; }

  /* Hits */
  .ll-hit { border-bottom:1px solid var(--brd);transition:background .08s; }
  .ll-hit-open { background:var(--surf); }
  .ll-hit-row { display:flex;align-items:center;height:64px;
    padding:0 14px;cursor:pointer;gap:11px; }
  .ll-hit:not(.ll-hit-open) .ll-hit-row:hover { background:var(--surf); }
  .ll-hit-ico { width:30px;height:30px;border-radius:7px;flex-shrink:0;border:1px solid;
    display:flex;align-items:center;justify-content:center; }
  .ll-hit-body { flex:1;min-width:0; }
  .ll-hit-meta { display:flex;align-items:center;gap:7px;margin-bottom:3px; }
  .ll-hit-fn { font-size:11px;font-weight:600;color:var(--txt);
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:260px; }
  .ll-hit-sn { font-size:8px;font-weight:600;color:var(--blue);
    background:rgba(56,189,248,.08);border:1px solid rgba(56,189,248,.25);
    border-radius:3px;padding:1px 4px;flex-shrink:0; }
  .ll-hit-rn { font-size:9px;color:var(--dim);flex-shrink:0;letter-spacing:.05em; }
  .ll-hit-prev { font-size:11px;color:var(--muted);
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
  .ll-hit-right { display:flex;align-items:center;gap:7px;flex-shrink:0; }
  .ll-hit-sc { font-size:9px;font-weight:700;color:var(--acc);
    background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.25);
    border-radius:4px;padding:1px 5px; }
  .ll-chev { color:var(--dim);transition:transform .15s;flex-shrink:0; }
  .ll-chev-open { transform:rotate(180deg);color:var(--acc); }

  /* Detail */
  .ll-detail { padding:10px 14px 14px;border-top:1px solid var(--brd);
    background:rgba(0,0,0,.18);animation:fadeIn .12s ease; }
  .ll-detail-load,.ll-detail-empty { display:flex;align-items:center;gap:7px;
    font-size:11px;color:var(--muted);padding:7px 0; }
  .ll-detail-grid { display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:7px; }
  .ll-detail-cell { padding:7px 9px;border-radius:5px;background:var(--surf2);border:1px solid var(--brd); }
  .ll-detail-k { font-size:8px;font-weight:600;color:var(--dim);
    letter-spacing:.12em;margin-bottom:3px;text-transform:uppercase; }
  .ll-detail-v { font-size:11px;color:var(--txt);word-break:break-word;line-height:1.4; }

  /* Highlight */
  .ll-hl { background:rgba(245,158,11,.2);color:var(--acc);border-radius:2px;padding:0 1px; }

  ::-webkit-scrollbar { width:4px;height:4px; }
  ::-webkit-scrollbar-track { background:transparent; }
  ::-webkit-scrollbar-thumb { background:var(--brd2);border-radius:4px; }
  button,input { font-family:inherit; }
`;

// Inject CSS once
(() => {
  const s = document.createElement('style');
  s.textContent = _css;
  document.head.appendChild(s);
})();
