import React, { useEffect, useRef, useState, useMemo, useCallback, memo } from 'react';
import {
  Zap, FolderOpen, Search, X, Download,
  Loader2, CheckCircle2, AlertCircle, ChevronDown,
  FileSpreadsheet, FileCode, FileText, ScanSearch
} from 'lucide-react';
import workerScript from './engine.js?raw';

/* ── Types ─────────────────────────────────────────────────────────────── */
interface FM {
  id:string; name:string; type:string; blob:File; size:number;
  st:'q'|'ing'|'ok'|'err'; rows:number; err?:string;
}
interface Hit {
  id:string; fid:string; fn:string; sn:string|null;
  rn:number; ss:string; sc:number; ri:number;
}

/* ── Constants ──────────────────────────────────────────────────────────── */
const NW      = 2;
const ROW_H   = 64;
const OVER    = 10;
const DEB     = 200;
const MAX_H   = 20000;
const SUPP    = new Set(['xlsx','xls','csv','txt']);

const fmtN = (n:number) => n>=1e6?(n/1e6).toFixed(1)+'M':n>=1e3?(n/1e3).toFixed(1)+'K':''+n;
const fmtB = (b:number) => b>=(1<<20)?(b/(1<<20)).toFixed(1)+' MB':b>=(1<<10)?(b/(1<<10)).toFixed(1)+' KB':b+' B';
const fmtT = (ms:number) => ms>=1000?(ms/1000).toFixed(2)+'s':ms+'ms';

type Ext='xlsx'|'xls'|'csv'|'txt';
const ES:Record<Ext,{c:string;dim:string;I:any}> = {
  xlsx:{c:'#38bdf8',dim:'rgba(56,189,248,0.08)',I:FileSpreadsheet},
  xls: {c:'#818cf8',dim:'rgba(129,140,248,0.08)',I:FileSpreadsheet},
  csv: {c:'#4ade80',dim:'rgba(74,222,128,0.08)',I:FileCode},
  txt: {c:'#fb923c',dim:'rgba(251,146,60,0.08)',I:FileText},
};
const eOf = (n:string) => (n.split('.').pop()?.toLowerCase()||'xlsx') as Ext;

/* ── Design tokens ──────────────────────────────────────────────────────── */
const T = {
  bg:      '#0a0b0e',
  surface: '#0f1117',
  border:  '#1c2030',
  border2: '#252b3b',
  text:    '#e2e8f0',
  muted:   '#4a5568',
  dim:     '#2d3748',
  accent:  '#f59e0b',
  blue:    '#38bdf8',
  green:   '#4ade80',
  red:     '#f87171',
};

/* ── App ────────────────────────────────────────────────────────────────── */
export default function App() {
  const [files,   setFiles]   = useState<FM[]>([]);
  const [totR,    setTotR]    = useState(0);
  const [query,   setQuery]   = useState('');
  const [hits,    setHits]    = useState<Hit[]>([]);
  const [busy,    setBusy]    = useState(false);
  const [elapsed, setElapsed] = useState<number|null>(null);
  const [fuzzy,   setFuzzy]   = useState(true);
  const [exact,   setExact]   = useState(false);
  const [exts,    setExts]    = useState(['xlsx','xls','csv','txt']);
  const [sortK,   setSortK]   = useState<'sc'|'fn'|'rn'>('sc');
  const [sortD,   setSortD]   = useState<1|-1>(-1);
  const [expId,   setExpId]   = useState<string|null>(null);
  const [expD,    setExpD]    = useState<Record<string,string>|null>(null);
  const [expLoad, setExpLoad] = useState(false);
  const [sTop,    setSTop]    = useState(0);
  const [vH,      setVH]      = useState(600);
  const scEl = useRef<HTMLDivElement>(null);

  const WW       = useRef<Worker[]>([]);
  const indexed  = useRef<Map<string,FM>>(new Map());
  const iQ       = useRef<FM[]>([]);
  const sid      = useRef(0);
  const sDone    = useRef(0);
  const sTotal   = useRef(0);
  const sStart   = useRef(0);
  const sBuf     = useRef<Hit[]>([]);
  const allHits  = useRef<Hit[]>([]);
  const debT     = useRef(0);
  const tickT    = useRef(0);
  const queryRef = useRef('');
  const fuzzyRef = useRef(true);
  const exactRef = useRef(false);
  const extsRef  = useRef(['xlsx','xls','csv','txt']);

  useEffect(() => { fuzzyRef.current = fuzzy; }, [fuzzy]);
  useEffect(() => { exactRef.current = exact; }, [exact]);
  useEffect(() => { extsRef.current  = exts;  }, [exts]);

  useEffect(() => {
    const ro = new ResizeObserver(() => {
      if (scEl.current) setVH(scEl.current.clientHeight);
    });
    if (scEl.current) ro.observe(scEl.current);
    return () => ro.disconnect();
  }, []);

  useEffect(() => {
    const blob = new Blob([workerScript], {type:'application/javascript'});
    const url  = URL.createObjectURL(blob);

    function flushHits() {
      if (!sBuf.current.length) return;
      const snap = sBuf.current.splice(0);
      allHits.current.push(...snap);
      setHits(allHits.current.slice(0, MAX_H));
    }

    function dispatchIdx(_wi:number) {
      // Always use worker 0 for indexing — serialised to prevent OOM
      if (!iQ.current.length) return;
      const f = iQ.current.shift()!;
      setFiles(prev => prev.map(x => x.id===f.id ? {...x,st:'ing' as const} : x));
      WW.current[0].postMessage({t:'idx', id:f.id, blob:f.blob, ft:f.type, name:f.name});
    }

    WW.current = Array.from({length:NW}, (_, wi) => {
      const w = new Worker(url);
      w.onmessage = (ev) => {
        const d = ev.data;
        if (d.t === 'ok') {
          setFiles(prev => {
            const nx = prev.map(x => x.id===d.id ? {...x, st:'ok' as const, rows:d.n} : x);
            const entry = nx.find(x => x.id===d.id);
            if (entry) indexed.current.set(d.id, {...entry, rows:d.n, st:'ok'});
            return nx;
          });
          setTotR(r => r + d.n);
          dispatchIdx(wi);
        } else if (d.t === 'err') {
          setFiles(prev => prev.map(x => x.id===d.id ? {...x, st:'err' as const, err:d.msg} : x));
          dispatchIdx(wi);
        } else if (d.t === 'hits') {
          if (d.sid !== sid.current) return;
          sBuf.current.push(...d.hits);
          if (sBuf.current.length > 300) flushHits();
        } else if (d.t === 'done') {
          if (d.sid !== sid.current) return;
          sDone.current++;
          if (sDone.current >= sTotal.current) {
            flushHits();
            setBusy(false);
            clearInterval(tickT.current);
            setElapsed(Date.now() - sStart.current);
          }
        } else if (d.t === 'row') {
          setExpD(d.d); setExpLoad(false);
        } else if (d.t === 'csv') {
          const a = document.createElement('a');
          a.href = URL.createObjectURL(d.blob);
          a.download = 'locallens_' + Date.now() + '.csv';
          a.click();
        }
      };
      return w;
    });

    return () => { WW.current.forEach(w => w.terminate()); URL.revokeObjectURL(url); };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const loadFolder = useCallback((e:React.ChangeEvent<HTMLInputElement>) => {
    const raw = e.target.files;
    if (!raw) return;
    WW.current.forEach(w => w.postMessage({t:'clear'}));
    indexed.current.clear(); iQ.current = [];
    allHits.current = []; sBuf.current = [];
    setHits([]); setQuery(''); queryRef.current = '';
    setElapsed(null); setTotR(0); setExpId(null); setExpD(null);
    const list:FM[] = [];
    for (let i = 0; i < raw.length; i++) {
      const f = raw[i];
      const ext = (f.name.split('.').pop()||'').toLowerCase();
      if (!SUPP.has(ext)) continue;
      list.push({ id:`${f.name}-${f.size}-${f.lastModified}`,
        name:f.name, type:ext, blob:f, size:f.size, st:'q', rows:0 });
    }
    if (!list.length) return;
    setFiles(list); iQ.current = list.slice();
    // Only seed worker 0 for indexing — one file at a time prevents OOM on large folders
    if (iQ.current.length) {
      const f = iQ.current.shift()!;
      setFiles(prev => prev.map(x => x.id===f.id ? {...x, st:'ing' as const} : x));
      WW.current[0].postMessage({t:'idx', id:f.id, blob:f.blob, ft:f.type, name:f.name});
    }
  }, []);

  const runSearch = useCallback((q:string) => {
    if (!q.trim()) { setHits([]); allHits.current = []; setElapsed(null); return; }
    const targets = Array.from(indexed.current.values()).filter(f => extsRef.current.includes(f.type));
    if (!targets.length) return;
    const newSid = ++sid.current;
    sBuf.current = []; allHits.current = [];
    sDone.current = 0; sTotal.current = targets.length;
    sStart.current = Date.now();
    setHits([]); setElapsed(null); setBusy(true); setExpId(null); setExpD(null);
    clearInterval(tickT.current);
    tickT.current = setInterval(() => setElapsed(Date.now()-sStart.current), 100) as unknown as number;
    targets.forEach((f, i) => {
      WW.current[i % NW].postMessage({t:'search', sid:newSid, id:f.id, q, fuzzy:fuzzyRef.current, exact:exactRef.current});
    });
  }, []);

  const trig = useCallback((q:string) => {
    clearTimeout(debT.current);
    queryRef.current = q;
    debT.current = setTimeout(() => runSearch(q), DEB) as unknown as number;
  }, [runSearch]);

  useEffect(() => {
    if (queryRef.current.trim()) runSearch(queryRef.current);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fuzzy, exact, exts]);

  const expand = useCallback((h:Hit) => {
    if (expId === h.id) { setExpId(null); setExpD(null); return; }
    setExpId(h.id); setExpD(null); setExpLoad(true);
    WW.current[0].postMessage({t:'row', id:h.fid, ri:h.ri});
  }, [expId]);

  const sorted = useMemo(() => {
    const a = hits.slice();
    if      (sortK==='sc') a.sort((x,y) => sortD*(y.sc-x.sc));
    else if (sortK==='fn') a.sort((x,y) => sortD*x.fn.localeCompare(y.fn));
    else                   a.sort((x,y) => sortD*(x.rn-y.rn));
    return a;
  }, [hits, sortK, sortD]);

  const {vis, vs} = useMemo(() => {
    const s = Math.max(0, Math.floor(sTop/ROW_H) - OVER);
    const e = Math.min(sorted.length, Math.ceil((sTop+vH)/ROW_H) + OVER);
    return {vis:sorted.slice(s,e), vs:s};
  }, [sorted, sTop, vH]);

  const nOk   = files.filter(f=>f.st==='ok').length;
  const isIdx = files.some(f=>f.st==='ing'||f.st==='q');
  const ready = !isIdx && nOk > 0;
  const totSz = files.reduce((s,f)=>s+f.size,0);
  const pct   = files.length ? Math.round(nOk/files.length*100) : 0;
  const capped = allHits.current.length >= MAX_H;

  return (
    <div className="ll-root">
      {/* ══ TOPBAR ════════════════════════════════════════════════════════ */}
      <header className="ll-header">
        <div className="ll-header-main">

          {/* Logo */}
          <div className="ll-logo">
            <div className="ll-logo-icon"><Zap size={15} color="#0a0b0e"/></div>
            <span className="ll-logo-text">LOCAL<span className="ll-logo-accent">LENS</span></span>
          </div>

          {/* Search */}
          <div className="ll-search-wrap">
            <div className="ll-search-icon">
              {busy
                ? <Loader2 size={14} className="ll-spin" style={{color:T.accent}}/>
                : <Search size={14} style={{color: query ? T.accent : T.muted}}/>}
            </div>
            <input
              className="ll-search-input"
              value={query}
              disabled={!ready}
              onChange={e => { setQuery(e.target.value); trig(e.target.value); }}
              placeholder={
                isIdx ? `Indexing ${nOk} / ${files.length} files…` :
                !files.length ? 'Open a folder to start searching' :
                !ready ? 'Building index…' :
                `Search across ${fmtN(totR)} rows in ${files.length} files`
              }
              autoComplete="off"
              spellCheck={false}
            />
            {query && (
              <button className="ll-search-clear" onClick={() => { setQuery(''); runSearch(''); }}>
                <X size={13}/>
              </button>
            )}
          </div>

          {/* Mode toggles */}
          <div className="ll-pill-group">
            <ModeBtn active={fuzzy}  onClick={() => setFuzzy(v=>!v)}  label="FUZZY"  title="Match all words anywhere in row"/>
            <ModeBtn active={exact}  onClick={() => setExact(v=>!v)}  label="EXACT"  title="Exact cell match"/>
          </div>

          <div className="ll-sep"/>

          {/* File type filters */}
          <div className="ll-pill-group">
            {(['xlsx','xls','csv','txt'] as Ext[]).map(t => (
              <TypeBtn key={t} t={t} active={exts.includes(t)}
                onClick={() => setExts(p => p.includes(t) ? p.filter(x=>x!==t) : [...p,t])}/>
            ))}
          </div>

          <div className="ll-sep"/>

          {/* Sort */}
          <div className="ll-pill-group">
            {(['sc','fn','rn'] as const).map((k,i) => (
              <SortBtn key={k}
                label={['SCORE','FILE','ROW'][i]}
                active={sortK===k} dir={sortD}
                onClick={() => {
                  if (sortK===k) setSortD(d => d===1?-1:1);
                  else { setSortK(k); setSortD(-1); }
                }}/>
            ))}
          </div>

          <div className="ll-sep"/>

          {/* Folder + Export */}
          <label className="ll-btn ll-btn-ghost">
            <FolderOpen size={13}/>
            {files.length ? 'CHANGE' : 'OPEN FOLDER'}
            <input type="file" style={{display:'none'}}
              // @ts-ignore
              webkitdirectory="" directory="" multiple onChange={loadFolder}/>
          </label>

          {allHits.current.length > 0 && (
            <button className="ll-btn ll-btn-accent"
              onClick={() => WW.current[0].postMessage({t:'export', hits:allHits.current})}>
              <Download size={13}/> EXPORT
            </button>
          )}
        </div>

        {/* Status bar */}
        <div className="ll-statusbar">
          <div className="ll-status-left">
            {files.length > 0 && (
              <>
                <StatusDot ok={!isIdx}/>
                <span className="ll-status-label" style={{color: isIdx ? T.accent : T.green}}>
                  {isIdx ? `INDEXING ${nOk}/${files.length}` : `${nOk} FILES INDEXED`}
                </span>
                <span className="ll-status-divider"/>
                <span className="ll-status-stat">{fmtN(totR)} ROWS</span>
                <span className="ll-status-divider"/>
                <span className="ll-status-stat">{fmtB(totSz)}</span>
                {isIdx && (
                  <>
                    <span className="ll-status-divider"/>
                    <div className="ll-progbar">
                      <div className="ll-progbar-fill" style={{width:`${pct}%`}}/>
                    </div>
                    <span className="ll-status-pct">{pct}%</span>
                  </>
                )}
              </>
            )}
          </div>
          <div className="ll-status-right">
            {hits.length > 0 && (
              <span className="ll-status-results" style={{color: capped ? T.accent : T.blue}}>
                {capped ? `${fmtN(MAX_H)}+ RESULTS` : `${fmtN(hits.length)} RESULTS`}
              </span>
            )}
            {elapsed !== null && (
              <span className="ll-status-time" style={{color: busy ? T.accent : T.muted}}>
                {fmtT(elapsed)}{busy ? '…' : ''}
              </span>
            )}
          </div>
        </div>
      </header>

      {/* ══ BODY ════════════════════════════════════════════════════════ */}
      <div className="ll-body" ref={scEl}
        onScroll={e => setSTop(e.currentTarget.scrollTop)}>

        {files.length === 0 && <WelcomeScreen/>}
        {files.length > 0 && isIdx && hits.length === 0 && !query && <IndexingScreen files={files}/>}
        {ready && !query && hits.length === 0 && <ReadyScreen n={files.length} rows={totR}/>}
        {ready && query && hits.length === 0 && !busy && <NoResults q={query}/>}

        {sorted.length > 0 && (
          <div style={{position:'relative', height: sorted.length * ROW_H}}>
            <div style={{position:'absolute', top:0, left:0, right:0,
              transform: `translateY(${vs * ROW_H}px)`}}>
              {vis.map(h => (
                <HitRow key={h.id} h={h}
                  open={expId===h.id}
                  data={expId===h.id ? expD : undefined}
                  loading={expId===h.id && expLoad}
                  onOpen={expand}
                  q={query}/>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ── Control buttons ────────────────────────────────────────────────────── */
const ModeBtn = memo(({active,onClick,label,title}:{active:boolean;onClick:()=>void;label:string;title?:string}) => (
  <button className={`ll-modebtn ${active?'ll-modebtn-on':''}`} onClick={onClick} title={title}>
    {active && <span className="ll-modebtn-dot"/>}
    {label}
  </button>
));

const TypeBtn = memo(({t,active,onClick}:{t:Ext;active:boolean;onClick:()=>void}) => {
  const e = ES[t];
  return (
    <button className="ll-typebtn" onClick={onClick}
      style={{
        color: active ? e.c : T.muted,
        background: active ? e.dim : 'transparent',
        borderColor: active ? e.c+'40' : T.border,
      }}>
      {t.toUpperCase()}
    </button>
  );
});

const SortBtn = memo(({label,active,dir,onClick}:{label:string;active:boolean;dir:1|-1;onClick:()=>void}) => (
  <button className={`ll-sortbtn ${active?'ll-sortbtn-on':''}`} onClick={onClick}>
    {label}{active ? (dir===-1?' ↓':' ↑') : ''}
  </button>
));

const StatusDot = ({ok}:{ok:boolean}) => (
  <span style={{
    display:'inline-block', width:6, height:6, borderRadius:'50%',
    background: ok ? T.green : T.accent,
    boxShadow: ok ? `0 0 6px ${T.green}` : `0 0 6px ${T.accent}`,
    flexShrink:0
  }}/>
);

/* ── Empty states ───────────────────────────────────────────────────────── */
const WelcomeScreen = memo(() => (
  <div className="ll-center">
    <div className="ll-welcome-icon">
      <ScanSearch size={32} strokeWidth={1.5} style={{color:T.muted}}/>
    </div>
    <div className="ll-welcome-title">OPEN A FOLDER TO BEGIN</div>
    <div className="ll-welcome-sub">
      Drop a folder of XLSX, XLS, CSV or TXT files.<br/>
      Each file is indexed <em>once</em> — every search runs at native JS speed.
    </div>
    <div className="ll-welcome-badges">
      {(['XLSX','XLS','CSV','TXT'] as const).map((x,i) => (
        <span key={x} className="ll-badge"
          style={{color:Object.values(ES)[i].c, borderColor:Object.values(ES)[i].c+'30',
            background:Object.values(ES)[i].dim}}>
          {x}
        </span>
      ))}
    </div>
    <div className="ll-welcome-features">
      {['⚡ Parse once, search instantly','🔒 100% local — no uploads','🔍 Fuzzy + exact modes','📤 Export matches to CSV'].map(f => (
        <div key={f} className="ll-feature-row">{f}</div>
      ))}
    </div>
  </div>
));

const IndexingScreen = memo(({files}:{files:FM[]}) => {
  const done = files.filter(f=>f.st==='ok').length;
  const pct  = files.length ? Math.round(done/files.length*100) : 0;
  return (
    <div className="ll-center">
      <div className="ll-idx-ring">
        <Loader2 size={28} className="ll-spin" style={{color:T.accent}}/>
      </div>
      <div className="ll-idx-title">BUILDING SEARCH INDEX</div>
      <div className="ll-idx-sub">{done} of {files.length} files · {pct}%</div>
      <div className="ll-idx-bar">
        <div className="ll-idx-bar-fill" style={{width:`${pct}%`}}/>
      </div>
      <div className="ll-idx-list">
        {files.slice(0,12).map(f => (
          <div key={f.id} className="ll-idx-row">
            <span className="ll-idx-icon">
              {f.st==='ok'  ? <CheckCircle2 size={10} style={{color:T.green}}/>
              :f.st==='err' ? <AlertCircle  size={10} style={{color:T.red}}/>
              :f.st==='ing' ? <Loader2 size={10} className="ll-spin" style={{color:T.blue}}/>
              : <span style={{display:'inline-block',width:10,height:10,borderRadius:'50%',background:T.dim}}/>}
            </span>
            <span className="ll-idx-name" style={{
              color: f.st==='ok' ? T.text : f.st==='err' ? T.red : T.muted
            }}>{f.name}</span>
            {f.st==='ok' && <span className="ll-idx-count">{fmtN(f.rows)}</span>}
            {f.st==='err' && <span className="ll-idx-err">ERR</span>}
          </div>
        ))}
        {files.length > 12 && (
          <div className="ll-idx-more">+{files.length-12} more files queued</div>
        )}
      </div>
    </div>
  );
});

const ReadyScreen = memo(({n,rows}:{n:number;rows:number}) => (
  <div className="ll-center" style={{opacity:0.4}}>
    <Search size={36} strokeWidth={1} style={{color:T.accent}}/>
    <div style={{fontSize:11,color:T.muted,letterSpacing:'0.12em',marginTop:12}}>
      {n} FILES · {fmtN(rows)} ROWS READY
    </div>
    <div style={{fontSize:10,color:T.dim,letterSpacing:'0.08em',marginTop:4}}>
      TYPE ANYTHING TO SEARCH
    </div>
  </div>
));

const NoResults = memo(({q}:{q:string}) => (
  <div className="ll-center" style={{opacity:0.5}}>
    <Search size={32} strokeWidth={1} style={{color:T.muted}}/>
    <div style={{fontSize:12,color:T.muted,marginTop:12}}>
      No matches for <span style={{color:T.text}}>"{q}"</span>
    </div>
    <div style={{fontSize:10,color:T.dim,marginTop:4}}>
      Try fuzzy mode or check file type filters
    </div>
  </div>
));

/* ── Highlight ──────────────────────────────────────────────────────────── */
const Hl = memo(({text,q}:{text:string;q:string}) => {
  if (!q.trim()||!text) return <>{text}</>;
  const terms = q.trim().split(/\s+/).filter(t=>t.length>1);
  if (!terms.length) return <>{text}</>;
  try {
    const re = new RegExp(`(${terms.map(t=>t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|')})`, 'gi');
    return <>{text.split(re).map((p,i) =>
      re.test(p) ? <mark key={i} className="ll-hl">{p}</mark> : p
    )}</>;
  } catch { return <>{text}</>; }
});

/* ── Hit row ────────────────────────────────────────────────────────────── */
const HitRow = memo(({h,open,data,loading,onOpen,q}:{
  h:Hit; open:boolean; data?:Record<string,string>|null;
  loading?:boolean; onOpen:(h:Hit)=>void; q:string;
}) => {
  const es = ES[eOf(h.fn)] || ES.xlsx;
  return (
    <div className={`ll-hit ${open?'ll-hit-open':''}`}>
      {/* Summary */}
      <div className="ll-hit-row" onClick={() => onOpen(h)}>
        <div className="ll-hit-icon" style={{background:es.dim, borderColor:es.c+'30', color:es.c}}>
          <es.I size={13} strokeWidth={1.5}/>
        </div>

        <div className="ll-hit-body">
          <div className="ll-hit-meta">
            <span className="ll-hit-filename">{h.fn}</span>
            {h.sn && <span className="ll-hit-sheet">{h.sn}</span>}
            <span className="ll-hit-row-num">ROW {h.rn}</span>
          </div>
          <div className="ll-hit-preview">
            <Hl text={h.ss} q={q}/>
          </div>
        </div>

        <div className="ll-hit-right">
          {h.sc > 1 && <span className="ll-hit-score">{h.sc}</span>}
          <ChevronDown size={13} className={`ll-chevron ${open?'ll-chevron-open':''}`}/>
        </div>
      </div>

      {/* Detail panel */}
      {open && (
        <div className="ll-hit-detail">
          {loading ? (
            <div className="ll-detail-loading">
              <Loader2 size={12} className="ll-spin"/>
              <span>Loading row data…</span>
            </div>
          ) : !data || !Object.keys(data).length ? (
            <div className="ll-detail-empty">
              <AlertCircle size={12}/> No data available
            </div>
          ) : (
            <div className="ll-detail-grid">
              {Object.entries(data).map(([k,v]) => (
                <div key={k} className="ll-detail-cell">
                  <div className="ll-detail-key">{k}</div>
                  <div className="ll-detail-val">
                    <Hl text={String(v||'—')} q={q}/>
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

/* ── CSS ─────────────────────────────────────────────────────────────────── */
{
  const css = `
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=IBM+Plex+Sans:wght@400;500;600&display=swap');

    @keyframes spin { to { transform:rotate(360deg) } }
    @keyframes fadeIn { from { opacity:0; transform:translateY(4px) } to { opacity:1; transform:none } }
    @keyframes pulse-dot { 0%,100% { opacity:1 } 50% { opacity:0.3 } }

    *, *::before, *::after { box-sizing:border-box; margin:0; padding:0; }
    :root {
      --bg:      #0a0b0e;
      --surface: #0f1117;
      --surface2:#141720;
      --border:  #1c2030;
      --border2: #252b3b;
      --text:    #e2e8f0;
      --muted:   #4a5568;
      --dim:     #2d3748;
      --accent:  #f59e0b;
      --blue:    #38bdf8;
      --green:   #4ade80;
      --red:     #f87171;
    }

    html, body, #root { height:100%; overflow:hidden; }
    body { background:var(--bg); color:var(--text); }

    .ll-root {
      display:flex; flex-direction:column; height:100vh; overflow:hidden;
      font-family:'IBM Plex Mono', 'Fira Code', monospace;
      background:var(--bg);
    }

    /* ── HEADER ── */
    .ll-header {
      flex-shrink:0;
      background:var(--surface);
      border-bottom:1px solid var(--border);
    }

    .ll-header-main {
      display:flex; align-items:center; gap:8px;
      padding:10px 16px; flex-wrap:wrap;
    }

    /* Logo */
    .ll-logo { display:flex; align-items:center; gap:9px; flex-shrink:0; margin-right:4px; }
    .ll-logo-icon {
      width:30px; height:30px; border-radius:8px; flex-shrink:0;
      background:var(--accent);
      display:flex; align-items:center; justify-content:center;
      box-shadow:0 0 16px rgba(245,158,11,0.4);
    }
    .ll-logo-text {
      font-size:14px; font-weight:600; letter-spacing:0.08em; color:var(--text);
      user-select:none; white-space:nowrap;
    }
    .ll-logo-accent { color:var(--accent); }

    /* Search */
    .ll-search-wrap {
      flex:1; min-width:180px; position:relative;
      display:flex; align-items:center;
    }
    .ll-search-icon {
      position:absolute; left:12px; top:50%; transform:translateY(-50%);
      pointer-events:none; line-height:0;
    }
    .ll-search-input {
      width:100%; padding:9px 36px 9px 36px;
      background:var(--surface2);
      border:1px solid var(--border2);
      border-radius:8px;
      color:var(--text); font-size:13px; font-family:inherit;
      outline:none; transition:border-color 0.15s, box-shadow 0.15s;
    }
    .ll-search-input:focus {
      border-color:var(--accent);
      box-shadow:0 0 0 3px rgba(245,158,11,0.1);
    }
    .ll-search-input::placeholder { color:var(--muted); }
    .ll-search-input:disabled { opacity:0.35; cursor:not-allowed; }
    .ll-search-clear {
      position:absolute; right:10px; top:50%; transform:translateY(-50%);
      background:none; border:none; cursor:pointer; color:var(--muted);
      line-height:0; padding:2px; border-radius:3px; transition:color 0.1s;
    }
    .ll-search-clear:hover { color:var(--text); }

    /* Separators */
    .ll-sep { width:1px; height:18px; background:var(--border); flex-shrink:0; }
    .ll-pill-group { display:flex; align-items:center; gap:4px; }

    /* Mode buttons */
    .ll-modebtn {
      display:flex; align-items:center; gap:5px;
      padding:4px 10px; border-radius:6px; cursor:pointer; flex-shrink:0;
      background:transparent; border:1px solid var(--border);
      color:var(--muted); font-size:9px; font-weight:700;
      letter-spacing:0.1em; font-family:inherit;
      transition:all 0.12s;
    }
    .ll-modebtn:hover { border-color:var(--border2); color:var(--text); }
    .ll-modebtn-on {
      background:rgba(245,158,11,0.08);
      border-color:rgba(245,158,11,0.5);
      color:var(--accent);
    }
    .ll-modebtn-dot {
      width:5px; height:5px; border-radius:50%;
      background:var(--accent); flex-shrink:0;
      box-shadow:0 0 5px var(--accent);
    }

    /* Type buttons */
    .ll-typebtn {
      padding:4px 8px; border-radius:5px; cursor:pointer; flex-shrink:0;
      border:1px solid; font-size:9px; font-weight:700;
      letter-spacing:0.08em; font-family:inherit;
      transition:all 0.12s;
    }

    /* Sort buttons */
    .ll-sortbtn {
      padding:4px 8px; border-radius:5px; cursor:pointer; flex-shrink:0;
      background:transparent; border:1px solid transparent;
      color:var(--muted); font-size:9px; font-weight:600;
      font-family:inherit; transition:all 0.12s; letter-spacing:0.06em;
    }
    .ll-sortbtn:hover { color:var(--text); }
    .ll-sortbtn-on {
      background:rgba(56,189,248,0.08);
      border-color:rgba(56,189,248,0.3);
      color:var(--blue);
    }

    /* Buttons */
    .ll-btn {
      display:flex; align-items:center; gap:6px;
      padding:6px 12px; border-radius:7px; cursor:pointer; flex-shrink:0;
      font-size:10px; font-weight:600; letter-spacing:0.07em;
      font-family:inherit; white-space:nowrap; transition:all 0.15s;
    }
    .ll-btn-ghost {
      background:var(--surface2); border:1px solid var(--border2);
      color:var(--muted);
    }
    .ll-btn-ghost:hover { border-color:var(--border2); color:var(--text); }
    .ll-btn-accent {
      background:var(--accent); border:none; color:#0a0b0e; font-weight:700;
      box-shadow:0 0 12px rgba(245,158,11,0.25);
    }
    .ll-btn-accent:hover { box-shadow:0 0 20px rgba(245,158,11,0.4); }

    /* Status bar */
    .ll-statusbar {
      display:flex; align-items:center; justify-content:space-between;
      padding:4px 16px 6px;
      border-top:1px solid var(--border);
      font-size:9px; letter-spacing:0.08em;
    }
    .ll-status-left { display:flex; align-items:center; gap:10px; }
    .ll-status-right { display:flex; align-items:center; gap:12px; }
    .ll-status-label { font-weight:700; }
    .ll-status-divider { width:1px; height:10px; background:var(--border); }
    .ll-status-stat { color:var(--muted); }
    .ll-status-results { font-weight:700; }
    .ll-status-time { font-family:'IBM Plex Mono', monospace; }
    .ll-status-pct { color:var(--blue); font-weight:700; }
    .ll-progbar {
      width:80px; height:2px; background:var(--border); border-radius:2px; overflow:hidden;
    }
    .ll-progbar-fill {
      height:100%; background:var(--accent); border-radius:2px; transition:width 0.3s;
    }

    /* ── BODY ── */
    .ll-body {
      flex:1; overflow-y:auto; overflow-x:hidden;
    }
    .ll-body::-webkit-scrollbar { width:4px; }
    .ll-body::-webkit-scrollbar-track { background:transparent; }
    .ll-body::-webkit-scrollbar-thumb { background:var(--border2); border-radius:4px; }

    /* Center layouts */
    .ll-center {
      height:100%; display:flex; flex-direction:column;
      align-items:center; justify-content:center;
      gap:0; padding:40px;
    }

    /* Welcome */
    .ll-welcome-icon {
      width:72px; height:72px; border-radius:20px;
      background:var(--surface); border:1px solid var(--border2);
      display:flex; align-items:center; justify-content:center;
      margin-bottom:20px;
    }
    .ll-welcome-title {
      font-size:14px; font-weight:600; color:var(--muted);
      letter-spacing:0.1em; margin-bottom:10px;
    }
    .ll-welcome-sub {
      font-family:'IBM Plex Sans', sans-serif;
      font-size:12px; color:var(--dim); text-align:center;
      line-height:1.8; max-width:340px; margin-bottom:20px;
    }
    .ll-welcome-sub em { color:var(--accent); font-style:normal; }
    .ll-welcome-badges {
      display:flex; gap:8px; margin-bottom:20px;
    }
    .ll-badge {
      padding:3px 10px; border-radius:20px; border:1px solid;
      font-size:9px; font-weight:700; letter-spacing:0.1em;
    }
    .ll-welcome-features {
      display:flex; flex-direction:column; gap:6px; align-items:flex-start;
    }
    .ll-feature-row {
      font-family:'IBM Plex Sans', sans-serif;
      font-size:11px; color:var(--dim);
    }

    /* Indexing */
    .ll-idx-ring { margin-bottom:16px; }
    .ll-idx-title {
      font-size:11px; font-weight:700; color:var(--muted);
      letter-spacing:0.1em; margin-bottom:6px;
    }
    .ll-idx-sub { font-size:10px; color:var(--dim); margin-bottom:14px; }
    .ll-idx-bar {
      width:200px; height:2px; background:var(--border); border-radius:2px;
      overflow:hidden; margin-bottom:20px;
    }
    .ll-idx-bar-fill {
      height:100%; background:var(--accent); border-radius:2px; transition:width 0.4s;
    }
    .ll-idx-list { display:flex; flex-direction:column; gap:5px; width:280px; }
    .ll-idx-row {
      display:flex; align-items:center; gap:8px;
      padding:4px 8px; border-radius:5px;
      background:var(--surface); border:1px solid var(--border);
    }
    .ll-idx-icon { line-height:0; flex-shrink:0; }
    .ll-idx-name {
      flex:1; font-size:10px; overflow:hidden;
      text-overflow:ellipsis; white-space:nowrap;
    }
    .ll-idx-count { font-size:9px; color:var(--muted); flex-shrink:0; }
    .ll-idx-err {
      font-size:8px; color:var(--red); font-weight:700;
      background:rgba(248,113,113,0.1); padding:1px 5px; border-radius:3px;
    }
    .ll-idx-more { font-size:9px; color:var(--dim); text-align:center; padding:4px; }

    /* ── HITS ── */
    .ll-hit {
      border-bottom:1px solid var(--border);
      transition:background 0.08s;
    }
    .ll-hit-open { background:var(--surface); }
    .ll-hit-row {
      display:flex; align-items:center;
      height:64px; padding:0 16px; cursor:pointer; gap:12px;
    }
    .ll-hit:not(.ll-hit-open) .ll-hit-row:hover {
      background:var(--surface);
    }
    .ll-hit-icon {
      width:32px; height:32px; border-radius:8px; flex-shrink:0;
      border:1px solid; display:flex; align-items:center; justify-content:center;
    }
    .ll-hit-body { flex:1; min-width:0; }
    .ll-hit-meta {
      display:flex; align-items:center; gap:8px; margin-bottom:4px;
    }
    .ll-hit-filename {
      font-size:11px; font-weight:600; color:var(--text);
      overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:260px;
    }
    .ll-hit-sheet {
      font-size:8px; font-weight:600; color:var(--blue);
      background:rgba(56,189,248,0.08); border:1px solid rgba(56,189,248,0.2);
      border-radius:3px; padding:1px 5px; flex-shrink:0; letter-spacing:0.05em;
    }
    .ll-hit-row-num {
      font-size:9px; color:var(--dim); flex-shrink:0; letter-spacing:0.06em;
    }
    .ll-hit-preview {
      font-size:11px; color:var(--muted);
      overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
      font-family:'IBM Plex Sans', sans-serif;
    }
    .ll-hit-right {
      display:flex; align-items:center; gap:8px; flex-shrink:0;
    }
    .ll-hit-score {
      font-size:9px; font-weight:700; color:var(--accent);
      background:rgba(245,158,11,0.08); border:1px solid rgba(245,158,11,0.2);
      border-radius:4px; padding:1px 6px;
    }
    .ll-chevron {
      color:var(--dim); transition:transform 0.15s; flex-shrink:0;
    }
    .ll-chevron-open { transform:rotate(180deg); color:var(--accent); }

    /* Detail panel */
    .ll-hit-detail {
      padding:12px 16px 16px;
      border-top:1px solid var(--border);
      background:rgba(0,0,0,0.2);
      animation:fadeIn 0.15s ease;
    }
    .ll-detail-loading, .ll-detail-empty {
      display:flex; align-items:center; gap:8px;
      font-size:11px; color:var(--muted); padding:8px 0;
      font-family:'IBM Plex Sans', sans-serif;
    }
    .ll-detail-grid {
      display:grid; grid-template-columns:repeat(auto-fill, minmax(160px,1fr));
      gap:8px;
    }
    .ll-detail-cell {
      padding:8px 10px; border-radius:6px;
      background:var(--surface2); border:1px solid var(--border);
    }
    .ll-detail-key {
      font-size:8px; font-weight:600; color:var(--dim);
      letter-spacing:0.12em; margin-bottom:4px; text-transform:uppercase;
    }
    .ll-detail-val {
      font-size:11px; color:var(--text); word-break:break-word; line-height:1.4;
      font-family:'IBM Plex Sans', sans-serif;
    }

    /* Highlight */
    .ll-hl {
      background:rgba(245,158,11,0.2); color:var(--accent);
      border-radius:2px; padding:0 2px;
    }

    /* Spin */
    .ll-spin { animation:spin 0.8s linear infinite; }

    ::-webkit-scrollbar { width:4px; height:4px; }
    ::-webkit-scrollbar-track { background:transparent; }
    ::-webkit-scrollbar-thumb { background:var(--border2); border-radius:4px; }
  `;
  const s = document.createElement('style');
  s.textContent = css;
  document.head.appendChild(s);
}
