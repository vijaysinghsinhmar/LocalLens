import React, { useEffect, useRef, useState, useMemo, memo, startTransition } from 'react';
import {
  Zap, FolderOpen, Search, X, Download, Loader2,
  CheckCircle2, AlertCircle, ChevronDown,
  FileSpreadsheet, FileCode, FileText, ScanSearch
} from 'lucide-react';
// Vite worker import — bundled at build time, no CDN, no importScripts
import SearchWorker from './worker.ts?worker&inline';

// ── Types ─────────────────────────────────────────────────────────────────────
interface FM {
  id: string; name: string; type: string; blob: File; size: number;
  st: 'q'|'ing'|'ok'|'err'; rows: number; err?: string;
}
interface Hit {
  id: string; fid: string; fn: string; sn: string|null;
  rn: number; ss: string; sc: number; ri: number;
}

// ── Constants ─────────────────────────────────────────────────────────────────
const ROW_H = 64;
const OVER  = 12;
const DEB   = 200;
const MAX_H = 20000;
const SUPP  = new Set(['xlsx','xls','csv','txt']);

const fmtN = (n:number) => n>=1e6?(n/1e6).toFixed(1)+'M':n>=1e3?(n/1e3).toFixed(1)+'K':''+n;
const fmtB = (b:number) => b>=(1<<20)?(b/(1<<20)).toFixed(1)+' MB':b>=(1<<10)?(b/(1<<10)).toFixed(1)+' KB':b+' B';
const fmtT = (ms:number) => ms>=60000?(ms/60000).toFixed(1)+'m':ms>=1000?(ms/1000).toFixed(2)+'s':ms+'ms';

type Ext = 'xlsx'|'xls'|'csv'|'txt';
const EXT: Record<Ext,{c:string;dim:string;I:any}> = {
  xlsx:{c:'#38bdf8',dim:'rgba(56,189,248,0.08)',I:FileSpreadsheet},
  xls: {c:'#818cf8',dim:'rgba(129,140,248,0.08)',I:FileSpreadsheet},
  csv: {c:'#4ade80',dim:'rgba(74,222,128,0.08)',I:FileCode},
  txt: {c:'#fb923c',dim:'rgba(251,146,60,0.08)',I:FileText},
};
const eOf = (n:string) => (n.split('.').pop()?.toLowerCase()||'xlsx') as Ext;

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [files,      setFiles]      = useState<FM[]>([]);
  const [totRows,    setTotRows]    = useState(0);
  const [query,      setQuery]      = useState('');
  const [hits,       setHits]       = useState<Hit[]>([]);
  const [sortedHits, setSortedHits] = useState<Hit[]>([]);
  const [busy,       setBusy]       = useState(false);
  const [elapsed,    setElapsed]    = useState<number|null>(null);
  const [fuzzy,      setFuzzy]      = useState(true);
  const [exact,      setExact]      = useState(false);
  const [exts,       setExts]       = useState<string[]>(['xlsx','xls','csv','txt']);
  const [sortK,      setSortK]      = useState<'sc'|'fn'|'rn'>('sc');
  const [sortD,      setSortD]      = useState<1|-1>(-1);
  const [expId,      setExpId]      = useState<string|null>(null);
  const [expD,       setExpD]       = useState<Record<string,string>|null>(null);
  const [expLoad,    setExpLoad]    = useState(false);
  const [hitCount,   setHitCount]   = useState(0);
  const [scrollTop,  setScrollTop]  = useState(0);
  const [viewH,      setViewH]      = useState(600);
  const bodyRef = useRef<HTMLDivElement>(null);

  // All mutable cross-render state in refs
  const workerRef  = useRef<InstanceType<typeof SearchWorker>|null>(null);
  const indexedRef = useRef<Map<string,FM>>(new Map());
  const iQueueRef  = useRef<FM[]>([]);
  const hitBufRef  = useRef<Hit[]>([]);
  const allHitsRef = useRef<Hit[]>([]);
  const sidRef     = useRef(0);
  const sDoneRef   = useRef(0);
  const sTotalRef  = useRef(0);
  const sStartRef  = useRef(0);
  const debRef     = useRef(0);
  const tickRef    = useRef(0);
  const queryRef   = useRef('');
  const fuzzyRef   = useRef(true);
  const exactRef   = useRef(false);
  const extsRef    = useRef<string[]>(['xlsx','xls','csv','txt']);
  const expIdRef   = useRef<string|null>(null);

  useEffect(()=>{ fuzzyRef.current = fuzzy; },[fuzzy]);
  useEffect(()=>{ exactRef.current = exact; },[exact]);
  useEffect(()=>{ extsRef.current  = exts;  },[exts]);

  // Sort
  useEffect(()=>{
    startTransition(()=>{
      const a = hits.slice();
      if      (sortK==='sc') a.sort((x,y)=>sortD*(y.sc-x.sc));
      else if (sortK==='fn') a.sort((x,y)=>sortD*x.fn.localeCompare(y.fn));
      else                   a.sort((x,y)=>sortD*(x.rn-y.rn));
      setSortedHits(a);
    });
  },[hits,sortK,sortD]);

  // Resize
  useEffect(()=>{
    const ro = new ResizeObserver(()=>{ if(bodyRef.current) setViewH(bodyRef.current.clientHeight); });
    if(bodyRef.current) ro.observe(bodyRef.current);
    return ()=>ro.disconnect();
  },[]);

  // ── Worker — boot once ────────────────────────────────────────────────────
  useEffect(()=>{
    const w = new SearchWorker();
    workerRef.current = w;

    function flushHits() {
      const snap = hitBufRef.current.splice(0);
      if (!snap.length) return;
      Array.prototype.push.apply(allHitsRef.current, snap);
      setHits(allHitsRef.current.slice(0, MAX_H));
      setHitCount(allHitsRef.current.length);
    }

    function dispatchNext() {
      if (!iQueueRef.current.length) return;
      const f = iQueueRef.current.shift()!;
      setFiles(prev => prev.map(x => x.id===f.id ? {...x,st:'ing' as const} : x));
      w.postMessage({t:'idx', id:f.id, blob:f.blob, ft:f.type, name:f.name});
    }

    w.onmessage = (ev: MessageEvent) => {
      const d = ev.data;

      if (d.t === 'ok') {
        setFiles(prev => {
          const nx = prev.map(x => x.id===d.id ? {...x,st:'ok' as const,rows:d.n} : x);
          const entry = nx.find(x => x.id===d.id);
          if (entry) indexedRef.current.set(d.id, entry);
          return nx;
        });
        setTotRows(r => r + d.n);
        dispatchNext();

      } else if (d.t === 'err') {
        setFiles(prev => prev.map(x => x.id===d.id ? {...x,st:'err' as const,err:d.msg} : x));
        dispatchNext();

      } else if (d.t === 'hits') {
        if (d.sid !== sidRef.current) return;
        Array.prototype.push.apply(hitBufRef.current, d.hits);
        if (hitBufRef.current.length >= 400) flushHits();

      } else if (d.t === 'done') {
        if (d.sid !== sidRef.current) return;
        sDoneRef.current++;
        if (sDoneRef.current >= sTotalRef.current) {
          flushHits();
          setBusy(false);
          clearInterval(tickRef.current);
          setElapsed(Date.now() - sStartRef.current);
        }

      } else if (d.t === 'row') {
        setExpD(d.d);
        setExpLoad(false);

      } else if (d.t === 'csv') {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(d.blob);
        a.download = 'locallens_'+Date.now()+'.csv';
        a.click();
      }
    };

    w.onerror = (e: ErrorEvent) => {
      console.error('Worker error:', e.message, e);
    };

    return ()=>{
      w.terminate();
      clearInterval(tickRef.current);
      clearTimeout(debRef.current);
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  // ── Load folder ────────────────────────────────────────────────────────────
  function loadFolder(e: React.ChangeEvent<HTMLInputElement>) {
    const input = e.target;
    if (!input.files || !input.files.length) return;
    // Copy FileList to array BEFORE clearing input value.
    // Clearing input.value can invalidate the FileList reference in Chrome.
    const raw: File[] = Array.from(input.files);
    input.value = '';

    const w = workerRef.current;
    if (!w) { console.error('Worker not ready'); return; }

    w.postMessage({t:'clear'});
    indexedRef.current.clear();
    iQueueRef.current  = [];
    hitBufRef.current  = [];
    allHitsRef.current = [];
    clearInterval(tickRef.current);
    clearTimeout(debRef.current);

    setFiles([]);
    setHits([]);
    setSortedHits([]);
    setQuery('');       queryRef.current = '';
    setTotRows(0);
    setElapsed(null);
    setBusy(false);
    setExpId(null);     expIdRef.current = null;
    setExpD(null);
    setHitCount(0);

    const list: FM[] = [];
    for (let i = 0; i < raw.length; i++) {
      const f = raw[i];
      // Skip directory entries — they have size 0 and no real extension
      if (f.size === 0) continue;
      const parts = f.name.split('.');
      if (parts.length < 2) continue;
      const ext   = parts[parts.length - 1].toLowerCase();
      if (!SUPP.has(ext)) continue;
      list.push({
        id:   `${f.name}-${f.size}-${f.lastModified}`,
        name: f.name, type: ext, blob: f, size: f.size, st:'q', rows:0
      });
    }
    if (!list.length) {
      // Show what was actually received for debugging
      const names = raw.slice(0,5).map((f:File)=>f.name+'('+f.size+')').join(', ');
      alert(`No supported files (xlsx/xls/csv/txt) found.\nReceived: ${names}${raw.length>5?'...':''}`);
      return;
    }

    setFiles(list);
    iQueueRef.current = list.slice();

    const first = iQueueRef.current.shift()!;
    setFiles(prev => prev.map(x => x.id===first.id ? {...x,st:'ing' as const} : x));
    w.postMessage({t:'idx', id:first.id, blob:first.blob, ft:first.type, name:first.name});
  }

  // ── Search ─────────────────────────────────────────────────────────────────
  function runSearch(q: string) {
    const w = workerRef.current;
    if (!w) return;

    if (!q.trim()) {
      setHits([]); setSortedHits([]);
      allHitsRef.current = []; hitBufRef.current = [];
      setElapsed(null); setHitCount(0);
      return;
    }

    const targets = Array.from(indexedRef.current.values())
      .filter(f => extsRef.current.includes(f.type));
    if (!targets.length) return;

    const newSid = ++sidRef.current;
    hitBufRef.current  = [];
    allHitsRef.current = [];
    sDoneRef.current   = 0;
    sTotalRef.current  = targets.length;
    sStartRef.current  = Date.now();

    setHits([]); setSortedHits([]);
    setElapsed(null); setHitCount(0);
    setBusy(true);
    setExpId(null); expIdRef.current = null;
    setExpD(null);

    clearInterval(tickRef.current);
    tickRef.current = setInterval(
      ()=>setElapsed(Date.now()-sStartRef.current), 100
    ) as unknown as number;

    for (const f of targets) {
      w.postMessage({t:'search', sid:newSid, id:f.id,
        q, fuzzy:fuzzyRef.current, exact:exactRef.current});
    }
  }

  function trigSearch(q: string) {
    clearTimeout(debRef.current);
    queryRef.current = q;
    debRef.current = setTimeout(()=>runSearch(q), DEB) as unknown as number;
  }

  useEffect(()=>{
    if (queryRef.current.trim()) runSearch(queryRef.current);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[fuzzy,exact,exts]);

  // ── Expand ─────────────────────────────────────────────────────────────────
  function expand(h: Hit) {
    if (expIdRef.current === h.id) {
      expIdRef.current = null;
      setExpId(null); setExpD(null); return;
    }
    expIdRef.current = h.id;
    setExpId(h.id); setExpD(null); setExpLoad(true);
    workerRef.current?.postMessage({t:'row', id:h.fid, ri:h.ri});
  }

  // ── Virtual scroll ─────────────────────────────────────────────────────────
  const {vis, vs} = useMemo(()=>{
    const s = Math.max(0, Math.floor(scrollTop/ROW_H)-OVER);
    const e = Math.min(sortedHits.length, Math.ceil((scrollTop+viewH)/ROW_H)+OVER);
    return {vis:sortedHits.slice(s,e), vs:s};
  },[sortedHits,scrollTop,viewH]);

  const nOk   = files.filter(f=>f.st==='ok').length;
  const nErr  = files.filter(f=>f.st==='err').length;
  const isIdx = files.some(f=>f.st==='ing'||f.st==='q');
  const ready = !isIdx && nOk > 0;
  const totSz = files.reduce((s,f)=>s+f.size,0);
  const pct   = files.length ? Math.round(nOk/files.length*100) : 0;
  const capped = hitCount >= MAX_H;

  return (
    <div className="ll-root">
      <header className="ll-header">
        <div className="ll-topbar">

          <div className="ll-logo">
            <div className="ll-logo-icon"><Zap size={15} color="#0a0b0e"/></div>
            <span className="ll-logo-text">LOCAL<span className="ll-acc">LENS</span></span>
          </div>

          <div className="ll-search-wrap">
            <span className="ll-search-ico">
              {busy
                ? <Loader2 size={14} className="ll-spin" style={{color:'#f59e0b'}}/>
                : <Search  size={14} style={{color:query?'#f59e0b':'#4a5568'}}/>}
            </span>
            <input className="ll-input" value={query} disabled={!ready}
              autoComplete="off" spellCheck={false}
              placeholder={
                isIdx  ? `Indexing ${nOk}/${files.length} files…` :
                !files.length ? 'Open a folder to start searching' :
                !ready ? 'Building index…' :
                `Search ${fmtN(totRows)} rows across ${files.length} files`
              }
              onChange={e=>{setQuery(e.target.value);trigSearch(e.target.value);}}
            />
            {query && (
              <button className="ll-clear" onClick={()=>{setQuery('');runSearch('');}}>
                <X size={13}/>
              </button>
            )}
          </div>

          <Chip on={fuzzy} click={()=>setFuzzy(v=>!v)} label="FUZZY" title="All terms must appear"/>
          <Chip on={exact} click={()=>setExact(v=>!v)} label="EXACT" title="Exact cell match"/>
          <Sep/>
          {(['xlsx','xls','csv','txt'] as Ext[]).map(t=>(
            <TypeChip key={t} t={t} on={exts.includes(t)}
              click={()=>setExts(p=>p.includes(t)?p.filter(x=>x!==t):[...p,t])}/>
          ))}
          <Sep/>
          {(['sc','fn','rn'] as const).map((k,i)=>(
            <SortChip key={k} label={['SCORE','FILE','ROW'][i]}
              active={sortK===k} dir={sortD}
              click={()=>{ if(sortK===k) setSortD(d=>d===1?-1:1); else{setSortK(k);setSortD(-1);} }}/>
          ))}
          <Sep/>

          <label className="ll-btn ll-btn-ghost">
            <FolderOpen size={13}/>
            {files.length?'CHANGE':'OPEN FOLDER'}
            <input type="file" style={{display:'none'}}
              // @ts-ignore
              webkitdirectory="" directory="" multiple onChange={loadFolder}/>
          </label>

          {hitCount > 0 && (
            <button className="ll-btn ll-btn-acc"
              onClick={()=>workerRef.current?.postMessage({t:'export',hits:allHitsRef.current})}>
              <Download size={13}/> EXPORT
            </button>
          )}
        </div>

        <div className="ll-status">
          <div className="ll-sl">
            {files.length > 0 && <>
              <span className="ll-dot" style={{background:isIdx?'#f59e0b':'#4ade80',boxShadow:`0 0 6px ${isIdx?'#f59e0b':'#4ade80'}`}}/>
              <span className="ll-lbl" style={{color:isIdx?'#f59e0b':'#4ade80'}}>
                {isIdx?`INDEXING ${nOk}/${files.length}`:`${nOk} FILES READY`}
                {nErr>0&&<span style={{color:'#f87171',marginLeft:8}}>{nErr} ERR</span>}
              </span>
              <span className="ll-sdiv"/>
              <span className="ll-stat">{fmtN(totRows)} ROWS</span>
              <span className="ll-sdiv"/>
              <span className="ll-stat">{fmtB(totSz)}</span>
              {isIdx&&<>
                <span className="ll-sdiv"/>
                <div className="ll-pbar"><div className="ll-pbar-f" style={{width:pct+'%'}}/></div>
                <span style={{fontSize:9,color:'#38bdf8',fontWeight:700}}>{pct}%</span>
              </>}
            </>}
          </div>
          <div className="ll-sr">
            {hitCount>0&&<span style={{fontSize:9,fontWeight:700,color:capped?'#f59e0b':'#38bdf8'}}>
              {capped?fmtN(MAX_H)+'+':fmtN(hitCount)} RESULTS
            </span>}
            {elapsed!==null&&<span style={{fontSize:9,color:busy?'#f59e0b':'#4a5568'}}>
              {fmtT(elapsed)}{busy?'…':''}
            </span>}
          </div>
        </div>
      </header>

      <div className="ll-body" ref={bodyRef} onScroll={e=>setScrollTop(e.currentTarget.scrollTop)}>
        {files.length===0 && <Welcome/>}
        {files.length>0 && isIdx && !query && sortedHits.length===0 && <Indexing files={files}/>}
        {ready && !query && sortedHits.length===0 && <Ready n={files.length} rows={totRows}/>}
        {ready && query && sortedHits.length===0 && !busy && <NoResults q={query}/>}

        {sortedHits.length>0 && (
          <div style={{position:'relative',height:sortedHits.length*ROW_H}}>
            <div style={{position:'absolute',top:0,left:0,right:0,
              transform:`translateY(${vs*ROW_H}px)`}}>
              {vis.map(h=>(
                <HitRow key={h.id} h={h}
                  open={expId===h.id}
                  data={expId===h.id?expD:undefined}
                  loading={expId===h.id&&expLoad}
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
const Sep=()=><div className="ll-sep"/>;

const Chip=memo(({on,click,label,title}:{on:boolean;click:()=>void;label:string;title?:string})=>(
  <button className={`ll-chip${on?' ll-chip-on':''}`} onClick={click} title={title}>
    {on&&<span className="ll-chip-dot"/>}{label}
  </button>
));

const TypeChip=memo(({t,on,click}:{t:Ext;on:boolean;click:()=>void})=>{
  const e=EXT[t];
  return <button className="ll-typechip" onClick={click}
    style={{color:on?e.c:'#4a5568',background:on?e.dim:'transparent',borderColor:on?e.c+'50':'#1c2030'}}>
    {t.toUpperCase()}
  </button>;
});

const SortChip=memo(({label,active,dir,click}:{label:string;active:boolean;dir:1|-1;click:()=>void})=>(
  <button className={`ll-sortchip${active?' ll-sortchip-on':''}`} onClick={click}>
    {label}{active?(dir===-1?' ↓':' ↑'):''}
  </button>
));

const Welcome=memo(()=>(
  <div className="ll-center">
    <div className="ll-wi"><ScanSearch size={28} strokeWidth={1.5} style={{color:'#4a5568'}}/></div>
    <div className="ll-wt">OPEN A FOLDER TO BEGIN</div>
    <div className="ll-ws">Index once. Search at native JS speed. Everything stays local.</div>
    <div className="ll-badges">
      {(['xlsx','xls','csv','txt'] as Ext[]).map(t=>(
        <span key={t} className="ll-badge" style={{color:EXT[t].c,borderColor:EXT[t].c+'40',background:EXT[t].dim}}>
          {t.toUpperCase()}
        </span>
      ))}
    </div>
    <div className="ll-features">
      {['⚡ Parse once, search instantly','🔒 100% local','🔍 Fuzzy + exact match','📤 Export to CSV'].map(s=>(
        <div key={s} style={{fontSize:11,color:'#2d3748'}}>{s}</div>
      ))}
    </div>
  </div>
));

const Indexing=memo(({files}:{files:FM[]})=>{
  const done=files.filter(f=>f.st==='ok').length;
  const pct=files.length?Math.round(done/files.length*100):0;
  return (
    <div className="ll-center">
      <Loader2 size={26} className="ll-spin" style={{color:'#f59e0b',marginBottom:14}}/>
      <div style={{fontSize:11,fontWeight:700,color:'#4a5568',letterSpacing:'.1em',marginBottom:5}}>BUILDING INDEX</div>
      <div style={{fontSize:10,color:'#2d3748',marginBottom:12}}>{done} / {files.length} files · {pct}%</div>
      <div className="ll-idx-bar"><div className="ll-idx-fill" style={{width:pct+'%'}}/></div>
      <div className="ll-idx-list">
        {files.slice(0,12).map(f=>(
          <div key={f.id} className="ll-idx-row">
            <span style={{lineHeight:0,flexShrink:0}}>
              {f.st==='ok'?<CheckCircle2 size={10} style={{color:'#4ade80'}}/>
              :f.st==='err'?<AlertCircle size={10} style={{color:'#f87171'}}/>
              :f.st==='ing'?<Loader2 size={10} className="ll-spin" style={{color:'#38bdf8'}}/>
              :<span style={{display:'inline-block',width:10,height:10,borderRadius:'50%',background:'#2d3748'}}/>}
            </span>
            <span style={{flex:1,fontSize:9,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',
              color:f.st==='ok'?'#e2e8f0':f.st==='err'?'#f87171':'#4a5568'}}>{f.name}</span>
            {f.st==='ok'&&<span style={{fontSize:8,color:'#4a5568',flexShrink:0}}>{f.rows.toLocaleString()}</span>}
            {f.st==='err'&&<span style={{fontSize:8,color:'#f87171',fontWeight:700}}>ERR</span>}
          </div>
        ))}
        {files.length>12&&<div style={{fontSize:9,color:'#2d3748',textAlign:'center',padding:4}}>+{files.length-12} more queued</div>}
      </div>
    </div>
  );
});

const Ready=memo(({n,rows}:{n:number;rows:number})=>(
  <div className="ll-center" style={{opacity:.35}}>
    <Search size={32} strokeWidth={1} style={{color:'#f59e0b'}}/>
    <div style={{fontSize:11,color:'#4a5568',letterSpacing:'.12em',marginTop:12}}>{n} FILES · {fmtN(rows)} ROWS · TYPE TO SEARCH</div>
  </div>
));

const NoResults=memo(({q}:{q:string})=>(
  <div className="ll-center" style={{opacity:.5}}>
    <Search size={28} strokeWidth={1} style={{color:'#4a5568'}}/>
    <div style={{fontSize:12,color:'#4a5568',marginTop:12}}>No results for <span style={{color:'#e2e8f0'}}>"{q}"</span></div>
    <div style={{fontSize:10,color:'#2d3748',marginTop:4}}>Try fuzzy mode or check type filters</div>
  </div>
));

const Hl=memo(({text,q}:{text:string;q:string})=>{
  if(!q.trim()||!text) return <>{text}</>;
  const terms=q.trim().split(/\s+/).filter(t=>t.length>0);
  if(!terms.length) return <>{text}</>;
  try {
    const re=new RegExp(`(${terms.map(t=>t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|')})`,'gi');
    return <>{text.split(re).map((p,i)=>re.test(p)?<mark key={i} className="ll-hl">{p}</mark>:p)}</>;
  } catch { return <>{text}</>; }
});

const HitRow=memo(({h,open,data,loading,onOpen,q}:{
  h:Hit;open:boolean;data?:Record<string,string>|null;loading?:boolean;onOpen:(h:Hit)=>void;q:string;
})=>{
  const es=EXT[eOf(h.fn)]||EXT.xlsx;
  return (
    <div className={`ll-hit${open?' ll-hit-open':''}`}>
      <div className="ll-hit-row" onClick={()=>onOpen(h)}>
        <div className="ll-hit-ico" style={{background:es.dim,borderColor:es.c+'35',color:es.c}}>
          <es.I size={13} strokeWidth={1.5}/>
        </div>
        <div className="ll-hit-body">
          <div className="ll-hit-meta">
            <span className="ll-hit-fn">{h.fn}</span>
            {h.sn&&<span className="ll-hit-sn">{h.sn}</span>}
            <span className="ll-hit-rn">ROW {h.rn}</span>
          </div>
          <div className="ll-hit-prev"><Hl text={h.ss} q={q}/></div>
        </div>
        <div className="ll-hit-right">
          {h.sc>1&&<span className="ll-hit-sc">{h.sc}</span>}
          <ChevronDown size={13} className={`ll-chev${open?' ll-chev-open':''}`}/>
        </div>
      </div>
      {open&&(
        <div className="ll-detail">
          {loading?<div className="ll-dl"><Loader2 size={11} className="ll-spin"/> Loading row…</div>
          :!data||!Object.keys(data).length?<div className="ll-dl"><AlertCircle size={11}/> No data</div>
          :<div className="ll-dgrid">
            {Object.entries(data).map(([k,v])=>(
              <div key={k} className="ll-dcell">
                <div className="ll-dk">{k}</div>
                <div className="ll-dv"><Hl text={String(v||'—')} q={q}/></div>
              </div>
            ))}
          </div>}
        </div>
      )}
    </div>
  );
});

// ── CSS ───────────────────────────────────────────────────────────────────────
{
  const s=document.createElement('style');
  s.textContent=`
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
    :root{--bg:#0a0b0e;--surf:#0f1117;--surf2:#141720;--brd:#1c2030;--brd2:#252b3b;
      --txt:#e2e8f0;--muted:#4a5568;--dim:#2d3748;
      --acc:#f59e0b;--blue:#38bdf8;--green:#4ade80;--red:#f87171}
    html,body,#root{height:100%;overflow:hidden}
    body{background:var(--bg);color:var(--txt)}
    @keyframes spin{to{transform:rotate(360deg)}}
    @keyframes fadeIn{from{opacity:0;transform:translateY(3px)}to{opacity:1;transform:none}}
    .ll-spin{animation:spin .8s linear infinite}
    .ll-root{display:flex;flex-direction:column;height:100vh;overflow:hidden;
      font-family:'IBM Plex Mono','Fira Code','Cascadia Code',monospace;background:var(--bg)}
    .ll-header{flex-shrink:0;background:var(--surf);border-bottom:1px solid var(--brd)}
    .ll-topbar{display:flex;align-items:center;gap:7px;padding:9px 14px;flex-wrap:wrap}
    .ll-logo{display:flex;align-items:center;gap:8px;flex-shrink:0;margin-right:4px}
    .ll-logo-icon{width:28px;height:28px;border-radius:7px;background:var(--acc);flex-shrink:0;
      display:flex;align-items:center;justify-content:center;box-shadow:0 0 14px rgba(245,158,11,.4)}
    .ll-logo-text{font-size:13px;font-weight:700;letter-spacing:.08em;color:var(--txt);user-select:none}
    .ll-acc{color:var(--acc)}
    .ll-search-wrap{flex:1;min-width:160px;position:relative;display:flex;align-items:center}
    .ll-search-ico{position:absolute;left:11px;top:50%;transform:translateY(-50%);pointer-events:none;line-height:0}
    .ll-input{width:100%;padding:8px 32px 8px 34px;background:var(--surf2);border:1px solid var(--brd2);
      border-radius:7px;color:var(--txt);font-size:13px;font-family:inherit;outline:none;
      transition:border-color .15s,box-shadow .15s}
    .ll-input:focus{border-color:var(--acc);box-shadow:0 0 0 3px rgba(245,158,11,.1)}
    .ll-input::placeholder{color:var(--muted)}
    .ll-input:disabled{opacity:.35;cursor:not-allowed}
    .ll-clear{position:absolute;right:8px;top:50%;transform:translateY(-50%);
      background:none;border:none;cursor:pointer;color:var(--muted);line-height:0;padding:2px;transition:color .1s}
    .ll-clear:hover{color:var(--txt)}
    .ll-sep{width:1px;height:16px;background:var(--brd);flex-shrink:0}
    .ll-chip{display:flex;align-items:center;gap:4px;padding:3px 9px;border-radius:5px;
      cursor:pointer;flex-shrink:0;background:transparent;border:1px solid var(--brd);
      color:var(--muted);font-size:9px;font-weight:700;letter-spacing:.1em;font-family:inherit;transition:all .1s}
    .ll-chip:hover{color:var(--txt)}
    .ll-chip-on{background:rgba(245,158,11,.08);border-color:rgba(245,158,11,.5);color:var(--acc)}
    .ll-chip-dot{width:5px;height:5px;border-radius:50%;background:var(--acc);flex-shrink:0;box-shadow:0 0 5px var(--acc)}
    .ll-typechip{padding:3px 8px;border-radius:5px;cursor:pointer;flex-shrink:0;border:1px solid;
      font-size:9px;font-weight:700;letter-spacing:.08em;font-family:inherit;transition:all .1s}
    .ll-sortchip{padding:3px 8px;border-radius:5px;cursor:pointer;flex-shrink:0;
      background:transparent;border:1px solid transparent;color:var(--muted);
      font-size:9px;font-weight:600;font-family:inherit;transition:all .1s}
    .ll-sortchip:hover{color:var(--txt)}
    .ll-sortchip-on{background:rgba(56,189,248,.08);border-color:rgba(56,189,248,.3);color:var(--blue)}
    .ll-btn{display:flex;align-items:center;gap:5px;padding:6px 11px;border-radius:6px;
      cursor:pointer;flex-shrink:0;font-size:10px;font-weight:600;letter-spacing:.07em;
      font-family:inherit;white-space:nowrap;transition:all .15s}
    .ll-btn-ghost{background:var(--surf2);border:1px solid var(--brd2);color:var(--muted)}
    .ll-btn-ghost:hover{color:var(--txt)}
    .ll-btn-acc{background:var(--acc);border:none;color:#0a0b0e;font-weight:700;box-shadow:0 0 12px rgba(245,158,11,.3)}
    .ll-btn-acc:hover{box-shadow:0 0 20px rgba(245,158,11,.45)}
    .ll-status{display:flex;align-items:center;justify-content:space-between;
      padding:3px 14px 5px;border-top:1px solid var(--brd);font-size:9px;letter-spacing:.08em}
    .ll-sl,.ll-sr{display:flex;align-items:center;gap:8px}
    .ll-lbl{font-weight:700}
    .ll-stat{color:var(--muted)}
    .ll-sdiv{width:1px;height:10px;background:var(--brd)}
    .ll-dot{display:inline-block;width:6px;height:6px;border-radius:50%;flex-shrink:0}
    .ll-pbar{width:70px;height:2px;background:var(--brd);border-radius:2px;overflow:hidden}
    .ll-pbar-f{height:100%;background:var(--acc);border-radius:2px;transition:width .3s}
    .ll-body{flex:1;overflow-y:auto;overflow-x:hidden}
    .ll-body::-webkit-scrollbar{width:4px}
    .ll-body::-webkit-scrollbar-track{background:transparent}
    .ll-body::-webkit-scrollbar-thumb{background:var(--brd2);border-radius:4px}
    .ll-center{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:32px}
    .ll-wi{width:60px;height:60px;border-radius:16px;background:var(--surf);border:1px solid var(--brd2);
      display:flex;align-items:center;justify-content:center;margin-bottom:16px}
    .ll-wt{font-size:12px;font-weight:700;color:var(--muted);letter-spacing:.1em;margin-bottom:9px}
    .ll-ws{font-size:11px;color:var(--dim);text-align:center;line-height:1.8;max-width:300px;margin-bottom:16px}
    .ll-badges{display:flex;gap:6px;margin-bottom:16px;flex-wrap:wrap;justify-content:center}
    .ll-badge{padding:2px 9px;border-radius:20px;border:1px solid;font-size:8px;font-weight:700;letter-spacing:.1em}
    .ll-features{display:flex;flex-direction:column;gap:4px;align-items:flex-start}
    .ll-idx-bar{width:180px;height:2px;background:var(--brd);border-radius:2px;overflow:hidden;margin-bottom:16px}
    .ll-idx-fill{height:100%;background:var(--acc);border-radius:2px;transition:width .4s}
    .ll-idx-list{display:flex;flex-direction:column;gap:4px;width:256px}
    .ll-idx-row{display:flex;align-items:center;gap:7px;padding:4px 8px;border-radius:5px;background:var(--surf);border:1px solid var(--brd)}
    .ll-hit{border-bottom:1px solid var(--brd);transition:background .08s}
    .ll-hit-open{background:var(--surf)}
    .ll-hit-row{display:flex;align-items:center;height:64px;padding:0 14px;cursor:pointer;gap:11px}
    .ll-hit:not(.ll-hit-open) .ll-hit-row:hover{background:var(--surf)}
    .ll-hit-ico{width:30px;height:30px;border-radius:7px;flex-shrink:0;border:1px solid;display:flex;align-items:center;justify-content:center}
    .ll-hit-body{flex:1;min-width:0}
    .ll-hit-meta{display:flex;align-items:center;gap:7px;margin-bottom:3px}
    .ll-hit-fn{font-size:11px;font-weight:600;color:var(--txt);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:260px}
    .ll-hit-sn{font-size:8px;font-weight:600;color:var(--blue);background:rgba(56,189,248,.08);border:1px solid rgba(56,189,248,.25);border-radius:3px;padding:1px 4px;flex-shrink:0}
    .ll-hit-rn{font-size:9px;color:var(--dim);flex-shrink:0;letter-spacing:.05em}
    .ll-hit-prev{font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .ll-hit-right{display:flex;align-items:center;gap:7px;flex-shrink:0}
    .ll-hit-sc{font-size:9px;font-weight:700;color:var(--acc);background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.25);border-radius:4px;padding:1px 5px}
    .ll-chev{color:var(--dim);transition:transform .15s;flex-shrink:0}
    .ll-chev-open{transform:rotate(180deg);color:var(--acc)}
    .ll-detail{padding:10px 14px 14px;border-top:1px solid var(--brd);background:rgba(0,0,0,.18);animation:fadeIn .12s ease}
    .ll-dl{display:flex;align-items:center;gap:7px;font-size:11px;color:var(--muted);padding:7px 0}
    .ll-dgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:7px}
    .ll-dcell{padding:7px 9px;border-radius:5px;background:var(--surf2);border:1px solid var(--brd)}
    .ll-dk{font-size:8px;font-weight:600;color:var(--dim);letter-spacing:.12em;margin-bottom:3px;text-transform:uppercase}
    .ll-dv{font-size:11px;color:var(--txt);word-break:break-word;line-height:1.4}
    .ll-hl{background:rgba(245,158,11,.2);color:var(--acc);border-radius:2px;padding:0 1px}
    ::-webkit-scrollbar{width:4px;height:4px}
    ::-webkit-scrollbar-track{background:transparent}
    ::-webkit-scrollbar-thumb{background:var(--brd2);border-radius:4px}
    button,input{font-family:inherit}
  `;
  document.head.appendChild(s);
}
