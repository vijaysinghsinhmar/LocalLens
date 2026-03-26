import React, { useEffect, useRef, useState, useMemo, useCallback, memo } from 'react';
import {
  Zap, FolderOpen, Search, X, Download,
  Loader2, CheckCircle2, AlertCircle, ChevronDown,
  FileSpreadsheet, FileCode, FileText
} from 'lucide-react';
import { workerScript } from './worker';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────
interface FileMeta {
  id: string; name: string; type: string; blob: File; size: number;
  state: 'pending'|'indexing'|'done'|'error';
  rows: number; err?: string;
}
interface Hit {
  id: string; fid: string; fn: string; sn: string|null;
  rn: number; ss: string; sc: number; ri: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────
const N_WORKERS  = 2;
const ROW_H      = 62;
const OVER       = 12;
const DEBOUNCE   = 220;
const MAX_HITS   = 25000;
const SUPPORTED  = new Set(['xlsx','xls','csv','txt']);

const fmt  = (n:number) => n>=1e6?(n/1e6).toFixed(1)+'M':n>=1e3?(n/1e3).toFixed(1)+'K':''+n;
const fmtB = (b:number) => b>=1<<20?(b/(1<<20)).toFixed(1)+' MB':b>=1<<10?(b/(1<<10)).toFixed(1)+' KB':b+' B';
const fmtT = (ms:number) => ms>=1000?(ms/1000).toFixed(2)+'s':ms+'ms';

type Ext = 'xlsx'|'xls'|'csv'|'txt';
const EXT: Record<Ext,{c:string;d:string;I:any}> = {
  xlsx:{c:'#60a5fa',d:'#0c1e3a',I:FileSpreadsheet},
  xls: {c:'#818cf8',d:'#150d35',I:FileSpreadsheet},
  csv: {c:'#34d399',d:'#011a0e',I:FileCode},
  txt: {c:'#fbbf24',d:'#120c02',I:FileText},
};
const eOf = (n:string) => (n.split('.').pop()?.toLowerCase()||'xlsx') as Ext;

// ─────────────────────────────────────────────────────────────────────────────
// App
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  // Files
  const [files,    setFiles]   = useState<FileMeta[]>([]);
  const [totRows,  setTotRows] = useState(0);

  // Search
  const [query,    setQuery]   = useState('');
  const [hits,     setHits]    = useState<Hit[]>([]);
  const [busy,     setBusy]    = useState(false);
  const [ms,       setMs]      = useState<number|null>(null);

  // Options
  const [fuzzy,    setFuzzy]   = useState(true);
  const [exact,    setExact]   = useState(false);
  const [exts,     setExts]    = useState(['xlsx','xls','csv','txt']);
  const [sortK,    setSortK]   = useState<'sc'|'fn'|'rn'>('sc');
  const [sortD,    setSortD]   = useState<1|-1>(-1);

  // Expand
  const [expId,    setExpId]   = useState<string|null>(null);
  const [expData,  setExpData] = useState<Record<string,string>|null>(null);
  const [expLoad,  setExpLoad] = useState(false);

  // Virtual scroll
  const [sTop,     setSTop]    = useState(0);
  const [vH,       setVH]      = useState(600);
  const scrollEl = useRef<HTMLDivElement>(null);

  // Internal refs — never trigger re-renders
  const ws        = useRef<Worker[]>([]);
  const wFree     = useRef<boolean[]>([]);        // true = worker is idle
  const iQueue    = useRef<FileMeta[]>([]);       // files waiting to be indexed
  const iTotal    = useRef(0);
  const iDone     = useRef(0);
  const sid       = useRef(0);
  const sTotal    = useRef(0);
  const sDone     = useRef(0);
  const sStart    = useRef(0);
  const sBuf      = useRef<Hit[]>([]);
  const allHits   = useRef<Hit[]>([]);
  const debTmr    = useRef(0);
  const tickTmr   = useRef(0);
  const filesRef  = useRef<FileMeta[]>([]);       // mirror of files state, readable from handlers

  // ── Resize ────────────────────────────────────────────────────────────────
  useEffect(() => {
    const ro = new ResizeObserver(() => {
      if (scrollEl.current) setVH(scrollEl.current.clientHeight);
    });
    if (scrollEl.current) ro.observe(scrollEl.current);
    return () => ro.disconnect();
  }, []);

  // ── Boot workers (once) ───────────────────────────────────────────────────
  useEffect(() => {
    const blob = new Blob([workerScript], {type:'application/javascript'});
    const url  = URL.createObjectURL(blob);

    // Dispatch next file from the index queue to a specific worker
    function dispatchNext(wi: number) {
      const q = iQueue.current;
      if (!q.length) return;
      const f = q.shift()!;
      wFree.current[wi] = false;
      // Mark file as indexing
      setFiles(prev => {
        const next = prev.map(x => x.id===f.id ? {...x,state:'indexing' as const} : x);
        filesRef.current = next;
        return next;
      });
      ws.current[wi].postMessage({
        type:'index', id:f.id, blob:f.blob, ftype:f.type, name:f.name
      });
    }

    ws.current = Array.from({length:N_WORKERS}, (_, wi) => {
      const w = new Worker(url);
      wFree.current[wi] = true;

      w.onmessage = (ev) => {
        const d = ev.data;

        if (d.type === 'indexed') {
          iDone.current++;
          wFree.current[wi] = true;
          setTotRows(r => r + d.count);
          setFiles(prev => {
            const next = prev.map(x => x.id===d.id ? {...x,state:'done' as const,rows:d.count} : x);
            filesRef.current = next;
            return next;
          });
          dispatchNext(wi);

        } else if (d.type === 'indexerr') {
          iDone.current++;
          wFree.current[wi] = true;
          setFiles(prev => {
            const next = prev.map(x => x.id===d.id ? {...x,state:'error' as const,err:d.err} : x);
            filesRef.current = next;
            return next;
          });
          dispatchNext(wi);

        } else if (d.type === 'results') {
          if (d.sid !== sid.current) return;
          sBuf.current.push(...d.hits);
          // Flush every 80ms — avoid flooding React
          if (sBuf.current.length > 500) flushSearchBuf();

        } else if (d.type === 'searchdone') {
          if (d.sid !== sid.current) return;
          sDone.current++;
          if (sDone.current >= sTotal.current) {
            flushSearchBuf();
            setBusy(false);
            clearInterval(tickTmr.current);
            setMs(Date.now() - sStart.current);
          }

        } else if (d.type === 'rowdata') {
          setExpData(d.data);
          setExpLoad(false);

        } else if (d.type === 'exportready') {
          const a = document.createElement('a');
          a.href = URL.createObjectURL(d.blob);
          a.download = 'locallens_export_'+Date.now()+'.csv';
          a.click();
        }
      };
      return w;
    });

    return () => {
      ws.current.forEach(w => w.terminate());
      URL.revokeObjectURL(url);
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function flushSearchBuf() {
    if (!sBuf.current.length) return;
    const snap = sBuf.current.splice(0);
    allHits.current.push(...snap);
    setHits(allHits.current.slice(0, MAX_HITS));
  }

  // ── Load folder ───────────────────────────────────────────────────────────
  const loadFolder = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const raw = e.target.files;
    if (!raw) return;
    // Reset everything
    ws.current.forEach(w => w.postMessage({type:'clear'}));
    iQueue.current = [];
    iTotal.current = 0;
    iDone.current  = 0;
    allHits.current = [];
    setHits([]);
    setQuery('');
    setMs(null);
    setTotRows(0);
    setExpId(null);
    setExpData(null);

    const list: FileMeta[] = [];
    for (let i = 0; i < raw.length; i++) {
      const f = raw[i];
      const ext = (f.name.split('.').pop()||'').toLowerCase();
      if (!SUPPORTED.has(ext)) continue;
      list.push({
        id:`${f.name}-${f.size}-${f.lastModified}`,
        name:f.name, type:ext, blob:f, size:f.size,
        state:'pending', rows:0
      });
    }

    filesRef.current = list;
    setFiles(list);

    // Fill index queue and seed workers
    iQueue.current = list.slice(); // copy
    iTotal.current = list.length;

    // Seed each worker with one file immediately
    for (let wi = 0; wi < N_WORKERS && iQueue.current.length; wi++) {
      const f = iQueue.current.shift()!;
      wFree.current[wi] = false;
      setFiles(prev => {
        const next = prev.map(x => x.id===f.id ? {...x,state:'indexing' as const} : x);
        filesRef.current = next;
        return next;
      });
      ws.current[wi].postMessage({
        type:'index', id:f.id, blob:f.blob, ftype:f.type, name:f.name
      });
    }
  }, []);

  // ── Run search ────────────────────────────────────────────────────────────
  const runSearch = useCallback((q: string, fz: boolean, ex: boolean, activeExts: string[]) => {
    if (!q.trim()) { setHits([]); allHits.current=[]; setMs(null); return; }

    const targets = filesRef.current.filter(f => f.state==='done' && activeExts.includes(f.type));
    if (!targets.length) return;

    const newSid = ++sid.current;
    sBuf.current    = [];
    allHits.current = [];
    sDone.current   = 0;
    sTotal.current  = targets.length;
    sStart.current  = Date.now();
    setHits([]);
    setMs(null);
    setBusy(true);
    setExpId(null);
    setExpData(null);

    clearInterval(tickTmr.current);
    tickTmr.current = setInterval(() => setMs(Date.now()-sStart.current), 120) as unknown as number;

    targets.forEach((f, i) => {
      ws.current[i % N_WORKERS].postMessage({
        type:'search', sid:newSid, id:f.id,
        query:q, fuzzy:fz, exact:ex
      });
    });
  }, []);

  const triggerSearch = useCallback((q: string) => {
    clearTimeout(debTmr.current);
    debTmr.current = setTimeout(() => runSearch(q, fuzzy, exact, exts), DEBOUNCE) as unknown as number;
  }, [runSearch, fuzzy, exact, exts]);

  // Re-search when options change
  useEffect(() => {
    if (query.trim()) runSearch(query, fuzzy, exact, exts);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fuzzy, exact, exts]);

  // ── Row expand ────────────────────────────────────────────────────────────
  const expand = useCallback((h: Hit) => {
    if (expId===h.id) { setExpId(null); setExpData(null); return; }
    setExpId(h.id);
    setExpData(null);
    setExpLoad(true);
    ws.current[0].postMessage({type:'row', id:h.fid, ri:h.ri});
  }, [expId]);

  // ── Sort ──────────────────────────────────────────────────────────────────
  const sorted = useMemo(() => {
    const a = hits.slice();
    if      (sortK==='sc') a.sort((x,y) => sortD*(y.sc-x.sc));
    else if (sortK==='fn') a.sort((x,y) => sortD*x.fn.localeCompare(y.fn));
    else                   a.sort((x,y) => sortD*(x.rn-y.rn));
    return a;
  }, [hits, sortK, sortD]);

  // ── Virtual window ────────────────────────────────────────────────────────
  const {visItems, visStart} = useMemo(() => {
    const start = Math.max(0, Math.floor(sTop/ROW_H)-OVER);
    const end   = Math.min(sorted.length, Math.ceil((sTop+vH)/ROW_H)+OVER);
    return {visItems:sorted.slice(start,end), visStart:start};
  }, [sorted, sTop, vH]);

  // ── Derived ───────────────────────────────────────────────────────────────
  const nDone    = files.filter(f=>f.state==='done').length;
  const nIndex   = files.filter(f=>f.state==='indexing'||f.state==='pending').length;
  const indexing = nIndex > 0;
  const ready    = !indexing && nDone > 0;
  const totalSz  = files.reduce((s,f)=>s+f.size,0);
  const capped   = allHits.current.length >= MAX_HITS;
  const pct      = files.length ? Math.round(nDone/files.length*100) : 0;

  return (
    <div style={{display:'flex',flexDirection:'column',height:'100vh',overflow:'hidden',
      background:'#06070a',color:'#94a3b8',fontFamily:"'DM Mono','Fira Code',monospace"}}>

      {/* ══ HEADER ══════════════════════════════════════════════════════════ */}
      <div style={{flexShrink:0,background:'#080a0e',borderBottom:'1px solid #0d1119'}}>

        <div style={{display:'flex',alignItems:'center',gap:8,padding:'9px 14px',flexWrap:'wrap'}}>

          {/* Logo */}
          <div style={{display:'flex',alignItems:'center',gap:7,flexShrink:0}}>
            <div style={{width:28,height:28,borderRadius:7,
              background:'linear-gradient(135deg,#1d4ed8,#4338ca)',
              display:'flex',alignItems:'center',justifyContent:'center',
              boxShadow:'0 0 12px #1d4ed850'}}>
              <Zap size={14} color="#fff"/>
            </div>
            <span style={{fontSize:12,fontWeight:700,letterSpacing:'0.08em',color:'#e2e8f0'}}>
              LOCAL<span style={{color:'#3b82f6'}}>LENS</span>
            </span>
          </div>

          {/* Search */}
          <div style={{flex:1,minWidth:180,position:'relative'}}>
            <Search size={13} style={{position:'absolute',left:10,top:'50%',
              transform:'translateY(-50%)',pointerEvents:'none',
              color:busy?'#3b82f6':'#1e3050',transition:'color 0.2s'}}/>
            <input value={query} disabled={!ready}
              onChange={e=>{setQuery(e.target.value);triggerSearch(e.target.value);}}
              placeholder={
                indexing ? `Indexing… (${nDone}/${files.length})` :
                !files.length ? 'Open a folder to start' :
                !ready ? 'Waiting…' :
                `Search ${fmt(totRows)} rows in ${files.length} files`
              }
              style={{width:'100%',boxSizing:'border-box',
                padding:'8px 32px 8px 30px',
                background:'#0d1017',border:'1px solid #141b28',
                borderRadius:7,color:'#e2e8f0',fontSize:13,outline:'none',
                fontFamily:'inherit',opacity:ready?1:0.45,
                transition:'border-color 0.15s,opacity 0.3s'}}
              onFocus={e=>e.target.style.borderColor='#1d4ed8'}
              onBlur={e =>e.target.style.borderColor='#141b28'}
            />
            {query && <button onClick={()=>{setQuery('');setHits([]);allHits.current=[];setMs(null);}}
              style={{position:'absolute',right:8,top:'50%',transform:'translateY(-50%)',
                background:'none',border:'none',cursor:'pointer',color:'#334155',padding:0,lineHeight:0}}>
              <X size={12}/></button>}
            {busy && <Loader2 size={12} color="#3b82f6"
              style={{position:'absolute',right:query?26:8,top:'50%',
                transform:'translateY(-50%)',animation:'spin 0.7s linear infinite'}}/>}
          </div>

          {/* Mode */}
          <T active={fuzzy}  on={()=>setFuzzy(!fuzzy)}  label="FUZZY"/>
          <T active={exact}  on={()=>setExact(!exact)}  label="EXACT"/>
          <Sep/>

          {/* File types */}
          {(['xlsx','xls','csv','txt'] as Ext[]).map(t=>(
            <T key={t} active={exts.includes(t)} color={EXT[t].c}
              on={()=>setExts(p=>p.includes(t)?p.filter(x=>x!==t):[...p,t])}
              label={t.toUpperCase()}/>
          ))}
          <Sep/>

          {/* Sort */}
          {([['sc','SCORE'],['fn','FILE'],['rn','ROW']] as [typeof sortK,string][]).map(([k,l])=>(
            <button key={k} onClick={()=>{
              if(sortK===k) setSortD(d=>d===1?-1:1); else{setSortK(k);setSortD(-1);}
            }} style={{padding:'3px 7px',borderRadius:5,cursor:'pointer',flexShrink:0,
              background:sortK===k?'#0a1425':'transparent',
              border:'1px solid '+(sortK===k?'#0f1e38':'transparent'),
              color:sortK===k?'#3b82f6':'#1e3050',fontSize:9,fontWeight:700}}>
              {l}{sortK===k?(sortD===-1?' ↓':' ↑'):''}
            </button>
          ))}
          <Sep/>

          {/* Folder */}
          <label style={{display:'flex',alignItems:'center',gap:5,padding:'5px 10px',
            borderRadius:6,cursor:'pointer',background:'#0d1017',
            border:'1px solid #141b28',fontSize:10,fontWeight:700,
            color:'#334155',whiteSpace:'nowrap',flexShrink:0}}>
            <FolderOpen size={12}/>
            {files.length?'CHANGE':'OPEN FOLDER'}
            <input type="file" style={{display:'none'}}
              // @ts-ignore
              webkitdirectory="" directory="" multiple onChange={loadFolder}/>
          </label>

          {/* Export */}
          {allHits.current.length>0 && (
            <button onClick={()=>ws.current[0].postMessage({type:'export',hits:allHits.current})}
              style={{display:'flex',alignItems:'center',gap:5,padding:'5px 10px',
                borderRadius:6,border:'none',cursor:'pointer',flexShrink:0,
                background:'linear-gradient(135deg,#1d4ed8,#4338ca)',
                fontSize:10,fontWeight:700,color:'#fff'}}>
              <Download size={11}/> CSV
            </button>
          )}
        </div>

        {/* Status strip */}
        <div style={{display:'flex',alignItems:'center',gap:10,
          padding:'3px 14px 5px',borderTop:'1px solid #08090d',fontSize:9,color:'#1e3050'}}>

          {files.length>0 && <>
            <span style={{color:indexing?'#f59e0b':'#22c55e',fontWeight:700,
              display:'flex',alignItems:'center',gap:4}}>
              {indexing
                ?<Loader2 size={8} style={{animation:'spin 0.7s linear infinite'}}/>
                :<CheckCircle2 size={8}/>}
              {indexing?`${nDone}/${files.length} indexed`:`${nDone} files`}
            </span>
            <span>{fmt(totRows)} rows</span>
            <span>{fmtB(totalSz)}</span>
            {indexing && <>
              <div style={{flex:1,height:2,background:'#0a0e17',borderRadius:1,minWidth:40}}>
                <div style={{height:'100%',width:pct+'%',background:'#1d4ed8',
                  borderRadius:1,transition:'width 0.3s'}}/>
              </div>
              <span style={{color:'#1d4ed8',fontWeight:700}}>{pct}%</span>
            </>}
          </>}

          <div style={{marginLeft:'auto',display:'flex',gap:10,alignItems:'center'}}>
            {hits.length>0 && <span style={{color:capped?'#f59e0b':'#3b82f6',fontWeight:700}}>
              {capped?fmt(MAX_HITS)+'+':fmt(hits.length)} results
            </span>}
            {ms!==null && <span style={{color:busy?'#f59e0b':'#1e3a5f'}}>
              {fmtT(ms)}{busy?'…':''}
            </span>}
          </div>
        </div>
      </div>

      {/* ══ BODY ════════════════════════════════════════════════════════════ */}
      <div ref={scrollEl} onScroll={e=>setSTop(e.currentTarget.scrollTop)}
        style={{flex:1,overflowY:'auto',overflowX:'hidden'}}>

        {files.length===0 && <Empty/>}
        {files.length>0 && indexing && hits.length===0 && !query && <IndexView files={files}/>}
        {ready && !query && hits.length===0 && <ReadyView n={files.length} rows={totRows}/>}
        {ready && query && hits.length===0 && !busy &&
          <div style={{height:'100%',display:'flex',flexDirection:'column',
            alignItems:'center',justifyContent:'center',gap:8,opacity:0.4}}>
            <Search size={30} color="#1e3a5f"/>
            <div style={{fontSize:12,color:'#334155'}}>No results for "{query}"</div>
          </div>}

        {sorted.length>0 && (
          <div style={{position:'relative',height:sorted.length*ROW_H}}>
            <div style={{position:'absolute',top:0,left:0,right:0,
              transform:`translateY(${visStart*ROW_H}px)`}}>
              {visItems.map(h=>(
                <Row key={h.id} h={h}
                  open={expId===h.id}
                  data={expId===h.id?expData:undefined}
                  loading={expId===h.id&&expLoad}
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

// ── Small shared bits ─────────────────────────────────────────────────────────
const Sep = ()=><div style={{width:1,height:14,background:'#0d1119',flexShrink:0}}/>;

const T = memo(({active,on,label,color}:{active:boolean;on:()=>void;label:string;color?:string})=>(
  <button onClick={on} style={{padding:'3px 8px',borderRadius:5,cursor:'pointer',flexShrink:0,
    background:active?(color?color+'12':'#0c1e3a'):'transparent',
    border:'1px solid '+(active?(color||'#1d4ed8'):'#0d1119'),
    color:active?(color||'#60a5fa'):'#1a2d47',
    fontSize:9,fontWeight:800,letterSpacing:'0.08em',transition:'all 0.1s'}}>
    {label}
  </button>
));

// ── Empty ─────────────────────────────────────────────────────────────────────
const Empty = memo(()=>(
  <div style={{height:'100%',display:'flex',flexDirection:'column',
    alignItems:'center',justifyContent:'center',gap:18}}>
    <div style={{width:60,height:60,borderRadius:14,background:'#0a0e17',
      border:'1px solid #0d1525',display:'flex',alignItems:'center',justifyContent:'center'}}>
      <FolderOpen size={24} color="#0d1e35"/>
    </div>
    <div style={{textAlign:'center'}}>
      <div style={{fontSize:13,fontWeight:700,color:'#0f1e30',letterSpacing:'0.08em',marginBottom:6}}>
        OPEN A FOLDER TO BEGIN
      </div>
      <div style={{fontSize:10,color:'#0a1628',lineHeight:1.8}}>
        Files are indexed <span style={{color:'#1d4ed8'}}>once</span>.
        Searches run at pure JS speed — no re-parsing.
      </div>
    </div>
    <div style={{display:'flex',gap:5}}>
      {['XLSX','XLS','CSV','TXT'].map(f=>(
        <span key={f} style={{padding:'2px 8px',borderRadius:20,background:'#08090d',
          border:'1px solid #0a1018',fontSize:8,color:'#0d1e30',fontWeight:700}}>{f}</span>
      ))}
    </div>
  </div>
));

// ── Indexing view ─────────────────────────────────────────────────────────────
const IndexView = memo(({files}:{files:FileMeta[]})=>{
  const done  = files.filter(f=>f.state==='done').length;
  const total = files.length;
  const pct   = total?Math.round(done/total*100):0;
  return (
    <div style={{height:'100%',display:'flex',flexDirection:'column',
      alignItems:'center',justifyContent:'center',gap:14,padding:32}}>
      <Loader2 size={26} color="#1d4ed8" style={{animation:'spin 0.8s linear infinite'}}/>
      <div style={{textAlign:'center'}}>
        <div style={{fontSize:11,fontWeight:700,color:'#1a2d47',
          letterSpacing:'0.08em',marginBottom:4}}>BUILDING SEARCH INDEX</div>
        <div style={{fontSize:10,color:'#0f1e30'}}>{done}/{total} files · {pct}%</div>
      </div>
      <div style={{width:180,height:3,background:'#09090e',borderRadius:2}}>
        <div style={{height:'100%',width:pct+'%',borderRadius:2,
          background:'linear-gradient(90deg,#1d4ed8,#4338ca)',transition:'width 0.4s'}}/>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:3,width:240}}>
        {files.slice(0,10).map(f=>(
          <div key={f.id} style={{display:'flex',alignItems:'center',gap:6,fontSize:9}}>
            {f.state==='done'
              ?<CheckCircle2 size={9} color="#22c55e"/>
              :f.state==='error'
              ?<AlertCircle  size={9} color="#ef4444"/>
              :f.state==='indexing'
              ?<Loader2 size={9} color="#3b82f6" style={{animation:'spin 0.8s linear infinite'}}/>
              :<div style={{width:9,height:9,borderRadius:'50%',background:'#0a1018'}}/>}
            <span style={{color:f.state==='done'?'#1a2d47':f.state==='error'?'#ef4444':'#0d1830',
              flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
              {f.name}
            </span>
            {f.state==='done'&&<span style={{color:'#0d1830',flexShrink:0}}>{fmt(f.rows)}</span>}
          </div>
        ))}
        {files.length>10&&<div style={{fontSize:9,color:'#09111e',textAlign:'center'}}>
          +{files.length-10} more</div>}
      </div>
    </div>
  );
});

// ── Ready view ────────────────────────────────────────────────────────────────
const ReadyView = memo(({n,rows}:{n:number;rows:number})=>(
  <div style={{height:'100%',display:'flex',flexDirection:'column',
    alignItems:'center',justifyContent:'center',gap:8,opacity:0.35}}>
    <Search size={30} color="#1d4ed8"/>
    <div style={{fontSize:10,color:'#1a2d47',letterSpacing:'0.1em'}}>
      {n} FILES · {fmt(rows)} ROWS · TYPE TO SEARCH
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
      re.test(p)?<mark key={i} style={{background:'#172a50',color:'#93c5fd',
        borderRadius:2,padding:'0 1px'}}>{p}</mark>:p
    )}</>;
  }catch{return <>{text}</>;}
});

// ── Result row ────────────────────────────────────────────────────────────────
const Row = memo(({h,open,data,loading,onOpen,q}:{
  h:Hit; open:boolean; data?:Record<string,string>|null; loading?:boolean;
  onOpen:(h:Hit)=>void; q:string;
})=>{
  const es=EXT[eOf(h.fn)]||EXT.xlsx;
  return (
    <div style={{borderBottom:'1px solid #08090e',background:open?'#080f1a':'transparent'}}>
      <div onClick={()=>onOpen(h)}
        style={{height:ROW_H,display:'flex',alignItems:'center',
          padding:'0 12px',cursor:'pointer',gap:9}}
        onMouseEnter={e=>{if(!open)(e.currentTarget as HTMLDivElement).style.background='#080c12';}}
        onMouseLeave={e=>{if(!open)(e.currentTarget as HTMLDivElement).style.background='transparent';}}>

        <div style={{width:28,height:28,borderRadius:6,flexShrink:0,
          background:es.d,border:'1px solid '+es.c+'20',
          display:'flex',alignItems:'center',justifyContent:'center',color:es.c}}>
          <es.I size={12}/>
        </div>

        <div style={{flex:1,minWidth:0}}>
          <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:2}}>
            <span style={{fontSize:10,fontWeight:700,color:'#2a3f5f',
              overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:220}}>
              {h.fn}
            </span>
            {h.sn&&<span style={{fontSize:8,color:'#2563eb',background:'#08142a',
              border:'1px solid #0d1e3a',borderRadius:3,padding:'1px 4px',flexShrink:0}}>
              {h.sn}</span>}
            <span style={{fontSize:8,color:'#0d1e30',flexShrink:0}}>#{h.rn}</span>
          </div>
          <div style={{fontSize:10,color:'#1a2d47',
            overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
            <Hl text={h.ss} q={q}/>
          </div>
        </div>

        {h.sc>1&&<div style={{flexShrink:0,padding:'1px 5px',borderRadius:4,
          background:'#071222',border:'1px solid #0c1e38',
          fontSize:8,color:'#1d4ed8',fontWeight:700}}>{h.sc}</div>}

        <ChevronDown size={11} color="#0d1a2a"
          style={{flexShrink:0,transition:'transform 0.15s',
            transform:open?'rotate(180deg)':'none'}}/>
      </div>

      {open&&(
        <div style={{padding:'0 12px 10px',borderTop:'1px solid #08090e'}}>
          {loading?(
            <div style={{display:'flex',alignItems:'center',gap:7,
              padding:'9px 0',color:'#1a2d47',fontSize:10}}>
              <Loader2 size={10} style={{animation:'spin 0.7s linear infinite'}}/> Loading…
            </div>
          ):!data||!Object.keys(data).length?(
            <div style={{display:'flex',alignItems:'center',gap:5,
              padding:'9px 0',color:'#0d1830',fontSize:10}}>
              <AlertCircle size={10}/> No data
            </div>
          ):(
            <div style={{display:'grid',
              gridTemplateColumns:'repeat(auto-fill,minmax(140px,1fr))',
              gap:4,paddingTop:7}}>
              {Object.entries(data).map(([k,v])=>(
                <div key={k} style={{padding:'4px 6px',borderRadius:4,
                  background:'#07090f',border:'1px solid #0a0f1a'}}>
                  <div style={{fontSize:7,color:'#0d1830',fontWeight:700,
                    letterSpacing:'0.1em',marginBottom:1}}>{k.toUpperCase()}</div>
                  <div style={{fontSize:10,color:'#1a2d47',wordBreak:'break-word'}}>
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

// ── Global CSS ────────────────────────────────────────────────────────────────
{
  const s=document.createElement('style');
  s.textContent=`
    @keyframes spin{to{transform:rotate(360deg)}}
    *{box-sizing:border-box;-webkit-font-smoothing:antialiased}
    ::-webkit-scrollbar{width:3px}
    ::-webkit-scrollbar-track{background:transparent}
    ::-webkit-scrollbar-thumb{background:#0a0e17;border-radius:3px}
    input::placeholder{color:#0d1830}
    button{font-family:inherit}
  `;
  document.head.appendChild(s);
}
