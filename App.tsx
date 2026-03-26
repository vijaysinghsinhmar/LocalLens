import React, { useEffect, useRef, useState, useMemo, useCallback, memo } from 'react';
import {
  Zap, FolderOpen, Search, X, Download,
  Loader2, CheckCircle2, AlertCircle, ChevronDown,
  FileSpreadsheet, FileCode, FileText
} from 'lucide-react';
import { workerScript } from './worker';

// ─────────────────────────────────────────────────────────────────────────────
//  Types
// ─────────────────────────────────────────────────────────────────────────────
interface FM {
  id:string; name:string; type:string; blob:File; size:number;
  st:'q'|'ing'|'ok'|'err'; rows:number; err?:string;
}
interface Hit {
  id:string; fid:string; fn:string; sn:string|null;
  rn:number; ss:string; sc:number; ri:number;
}

// ─────────────────────────────────────────────────────────────────────────────
//  Constants / helpers
// ─────────────────────────────────────────────────────────────────────────────
const NW       = 2;          // workers
const ROW_H    = 58;
const OVER     = 12;
const DEB      = 200;
const MAX_HITS = 20000;
const SUPP     = new Set(['xlsx','xls','csv','txt']);

const fmtN = (n:number) => n>=1e6?(n/1e6).toFixed(1)+'M':n>=1e3?(n/1e3).toFixed(1)+'K':''+n;
const fmtB = (b:number) => b>=(1<<20)?(b/(1<<20)).toFixed(1)+' MB':b>=(1<<10)?(b/(1<<10)).toFixed(1)+' KB':b+' B';
const fmtT = (ms:number) => ms>=1000?(ms/1000).toFixed(2)+'s':ms+'ms';

type Ext='xlsx'|'xls'|'csv'|'txt';
const ES:Record<Ext,{c:string;bg:string;I:any}> = {
  xlsx:{c:'#60a5fa',bg:'#0c1e3a',I:FileSpreadsheet},
  xls: {c:'#818cf8',bg:'#150d35',I:FileSpreadsheet},
  csv: {c:'#34d399',bg:'#011a0e',I:FileCode},
  txt: {c:'#fbbf24',bg:'#120c02',I:FileText},
};
const eOf = (n:string)=>(n.split('.').pop()?.toLowerCase()||'xlsx') as Ext;

// ─────────────────────────────────────────────────────────────────────────────
//  App
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  const [files,   setFiles]   = useState<FM[]>([]);
  const [totR,    setTotR]    = useState(0);
  const [query,   setQuery]   = useState('');
  const [hits,    setHits]    = useState<Hit[]>([]);
  const [busy,    setBusy]    = useState(false);
  const [ms,      setMs]      = useState<number|null>(null);
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

  // Refs — never re-render
  const WW      = useRef<Worker[]>([]);
  const iQ      = useRef<FM[]>([]);       // index queue
  const iDone   = useRef(0);
  const iTotal  = useRef(0);
  const sid     = useRef(0);
  const sDone   = useRef(0);
  const sTotal  = useRef(0);
  const sStart  = useRef(0);
  const sBuf    = useRef<Hit[]>([]);
  const allHits = useRef<Hit[]>([]);
  const fRef    = useRef<FM[]>([]);       // mirror of files
  const debT    = useRef(0);
  const tickT   = useRef(0);

  // Resize
  useEffect(()=>{
    const ro=new ResizeObserver(()=>{ if(scEl.current) setVH(scEl.current.clientHeight); });
    if(scEl.current) ro.observe(scEl.current);
    return ()=>ro.disconnect();
  },[]);

  // ── Boot workers ONCE ─────────────────────────────────────────────────
  useEffect(()=>{
    const blob = new Blob([workerScript],{type:'application/javascript'});
    const url  = URL.createObjectURL(blob);

    function dispatchIdx(wi:number) {
      if (!iQ.current.length) return;
      const f = iQ.current.shift()!;
      // mark indexing
      setFiles(prev=>{
        const nx=prev.map(x=>x.id===f.id?{...x,st:'ing' as const}:x);
        fRef.current=nx; return nx;
      });
      WW.current[wi].postMessage({t:'idx',id:f.id,blob:f.blob,ft:f.type,name:f.name});
    }

    function flushHits() {
      if (!sBuf.current.length) return;
      const snap=sBuf.current.splice(0);
      allHits.current.push(...snap);
      if (allHits.current.length<=MAX_HITS)
        setHits(allHits.current.slice());
    }

    WW.current = Array.from({length:NW},(_,wi)=>{
      const w = new Worker(url);
      w.onmessage=(ev)=>{
        const d=ev.data;

        if (d.t==='ok') {
          iDone.current++;
          setTotR(r=>r+d.n);
          setFiles(prev=>{
            const nx=prev.map(x=>x.id===d.id?{...x,st:'ok' as const,rows:d.n}:x);
            fRef.current=nx; return nx;
          });
          dispatchIdx(wi);

        } else if (d.t==='err') {
          iDone.current++;
          setFiles(prev=>{
            const nx=prev.map(x=>x.id===d.id?{...x,st:'err' as const,err:d.msg}:x);
            fRef.current=nx; return nx;
          });
          dispatchIdx(wi);

        } else if (d.t==='hits') {
          if (d.sid!==sid.current) return;
          sBuf.current.push(...d.hits);
          if (sBuf.current.length>400) flushHits();

        } else if (d.t==='done') {
          if (d.sid!==sid.current) return;
          sDone.current++;
          if (sDone.current>=sTotal.current) {
            flushHits();
            setBusy(false);
            clearInterval(tickT.current);
            setMs(Date.now()-sStart.current);
          }

        } else if (d.t==='row') {
          setExpD(d.d); setExpLoad(false);

        } else if (d.t==='csv') {
          const a=document.createElement('a');
          a.href=URL.createObjectURL(d.blob);
          a.download='locallens_'+Date.now()+'.csv';
          a.click();
        }
      };
      return w;
    });

    return ()=>{
      WW.current.forEach(w=>w.terminate());
      URL.revokeObjectURL(url);
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  // ── Load folder ───────────────────────────────────────────────────────
  const loadFolder=useCallback((e:React.ChangeEvent<HTMLInputElement>)=>{
    const raw=e.target.files; if(!raw) return;
    WW.current.forEach(w=>w.postMessage({t:'clear'}));
    iQ.current=[];
    iDone.current=0; iTotal.current=0;
    allHits.current=[]; sBuf.current=[];
    setHits([]); setQuery(''); setMs(null); setTotR(0);
    setExpId(null); setExpD(null);

    const list:FM[]=[];
    for (let i=0;i<raw.length;i++) {
      const f=raw[i];
      const ext=(f.name.split('.').pop()||'').toLowerCase();
      if (!SUPP.has(ext)) continue;
      list.push({
        id:`${f.name}-${f.size}-${f.lastModified}`,
        name:f.name,type:ext,blob:f,size:f.size,st:'q',rows:0
      });
    }
    if (!list.length) return;

    fRef.current=list;
    setFiles(list);
    iQ.current=list.slice();
    iTotal.current=list.length;

    // Seed workers
    for (let wi=0;wi<NW&&iQ.current.length;wi++) {
      const f=iQ.current.shift()!;
      setFiles(prev=>{
        const nx=prev.map(x=>x.id===f.id?{...x,st:'ing' as const}:x);
        fRef.current=nx; return nx;
      });
      WW.current[wi].postMessage({t:'idx',id:f.id,blob:f.blob,ft:f.type,name:f.name});
    }
  },[]);

  // ── Search ────────────────────────────────────────────────────────────
  const runSearch=useCallback((q:string,fz:boolean,ex:boolean,ax:string[])=>{
    if (!q.trim()) { setHits([]); allHits.current=[]; setMs(null); return; }
    const targets=fRef.current.filter(f=>f.st==='ok'&&ax.includes(f.type));
    if (!targets.length) return;

    const newSid=++sid.current;
    sBuf.current=[]; allHits.current=[];
    sDone.current=0; sTotal.current=targets.length;
    sStart.current=Date.now();
    setHits([]); setMs(null); setBusy(true);
    setExpId(null); setExpD(null);

    clearInterval(tickT.current);
    tickT.current=setInterval(()=>setMs(Date.now()-sStart.current),120) as unknown as number;

    targets.forEach((f,i)=>{
      WW.current[i%NW].postMessage({t:'search',sid:newSid,id:f.id,q:q,fuzzy:fz,exact:ex});
    });
  },[]);

  const trig=useCallback((q:string)=>{
    clearTimeout(debT.current);
    debT.current=setTimeout(()=>runSearch(q,fuzzy,exact,exts),DEB) as unknown as number;
  },[runSearch,fuzzy,exact,exts]);

  useEffect(()=>{
    if (query.trim()&&fRef.current.some(f=>f.st==='ok')) runSearch(query,fuzzy,exact,exts);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[fuzzy,exact,exts]);

  // ── Expand ────────────────────────────────────────────────────────────
  const expand=useCallback((h:Hit)=>{
    if (expId===h.id){setExpId(null);setExpD(null);return;}
    setExpId(h.id); setExpD(null); setExpLoad(true);
    WW.current[0].postMessage({t:'row',id:h.fid,ri:h.ri});
  },[expId]);

  // ── Sort + virtual slice ──────────────────────────────────────────────
  const sorted=useMemo(()=>{
    const a=hits.slice();
    if      (sortK==='sc') a.sort((x,y)=>sortD*(y.sc-x.sc));
    else if (sortK==='fn') a.sort((x,y)=>sortD*x.fn.localeCompare(y.fn));
    else                   a.sort((x,y)=>sortD*(x.rn-y.rn));
    return a;
  },[hits,sortK,sortD]);

  const {vis,vs}=useMemo(()=>{
    const s=Math.max(0,Math.floor(sTop/ROW_H)-OVER);
    const e=Math.min(sorted.length,Math.ceil((sTop+vH)/ROW_H)+OVER);
    return {vis:sorted.slice(s,e),vs:s};
  },[sorted,sTop,vH]);

  // ── Derived ───────────────────────────────────────────────────────────
  const nOk  = files.filter(f=>f.st==='ok').length;
  const nPend= files.filter(f=>f.st!=='ok'&&f.st!=='err').length;
  const isIdx= nPend>0;
  const ready= !isIdx&&nOk>0;
  const totSz= files.reduce((s,f)=>s+f.size,0);
  const pct  = files.length?Math.round(nOk/files.length*100):0;
  const capped=allHits.current.length>=MAX_HITS;

  return (
    <div style={{display:'flex',flexDirection:'column',height:'100vh',overflow:'hidden',
      background:'#05060a',color:'#8899aa',fontFamily:"'DM Mono','Cascadia Code','Fira Code',monospace"}}>

      {/* ═══ HEADER ══════════════════════════════════════════════════════ */}
      <div style={{flexShrink:0,background:'#070810',borderBottom:'1px solid #0c1018'}}>

        {/* Row 1: controls */}
        <div style={{display:'flex',alignItems:'center',gap:7,
          padding:'8px 12px',flexWrap:'wrap'}}>

          {/* Logo */}
          <div style={{display:'flex',alignItems:'center',gap:7,flexShrink:0,marginRight:4}}>
            <div style={{width:27,height:27,borderRadius:7,flexShrink:0,
              background:'linear-gradient(135deg,#1d4ed8,#4338ca)',
              display:'flex',alignItems:'center',justifyContent:'center',
              boxShadow:'0 0 10px #1d4ed840'}}>
              <Zap size={14} color="#fff"/>
            </div>
            <span style={{fontSize:12,fontWeight:700,letterSpacing:'0.07em',color:'#dde6f0',userSelect:'none'}}>
              LOCAL<span style={{color:'#3b82f6'}}>LENS</span>
            </span>
          </div>

          {/* Search */}
          <div style={{flex:1,minWidth:160,position:'relative'}}>
            <Search size={13} style={{position:'absolute',left:9,top:'50%',
              transform:'translateY(-50%)',pointerEvents:'none',
              color:busy?'#3b82f6':'#1c2e44',transition:'color 0.15s'}}/>
            <input value={query} disabled={!ready}
              onChange={e=>{setQuery(e.target.value);trig(e.target.value);}}
              placeholder={
                isIdx?`Indexing… ${nOk}/${files.length}`:
                !files.length?'Open a folder':
                !ready?'Waiting for index…':
                `Search ${fmtN(totR)} rows across ${files.length} files`
              }
              style={{width:'100%',boxSizing:'border-box',
                padding:'7px 30px 7px 28px',
                background:'#0c0f18',border:'1px solid #121929',
                borderRadius:7,color:'#dde6f0',fontSize:13,outline:'none',
                fontFamily:'inherit',opacity:ready?1:0.4,
                transition:'border-color 0.15s,opacity 0.3s'}}
              onFocus={e=>e.target.style.borderColor='#1d4ed8'}
              onBlur={e=>e.target.style.borderColor='#121929'}
            />
            {query&&<button onClick={()=>{setQuery('');setHits([]);allHits.current=[];setMs(null);}}
              style={{position:'absolute',right:7,top:'50%',transform:'translateY(-50%)',
                background:'none',border:'none',cursor:'pointer',color:'#2a3a50',padding:0,lineHeight:0}}>
              <X size={12}/></button>}
            {busy&&<Loader2 size={12} color="#3b82f6"
              style={{position:'absolute',right:query?24:7,top:'50%',
                transform:'translateY(-50%)',animation:'spin 0.7s linear infinite'}}/>}
          </div>

          {/* Mode toggles */}
          <Tog a={fuzzy}  f={()=>setFuzzy(!fuzzy)}  l="FUZZY"/>
          <Tog a={exact}  f={()=>setExact(!exact)}  l="EXACT"/>
          <Div/>

          {/* File type filters */}
          {(['xlsx','xls','csv','txt'] as Ext[]).map(t=>(
            <Tog key={t} a={exts.includes(t)} c={ES[t].c}
              f={()=>setExts(p=>p.includes(t)?p.filter(x=>x!==t):[...p,t])}
              l={t.toUpperCase()}/>
          ))}
          <Div/>

          {/* Sort */}
          {([['sc','SCORE'],['fn','FILE'],['rn','ROW']] as ['sc'|'fn'|'rn',string][]).map(([k,l])=>(
            <button key={k} onClick={()=>{
              if(sortK===k) setSortD(d=>d===1?-1:1);
              else {setSortK(k);setSortD(-1);}
            }} style={{padding:'3px 7px',borderRadius:5,cursor:'pointer',flexShrink:0,
              background:sortK===k?'#0a1425':'transparent',
              border:'1px solid '+(sortK===k?'#0f1e38':'transparent'),
              color:sortK===k?'#3b82f6':'#1a2d3f',fontSize:9,fontWeight:700}}>
              {l}{sortK===k?(sortD===-1?' ↓':' ↑'):''}
            </button>
          ))}
          <Div/>

          {/* Folder */}
          <label style={{display:'flex',alignItems:'center',gap:5,padding:'5px 10px',
            borderRadius:6,cursor:'pointer',background:'#0c0f18',
            border:'1px solid #121929',fontSize:10,fontWeight:700,
            color:'#2a3a50',whiteSpace:'nowrap',flexShrink:0}}>
            <FolderOpen size={12}/>
            {files.length?'CHANGE':'OPEN FOLDER'}
            <input type="file" style={{display:'none'}}
              // @ts-ignore
              webkitdirectory="" directory="" multiple onChange={loadFolder}/>
          </label>

          {/* Export */}
          {allHits.current.length>0&&(
            <button onClick={()=>WW.current[0].postMessage({t:'export',hits:allHits.current})}
              style={{display:'flex',alignItems:'center',gap:5,padding:'5px 9px',
                borderRadius:6,border:'none',cursor:'pointer',flexShrink:0,
                background:'linear-gradient(135deg,#1d4ed8,#4338ca)',
                fontSize:10,fontWeight:700,color:'#fff'}}>
              <Download size={11}/> CSV
            </button>
          )}
        </div>

        {/* Row 2: status */}
        <div style={{display:'flex',alignItems:'center',gap:10,
          padding:'3px 12px 5px',borderTop:'1px solid #080b12',fontSize:9}}>

          {files.length>0&&<>
            <span style={{display:'flex',alignItems:'center',gap:4,fontWeight:700,
              color:isIdx?'#f59e0b':'#22c55e'}}>
              {isIdx
                ?<Loader2 size={8} style={{animation:'spin 0.7s linear infinite'}}/>
                :<CheckCircle2 size={8}/>}
              {isIdx?`${nOk}/${files.length} indexed`:`${nOk} files`}
            </span>
            <span style={{color:'#1c2e44'}}>{fmtN(totR)} rows</span>
            <span style={{color:'#1c2e44'}}>{fmtB(totSz)}</span>
            {isIdx&&<>
              <div style={{flex:1,height:2,background:'#090c14',borderRadius:1,minWidth:30}}>
                <div style={{height:'100%',width:pct+'%',
                  background:'#1d4ed8',borderRadius:1,transition:'width 0.35s'}}/>
              </div>
              <span style={{color:'#1d4ed8',fontWeight:700}}>{pct}%</span>
            </>}
          </>}

          <div style={{marginLeft:'auto',display:'flex',gap:10,alignItems:'center'}}>
            {hits.length>0&&<span style={{fontWeight:700,
              color:capped?'#f59e0b':'#3b82f6'}}>
              {capped?fmtN(MAX_HITS)+'+':fmtN(hits.length)} results
            </span>}
            {ms!==null&&<span style={{color:busy?'#f59e0b':'#1a2d3f'}}>
              {fmtT(ms)}{busy?'…':''}
            </span>}
          </div>
        </div>
      </div>

      {/* ═══ BODY ════════════════════════════════════════════════════════ */}
      <div ref={scEl} onScroll={e=>setSTop(e.currentTarget.scrollTop)}
        style={{flex:1,overflowY:'auto',overflowX:'hidden'}}>

        {files.length===0&&<EmptyScreen/>}
        {files.length>0&&isIdx&&hits.length===0&&!query&&<IdxScreen files={files}/>}
        {ready&&!query&&hits.length===0&&<ReadyScreen n={files.length} rows={totR}/>}
        {ready&&query&&hits.length===0&&!busy&&
          <Center><Search size={28} color="#1c2e44"/><div style={{fontSize:11,color:'#1c2e44',marginTop:8}}>No results for "{query}"</div></Center>}

        {sorted.length>0&&(
          <div style={{position:'relative',height:sorted.length*ROW_H}}>
            <div style={{position:'absolute',top:0,left:0,right:0,
              transform:`translateY(${vs*ROW_H}px)`}}>
              {vis.map(h=>(
                <ResultRow key={h.id} h={h}
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

// ── Small atoms ───────────────────────────────────────────────────────────────
const Div=()=><div style={{width:1,height:14,background:'#0c1018',flexShrink:0}}/>;

const Tog=memo(({a,f,l,c}:{a:boolean;f:()=>void;l:string;c?:string})=>(
  <button onClick={f} style={{padding:'3px 7px',borderRadius:5,cursor:'pointer',flexShrink:0,
    background:a?(c?c+'12':'#0c1d38'):'transparent',
    border:'1px solid '+(a?(c||'#1d4ed8'):'#0c1018'),
    color:a?(c||'#60a5fa'):'#182333',
    fontSize:9,fontWeight:800,letterSpacing:'0.07em',transition:'all 0.1s'}}>
    {l}
  </button>
));

const Center=({children}:{children:React.ReactNode})=>(
  <div style={{height:'100%',display:'flex',flexDirection:'column',
    alignItems:'center',justifyContent:'center',gap:4,opacity:0.5}}>
    {children}
  </div>
);

const EmptyScreen=memo(()=>(
  <Center>
    <div style={{width:56,height:56,borderRadius:14,background:'#090c14',
      border:'1px solid #0c1220',display:'flex',alignItems:'center',justifyContent:'center',marginBottom:8}}>
      <FolderOpen size={22} color="#0d1e2e"/>
    </div>
    <div style={{fontSize:13,fontWeight:700,color:'#0e1d2e',letterSpacing:'0.07em',marginBottom:4}}>
      OPEN A FOLDER TO START
    </div>
    <div style={{fontSize:10,color:'#091422',textAlign:'center',lineHeight:1.8,maxWidth:300}}>
      Files index <span style={{color:'#1d4ed8'}}>once</span>.
      Every search is pure JS — no re-parsing, no network.
    </div>
    <div style={{display:'flex',gap:5,marginTop:8}}>
      {['XLSX','XLS','CSV','TXT'].map(x=>(
        <span key={x} style={{padding:'2px 7px',borderRadius:20,background:'#080a12',
          border:'1px solid #0a0e18',fontSize:8,color:'#0d1a28',fontWeight:700}}>{x}</span>
      ))}
    </div>
  </Center>
));

const IdxScreen=memo(({files}:{files:FM[]})=>{
  const done=files.filter(f=>f.st==='ok').length;
  const pct=files.length?Math.round(done/files.length*100):0;
  return (
    <Center>
      <Loader2 size={24} color="#1d4ed8" style={{animation:'spin 0.9s linear infinite',marginBottom:8}}/>
      <div style={{fontSize:11,fontWeight:700,color:'#1a2d40',letterSpacing:'0.07em'}}>
        INDEXING {done}/{files.length} · {pct}%
      </div>
      <div style={{width:160,height:2,background:'#08090e',borderRadius:1,margin:'8px 0'}}>
        <div style={{height:'100%',width:pct+'%',background:'#1d4ed8',
          borderRadius:1,transition:'width 0.4s'}}/>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:3,width:220}}>
        {files.slice(0,8).map(f=>(
          <div key={f.id} style={{display:'flex',alignItems:'center',gap:6,fontSize:9}}>
            {f.st==='ok'?<CheckCircle2 size={9} color="#22c55e"/>
             :f.st==='err'?<AlertCircle size={9} color="#ef4444"/>
             :f.st==='ing'?<Loader2 size={9} color="#3b82f6" style={{animation:'spin 0.8s linear infinite'}}/>
             :<div style={{width:9,height:9,borderRadius:'50%',background:'#0a0e18'}}/>}
            <span style={{flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',
              color:f.st==='ok'?'#1a2d40':f.st==='err'?'#ef4444':'#0e1d2e'}}>{f.name}</span>
            {f.st==='ok'&&<span style={{color:'#0e1d2e',flexShrink:0}}>{fmtN(f.rows)}</span>}
          </div>
        ))}
        {files.length>8&&<div style={{fontSize:9,color:'#08111e',textAlign:'center'}}>
          +{files.length-8} more files</div>}
      </div>
    </Center>
  );
});

const ReadyScreen=memo(({n,rows}:{n:number;rows:number})=>(
  <Center>
    <Search size={28} color="#1d4ed8"/>
    <div style={{fontSize:10,color:'#1a2d40',letterSpacing:'0.08em',marginTop:6}}>
      {n} FILES · {fmtN(rows)} ROWS INDEXED · TYPE TO SEARCH
    </div>
  </Center>
));

// ── Highlight ─────────────────────────────────────────────────────────────────
const Hl=memo(({text,q}:{text:string;q:string})=>{
  if (!q.trim()||!text) return <>{text}</>;
  const terms=q.trim().split(/\s+/).filter(t=>t.length>1);
  if (!terms.length) return <>{text}</>;
  try {
    const re=new RegExp(`(${terms.map(t=>t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|')})`, 'gi');
    return <>{text.split(re).map((p,i)=>
      re.test(p)?<mark key={i} style={{background:'#152040',color:'#93c5fd',borderRadius:2,padding:'0 1px'}}>{p}</mark>:p
    )}</>;
  } catch { return <>{text}</>; }
});

// ── Result row ────────────────────────────────────────────────────────────────
const ResultRow=memo(({h,open,data,loading,onOpen,q}:{
  h:Hit;open:boolean;data?:Record<string,string>|null;
  loading?:boolean;onOpen:(h:Hit)=>void;q:string;
})=>{
  const es=ES[eOf(h.fn)]||ES.xlsx;
  return (
    <div style={{borderBottom:'1px solid #07080d',background:open?'#070d18':'transparent'}}>
      <div onClick={()=>onOpen(h)}
        style={{height:ROW_H,display:'flex',alignItems:'center',
          padding:'0 12px',cursor:'pointer',gap:9}}
        onMouseEnter={e=>{if(!open)(e.currentTarget as HTMLDivElement).style.background='#070a10';}}
        onMouseLeave={e=>{if(!open)(e.currentTarget as HTMLDivElement).style.background='transparent';}}>

        <div style={{width:26,height:26,borderRadius:6,flexShrink:0,
          background:es.bg,border:'1px solid '+es.c+'18',
          display:'flex',alignItems:'center',justifyContent:'center',color:es.c}}>
          <es.I size={12}/>
        </div>

        <div style={{flex:1,minWidth:0}}>
          <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:2}}>
            <span style={{fontSize:10,fontWeight:700,color:'#243648',
              overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:240}}>
              {h.fn}
            </span>
            {h.sn&&<span style={{fontSize:7,color:'#1d4ed8',background:'#07122a',
              border:'1px solid #0c1e3a',borderRadius:3,padding:'1px 4px',flexShrink:0}}>
              {h.sn}</span>}
            <span style={{fontSize:8,color:'#0c1828',flexShrink:0}}>#{h.rn}</span>
          </div>
          <div style={{fontSize:10,color:'#182838',
            overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
            <Hl text={h.ss} q={q}/>
          </div>
        </div>

        {h.sc>1&&<div style={{flexShrink:0,padding:'1px 4px',borderRadius:4,
          background:'#061020',border:'1px solid #0a1830',
          fontSize:8,color:'#1d4ed8',fontWeight:700}}>{h.sc}</div>}

        <ChevronDown size={11} color="#0c1828"
          style={{flexShrink:0,transition:'transform 0.15s',
            transform:open?'rotate(180deg)':'none'}}/>
      </div>

      {open&&(
        <div style={{padding:'2px 12px 10px',borderTop:'1px solid #07080d'}}>
          {loading?(
            <div style={{display:'flex',alignItems:'center',gap:7,
              padding:'8px 0',color:'#182838',fontSize:10}}>
              <Loader2 size={10} style={{animation:'spin 0.7s linear infinite'}}/> Loading…
            </div>
          ):!data||!Object.keys(data).length?(
            <div style={{display:'flex',alignItems:'center',gap:5,
              padding:'8px 0',color:'#0c1828',fontSize:10}}>
              <AlertCircle size={10}/> No data
            </div>
          ):(
            <div style={{display:'grid',
              gridTemplateColumns:'repeat(auto-fill,minmax(140px,1fr))',
              gap:4,paddingTop:6}}>
              {Object.entries(data).map(([k,v])=>(
                <div key={k} style={{padding:'4px 6px',borderRadius:4,
                  background:'#060810',border:'1px solid #090c18'}}>
                  <div style={{fontSize:7,color:'#0c1828',fontWeight:700,
                    letterSpacing:'0.1em',marginBottom:1}}>{k.toUpperCase()}</div>
                  <div style={{fontSize:10,color:'#182838',wordBreak:'break-word'}}>
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

// ── CSS ───────────────────────────────────────────────────────────────────────
{
  const s=document.createElement('style');
  s.textContent=`
    @keyframes spin{to{transform:rotate(360deg)}}
    *{box-sizing:border-box;-webkit-font-smoothing:antialiased}
    ::-webkit-scrollbar{width:3px}
    ::-webkit-scrollbar-track{background:transparent}
    ::-webkit-scrollbar-thumb{background:#090c14;border-radius:3px}
    input::placeholder{color:#0d1a28}
    button,input{font-family:inherit}
  `;
  document.head.appendChild(s);
}
