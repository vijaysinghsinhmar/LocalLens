
import React, { useState, useCallback, useEffect, useRef, useMemo } from 'react';
import { 
  Search, 
  FolderOpen, 
  ShieldCheck, 
  Loader2, 
  Clock, 
  Database,
  Download,
  Zap,
  AlertTriangle,
  ChevronDown,
  CheckCircle2,
  FileSpreadsheet,
  FileCode,
  Sparkles,
  FileWarning
} from 'lucide-react';
import { SearchResult, RecentSearch, SortConfig } from './types';
import { workerScript } from './worker';

const MAX_WORKERS = Math.min(navigator.hardwareConcurrency || 4, 4);
const ROW_HEIGHT = 80; 
const VIRTUAL_BUFFER = 12;
const BATCH_FLUSH_INTERVAL = 150; 

interface FileMetadata {
  id: string;
  name: string;
  path: string;
  type: string;
  blob: File;
  size: number;
}

const App: React.FC = () => {
  const [query, setQuery] = useState('');
  const [allResults, setAllResults] = useState<SearchResult[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exactMatch, setExactMatch] = useState(false);
  const [fuzzySearch, setFuzzySearch] = useState(true);
  const [selectedFileTypes, setSelectedFileTypes] = useState<string[]>(['xlsx', 'xls', 'csv', 'txt']);
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const [expandedData, setExpandedData] = useState<Record<string, any> | null>(null);
  const [showSummary, setShowSummary] = useState(false);
  
  const [files, setFiles] = useState<FileMetadata[]>([]);
  const [searchProgress, setSearchProgress] = useState({ scanned: 0, total: 0, errors: [] as string[] });
  
  const [scrollTop, setScrollTop] = useState(0);
  const [containerHeight, setContainerHeight] = useState(0);
  
  const workers = useRef<Worker[]>([]);
  const searchTimeout = useRef<number | null>(null);
  const activeSearchId = useRef<number>(0);
  const startTime = useRef<number>(0);
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const resultsBuffer = useRef<SearchResult[]>([]);
  const flushTimeout = useRef<number | null>(null);

  useEffect(() => {
    const handleResize = () => {
      if (scrollContainerRef.current) setContainerHeight(scrollContainerRef.current.clientHeight);
    };
    handleResize();
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const flushResults = useCallback(() => {
    if (resultsBuffer.current.length > 0) {
      const newItems = [...resultsBuffer.current];
      resultsBuffer.current = [];
      setAllResults(prev => [...prev, ...newItems]);
    }
    flushTimeout.current = null;
  }, []);

  const queueResults = useCallback((items: SearchResult[]) => {
    resultsBuffer.current.push(...items);
    if (!flushTimeout.current) flushTimeout.current = window.setTimeout(flushResults, BATCH_FLUSH_INTERVAL);
  }, [flushResults]);

  useEffect(() => {
    const blob = new Blob([workerScript], { type: 'application/javascript' });
    const url = URL.createObjectURL(blob);
    for (let i = 0; i < MAX_WORKERS; i++) {
      const w = new Worker(url);
      w.onmessage = (e) => {
        const { action, payload } = e.data;
        if (action === 'MATCH_CHUNK') queueResults(payload.matches);
        if (action === 'ROW_DETAIL_RESULT') setExpandedData(payload.rowData);
        if (action === 'EXPORT_READY') {
          const link = document.createElement("a");
          link.href = URL.createObjectURL(payload.blob);
          link.download = `xsearch_export_${Date.now()}.csv`;
          link.click();
          setIsExporting(false);
        }
      };
      workers.current.push(w);
    }
    return () => workers.current.forEach(w => w.terminate());
  }, [queueResults]);

  const handleFolderSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const rawFiles = e.target.files;
    if (!rawFiles) return;
    const fileList: FileMetadata[] = [];
    for (let i = 0; i < rawFiles.length; i++) {
      const file = rawFiles[i];
      const extension = file.name.split('.').pop()?.toLowerCase() || '';
      if (['xlsx', 'xls', 'csv', 'txt'].includes(extension)) {
        fileList.push({
          id: `${file.name}-${file.size}-${file.lastModified}`,
          name: file.name,
          path: (file as any).webkitRelativePath || file.name,
          type: extension,
          blob: file,
          size: file.size
        });
      }
    }
    setFiles(fileList);
  };

  const performSearch = useCallback(async (searchTerm: string, exact: boolean, fuzzy: boolean, types: string[]) => {
    const searchId = ++activeSearchId.current;
    if (flushTimeout.current) window.clearTimeout(flushTimeout.current);
    flushTimeout.current = null;
    resultsBuffer.current = [];
    setIsSearching(true);
    setAllResults([]);
    setShowSummary(false);
    startTime.current = Date.now();
    
    const filteredFiles = files.filter(f => types.includes(f.type));
    setSearchProgress({ scanned: 0, total: filteredFiles.length, errors: [] });

    let currentIdx = 0;
    let completedCount = 0;

    const startWorker = (worker: Worker) => {
      if (currentIdx >= filteredFiles.length || searchId !== activeSearchId.current) return;
      const file = filteredFiles[currentIdx++];
      worker.postMessage({
        action: 'PROCESS_FILE',
        payload: { fileId: file.id, blob: file.blob, type: file.type, name: file.name, path: file.path, query: searchTerm, exactMatch: exact, fuzzy }
      });
      
      const oldMsg = worker.onmessage;
      worker.onmessage = (e) => {
        const { action, payload } = e.data;
        if (action === 'FILE_COMPLETE' || action === 'FILE_ERROR') {
          completedCount++;
          setSearchProgress(p => ({ ...p, scanned: completedCount, errors: action === 'FILE_ERROR' ? [...p.errors, payload.error] : p.errors }));
          if (completedCount === filteredFiles.length) {
            flushResults();
            setIsSearching(false);
            setShowSummary(true);
          } else {
            startWorker(worker);
          }
        }
        // Fix for Error: The 'this' context of type 'void' is not assignable to method's 'this' of type 'Worker'
        // Using .call(worker, e) explicitly provides the Worker instance as the 'this' context for the original message handler.
        if (oldMsg) oldMsg.call(worker, e);
      };
    };
    workers.current.forEach(w => startWorker(w));
  }, [files, flushResults]);

  const onQueryChange = (val: string) => {
    setQuery(val);
    if (searchTimeout.current) window.clearTimeout(searchTimeout.current);
    searchTimeout.current = window.setTimeout(() => performSearch(val, exactMatch, fuzzySearch, selectedFileTypes), 400);
  };

  const handleExpand = (res: SearchResult) => {
    if (expandedId === res.id) {
      setExpandedId(null);
      setExpandedData(null);
    } else {
      setExpandedId(res.id);
      setExpandedData(null);
      workers.current[0].postMessage({
        action: 'FETCH_ROW_DETAIL',
        payload: { fileId: res.fileId, rowNumber: res.rowNumber, sheetName: res.sheetName }
      });
    }
  };

  const startExport = () => {
    setIsExporting(true);
    workers.current[0].postMessage({ action: 'GENERATE_EXPORT', payload: {} });
  };

  const visibleItems = useMemo(() => {
    const start = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - VIRTUAL_BUFFER);
    const end = Math.min(allResults.length, Math.ceil((scrollTop + containerHeight) / ROW_HEIGHT) + VIRTUAL_BUFFER);
    return { items: allResults.slice(start, end), startIdx: start };
  }, [allResults, scrollTop, containerHeight]);

  return (
    <div className="flex flex-col h-screen bg-[#0a0a0c] text-gray-200 overflow-hidden font-sans">
      <header className="px-6 py-4 bg-[#111114] border-b border-gray-800 flex justify-between items-center z-50">
        <div className="flex items-center gap-4">
          <div className="bg-blue-600 p-2 rounded-xl">
            <Zap className="w-5 h-5 text-white" />
          </div>
          <h1 className="text-xl font-black uppercase tracking-tight">X-SEARCH <span className="text-blue-500">PRO</span></h1>
        </div>
        
        <div className="flex items-center gap-4">
          <label className="flex items-center gap-3 cursor-pointer bg-[#1a1a1e] px-5 py-2 rounded-xl text-xs font-bold border border-gray-800">
            <FolderOpen className="w-4 h-4 text-gray-400" />
            {files.length || 'Select Directory'}
            <input type="file" className="hidden" webkitdirectory="" directory="" multiple onChange={handleFolderSelect} />
          </label>
          
          {allResults.length > 0 && (
            <button 
              onClick={startExport}
              disabled={isExporting}
              className="flex items-center gap-2 px-5 py-2 bg-blue-600 text-white rounded-xl text-xs font-bold disabled:opacity-50"
            >
              {isExporting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Download className="w-4 h-4" />}
              Export CSV
            </button>
          )}
        </div>
      </header>

      <main className="flex-1 flex flex-col bg-[#0d0d0f]">
        <div className="px-6 py-6 bg-[#111114] border-b border-gray-800">
          <div className="max-w-6xl mx-auto space-y-4">
            <div className="relative">
              <Search className={`absolute left-5 top-1/2 -translate-y-1/2 w-5 h-5 ${isSearching ? 'text-blue-500' : 'text-gray-500'}`} />
              <input 
                type="text"
                placeholder="Deep search local records..."
                className="w-full bg-[#1a1a1e] border border-gray-800 rounded-2xl py-4 pl-14 pr-6 text-lg"
                disabled={files.length === 0}
                value={query}
                onChange={(e) => onQueryChange(e.target.value)}
              />
              {isSearching && (
                <div className="absolute right-5 top-1/2 -translate-y-1/2 flex items-center gap-3">
                  <span className="text-[10px] text-blue-500 font-bold uppercase">{Math.round((searchProgress.scanned / (searchProgress.total || 1)) * 100)}%</span>
                  <Loader2 className="w-4 h-4 animate-spin text-blue-500" />
                </div>
              )}
            </div>

            <div className="flex items-center gap-6">
              <button onClick={() => setFuzzySearch(!fuzzySearch)} className={`flex items-center gap-2 text-[10px] font-black tracking-widest ${fuzzySearch ? 'text-blue-400' : 'text-gray-600'}`}>
                <Sparkles className="w-4 h-4" /> FUZZY ENGINE
              </button>
              <div className="h-3 w-px bg-gray-800" />
              <button onClick={() => setExactMatch(!exactMatch)} className={`flex items-center gap-2 text-[10px] font-black tracking-widest ${exactMatch ? 'text-blue-400' : 'text-gray-600'}`}>
                <Database className="w-4 h-4" /> EXACT MATCH
              </button>
            </div>
          </div>
        </div>

        <div ref={scrollContainerRef} onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)} className="flex-1 overflow-auto bg-[#0a0a0c]">
          {files.length > 0 ? (
            <div className="max-w-6xl mx-auto py-6" style={{ height: allResults.length * ROW_HEIGHT }}>
              <div style={{ transform: `translateY(${visibleItems.startIdx * ROW_HEIGHT}px)` }}>
                {visibleItems.items.map((res) => (
                  <div 
                    key={res.id} 
                    className={`mb-2 mx-4 bg-[#111114] border rounded-xl overflow-hidden transition-all ${expandedId === res.id ? 'border-blue-500 ring-1 ring-blue-500/20' : 'border-gray-800'}`}
                  >
                    <div className="p-4 flex items-center justify-between cursor-pointer" onClick={() => handleExpand(res)}>
                      <div className="flex items-center gap-4 min-w-0">
                        <div className="w-10 h-10 bg-[#1a1a1e] rounded-lg flex items-center justify-center">
                          {res.fileName.endsWith('.csv') ? <FileCode className="w-5 h-5 text-gray-400" /> : <FileSpreadsheet className="w-5 h-5 text-gray-400" />}
                        </div>
                        <div className="min-w-0">
                          <div className="text-xs font-bold truncate">{res.fileName} <span className="text-[10px] text-gray-600 ml-2">ROW {res.rowNumber}</span></div>
                          <div className="text-[10px] text-gray-500 truncate italic mt-0.5">{res.searchString}</div>
                        </div>
                      </div>
                      <ChevronDown className={`w-4 h-4 transition-transform ${expandedId === res.id ? 'rotate-180 text-blue-500' : 'text-gray-600'}`} />
                    </div>
                    {expandedId === res.id && (
                      <div className="px-6 pb-6 pt-2 border-t border-gray-800/50 bg-[#0d0d0f]">
                        {expandedData ? (
                          <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-3 mt-2">
                            {Object.entries(expandedData).map(([key, value]) => (
                              <div key={key} className="p-2 bg-[#1a1a1e] rounded-lg border border-gray-800">
                                <div className="text-[8px] font-bold text-gray-500 uppercase mb-1">{key}</div>
                                <div className="text-[10px] text-gray-200">{String(value) || '—'}</div>
                              </div>
                            ))}
                          </div>
                        ) : (
                          <div className="flex items-center gap-3 py-4 text-xs text-gray-500 italic">
                            <Loader2 className="w-4 h-4 animate-spin" /> Fetching row details from source...
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          ) : (
            <div className="h-full flex flex-col items-center justify-center opacity-20">
              <Database className="w-24 h-24 mb-4" />
              <div className="text-xl font-black uppercase">Mount Directory</div>
            </div>
          )}
        </div>
      </main>

      <footer className="px-6 py-3 bg-[#111114] border-t border-gray-800 flex justify-between items-center text-[9px] font-black text-gray-600 uppercase">
        <div className="flex gap-8">
          <span>CORES: {MAX_WORKERS}</span>
          <span>FILES: {files.length}</span>
          {searchProgress.errors.length > 0 && <span className="text-red-500 animate-pulse">ERRORS: {searchProgress.errors.length}</span>}
        </div>
        <span>X-SEARCH PRO V4.1.0-DEFERRED</span>
      </footer>

      {showSummary && allResults.length > 0 && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 pointer-events-none">
          <div className="bg-[#111114] border border-gray-800 rounded-3xl p-8 text-center shadow-2xl pointer-events-auto">
            <CheckCircle2 className="w-12 h-12 text-blue-500 mx-auto mb-4" />
            <div className="text-lg font-bold mb-1">{allResults.length.toLocaleString()} Matches</div>
            <div className="text-[10px] text-gray-500 mb-6 uppercase">Time: {((Date.now() - startTime.current) / 1000).toFixed(2)}s</div>
            <button onClick={() => setShowSummary(false)} className="px-8 py-3 bg-blue-600 text-white rounded-xl text-[10px] font-bold uppercase">Explore</button>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
