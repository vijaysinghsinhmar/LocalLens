export const workerScript = `
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js');

  // ─────────────────────────────────────────────────────────────────────────
  // Memory strategy: never cache full parsed rows.
  // rowCache stores ONLY matched rows (tiny). No parsedCache.
  // Workers cap at 4 to limit peak RAM.
  // ─────────────────────────────────────────────────────────────────────────

  const CHUNK_SIZE = 200;

  const PRIORITY_SET = new Set([
    'date','amount','nature','debit','credit','balance','name','particulars',
    'ref','account','description','memo','payee','vendor','customer',
    'invoice','total','qty','price','narration','remarks','note','utr','trn'
  ]);

  let fileCache  = new Map();  // fileId -> { blob, type, name, path }
  let matchCache = new Map();  // fileId -> Set<rowIdx>
  let rowCache   = new Map();  // fileId -> Map<rowIdx, cells[]>  (matched rows ONLY)
  let hdrCache   = new Map();  // fileId -> string[]

  self.onmessage = async function(e) {
    const { action, payload } = e.data;
    if      (action === 'PROCESS_FILE')     await processFile(payload);
    else if (action === 'FETCH_ROW_DETAIL') fetchRowDetail(payload);
    else if (action === 'GENERATE_EXPORT')  await generateExport();
    else if (action === 'CLEAR_CACHE') {
      fileCache.clear(); matchCache.clear(); rowCache.clear(); hdrCache.clear();
    }
  };

  // ── Fast match check — no allocations ────────────────────────────────────
  function checkMatch(cells, lq, terms, exactMatch, fuzzy) {
    // Build search string inline — single pass
    let s = '';
    for (let i = 0; i < cells.length; i++) {
      if (cells[i]) { s += cells[i]; s += '|'; }
    }
    const ls = s.toLowerCase();

    if (exactMatch) {
      for (let i = 0; i < cells.length; i++)
        if (cells[i].toLowerCase() === lq) return 10;
      return 0;
    }
    if (fuzzy) {
      let score = 0;
      for (let t = 0; t < terms.length; t++) {
        const idx = ls.indexOf(terms[t]);
        if (idx === -1) return 0;
        score += idx === 0 ? 3 : 1;
      }
      return score;
    }
    const idx = ls.indexOf(lq);
    return idx === -1 ? 0 : (idx === 0 ? 3 : 1);
  }

  function buildPreview(cells, headers, previewIdx) {
    const parts = [];
    for (let pi = 0; pi < previewIdx.length && parts.length < 4; pi++) {
      const ci = previewIdx[pi];
      if (cells[ci]) parts.push(headers[ci] + ': ' + cells[ci]);
    }
    return parts.length ? parts.join(' • ') : (cells[0] || '').slice(0, 120);
  }

  function getPreviewIdx(headers) {
    const idx = [];
    for (let i = 0; i < headers.length && idx.length < 4; i++) {
      const bare = headers[i].toLowerCase().replace(/^[^:]+::/, '');
      for (const pk of PRIORITY_SET) {
        if (bare.includes(pk)) { idx.push(i); break; }
      }
    }
    return idx.length ? idx : [0,1,2].filter(i => i < headers.length);
  }

  // ── XLSX processor ────────────────────────────────────────────────────────
  // Uses dense array-of-arrays (fastest XLSX layout).
  // Does NOT call sheet_to_json — that creates a massive object array.
  async function processXLSX(fileId, blob, name, path, lq, terms, exactMatch, fuzzy) {
    const buf = await blob.arrayBuffer();
    // raw:true skips number/date formatting — pure string values, much faster
    const wb  = XLSX.read(buf, { type:'array', raw:true, dense:true, sheetStubs:false });

    const matchedRows = new Set();
    const rowData     = new Map();
    let   pending     = [];
    let   totalRows   = 0;
    const multiSheet  = wb.SheetNames.length > 1;
    let   fileHeaders = null;
    let   previewIdx  = null;

    const flush = () => {
      if (pending.length) {
        self.postMessage({ action:'MATCH_CHUNK', payload:{ fileId, matches: pending } });
        pending = [];
      }
    };

    for (const sName of wb.SheetNames) {
      const ws = wb.Sheets[sName];
      if (!ws || !ws['!data'] || ws['!data'].length < 2) continue;

      const raw    = ws['!data'];
      const hdrRaw = raw[0] || [];
      const headers = hdrRaw.map(c => c ? String(c.v ?? '') : '');
      const prefixed = multiSheet ? headers.map(h => sName+'::'+h) : headers;

      // Only set once — first sheet defines file headers for hdrCache
      if (!fileHeaders) {
        fileHeaders = prefixed;
        previewIdx  = getPreviewIdx(fileHeaders);
        hdrCache.set(fileId, fileHeaders);
      }

      for (let r = 1; r < raw.length; r++) {
        totalRows++;
        const srcRow = raw[r] || [];
        const cells  = new Array(headers.length);
        for (let c = 0; c < headers.length; c++) {
          const cell = srcRow[c];
          cells[c]   = cell ? String(cell.v ?? '') : '';
        }

        const score = checkMatch(cells, lq, terms, exactMatch, fuzzy);
        if (!score) continue;

        const rowIdx = totalRows - 1;
        matchedRows.add(rowIdx);
        rowData.set(rowIdx, cells);

        pending.push({
          id: fileId+'-'+rowIdx, fileId,
          fileName: name, filePath: path, sheetName: sName,
          rowNumber: r + 1,
          searchString: buildPreview(cells, prefixed, previewIdx || []),
          score
        });
        if (pending.length >= CHUNK_SIZE) flush();
      }
    }

    flush();
    return { matchedRows, rowData, totalRows };
  }

  // ── CSV processor ─────────────────────────────────────────────────────────
  async function processCSV(fileId, blob, name, path, lq, terms, exactMatch, fuzzy) {
    const text    = await blob.text();
    const matchedRows = new Set();
    const rowData     = new Map();
    let   pending     = [];
    let   totalRows   = 0;
    let   headers     = null;
    let   previewIdx  = null;

    const flush = () => {
      if (pending.length) {
        self.postMessage({ action:'MATCH_CHUNK', payload:{ fileId, matches: pending } });
        pending = [];
      }
    };

    Papa.parse(text, {
      skipEmptyLines: true,
      step: ({ data }) => {
        if (!headers) {
          headers    = data.map(v => String(v ?? ''));
          previewIdx = getPreviewIdx(headers);
          hdrCache.set(fileId, headers);
          return;
        }
        totalRows++;
        const cells  = data.map(v => String(v ?? ''));
        const rowIdx = totalRows - 1;
        const score  = checkMatch(cells, lq, terms, exactMatch, fuzzy);
        if (!score) return;

        matchedRows.add(rowIdx);
        rowData.set(rowIdx, cells);
        pending.push({
          id: fileId+'-'+rowIdx, fileId,
          fileName: name, filePath: path, sheetName: null,
          rowNumber: totalRows + 1,
          searchString: buildPreview(cells, headers, previewIdx || []),
          score
        });
        if (pending.length >= CHUNK_SIZE) flush();
      }
    });

    flush();
    return { matchedRows, rowData, totalRows };
  }

  // ── TXT processor ─────────────────────────────────────────────────────────
  async function processTXT(fileId, blob, name, path, lq, terms, exactMatch, fuzzy) {
    const text    = await blob.text();
    const lines   = text.split(/\r?\n/);
    const matchedRows = new Set();
    const rowData     = new Map();
    let   pending     = [];
    let   totalRows   = 0;

    hdrCache.set(fileId, ['line']);

    const flush = () => {
      if (pending.length) {
        self.postMessage({ action:'MATCH_CHUNK', payload:{ fileId, matches: pending } });
        pending = [];
      }
    };

    for (let i = 0; i < lines.length; i++) {
      const l = lines[i].trim();
      if (!l) continue;
      totalRows++;
      const cells  = [l];
      const rowIdx = totalRows - 1;
      const score  = checkMatch(cells, lq, terms, exactMatch, fuzzy);
      if (!score) continue;

      matchedRows.add(rowIdx);
      rowData.set(rowIdx, cells);
      pending.push({
        id: fileId+'-'+rowIdx, fileId,
        fileName: name, filePath: path, sheetName: null,
        rowNumber: i + 1,
        searchString: l.slice(0, 120),
        score
      });
      if (pending.length >= CHUNK_SIZE) flush();
    }

    flush();
    return { matchedRows, rowData, totalRows };
  }

  // ── Main dispatcher ───────────────────────────────────────────────────────
  async function processFile({ fileId, blob, type, name, path, query, exactMatch, fuzzy }) {
    fileCache.set(fileId, { blob, type, name, path });
    matchCache.delete(fileId);
    rowCache.delete(fileId);
    hdrCache.delete(fileId);

    const lq    = query.toLowerCase().trim();
    const terms = lq.split(/\s+/).filter(t => t.length > 0);

    try {
      let result;
      if      (type === 'xlsx' || type === 'xls') result = await processXLSX(fileId, blob, name, path, lq, terms, exactMatch, fuzzy);
      else if (type === 'csv')                    result = await processCSV(fileId, blob, name, path, lq, terms, exactMatch, fuzzy);
      else                                        result = await processTXT(fileId, blob, name, path, lq, terms, exactMatch, fuzzy);

      matchCache.set(fileId, result.matchedRows);
      rowCache.set(fileId, result.rowData);
      self.postMessage({ action:'FILE_COMPLETE', payload:{ fileId, totalRows: result.totalRows } });
    } catch(err) {
      self.postMessage({ action:'FILE_ERROR', payload:{ fileId, error: err.message } });
    }
  }

  // ── Row detail — instant from rowCache ────────────────────────────────────
  function fetchRowDetail({ fileId, rowNumber }) {
    const rows    = rowCache.get(fileId);
    const headers = hdrCache.get(fileId);
    const rowIdx  = rowNumber - 2;

    if (!rows || !headers || !rows.has(rowIdx)) {
      self.postMessage({ action:'ROW_DETAIL_RESULT', payload:{ fileId, rowNumber, rowData: null } });
      return;
    }
    const cells = rows.get(rowIdx);
    const obj   = {};
    for (let i = 0; i < headers.length; i++) {
      const bare = headers[i].replace(/^[^:]+::/, '');
      obj[bare]  = cells[i] || '';
    }
    self.postMessage({ action:'ROW_DETAIL_RESULT', payload:{ fileId, rowNumber, rowData: obj } });
  }

  // ── Export — re-streams from blob (export is infrequent) ─────────────────
  async function generateExport() {
    const parts   = [];
    let firstFile = true;

    for (const [fileId, matched] of matchCache.entries()) {
      if (!matched.size) continue;
      const info    = fileCache.get(fileId);
      const headers = hdrCache.get(fileId);
      if (!info || !headers) continue;

      const bareHdrs = headers.map(h => h.replace(/^[^:]+::/, ''));
      if (firstFile) {
        parts.push(['FileName','Sheet','RowNumber',...bareHdrs]
          .map(h => '"'+h.replace(/"/g,'""')+'"').join(','));
        firstFile = false;
      }

      for (const [rowIdx, cells] of rowCache.get(fileId)?.entries() || []) {
        if (!matched.has(rowIdx)) continue;
        parts.push([
          '"'+info.name.replace(/"/g,'""')+'"',
          '"N/A"',
          rowIdx+2,
          ...cells.map(v => '"'+String(v).replace(/"/g,'""')+'"')
        ].join(','));
      }
    }

    self.postMessage({ action:'EXPORT_READY',
      payload:{ blob: new Blob([parts.join('\n')], { type:'text/csv' }) } });
  }
`;
