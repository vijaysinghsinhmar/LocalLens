export const workerScript = `
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js');

  // ─────────────────────────────────────────────────────────────────────────
  // ARCHITECTURE: Zero persistent parsed-data cache to prevent OOM.
  //
  //  • fileCache   = blob + metadata only (tiny)
  //  • matchCache  = Set<rowIndex> per file  (tiny)
  //  • rowCache    = Map<fileId, Map<rowIdx, string[]>> — ONLY matched rows
  //                  (never the whole file)
  //
  // Parsing happens once per search, streaming row-by-row.
  // Row detail reads from rowCache (already in memory for matched rows).
  // Export re-streams from blob (acceptable — export is infrequent).
  // ─────────────────────────────────────────────────────────────────────────

  const CHUNK_SIZE = 200;   // send matches to main thread every N hits

  const PRIORITY_SET = new Set([
    'date','amount','nature','debit','credit','balance','name','particulars',
    'ref','account','description','memo','payee','vendor','customer',
    'invoice','total','qty','price','narration','remarks','note','trn','utr'
  ]);

  let fileCache  = new Map();  // fileId -> { blob, type, name, path }
  let matchCache = new Map();  // fileId -> Set<rowIdx>
  let rowCache   = new Map();  // fileId -> Map<rowIdx, string[]>  (matched rows only)
  let hdrCache   = new Map();  // fileId -> string[]  (headers only)

  self.onmessage = async function(e) {
    const { action, payload } = e.data;
    if      (action === 'PROCESS_FILE')     await processFile(payload);
    else if (action === 'FETCH_ROW_DETAIL') fetchRowDetail(payload);
    else if (action === 'GENERATE_EXPORT')  await generateExport();
    else if (action === 'CLEAR_CACHE') {
      fileCache.clear(); matchCache.clear();
      rowCache.clear();  hdrCache.clear();
    }
  };

  // ── Helpers ───────────────────────────────────────────────────────────────

  function buildSearchStr(cells) {
    // single pass, no intermediate array — fastest concat
    let s = '';
    for (let i = 0; i < cells.length; i++) {
      if (cells[i]) { s += cells[i]; s += '|'; }
    }
    return s.toLowerCase();
  }

  function isMatch(lowerStr, lq, terms, exactCells, exactMatch, fuzzy) {
    if (exactMatch) {
      for (let i = 0; i < exactCells.length; i++)
        if (exactCells[i] === lq) return 10;
      return 0;
    }
    if (fuzzy) {
      let score = 0;
      for (let t = 0; t < terms.length; t++) {
        const idx = lowerStr.indexOf(terms[t]);
        if (idx === -1) return 0;
        score += idx === 0 ? 3 : 1;
      }
      return score;
    }
    // plain substring
    const idx = lowerStr.indexOf(lq);
    return idx === -1 ? 0 : idx === 0 ? 3 : 1;
  }

  function buildPreview(cells, headers, previewIdx) {
    const parts = [];
    for (let pi = 0; pi < previewIdx.length && parts.length < 4; pi++) {
      const ci = previewIdx[pi];
      const v  = cells[ci];
      if (v) parts.push(headers[ci] + ': ' + v);
    }
    return parts.length ? parts.join(' • ') : (cells[0] || '').slice(0, 100);
  }

  function resolvePreviewIdx(headers) {
    const idx = [];
    for (let i = 0; i < headers.length && idx.length < 4; i++) {
      const bare = headers[i].toLowerCase().replace(/^[^:]+::/, '');
      for (const pk of PRIORITY_SET) {
        if (bare.includes(pk)) { idx.push(i); break; }
      }
    }
    return idx.length ? idx : [0,1,2].filter(i => i < headers.length);
  }

  // ── XLSX streaming ────────────────────────────────────────────────────────
  // We do NOT call sheet_to_json — that builds a giant object array.
  // Instead we iterate ws['!data'] (dense array-of-arrays) directly.

  async function* xlsxRows(blob) {
    const buf = await blob.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', raw: true, dense: true, sheetStubs: false });
    const multiSheet = wb.SheetNames.length > 1;

    for (const sName of wb.SheetNames) {
      const ws = wb.Sheets[sName];
      if (!ws || !ws['!data'] || ws['!data'].length < 2) continue;

      const raw     = ws['!data'];
      const hdrRaw  = raw[0] || [];
      const headers = hdrRaw.map(c => c ? String(c.v ?? '') : '');
      const prefixed = multiSheet ? headers.map(h => sName + '::' + h) : headers;

      for (let r = 1; r < raw.length; r++) {
        const srcRow = raw[r] || [];
        const cells  = new Array(headers.length);
        for (let c = 0; c < headers.length; c++) {
          const cell = srcRow[c];
          cells[c]   = cell ? String(cell.v ?? '') : '';
        }
        yield { cells, headers: prefixed, sheet: sName, rowIdx: r - 1 };
      }
    }
  }

  async function* csvRows(blob) {
    const text    = await blob.text();
    let   headers = null;
    let   rowIdx  = 0;
    // step-parse: never holds all rows in memory
    let resolve;
    let queue = [];
    let done  = false;

    // PapaParse doesn't support async generators natively,
    // so we collect synchronously (CSV is usually smaller than XLSX anyway)
    Papa.parse(text, {
      skipEmptyLines: true,
      step: ({ data }) => {
        if (!headers) { headers = data.map(v => String(v ?? '')); return; }
        queue.push({ cells: data.map(v => String(v ?? '')), headers, sheet: null, rowIdx: rowIdx++ });
      }
    });

    for (const row of queue) yield row;
  }

  async function* txtRows(blob) {
    const text  = await blob.text();
    const lines = text.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const l = lines[i].trim();
      if (l) yield { cells: [l], headers: ['line'], sheet: null, rowIdx: i };
    }
  }

  function rowGenerator(type, blob) {
    if (type === 'xlsx' || type === 'xls') return xlsxRows(blob);
    if (type === 'csv')                    return csvRows(blob);
    return txtRows(blob);
  }

  // ── Main search ───────────────────────────────────────────────────────────
  async function processFile({ fileId, blob, type, name, path, query, exactMatch, fuzzy }) {
    fileCache.set(fileId, { blob, type, name, path });

    // Clear stale match/row data from previous search
    matchCache.delete(fileId);
    rowCache.delete(fileId);
    hdrCache.delete(fileId);

    const matchedRows = new Set();
    const rowData     = new Map();   // rowIdx -> cells  (matched only)

    const lq    = query.toLowerCase().trim();
    const terms = lq.split(/\s+/).filter(t => t.length > 0);
    const empty = terms.length === 0;

    let pending    = [];
    let totalRows  = 0;
    let headers    = null;
    let previewIdx = null;

    const flush = () => {
      if (pending.length) {
        self.postMessage({ action: 'MATCH_CHUNK', payload: { fileId, matches: pending } });
        pending = [];
      }
    };

    try {
      const gen = rowGenerator(type, blob);
      let   yieldCounter = 0;

      for await (const { cells, headers: rowHdrs, sheet, rowIdx } of gen) {
        // First row sets headers + previewIdx (same for all rows in a file)
        if (!headers) {
          headers    = rowHdrs;
          previewIdx = resolvePreviewIdx(headers);
          hdrCache.set(fileId, headers);
        }

        totalRows++;
        yieldCounter++;

        // Yield to event loop every 8k rows to keep worker responsive
        if (yieldCounter === 8000) {
          yieldCounter = 0;
          await new Promise(r => setTimeout(r, 0));
        }

        const lowerStr   = buildSearchStr(cells);
        const exactCells = exactMatch ? cells.map(c => c.toLowerCase()) : cells;
        const score      = empty ? 1 : isMatch(lowerStr, lq, terms, exactCells, exactMatch, fuzzy);

        if (!score) continue;

        matchedRows.add(rowIdx);
        rowData.set(rowIdx, cells);  // store ONLY matched row cells

        pending.push({
          id:           fileId + '-' + rowIdx,
          fileId,
          fileName:     name,
          filePath:     path,
          sheetName:    sheet,
          rowNumber:    rowIdx + 2,
          searchString: buildPreview(cells, headers, previewIdx),
          score
        });

        if (pending.length >= CHUNK_SIZE) flush();
      }

      flush();
      matchCache.set(fileId, matchedRows);
      rowCache.set(fileId, rowData);
      self.postMessage({ action: 'FILE_COMPLETE', payload: { fileId, totalRows } });

    } catch(err) {
      self.postMessage({ action: 'FILE_ERROR', payload: { fileId, error: err.message } });
    }
  }

  // ── Row detail — instant from rowCache ────────────────────────────────────
  function fetchRowDetail({ fileId, rowNumber }) {
    const rows    = rowCache.get(fileId);
    const headers = hdrCache.get(fileId);
    const rowIdx  = rowNumber - 2;

    if (!rows || !headers || !rows.has(rowIdx)) {
      self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData: null } });
      return;
    }

    const cells = rows.get(rowIdx);
    const obj   = {};
    for (let i = 0; i < headers.length; i++) {
      const bare = headers[i].replace(/^[^:]+::/, '');
      obj[bare]  = cells[i] || '';
    }
    self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData: obj } });
  }

  // ── Export — re-streams from blob (export is rare, correctness > speed) ──
  async function generateExport() {
    const parts    = [];
    let   firstHdr = true;

    for (const [fileId, matched] of matchCache.entries()) {
      if (!matched.size) continue;
      const info    = fileCache.get(fileId);
      const headers = hdrCache.get(fileId);
      if (!info || !headers) continue;

      const bareHdrs = headers.map(h => h.replace(/^[^:]+::/, ''));

      if (firstHdr) {
        parts.push(['FileName','Sheet','RowNumber',...bareHdrs]
          .map(h => '"' + h.replace(/"/g,'""') + '"').join(','));
        firstHdr = false;
      }

      const gen = rowGenerator(info.type, info.blob);
      for await (const { cells, sheet, rowIdx } of gen) {
        if (!matched.has(rowIdx)) continue;
        parts.push([
          '"' + info.name.replace(/"/g,'""') + '"',
          '"' + (sheet || 'N/A') + '"',
          rowIdx + 2,
          ...cells.map(v => '"' + v.replace(/"/g,'""') + '"')
        ].join(','));
      }
    }

    const csv  = parts.join('\n');
    self.postMessage({ action: 'EXPORT_READY', payload: { blob: new Blob([csv], { type: 'text/csv' }) } });
  }
`;
