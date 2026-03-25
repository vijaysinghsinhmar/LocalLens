export const workerScript = `
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js');

  // ── Tuning constants ──────────────────────────────────────────────────────
  const CHUNK_SIZE   = 500;   // flush to main thread every N matches
  const YIELD_EVERY  = 5000;  // yield to event loop every N rows (keeps worker responsive)

  // Priority key set — O(1) lookup instead of O(N) .some() per cell
  const PRIORITY_SET = new Set([
    'date','amount','nature','debit','credit','balance','name','particulars',
    'ref','account','description','memo','payee','vendor','customer',
    'invoice','total','qty','price','narration','remarks','note'
  ]);

  let fileCache  = new Map();   // fileId -> { blob, type, name }
  let matchCache = new Map();   // fileId -> Set<rowIndex>
  // Parsed data cache: avoids re-parsing the same file for row-detail / export
  let parsedCache = new Map();  // fileId -> { headers: string[], rows: string[][] }

  self.onmessage = async function(e) {
    const { action, payload } = e.data;
    if      (action === 'PROCESS_FILE')    await processFile(payload);
    else if (action === 'FETCH_ROW_DETAIL') fetchRowDetail(payload);
    else if (action === 'GENERATE_EXPORT')  await generateExport();
    else if (action === 'CLEAR_CACHE')     { fileCache.clear(); matchCache.clear(); parsedCache.clear(); }
  };

  // ── Parse any file into { headers, rows } where every cell is a string ────
  // This is the hot path — optimised to avoid all unnecessary allocations.
  async function parseToStringRows(fileId, blob, type) {
    if (parsedCache.has(fileId)) return parsedCache.get(fileId);

    let result = { headers: [], rows: [] };

    if (type === 'xlsx' || type === 'xls') {
      const buf = await blob.arrayBuffer();
      // raw:true = no date parsing, no number formatting — just raw values
      // dense:true = array-of-arrays layout, much faster than object layout
      const wb  = XLSX.read(buf, { type: 'array', raw: true, dense: true });

      const allHeaders = [];
      const allRows    = [];

      for (const sName of wb.SheetNames) {
        const ws = wb.Sheets[sName];
        if (!ws || !ws['!data']) continue;
        const data = ws['!data']; // array of arrays (dense mode)
        if (data.length === 0) continue;

        // First row = headers
        const headerRow = data[0];
        const headers   = headerRow.map(c => (c ? String(c.v ?? '') : ''));

        // Build a combined header prefix for this sheet: "sheetname::colname"
        // so multi-sheet xlsx keeps columns distinguishable
        const prefixed = wb.SheetNames.length > 1
          ? headers.map(h => sName + '::' + h)
          : headers;

        allHeaders.push(...prefixed);

        for (let r = 1; r < data.length; r++) {
          const srcRow  = data[r] || [];
          const strRow  = new Array(headers.length);
          for (let c = 0; c < headers.length; c++) {
            const cell = srcRow[c];
            strRow[c]  = cell ? String(cell.v ?? '') : '';
          }
          // Attach sheet name as last "column" for display
          strRow._sheet = sName;
          allRows.push(strRow);
        }
      }

      result = { headers: allHeaders, rows: allRows };

    } else if (type === 'csv') {
      const text = await blob.text();
      // Worker-local parse — step mode to avoid holding entire parsed structure
      const headers = [];
      const rows    = [];
      let   first   = true;
      Papa.parse(text, {
        skipEmptyLines: true,
        step: ({ data }) => {
          if (first) { headers.push(...data); first = false; return; }
          rows.push(data.map(v => String(v ?? '')));
        }
      });
      result = { headers, rows };

    } else {
      // Plain text: one "column" called "line"
      const text  = await blob.text();
      const lines = text.split(/\\r?\\n/);
      result = {
        headers: ['line'],
        rows: lines.filter(l => l.trim()).map(l => [l.trim()])
      };
    }

    parsedCache.set(fileId, result);
    return result;
  }

  // ── Main search ───────────────────────────────────────────────────────────
  async function processFile({ fileId, blob, type, name, path, query, exactMatch, fuzzy }) {
    fileCache.set(fileId, { blob, type, name });

    const matchedRows = new Set();
    let pending       = [];
    let totalRows     = 0;

    const flush = () => {
      if (pending.length) {
        self.postMessage({ action: 'MATCH_CHUNK', payload: { fileId, matches: pending } });
        pending = [];
      }
    };

    try {
      const lq    = query.toLowerCase().trim();
      const terms = lq.split(/\\s+/).filter(t => t.length > 0);
      const empty = terms.length === 0;

      const { headers, rows } = await parseToStringRows(fileId, blob, type);

      // Pre-compute which header indices are "priority" for preview
      // Do this ONCE per file, not per row
      const priorityIdx = [];
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i].toLowerCase();
        // strip sheet prefix if present
        const bare = h.includes('::') ? h.split('::')[1] : h;
        for (const pk of PRIORITY_SET) {
          if (bare.includes(pk)) { priorityIdx.push(i); break; }
        }
        if (priorityIdx.length >= 4) break;
      }
      // Fallback: first 3 columns
      const previewIdx = priorityIdx.length ? priorityIdx : [0, 1, 2].filter(i => i < headers.length);

      // ── Hot loop ─────────────────────────────────────────────────────────
      for (let ri = 0; ri < rows.length; ri++) {
        // Yield to event loop periodically so worker stays responsive
        if ((ri & 4999) === 4999) await new Promise(r => setTimeout(r, 0));

        totalRows++;
        const row = rows[ri];

        // Build single concatenated search string — ONE allocation per row
        // Use a separator unlikely to appear in data
        let searchStr = '';
        for (let ci = 0; ci < row.length; ci++) {
          if (row[ci]) searchStr += row[ci] + '|';
        }
        const lowerStr = searchStr.toLowerCase();

        // ── Match check ───────────────────────────────────────────────────
        let isMatch  = empty;
        let rowScore = 0;

        if (!isMatch) {
          if (exactMatch) {
            // Exact: any single cell equals query exactly
            for (let ci = 0; ci < row.length; ci++) {
              if (row[ci].toLowerCase() === lq) { isMatch = true; rowScore = 10; break; }
            }
          } else if (fuzzy) {
            // Fuzzy: ALL terms must appear somewhere in the row
            let allFound = true;
            let s = 0;
            for (let ti = 0; ti < terms.length; ti++) {
              const idx = lowerStr.indexOf(terms[ti]);
              if (idx === -1) { allFound = false; break; }
              s += idx === 0 ? 3 : 1;
            }
            if (allFound) { isMatch = true; rowScore = s; }
          } else {
            // Substring: query appears anywhere in the row
            const idx = lowerStr.indexOf(lq);
            if (idx !== -1) { isMatch = true; rowScore = idx === 0 ? 3 : 1; }
          }
        }

        if (!isMatch) continue;

        matchedRows.add(ri);

        // Build preview — only for matched rows, using pre-computed indices
        const parts = [];
        for (let pi = 0; pi < previewIdx.length; pi++) {
          const ci  = previewIdx[pi];
          const val = row[ci];
          if (val) {
            const hdr = headers[ci];
            const bare = hdr.includes('::') ? hdr.split('::')[1] : hdr;
            parts.push(bare + ': ' + val);
          }
        }

        pending.push({
          id:           fileId + '-' + ri,
          fileId,
          fileName:     name,
          filePath:     path,
          sheetName:    row._sheet || null,
          rowNumber:    ri + 2,   // 1-based + header offset
          searchString: parts.length ? parts.join(' • ') : (row[0] || '').slice(0, 120),
          score:        rowScore
        });

        if (pending.length >= CHUNK_SIZE) flush();
      }

      flush();
      matchCache.set(fileId, matchedRows);
      self.postMessage({ action: 'FILE_COMPLETE', payload: { fileId, totalRows } });
    } catch(err) {
      self.postMessage({ action: 'FILE_ERROR', payload: { fileId, error: err.message } });
    }
  }

  // ── Row detail — uses parsed cache, no re-parse ───────────────────────────
  function fetchRowDetail({ fileId, rowNumber, sheetName }) {
    const parsed = parsedCache.get(fileId);
    if (!parsed) {
      self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData: null } });
      return;
    }
    const { headers, rows } = parsed;
    const row = rows[rowNumber - 2]; // rowNumber is 1-based + header offset
    if (!row) {
      self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData: null } });
      return;
    }
    const obj = {};
    for (let i = 0; i < headers.length; i++) {
      const hdr  = headers[i];
      const bare = hdr.includes('::') ? hdr.split('::')[1] : hdr;
      obj[bare]  = row[i] || '';
    }
    self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData: obj } });
  }

  // ── Export — uses parsed cache ────────────────────────────────────────────
  async function generateExport() {
    const csvParts = [];
    let headerWritten = false;
    let globalHeaders = [];

    for (const [fileId, matched] of matchCache.entries()) {
      if (!matched.size) continue;
      const file   = fileCache.get(fileId);
      const parsed = parsedCache.get(fileId);
      if (!file || !parsed) continue;

      const { headers, rows } = parsed;

      if (!headerWritten) {
        globalHeaders = ['FileName', 'Sheet', 'RowNumber', ...headers.map(h => h.includes('::') ? h.split('::')[1] : h)];
        csvParts.push(globalHeaders.map(h => '"' + h.replace(/"/g, '""') + '"').join(','));
        headerWritten = true;
      }

      for (const ri of matched) {
        const row = rows[ri];
        if (!row) continue;
        const cells = [
          '"' + file.name.replace(/"/g, '""') + '"',
          '"' + (row._sheet || 'N/A') + '"',
          ri + 2,
          ...row.map(v => '"' + String(v ?? '').replace(/"/g, '""') + '"')
        ];
        csvParts.push(cells.join(','));
      }
    }

    const csv  = csvParts.join('\\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    self.postMessage({ action: 'EXPORT_READY', payload: { blob } });
  }
`;
