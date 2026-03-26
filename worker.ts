export const workerScript = `
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js');

  // ═══════════════════════════════════════════════════════════════════════════
  //  ARCHITECTURE: Parse-once, search-many flat string index
  //
  //  On first load of a file we build TWO things and keep them forever:
  //    1. flatLines[]  — one string per row: "cell1|cell2|cell3"  (lowercase)
  //    2. rawRows[]    — original string[][] for display / export
  //
  //  On every search we do plain indexOf on flatLines[] — no re-parsing,
  //  no arrayBuffer(), no PapaParse, no XLSX.read().
  //
  //  Memory: ~2–4 bytes per character. A 315MB XLSX with text data compresses
  //  to roughly the same after string conversion. We accept this trade-off
  //  because the alternative (re-parsing on every search) makes the app unusable.
  //
  //  Workers = 2 (never more). Two workers parse two files in parallel.
  //  More workers = more simultaneous arrayBuffer() calls = OOM.
  // ═══════════════════════════════════════════════════════════════════════════

  const CHUNK   = 300;   // batch size before postMessage

  const PRIO = new Set([
    'date','amount','nature','debit','credit','balance','name','particulars',
    'ref','account','description','memo','payee','vendor','customer',
    'invoice','total','qty','price','narration','remarks','note','utr','trn','tran'
  ]);

  // Per-worker persistent index — survives across searches
  // fileId -> { flatLines: string[], rawRows: string[][], headers: string[], sheets: string[] }
  const INDEX = new Map();

  self.onmessage = async function(e) {
    const { action, payload } = e.data;
    if      (action === 'INDEX_FILE')      await indexFile(payload);
    else if (action === 'SEARCH_FILE')     searchFile(payload);
    else if (action === 'FETCH_ROW')       fetchRow(payload);
    else if (action === 'EXPORT_MATCHES')  exportMatches(payload);
    else if (action === 'DROP_INDEX')      INDEX.delete(payload.fileId);
    else if (action === 'CLEAR_ALL')       INDEX.clear();
  };

  // ── PHASE 1: Index a file (done once per file load) ──────────────────────
  async function indexFile({ fileId, blob, type, name }) {
    try {
      const flatLines = [];  // lowercase concatenated row strings for search
      const rawRows   = [];  // original cells for display
      const sheets    = [];  // sheet name per row
      let   headers   = [];

      if (type === 'xlsx' || type === 'xls') {
        const buf = await blob.arrayBuffer();
        const wb  = XLSX.read(buf, { type:'array', raw:true, dense:true, sheetStubs:false });
        const multi = wb.SheetNames.length > 1;

        for (const sName of wb.SheetNames) {
          const ws = wb.Sheets[sName];
          if (!ws || !ws['!data'] || ws['!data'].length < 2) continue;
          const data   = ws['!data'];
          const hdrRow = data[0] || [];
          const hdrs   = hdrRow.map(c => c ? String(c.v ?? '') : '');

          // Only set headers from first sheet
          if (!headers.length) {
            headers = multi ? hdrs.map(h => sName+'::'+h) : hdrs;
          }

          for (let r = 1; r < data.length; r++) {
            const src   = data[r] || [];
            const cells = new Array(hdrs.length);
            let   flat  = '';
            for (let c = 0; c < hdrs.length; c++) {
              const v  = src[c] ? String(src[c].v ?? '') : '';
              cells[c] = v;
              if (v) { flat += v; flat += '|'; }
            }
            flatLines.push(flat.toLowerCase());
            rawRows.push(cells);
            sheets.push(sName);
          }
        }

      } else if (type === 'csv') {
        const text  = await blob.text();
        let   first = true;
        Papa.parse(text, {
          skipEmptyLines: true,
          step: ({ data }) => {
            if (first) { headers = data.map(v => String(v ?? '')); first = false; return; }
            const cells = data.map(v => String(v ?? ''));
            let flat = '';
            for (let i = 0; i < cells.length; i++) {
              if (cells[i]) { flat += cells[i]; flat += '|'; }
            }
            flatLines.push(flat.toLowerCase());
            rawRows.push(cells);
            sheets.push(null);
          }
        });

      } else {
        // TXT
        headers = ['line'];
        const text  = await blob.text();
        const lines = text.split(/\r?\n/);
        for (const l of lines) {
          const t = l.trim();
          if (!t) continue;
          flatLines.push(t.toLowerCase());
          rawRows.push([t]);
          sheets.push(null);
        }
      }

      INDEX.set(fileId, { flatLines, rawRows, headers, sheets, name });
      self.postMessage({ action:'INDEX_DONE', payload:{ fileId, rowCount: flatLines.length } });
    } catch(err) {
      self.postMessage({ action:'INDEX_ERROR', payload:{ fileId, error: err.message } });
    }
  }

  // ── PHASE 2: Search (pure JS, no I/O, no parsing) ────────────────────────
  function searchFile({ fileId, query, exactMatch, fuzzy, searchId }) {
    const idx = INDEX.get(fileId);
    if (!idx) {
      self.postMessage({ action:'SEARCH_ERROR', payload:{ fileId, searchId, error:'Not indexed' } });
      return;
    }

    const { flatLines, rawRows, headers, sheets, name } = idx;
    const lq    = query.toLowerCase().trim();
    const terms = lq.split(/\s+/).filter(t => t.length > 0);
    const empty = !lq;

    // Pre-compute priority preview column indices once
    const previewIdx = [];
    for (let i = 0; i < headers.length && previewIdx.length < 4; i++) {
      const bare = headers[i].toLowerCase().replace(/^[^:]+::/, '');
      for (const pk of PRIO) { if (bare.includes(pk)) { previewIdx.push(i); break; } }
    }
    if (!previewIdx.length) {
      for (let i = 0; i < Math.min(3, headers.length); i++) previewIdx.push(i);
    }

    const matches = [];

    for (let ri = 0; ri < flatLines.length; ri++) {
      const flat = flatLines[ri];
      let score  = 0;

      if (empty) {
        score = 1;
      } else if (exactMatch) {
        // Exact: check each raw cell
        const cells = rawRows[ri];
        for (let ci = 0; ci < cells.length; ci++) {
          if (cells[ci].toLowerCase() === lq) { score = 10; break; }
        }
      } else if (fuzzy) {
        // All terms must appear
        let s = 0, ok = true;
        for (let t = 0; t < terms.length; t++) {
          const pos = flat.indexOf(terms[t]);
          if (pos === -1) { ok = false; break; }
          s += pos === 0 ? 3 : 1;
        }
        if (ok) score = s;
      } else {
        // Substring
        const pos = flat.indexOf(lq);
        if (pos !== -1) score = pos === 0 ? 3 : 1;
      }

      if (!score) continue;

      // Build preview from raw cells
      const cells = rawRows[ri];
      const parts = [];
      for (let pi = 0; pi < previewIdx.length; pi++) {
        const v = cells[previewIdx[pi]];
        if (v) {
          const h = headers[previewIdx[pi]].replace(/^[^:]+::/, '');
          parts.push(h + ': ' + v);
        }
      }

      matches.push({
        id:           fileId + '-' + ri,
        fileId,
        fileName:     name,
        sheetName:    sheets[ri],
        rowNumber:    ri + 2,
        searchString: parts.length ? parts.join(' • ') : (cells[0] || '').slice(0, 100),
        score,
        _ri: ri   // keep for row detail lookup
      });

      if (matches.length >= CHUNK) {
        self.postMessage({ action:'SEARCH_CHUNK', payload:{ fileId, searchId, matches: matches.splice(0) } });
      }
    }

    if (matches.length) {
      self.postMessage({ action:'SEARCH_CHUNK', payload:{ fileId, searchId, matches } });
    }
    self.postMessage({ action:'SEARCH_DONE', payload:{ fileId, searchId } });
  }

  // ── Row detail — O(1) from index ──────────────────────────────────────────
  function fetchRow({ fileId, ri }) {
    const idx = INDEX.get(fileId);
    if (!idx || !idx.rawRows[ri]) {
      self.postMessage({ action:'ROW_DATA', payload:{ fileId, ri, data: null } });
      return;
    }
    const cells = idx.rawRows[ri];
    const obj   = {};
    for (let i = 0; i < idx.headers.length; i++) {
      obj[idx.headers[i].replace(/^[^:]+::/, '')] = cells[i] || '';
    }
    self.postMessage({ action:'ROW_DATA', payload:{ fileId, ri, data: obj } });
  }

  // ── Export matched rows ───────────────────────────────────────────────────
  function exportMatches({ matches }) {
    if (!matches.length) {
      self.postMessage({ action:'EXPORT_READY', payload:{ blob: new Blob([''], { type:'text/csv' }) } });
      return;
    }
    const lines = [];
    // Group by fileId to get headers
    const byFile = new Map();
    for (const m of matches) {
      if (!byFile.has(m.fileId)) byFile.set(m.fileId, []);
      byFile.get(m.fileId).push(m);
    }

    let wroteHeader = false;
    for (const [fileId, ms] of byFile) {
      const idx = INDEX.get(fileId);
      if (!idx) continue;
      const bareHdrs = idx.headers.map(h => h.replace(/^[^:]+::/, ''));
      if (!wroteHeader) {
        lines.push(['FileName','Sheet','RowNumber',...bareHdrs]
          .map(h => '"'+h.replace(/"/g,'""')+'"').join(','));
        wroteHeader = true;
      }
      for (const m of ms) {
        const cells = idx.rawRows[m._ri] || [];
        lines.push([
          '"'+idx.name.replace(/"/g,'""')+'"',
          '"'+(m.sheetName||'N/A')+'"',
          m.rowNumber,
          ...cells.map(v => '"'+String(v).replace(/"/g,'""')+'"')
        ].join(','));
      }
    }
    self.postMessage({ action:'EXPORT_READY',
      payload:{ blob: new Blob([lines.join('\n')], { type:'text/csv' }) } });
  }
`;
