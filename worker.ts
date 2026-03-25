export const workerScript = `
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js');

  const CHUNK_SIZE = 300;
  const PRIORITY_KEYS = ['date','amount','nature','debit','credit','balance','name','particulars','ref','account','description','memo','payee','vendor','customer','invoice','total','qty','price'];

  // Per-worker caches
  let fileCache = new Map();   // fileId -> { blob, type, name }
  let matchCache = new Map();  // fileId -> Set<rowIndex>

  self.onmessage = async function(e) {
    const { action, payload } = e.data;
    if (action === 'PROCESS_FILE')    await processFile(payload);
    else if (action === 'FETCH_ROW_DETAIL') await fetchRowDetail(payload);
    else if (action === 'GENERATE_EXPORT')  await generateExport(payload);
    else if (action === 'CLEAR_CACHE') { fileCache.clear(); matchCache.clear(); }
  };

  // ── Fuzzy: all terms must appear ─────────────────────────────────────────
  function matchesFuzzy(queryTerms, str) {
    for (const t of queryTerms) if (!str.includes(t)) return false;
    return true;
  }

  // ── Relevance score (higher = better) ────────────────────────────────────
  function score(queryTerms, str) {
    let s = 0;
    for (const t of queryTerms) {
      const idx = str.indexOf(t);
      if (idx === -1) return 0;
      if (idx === 0) s += 3;
      else s += 1;
    }
    return s;
  }

  // ── Row detail fetch ──────────────────────────────────────────────────────
  async function fetchRowDetail({ fileId, rowNumber, sheetName }) {
    const file = fileCache.get(fileId);
    if (!file) { self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData: null } }); return; }
    try {
      let rowData = null;
      if (file.type === 'xlsx' || file.type === 'xls') {
        const buf = await file.blob.arrayBuffer();
        const wb  = XLSX.read(buf, { type: 'array', dense: true });
        const ws  = wb.Sheets[sheetName || wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        rowData = rows[rowNumber - 2] || null;
      } else if (file.type === 'csv') {
        const text = await file.blob.text();
        const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
        rowData = parsed.data[rowNumber - 2] || null;
      } else {
        const text = await file.blob.text();
        const lines = text.split(/\r?\n/);
        rowData = { line: lines[rowNumber - 2] || '' };
      }
      self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData } });
    } catch(err) {
      self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData: null } });
    }
  }

  // ── Main file processor ───────────────────────────────────────────────────
  async function processFile({ fileId, blob, type, name, path, query, exactMatch, fuzzy }) {
    fileCache.set(fileId, { blob, type, name });
    const matchedRows = new Set();

    try {
      const lq = query.toLowerCase().trim();
      const terms = lq.split(/\s+/).filter(t => t.length > 0);
      const empty = terms.length === 0;
      let pending = [];
      let totalRows = 0;

      const flush = () => {
        if (pending.length) {
          self.postMessage({ action: 'MATCH_CHUNK', payload: { fileId, matches: pending } });
          pending = [];
        }
      };

      const onRow = (row, sName) => {
        totalRows++;
        const vals = Object.values(row).map(v => String(v));
        const full  = vals.map(v => v.toLowerCase()).join(' ');
        const previews = [];
        for (const [k, v] of Object.entries(row)) {
          if (previews.length >= 4) break;
          if (PRIORITY_KEYS.some(pk => k.toLowerCase().includes(pk))) previews.push(k + ': ' + v);
        }
        if (!previews.length) previews.push(...vals.slice(0, 3));

        let isMatch = empty;
        let rowScore = 0;
        if (!isMatch) {
          if (exactMatch) {
            isMatch = vals.some(v => v.toLowerCase() === lq);
            rowScore = isMatch ? 10 : 0;
          } else if (fuzzy) {
            rowScore = score(terms, full);
            isMatch  = rowScore > 0;
          } else {
            isMatch  = full.includes(lq);
            rowScore = isMatch ? full.indexOf(lq) === 0 ? 3 : 1 : 0;
          }
        }

        if (isMatch) {
          matchedRows.add(totalRows);
          pending.push({
            id: fileId + '-' + totalRows,
            fileId,
            fileName: name,
            filePath: path,
            sheetName: sName || null,
            rowNumber: totalRows + 1,
            searchString: previews.join(' • '),
            score: rowScore
          });
          if (pending.length >= CHUNK_SIZE) flush();
        }
      };

      if (type === 'xlsx' || type === 'xls') {
        const buf = await blob.arrayBuffer();
        const wb  = XLSX.read(buf, { type: 'array', dense: true, cellDates: true });
        for (const sName of wb.SheetNames) {
          const ws   = wb.Sheets[sName];
          const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
          for (const row of data) onRow(row, sName);
        }
      } else if (type === 'csv') {
        const text = await blob.text();
        Papa.parse(text, { header: true, skipEmptyLines: true, step: r => onRow(r.data, null) });
      } else {
        const text = await blob.text();
        for (const line of text.split(/\r?\n/)) {
          if (line.trim()) onRow({ line: line.trim() }, null);
        }
      }

      flush();
      matchCache.set(fileId, matchedRows);
      self.postMessage({ action: 'FILE_COMPLETE', payload: { fileId, totalRows } });
    } catch(err) {
      self.postMessage({ action: 'FILE_ERROR', payload: { fileId, error: err.message } });
    }
  }

  // ── Export all matched rows as CSV ────────────────────────────────────────
  async function generateExport() {
    const headers = new Set(['FileName', 'Sheet', 'RowNumber']);
    const rows    = [];

    for (const [fileId, matched] of matchCache.entries()) {
      if (!matched.size) continue;
      const file = fileCache.get(fileId);
      if (!file) continue;

      if (file.type === 'xlsx' || file.type === 'xls') {
        const buf = await file.blob.arrayBuffer();
        const wb  = XLSX.read(buf, { type: 'array', dense: true });
        for (const sName of wb.SheetNames) {
          const data = XLSX.utils.sheet_to_json(wb.Sheets[sName], { defval: '' });
          data.forEach((r, i) => {
            if (!matched.has(i + 1)) return;
            const out = { FileName: file.name, Sheet: sName, RowNumber: i + 2, ...r };
            Object.keys(out).forEach(k => headers.add(k));
            rows.push(out);
          });
        }
      } else if (file.type === 'csv') {
        const text = await file.blob.text();
        const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
        parsed.data.forEach((r, i) => {
          if (!matched.has(i + 1)) return;
          const out = { FileName: file.name, Sheet: 'N/A', RowNumber: i + 2, ...r };
          Object.keys(out).forEach(k => headers.add(k));
          rows.push(out);
        });
      }
    }

    const hArr = Array.from(headers);
    const csv  = [
      hArr.join(','),
      ...rows.map(r => hArr.map(h => '"' + String(r[h] ?? '').replace(/"/g, '""') + '"').join(','))
    ].join('\n');

    self.postMessage({ action: 'EXPORT_READY', payload: { blob: new Blob([csv], { type: 'text/csv' }) } });
  }
`;
