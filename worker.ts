
/**
 * X-Search Pro Background Worker (High-Throughput Edition)
 * Optimized for O(1) memory growth and deferred data fetching.
 */
export const workerScript = `
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
  self.importScripts('https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js');

  const CHUNK_SIZE = 500; 
  const PRIORITY_KEYS = ['date', 'amount', 'nature', 'debit', 'credit', 'balance', 'name', 'particulars', 'ref', 'account'];
  
  // Cache for the current session to avoid re-parsing during detail fetching or export
  let fileCache = new Map(); // id -> { blob, type, name }
  let matchCache = new Map(); // fileId -> Set of matched row numbers

  self.onmessage = async function(e) {
    const { action, payload } = e.data;
    
    if (action === 'PROCESS_FILE') {
      await processFile(payload);
    } else if (action === 'FETCH_ROW_DETAIL') {
      await fetchRowDetail(payload);
    } else if (action === 'GENERATE_EXPORT') {
      await generateExport(payload);
    }
  };

  function matchesFuzzy(queryTerms, searchableString) {
    for (const term of queryTerms) {
      if (!searchableString.includes(term)) return false;
    }
    return true;
  }

  async function fetchRowDetail({ fileId, rowNumber, sheetName }) {
    const file = fileCache.get(fileId);
    if (!file) return;

    try {
      let rowData = null;
      if (file.type === 'xlsx' || file.type === 'xls') {
        const buffer = await file.blob.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', dense: true });
        const sheet = workbook.Sheets[sheetName || workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        rowData = jsonData[rowNumber - 2]; // Adjust for 1-based row numbers and header
      } else if (file.type === 'csv') {
        const text = await file.blob.text();
        const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
        rowData = parsed.data[rowNumber - 2];
      } else {
        const text = await file.blob.text();
        const lines = text.split(/\\r?\\n/);
        rowData = { content: lines[rowNumber - 2] };
      }
      self.postMessage({ action: 'ROW_DETAIL_RESULT', payload: { fileId, rowNumber, rowData } });
    } catch (err) {
      console.error("Detail fetch error", err);
    }
  }

  async function processFile({ fileId, blob, type, name, path, query, exactMatch, fuzzy }) {
    fileCache.set(fileId, { blob, type, name });
    const fileMatchedRows = new Set();
    
    try {
      const lowerQuery = query.toLowerCase();
      const queryTerms = lowerQuery.split(/\\s+/).filter(t => t.length > 0);
      const isQueryEmpty = queryTerms.length === 0;
      let matches = [];
      let totalRows = 0;

      const onRowProcessed = (row, sName = null) => {
        totalRows++;
        let searchable = [];
        let preview = [];
        
        for (const key in row) {
          const val = String(row[key]);
          searchable.push(val.toLowerCase());
          const lKey = key.toLowerCase();
          if (PRIORITY_KEYS.some(pk => lKey.includes(pk)) && preview.length < 4) {
            preview.push(\`\${key}: \${val}\`);
          }
        }

        const fullStr = searchable.join(' ');
        let isMatch = isQueryEmpty;
        if (!isMatch) {
          if (exactMatch) isMatch = searchable.some(v => v === lowerQuery);
          else if (fuzzy) isMatch = matchesFuzzy(queryTerms, fullStr);
          else isMatch = fullStr.includes(lowerQuery);
        }

        if (isMatch) {
          fileMatchedRows.add(totalRows);
          matches.push({
            id: \`\${fileId}-\${totalRows}\`,
            fileId,
            fileName: name,
            filePath: path,
            sheetName: sName,
            rowNumber: totalRows + 1,
            searchString: preview.length ? preview.join(' • ') : searchable.slice(0, 3).join(' • ')
          });

          if (matches.length >= CHUNK_SIZE) {
            self.postMessage({ action: 'MATCH_CHUNK', payload: { fileId, matches } });
            matches = [];
          }
        }
      };

      if (type === 'xlsx' || type === 'xls') {
        const buffer = await blob.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', dense: true, cellDates: true });
        for (const sName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sName];
          const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          for (const row of data) onRowProcessed(row, sName);
        }
      } else if (type === 'csv') {
        const text = await blob.text();
        Papa.parse(text, { 
          header: true, 
          skipEmptyLines: true, 
          step: (results) => onRowProcessed(results.data) 
        });
      } else {
        const text = await blob.text();
        text.split(/\\r?\\n/).forEach(line => {
          if (line.trim()) onRowProcessed({ content: line.trim() });
        });
      }

      if (matches.length > 0) self.postMessage({ action: 'MATCH_CHUNK', payload: { fileId, matches } });
      matchCache.set(fileId, fileMatchedRows);
      self.postMessage({ action: 'FILE_COMPLETE', payload: { fileId, totalRows } });
    } catch (err) {
      self.postMessage({ action: 'FILE_ERROR', payload: { fileId, error: err.message } });
    }
  }

  async function generateExport({ queries }) {
    // Generate CSV for all matches in the current session
    const allHeaders = new Set(['FileName', 'RowNumber', 'Sheet']);
    const allRows = [];

    for (const [fileId, matchedRowIndices] of matchCache.entries()) {
      const file = fileCache.get(fileId);
      if (!file || matchedRowIndices.size === 0) continue;

      let fileRows = [];
      if (file.type === 'xlsx' || file.type === 'xls') {
        const buffer = await file.blob.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', dense: true });
        workbook.SheetNames.forEach(sName => {
          const sheet = workbook.Sheets[sName];
          const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          data.forEach((r, idx) => {
            if (matchedRowIndices.has(idx + 1)) {
              const exportRow = { FileName: file.name, RowNumber: idx + 2, Sheet: sName, ...r };
              Object.keys(exportRow).forEach(k => allHeaders.add(k));
              fileRows.push(exportRow);
            }
          });
        });
      } else if (file.type === 'csv') {
        const text = await file.blob.text();
        const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
        parsed.data.forEach((r, idx) => {
          if (matchedRowIndices.has(idx + 1)) {
            const exportRow = { FileName: file.name, RowNumber: idx + 2, Sheet: 'N/A', ...r };
            Object.keys(exportRow).forEach(k => allHeaders.add(k));
            fileRows.push(exportRow);
          }
        });
      }
      allRows.push(...fileRows);
    }

    const headers = Array.from(allHeaders);
    const csvContent = [
      headers.join(','),
      ...allRows.map(r => headers.map(h => \`"\${String(r[h] || '').replace(/"/g, '""')}"\`).join(','))
    ].join('\\n');

    const blob = new Blob([csvContent], { type: 'text/csv' });
    self.postMessage({ action: 'EXPORT_READY', payload: { blob } });
  }
`;
