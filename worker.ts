import * as XLSX from 'xlsx';
import Papa from 'papaparse';

// Unit separator — won't appear in normal spreadsheet data
const SEP = '\x1F';

interface DBEntry {
  name: string;
  type: string;
  flat:    string[];   // lowercase search string per row
  compact: string[];   // display values joined by SEP
  headers: string[];
  sheets:  (string|null)[];
  rowNums: number[];
}

const DB = new Map<string, DBEntry>();

// Split a compact string back into cell values
function getCells(entry: DBEntry, ri: number): string[] {
  const s = entry.compact[ri];
  if (!s) return [];
  return s.split(SEP);
}

// ── Serialised message queue — guarantees idx completes before search runs ───
let queue = Promise.resolve();

self.onmessage = (ev: MessageEvent) => {
  const d = ev.data;
  if (d.t === 'clear') { DB.clear(); queue = Promise.resolve(); return; }
  queue = queue
    .then(() => dispatch(d))
    .catch(e => console.error('[worker] unhandled in', d.t, e));
};

async function dispatch(d: any) {
  if      (d.t === 'idx')    await doIndex(d);
  else if (d.t === 'search')      doSearch(d);
  else if (d.t === 'row')         doRow(d);
  else if (d.t === 'export')      doExport(d);
}

// ── Cell value extraction ─────────────────────────────────────────────────────
// Returns { display, raw } where:
//   display = cell.w (formatted text: "15/03/2025", "1,234.56") — for compact[]
//   raw     = cell.v as string ("45678", "1234.56") — added to flat[] too
// This way users can search either "15/03/2025" OR "45678" and find the row.
function cellValues(cell: any): { display: string; raw: string } {
  if (cell == null || cell.v == null) return { display: '', raw: '' };
  const raw     = String(cell.v);
  const display = (cell.w != null && cell.w !== raw) ? String(cell.w) : raw;
  return { display, raw };
}

// ── INDEX ─────────────────────────────────────────────────────────────────────
async function doIndex(msg: any) {
  const { id, blob, ft: type, name } = msg;
  const flat:    string[]        = [];
  const compact: string[]        = [];
  const headers: string[]        = [];
  const sheets:  (string|null)[] = [];
  const rowNums: number[]        = [];

  try {
    if (type === 'xlsx' || type === 'xls') {

      const buf = await blob.arrayBuffer();
      const wb  = XLSX.read(buf, {
        type:       'array',
        dense:      true,
        // FIX: cellText:true populates cell.w with formatted string
        // raw:false alone would lose numeric precision; cellText gives us both
        cellText:   true,
        cellDates:  false,  // keep dates as serial numbers; cell.w gives formatted date
        cellNF:     false,
        cellStyles: false,
        cellHTML:   false,
        sheetStubs: false,
      });

      const multi = wb.SheetNames.length > 1;

      for (const sName of wb.SheetNames) {
        const ws = wb.Sheets[sName] as any;
        if (!ws?.['!data'] || ws['!data'].length < 2) {
          if (ws) ws['!data'] = null;
          continue;
        }

        const data   = ws['!data'] as any[][];
        const hdrRow = data[0] || [];

        // Headers: use formatted text (cell.w) so header names are readable
        const hdrs: string[] = hdrRow.map((c: any) => {
          if (c == null) return '';
          return c.w != null ? String(c.w) : (c.v != null ? String(c.v) : '');
        });

        if (!hdrs.every(h => h === '')) {
          // Only set headers from first sheet with actual content
          if (!headers.length) {
            for (const h of hdrs) headers.push(multi ? sName + '::' + h : h);
          }
        }
        if (!headers.length) { ws['!data'] = null; continue; }

        const nCols = hdrs.length;

        for (let r = 1; r < data.length; r++) {
          const src = data[r];
          let searchStr = '';  // goes into flat[] — lowercased
          let compStr   = '';  // goes into compact[] — display values

          if (src) {
            for (let c = 0; c < nCols; c++) {
              const cell = src[c] as any;
              const { display, raw } = cellValues(cell);

              // compact: use display value for showing to user
              if (c > 0) compStr += SEP;
              compStr += display;

              // flat: add BOTH display and raw so either is searchable
              if (display !== '') searchStr += display + ' ';
              // Add raw only if different from display (e.g. date serial vs formatted)
              if (raw !== '' && raw !== display) searchStr += raw + ' ';
            }
          }

          flat.push(searchStr.toLowerCase());
          compact.push(compStr);
          sheets.push(sName);
          rowNums.push(r + 1);
        }

        ws['!data'] = null; // free memory immediately
      }

    } else if (type === 'csv') {

      const text = await blob.text();
      let first  = true;
      let rn     = 1;

      Papa.parse(text, {
        skipEmptyLines: true,
        step: (res: any) => {
          const row: string[] = res.data;
          if (first) {
            for (const h of row) headers.push(String(h ?? '').trim());
            first = false;
            return;
          }
          rn++;
          let searchStr = '', compStr = '';
          for (let i = 0; i < row.length; i++) {
            const v = String(row[i] ?? '').trim();
            if (i > 0) compStr += SEP;
            compStr   += v;
            if (v !== '') searchStr += v + ' ';
          }
          flat.push(searchStr.toLowerCase());
          compact.push(compStr);
          sheets.push(null);
          rowNums.push(rn);
        },
      });

    } else {
      // TXT — one line per row
      const text  = await blob.text();
      const lines = text.split(/\r?\n/);
      headers.push('line');
      for (let i = 0; i < lines.length; i++) {
        const l = lines[i].trim();
        if (!l) continue;
        flat.push(l.toLowerCase());
        compact.push(l);
        sheets.push(null);
        rowNums.push(i + 1);
      }
    }

    if (!headers.length && flat.length) headers.push('col1');

    DB.set(id, { name, type, flat, compact, headers, sheets, rowNums });
    self.postMessage({ t: 'ok', id, n: flat.length });

  } catch (e: any) {
    self.postMessage({ t: 'err', id, msg: String(e?.message ?? e) });
  }
}

// ── SEARCH ────────────────────────────────────────────────────────────────────
function doSearch(msg: any) {
  const { sid, id } = msg;
  const entry = DB.get(id);

  if (!entry || !entry.flat.length) {
    self.postMessage({ t: 'done', sid, id });
    return;
  }

  const { flat, headers: hdrs, sheets: shts, rowNums, name } = entry;

  // Normalise query
  const q     = String(msg.q ?? '').toLowerCase().trim();
  const exact = !!msg.exact;
  const fuzzy = !!msg.fuzzy;
  const terms = q ? q.split(/\s+/).filter(t => t.length > 0) : [];
  const empty = q === '';

  // Compute priority preview column indices once per file
  const PRIO = [
    'date','time','amount','debit','credit','balance',
    'particular','narration','description','detail',
    'ref','chq','cheque','utr','name','account','remarks','note',
  ];
  const pi: number[] = [];
  for (let hi = 0; hi < hdrs.length && pi.length < 5; hi++) {
    const bare = hdrs[hi].toLowerCase().replace(/^[^:]+::/, '');
    if (PRIO.some(p => bare.includes(p))) pi.push(hi);
  }
  // Fallback: use first 4 columns
  if (!pi.length) {
    for (let i = 0; i < Math.min(4, hdrs.length); i++) pi.push(i);
  }

  const hits: any[] = [];
  const CHUNK = 400;

  for (let ri = 0; ri < flat.length; ri++) {
    const f = flat[ri];
    let sc  = 0;
    let cs: string[] | null = null;

    if (empty) {
      sc = 1;

    } else if (exact) {
      // Exact: any single cell equals query exactly (case-insensitive)
      cs = getCells(entry, ri);
      for (const c of cs) {
        if (c.toLowerCase() === q) { sc = 10; break; }
      }

    } else if (fuzzy) {
      // Fuzzy: ALL terms must appear somewhere in the row (substring)
      let ok = true;
      let s  = 0;
      for (const t of terms) {
        const pos = f.indexOf(t);
        if (pos === -1) { ok = false; break; }
        s += pos === 0 ? 3 : 1;
      }
      if (ok) sc = s || 1;

    } else {
      // Plain: entire query must appear as a substring
      const pos = f.indexOf(q);
      if (pos !== -1) sc = pos === 0 ? 3 : 1;
    }

    if (!sc) continue;

    if (!cs) cs = getCells(entry, ri);

    // Build preview from priority columns
    const parts: string[] = [];
    for (const pj of pi) {
      const v = cs[pj];
      if (v && v.trim()) {
        parts.push(hdrs[pj].replace(/^[^:]+::/, '') + ': ' + v);
      }
    }

    hits.push({
      id:  id + '-' + ri,
      fid: id,
      fn:  name,
      sn:  shts[ri],
      rn:  rowNums[ri],
      ss:  parts.length ? parts.join(' | ') : (cs[0] ?? '').slice(0, 140),
      sc,
      ri,
    });

    if (hits.length >= CHUNK) {
      self.postMessage({ t: 'hits', sid, hits: hits.splice(0) });
    }
  }

  if (hits.length) self.postMessage({ t: 'hits', sid, hits });
  self.postMessage({ t: 'done', sid, id });
}

// ── ROW DETAIL ────────────────────────────────────────────────────────────────
function doRow(msg: any) {
  const entry = DB.get(msg.id);
  if (!entry) {
    self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: null });
    return;
  }
  const cs = getCells(entry, msg.ri);
  const obj: Record<string, string> = {};
  for (let i = 0; i < entry.headers.length; i++) {
    const k = entry.headers[i].replace(/^[^:]+::/, '') || ('col' + i);
    obj[k] = cs[i] ?? '';
  }
  self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: obj });
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
function doExport(msg: any) {
  const hits: any[] = msg.hits ?? [];
  if (!hits.length) {
    self.postMessage({ t: 'csv', blob: new Blob([''], { type: 'text/csv' }) });
    return;
  }

  const lines: string[]       = [];
  const byFile                = new Map<string, any[]>();
  for (const h of hits) {
    if (!byFile.has(h.fid)) byFile.set(h.fid, []);
    byFile.get(h.fid)!.push(h);
  }

  let wrote = false;
  for (const [fid, fhits] of byFile) {
    const e = DB.get(fid);
    if (!e) continue;
    const bh = e.headers.map(h => h.replace(/^[^:]+::/, ''));
    if (!wrote) {
      lines.push(
        ['File', 'Sheet', 'Row', ...bh]
          .map(x => '"' + String(x).replace(/"/g, '""') + '"')
          .join(',')
      );
      wrote = true;
    }
    for (const m of fhits) {
      const cs = getCells(e, m.ri);
      lines.push([
        '"' + e.name.replace(/"/g, '""') + '"',
        '"' + (m.sn ?? '') + '"',
        m.rn,
        ...cs.map(v => '"' + String(v ?? '').replace(/"/g, '""') + '"'),
      ].join(','));
    }
  }

  self.postMessage({
    t: 'csv',
    blob: new Blob([lines.join('\n')], { type: 'text/csv' }),
  });
}
