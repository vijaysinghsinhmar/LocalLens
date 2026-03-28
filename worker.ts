// LocalLens Search Worker
// Imported as a Vite worker — XLSX and PapaParse are bundled at build time.
// No importScripts, no CDN, no CSP issues.

import * as XLSX from 'xlsx';
import Papa from 'papaparse';

const SEP = '\x1F';

interface DBEntry {
  name: string;
  type: string;
  flat: string[];
  compact: string[];
  headers: string[];
  sheets: (string|null)[];
  rowNums: number[];
}

const DB = new Map<string, DBEntry>();

function getCells(entry: DBEntry, ri: number): string[] {
  const s = entry.compact[ri];
  return s ? s.split(SEP) : [];
}

self.onmessage = async (ev: MessageEvent) => {
  const d = ev.data;
  if      (d.t === 'idx')    doIndex(d);
  else if (d.t === 'search') doSearch(d);
  else if (d.t === 'row')    doRow(d);
  else if (d.t === 'export') doExport(d);
  else if (d.t === 'clear')  DB.clear();
};

function doIndex(msg: any) {
  const { id, blob, ft: type, name } = msg;
  const flat:    string[]         = [];
  const compact: string[]         = [];
  const headers: string[]         = [];
  const sheets:  (string|null)[]  = [];
  const rowNums: number[]         = [];

  try {
    if (type === 'xlsx' || type === 'xls') {
      blob.arrayBuffer().then((buf: ArrayBuffer) => {
        try {
          const wb = XLSX.read(buf, {
            type: 'array', raw: true, dense: true,
            cellDates: false, cellNF: false, cellStyles: false,
            cellHTML: false, sheetStubs: false,
          });

          const multi = wb.SheetNames.length > 1;

          for (const sName of wb.SheetNames) {
            const ws = wb.Sheets[sName] as any;
            if (!ws || !ws['!data'] || ws['!data'].length < 2) {
              if (ws) ws['!data'] = null;
              continue;
            }
            const data   = ws['!data'] as any[][];
            const hdrRow = data[0] || [];
            const hdrs: string[] = hdrRow.map((c: any) =>
              c != null && c.v != null ? String(c.v) : ''
            );
            if (!hdrs.length) { ws['!data'] = null; continue; }

            if (!headers.length) {
              for (const h of hdrs) headers.push(multi ? sName + '::' + h : h);
            }

            for (let r = 1; r < data.length; r++) {
              const src = data[r];
              let f = '', cmp = '';
              if (src) {
                for (let c = 0; c < hdrs.length; c++) {
                  const cell = src[c] as any;
                  const v = (cell != null && cell.v != null) ? String(cell.v) : '';
                  if (v !== '') f += v + ' ';
                  if (c) cmp += SEP;
                  cmp += v;
                }
              }
              flat.push(f.toLowerCase());
              compact.push(cmp);
              sheets.push(sName);
              rowNums.push(r + 1);
            }
            ws['!data'] = null;
          }

          if (!headers.length && flat.length) headers.push('col1');
          DB.set(id, { name, type, flat, compact, headers, sheets, rowNums });
          self.postMessage({ t: 'ok', id, n: flat.length });
        } catch(e: any) {
          self.postMessage({ t: 'err', id, msg: String(e?.message || e) });
        }
      });

    } else if (type === 'csv') {
      blob.text().then((text: string) => {
        try {
          let first = true;
          let rn = 1;
          Papa.parse(text, {
            skipEmptyLines: true,
            step: (res: any) => {
              const row: string[] = res.data;
              if (first) {
                for (const h of row) headers.push(String(h || ''));
                first = false;
                return;
              }
              rn++;
              let f = '', cmp = '';
              for (let i = 0; i < row.length; i++) {
                const v = String(row[i] || '');
                if (v !== '') f += v + ' ';
                if (i) cmp += SEP;
                cmp += v;
              }
              flat.push(f.toLowerCase());
              compact.push(cmp);
              sheets.push(null);
              rowNums.push(rn);
            },
          });
          if (!headers.length && flat.length) headers.push('col1');
          DB.set(id, { name, type, flat, compact, headers, sheets, rowNums });
          self.postMessage({ t: 'ok', id, n: flat.length });
        } catch(e: any) {
          self.postMessage({ t: 'err', id, msg: String(e?.message || e) });
        }
      });

    } else {
      // TXT
      blob.text().then((text: string) => {
        try {
          headers.push('line');
          const lines = text.split('\n');
          for (let i = 0; i < lines.length; i++) {
            const l = lines[i].trim();
            if (!l) continue;
            flat.push(l.toLowerCase());
            compact.push(l);
            sheets.push(null);
            rowNums.push(i + 1);
          }
          DB.set(id, { name, type, flat, compact, headers, sheets, rowNums });
          self.postMessage({ t: 'ok', id, n: flat.length });
        } catch(e: any) {
          self.postMessage({ t: 'err', id, msg: String(e?.message || e) });
        }
      });
    }
  } catch(e: any) {
    self.postMessage({ t: 'err', id, msg: String(e?.message || e) });
  }
}

function doSearch(msg: any) {
  const { sid, id, q: rawQ, exact, fuzzy } = msg;
  const entry = DB.get(id);
  if (!entry || !entry.flat.length) {
    self.postMessage({ t: 'done', sid, id });
    return;
  }

  const { flat, headers: hdrs, sheets: shts, rowNums, name } = entry;
  const q     = String(rawQ || '').toLowerCase().trim();
  const terms = q ? q.split(/\s+/).filter(t => t.length > 0) : [];
  const empty = q === '';

  const PRIO = ['date','amount','debit','credit','balance','particular',
                'narration','description','ref','name','utr','account','remarks'];
  const pi: number[] = [];
  for (let hi = 0; hi < hdrs.length && pi.length < 4; hi++) {
    const bare = hdrs[hi].toLowerCase().replace(/^[^:]+::/, '');
    for (const pk of PRIO) {
      if (bare.includes(pk)) { pi.push(hi); break; }
    }
  }
  if (!pi.length) for (let i = 0; i < Math.min(4, hdrs.length); i++) pi.push(i);

  const hits: any[] = [];
  const CHUNK = 300;

  for (let ri = 0; ri < flat.length; ri++) {
    const f = flat[ri];
    let sc  = 0;
    let cs: string[]|null = null;

    if (empty) {
      sc = 1;
    } else if (exact) {
      cs = getCells(entry, ri);
      for (const c of cs) {
        if (c.toLowerCase() === q) { sc = 10; break; }
      }
    } else if (fuzzy) {
      let ok = true, s = 0;
      for (const t of terms) {
        const pos = f.indexOf(t);
        if (pos === -1) { ok = false; break; }
        s += pos === 0 ? 3 : 1;
      }
      if (ok) sc = s || 1;
    } else {
      const pos = f.indexOf(q);
      if (pos !== -1) sc = pos === 0 ? 3 : 1;
    }

    if (!sc) continue;
    if (!cs) cs = getCells(entry, ri);

    const parts: string[] = [];
    for (const pj of pi) {
      const v = cs[pj];
      if (v) parts.push(hdrs[pj].replace(/^[^:]+::/, '') + ': ' + v);
    }

    hits.push({
      id:  id + '-' + ri, fid: id, fn: name,
      sn:  shts[ri], rn: rowNums[ri],
      ss:  parts.length ? parts.join(' | ') : (cs[0] || '').slice(0, 120),
      sc, ri,
    });

    if (hits.length >= CHUNK) {
      self.postMessage({ t: 'hits', sid, hits: hits.splice(0) });
    }
  }

  if (hits.length) self.postMessage({ t: 'hits', sid, hits });
  self.postMessage({ t: 'done', sid, id });
}

function doRow(msg: any) {
  const entry = DB.get(msg.id);
  if (!entry) { self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: null }); return; }
  const cs = getCells(entry, msg.ri);
  if (!cs.length) { self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: null }); return; }
  const obj: Record<string,string> = {};
  for (let i = 0; i < entry.headers.length; i++) {
    obj[entry.headers[i].replace(/^[^:]+::/, '') || ('col' + i)] = cs[i] || '';
  }
  self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: obj });
}

function doExport(msg: any) {
  const hits = msg.hits || [];
  if (!hits.length) {
    self.postMessage({ t: 'csv', blob: new Blob([''], { type: 'text/csv' }) }); return;
  }
  const lines: string[] = [];
  const byFile = new Map<string, any[]>();
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
      lines.push(['File','Sheet','Row',...bh]
        .map(x => '"' + String(x).replace(/"/g, '""') + '"').join(','));
      wrote = true;
    }
    for (const m of fhits) {
      const cs = getCells(e, m.ri);
      lines.push([
        '"' + e.name.replace(/"/g, '""') + '"',
        '"' + (m.sn || '') + '"',
        m.rn,
        ...cs.map(v => '"' + String(v || '').replace(/"/g, '""') + '"'),
      ].join(','));
    }
  }
  self.postMessage({ t: 'csv', blob: new Blob([lines.join('\n')], { type: 'text/csv' }) });
}
