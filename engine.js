'use strict';

// ─────────────────────────────────────────────────────────────────────────────
// LocalLens Engine — fixed all critical bugs:
//
// ARCHITECTURE CHANGE:
//   Old: store flat[] only, re-parse blob on every search for cell data → OOM
//   New: store flat[] + rows[] during index, but use compact storage:
//        - rows stored as a SINGLE joined string per row (pipe-separated)
//        - split back into cells only when needed (row detail / preview)
//        - this halves object overhead vs string[][] (no inner array per row)
//
// MEMORY: 500k rows × 20 cols × 15 chars = ~150MB flat + ~150MB compact rows
//         vs old approach of re-parsing 315MB XLSX on every search keystroke
//
// BUGS FIXED:
//   1. Removed readMatchedRows() — no re-parse on search (was causing OOM)
//   2. terms split on /\s+/ not ' ' — handles multiple spaces
//   3. Export uses stored row data not entry.matched (which gets overwritten)
//   4. Cleanup: blob ref not held in DB after indexing (no longer needed)
//   5. Single-char highlight fixed (min length 1 not 2)
//   6. tickT/debT cleanup on worker terminate handled via App fixes
// ─────────────────────────────────────────────────────────────────────────────

self.importScripts(
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'
);

var SEP = '\x00'; // separator between cells in compact row string

// DB[id] = { name, flat:string[], compact:string[], headers:string[], sheets:(string|null)[] }
// compact[ri] = cells joined by SEP — split only when needed
var DB = {};

self.onmessage = function(ev) {
  var d = ev.data;
  if      (d.t === 'idx')    doIndex(d);
  else if (d.t === 'search') doSearch(d);
  else if (d.t === 'row')    doRow(d);
  else if (d.t === 'export') doExport(d);
  else if (d.t === 'clear')  { DB = {}; }
};

// ── Helpers ───────────────────────────────────────────────────────────────────
function getCells(entry, ri) {
  var c = entry.compact[ri];
  return c ? c.split(SEP) : [];
}

// ── INDEX ─────────────────────────────────────────────────────────────────────
function doIndex(msg) {
  var id   = msg.id;
  var type = msg.ft;
  var name = msg.name;
  var flat    = [];
  var compact = []; // one string per row: "cell1\x00cell2\x00cell3"
  var headers = [];
  var sheets  = [];

  try {
    var fr = new FileReaderSync();

    if (type === 'xlsx' || type === 'xls') {
      var buf = fr.readAsArrayBuffer(msg.blob);
      var wb  = XLSX.read(buf, {
        type:'array', raw:true, dense:true,
        cellDates:false, cellNF:false, cellStyles:false, cellHTML:false, sheetStubs:false
      });
      buf = null;

      var multi = wb.SheetNames.length > 1;

      for (var si = 0; si < wb.SheetNames.length; si++) {
        var sName = wb.SheetNames[si];
        var ws    = wb.Sheets[sName];
        if (!ws || !ws['!data'] || ws['!data'].length < 2) {
          if (ws) ws['!data'] = null;
          continue;
        }

        var data   = ws['!data'];
        var hdrRow = data[0] || [];
        var hdrs   = [];
        for (var c = 0; c < hdrRow.length; c++) {
          var hc = hdrRow[c];
          hdrs.push(hc != null ? String(hc.v != null ? hc.v : '') : '');
        }
        if (!hdrs.length) { ws['!data'] = null; continue; }

        if (!headers.length) {
          headers = multi
            ? hdrs.map(function(h){ return sName + '::' + h; })
            : hdrs.slice();
        }

        for (var r = 1; r < data.length; r++) {
          var src = data[r];
          if (!src) {
            flat.push('');
            compact.push('');
            sheets.push(sName);
            continue;
          }
          var f    = '';
          var comp = '';
          for (var c = 0; c < hdrs.length; c++) {
            var cell = src[c];
            var v    = (cell != null && cell.v != null) ? String(cell.v) : '';
            if (v) f += v + ' ';
            comp += v;
            if (c < hdrs.length - 1) comp += SEP;
          }
          flat.push(f.toLowerCase());
          compact.push(comp);
          sheets.push(sName);
        }

        ws['!data'] = null;
        data = null;
      }
      wb = null;

    } else if (type === 'csv') {
      var text  = fr.readAsText(msg.blob);
      var first = true;
      Papa.parse(text, {
        skipEmptyLines: true,
        step: function(res) {
          var row = res.data;
          if (first) {
            for (var i = 0; i < row.length; i++) headers.push(String(row[i] || ''));
            first = false;
            return;
          }
          var f = '', comp = '';
          for (var i = 0; i < row.length; i++) {
            var v = String(row[i] || '');
            if (v) f += v + ' ';
            comp += v;
            if (i < row.length - 1) comp += SEP;
          }
          flat.push(f.toLowerCase());
          compact.push(comp);
          sheets.push(null);
        }
      });
      text = null;

    } else {
      headers = ['line'];
      var text  = fr.readAsText(msg.blob);
      var lines = text.split('\n');
      text = null;
      for (var i = 0; i < lines.length; i++) {
        var l = lines[i].trim();
        if (!l) continue;
        flat.push(l.toLowerCase());
        compact.push(l);
        sheets.push(null);
      }
      lines = null;
    }

    if (!headers.length && flat.length) headers = ['col1'];

    // Do NOT store blob — no longer needed, avoids holding 315MB File ref
    DB[id] = { name:name, type:type, flat:flat, compact:compact, headers:headers, sheets:sheets };

    self.postMessage({ t:'ok', id:id, n:flat.length });

  } catch(e) {
    self.postMessage({ t:'err', id:id, msg:String(e && e.message ? e.message : e) });
  }
}

// ── SEARCH — single pass, no re-parse ─────────────────────────────────────────
function doSearch(msg) {
  var sid   = msg.sid;
  var id    = msg.id;
  var entry = DB[id];

  if (!entry || !entry.flat.length) {
    self.postMessage({ t:'done', sid:sid, id:id });
    return;
  }

  var flat    = entry.flat;
  var compact = entry.compact;
  var hdrs    = entry.headers;
  var shts    = entry.sheets;
  var name    = entry.name;

  var q     = String(msg.q || '').toLowerCase().trim();
  var exact = !!msg.exact;
  var fuzzy = !!msg.fuzzy;
  // FIX: use /\s+/ not ' ' to handle multiple spaces / tabs
  var terms = q ? q.split(/\s+/).filter(function(t){ return t.length > 0; }) : [];
  var empty = (q === '');

  // Priority preview columns — computed once per file, not per row
  var PRIO = ['date','amount','debit','credit','balance','particular',
              'narration','description','ref','name','utr','account','remarks'];
  var pi = [];
  for (var hi = 0; hi < hdrs.length; hi++) {
    if (pi.length >= 4) break;
    var bare = hdrs[hi].toLowerCase().replace(/^[^:]+::/, '');
    for (var p = 0; p < PRIO.length; p++) {
      if (bare.indexOf(PRIO[p]) !== -1) { pi.push(hi); break; }
    }
  }
  if (!pi.length) {
    for (var i = 0; i < Math.min(4, hdrs.length); i++) pi.push(i);
  }

  var hits  = [];
  var CHUNK = 250;

  for (var ri = 0; ri < flat.length; ri++) {
    var f  = flat[ri];
    var sc = 0;

    if (empty) {
      sc = 1;
    } else if (exact) {
      // Exact: any cell value equals query exactly
      var cells0 = getCells(entry, ri);
      for (var ci = 0; ci < cells0.length; ci++) {
        if (cells0[ci].toLowerCase() === q) { sc = 10; break; }
      }
    } else if (fuzzy) {
      var ok = true, s = 0;
      for (var ti = 0; ti < terms.length; ti++) {
        var pos = f.indexOf(terms[ti]);
        if (pos === -1) { ok = false; break; }
        s += (pos === 0 ? 3 : 1);
      }
      if (ok) sc = s || 1;
    } else {
      var pos2 = f.indexOf(q);
      if (pos2 !== -1) sc = (pos2 === 0 ? 3 : 1);
    }

    if (!sc) continue;

    // Build preview from compact row (split only for matched rows)
    var cells = getCells(entry, ri);
    var parts = [];
    for (var pj = 0; pj < pi.length; pj++) {
      var v = cells[pi[pj]];
      if (v) parts.push(hdrs[pi[pj]].replace(/^[^:]+::/, '') + ': ' + v);
    }

    hits.push({
      id:  id + '-' + ri,
      fid: id,
      fn:  name,
      sn:  shts[ri],
      rn:  ri + 2,
      ss:  parts.length ? parts.join(' | ') : (cells[0] || '').slice(0, 120),
      sc:  sc,
      ri:  ri
    });

    if (hits.length >= CHUNK) {
      self.postMessage({ t:'hits', sid:sid, hits:hits.splice(0) });
    }
  }

  if (hits.length) self.postMessage({ t:'hits', sid:sid, hits:hits });
  self.postMessage({ t:'done', sid:sid, id:id });
}

// ── ROW DETAIL ────────────────────────────────────────────────────────────────
function doRow(msg) {
  var e = DB[msg.id];
  if (!e) { self.postMessage({ t:'row', id:msg.id, ri:msg.ri, d:null }); return; }

  var cells = getCells(e, msg.ri);
  if (!cells.length) { self.postMessage({ t:'row', id:msg.id, ri:msg.ri, d:null }); return; }

  var obj = {};
  for (var i = 0; i < e.headers.length; i++) {
    var k = e.headers[i].replace(/^[^:]+::/, '') || ('col' + i);
    obj[k] = cells[i] || '';
  }
  self.postMessage({ t:'row', id:msg.id, ri:msg.ri, d:obj });
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
// FIX: reads from compact[] using ri from each hit — not from entry.matched
// which was overwritten on every search
function doExport(msg) {
  var hits = msg.hits || [];
  if (!hits.length) {
    self.postMessage({ t:'csv', blob: new Blob([''], {type:'text/csv'}) });
    return;
  }

  var lines  = [];
  var byFile = {};
  for (var i = 0; i < hits.length; i++) {
    var h = hits[i];
    if (!byFile[h.fid]) byFile[h.fid] = [];
    byFile[h.fid].push(h);
  }

  var wrote = false;
  for (var fid in byFile) {
    var e = DB[fid];
    if (!e) continue;
    var bh = e.headers.map(function(hh){ return hh.replace(/^[^:]+::/, ''); });
    if (!wrote) {
      lines.push(['File', 'Sheet', 'Row'].concat(bh)
        .map(function(x){ return '"' + String(x).replace(/"/g, '""') + '"'; }).join(','));
      wrote = true;
    }
    var fhits = byFile[fid];
    for (var j = 0; j < fhits.length; j++) {
      var m     = fhits[j];
      var cells = getCells(e, m.ri); // FIX: read from compact[], not entry.matched
      lines.push([
        '"' + e.name.replace(/"/g, '""') + '"',
        '"' + (m.sn || '') + '"',
        m.rn
      ].concat(cells.map(function(v){
        return '"' + String(v || '').replace(/"/g, '""') + '"';
      })).join(','));
    }
  }

  self.postMessage({ t:'csv', blob: new Blob([lines.join('\n')], {type:'text/csv'}) });
}
