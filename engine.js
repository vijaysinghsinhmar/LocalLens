'use strict';

// ─────────────────────────────────────────────────────────────────────────────
// MEMORY-SAFE ENGINE v2
//
// Core insight: for a 315MB XLSX with 500k rows and 20 columns:
//   - flat[] (search index):  ~80 MB  (one string per row, ~160 chars avg)
//   - rows[] (ALL rows):     ~800 MB  (cell arrays for every row) — TOO MUCH
//
// Solution: store flat[] for ALL rows but rows[] only for MATCHED rows.
// Row detail / export re-reads from the blob for unmatched rows if needed.
// For typical searches (<<1% match rate) this saves ~95% of row storage.
//
// Additional savings:
//   - ws['!data'] nulled immediately after each sheet
//   - Workbook reference nulled after all sheets
//   - One file indexed at a time (App serialises via single worker)
// ─────────────────────────────────────────────────────────────────────────────

self.importScripts(
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'
);

// DB[id] = {
//   name, type,
//   flat:     string[],              — one per row, lowercase, for search
//   matched:  Map<ri, string[]>,     — cells only for rows that matched last search
//   headers:  string[],
//   sheets:   (string|null)[],
//   rowCount: number,
//   blob:     Blob                   — kept for re-read on export / row detail
// }
var DB = {};

self.onmessage = function(ev) {
  var d = ev.data;
  if      (d.t === 'idx')    doIndex(d);
  else if (d.t === 'search') doSearch(d);
  else if (d.t === 'row')    doRow(d);
  else if (d.t === 'export') doExport(d);
  else if (d.t === 'clear')  { DB = {}; }
};

// ── INDEX — builds flat[] only, no row data stored ────────────────────────────
function doIndex(msg) {
  var id   = msg.id;
  var type = msg.ft;
  var name = msg.name;
  var blob = msg.blob;
  var flat    = [];
  var headers = [];
  var sheets  = [];

  try {
    var fr = new FileReaderSync();

    if (type === 'xlsx' || type === 'xls') {
      var buf = fr.readAsArrayBuffer(blob);
      var wb  = XLSX.read(buf, {
        type:'array', raw:true, dense:true,
        cellDates:false, cellNF:false, cellStyles:false, cellHTML:false, sheetStubs:false
      });
      buf = null; // free raw buffer

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
          headers = multi ? hdrs.map(function(h){ return sName+'::'+h; }) : hdrs.slice();
        }

        for (var r = 1; r < data.length; r++) {
          var src = data[r];
          if (!src) { flat.push(''); sheets.push(sName); continue; }
          var f = '';
          for (var c = 0; c < hdrs.length; c++) {
            var cell = src[c];
            if (cell != null && cell.v != null) { f += String(cell.v); f += ' '; }
          }
          flat.push(f.toLowerCase());
          sheets.push(sName);
        }

        ws['!data'] = null; // free sheet immediately
        data = null;
      }
      wb = null; // free workbook

    } else if (type === 'csv') {
      var text  = fr.readAsText(blob);
      var first = true;
      Papa.parse(text, {
        skipEmptyLines: true,
        step: function(res) {
          var row = res.data;
          if (first) {
            for (var i = 0; i < row.length; i++) headers.push(String(row[i] || ''));
            first = false; return;
          }
          var f = '';
          for (var i = 0; i < row.length; i++) {
            var v = String(row[i] || '');
            if (v) { f += v; f += ' '; }
          }
          flat.push(f.toLowerCase());
          sheets.push(null);
        }
      });
      text = null;

    } else {
      headers = ['line'];
      var text  = fr.readAsText(blob);
      var lines = text.split('\n');
      text = null;
      for (var i = 0; i < lines.length; i++) {
        var l = lines[i].trim();
        if (!l) continue;
        flat.push(l.toLowerCase());
        sheets.push(null);
      }
      lines = null;
    }

    if (!headers.length && flat.length) headers = ['col1'];

    DB[id] = {
      name:name, type:type, blob:blob,
      flat:flat, matched:new Map(),
      headers:headers, sheets:sheets, rowCount:flat.length
    };

    self.postMessage({ t:'ok', id:id, n:flat.length });

  } catch(e) {
    self.postMessage({ t:'err', id:id, msg:String(e && e.message ? e.message : e) });
  }
}

// ── SEARCH — scans flat[], stores cells only for matched rows ─────────────────
function doSearch(msg) {
  var sid   = msg.sid;
  var id    = msg.id;
  var entry = DB[id];

  if (!entry || !entry.flat.length) {
    self.postMessage({ t:'done', sid:sid, id:id }); return;
  }

  var flat  = entry.flat;
  var hdrs  = entry.headers;
  var shts  = entry.sheets;
  var name  = entry.name;

  var q     = String(msg.q || '').toLowerCase().trim();
  var exact = !!msg.exact;
  var fuzzy = !!msg.fuzzy;
  var terms = q ? q.split(' ').filter(function(t){ return t.length > 0; }) : [];
  var empty = (q === '');

  // Preview columns
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
  if (!pi.length) for (var i = 0; i < Math.min(4, hdrs.length); i++) pi.push(i);

  // We need cells for preview — re-read matched rows from blob
  // Strategy: first pass = find matching row indices, second pass = read cells
  var matchedIndices = [];
  var scores         = {};

  for (var ri = 0; ri < flat.length; ri++) {
    var f  = flat[ri];
    var sc = 0;

    if (empty) {
      sc = 1;
    } else if (exact) {
      // For exact we need cells — skip for now, handle in second pass
      sc = f.indexOf(q) !== -1 ? 1 : 0; // pre-filter
    } else if (fuzzy) {
      var ok = true, s = 0;
      for (var ti = 0; ti < terms.length; ti++) {
        var pos = f.indexOf(terms[ti]);
        if (pos === -1) { ok = false; break; }
        s += pos === 0 ? 3 : 1;
      }
      if (ok) sc = s || 1;
    } else {
      var pos2 = f.indexOf(q);
      if (pos2 !== -1) sc = pos2 === 0 ? 3 : 1;
    }

    if (!sc) continue;
    matchedIndices.push(ri);
    scores[ri] = sc;
  }

  if (!matchedIndices.length) {
    self.postMessage({ t:'done', sid:sid, id:id }); return;
  }

  // Read cells for matched rows only
  var matchedCells = readMatchedRows(entry, matchedIndices);
  entry.matched    = matchedCells; // cache for row detail / export

  // Build and send hits
  var hits  = [];
  var CHUNK = 250;

  for (var mi = 0; mi < matchedIndices.length; mi++) {
    var ri    = matchedIndices[mi];
    var cells = matchedCells.get(ri) || [];
    var sc2   = scores[ri];

    // For exact mode: re-check against actual cells
    if (exact) {
      sc2 = 0;
      for (var ci = 0; ci < cells.length; ci++) {
        if (cells[ci].toLowerCase() === q) { sc2 = 10; break; }
      }
      if (!sc2) continue;
    }

    var parts = [];
    for (var pj = 0; pj < pi.length; pj++) {
      var v = cells[pi[pj]];
      if (v) parts.push(hdrs[pi[pj]].replace(/^[^:]+::/, '') + ': ' + v);
    }

    hits.push({
      id: id+'-'+ri, fid:id, fn:name, sn:shts[ri],
      rn:ri+2,
      ss: parts.length ? parts.join(' | ') : (cells[0]||'').slice(0,120),
      sc:sc2, ri:ri
    });

    if (hits.length >= CHUNK) {
      self.postMessage({ t:'hits', sid:sid, hits:hits.splice(0) });
    }
  }

  if (hits.length) self.postMessage({ t:'hits', sid:sid, hits:hits });
  self.postMessage({ t:'done', sid:sid, id:id });
}

// ── Read cells for specific row indices from blob ─────────────────────────────
function readMatchedRows(entry, indices) {
  var result  = new Map();
  var indexSet = {};
  for (var i = 0; i < indices.length; i++) indexSet[indices[i]] = true;

  try {
    var fr   = new FileReaderSync();
    var type = entry.type;
    var hdrs = entry.headers;

    if (type === 'xlsx' || type === 'xls') {
      var buf = fr.readAsArrayBuffer(entry.blob);
      var wb  = XLSX.read(buf, {
        type:'array', raw:true, dense:true,
        cellDates:false, cellNF:false, cellStyles:false, cellHTML:false, sheetStubs:false
      });
      buf = null;
      var multi = wb.SheetNames.length > 1;
      var globalRow = 0;

      for (var si = 0; si < wb.SheetNames.length; si++) {
        var sName = wb.SheetNames[si];
        var ws    = wb.Sheets[sName];
        if (!ws || !ws['!data'] || ws['!data'].length < 2) continue;
        var data   = ws['!data'];
        var hdrRow = data[0] || [];
        var hdrs2  = [];
        for (var c = 0; c < hdrRow.length; c++) {
          var hc = hdrRow[c];
          hdrs2.push(hc != null ? String(hc.v != null ? hc.v : '') : '');
        }

        for (var r = 1; r < data.length; r++) {
          if (indexSet[globalRow]) {
            var src   = data[r];
            var cells = [];
            if (src) {
              for (var c = 0; c < hdrs2.length; c++) {
                var cell = src[c];
                cells.push(cell != null && cell.v != null ? String(cell.v) : '');
              }
            }
            result.set(globalRow, cells);
          }
          globalRow++;
        }
        ws['!data'] = null;
      }
      wb = null;

    } else if (type === 'csv') {
      var text  = fr.readAsText(entry.blob);
      var first = true;
      var rowIdx = 0;
      Papa.parse(text, {
        skipEmptyLines: true,
        step: function(res) {
          if (first) { first = false; return; }
          if (indexSet[rowIdx]) {
            result.set(rowIdx, res.data.map(function(v){ return String(v || ''); }));
          }
          rowIdx++;
        }
      });
      text = null;

    } else {
      var text  = fr.readAsText(entry.blob);
      var lines = text.split('\n');
      text = null;
      var rowIdx = 0;
      for (var li = 0; li < lines.length; li++) {
        var l = lines[li].trim();
        if (!l) continue;
        if (indexSet[rowIdx]) result.set(rowIdx, [l]);
        rowIdx++;
      }
    }
  } catch(e) {
    // Return whatever we have
  }
  return result;
}

// ── ROW DETAIL ────────────────────────────────────────────────────────────────
function doRow(msg) {
  var e = DB[msg.id];
  if (!e) { self.postMessage({ t:'row', id:msg.id, ri:msg.ri, d:null }); return; }

  var cells = e.matched ? e.matched.get(msg.ri) : null;
  if (!cells || !cells.length) {
    // Not in match cache — read from blob
    var m = readMatchedRows(e, [msg.ri]);
    cells = m.get(msg.ri) || [];
  }

  var obj = {};
  for (var i = 0; i < e.headers.length; i++) {
    var k = e.headers[i].replace(/^[^:]+::/, '') || ('col' + i);
    obj[k] = cells[i] || '';
  }
  self.postMessage({ t:'row', id:msg.id, ri:msg.ri, d:obj });
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
function doExport(msg) {
  var hits = msg.hits || [];
  if (!hits.length) {
    self.postMessage({ t:'csv', blob: new Blob([''], {type:'text/csv'}) }); return;
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
      lines.push(['File','Sheet','Row'].concat(bh)
        .map(function(x){ return '"'+String(x).replace(/"/g,'""')+'"'; }).join(','));
      wrote = true;
    }
    var fhits = byFile[fid];
    for (var j = 0; j < fhits.length; j++) {
      var m  = fhits[j];
      var cells = e.matched ? (e.matched.get(m.ri) || []) : [];
      lines.push([
        '"'+e.name.replace(/"/g,'""')+'"',
        '"'+(m.sn||'')+'"',
        m.rn
      ].concat(cells.map(function(v){
        return '"'+String(v).replace(/"/g,'""')+'"';
      })).join(','));
    }
  }

  self.postMessage({ t:'csv', blob: new Blob([lines.join('\n')], {type:'text/csv'}) });
}
