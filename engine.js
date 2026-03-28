'use strict';

// ── LocalLens Search Engine ───────────────────────────────────────────────────
// Single worker handles everything: index, search, row detail, export.
// App uses 1 worker only — simplest and most reliable.
//
// Storage: flat[] for search, compact[] for display (null-safe SEP).
// Memory: ~300MB for 500k rows × 20 cols — acceptable for modern browsers.

self.importScripts(
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'
);

var SEP = '\x1F'; // ASCII unit separator — never appears in normal text

// DB: fileId → { name, flat, compact, headers, sheets, rowNums }
var DB = {};

self.onmessage = function(ev) {
  var d = ev.data;
  if      (d.t === 'idx')    doIndex(d);
  else if (d.t === 'search') doSearch(d);
  else if (d.t === 'row')    doRow(d);
  else if (d.t === 'export') doExport(d);
  else if (d.t === 'clear')  { DB = {}; }
};

function cells(entry, ri) {
  var s = entry.compact[ri];
  return s ? s.split(SEP) : [];
}

// ── INDEX ─────────────────────────────────────────────────────────────────────
function doIndex(msg) {
  var id   = msg.id;
  var type = msg.ft;
  var name = msg.name;

  var flat    = [];
  var compact = [];
  var headers = [];
  var sheets  = [];
  var rowNums = [];

  try {
    var fr = new FileReaderSync();

    if (type === 'xlsx' || type === 'xls') {
      var buf = fr.readAsArrayBuffer(msg.blob);
      var wb  = XLSX.read(buf, {
        type: 'array', raw: true, dense: true,
        cellDates: false, cellNF: false, cellStyles: false,
        cellHTML: false, sheetStubs: false
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
          hdrs.push(hc != null && hc.v != null ? String(hc.v) : '');
        }
        if (!hdrs.length) { ws['!data'] = null; continue; }

        // Set headers from first sheet (multi-sheet: prefix with sheet name)
        if (!headers.length) {
          for (var h = 0; h < hdrs.length; h++) {
            headers.push(multi ? sName + '::' + hdrs[h] : hdrs[h]);
          }
        }

        for (var r = 1; r < data.length; r++) {
          var src = data[r];
          var f   = '';
          var cmp = '';
          if (src) {
            for (var c = 0; c < hdrs.length; c++) {
              var cell = src[c];
              var v    = (cell != null && cell.v != null) ? String(cell.v) : '';
              if (v !== '') f += v + ' '; // include "0", "false" etc
              if (c) cmp += SEP;
              cmp += v;
            }
          }
          flat.push(f.toLowerCase());
          compact.push(cmp);
          sheets.push(sName);
          rowNums.push(r + 1); // actual 1-based row in this sheet
        }

        ws['!data'] = null; // free immediately
        data = null;
      }
      wb = null;

    } else if (type === 'csv') {
      var text  = fr.readAsText(msg.blob);
      var first = true;
      var rn    = 1;
      Papa.parse(text, {
        skipEmptyLines: true,
        step: function(res) {
          var row = res.data;
          if (first) {
            for (var i = 0; i < row.length; i++) headers.push(String(row[i] || ''));
            first = false;
            return;
          }
          rn++;
          var f = '', cmp = '';
          for (var i = 0; i < row.length; i++) {
            var v = String(row[i] || '');
            if (v !== '') f += v + ' ';
            if (i) cmp += SEP;
            cmp += v;
          }
          flat.push(f.toLowerCase());
          compact.push(cmp);
          sheets.push(null);
          rowNums.push(rn);
        }
      });
      text = null;

    } else {
      // TXT
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
        rowNums.push(i + 1);
      }
      lines = null;
    }

    if (!headers.length && flat.length) headers = ['col1'];

    DB[id] = { name: name, type: type,
               flat: flat, compact: compact,
               headers: headers, sheets: sheets, rowNums: rowNums };

    self.postMessage({ t: 'ok', id: id, n: flat.length });

  } catch(e) {
    self.postMessage({ t: 'err', id: id, msg: String(e && e.message ? e.message : e) });
  }
}

// ── SEARCH ────────────────────────────────────────────────────────────────────
function doSearch(msg) {
  var sid   = msg.sid;
  var id    = msg.id;
  var entry = DB[id];

  if (!entry || !entry.flat.length) {
    self.postMessage({ t: 'done', sid: sid, id: id });
    return;
  }

  var flat    = entry.flat;
  var hdrs    = entry.headers;
  var shts    = entry.sheets;
  var rowNums = entry.rowNums;
  var name    = entry.name;

  var q     = String(msg.q || '').toLowerCase().trim();
  var exact = !!msg.exact;
  var fuzzy = !!msg.fuzzy;
  var terms = q ? q.split(/\s+/).filter(function(t) { return t.length > 0; }) : [];
  var empty = (q === '');

  // Priority preview columns (computed once per file per search)
  var PRIO = ['date','amount','debit','credit','balance','particular',
              'narration','description','ref','name','utr','account','remarks'];
  var pi = [];
  for (var hi = 0; hi < hdrs.length && pi.length < 4; hi++) {
    var bare = hdrs[hi].toLowerCase().replace(/^[^:]+::/, '');
    for (var p = 0; p < PRIO.length; p++) {
      if (bare.indexOf(PRIO[p]) !== -1) { pi.push(hi); break; }
    }
  }
  if (!pi.length) for (var i = 0; i < Math.min(4, hdrs.length); i++) pi.push(i);

  var hits  = [];
  var CHUNK = 300;

  for (var ri = 0; ri < flat.length; ri++) {
    var f  = flat[ri];
    var sc = 0;
    var cs = null; // lazy cell split

    if (empty) {
      sc = 1;
    } else if (exact) {
      cs = cells(entry, ri);
      for (var ci = 0; ci < cs.length; ci++) {
        if (cs[ci].toLowerCase() === q) { sc = 10; break; }
      }
    } else if (fuzzy) {
      var ok = true, s = 0;
      for (var ti = 0; ti < terms.length; ti++) {
        var pos = f.indexOf(terms[ti]);
        if (pos === -1) { ok = false; break; }
        s += pos === 0 ? 3 : 1;
      }
      if (ok) sc = s || 1;
    } else {
      var idx = f.indexOf(q);
      if (idx !== -1) sc = idx === 0 ? 3 : 1;
    }

    if (!sc) continue;

    if (!cs) cs = cells(entry, ri);

    var parts = [];
    for (var pj = 0; pj < pi.length; pj++) {
      var v = cs[pi[pj]];
      if (v) parts.push(hdrs[pi[pj]].replace(/^[^:]+::/, '') + ': ' + v);
    }

    hits.push({
      id:  id + '-' + ri,
      fid: id,
      fn:  name,
      sn:  shts[ri],
      rn:  rowNums[ri],
      ss:  parts.length ? parts.join(' | ') : (cs[0] || '').slice(0, 120),
      sc:  sc,
      ri:  ri
    });

    if (hits.length >= CHUNK) {
      self.postMessage({ t: 'hits', sid: sid, hits: hits.splice(0) });
    }
  }

  if (hits.length) self.postMessage({ t: 'hits', sid: sid, hits: hits });
  self.postMessage({ t: 'done', sid: sid, id: id });
}

// ── ROW DETAIL ────────────────────────────────────────────────────────────────
function doRow(msg) {
  var e = DB[msg.id];
  if (!e) { self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: null }); return; }
  var cs = cells(e, msg.ri);
  if (!cs.length) { self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: null }); return; }
  var obj = {};
  for (var i = 0; i < e.headers.length; i++) {
    obj[e.headers[i].replace(/^[^:]+::/, '') || ('col' + i)] = cs[i] || '';
  }
  self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: obj });
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
function doExport(msg) {
  var hits = msg.hits || [];
  if (!hits.length) {
    self.postMessage({ t: 'csv', blob: new Blob([''], { type: 'text/csv' }) });
    return;
  }
  var lines  = [];
  var byFile = {};
  for (var i = 0; i < hits.length; i++) {
    if (!byFile[hits[i].fid]) byFile[hits[i].fid] = [];
    byFile[hits[i].fid].push(hits[i]);
  }
  var wrote = false;
  for (var fid in byFile) {
    if (!Object.prototype.hasOwnProperty.call(byFile, fid)) continue;
    var e = DB[fid];
    if (!e) continue;
    var bh = e.headers.map(function(h) { return h.replace(/^[^:]+::/, ''); });
    if (!wrote) {
      lines.push(['File','Sheet','Row'].concat(bh)
        .map(function(x) { return '"' + String(x).replace(/"/g, '""') + '"'; }).join(','));
      wrote = true;
    }
    var fhits = byFile[fid];
    for (var j = 0; j < fhits.length; j++) {
      var m  = fhits[j];
      var cs = cells(e, m.ri);
      lines.push([
        '"' + e.name.replace(/"/g, '""') + '"',
        '"' + (m.sn || '') + '"',
        m.rn
      ].concat(cs.map(function(v) {
        return '"' + String(v || '').replace(/"/g, '""') + '"';
      })).join(','));
    }
  }
  self.postMessage({ t: 'csv', blob: new Blob([lines.join('\n')], { type: 'text/csv' }) });
}
