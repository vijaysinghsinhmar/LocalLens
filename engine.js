'use strict';

self.importScripts(
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'
);

// DB: fileId -> { name, flat[], rows[], headers[], sheets[] }
var DB = {};

self.onmessage = function(ev) {
  var d = ev.data;
  if      (d.t === 'idx')    doIndex(d);
  else if (d.t === 'search') doSearch(d);
  else if (d.t === 'row')    doRow(d);
  else if (d.t === 'export') doExport(d);
  else if (d.t === 'clear')  DB = {};
};

// ── INDEX ─────────────────────────────────────────────────────────────────────
function doIndex(msg) {
  var id   = msg.id;
  var type = msg.ft;
  var name = msg.name;
  var flat = [], rows = [], headers = [], sheets = [];

  try {
    var fr = new FileReaderSync();

    if (type === 'xlsx' || type === 'xls') {
      var buf = fr.readAsArrayBuffer(msg.blob);
      var wb  = XLSX.read(new Uint8Array(buf), {
        type: 'array', raw: true, dense: true,
        cellDates: false, cellNF: false, cellStyles: false
      });
      var multi = wb.SheetNames.length > 1;

      for (var si = 0; si < wb.SheetNames.length; si++) {
        var sName = wb.SheetNames[si];
        var ws    = wb.Sheets[sName];
        if (!ws || !ws['!data'] || ws['!data'].length < 2) continue;

        var data = ws['!data'];
        var hdrRow = data[0] || [];
        var hdrs = [];
        for (var hi = 0; hi < hdrRow.length; hi++) {
          var hc = hdrRow[hi];
          hdrs.push(hc ? String(hc.v != null ? hc.v : '') : '');
        }

        if (!headers.length) {
          headers = multi ? hdrs.map(function(h){ return sName + '::' + h; }) : hdrs.slice();
        }

        for (var r = 1; r < data.length; r++) {
          var src   = data[r];
          if (!src)  { flat.push(''); rows.push([]); sheets.push(sName); continue; }
          var cells = [];
          var f     = '';
          for (var c = 0; c < hdrs.length; c++) {
            var cell = src[c];
            var v    = cell != null ? String(cell.v != null ? cell.v : '') : '';
            cells.push(v);
            if (v) f += v + ' ';
          }
          flat.push(f.toLowerCase());
          rows.push(cells);
          sheets.push(sName);
        }
      }

    } else if (type === 'csv') {
      var text  = fr.readAsText(msg.blob);
      var first = true;
      Papa.parse(text, {
        skipEmptyLines: true,
        step: function(res) {
          var data = res.data;
          if (first) {
            for (var i = 0; i < data.length; i++) headers.push(String(data[i] || ''));
            first = false;
            return;
          }
          var cells = [], f = '';
          for (var i = 0; i < data.length; i++) {
            var v = String(data[i] || '');
            cells.push(v);
            if (v) f += v + ' ';
          }
          flat.push(f.toLowerCase());
          rows.push(cells);
          sheets.push(null);
        }
      });

    } else {
      // TXT
      headers = ['line'];
      var text  = fr.readAsText(msg.blob);
      var lines = text.split('\n');
      for (var li = 0; li < lines.length; li++) {
        var l = lines[li].trim();
        if (!l) continue;
        flat.push(l.toLowerCase());
        rows.push([l]);
        sheets.push(null);
      }
    }

    DB[id] = { name: name, flat: flat, rows: rows, headers: headers, sheets: sheets };
    self.postMessage({ t: 'ok', id: id, n: flat.length });

  } catch(e) {
    self.postMessage({ t: 'err', id: id, msg: String(e ? (e.message || e) : 'unknown') });
  }
}

// ── SEARCH ────────────────────────────────────────────────────────────────────
function doSearch(msg) {
  var sid   = msg.sid;
  var entry = DB[msg.id];
  if (!entry) {
    self.postMessage({ t: 'done', sid: sid, id: msg.id });
    return;
  }

  var flat  = entry.flat;
  var rows  = entry.rows;
  var hdrs  = entry.headers;
  var shts  = entry.sheets;
  var name  = entry.name;

  var q     = String(msg.q || '').toLowerCase().trim();
  var exact = !!msg.exact;
  var fuzzy = !!msg.fuzzy;
  var terms = q ? q.split(' ').filter(function(t){ return t.length > 0; }) : [];
  var empty = (q === '');

  // Pre-compute priority preview indices (once per search per file)
  var PRIO = ['date','amount','debit','credit','balance','name','particulars',
              'ref','narration','description','utr','account'];
  var pi = [];
  for (var hi = 0; hi < hdrs.length && pi.length < 4; hi++) {
    var bare = hdrs[hi].toLowerCase().replace(/^[^:]+::/, '');
    for (var p = 0; p < PRIO.length; p++) {
      if (bare.indexOf(PRIO[p]) !== -1) { pi.push(hi); break; }
    }
  }
  if (!pi.length) {
    for (var i = 0; i < Math.min(3, hdrs.length); i++) pi.push(i);
  }

  var hits  = [];
  var CHUNK = 300;

  for (var ri = 0; ri < flat.length; ri++) {
    var f  = flat[ri];
    var sc = 0;

    if (empty) {
      sc = 1;
    } else if (exact) {
      var cr = rows[ri];
      for (var ci = 0; ci < cr.length; ci++) {
        if (cr[ci].toLowerCase() === q) { sc = 10; break; }
      }
    } else if (fuzzy) {
      var ok = true, s = 0;
      for (var ti = 0; ti < terms.length; ti++) {
        var pos = f.indexOf(terms[ti]);
        if (pos === -1) { ok = false; break; }
        s += (pos === 0 ? 3 : 1);
      }
      if (ok) sc = (s || 1);
    } else {
      var idx = f.indexOf(q);
      if (idx !== -1) sc = (idx === 0 ? 3 : 1);
    }

    if (!sc) continue;

    var cells = rows[ri];
    var parts = [];
    for (var pj = 0; pj < pi.length; pj++) {
      var v = cells[pi[pj]];
      if (v) parts.push(hdrs[pi[pj]].replace(/^[^:]+::/, '') + ': ' + v);
    }

    hits.push({
      id:  msg.id + '-' + ri,
      fid: msg.id,
      fn:  name,
      sn:  shts[ri],
      rn:  ri + 2,
      ss:  parts.length ? parts.join(' | ') : (cells[0] || '').slice(0, 100),
      sc:  sc,
      ri:  ri
    });

    if (hits.length >= CHUNK) {
      self.postMessage({ t: 'hits', sid: sid, hits: hits.slice() });
      hits.length = 0;
    }
  }

  if (hits.length) self.postMessage({ t: 'hits', sid: sid, hits: hits });
  self.postMessage({ t: 'done', sid: sid, id: msg.id });
}

// ── ROW DETAIL ────────────────────────────────────────────────────────────────
function doRow(msg) {
  var e = DB[msg.id];
  if (!e || !e.rows[msg.ri]) {
    self.postMessage({ t: 'row', id: msg.id, ri: msg.ri, d: null });
    return;
  }
  var cells = e.rows[msg.ri];
  var obj   = {};
  for (var i = 0; i < e.headers.length; i++) {
    var k  = e.headers[i].replace(/^[^:]+::/, '') || ('col' + i);
    obj[k] = cells[i] || '';
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
    var h = hits[i];
    if (!byFile[h.fid]) byFile[h.fid] = [];
    byFile[h.fid].push(h);
  }

  var wroteHdr = false;
  for (var fid in byFile) {
    var e = DB[fid];
    if (!e) continue;
    var bh = e.headers.map(function(h){ return h.replace(/^[^:]+::/, ''); });
    if (!wroteHdr) {
      lines.push(['File','Sheet','Row'].concat(bh)
        .map(function(x){ return '"' + String(x).replace(/"/g, '""') + '"'; }).join(','));
      wroteHdr = true;
    }
    var fhits = byFile[fid];
    for (var j = 0; j < fhits.length; j++) {
      var m  = fhits[j];
      var cr = e.rows[m.ri] || [];
      lines.push([
        '"' + e.name.replace(/"/g, '""') + '"',
        '"' + (m.sn || 'N/A') + '"',
        m.rn
      ].concat(cr.map(function(v){
        return '"' + String(v).replace(/"/g, '""') + '"';
      })).join(','));
    }
  }

  self.postMessage({ t: 'csv', blob: new Blob([lines.join('\n')], { type: 'text/csv' }) });
}
