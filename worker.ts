// Worker runs in its own scope — no React, no modules
// Blob URL created in App.tsx and passed to new Worker(url)

export const workerScript = `
(function() {
  'use strict';

  self.importScripts(
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'
  );

  // fileId -> { flat: string[], rows: string[][], headers: string[], sheets: (string|null)[], name: string }
  var DB = {};

  self.onmessage = function(ev) {
    var msg = ev.data;
    if      (msg.type === 'index')  doIndex(msg);
    else if (msg.type === 'search') doSearch(msg);
    else if (msg.type === 'row')    doRow(msg);
    else if (msg.type === 'export') doExport(msg);
    else if (msg.type === 'clear')  DB = {};
  };

  // ── INDEX ─────────────────────────────────────────────────────────────────
  function doIndex(msg) {
    var id   = msg.id;
    var blob = msg.blob;
    var type = msg.ftype;
    var name = msg.name;

    var flat    = [];
    var rows    = [];
    var sheets  = [];
    var headers = [];

    try {
      if (type === 'xlsx' || type === 'xls') {
        var buf = new Uint8Array(blobToArrayBufferSync(blob));
        var wb  = XLSX.read(buf, { type: 'array', raw: true, dense: true });
        var multi = wb.SheetNames.length > 1;

        for (var si = 0; si < wb.SheetNames.length; si++) {
          var sName = wb.SheetNames[si];
          var ws    = wb.Sheets[sName];
          if (!ws || !ws['!data'] || ws['!data'].length < 2) continue;

          var data   = ws['!data'];
          var hdrRow = data[0] || [];
          var hdrs   = [];
          for (var hi = 0; hi < hdrRow.length; hi++) {
            hdrs.push(hdrRow[hi] ? String(hdrRow[hi].v || '') : '');
          }

          if (!headers.length) {
            headers = multi ? hdrs.map(function(h){ return sName+'::'+h; }) : hdrs.slice();
          }

          for (var r = 1; r < data.length; r++) {
            var src   = data[r] || [];
            var cells = [];
            var f     = '';
            for (var c = 0; c < hdrs.length; c++) {
              var v = src[c] ? String(src[c].v || '') : '';
              cells.push(v);
              if (v) { f += v; f += '|'; }
            }
            flat.push(f.toLowerCase());
            rows.push(cells);
            sheets.push(sName);
          }
        }

      } else if (type === 'csv') {
        var text  = blobToTextSync(blob);
        var first = true;
        Papa.parse(text, {
          skipEmptyLines: true,
          step: function(result) {
            var data = result.data;
            if (first) {
              for (var i = 0; i < data.length; i++) headers.push(String(data[i] || ''));
              first = false;
              return;
            }
            var cells = [];
            var f = '';
            for (var i = 0; i < data.length; i++) {
              var v = String(data[i] || '');
              cells.push(v);
              if (v) { f += v; f += '|'; }
            }
            flat.push(f.toLowerCase());
            rows.push(cells);
            sheets.push(null);
          }
        });

      } else {
        // txt
        headers = ['line'];
        var text  = blobToTextSync(blob);
        var lines = text.split(/\r?\n/);
        for (var li = 0; li < lines.length; li++) {
          var t = lines[li].trim();
          if (!t) continue;
          flat.push(t.toLowerCase());
          rows.push([t]);
          sheets.push(null);
        }
      }

      DB[id] = { flat: flat, rows: rows, headers: headers, sheets: sheets, name: name };
      self.postMessage({ type: 'indexed', id: id, count: flat.length });

    } catch(e) {
      self.postMessage({ type: 'indexerr', id: id, err: String(e && e.message || e) });
    }
  }

  // ── SEARCH ────────────────────────────────────────────────────────────────
  function doSearch(msg) {
    var sid   = msg.sid;
    var id    = msg.id;
    var q     = (msg.query || '').toLowerCase().trim();
    var exact = msg.exact;
    var fuzzy = msg.fuzzy;
    var entry = DB[id];

    if (!entry) {
      self.postMessage({ type: 'searchdone', sid: sid, id: id });
      return;
    }

    var flat  = entry.flat;
    var rows  = entry.rows;
    var hdrs  = entry.headers;
    var shts  = entry.sheets;
    var name  = entry.name;

    // Priority columns for preview
    var PRIO  = ['date','amount','nature','debit','credit','balance','name',
                 'particulars','ref','account','description','narration','utr'];
    var prvIdx = [];
    for (var hi = 0; hi < hdrs.length && prvIdx.length < 4; hi++) {
      var bare = hdrs[hi].toLowerCase().replace(/^[^:]+::/, '');
      for (var pi = 0; pi < PRIO.length; pi++) {
        if (bare.indexOf(PRIO[pi]) !== -1) { prvIdx.push(hi); break; }
      }
    }
    if (!prvIdx.length) {
      for (var i = 0; i < Math.min(3, hdrs.length); i++) prvIdx.push(i);
    }

    var terms   = q ? q.split(/\s+/).filter(function(t){ return t.length > 0; }) : [];
    var empty   = !q;
    var matches = [];
    var CHUNK   = 250;

    for (var ri = 0; ri < flat.length; ri++) {
      var f     = flat[ri];
      var score = 0;

      if (empty) {
        score = 1;
      } else if (exact) {
        var cr = rows[ri];
        for (var ci = 0; ci < cr.length; ci++) {
          if (cr[ci].toLowerCase() === q) { score = 10; break; }
        }
      } else if (fuzzy) {
        var ok = true; var s = 0;
        for (var ti = 0; ti < terms.length; ti++) {
          var pos = f.indexOf(terms[ti]);
          if (pos === -1) { ok = false; break; }
          s += pos === 0 ? 3 : 1;
        }
        if (ok) score = s || 1;
      } else {
        var pos2 = f.indexOf(q);
        if (pos2 !== -1) score = pos2 === 0 ? 3 : 1;
      }

      if (!score) continue;

      var cells = rows[ri];
      var parts = [];
      for (var pj = 0; pj < prvIdx.length; pj++) {
        var cv = cells[prvIdx[pj]];
        if (cv) {
          var ch = hdrs[prvIdx[pj]].replace(/^[^:]+::/, '');
          parts.push(ch + ': ' + cv);
        }
      }

      matches.push({
        id:  id + '-' + ri,
        fid: id,
        fn:  name,
        sn:  shts[ri],
        rn:  ri + 2,
        ss:  parts.length ? parts.join(' • ') : (cells[0] || '').slice(0, 100),
        sc:  score,
        ri:  ri
      });

      if (matches.length >= CHUNK) {
        self.postMessage({ type: 'results', sid: sid, hits: matches.slice() });
        matches.length = 0;
      }
    }

    if (matches.length) {
      self.postMessage({ type: 'results', sid: sid, hits: matches });
    }
    self.postMessage({ type: 'searchdone', sid: sid, id: id });
  }

  // ── ROW DETAIL ────────────────────────────────────────────────────────────
  function doRow(msg) {
    var entry = DB[msg.id];
    if (!entry || !entry.rows[msg.ri]) {
      self.postMessage({ type: 'rowdata', id: msg.id, ri: msg.ri, data: null });
      return;
    }
    var cells = entry.rows[msg.ri];
    var hdrs  = entry.headers;
    var obj   = {};
    for (var i = 0; i < hdrs.length; i++) {
      obj[hdrs[i].replace(/^[^:]+::/, '') || ('col'+i)] = cells[i] || '';
    }
    self.postMessage({ type: 'rowdata', id: msg.id, ri: msg.ri, data: obj });
  }

  // ── EXPORT ────────────────────────────────────────────────────────────────
  function doExport(msg) {
    var hits = msg.hits || [];
    if (!hits.length) {
      self.postMessage({ type: 'exportready', blob: new Blob([''], { type: 'text/csv' }) });
      return;
    }
    var lines = [];
    var byFile = {};
    for (var i = 0; i < hits.length; i++) {
      var h = hits[i];
      if (!byFile[h.fid]) byFile[h.fid] = [];
      byFile[h.fid].push(h);
    }
    var wroteHdr = false;
    for (var fid in byFile) {
      var entry = DB[fid];
      if (!entry) continue;
      var bh = entry.headers.map(function(h){ return h.replace(/^[^:]+::/, ''); });
      if (!wroteHdr) {
        lines.push(['FileName','Sheet','Row'].concat(bh)
          .map(function(x){ return '"'+String(x).replace(/"/g,'""')+'"'; }).join(','));
        wroteHdr = true;
      }
      var fhits = byFile[fid];
      for (var j = 0; j < fhits.length; j++) {
        var m  = fhits[j];
        var cr = entry.rows[m.ri] || [];
        lines.push([
          '"'+entry.name.replace(/"/g,'""')+'"',
          '"'+(m.sn||'N/A')+'"',
          m.rn
        ].concat(cr.map(function(v){ return '"'+String(v).replace(/"/g,'""')+'"'; })).join(','));
      }
    }
    self.postMessage({ type: 'exportready',
      blob: new Blob([lines.join('\n')], { type: 'text/csv' }) });
  }

  // ── Sync helpers (workers can use synchronous XHR on blob URLs) ───────────
  function blobToArrayBufferSync(blob) {
    // FileReaderSync is available in workers
    var fr = new FileReaderSync();
    return fr.readAsArrayBuffer(blob);
  }

  function blobToTextSync(blob) {
    var fr = new FileReaderSync();
    return fr.readAsText(blob);
  }

})();
`;
