const XLSX = (function() {
  try { return require('xlsx'); }
  catch(e) {
    try { return require(require('path').join(require('os').homedir(), 'AppData/Roaming/npm/node_modules/xlsx/xlsx.js')); }
    catch(e2) {
      try { return require(require('path').join(require('os').homedir(), '.npm-global/node_modules/xlsx/xlsx.js')); }
      catch(e3) { throw new Error('Please install xlsx: npm install -g xlsx'); }
    }
  }
})();
const fs = require('fs');
const path = require('path');

var inputPath = process.argv[2];
if (!inputPath) {
  inputPath = 'D:/OpenCode/バージョン管理表.xlsx';
  if (!fs.existsSync(inputPath)) {
    console.error('Usage: node gen_html.js <path-to-xlsx>');
    process.exit(1);
  }
}

var outputPath = process.argv[3] || path.join(path.dirname(inputPath), path.basename(inputPath, path.extname(inputPath)) + '.html');

var wb = XLSX.readFile(inputPath);

const sheets = {};
wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const ref = XLSX.utils.decode_range(ws['!ref']);
    const headers = XLSX.utils.sheet_to_json(ws, {header: 1})[0];
    const rows = [];
    for (let r = ref.s.r + 1; r <= ref.e.r; r++) {
        const rw = [];
        for (let c = ref.s.c; c <= ref.e.c; c++) {
            const addr = XLSX.utils.encode_cell({r, c});
            const cell = ws[addr];
            let v = '';
            if (cell) {
                if (cell.w !== undefined) v = String(cell.w).trim();
                else if (cell.v !== undefined) v = String(cell.v);
            }
            if (c === ref.s.c && cell && typeof cell.v === 'number' && cell.v > 44000 && cell.v < 48000) {
                const d = new Date((cell.v - 25569) * 86400000);
                v = d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + String(d.getDate()).padStart(2, '0');
            }
            rw.push(v);
        }
        if (rw.some(x => x !== '')) rows.push(rw);
    }
    sheets[name] = { h: headers, r: rows };
});

function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&#34;');
}

// Build data JSONs
var jsonData = JSON.stringify(Object.keys(sheets).map(function(n) { return sheets[n].r; }));
var jsonHeaders = JSON.stringify(Object.keys(sheets).map(function(n) { return sheets[n].h; }));
var jsonNames = JSON.stringify(Object.keys(sheets));

// Tabs HTML
var tabsHtml = '';
Object.keys(sheets).forEach(function(n, i) {
    tabsHtml += '<button class="tab' + (i === 0 ? ' on' : '') + '" data-t="' + i + '">' + esc(n) + ' <span class="n">' + sheets[n].r.length + '</span></button>';
});

// Tables HTML + per-sheet column selectors
var tablesHtml = '';
Object.keys(sheets).forEach(function(n, si) {
    var s = sheets[n];
    var colsHtml = '<option value="">すべての列</option>';
    s.h.forEach(function(h) {
        colsHtml += '<option value="' + esc(h) + '">' + esc(h) + '</option>';
    });

    var rowsHtml = '';
    s.r.forEach(function(row) {
        var cells = '';
        row.forEach(function(v, ci) {
            var isMono = (ci === 0 && /^\d{4}\/\d{2}\/\d{2}$/.test(v)) || (ci > 0 && ci < 3 && /^[\d.]+$/.test(v));
            if (isMono) cells += '<td><span class="m">' + esc(v) + '</span></td>';
            else cells += '<td>' + esc(v) + '</td>';
        });
        rowsHtml += '<tr>' + cells + '</tr>';
    });

    var hdHtml = '';
    s.h.forEach(function(h) { hdHtml += '<th>' + esc(h) + '</th>'; });

    tablesHtml += '<div class="sec' + (si === 0 ? ' on' : '') + '" data-s="' + si + '">'
        + '<div class="srow" data-sr="' + si + '">'
        + '<select class="cs" data-cs="' + si + '">' + colsHtml + '</select>'
        + '<input class="q" data-q="' + si + '" placeholder="検索...">'
        + '<span class="cr" id="cr-' + si + '"></span>'
        + '</div>'
        + '<div class="wrap"><table><thead><tr>' + hdHtml + '</tr></thead><tbody>' + rowsHtml + '</tbody></table></div>'
        + '</div>';
});

var html = '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>NASDA バージョン管理表</title>'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Hiragino Sans","Yu Gothic",sans-serif;background:#eef1f5;color:#1f1f1f;padding:24px}'
    + '.app{max-width:1640px;margin:0 auto}'

    // Header
    + 'header{background:linear-gradient(135deg,#1a3a5c 0%,#1565c0 100%);border-radius:14px;padding:28px 32px;margin-bottom:22px;color:#fff}'
    + 'header h1{font-size:24px;font-weight:700;letter-spacing:.3px}'
    + 'header .sub{font-size:13px;margin-top:5px;opacity:.78}'

    // Tabs
    + '.bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px}'
    + '.tab{padding:9px 22px;border-radius:22px;border:1px solid #d0d5dd;background:#fff;cursor:pointer;font-size:14px;font-weight:500;color:#475467;transition:all .15s;box-shadow:0 1px 2px rgba(0,0,0,.04)}'
    + '.tab:hover{background:#f5f6fa;border-color:#b0b7c3}'
    + '.tab.on{background:#1565c0;color:#fff;border-color:#1565c0;box-shadow:0 2px 8px rgba(21,101,192,.25)}'
    + '.tab .n{display:inline-block;margin-left:6px;padding:0 8px;border-radius:10px;font-size:11px;background:rgba(0,0,0,.07)}'
    + '.tab.on .n{background:rgba(255,255,255,.18)}'

    // Search row
    + '.srow{display:flex;gap:8px;align-items:center;margin-bottom:12px}'
    + '.cs{padding:8px 12px;border:1px solid #d0d5dd;border-radius:8px;font-size:13px;background:#fff;min-width:130px;color:#1f1f1f;cursor:pointer}'
    + '.cs:focus{outline:none;border-color:#1565c0;box-shadow:0 0 0 3px rgba(21,101,192,.1)}'
    + '.q{flex:1;min-width:160px;padding:8px 14px;border:1px solid #d0d5dd;border-radius:8px;font-size:13px;outline:none;transition:border .15s}'
    + '.q:focus{border-color:#1565c0;box-shadow:0 0 0 3px rgba(21,101,192,.1)}'
    + '.cr{font-size:13px;color:#475467;white-space:nowrap;min-width:80px;text-align:right;font-weight:500}'

    // Content
    + '.sec{display:none}'
    + '.sec.on{display:block}'
    + '.wrap{background:#fff;border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.07);overflow-x:auto}'
    + 'table{width:100%;border-collapse:collapse;font-size:13.5px}'
    + 'th{background:#f8f9fc;padding:12px 16px;text-align:left;font-weight:600;border-bottom:2px solid #e4e7ec;white-space:nowrap;color:#344054}'
    + 'td{padding:10px 16px;vertical-align:top;border-bottom:1px solid #eaecf0;line-height:1.5}'
    + 'tr:hover td{background:#f8f9fc}'
    + 'tr.h{display:none}'
    + '.m{font-family:"SF Mono","Consolas","Menlo",monospace;font-size:12px;background:#f2f4f7;padding:2px 10px;border-radius:4px;white-space:nowrap;color:#344054}'

    // Footer
    + '.ft{text-align:center;padding:24px;font-size:12px;color:#98a2b3;margin-top:8px}'

    // Responsive
    + '@media(max-width:640px){body{padding:12px}header{padding:20px 24px}header h1{font-size:20px}.srow{flex-wrap:wrap}.cs{min-width:100px}.q{min-width:120px}td,th{padding:8px 12px}}'
    + '</style></head><body><div class="app">'
    + '<header><h1>NASDA バージョン管理表</h1><div class="sub">Nelson Project — ' + esc(wb.SheetNames.join(' · ')) + '</div></header>'
    + '<div class="bar">' + tabsHtml + '</div>'
    + tablesHtml
    + '<div class="ft"><strong>NASDA バージョン管理表</strong> &middot; Nelson Project<br><span style="opacity:.6">SVN: /svn/smartauto/Nelson/01_開発庫/22_Release/バージョン管理表.xlsx</span></div>'
    + '</div>'
    + '<script>'
    + 'var SN=' + jsonNames + ';'
    + 'var H=' + jsonHeaders + ';'
    + 'var D=' + jsonData + ';'
    + 'var P=0;'

    // Tab switching
    + 'document.querySelectorAll(".tab").forEach(function(b,i){b.addEventListener("click",function(){'
    + 'P=+this.dataset.t;'
    + 'document.querySelectorAll(".tab").forEach(function(e){e.classList.toggle("on",+e.dataset.t===P)});'
    + 'document.querySelectorAll(".sec").forEach(function(e){e.classList.toggle("on",+e.dataset.s===P)});'
    + 'doSearch(P);'
    + '})});'

    // Search function
    + 'function doSearch(si){'
    + 'var q=document.querySelector(\'.q[data-q="\'+si+\'"]\');if(!q)return;'
    + 'var v=q.value.toLowerCase();'
    + 'var cs=document.querySelector(\'.cs[data-cs="\'+si+\'"]\');'
    + 'var c=cs?cs.value:"";'
    + 'var sec=document.querySelector(\'.sec[data-s="\'+si+\'"]\');if(!sec)return;'
    + 'var rows=sec.querySelectorAll("tbody tr");'
    + 'var n=0;'
    + 'rows.forEach(function(tr){'
    + 'var ok=!v;'
    + 'if(v){'
    + 'var cells=Array.from(tr.cells).map(function(x){return x.textContent});'
    + 'if(c){var idx=H[si].indexOf(c);ok=idx>=0&&idx<cells.length&&cells[idx].toLowerCase().indexOf(v)!==-1}'
    + 'else{ok=cells.some(function(x){return x.toLowerCase().indexOf(v)!==-1})}'
    + '}'
    + 'tr.classList.toggle("h",!ok);if(ok)n++;'
    + '});'
    + 'document.getElementById("cr-"+si).textContent=n+"/"+D[si].length+"件";'
    + '}'

    // Bind search events for each sheet
    + 'for(var si=0;si<SN.length;si++){(function(si){'
    + 'var q=document.querySelector(\'.q[data-q="\'+si+\'"]\');'
    + 'var cs=document.querySelector(\'.cs[data-cs="\'+si+\'"]\');'
    + 'if(q)q.addEventListener("input",function(){doSearch(si)});'
    + 'if(cs)cs.addEventListener("change",function(){doSearch(si)});'
    + '})(si)}'

    // Click row filter on search on current tab
    + 'doSearch(0);'
    + '</script></body></html>';

fs.writeFileSync(outputPath, html, 'utf8');
console.log('Generated: ' + outputPath);
console.log('Size: ' + html.length + ' bytes');
console.log('Sheets: ' + Object.keys(sheets).join(', '));
console.log('Rows: ' + Object.values(sheets).reduce(function(s, v) { return s + v.r.length; }, 0));