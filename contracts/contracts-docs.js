/* contracts-docs.js — fills the commitment attachments (Contracts desk, 19 Nov 2026).
   Works in the browser (window.BHDocs) and in node (module.exports) for testing.
   Needs JSZip (window.JSZip in the browser, require('jszip') in node).
   Text-level edits only: whole paragraphs are rewritten into one run that keeps the
   paragraph's first run formatting; yellow highlight is removed from the output. */
(function (root) {
  'use strict';
  var JSZipRef = (typeof module !== 'undefined' && module.exports) ? require('jszip') : root.JSZip;

  var esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var unesc = function (s) { return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); };
  var norm = function (s) { return String(s).replace(/[‐-―]/g, '-').replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/\s+/g, ' ').trim(); };

  /* ---------- body element list (top-level <w:p> and <w:tbl>) ---------- */
  function splitBody(xml) {
    var a = xml.indexOf('<w:body>') + 8, b = xml.lastIndexOf('</w:body>');
    var body = xml.slice(a, b), out = [], re = /<(\/?)w:(p|tbl|sectPr)([ >\/])/g, m, lvl = 0, kind = null, start = 0;
    while ((m = re.exec(body))) {
      var close = m[1] === '/', tag = m[2];
      if (lvl === 0 && !close) { start = m.index; kind = tag; }
      if (tag !== kind) continue;
      if (!close) {
        var gt = body.indexOf('>', m.index);
        if (body[gt - 1] === '/') { if (lvl === 0) { out.push(body.slice(start, gt + 1)); kind = null; } continue; }
        lvl++;
      } else {
        lvl--;
        if (lvl === 0) { var e = body.indexOf('>', m.index) + 1; out.push(body.slice(start, e)); kind = null; }
      }
    }
    return { head: xml.slice(0, a), els: out, tail: xml.slice(b) };
  }
  function joinBody(d) { return d.head + d.els.join('') + d.tail; }
  function textOf(el) { var t = '', re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g, m; while ((m = re.exec(el))) t += m[1]; return unesc(t); }

  /* rewrite a paragraph's text: keep pPr and the first run's rPr */
  function setParaText(p, text) {
    var pPr = (p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
    var rPr = (p.match(/<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/) || [, ''])[1];
    var open = p.match(/^<w:p(?:\s[^>]*)?>/)[0];
    return open + pPr + '<w:r>' + rPr + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
  }
  function replaceInPara(p, find, repl) {
    var t = textOf(p); if (t.indexOf(find) < 0) return p;
    return setParaText(p, t.split(find).join(repl));
  }
  function stripHighlight(xml) { return xml.replace(/<w:highlight [^>]*\/>/g, ''); }

  function find(els, pred, from) { for (var i = from || 0; i < els.length; i++) if (pred(norm(textOf(els[i])), els[i])) return i; return -1; }
  var starts = function (s) { s = norm(s); return function (t) { return t.indexOf(s) === 0; }; };

  /* ---------- Attachment B ---------- */
  /* d = { projectNumName, projectCity, subName, taxExempt, ownerPlatform, bondRequired,
           builderRiskOwner, projectItems:[..], alternates:[{label,status}], scopeLines:[{desc,amount}],
           submittalDate, scopeName, csi:[..], inclusions:[..], exclusions:[..],
           submitDay:'10th', billingDate:'20th', retPct:5, retDuration:'Thru 50%' } */
  async function buildAttachmentB(bytes, d) {
    var zip = await JSZipRef.loadAsync(bytes);
    var xml = await zip.file('word/document.xml').async('string');
    var D = splitBody(xml), els = D.els, i, j;
    var money = function (v) { return '$' + Number(v || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
    var log = [];

    // header placeholders
    els = els.map(function (p) {
      if (p.indexOf('<w:p') !== 0) return p;
      var t = textOf(p);
      if (t.indexOf('<<') < 0) return p;
      t = t.replace(/<<\s*Project Number and Name\s*>>/g, d.projectNumName || '')
           .replace(/<<\s*Project City, State\s*>>/g, d.projectCity || '')
           .replace(/<<\s*Subcontractor Name\s*>>/g, d.subName || '')
           .replace(/<<\s*Scope of Work\s*>>/g, d.scopeName || '');
      return setParaText(p, t);
    });

    // general item 4 (sales tax exemption) and item 21 (Procore)
    if (!d.taxExempt) { i = find(els, starts('This Subcontract Agreement does not include Sales Taxes')); if (i >= 0) { els.splice(i, 1); log.push('removed general item: sales tax exemption'); } }
    if (d.ownerPlatform) { i = find(els, starts('Unless otherwise noted, Procore will be utilized')); if (i >= 0) { els.splice(i, 1); log.push('removed general item: Procore'); } }

    // project specific general items: replace the examples with the PC's items
    var ps = find(els, starts('Project Specific General Items')), pe = find(els, starts('Subcontract Scope'), ps + 1);
    if (ps >= 0 && pe > ps) {
      var firstItem = -1;
      for (i = ps + 1; i < pe; i++) if (/<w:numPr>/.test(els[i]) && /<w:ilvl w:val="1"\/>/.test(els[i])) { firstItem = i; break; }
      var tmpl = firstItem >= 0 ? els[firstItem] : null;
      var items = [];
      if (d.builderRiskOwner && tmpl) items.push(textOf(tmpl)); // the template's first item is the Owner builder's risk sentence
      (d.projectItems || []).forEach(function (s) { if (String(s).trim()) items.push(String(s).trim()); });
      var lastItem = firstItem; for (i = firstItem; i < pe; i++) if (/<w:numPr>/.test(els[i])) lastItem = i;
      if (tmpl) {
        var repl = items.map(function (s) { return setParaText(tmpl, s); });
        els.splice(firstItem, lastItem - firstItem + 1, ...repl);
        log.push('project specific items: ' + items.length);
      }
    }

    // alternates and scope & cost table
    var ti = find(els, function (t, el) { return el.indexOf('<w:tbl') === 0 && t.indexOf('Owner Contract Base Bid') >= 0; });
    if (ti >= 0) els[ti] = rebuildScopeTable(els[ti], d.alternates || [], d.scopeLines || [], money, log);

    // scope items
    var ss = find(els, starts("Subcontractor's scope of work includes the following")), se = find(els, starts('Accepted Exclusions'), ss + 1);
    if (ss >= 0 && se > ss) {
      i = find(els, starts('This Subcontractor will provide a complete Submittal Package'), ss);
      if (i >= 0 && i < se && d.submittalDate) els[i] = setParaText(els[i], norm(textOf(els[i])).replace(/no later than [^.]*\d{4}\./, 'no later than ' + d.submittalDate + '.'));
      i = find(els, starts("Utilize the Owner's Sales and Use Tax Exemption"), ss);
      if (i >= 0 && i < se && !d.taxExempt) { els.splice(i, 1); se--; log.push('removed scope line: owner tax exemption'); }
      // scope name + CSI lines (ilvl 2)
      var csiIdx = -1; for (i = ss; i < se; i++) if (/<w:ilvl w:val="2"\/>/.test(els[i])) { csiIdx = i; break; }
      if (csiIdx >= 0) {
        var csi = (d.csi && d.csi.length) ? d.csi : [];
        var tplCsi = els[csiIdx];
        var newCsi = csi.map(function (c) { return setParaText(tplCsi, ' ' + c); });
        if (!newCsi.length) newCsi = [];
        els.splice(csiIdx, 1, ...newCsi); se += newCsi.length - 1;
        var furnish = csiIdx - 1;
        if (d.scopeName && /Furnish all necessary labor/i.test(textOf(els[furnish]))) {
          els[furnish] = setParaText(els[furnish], norm(textOf(els[furnish])).replace(/to complete all .*? scope per/i, 'to complete all ' + d.scopeName + ' scope per'));
        }
      }
      // inclusions: everything after "Coordination and acceptance…" up to the exclusions heading
      var co = find(els, starts('Coordination and acceptance of all documents'), ss);
      if (co >= 0 && co < se) {
        var incl = (d.inclusions || []).filter(function (s) { return String(s).trim(); });
        var endInc = se - 1; while (endInc > co && !/<w:numPr>/.test(els[endInc])) endInc--;
        var tplInc = els[co + 1] && /<w:numPr>/.test(els[co + 1]) ? els[co + 1] : els[co];
        els.splice(co + 1, endInc - co, ...incl.map(function (s) { return setParaText(tplInc, s); }));
        log.push('inclusions: ' + incl.length);
      }
    }

    // exclusions
    var xs = find(els, starts('Agreement does not include the following')), xe = find(els, starts('Terms of Agreement'), xs + 1);
    if (xs >= 0 && xe > xs) {
      if (d.bondRequired) { i = find(els, function (t) { return t === 'Bonds.'; }, xs); if (i >= 0 && i < xe) { els.splice(i, 1); xe--; log.push('removed exclusion: Bonds'); } }
      if (!d.taxExempt) { i = find(els, function (t) { return t === 'Sales taxes.'; }, xs); if (i >= 0 && i < xe) { els.splice(i, 1); xe--; log.push('removed exclusion: Sales taxes'); } }
      var dump = find(els, starts('Dumpsters or trash removal'), xs);
      var tplEx = els[dump];
      var after = dump + 1, endEx = xe - 1; while (endEx > dump && !/<w:numPr>/.test(els[endEx])) endEx--;
      var exc = (d.exclusions || []).filter(function (s) { return String(s).trim(); });
      if (dump >= 0) {
        // keep "Sales taxes." if still present right after Dumpsters
        if (after <= endEx && norm(textOf(els[after])) === 'Sales taxes.') after++;
        els.splice(after, Math.max(0, endEx - after + 1), ...exc.map(function (s) { return setParaText(tplEx, s); }));
        log.push('exclusions added: ' + exc.length);
      }
    }

    // terms of agreement: #60 billing, #61 retainage
    i = find(els, starts('Monthly Progress Pay Requests'));
    if (i >= 0) els[i] = setParaText(els[i], 'Monthly Progress Pay Requests should be submitted to Bailey-Harris by the ' + (d.submitDay || '10th') +
      ' of each month, with projection of work to be completed by the last day of the month, to allow time for review and any corrections. Pay requests, including any corrections, not received complete by the ' +
      (d.billingDate || '20th') + ' (the Monthly Billing Date) will be returned and processed the following month.');
    i = find(els, function (t) { return /Retainage will be withheld/.test(t); });
    if (i >= 0) els[i] = setParaText(els[i], retainageSentence(d.retPct, d.retDuration));

    D.els = els;
    zip.file('word/document.xml', stripHighlight(joinBody(D)));
    return { bytes: await zip.generateAsync({ type: 'uint8array' }), log: log };
  }

  function retainageSentence(pct, dur) {
    pct = Number(pct || 0);
    if (!pct) return 'No Retainage will be withheld from Monthly Pay Requests.';
    var p = (Math.round(pct * 100) / 100) + '%';
    if (/50/.test(dur || '')) return p + ' Retainage will be withheld from each Monthly Pay Request on completed work and stored materials until 50% of the Subcontract Price has been billed, after which no additional Retainage will be withheld.';
    return p + ' Retainage will be withheld from each Monthly Pay Request on completed work and stored materials throughout the Subcontract.';
  }

  function rebuildScopeTable(tbl, alts, lines, money, log) {
    var rows = tbl.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
    var head = tbl.slice(0, tbl.indexOf(rows[0])), tail = tbl.slice(tbl.lastIndexOf(rows[rows.length - 1]) + rows[rows.length - 1].length);
    var cellsOf = function (r) { return r.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []; };
    var setRow = function (r, vals) {
      var cs = cellsOf(r), out = r;
      cs.forEach(function (c, k) {
        var ps = c.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
        if (!ps.length) return;
        var nc = c.replace(ps[0], setParaText(ps[0], vals[k] == null ? '' : vals[k]));
        for (var z = 1; z < ps.length; z++) nc = nc.replace(ps[z], '');
        out = out.replace(c, nc);
      });
      return out;
    };
    var hdrIdx = rows.map(function (r, k) { return cellsOf(r).length === 1 ? k : -1; }).filter(function (k) { return k >= 0; });
    if (hdrIdx.length < 2) return tbl;
    var altHdr = rows[hdrIdx[0]], altTpl = rows[hdrIdx[0] + 1], costHdr = rows[hdrIdx[1]], costTpl = rows[hdrIdx[1] + 1], totTpl = rows[rows.length - 1];
    var out = [altHdr];
    alts.forEach(function (a) { out.push(setRow(altTpl, [a.label, a.status])); });
    out.push(costHdr);
    var total = 0;
    lines.forEach(function (l) { total += Number(l.amount || 0); out.push(setRow(costTpl, [l.desc, money(l.amount)])); });
    out.push(setRow(totTpl, ['Total', money(total)]));
    log.push('alternates rows: ' + alts.length + ', cost rows: ' + lines.length);
    return head + out.join('') + tail;
  }

  /* ---------- simple docx fills (C, D, F) ---------- */
  /* placeholders anywhere in a paragraph, plus "label cell → value in the next cell" for F */
  async function fillDocx(bytes, d) {
    var zip = await JSZipRef.loadAsync(bytes);
    var names = Object.keys(zip.files).filter(function (n) { return /^word\/(document|header\d*|footer\d*)\.xml$/.test(n); });
    for (var k = 0; k < names.length; k++) {
      var xml = await zip.file(names[k]).async('string');
      xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, function (p) {
        var t = textOf(p);
        if (!/<<|«/.test(t)) return p;
        var n = t.replace(/<<\s*Project Number\s*(&|and)\s*Name\s*>>/gi, d.projectNumName || '')
                 .replace(/<<\s*Project Name\s*>>/gi, d.projectNumName || '')
                 .replace(/<<\s*Project City, State\s*>>/gi, d.projectCity || '')
                 .replace(/<<\s*Subcontractor Name\s*>>/gi, d.subName || '');
        return n === t ? p : setParaText(p, n);
      });
      // table rows: label cell then empty value cell
      xml = xml.replace(/<w:tr[ >][\s\S]*?<\/w:tr>/g, function (r) {
        var cs = r.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || [];
        for (var c = 0; c < cs.length - 1; c++) {
          var lab = norm(textOf(cs[c])).replace(/:$/, '');
          var val = /^Project Number\/Name$/i.test(lab) ? d.projectNumName : /^Project Location$/i.test(lab) ? d.projectCity : null;
          if (val == null || norm(textOf(cs[c + 1]))) continue;
          var p = (cs[c + 1].match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/) || [null])[0];
          if (!p) continue;
          var np = p === '<w:p/>' ? '<w:p><w:r><w:t xml:space="preserve">' + esc(val) + '</w:t></w:r></w:p>' : setParaText(p, val);
          r = r.replace(cs[c + 1], cs[c + 1].replace(p, np));
        }
        return r;
      });
      zip.file(names[k], stripHighlight(xml));
    }
    return await zip.generateAsync({ type: 'uint8array' });
  }

  /* ---------- Attachment G (xlsx): shared-string placeholders + label → next cell ---------- */
  async function fillXlsx(bytes, d) {
    var zip = await JSZipRef.loadAsync(bytes);
    var ssx = zip.file('xl/sharedStrings.xml') ? await zip.file('xl/sharedStrings.xml').async('string') : '';
    var strings = [];
    ssx = ssx.replace(/<si>([\s\S]*?)<\/si>/g, function (all, inner) {
      var t = '', re = /<t(?:\s[^>]*)?>([^<]*)<\/t>/g, m; while ((m = re.exec(inner))) t += m[1];
      t = unesc(t); strings.push(t);
      if (!/<</.test(t)) return all;
      var n = t.replace(/<<\s*Subcontractor Name\s*>>/gi, d.subName || '').replace(/<<\s*Project Name\s*>>/gi, d.projectNumName || '');
      return '<si><t xml:space="preserve">' + esc(n) + '</t></si>';
    });
    if (ssx) zip.file('xl/sharedStrings.xml', ssx);
    var labels = { 'BH PROJECT #:': d.projectNumber, 'SUBCONTRACT #:': d.commitmentNumber };
    var sheets = Object.keys(zip.files).filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); });
    for (var s = 0; s < sheets.length; s++) {
      var x = await zip.file(sheets[s]).async('string');
      Object.keys(labels).forEach(function (lab) {
        var idx = strings.findIndex(function (t) { return norm(t) === lab; }); if (idx < 0 || !labels[lab]) return;
        var re = new RegExp('<c r="([A-Z]+)(\\d+)"([^>]*?)t="s"([^>]*)><v>' + idx + '</v></c>');
        var m = x.match(re); if (!m) return;
        var col = m[1], row = m[2];
        var nextCol = String.fromCharCode(col.charCodeAt(col.length - 1) + 1);
        // find the first empty cell to the right in the same row (up to 4 columns)
        for (var hop = 0; hop < 4; hop++) {
          var ref = (col.length > 1 ? col.slice(0, -1) : '') + nextCol + row;
          var cre = new RegExp('<c r="' + ref + '"([^>]*?)(/>|>([\\s\\S]*?)</c>)');
          var cm = x.match(cre);
          if (cm && !/<v>|<is>/.test(cm[3] || '')) {
            var attrs = cm[1].replace(/\s*t="[^"]*"/, '');
            x = x.replace(cm[0], '<c r="' + ref + '"' + attrs + ' t="inlineStr"><is><t>' + esc(labels[lab]) + '</t></is></c>');
            break;
          }
          nextCol = String.fromCharCode(nextCol.charCodeAt(0) + 1);
        }
      });
      zip.file(sheets[s], x);
    }
    return await zip.generateAsync({ type: 'uint8array' });
  }

  var api = { buildAttachmentB: buildAttachmentB, fillDocx: fillDocx, fillXlsx: fillXlsx, retainageSentence: retainageSentence, _textOf: textOf, _splitBody: splitBody };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.BHDocs = api;
})(this);
