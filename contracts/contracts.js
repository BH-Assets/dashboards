/* Contracts Desk (dashboard.baileyharris.com/contracts) — 19 Nov 2026.
   Prime Contract tab: read the owner contract, propose the terms, PM confirms, saved to Procore.
   Commitments tab: rows → Draft commitments in Procore with lines, custom fields and filled attachments.
   All Procore calls go through the bh-contracts Worker as the signed-in PM. */
(function () {
  'use strict';
  /* ---------------- config ---------------- */
  var API = 'https://contracts-api.baileyharris.com';
  var CID = '15326';
  var CLIENT_ID = '1825a1c0-1ce3-4078-aeb7-a4714562c422', TENANT_ID = '848aa890-7982-4028-befa-3323d82583ce';
  var SCOPES = ['User.Read', 'Sites.Selected'];
  var SP_HOST = 'baileyharris.sharepoint.com';
  var OPS_SITE = '/sites/Operations', OPS_LIB = 'Documents';
  var TEMPLATE_DIR = '03 Forms and Templates/Contract Documents/Commitments';
  var RIDER_DIR = '01 Working/Project Management/Approved Riders';
  var CF = { billing: 194473, taxExempt: 640523, ownerPlatform: 640528, subNotice: 640533, subItems: 640530, alternates: 640537,
    verifiedBy: 640531, verifiedDate: 640532, retDuration: 318332, ld: 269476, payTerms: 269489, finalPay: 322642, builderRisk: 320426,
    pdfType: 194471, bhSignatory: 300705, subSignatory: 302837, bond: 526946, riderNeeded: 542583 };
  var TYPES = {
    LF:  { label: 'Long Form', ep: 'work_order_contracts', series: 400, pdf: 'Standard Subcontract', att: ['B', 'C', 'D', 'F-Subcontract', 'G'], rider: 'Approved Subcontract Riders_Longform' },
    SF:  { label: 'Short Form', ep: 'work_order_contracts', series: 900, pdf: 'Short Form Subcontract', att: ['C', 'F-Subcontract'], rider: 'Approved Subcontract Riders_Shortform' },
    PSA: { label: 'PSA', ep: 'work_order_contracts', series: 100, pdf: 'Professional Services Agreement', att: ['F-PSA'], rider: 'Approved Professional Service Agreement Riders' },
    POON:{ label: 'PO on site', ep: 'purchase_order_contracts', series: 300, pdf: null, att: ['F-Purchase Order'], rider: 'Approved Purchase Order Riders' },
    POOFF:{ label: 'PO off site', ep: 'purchase_order_contracts', series: 300, pdf: null, att: ['F-Purchase Order'], rider: 'Approved Purchase Order Riders' }
  };
  var ATT_FILES = {
    'B': 'Attachment B - Scope of Work.docx', 'C': 'Attachment C - Safety Agreement.docx', 'D': 'Attachment D - OPS.docx',
    'F-Subcontract': 'Attachment F (NEW)/Attachment F - Subcontract.docx', 'F-PSA': 'Attachment F (NEW)/Attachment F - PSA.docx',
    'F-Purchase Order': 'Attachment F (NEW)/Attachment F - Purchase Order.docx'
  };
  var ATT_NAMES = { 'B': 'Attachment B - Scope of Work.docx', 'C': 'Attachment C - Safety Agreement.docx', 'D': 'Attachment D - OPS.docx',
    'F-Subcontract': 'Attachment F - Insurance Requirements.docx', 'F-PSA': 'Attachment F - Insurance Requirements.docx', 'F-Purchase Order': 'Attachment F - Insurance Requirements.docx' };

  /* ---------------- small helpers ---------------- */
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var money = function (v) { return '$' + Number(v || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
  var num = function (v) { var x = Number(String(v == null ? '' : v).replace(/[$,\s]/g, '')); return isNaN(x) ? 0 : Math.round(x * 100) / 100; };
  var htmlToText = function (h) { if (h == null) return ''; var d = document.createElement('div'); d.innerHTML = String(h).replace(/<\/(p|li|div)>/gi, '\n').replace(/<br\s*\/?>/gi, '\n'); return d.textContent.replace(/\n{2,}/g, '\n').trim(); };
  var linesOf = function (t) { return String(t || '').split(/\r?\n/).map(function (s) { return s.replace(/^\s*([-*•]|\d+[.)])\s*/, '').trim(); }).filter(Boolean); };
  var normName = function (s) { return String(s || '').toLowerCase().replace(/&/g, 'and').replace(/\b(inc|llc|co|corp|company|ltd|the)\b\.?/g, '').replace(/[^a-z0-9]/g, ''); };
  function msg(t, bad) { var m = $('#msg'); m.hidden = !t; m.className = 'cd-msg' + (bad ? ' is-bad' : ''); m.textContent = t || ''; }
  function stamp() { $('#asof').textContent = new Date().toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' }); }

  /* ---------------- Microsoft sign-in ---------------- */
  var app, acct;
  async function signIn() {
    app = new msal.PublicClientApplication({ auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/' + TENANT_ID, redirectUri: location.origin + '/contracts/', navigateToLoginRequestUrl: false }, cache: { cacheLocation: 'localStorage' } });
    await app.initialize(); await app.handleRedirectPromise();
    acct = app.getAllAccounts()[0];
    if (!acct) { await app.loginRedirect({ scopes: SCOPES }); return false; }
    app.setActiveAccount(acct);
    var n = (acct.name || acct.username || '').split(/\s+/); $('#avatar').textContent = ((n[0] || '')[0] || '') + ((n[1] || '')[0] || '');
    return true;
  }
  async function tokens() {
    try { return await app.acquireTokenSilent({ scopes: SCOPES, account: acct }); }
    catch (e) { await app.acquireTokenRedirect({ scopes: SCOPES, account: acct }); throw e; }
  }

  /* ---------------- Worker / Procore ---------------- */
  async function api(path, opt) {
    var t = await tokens(); opt = opt || {};
    var r = await fetch(API + path, { method: opt.method || 'GET', headers: Object.assign({ Authorization: 'Bearer ' + t.idToken }, opt.headers || {}), body: opt.body });
    var txt = await r.text(), j = null; try { j = txt ? JSON.parse(txt) : null; } catch (_) { j = { raw: txt }; }
    if (!r.ok) { var e = new Error((j && j.error) || ('HTTP ' + r.status)); e.status = r.status; e.body = j; throw e; }
    return j;
  }
  function pc(method, path, body, isForm) {
    var h = { 'X-Procore-Method': method, 'X-Procore-Path': path };
    if (body && !isForm) h['Content-Type'] = 'application/json';
    return api('/procore', { method: 'POST', headers: h, body: body ? (isForm ? body : JSON.stringify(body)) : undefined });
  }
  async function pcAll(path) { // simple pager for v1 lists
    var out = [], page = 1;
    while (page < 30) { var sep = path.indexOf('?') >= 0 ? '&' : '?'; var b = await pc('GET', path + sep + 'page=' + page + '&per_page=100'); out = out.concat(b || []); if (!b || b.length < 100) break; page++; }
    return out;
  }

  /* ---------------- Graph (templates and riders) ---------------- */
  var OPS = null;
  async function G(path, raw) {
    var t = await tokens();
    var r = await fetch(path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path, { headers: { Authorization: 'Bearer ' + t.accessToken } });
    if (!r.ok) throw new Error('SharePoint ' + r.status + ' on ' + path.split('?')[0]);
    return raw ? r : r.json();
  }
  async function ops() {
    if (OPS) return OPS;
    var s = await G('/sites/' + SP_HOST + ':' + OPS_SITE + '?$select=id');
    var d = await G('/sites/' + s.id + '/drives?$select=id,name');
    var drive = (d.value || []).find(function (x) { return x.name === OPS_LIB; }) || (d.value || [])[0];
    OPS = { site: s.id, drive: drive.id }; return OPS;
  }
  var encPath = function (p) { return p.split('/').map(encodeURIComponent).join('/'); };
  async function opsFile(path) { var o = await ops(); var r = await G('/drives/' + o.drive + '/root:/' + encPath(path) + ':/content', true); return new Uint8Array(await r.arrayBuffer()); }
  async function opsList(path) { var o = await ops(); var j = await G('/drives/' + o.drive + '/root:/' + encPath(path) + ':/children?$select=name,id,lastModifiedDateTime&$top=999'); return j.value || []; }

  /* ---------------- state ---------------- */
  var S = { me: null, projects: [], project: null, pcRec: null, lov: {}, defs: {}, proposals: {}, docs: [], rows: [], ref: null };

  /* ---------------- start ---------------- */
  async function start() {
    bindNav();
    try { if (!(await signIn())) return; } catch (e) { msg('Microsoft sign-in failed: ' + e.message, true); return; }
    await connection();
  }
  async function connection() {
    var c = $('#conn');
    if (/procore=connected/.test(location.hash)) history.replaceState(null, '', location.pathname + '#prime');
    if (/procore=denied/.test(location.hash)) msg('Procore access was not approved, so the desk cannot read or create anything yet.', true);
    var fm = location.hash.match(/procore=failed&why=([^&]*)/);
    if (fm) { msg('Connecting Procore failed: ' + decodeURIComponent(fm[1]), true); history.replaceState(null, '', location.pathname + '#prime'); }
    var st;
    try { st = await api('/procore/status'); }
    catch (e) { c.innerHTML = '<span><span class="dot"></span>The Contracts service is not reachable (' + esc(e.message) + ').</span>'; return; }
    if (!st.connected) {
      c.innerHTML = '<span><span class="dot"></span>Procore not connected</span><button class="bh-btn" id="pc-connect">Connect my Procore login</button>';
      $('#pc-connect').onclick = async function () { var r = await api('/procore/start', { method: 'POST' }); location.href = r.url; };
      return;
    }
    c.innerHTML = '<span><span class="dot is-on"></span>Procore connected as you</span>';
    await loadBase();
  }
  async function loadBase() {
    try {
      S.me = await pc('GET', '/rest/v1.0/me');
      S.projects = (await pc('GET', '/rest/v1.0/projects?company_id=' + CID + '&per_page=300')).filter(function (p) { return p.active !== false; })
        .sort(function (a, b) { return String(a.project_number || a.name).localeCompare(String(b.project_number || b.name)); });
      var defs = await pc('GET', '/rest/v1.0/custom_field_definitions?company_id=' + CID);
      (defs || []).forEach(function (d) { S.defs[d.id] = d; });
      var lovIds = [CF.retDuration, CF.taxExempt, CF.builderRisk, CF.pdfType, CF.bond, CF.riderNeeded];
      for (var i = 0; i < lovIds.length; i++) {
        try { S.lov[lovIds[i]] = await pc('GET', '/rest/v1.0/custom_field_definitions/' + lovIds[i] + '/custom_field_lov_entries?company_id=' + CID); } catch (_) { S.lov[lovIds[i]] = []; }
      }
    } catch (e) { msg('Could not load from Procore: ' + e.message, true); return; }
    var sel = $('#job');
    sel.innerHTML = '<option value="">Pick a job</option>' + S.projects.map(function (p) { return '<option value="' + p.id + '">' + esc((p.project_number ? p.project_number + ' · ' : '') + p.name) + '</option>'; }).join('');
    sel.disabled = false;
    var last = null; try { last = localStorage.getItem('bh.contracts.job'); } catch (_) {}
    if (last && S.projects.some(function (p) { return String(p.id) === last; })) { sel.value = last; await pickJob(); }
    sel.onchange = pickJob;
    stamp();
  }
  async function pickJob() {
    var id = $('#job').value; msg('');
    S.project = S.projects.find(function (p) { return String(p.id) === id; }) || null;
    try { localStorage.setItem('bh.contracts.job', id); } catch (_) {}
    S.proposals = {}; S.docs = []; S.rows = []; S.ref = null;
    if (!S.project) return;
    await loadPrime();
    renderRows();
    stamp();
  }

  /* =====================================================================
     PRIME CONTRACT TAB
     ===================================================================== */
  var FIELDS = [
    { key: 'retainage_percent', label: 'Default retainage %', std: 'retainage_percent', kind: 'num' },
    { key: 'retainage_duration', label: 'Retainage duration', cf: CF.retDuration, kind: 'lov' },
    { key: 'contract_date', label: 'Contract date', std: 'contract_date', kind: 'date' },
    { key: 'billing_date', label: 'Monthly Billing Date (trade partner pay apps)', cf: CF.billing, kind: 'text', dflt: '20th', hint: 'Last date a trade partner pay app (with corrections) is accepted for the month. Trade partners are asked to submit by the 10th. This is not in the owner contract; the owner date is under Payment terms.' },
    { key: 'tax_exempt', label: 'Tax exempt', cf: CF.taxExempt, kind: 'lov' },
    { key: 'liquidated_damages', label: 'Liquidated damages', cf: CF.ld, kind: 'text' },
    { key: 'payment_terms', label: 'Payment terms (owner)', cf: CF.payTerms, kind: 'rich' },
    { key: 'sub_notice', label: 'Sub notice deadlines', cf: CF.subNotice, kind: 'text', dflt: 'Same as PC.', hint: 'Leave "Same as PC." unless trade partners get fewer days than the owner gives Bailey-Harris. Long Form 8.3 already sets 48 hours for change and differing-condition proposals.' },
    { key: 'owner_platform', label: 'Owner specific platform required (unticked = Procore)', cf: CF.ownerPlatform, kind: 'bool' },
    { key: 'sub_project_items', label: 'Sub project items (owner and site rules)', cf: CF.subItems, kind: 'list', hint: 'One rule per line. Goes into Attachment B, Project Specific General Items.' },
    { key: 'accepted_alternates', label: 'Accepted alternates and VE', cf: CF.alternates, kind: 'alts', hint: 'One per line: description and amount | Accepted (or Not Accepted).' },
    { key: 'builders_risk_provider', label: "Builder's risk provider", cf: CF.builderRisk, kind: 'lov' },
    { key: 'final_payment', label: 'Final payment / closeout', cf: CF.finalPay, kind: 'rich' }
  ];
  function cfVal(rec, id) { var v = rec && rec.custom_fields && rec.custom_fields['custom_field_' + id]; return v && v.value !== undefined ? v.value : v; }
  function lovLabel(v) { if (v == null) return ''; if (Array.isArray(v)) return v.map(lovLabel).join(', '); if (typeof v === 'object') return v.label || v.name || ''; return String(v); }
  function currentText(f, rec) {
    var v = f.std ? rec[f.std] : cfVal(rec, f.cf);
    if (f.kind === 'lov') return lovLabel(v);
    if (f.kind === 'bool') return v === true ? 'Ticked' : v === false ? 'Not ticked' : '';
    if (f.kind === 'rich' || f.kind === 'list' || f.kind === 'alts') return htmlToText(v);
    if (f.kind === 'num') return v == null ? '' : String(Number(v));
    return v == null ? '' : String(v);
  }
  async function loadPrime() {
    $('#pc-table tbody').innerHTML = '<tr><td colspan="4">Loading the prime contract&hellip;</td></tr>';
    try { S.pcRec = await pc('GET', '/rest/v1.0/prime_contract?project_id=' + S.project.id); }
    catch (e) { S.pcRec = null; $('#pc-table tbody').innerHTML = '<tr><td colspan="4">No prime contract found in Procore for this job. Create it in Procore first (Part 1, step 1 of the PM guide).</td></tr>'; $('#pc-read').disabled = true; $('#pc-save').disabled = true; return; }
    S.docs = (S.pcRec.attachments || []).filter(function (a) { return /\.pdf$/i.test(a.name || a.filename || ''); }).map(function (a) { return { name: a.name || a.filename, url: a.url, source: 'procore' }; });
    renderDocs(); renderPrime();
    $('#pc-read').disabled = false; $('#pc-save').disabled = false;
    var by = lovLabel(cfVal(S.pcRec, CF.verifiedBy)) || (cfVal(S.pcRec, CF.verifiedBy) && cfVal(S.pcRec, CF.verifiedBy).name) || '';
    var dt = cfVal(S.pcRec, CF.verifiedDate);
    $('#pc-verified').textContent = by || dt ? 'Last verified by ' + (by || 'unknown') + (dt ? ' on ' + String(dt).slice(0, 10) : '') + '.' : 'Not verified yet. Commitments cannot be created until this is confirmed.';
  }
  function renderDocs() {
    var el = $('#pc-docs');
    if (!S.docs.length) { el.innerHTML = '<div class="doc">No PDFs attached to the prime contract in Procore. Attach the executed contract and its conditions there, or add the PDFs here.</div>'; return; }
    el.innerHTML = S.docs.map(function (d, i) { return '<div class="doc"><b>' + esc(d.name) + '</b><span class="bh-caption" id="doc-' + i + '">' + esc(d.status || (d.source === 'procore' ? 'In Procore' : 'Added here')) + '</span></div>'; }).join('');
  }
  function renderPrime() {
    var rec = S.pcRec;
    $('#pc-table tbody').innerHTML = FIELDS.map(function (f) {
      var cur = currentText(f, rec), p = S.proposals[f.key], val = p && p.value != null ? p.value : (cur || f.dflt || '');
      var input;
      if (f.kind === 'lov') {
        var opts = S.lov[f.cf] || [];
        var curLabel = typeof val === 'string' ? val : lovLabel(val);
        input = '<select data-k="' + f.key + '"><option value=""></option>' + opts.map(function (o) { return '<option value="' + o.id + '"' + (o.label === curLabel ? ' selected' : '') + '>' + esc(o.label) + '</option>'; }).join('') + '</select>';
      } else if (f.kind === 'bool') {
        var on = p ? !!p.value : cfVal(rec, f.cf) === true;
        input = '<label><input type="checkbox" data-k="' + f.key + '"' + (on ? ' checked' : '') + '> Owner requires its own system</label>';
      } else if (f.kind === 'rich' || f.kind === 'list' || f.kind === 'alts') {
        if (Array.isArray(val)) val = val.map(function (x) { return typeof x === 'object' ? (x.label + ' | ' + (x.status || 'Accepted')) : x; }).join('\n');
        input = '<textarea data-k="' + f.key + '" rows="' + (f.kind === 'rich' ? 4 : 6) + '">' + esc(val) + '</textarea>';
      } else {
        input = '<input type="text" data-k="' + f.key + '" value="' + esc(val) + '">';
      }
      var src = p && p.source ? '<div class="src">' + esc(p.source) + (p.quote ? '<q>' + esc(p.quote) + '</q>' : '') + '</div>' : (f.dflt && !cur ? '<div class="src">Company default</div>' : '');
      var diff = p && p.value != null && normCmp(p.value) !== normCmp(cur) && cur;
      return '<tr' + (diff ? ' class="is-diff" title="The contract reading differs from what Procore holds"' : '') + '><td><b>' + esc(f.label) + '</b>' + (f.hint ? '<div class="hint">' + esc(f.hint) + '</div>' : '') + '</td><td class="cur"><div>' + esc(cur || '—') + '</div></td><td>' + input + '</td><td>' + src + '</td></tr>';
    }).join('');
  }
  var normCmp = function (v) { if (Array.isArray(v)) v = v.map(function (x) { return typeof x === 'object' ? x.label : x; }).join(' '); return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9%]/g, ''); };

  /* ---- reading the contract (in this browser) ---- */
  var KEYWORDS = /retain|retention|application for payment|pay(ment)? application|progress payment|final payment|liquidated|notice|calendar days|days after|weather|delay|change order|claim|sales (and use )?tax|tax[- ]exempt|exemption|builder'?s risk|alternate|value engineering|football|graduation|parking|signage|drone|unmanned|smok|firearm|harass|radio|safety|confidential|non-?disclosure|platform|e-builder|procore|textura|contract sum|contract time|substantial completion/i;
  var BANK = /(routing|aba)\s*(no\.?|number|#)?\s*[:#]?\s*\d{9}|account\s*(no\.?|number|#)\s*[:#]?\s*\d{4,}/i;
  async function pdfPages(bytes) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = '/contracts/pdf.worker.min.js';
    var doc = await pdfjsLib.getDocument({ data: bytes }).promise, out = [];
    for (var i = 1; i <= doc.numPages; i++) {
      var pg = await doc.getPage(i), tc = await pg.getTextContent();
      out.push(tc.items.map(function (it) { return it.str + (it.hasEOL ? '\n' : ' '); }).join('').replace(/[ \t]+/g, ' '));
    }
    return out;
  }
  function confidentiality(allText) {
    var hits = [], re = /[^.]{0,300}(confidential|non-?disclosure)[^.]{0,300}\./gi, m;
    while ((m = re.exec(allText))) hits.push(m[0].replace(/\s+/g, ' ').trim());
    if (!hits.length) return { found: false };
    var bars = hits.filter(function (h) { return /(shall|will|may) not (disclose|divulge|share|release)|not be disclosed|prohibit/i.test(h) && /third part|any person|any party|others|artificial intelligence|\bAI\b|machine learning/i.test(h); });
    var permits = hits.some(function (h) { return /consultant|advisor|agent|service provider|subcontractor|representative|professional|need to know/i.test(h); });
    var ai = hits.some(function (h) { return /artificial intelligence|\bAI\b|machine learning|large language/i.test(h); });
    return { found: true, blocked: ai || (bars.length > 0 && !permits), quote: (bars[0] || hits[0]).slice(0, 400) };
  }
  async function readContract() {
    var btn = $('#pc-read'); btn.disabled = true; msg('Reading the contract PDFs in this browser…');
    try {
      var sections = [], all = '', scanned = [];
      for (var i = 0; i < S.docs.length; i++) {
        var d = S.docs[i], st = $('#doc-' + i);
        try {
          if (!d.bytes) { var r = await fetch(d.url); if (!r.ok) throw new Error('HTTP ' + r.status); d.bytes = new Uint8Array(await r.arrayBuffer()); }
        } catch (e) { d.status = 'Could not download from Procore in this browser. Add this PDF with "Add contract PDFs".'; if (st) st.textContent = d.status; continue; }
        var pages = await pdfPages(d.bytes.slice(0));
        var textPages = pages.filter(function (t) { return t.trim().length > 80; }).length;
        if (textPages < pages.length / 2) { scanned.push(d.name); d.status = 'Scanned (' + pages.length + ' pages, no text). Server text recognition is not set up yet; fill those terms by hand.'; if (st) st.textContent = d.status; continue; }
        var picked = 0;
        pages.forEach(function (t, k) {
          all += '\n' + t;
          if (BANK.test(t)) return;          // never send a page with banking details
          if (KEYWORDS.test(t)) { sections.push({ doc: d.name, page: k + 1, text: t }); picked++; }
        });
        d.status = pages.length + ' pages, ' + picked + ' with contract terms'; if (st) st.textContent = d.status;
      }
      var conf = confidentiality(all);
      if (conf.found && conf.blocked) { msg('This contract has a confidentiality clause that does not allow sharing with service providers, so nothing was sent for reading. Fill the fields from the contract by hand. Clause: "' + conf.quote + '"', true); return; }
      if (!sections.length) { msg(scanned.length ? 'Only scanned PDFs were found (' + scanned.join(', ') + '). Fill the fields by hand from the contract.' : 'No contract terms found in the PDFs.', true); return; }
      msg('Sending ' + sections.length + ' pages of contract terms for reading…' + (conf.found ? ' (A confidentiality clause was found; it allows sharing with service providers.)' : ''));
      var res = await api('/read', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ project: (S.project.project_number || '') + ' ' + S.project.name, sections: sections }) });
      var f = res.fields || {};
      S.proposals = {};
      var map = { retainage_percent: 'retainage_percent', retainage_duration: 'retainage_duration', contract_date: 'contract_date', tax_exempt: 'tax_exempt', liquidated_damages: 'liquidated_damages',
        payment_terms: 'payment_terms', sub_project_items: 'sub_project_items', accepted_alternates: 'accepted_alternates', builders_risk_provider: 'builders_risk_provider', final_payment: 'final_payment' };
      Object.keys(map).forEach(function (k) { if (f[k] && f[k].value != null) S.proposals[map[k]] = f[k]; });
      if (f.owner_platform && f.owner_platform.value) S.proposals.owner_platform = { value: true, source: f.owner_platform.source, quote: 'Owner system: ' + f.owner_platform.value };
      if (f.owner_notice_deadlines && f.owner_notice_deadlines.value) S.proposals.sub_notice = { value: null, source: 'Owner deadlines: ' + (Array.isArray(f.owner_notice_deadlines.value) ? f.owner_notice_deadlines.value.join('; ') : f.owner_notice_deadlines.value), quote: '' };
      renderPrime();
      msg('Proposed values are filled in. Check each one against the contract (source shown on the right), correct anything wrong, then confirm. Rows in yellow differ from what Procore holds.' + ((res.notes || []).length ? ' Notes: ' + res.notes.join(' ') : ''));
    } catch (e) { msg('Reading failed: ' + e.message, true); }
    finally { btn.disabled = false; }
  }

  /* ---- saving to Procore (v2; v1 ignores custom fields on prime contracts) ---- */
  function toRich(kind, text) {
    var L = linesOf(text);
    if (!L.length) return null;
    if (kind === 'list') return '<ul>' + L.map(function (s) { return '<li>' + esc(s) + '</li>'; }).join('') + '</ul>';
    return L.map(function (s) { return '<p>' + esc(s) + '</p>'; }).join('');
  }
  async function savePrime() {
    var btn = $('#pc-save'); btn.disabled = true; msg('Saving to Procore…');
    try {
      var rec = S.pcRec, body = {}, changed = [];
      FIELDS.forEach(function (f) {
        var el = $('[data-k="' + f.key + '"]'); if (!el) return;
        var cur = currentText(f, rec), v;
        if (f.kind === 'bool') { v = el.checked; if (v === (cfVal(rec, f.cf) === true)) return; }
        else if (f.kind === 'lov') { if (!el.value) return; var o = (S.lov[f.cf] || []).find(function (x) { return String(x.id) === el.value; }); if (o && o.label === cur) return; v = Number(el.value); }
        else if (f.kind === 'rich' || f.kind === 'list' || f.kind === 'alts') { if (normCmp(el.value) === normCmp(cur)) return; v = toRich(f.kind, el.value); }
        else if (f.kind === 'num') { if (el.value === '' || Number(el.value) === Number(cur)) return; v = Number(el.value); }
        else { if (el.value.trim() === cur.trim()) return; v = el.value.trim(); }
        body[f.std ? f.std : 'custom_field_' + f.cf] = v; changed.push(f.label);
      });
      var ok = confirm('Save ' + (changed.length ? changed.length + ' changed field(s) (' + changed.join(', ') + ')' : 'no changed fields') + ' and stamp the prime contract as verified by you today?');
      if (!ok) { msg(''); return; }
      body['custom_field_' + CF.verifiedBy] = S.me.id;
      var t = new Date(); body['custom_field_' + CF.verifiedDate] = t.toISOString().slice(0, 10) + 'T18:00:00Z';
      await pc('PATCH', '/rest/v2.0/companies/' + CID + '/projects/' + S.project.id + '/prime_contracts/' + rec.id, body);
      await loadPrime();
      // read back and check every field we sent
      var missed = [];
      Object.keys(body).forEach(function (k) {
        if (k.indexOf('custom_field_') !== 0) return;
        var v = cfVal(S.pcRec, k.replace('custom_field_', ''));
        if (v == null || v === '' || (Array.isArray(v) && !v.length)) missed.push(k);
      });
      msg(missed.length ? 'Saved, but these fields did not come back from Procore: ' + missed.join(', ') + '. Check them in Procore.' : 'Saved to Procore and read back: ' + (changed.length ? changed.join(', ') + '; ' : '') + 'verified by you today.', !!missed.length);
    } catch (e) { msg('Save failed: ' + e.message + (e.body && e.body.error && e.body.error.details ? ' — ' + JSON.stringify(e.body.error.details) : ''), true); }
    finally { btn.disabled = false; }
  }

  /* =====================================================================
     COMMITMENTS TAB
     ===================================================================== */
  function pcTerms() {
    var r = S.pcRec; if (!r) return null;
    var dur = lovLabel(cfVal(r, CF.retDuration));
    return {
      verified: !!cfVal(r, CF.verifiedBy) && !!cfVal(r, CF.verifiedDate),
      retPct: Number(r.retainage_percent || 0), retDuration: dur || 'N/A',
      billing: cfVal(r, CF.billing) || '20th', taxExempt: /^yes$/i.test(lovLabel(cfVal(r, CF.taxExempt))),
      ownerPlatform: cfVal(r, CF.ownerPlatform) === true, builderRiskOwner: /owner/i.test(lovLabel(cfVal(r, CF.builderRisk))),
      items: linesOf(htmlToText(cfVal(r, CF.subItems))),
      alternates: linesOf(htmlToText(cfVal(r, CF.alternates))).map(function (l) { var p = l.split('|'); return { label: p[0].trim(), status: (p[1] || 'Accepted').trim() }; })
    };
  }
  async function loadRef() {
    if (S.ref && S.ref.pid === S.project.id) return S.ref;
    var pid = S.project.id;
    var R = { pid: pid };
    var get = async function (p) { try { return await pcAll(p); } catch (e) { return { error: e.message }; } };
    R.vendors = await get('/rest/v1.0/projects/' + pid + '/vendors');
    R.users = await get('/rest/v1.0/projects/' + pid + '/users');
    R.costCodes = await get('/rest/v1.0/cost_codes?project_id=' + pid);
    R.types = await get('/rest/v1.0/line_item_types?company_id=' + CID);
    var woc = await get('/rest/v1.0/work_order_contracts?project_id=' + pid), poc = await get('/rest/v1.0/purchase_order_contracts?project_id=' + pid);
    R.numbers = [].concat(Array.isArray(woc) ? woc : [], Array.isArray(poc) ? poc : []).map(function (c) { return String(c.number || ''); });
    ['vendors', 'users', 'costCodes', 'types'].forEach(function (k) { if (!Array.isArray(R[k])) { msg('Could not load ' + k + ' from Procore: ' + R[k].error, true); R[k] = []; } });
    R.vendors.sort(function (a, b) { return String(a.name).localeCompare(String(b.name)); });
    S.ref = R; return R;
  }
  var typeCode = function (t) { return String((t && (t.code || t.name)) || '').trim().charAt(0).toUpperCase(); };
  function nextNumber(typeKey, taken) {
    var pn = String(S.project.project_number || '').trim(), base = TYPES[typeKey].series;
    var newStyle = /^\d{4}-\d{2}$/.test(pn) || S.ref.numbers.some(function (n) { return pn && n.indexOf(pn) === 0; });
    var used = S.ref.numbers.concat(taken || []);
    for (var k = base + 1; k < base + 100; k++) {
      var n = newStyle ? pn + String(k) : String(k) + pn;
      if (used.indexOf(n) < 0) return n;
    }
    return '';
  }
  function newRow(seed) {
    return Object.assign({ id: Math.random().toString(36).slice(2), type: 'LF', vendorId: '', title: '', signatoryId: '', bhSignatoryId: S.me ? S.me.id : '', bond: 'No', scopeName: '', csi: '', submittalDate: '', inclusions: '', exclusions: '', description: '', lines: [{ costCodeId: '', typeId: '', desc: '', amount: '' }], rider: null, riderPick: '' }, seed || {});
  }
  async function renderRows() {
    var host = $('#cm-rows'), T = pcTerms();
    $('#cm-add').disabled = $('#cm-paste').disabled = $('#cm-create').disabled = true;
    if (!S.project) { host.innerHTML = ''; return; }
    if (!T) { $('#cm-pc').innerHTML = 'No prime contract in Procore for this job.'; host.innerHTML = ''; return; }
    $('#cm-pc').innerHTML = '<span>Retainage <b>' + T.retPct + '% ' + esc(T.retDuration) + '</b></span><span>Monthly Billing Date <b>' + esc(T.billing) + '</b> (submit by the 10th)</span><span>Tax exempt <b>' + (T.taxExempt ? 'Yes' : 'No') + '</b></span><span>Platform <b>' + (T.ownerPlatform ? 'Owner system' : 'Procore') + '</b></span><span>Owner rules <b>' + T.items.length + '</b></span><span>Alternates <b>' + T.alternates.length + '</b></span>';
    if (!T.verified) { host.innerHTML = '<div class="cd-msg is-bad">The prime contract is not verified yet. Confirm it on the Prime Contract tab first.</div>'; return; }
    host.innerHTML = '<p class="bh-caption">Loading vendors, cost codes and existing commitments&hellip;</p>';
    await loadRef();
    if (!S.rows.length) S.rows.push(newRow());
    $('#cm-add').disabled = $('#cm-paste').disabled = $('#cm-create').disabled = false;
    drawRows();
  }
  function drawRows() {
    var R = S.ref, taken = [];
    S.rows.forEach(function (row) { row.number = nextNumber(row.type, taken); taken.push(row.number); });
    $('#c-rows').textContent = S.rows.length || '';
    var vOpts = function (sel) { return '<option value=""></option>' + R.vendors.map(function (v) { return '<option value="' + v.id + '"' + (String(v.id) === String(sel) ? ' selected' : '') + '>' + esc(v.name) + '</option>'; }).join(''); };
    var uOpts = function (sel, vendorId) {
      var list = R.users.filter(function (u) { return !vendorId || (u.vendor && String(u.vendor.id) === String(vendorId)); });
      return '<option value=""></option>' + list.map(function (u) { return '<option value="' + u.id + '"' + (String(u.id) === String(sel) ? ' selected' : '') + '>' + esc(u.name || ((u.first_name || '') + ' ' + (u.last_name || ''))) + '</option>'; }).join('');
    };
    var ccOpts = function (sel) { return '<option value=""></option>' + R.costCodes.map(function (c) { return '<option value="' + c.id + '"' + (String(c.id) === String(sel) ? ' selected' : '') + '>' + esc((c.full_code || c.code || '') + ' ' + (c.name || '')) + '</option>'; }).join(''); };
    var tOpts = function (sel) { return '<option value=""></option>' + R.types.map(function (t) { return '<option value="' + t.id + '"' + (String(t.id) === String(sel) ? ' selected' : '') + '>' + esc((t.code ? t.code + ' ' : '') + t.name) + '</option>'; }).join(''); };
    $('#cm-rows').innerHTML = S.rows.map(function (row, i) {
      var T = TYPES[row.type], iss = issues(row), lf = row.type === 'LF';
      var total = row.lines.reduce(function (s, l) { return s + num(l.amount); }, 0);
      var state = iss.some(function (x) { return x.bad; }) ? '<span class="bh-chip is-bad">Blocked</span>' : iss.length ? '<span class="bh-chip is-warn">Needs review</span>' : '<span class="bh-chip is-ok">Ready</span>';
      if (row.done) state = '<span class="bh-chip">Created ' + esc(row.done) + '</span>';
      return '<div class="cd-row" data-i="' + i + '"><div class="cd-row-h"><span class="num">' + esc(row.number) + '</span><span class="ttl">' + esc(row.title || 'New commitment') + ' · ' + money(total) + '</span>' + state +
        '<button class="bh-btn-outline" data-act="del">Remove</button></div><div class="cd-row-b">' +
        '<label>Type<select data-f="type">' + Object.keys(TYPES).map(function (k) { return '<option value="' + k + '"' + (k === row.type ? ' selected' : '') + '>' + TYPES[k].label + '</option>'; }).join('') + '</select></label>' +
        '<label class="wide">Vendor<select data-f="vendorId">' + vOpts(row.vendorId) + '</select></label>' +
        '<label>P&amp;P bond required<select data-f="bond"><option' + (row.bond === 'No' ? ' selected' : '') + '>No</option><option' + (row.bond === 'Yes' ? ' selected' : '') + '>Yes</option></select></label>' +
        '<label class="wide">Title<input data-f="title" value="' + esc(row.title) + '"></label>' +
        '<label>Commitment signatory (vendor)<select data-f="signatoryId">' + uOpts(row.signatoryId, row.vendorId) + '</select></label>' +
        '<label>BH signatory (project PM)<select data-f="bhSignatoryId">' + uOpts(row.bhSignatoryId) + '</select></label>' +
        (lf ? '<label class="wide">Scope name (Attachment B)<input data-f="scopeName" value="' + esc(row.scopeName) + '" placeholder="e.g. Drywall and Framing"></label>' +
              '<label class="wide">Submittal deadline<input data-f="submittalDate" value="' + esc(row.submittalDate) + '" placeholder="e.g. Friday, December 18, 2026"></label>' +
              '<label class="full">Specification sections (one per line)<textarea data-f="csi" rows="2">' + esc(row.csi) + '</textarea></label>'
            : '<label class="full">Description<textarea data-f="description" rows="2">' + esc(row.description) + '</textarea></label>') +
        '<label class="wide">Inclusions (one per line)<textarea data-f="inclusions">' + esc(row.inclusions) + '</textarea></label>' +
        '<label class="wide">Exclusions accepted (one per line)<textarea data-f="exclusions">' + esc(row.exclusions) + '</textarea></label>' +
        '<div class="cd-lines"><table><thead><tr><th style="width:30%">Cost code</th><th style="width:16%">Cost type</th><th>Description</th><th style="width:14%">Amount</th><th style="width:40px"></th></tr></thead><tbody>' +
        row.lines.map(function (l, j) { return '<tr data-j="' + j + '"><td><select data-l="costCodeId">' + ccOpts(l.costCodeId) + '</select></td><td><select data-l="typeId">' + tOpts(l.typeId) + '</select></td><td><input data-l="desc" value="' + esc(l.desc) + '"></td><td><input class="amt" data-l="amount" value="' + esc(l.amount) + '"></td><td><button class="bh-btn-outline" data-act="dl" aria-label="Remove line">&times;</button></td></tr>'; }).join('') +
        '</tbody></table><button class="bh-btn-outline" data-act="al" style="margin-top:6px">Add line</button></div>' +
        '<label class="wide">Rider<select data-f="riderPick">' + riderOpts(row) + '</select></label>' +
        '<div class="wide bh-caption" style="align-self:end">Attachments: ' + T.att.map(function (a) { return a === 'B' ? 'B (built)' : a; }).join(', ') + (row.type === 'SF' ? ' (Short Form F version not in the templates yet; the Subcontract F is used)' : '') + '</div>' +
        (iss.length ? '<ul class="cd-issues">' + iss.map(function (x) { return '<li class="' + (x.bad ? 'bad' : 'warn') + '">' + esc(x.t) + '</li>'; }).join('') + '</ul>' : '') +
        '</div></div>';
    }).join('');
  }
  function riderOpts(row) {
    if (!row.vendorId) return '<option value="">Pick a vendor first</option>';
    if (!row.rider) return '<option value="">Looking for an approved rider…</option>';
    if (row.rider.error) return '<option value="">Rider folder not reachable: ' + esc(row.rider.error) + '</option>';
    if (!row.rider.matches.length) return '<option value="">No approved rider for this vendor</option>';
    return (row.rider.matches.length > 1 ? '<option value="">Pick one of ' + row.rider.matches.length + ' riders</option>' : '') +
      row.rider.matches.map(function (m) { return '<option value="' + esc(m.name) + '"' + (row.riderPick === m.name ? ' selected' : '') + '>' + esc(m.name) + '</option>'; }).join('') + '<option value="none"' + (row.riderPick === 'none' ? ' selected' : '') + '>No rider</option>';
  }
  var RIDER_CACHE = {};
  async function findRider(row) {
    var v = S.ref.vendors.find(function (x) { return String(x.id) === String(row.vendorId); }); if (!v) return;
    var folder = TYPES[row.type].rider;
    try {
      if (!RIDER_CACHE[folder]) RIDER_CACHE[folder] = await opsList(RIDER_DIR + '/' + folder);
      var key = normName(v.name);
      var matches = RIDER_CACHE[folder].filter(function (f) { return /\.(docx|pdf)$/i.test(f.name) && key && normName(f.name.split(/ - |_/)[0]).indexOf(key.slice(0, 12)) === 0; })
        .sort(function (a, b) { return String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)); });
      row.rider = { folder: folder, matches: matches };
      if (matches.length === 1) row.riderPick = matches[0].name;
    } catch (e) { row.rider = { error: e.message, matches: [] }; }
    drawRows();
  }
  var DUMPSTER = /dumpster|roll[- ]?off|toilet|porta|handwash|hand wash|trailer/i;
  function issues(row) {
    var T = pcTerms(), out = [], R = S.ref, isPO = /^PO/.test(row.type);
    var bad = function (t) { out.push({ bad: true, t: t }); }, warn = function (t) { out.push({ t: t }); };
    if (!row.vendorId) bad('Pick the vendor.');
    if (!row.title.trim()) bad('Add a title.');
    if (!row.signatoryId) bad('Pick the commitment signatory (the vendor\'s signer); Procore requires it.');
    if (!row.bhSignatoryId) bad('Pick the BH signatory (the project PM); Procore requires it.');
    var total = 0;
    row.lines.forEach(function (l, j) {
      var n = j + 1, t = R.types.find(function (x) { return String(x.id) === String(l.typeId); }), c = typeCode(t);
      total += num(l.amount);
      if (!l.costCodeId) bad('Line ' + n + ': pick a cost code.');
      if (!l.typeId) bad('Line ' + n + ': pick a cost type.');
      if (!(num(l.amount) > 0)) bad('Line ' + n + ': enter an amount.');
      if (!c) return;
      if (c === 'S' && !/^(LF|SF)$/.test(row.type)) bad('Line ' + n + ': cost type S is only for Long Form and Short Form subcontracts.');
      if (row.type === 'PSA' && c !== 'O') bad('Line ' + n + ': PSA lines must be cost type O.');
      if (isPO && 'OMX'.indexOf(c) < 0) bad('Line ' + n + ': PO lines must be O, M or X.');
      if (c === 'X' && !T.taxExempt) bad('Line ' + n + ': X is only for tax-exempt jobs.');
      if (c === 'X' && /temp/i.test(l.desc)) bad('Line ' + n + ': X is never used for temporary materials.');
      if (isPO && DUMPSTER.test(l.desc + ' ' + row.title) && c !== 'O') bad('Line ' + n + ': dumpsters, temporary toilets and handwash stations, and jobsite trailers on a PO are always O.');
    });
    if (row.type === 'SF' && total > 25000) bad('Short Form is limited to $25,000; use a Long Form.');
    if (row.type === 'LF') {
      if (!row.scopeName.trim()) bad('Add the scope name for Attachment B.');
      if (!linesOf(row.inclusions).length) warn('No inclusions entered; Attachment B will have an empty scope list.');
      if (!row.submittalDate.trim()) bad('Add the submittal deadline for Attachment B.');
    }
    if (row.rider && row.rider.matches && row.rider.matches.length > 1 && !row.riderPick) bad('Several approved riders match this vendor; pick one.');
    if (!row.number) bad('No free commitment number left in the ' + TYPES[row.type].series + ' series.');
    return out;
  }
  function bindRows() {
    var host = $('#cm-rows');
    host.addEventListener('change', function (e) {
      var rEl = e.target.closest('.cd-row'); if (!rEl) return;
      var row = S.rows[+rEl.dataset.i], f = e.target.dataset.f, l = e.target.dataset.l;
      if (f) { row[f] = e.target.value; if (f === 'vendorId' || f === 'type') { row.rider = null; row.riderPick = ''; if (row.vendorId) findRider(row); } }
      if (l) { var j = +e.target.closest('tr').dataset.j; row.lines[j][l] = e.target.value; }
      drawRows();
    });
    host.addEventListener('click', function (e) {
      var b = e.target.closest('[data-act]'); if (!b) return;
      var rEl = b.closest('.cd-row'), row = S.rows[+rEl.dataset.i];
      if (b.dataset.act === 'del') S.rows.splice(+rEl.dataset.i, 1);
      if (b.dataset.act === 'al') row.lines.push({ costCodeId: '', typeId: '', desc: '', amount: '' });
      if (b.dataset.act === 'dl') { row.lines.splice(+b.closest('tr').dataset.j, 1); if (!row.lines.length) row.lines.push({ costCodeId: '', typeId: '', desc: '', amount: '' }); }
      drawRows();
    });
    $('#cm-add').onclick = function () { S.rows.push(newRow()); drawRows(); };
    $('#cm-paste').onclick = function () { $('#paste-text').value = ''; $('#dlg-paste').showModal(); };
    $('#dlg-paste').addEventListener('close', function () { if ($('#dlg-paste').returnValue === 'ok') pasteRows($('#paste-text').value); });
    $('#cm-create').onclick = createReady;
  }
  function pasteRows(text) {
    var R = S.ref, typeOf = function (s) { s = String(s).toLowerCase(); return /long/.test(s) ? 'LF' : /short/.test(s) ? 'SF' : /psa|professional/.test(s) ? 'PSA' : /off/.test(s) ? 'POOFF' : /po|purchase/.test(s) ? 'POON' : 'LF'; };
    var groups = {}, unknown = [];
    text.split(/\r?\n/).forEach(function (ln) {
      var c = ln.split('\t'); if (c.length < 4 || !c[0].trim()) return;
      var v = R.vendors.find(function (x) { return normName(x.name) === normName(c[0]); }) || R.vendors.find(function (x) { return normName(x.name).indexOf(normName(c[0])) === 0; });
      if (!v) unknown.push(c[0].trim());
      var key = (v ? v.id : c[0]) + '|' + (c[2] || '').trim();
      var cc = R.costCodes.find(function (x) { return String(x.full_code || x.code || '').replace(/\s/g, '') === String(c[4] || '').replace(/\s/g, ''); });
      var tt = R.types.find(function (x) { return typeCode(x) === String(c[5] || '').trim().toUpperCase().charAt(0); });
      if (!groups[key]) groups[key] = newRow({ type: typeOf(c[1]), vendorId: v ? String(v.id) : '', title: (c[2] || '').trim(), lines: [] });
      groups[key].lines.push({ costCodeId: cc ? String(cc.id) : '', typeId: tt ? String(tt.id) : '', desc: (c[2] || '').trim(), amount: num(c[3]) });
    });
    var add = Object.keys(groups).map(function (k) { return groups[k]; });
    if (S.rows.length === 1 && !S.rows[0].vendorId && !S.rows[0].title) S.rows = [];
    S.rows = S.rows.concat(add);
    add.forEach(function (r) { if (r.vendorId) findRider(r); });
    drawRows();
    if (unknown.length) msg('Not found in the project directory, pick these vendors on their rows: ' + unknown.join(', '), true);
  }

  /* ---- create in Procore ---- */
  function log(t) { var l = $('#cm-log'); l.hidden = false; l.textContent += t + '\n'; }
  async function createReady() {
    var ready = S.rows.filter(function (r) { return !r.done && !issues(r).some(function (x) { return x.bad; }); });
    if (!ready.length) { msg('No rows are ready. Fix the items marked in red first.', true); return; }
    if (!confirm('Create ' + ready.length + ' Draft commitment(s) in Procore on ' + (S.project.project_number || S.project.name) + '?')) return;
    $('#cm-create').disabled = true; $('#cm-log').textContent = '';
    for (var i = 0; i < ready.length; i++) {
      try { await createOne(ready[i]); }
      catch (e) { log(ready[i].number + ': FAILED — ' + e.message + (e.body ? ' ' + JSON.stringify(e.body).slice(0, 300) : '')); }
    }
    S.ref = null; await loadRef(); drawRows();
    $('#cm-create').disabled = false;
  }
  function lovId(cf, label) { var o = (S.lov[cf] || []).find(function (x) { return String(x.label).toLowerCase() === String(label).toLowerCase(); }); return o ? o.id : null; }
  async function createOne(row) {
    var T = pcTerms(), TY = TYPES[row.type], pid = S.project.id, ep = TY.ep;
    var vendor = S.ref.vendors.find(function (v) { return String(v.id) === String(row.vendorId); });
    var ret = row.type === 'PSA' ? 0 : T.retPct;
    var cfs = {};
    if (TY.pdf && lovId(CF.pdfType, TY.pdf)) cfs['custom_field_' + CF.pdfType] = lovId(CF.pdfType, TY.pdf);
    cfs['custom_field_' + CF.billing] = T.billing;
    if (row.bhSignatoryId) cfs['custom_field_' + CF.bhSignatory] = Number(row.bhSignatoryId);
    if (row.signatoryId) cfs['custom_field_' + CF.subSignatory] = Number(row.signatoryId);
    if (lovId(CF.bond, row.bond)) cfs['custom_field_' + CF.bond] = lovId(CF.bond, row.bond);
    var durLabel = row.type === 'PSA' ? 'N/A' : T.retDuration;
    if (lovId(CF.retDuration, durLabel)) cfs['custom_field_' + CF.retDuration] = lovId(CF.retDuration, durLabel);
    var hasRider = row.riderPick && row.riderPick !== 'none';
    if (lovId(CF.riderNeeded, hasRider ? 'Yes, attached' : 'No')) cfs['custom_field_' + CF.riderNeeded] = lovId(CF.riderNeeded, hasRider ? 'Yes, attached' : 'No');
    var c = { number: row.number, title: row.title, vendor_id: Number(row.vendorId), status: 'Draft', executed: false, retainage_percent: ret };
    if (row.type !== 'LF') { c.inclusions = toRich('rich', row.inclusions) || ''; c.exclusions = toRich('rich', row.exclusions) || ''; c.description = toRich('rich', row.description) || ''; }
    var key = ep === 'work_order_contracts' ? 'work_order_contract' : 'purchase_order_contract';
    var body = {}; body[key] = c;
    log(row.number + ': creating ' + TY.label + ' for ' + (vendor ? vendor.name : row.vendorId) + '…');
    var made = await pc('POST', '/rest/v1.0/' + ep + '?project_id=' + pid, body);
    log(row.number + ': created as Draft (id ' + made.id + ').');
    // custom fields: v1 ignores them, v2 saves them but wants the contract date and every required field in the same call
    var v2 = Object.assign({ contract_date: new Date().toISOString().slice(0, 10) }, cfs);
    await pc('PATCH', '/rest/v2.0/companies/' + CID + '/projects/' + pid + '/commitment_contracts/' + made.id, v2);
    // lines
    for (var j = 0; j < row.lines.length; j++) {
      var l = row.lines[j];
      await pc('POST', '/rest/v1.0/' + ep + '/' + made.id + '/line_items?project_id=' + pid, { line_item: { description: l.desc || row.title, amount: num(l.amount), cost_code_id: Number(l.costCodeId), line_item_type_id: Number(l.typeId), quantity: 1, unit_cost: num(l.amount), uom: 'LS' } });
    }
    log(row.number + ': ' + row.lines.length + ' line(s) added.');
    // attachments: upload each file to Procore storage, then link them by upload id
    var files = await buildAttachments(row, T, vendor), ids = [];
    for (var f = 0; f < files.length; f++) {
      try {
        var u = await api('/procore/upload?project_id=' + pid + '&name=' + encodeURIComponent(files[f].name), { method: 'POST', headers: { 'Content-Type': files[f].type }, body: new Blob([files[f].bytes], { type: files[f].type }) });
        ids.push(u.uuid);
      } catch (e) { log(row.number + ': upload failed for ' + files[f].name + ' — ' + e.message); }
    }
    if (ids.length) {
      var ab = {}; ab[key] = { upload_ids: ids };
      await pc('PATCH', '/rest/v1.0/' + ep + '/' + made.id + '?project_id=' + pid, ab);
      log(row.number + ': attached ' + files.map(function (x) { return x.name; }).join(', ') + '.');
    }
    // read back
    var back = await pc('GET', '/rest/v1.0/' + ep + '/' + made.id + '?project_id=' + pid);
    var miss = Object.keys(cfs).filter(function (k) { var v = back.custom_fields && back.custom_fields[k]; v = v && v.value !== undefined ? v.value : v; return v == null || v === ''; });
    var nAtt = (back.attachments || []).length;
    log(row.number + ': checked in Procore — ' + nAtt + ' attachment(s), retainage ' + back.retainage_percent + '%' + (miss.length ? ', custom fields NOT saved: ' + miss.join(', ') : ', custom fields saved') + '.');
    row.done = 'id ' + made.id;
  }
  async function buildAttachments(row, T, vendor) {
    var TY = TYPES[row.type], out = [], pn = (S.project.project_number ? S.project.project_number + ' ' : '') + S.project.name;
    var city = [S.project.city, S.project.state_code].filter(Boolean).join(', ');
    var d = { projectNumName: pn, projectCity: city, projectNumber: S.project.project_number || '', commitmentNumber: row.number, subName: vendor ? vendor.name : '',
      taxExempt: T.taxExempt, ownerPlatform: T.ownerPlatform, bondRequired: row.bond === 'Yes', builderRiskOwner: T.builderRiskOwner, projectItems: T.items, alternates: T.alternates,
      scopeLines: row.lines.map(function (l) { return { desc: l.desc || row.title, amount: num(l.amount) }; }), submittalDate: row.submittalDate.trim(), scopeName: row.scopeName.trim(),
      csi: linesOf(row.csi), inclusions: linesOf(row.inclusions), exclusions: linesOf(row.exclusions), submitDay: '10th', billingDate: T.billing, retPct: T.retPct, retDuration: T.retDuration };
    var DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    for (var i = 0; i < TY.att.length; i++) {
      var a = TY.att[i];
      try {
        if (a === 'G') {
          var g = 'Attachment G - Subs Pay App - ' + (T.retPct || 5) + (/50/.test(T.retDuration) ? ' thru 50' : '') + '.xlsx';
          out.push({ name: 'Attachment G - Subcontractors Application for Payment.xlsx', type: XLSX, bytes: await BHDocs.fillXlsx(await opsFile(TEMPLATE_DIR + '/' + g), d) });
        } else if (a === 'B') {
          var b = await BHDocs.buildAttachmentB(await opsFile(TEMPLATE_DIR + '/' + ATT_FILES.B), d);
          out.push({ name: 'Attachment B - Scope of Work (' + (d.subName || row.number) + ').docx', type: DOCX, bytes: b.bytes });
          log(row.number + ': Attachment B — ' + b.log.join('; '));
        } else {
          out.push({ name: ATT_NAMES[a], type: DOCX, bytes: await BHDocs.fillDocx(await opsFile(TEMPLATE_DIR + '/' + ATT_FILES[a]), d) });
        }
      } catch (e) { log(row.number + ': could not build attachment ' + a + ' — ' + e.message); }
    }
    if (row.riderPick && row.riderPick !== 'none') {
      try { out.push({ name: row.riderPick, type: /\.pdf$/i.test(row.riderPick) ? 'application/pdf' : DOCX, bytes: await opsFile(RIDER_DIR + '/' + row.rider.folder + '/' + row.riderPick) }); }
      catch (e) { log(row.number + ': could not fetch the rider — ' + e.message); }
    }
    return out;
  }

  /* ---------------- navigation ---------------- */
  function bindNav() {
    var show = function (v) {
      $$('.cd-view').forEach(function (s) { s.hidden = s.id !== 'v-' + v; });
      $$('.bh-side-item[data-view]').forEach(function (a) { if (a.dataset.view === v) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current'); });
      if (v === 'commitments' && S.project) renderRows();
    };
    $$('.bh-side-item[data-view]').forEach(function (a) { a.addEventListener('click', function (e) { e.preventDefault(); history.replaceState(null, '', '#' + a.dataset.view); show(a.dataset.view); }); });
    show(/commitments/.test(location.hash) ? 'commitments' : 'prime');
    $('#pc-read').onclick = readContract;
    $('#pc-save').onclick = savePrime;
    $('#pc-files').onchange = async function (e) {
      var fs = Array.prototype.slice.call(e.target.files || []);
      for (var i = 0; i < fs.length; i++) S.docs.push({ name: fs[i].name, bytes: new Uint8Array(await fs[i].arrayBuffer()), source: 'local' });
      renderDocs(); $('#pc-read').disabled = !S.pcRec; e.target.value = '';
    };
    bindRows();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
