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
  var SCOPES = ['User.Read', 'Sites.Selected', 'Mail.Send'];   // Mail.Send: the hub emails payables@ on approval and sends test reminders
  var SP_HOST = 'baileyharris.sharepoint.com';
  var OPS_SITE = '/sites/Operations', OPS_LIB = 'Documents';
  var TEMPLATE_DIR = '03 Forms and Templates/Contract Documents/Commitments';
  var RIDER_DIR = '01 Working/Project Management/Approved Riders';
  var CF = { contractType: 254924, aiaForm: 459730, costPlus: 269473, splitSavings: 269474, ppBond: 307547, coTerms: 426345, weather: 269484, ownerDelays: 269485, claims: 269490, insDocs: 488034, brExpires: 488035, coiExpires: 488036,
    validatedBy: null, validatedDate: null, validationNotes: null,   // not written to commitments: the PC Verified By / Date on the prime contract is the record
    billing: 194473, taxExempt: 640523, ownerPlatform: 640528, subNotice: 640533, subItems: 640530, alternates: 640537,
    verifiedBy: 640531, verifiedDate: 640532, retDuration: 318332, ld: 269476, payTerms: 269489, finalPay: 322642, builderRisk: 320426,
    ownerAI: 640766 /* PC: Owner Additional Insureds (flows to Attachment F) */, insReq: 274391, pdfType: 194471, poPdfType: 325741, vendorType: 536773, bhSignatory: 300705, subSignatory: 302837, bond: 526946, riderNeeded: 542583 };
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
  async function tokens(force) {
    try {
      var t = await app.acquireTokenSilent({ scopes: SCOPES, account: acct, forceRefresh: !!force });
      // MSAL can hand back a cached ID token that has already expired; the Worker checks it, so refresh it here
      var exp = t.idTokenClaims && t.idTokenClaims.exp;
      if (!force && exp && exp * 1000 < Date.now() + 5 * 60e3) t = await app.acquireTokenSilent({ scopes: SCOPES, account: acct, forceRefresh: true });
      return t;
    }
    catch (e) { await app.acquireTokenRedirect({ scopes: SCOPES, account: acct }); throw e; }
  }

  /* ---------------- Worker / Procore ---------------- */
  async function api(path, opt) {
    var t = await tokens(); opt = opt || {};
    for (var attempt = 0; ; attempt++) {
      var r = await fetch(API + path, { method: opt.method || 'GET', headers: Object.assign({ Authorization: 'Bearer ' + t.idToken }, opt.headers || {}), body: opt.body });
      if (r.status === 401 && attempt === 0) { t = await tokens(true); continue; }
      if (r.status === 429 && attempt < 4) {   // Procore rate limit: wait and try again
        var wait = Number(r.headers.get('Retry-After')) || (3 + attempt * 5);
        msg('Procore is busy (rate limit). Retrying in ' + wait + ' seconds…');
        await new Promise(function (ok) { setTimeout(ok, wait * 1000); });
        continue;
      }
      var txt = await r.text(), j = null; try { j = txt ? JSON.parse(txt) : null; } catch (_) { j = { raw: txt }; }
      if (!r.ok) {
        var er = j && j.error, text = typeof er === 'string' ? er : er && er.message ? er.message : 'HTTP ' + r.status;
        var det = er && Array.isArray(er.details) ? er.details.map(function (d) { return d.message; }).join('; ') : er && er.details && er.details.message ? er.details.message : '';
        var e = new Error(text + (det ? ' — ' + det : '') + (r.status === 429 ? ' (Procore rate limit; try again in a few minutes)' : '')); e.status = r.status; e.body = j; throw e;
      }
      if (attempt) msg('');
      return j;
    }
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
      var lovIds = [CF.contractType, CF.aiaForm, CF.ppBond, CF.retDuration, CF.taxExempt, CF.builderRisk, CF.pdfType, CF.poPdfType, CF.vendorType, CF.bond, CF.riderNeeded];
      for (var i = 0; i < lovIds.length; i++) {
        try { S.lov[lovIds[i]] = await pc('GET', '/rest/v1.0/custom_field_definitions/' + lovIds[i] + '/custom_field_lov_entries?company_id=' + CID); }
        catch (e) { S.lov[lovIds[i]] = []; msg('Could not load the options for Procore field ' + lovIds[i] + ': ' + e.message + '. Reload the page before saving or creating.', true); }
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
    if (window.BHHub) window.BHHub.job();
    stamp();
  }

  /* =====================================================================
     PRIME CONTRACT TAB
     ===================================================================== */
  /* every field on the Prime Contract fieldset (443377), in its sections and order. kind 'ro' = shown, edited in Procore */
  var who = function (o) { return o ? o.name || o.company || o.login || '' : ''; };
  var FIELDS = [
    { section: 'General information' },
    { key: 'title', label: 'Title', std: 'title', kind: 'text' },
    { key: 'number', label: 'Contract #', std: 'number', kind: 'ro' },
    { key: 'owner', label: 'Owner / client', kind: 'ro', get: function (r) { return who(r.vendor); } },
    { key: 'contractor', label: 'Contractor', kind: 'ro', get: function (r) { return who(r.contractor); } },
    { key: 'architect', label: 'Architect / engineer', kind: 'ro', get: function (r) { return who(r.architect); } },
    { key: 'status', label: 'Status', kind: 'ro', get: function (r) { return r.status || ''; } },
    { key: 'executed', label: 'Executed', kind: 'ro', get: function (r) { return r.executed ? 'Yes' : 'No'; } },
    { key: 'signature_required', label: 'Sign with DocuSign', kind: 'ro', get: function (r) { return r.signature_required ? 'Yes' : 'No'; } },
    { key: 'description', label: 'Description', std: 'description', kind: 'rich' },
    { key: 'inclusions', label: 'Inclusions', std: 'inclusions', kind: 'rich' },
    { key: 'exclusions', label: 'Exclusions', std: 'exclusions', kind: 'rich' },
    { key: 'contract_date', label: 'Contract date', std: 'contract_date', kind: 'date' },
    { key: 'contract_start_date', label: 'Start date', std: 'contract_start_date', kind: 'date' },
    { key: 'contract_estimated_completion_date', label: 'Estimated completion date', std: 'contract_estimated_completion_date', kind: 'date' },
    { key: 'substantial_completion_date', label: 'Substantial completion date', std: 'substantial_completion_date', kind: 'date' },
    { key: 'actual_completion_date', label: 'Actual completion date', std: 'actual_completion_date', kind: 'date' },
    { key: 'contract_termination_date', label: 'Termination date', std: 'contract_termination_date', kind: 'date' },
    { key: 'execution_date', label: 'Execution date', std: 'execution_date', kind: 'date' },
    { key: 'retainage_percent', label: 'Default retainage %', std: 'retainage_percent', kind: 'num' },
    { key: 'attachments', label: 'Attachments (the contract the desk reads)', kind: 'ro', get: function (r) { var a = r.attachments || []; return a.length ? a.length + ' file(s): ' + a.map(function (x) { return x.name || x.filename; }).join(', ') : ''; } },
    { section: 'Prime contract summary' },
    { key: 'contract_type', label: 'Contract type', cf: CF.contractType, kind: 'lovm' },
    { key: 'aia_form', label: 'AIA contract form', cf: CF.aiaForm, kind: 'lov' },
    { key: 'cost_plus_percent', label: 'Cost-Plus %', cf: CF.costPlus, kind: 'text' },
    { key: 'split_savings_percent', label: 'Split-Savings %', cf: CF.splitSavings, kind: 'text' },
    { key: 'retainage_duration', label: 'Retainage duration', cf: CF.retDuration, kind: 'lov' },
    { key: 'builders_risk_provider', label: "Builder's risk provider", cf: CF.builderRisk, kind: 'lov' },
    { key: 'liquidated_damages', label: 'Liquidated damages', cf: CF.ld, kind: 'text' },
    { key: 'pp_bond', label: 'Performance and payment bond', cf: CF.ppBond, kind: 'lovm' },
    { key: 'change_order_terms', label: 'Change order terms', cf: CF.coTerms, kind: 'rich', hint: 'Article reference? Allowed markup? What can and cannot be included?' },
    { key: 'payment_terms', label: 'Payment terms (owner)', cf: CF.payTerms, kind: 'rich', hint: 'Article reference? Date BH pay app is due to Owner/Architect, payment terms once submitted, remedy for late or non-payment, items not subject to retainage.' },
    { key: 'weather_delays', label: 'Weather delays', cf: CF.weather, kind: 'rich', hint: 'Article reference? How are weather delays determined? Notice requirements? Dry-out days?' },
    { key: 'owner_delays', label: 'Owner delays', cf: CF.ownerDelays, kind: 'rich', hint: 'Article reference? Are owner or architect delays permissible and compensable? Notice requirements?' },
    { key: 'claims', label: 'Claims', cf: CF.claims, kind: 'rich', hint: 'Article reference? Window to start a claim after discovery, what counts as notice, special notice conditions.' },
    { key: 'final_payment', label: 'Final payment / closeout and other notes', cf: CF.finalPay, kind: 'rich' },
    { section: 'Subcontractor flow-down (builds every commitment on this job)' },
    { key: 'billing_date', label: 'Monthly Billing Date (trade partner pay apps)', cf: CF.billing, kind: 'text', dflt: '20th', hint: 'Last date a trade partner pay app is accepted for the month (Attachment B: no later than this date, no projection of work). This is not in the owner contract; the owner date is under Payment terms.' },
    { key: 'tax_exempt', label: 'Tax exempt', cf: CF.taxExempt, kind: 'lov' },
    { key: 'owner_platform', label: 'Owner specific platform required (unticked = Procore)', cf: CF.ownerPlatform, kind: 'bool' },
    { key: 'sub_notice', label: 'Sub notice deadlines', cf: CF.subNotice, kind: 'text', dflt: 'Same as PC.', hint: 'Leave "Same as PC." unless trade partners get fewer days than the owner gives Bailey-Harris. Long Form 8.3 already sets 48 hours for change and differing-condition proposals.' },
    { key: 'sub_project_items', label: 'Sub project items (everything job-specific for Attachment B)', cf: CF.subItems, kind: 'list', rows: 14, hint: 'One item per line. Goes into Attachment B, Project Specific General Items, in this order. Include the owner rules and the site and operational items PMs usually add (parking, lay-down area, utility shutdown notice, testing agency, superintendent on site, inspections, and so on).' },
    { key: 'accepted_alternates', label: 'Accepted alternates and VE', cf: CF.alternates, kind: 'alts', hint: 'One per line: description and amount | Accepted (or Not Accepted).' },
    { key: 'owner_additional_insureds', label: 'Owner additional insureds (flows to Attachment F)', cf: CF.ownerAI, kind: 'rich', hint: 'Exactly as the owner contract lists them, with the coverages and endorsements it requires. Leave blank to use our default: Bailey-Harris Construction Co., Inc., Owner, Architect, and their officers, agents, consultants, and employees.' },
    { key: 'verified_by', label: 'PC verified by', kind: 'ro', get: function (r) { return who(cfVal(r, CF.verifiedBy)) || lovLabel(cfVal(r, CF.verifiedBy)); }, hint: 'Stamped when you confirm.' },
    { key: 'verified_date', label: 'PC verified date', kind: 'ro', get: function (r) { return String(cfVal(r, CF.verifiedDate) || '').slice(0, 10); }, hint: 'Stamped when you confirm.' },
    { section: 'Project insurance' },
    { key: 'ins_docs', label: 'Insurance and bond documents (builders risk, BH project COI, P&P bonds)', kind: 'ro', get: function (r) { var v = cfVal(r, CF.insDocs) || []; return v.length ? v.length + ' file(s): ' + v.map(function (x) { return x.name || x.filename || ''; }).join(', ') : ''; } },
    { key: 'br_expires', label: "Builder's risk insurance expiration date", cf: CF.brExpires, kind: 'datetime' },
    { key: 'coi_expires', label: 'BH project COI expiration date', cf: CF.coiExpires, kind: 'datetime' }
  ];
  FIELDS = FIELDS.filter(function (f) { return f.section || f.std || f.cf != null || f.get; });
  function cfVal(rec, id) { var v = rec && rec.custom_fields && rec.custom_fields['custom_field_' + id]; return v && v.value !== undefined ? v.value : v; }
  function lovLabel(v) { if (v == null) return ''; if (Array.isArray(v)) return v.map(lovLabel).join(', '); if (typeof v === 'object') return v.label || v.name || ''; return String(v); }
  function currentText(f, rec) {
    if (f.get) return f.get(rec) || '';
    var v = f.std ? rec[f.std] : cfVal(rec, f.cf);
    if (f.kind === 'lov' || f.kind === 'lovm') return lovLabel(v);
    if (f.kind === 'date' || f.kind === 'datetime') return v ? String(v).slice(0, 10) : '';
    if (f.kind === 'bool') return v === true ? 'Ticked' : v === false ? 'Not ticked' : '';
    if (f.kind === 'rich' || f.kind === 'list' || f.kind === 'alts') return htmlToText(v);
    if (f.kind === 'num') return v == null ? '' : String(Number(v));
    return v == null ? '' : String(v);
  }
  async function loadPrime() {
    $('#pc-table tbody').innerHTML = '<tr><td colspan="4">Loading the prime contract&hellip;</td></tr>';
    try { S.pcRec = await pc('GET', '/rest/v1.0/prime_contract?project_id=' + S.project.id); }
    catch (e) { S.pcRec = null; $('#pc-table tbody').innerHTML = '<tr><td colspan="4">' + (e.status === 404 ? 'No prime contract found in Procore for this job. Create it in Procore first (Part 1, step 1 of the PM guide).' : 'Could not load the prime contract: ' + esc(e.message) + '. Pick the job again or reload.') + '</td></tr>'; $('#pc-read').disabled = true; $('#pc-save').disabled = true; return; }
    S.docs = (S.pcRec.attachments || []).filter(function (a) { return /\.pdf$/i.test(a.name || a.filename || ''); }).map(function (a) { return { name: a.name || a.filename, url: a.url, source: 'procore' }; });
    renderDocs(); renderPrime();
    if (window.BHHub) window.BHHub.summary();
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
      if (f.section) return '<tr class="cd-sec"><td colspan="4">' + esc(f.section) + '</td></tr>';
      var cur = currentText(f, rec), p = S.proposals[f.key], val = p && p.value != null ? p.value : (cur || f.dflt || '');
      var input;
      if (f.kind === 'ro') {
        input = '<span class="bh-caption">Edit in Procore</span>';
      } else if (f.kind === 'lov' || f.kind === 'lovm') {
        var opts = S.lov[f.cf] || [];
        // the reader's answer matched to Procore's own labels; no match keeps what Procore has and says so
        var curIds = [].concat(cfVal(rec, f.cf) || []).map(function (x) { return String(x && x.id != null ? x.id : ''); }).filter(Boolean);
        var hits = p && p.value != null ? matchOpts(f, p.value) : null;
        var selIds = hits && hits.length ? hits.map(function (o) { return String(o.id); }) : curIds;
        if (hits && !hits.length) p.nomatch = true;
        var isSel = function (o) { return selIds.indexOf(String(o.id)) >= 0; };
        input = '<select data-k="' + f.key + '"' + (f.kind === 'lovm' ? ' multiple size="' + Math.min(4, Math.max(2, opts.length)) + '"' : '') + '>' + (f.kind === 'lovm' ? '' : '<option value=""></option>') + opts.map(function (o) { return '<option value="' + o.id + '"' + (isSel(o) ? ' selected' : '') + '>' + esc(o.label) + '</option>'; }).join('') + '</select>';
      } else if (f.kind === 'date' || f.kind === 'datetime') {
        input = '<input type="date" data-k="' + f.key + '" value="' + esc(String(val || '').slice(0, 10)) + '">';
      } else if (f.kind === 'bool') {
        var on = p ? !!p.value : cfVal(rec, f.cf) === true;
        input = '<label><input type="checkbox" data-k="' + f.key + '"' + (on ? ' checked' : '') + '> Owner requires its own system</label>';
      } else if (f.kind === 'rich' || f.kind === 'list' || f.kind === 'alts') {
        if (Array.isArray(val)) val = val.map(function (x) { return typeof x === 'object' ? (x.label + ' | ' + (x.status || 'Accepted')) : x; }).join('\n');
        input = '<textarea data-k="' + f.key + '" rows="' + (f.rows || (f.kind === 'rich' ? 4 : 6)) + '">' + esc(val) + '</textarea>';
      } else {
        input = '<input type="text" data-k="' + f.key + '" value="' + esc(val) + '">';
      }
      var src = p && p.source ? '<div class="src">' + (p.nomatch ? '<b class="nomatch">The contract says "' + esc(Array.isArray(p.value) ? p.value.join(', ') : p.value) + '": no Procore option matches it, so the current value is kept. Pick one, or add the option to the list in Procore.</b>' : '') + esc(p.source) + (p.quote ? '<q>' + esc(p.quote) + '</q>' : '') + '</div>' : (f.dflt && !cur ? '<div class="src">Company default</div>' : '');
      var diff = p && p.value != null && normCmp(p.value) !== normCmp(cur) && cur;
      return '<tr' + (diff ? ' class="is-diff" title="The contract reading differs from what Procore holds"' : '') + '><td><b>' + esc(f.label) + '</b>' + (f.hint ? '<div class="hint">' + esc(f.hint) + '</div>' : '') + '</td><td class="cur"><div>' + esc(cur || '—') + '</div></td><td>' + input + '</td><td>' + src + '</td></tr>';
    }).join('');
  }
  /* match a reader answer ("Required", "Contractor", "Stipulated Sum") to the field's Procore options ("Yes", "Bailey Harris", "Fixed Price / Lump Sum") */
  var optNorm = function (x) { return String(x == null ? '' : x).toLowerCase().replace(/bailey[\s-]*harris(\s+construction)?(\s+co\.?)?(,?\s+inc\.?)?/g, 'bailey harris').replace(/[^a-z0-9%]+/g, ' ').trim(); };
  var OPT_SYN = [[/^(required|yes|true|y|provided|included)$/, 'yes'], [/^(not required|no|false|n|none|waived|not included)$/, 'no'],
    [/(payment|performance) bonds? (equal|required|in the amount|of 100)/, 'yes'],
    [/cost plus.*(gmp|guaranteed maximum)|(gmp|guaranteed maximum).*cost plus/, 'cost plus gmp'], [/guaranteed maximum|\bgmp\b/, 'gmp'], [/stipulated sum|lump sum|fixed price|fixed sum/, 'lump sum'], [/cost plus|cost of the work plus/, 'cost plus'],
    [/time and materials?|^t m$|\bt ?& ?m\b/, 't m'], [/unit price/, 'unit price'],
    [/^contractor$|design builder|general contractor|\bgc\b|bailey harris|construction manager/, 'bailey harris'], [/^owner\b/, 'owner'],
    [/(thru|through|until) 50|50 ?% (complete|of)/, 'thru 50'], [/throughout|entire|until final|full term/, 'throughout']];
  var optKey = function (x) { var n = optNorm(x); for (var i = 0; i < OPT_SYN.length; i++) if (OPT_SYN[i][0].test(n)) return OPT_SYN[i][1]; return n; };
  function matchOpts(f, val) {
    var opts = S.lov[f.cf] || [];
    var vals = Array.isArray(val) ? val.map(function (x) { return x && typeof x === 'object' ? (x.label || x.name || '') : x; })
      : f.kind === 'lovm' ? String(val && typeof val === 'object' ? lovLabel(val) : val).split(/\s*[;\n]\s*|\s*,\s+(?=[A-Za-z])/) : [val && typeof val === 'object' ? lovLabel(val) : val];
    var out = [];
    vals.forEach(function (v) {
      if (!String(v || '').trim()) return;
      var nv = optNorm(v), kv = optKey(v);
      var hit = opts.find(function (o) { return optNorm(o.label) === nv; }) || opts.find(function (o) { return optKey(o.label) === kv; });
      if (!hit) { var c = opts.filter(function (o) { var ko = optKey(o.label); return ko.length > 2 && kv.length > 2 && (kv.indexOf(ko) >= 0 || ko.indexOf(kv) >= 0); }); if (c.length === 1) hit = c[0]; }
      if (hit && out.indexOf(hit) < 0) out.push(hit);
    });
    return f.kind === 'lovm' ? out : out.slice(0, 1);
  }
  var normCmp = function (v) { if (Array.isArray(v)) v = v.map(function (x) { return typeof x === 'object' ? x.label : x; }).join(' '); return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9%]/g, ''); };

  /* ---- reading the contract (in this browser) ---- */
  var KEYWORDS = /retain|retention|application for payment|pay(ment)? application|progress payment|final payment|liquidated|notice|calendar days|days after|weather|delay|change order|claim|sales (and use )?tax|tax[- ]exempt|exemption|builder'?s risk|alternate|value engineering|football|graduation|parking|signage|drone|unmanned|smok|firearm|harass|radio|safety|confidential|non-?disclosure|platform|e-builder|procore|textura|contract sum|contract time|substantial completion/i;
  // a page is sent only when it carries at least two of the terms the desk fills (scored by distinct topics)
  var TOPICS = [/retain|retention/i, /application for payment|pay(ment)? application|progress payment/i, /final payment|closeout/i, /liquidated damages/i,
    /sales (and use )?tax|tax[- ]exempt|exemption/i, /builder'?s risk/i, /additional insured/i, /alternate|value engineering/i,
    /weather|rain day/i, /notice of (claim|delay)|written notice|within \w+ \(\d+\) (calendar )?days/i, /change order|change proposal/i,
    /football|graduation|parking|signage|drone|unmanned|smok|firearm|harass|radio/i, /confidential|non-?disclosure/i, /contract sum|contract time|substantial completion/i];
  // pages on these terms are always sent: each one sets a field on its own
  var KEY = /retainage|liquidated damages|additional insured|tax[- ]exempt|sales and use tax|builder'?s risk|alternate|confidential|non-?disclosure/i;
  function score(t) { return TOPICS.reduce(function (n, re) { return n + (re.test(t) ? 1 : 0); }, 0); }
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
  /* the Procore option labels for the reader's pick-list fields, so it answers with a label the desk can select */
  function readerOptions() {
    var keys = { contract_type: 'contract_type', aia_form: 'aia_form', retainage_duration: 'retainage_duration', builders_risk_provider: 'builders_risk_provider', pp_bond: 'performance_payment_bond', tax_exempt: 'tax_exempt' }, out = {};
    FIELDS.forEach(function (f) { if (keys[f.key] && f.cf && (S.lov[f.cf] || []).length) out[keys[f.key]] = { labels: S.lov[f.cf].map(function (o) { return o.label; }), multiple: f.kind === 'lovm' }; });
    return out;
  }
  async function readContract() {
    var btn = $('#pc-read'); btn.disabled = true; msg('Reading the contract PDFs in this browser…');
    try {
      var sections = [], all = '', scanned = [];
      for (var i = 0; i < S.docs.length; i++) {
        var d = S.docs[i], st = $('#doc-' + i);
        try {
          if (!d.bytes) {   // Procore files come through the Worker (the browser can't fetch Procore storage directly)
            var tk = await tokens();
            var r = await fetch(API + '/procore/file?url=' + encodeURIComponent(d.url), { method: 'POST', headers: { Authorization: 'Bearer ' + tk.idToken } });
            if (!r.ok) throw new Error('HTTP ' + r.status);
            d.bytes = new Uint8Array(await r.arrayBuffer());
          }
        } catch (e) { d.status = 'Could not download from Procore in this browser. Add this PDF with "Add contract PDFs".'; if (st) st.textContent = d.status; continue; }
        var pages = await pdfPages(d.bytes.slice(0));
        var textPages = pages.filter(function (t) { return t.trim().length > 80; }).length;
        if (textPages < pages.length / 2) { scanned.push(d.name); d.status = 'Scanned (' + pages.length + ' pages, no text). Server text recognition is not set up yet; fill those terms by hand.'; if (st) st.textContent = d.status; continue; }
        var picked = 0, cands = [];
        pages.forEach(function (t, k) {
          all += '\n' + t;
          if (BANK.test(t)) return;          // never send a page with banking details
          var sc = score(t);
          if (sc >= 2 || KEY.test(t)) cands.push({ doc: d.name, page: k + 1, text: t, sc: sc + (KEY.test(t) ? 2 : 0) });
        });
        cands.sort(function (a, b) { return b.sc - a.sc; }).slice(0, 24).sort(function (a, b) { return a.page - b.page; })
          .forEach(function (c) { sections.push({ doc: c.doc, page: c.page, text: c.text }); picked++; });
        d.status = pages.length + ' pages, ' + picked + ' with contract terms'; if (st) st.textContent = d.status;
      }
      var conf = confidentiality(all);
      if (conf.found && conf.blocked) { msg('This contract has a confidentiality clause that does not allow sharing with service providers, so nothing was sent for reading. Fill the fields from the contract by hand. Clause: "' + conf.quote + '"', true); return; }
      if (!sections.length) { msg(scanned.length ? 'Only scanned PDFs were found (' + scanned.join(', ') + '). Fill the fields by hand from the contract.' : 'No contract terms found in the PDFs.', true); return; }
      msg('Sending ' + sections.length + ' pages of contract terms for reading…' + (conf.found ? ' (A confidentiality clause was found; it allows sharing with service providers.)' : ''));
      var res = await api('/read', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ project: (S.project.project_number || '') + ' ' + S.project.name, sections: sections, options: readerOptions() }) });
      var f = res.fields || {};
      S.proposals = {};
      var map = { retainage_percent: 'retainage_percent', retainage_duration: 'retainage_duration', contract_date: 'contract_date', tax_exempt: 'tax_exempt', liquidated_damages: 'liquidated_damages',
        payment_terms: 'payment_terms', sub_project_items: 'sub_project_items', contract_type: 'contract_type', aia_form: 'aia_form', cost_plus_percent: 'cost_plus_percent', split_savings_percent: 'split_savings_percent',
        performance_payment_bond: 'pp_bond', change_order_terms: 'change_order_terms', weather_delays: 'weather_delays', owner_delays: 'owner_delays', claims: 'claims', contract_start_date: 'contract_start_date',
        substantial_completion_date: 'substantial_completion_date', contract_estimated_completion_date: 'contract_estimated_completion_date', description: 'description', accepted_alternates: 'accepted_alternates', builders_risk_provider: 'builders_risk_provider', final_payment: 'final_payment', owner_additional_insureds: 'owner_additional_insureds' };
      var empty = function (v) { return v == null || (typeof v === 'string' && !v.trim()) || (Array.isArray(v) && !v.length); };
      Object.keys(map).forEach(function (k) { if (f[k] && !empty(f[k].value)) S.proposals[map[k]] = f[k]; });
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
        else if (f.kind === 'ro') return;
        else if (f.kind === 'lovm') { var ids = Array.prototype.slice.call(el.selectedOptions).map(function (o) { return Number(o.value); }).sort(); var curIds = (cfVal(rec, f.cf) || []).map(function (x) { return Number(x.id); }).sort(); if (!ids.length || ids.join() === curIds.join()) return; v = ids; }
        else if (f.kind === 'date') { if (!el.value || el.value === cur) return; v = el.value; }
        else if (f.kind === 'datetime') { if (!el.value || el.value === cur) return; v = el.value + 'T18:00:00Z'; }
        else if (f.kind === 'lov') { if (!el.value) return; var o = (S.lov[f.cf] || []).find(function (x) { return String(x.id) === el.value; }); if (o && o.label === cur) return; v = Number(el.value); }
        else if (f.kind === 'rich' || f.kind === 'list' || f.kind === 'alts') { if (normCmp(el.value) === normCmp(cur) || !el.value.trim()) return; v = toRich(f.kind, el.value); }
        else if (f.kind === 'num') { if (el.value === '' || Number(el.value) === Number(cur)) return; v = Number(el.value); }
        else { if (el.value.trim() === cur.trim() || !el.value.trim()) return; v = el.value.trim(); }
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
      ownerAI: (function () { var t = CF.ownerAI ? htmlToText(cfVal(r, CF.ownerAI)).trim() : ''; return /^Bailey-Harris Construction Co\., Inc\., Owner, Architect, and their officers, agents, consultants, and employees\.?$/i.test(t) ? '' : t; })(), // blank or our default = Attachment F's standard line only
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
    R.woc = Array.isArray(woc) ? woc : []; R.poc = Array.isArray(poc) ? poc : [];
    R.numbers = [].concat(R.woc, R.poc).map(function (c) { return String(c.number || ''); });
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
    return Object.assign({ id: Math.random().toString(36).slice(2), open: true, type: 'LF', vendorId: '', title: '', signatoryId: '', bhSignatoryId: S.me ? S.me.id : '', bond: 'No', scopeName: '', csi: '', submittalDate: '', proposalDate: '', adds: '', bMode: '', ownB: null, items: (function () { try { var T = pcTerms(); return T ? T.items.join('\n') : ''; } catch (_) { return ''; } })(), equipAvail: false, inclusions: '', exclusions: '', description: '', lines: [{ costCodeId: '', typeId: '', desc: '', amount: '' }], rider: null, riderPick: '' }, seed || {});
  }
  async function renderRows() {
    var host = $('#cm-rows'), T = pcTerms();
    $('#cm-add').disabled = $('#cm-paste').disabled = $('#cm-bfiles').disabled = $('#cm-create').disabled = true;
    if (!S.project) { host.innerHTML = ''; return; }
    if (!T) { $('#cm-pc').innerHTML = 'No prime contract in Procore for this job.'; host.innerHTML = ''; return; }
    $('#cm-pc').innerHTML = '<span>Retainage <b>' + T.retPct + '% ' + esc(T.retDuration) + '</b></span><span>Monthly Billing Date <b>' + esc(T.billing) + '</b></span><span>Tax exempt <b>' + (T.taxExempt ? 'Yes' : 'No') + '</b></span><span>Platform <b>' + (T.ownerPlatform ? 'Owner system' : 'Procore') + '</b></span><span>Owner rules <b>' + T.items.length + '</b></span><span>Alternates <b>' + T.alternates.length + '</b></span>';
    if (!T.verified) { host.innerHTML = '<div class="cd-msg is-bad">The prime contract is not verified yet. Confirm it on the Prime Contract tab first.</div>'; return; }
    host.innerHTML = '<p class="bh-caption">Loading vendors, cost codes and existing commitments&hellip;</p>';
    await loadRef();
    if (!S.rows.length) S.rows.push(newRow());
    $('#cm-add').disabled = $('#cm-paste').disabled = $('#cm-bfiles').disabled = $('#cm-create').disabled = false;
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
      var T = TYPES[row.type], iss = issues(row), lf = row.type === 'LF', own = lf && row.bMode === 'own';
      var total = row.lines.reduce(function (s, l) { return s + num(l.amount); }, 0);
      var signed = acked(row, iss), hasWarn = iss.some(function (x) { return !x.bad; });
      var state = iss.some(function (x) { return x.bad; }) ? '<span class="bh-chip is-bad">Blocked</span>' : !signed ? '<span class="bh-chip is-warn">Needs PM sign-off</span>' : hasWarn ? '<span class="bh-chip is-ok">Signed off</span>' : '<span class="bh-chip is-ok">Validated</span>';
      if (row.done) state = '<span class="bh-chip">Created ' + esc(row.done) + '</span>';
      return '<div class="cd-row" data-i="' + i + '"><div class="cd-row-h"><span class="num">' + esc(row.number) + '</span><span class="ttl">' + esc(row.title || 'New commitment') + ' · ' + esc(vendorName(row.vendorId)) + ' · ' + money(total) + '</span>' + issueSummary(iss) + state +
        '<button class="bh-btn-outline" data-act="toggle">' + (row.open ? 'Hide details' : 'Details') + '</button><button class="bh-btn-outline" data-act="del">Remove</button></div>' + (!row.open ? '</div>' : '<div class="cd-row-b">' +
        '<label>Type<select data-f="type">' + Object.keys(TYPES).map(function (k) { return '<option value="' + k + '"' + (k === row.type ? ' selected' : '') + '>' + TYPES[k].label + '</option>'; }).join('') + '</select></label>' +
        '<label class="wide">Vendor<select data-f="vendorId">' + vOpts(row.vendorId) + '</select></label>' +
        '<label>P&amp;P bond required<select data-f="bond"><option' + (row.bond === 'No' ? ' selected' : '') + '>No</option><option' + (row.bond === 'Yes' ? ' selected' : '') + '>Yes</option></select></label>' +
        '<label class="wide">Title<input data-f="title" value="' + esc(row.title) + '"></label>' +
        '<label>Commitment signatory (vendor)<select data-f="signatoryId">' + uOpts(row.signatoryId, row.vendorId) + '</select></label>' +
        '<label>BH signatory (project PM)<select data-f="bhSignatoryId">' + uOpts(row.bhSignatoryId) + '</select></label>' +
        (lf ? '<label class="wide">Attachment B: choose one (required)<select data-f="bMode"' + (row.bMode ? '' : ' class="is-missing"') + '><option value=""' + (row.bMode ? '' : ' selected') + ' disabled>Choose one…</option><option value="build"' + (row.bMode === 'build' ? ' selected' : '') + '>Build it here, pre-filled from the prime contract</option><option value="own"' + (own ? ' selected' : '') + '>Upload my own; the desk checks it</option></select></label>' : '') +
        (lf && own ? '<label class="full">Your Attachment B (Word or PDF). The desk checks it against the prime contract and attaches it as is.<input type="file" data-own-b accept=".docx,.pdf">' + (row.ownB ? '<span class="bh-caption">' + esc(row.ownB.name) + (row.ownB.reading ? ' — reading…' : '') + '</span>' : '') + '</label>' : '') +
        (lf && row.bMode === 'build' ? '<label class="wide">Scope name (Attachment B)<input data-f="scopeName" value="' + esc(row.scopeName) + '" placeholder="e.g. Drywall and Framing"></label>' +
              '<label class="wide">Submittal deadline<input data-f="submittalDate" value="' + esc(row.submittalDate) + '" placeholder="e.g. Friday, December 18, 2026"></label>' +
              '<label class="full">Specification sections (one per line)<textarea data-f="csi" rows="2">' + esc(row.csi) + '</textarea></label>' +
              '<label class="wide">Proposal date (cost table)<input data-f="proposalDate" value="' + esc(row.proposalDate) + '" placeholder="e.g. June 5th, 2025"></label>' +
              '<div class="full cd-field"><label for="adds-' + row.id + '">Alternates and adds in this trade partner\'s price (cost table). One per line: description | amount. Base Bid = commitment total minus these.</label>' +
              altPicks(row) +
              '<textarea id="adds-' + row.id + '" data-f="adds" rows="2" placeholder="Alternate 1 – Wood Panels, Circulation Areas | 142000">' + esc(row.adds) + '</textarea></div>' +
              '<label class="full">Project specific items (Section B), one per line. Starts from the prime contract; edit for this trade partner.<textarea data-f="items" rows="6">' + esc(row.items) + '</textarea></label>' +
              '<label class="full cd-check"><input type="checkbox" data-f="equipAvail"' + (row.equipAvail ? ' checked' : '') + '><span>Let this trade partner use BH equipment on site for loading/unloading, with a usage agreement. Leave unticked for the standard General Item 33 (BH equipment not available).</span></label>'
            : '') +
        (!lf ? '<label class="full">Description<textarea data-f="description" rows="2">' + esc(row.description) + '</textarea></label>' : '') +
        (lf && row.bMode !== 'build' ? '' : '<label class="wide">' + (lf ? 'Scope inclusions for Attachment B (one per line)' : 'Inclusions (one per line)') + '<textarea data-f="inclusions">' + esc(row.inclusions) + '</textarea></label>' +
        '<label class="wide">' + (lf ? 'Accepted exclusions for Attachment B (one per line). Already added for you: Dumpsters (always), Bonds (unless this commitment needs a P&P bond), Sales taxes (only on tax-exempt jobs).' : 'Exclusions accepted (one per line)') + '<textarea data-f="exclusions">' + esc(row.exclusions) + '</textarea></label>') +
        '<div class="cd-lines"><table><thead><tr><th style="width:30%">Cost code</th><th style="width:16%">Cost type</th><th>Description</th><th style="width:14%">Amount</th><th style="width:40px"></th></tr></thead><tbody>' +
        row.lines.map(function (l, j) { return '<tr data-j="' + j + '"><td><select data-l="costCodeId">' + ccOpts(l.costCodeId) + '</select></td><td><select data-l="typeId">' + tOpts(l.typeId) + '</select></td><td><input data-l="desc" value="' + esc(l.desc) + '"></td><td><input class="amt" data-l="amount" value="' + esc(l.amount) + '"></td><td><button class="bh-btn-outline" data-act="dl" aria-label="Remove line">&times;</button></td></tr>'; }).join('') +
        '</tbody></table><button class="bh-btn-outline" data-act="al" style="margin-top:6px">Add line</button></div>' +
        '<label class="wide">Rider<select data-f="riderPick">' + riderOpts(row) + '</select></label>' +
        '<div class="wide bh-caption" style="align-self:end">Attachments: ' + (lf ? 'A (copied from the job), ' : '') + T.att.map(function (a) { return a === 'B' ? (own ? 'B (yours)' : 'B (built)') : a; }).join(', ') + (row.type === 'SF' ? ' (Short Form F version not in the templates yet; the Subcontract F is used)' : '') + '</div>' +
        (iss.length ? '<ul class="cd-issues">' + iss.map(function (x) { return '<li class="' + (x.bad ? 'bad' : 'warn') + '">' + esc(x.t) + '</li>'; }).join('') + '</ul>' : '') +
        '</div></div>');
    }).join('') + signoffPanel();
  }
  function vendorName(id) { var v = S.ref && S.ref.vendors.find(function (x) { return String(x.id) === String(id); }); return v ? v.name : 'no vendor'; }
  function issueSummary(iss) {
    var nb = iss.filter(function (x) { return x.bad; }).length, nw = iss.length - nb;
    return nb ? '<span class="cd-sum bad">' + nb + ' to fix</span>' : nw ? '<span class="cd-sum warn">' + nw + ' warning' + (nw > 1 ? 's' : '') + '</span>' : '';
  }
  /* one sign-off for every warning in the batch; red items still block their own rows */
  function signoffPanel() {
    var pend = S.rows.filter(function (r) { return !r.done; }).map(function (r) { return { r: r, iss: issues(r) }; })
      .filter(function (x) { return !x.iss.some(function (i) { return i.bad; }) && warnKey(x.iss); });
    if (!pend.length) return '';
    var all = pend.every(function (x) { return acked(x.r, x.iss); }), n = pend.reduce(function (s, x) { return s + x.iss.length; }, 0);
    return '<div class="cd-signoff"><div class="bh-section-title">Warnings to sign off (' + n + ' on ' + pend.length + ' commitment' + (pend.length > 1 ? 's' : '') + ')</div><ul>' +
      pend.map(function (x) { return '<li><b>' + esc(x.r.number) + ' ' + esc(x.r.title) + '</b> · ' + esc(vendorName(x.r.vendorId)) + '<ul>' + x.iss.map(function (i) { return '<li>' + esc(i.t) + '</li>'; }).join('') + '</ul></li>'; }).join('') +
      '</ul><label class="cd-ack"><input type="checkbox" data-ack-all' + (all ? ' checked' : '') + '> I reviewed every warning above and they are right for these commitments. I am the validator of record (' + esc(S.me ? S.me.name || S.me.login : 'me') + ').</label></div>';
  }
  /* the prime contract's accepted alternates, as one-click lines for the cost table (the PM adds this trade partner's amount) */
  function pcAltName(label) { return String(label).replace(/\s*[-–]?\s*\$[\d,]+(\.\d+)?\s*$/, '').trim(); }
  function altPicks(row) {
    var T = pcTerms(), alts = (T ? T.alternates : []).filter(function (a) { return /^accepted/i.test(a.status) && !/^base bid/i.test(a.label); });
    if (!alts.length) return '<div class="full bh-caption">No accepted alternates on the prime contract.</div>';
    var inAdds = parseAdds(row.adds).map(function (a) { return a.desc.toLowerCase(); });
    return '<div class="full bh-caption cd-alts">Add any of the following to this subcontract: ' + alts.map(function (a) {
      var n = pcAltName(a.label), used = inAdds.indexOf(n.toLowerCase()) >= 0;
      return '<button type="button" class="cd-altlink" data-act="addalt" data-alt="' + esc(n) + '"' + (used ? ' disabled' : '') + '>' + esc(n) + (used ? ' ✓' : '') + '</button>';
    }).join('<span class="cd-sep"> · </span>') + '</div>';
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
  function parseAdds(text) {
    return linesOf(text).map(function (l) { var k = l.lastIndexOf('|'); return k < 0 ? { desc: l.trim(), amount: 0 } : { desc: l.slice(0, k).trim(), amount: num(l.slice(k + 1)) }; });
  }
  /* checks a PM's own Attachment B against the prime contract terms the desk would otherwise write */
  function checkOwnB(row) {
    var T = pcTerms(), B = row.ownB, out = [];
    var bad = function (t) { out.push({ bad: true, t: 'Your Attachment B ' + t }); }, warn = function (t) { out.push({ t: 'Your Attachment B ' + t }); };
    if (!B) { out.push({ bad: true, t: 'Upload your Attachment B, or switch it to "Build it here".' }); return out; }
    if (B.reading) { out.push({ t: 'Reading your Attachment B…' }); return out; }
    if (B.error) { bad('could not be read: ' + B.error); return out; }
    if (B.scanned) { warn('is a scanned PDF with no readable text, so the desk could not check it. Check retainage, billing, tax and exclusions by hand.'); return out; }
    var t = String(B.text).replace(/[\u2010-\u2015]/g, '-').replace(/[\u2018\u2019]/g, "'").replace(/[ \t]+/g, ' ');
    if (/<<[^>]*>>/.test(t)) bad('still has << >> placeholders to fill in.');
    if (B.highlight) warn('still has yellow highlight. Check every highlighted part was edited for this job.');
    // retainage
    var m = t.match(/(\d+(?:\.\d+)?)\s*%\s*Retainage will be withheld/i), has50 = /50\s*% complete|50\s*% of the Subcontract/i.test(t);
    if (T.retPct) {
      if (!m) warn('has no retainage paragraph (the prime contract retainage is ' + T.retPct + '% ' + T.retDuration + ').');
      else if (Number(m[1]) !== Number(T.retPct)) bad('withholds ' + m[1] + '% retainage; the prime contract says ' + T.retPct + '%.');
      if (m && /50/.test(T.retDuration) && !has50) warn('does not reduce retainage at 50% (prime contract: ' + T.retDuration + ').');
      if (m && !/50/.test(T.retDuration) && has50) bad('reduces retainage at 50%, but the prime contract retainage is ' + T.retDuration + '.');
    } else if (m && Number(m[1]) > 0) bad('withholds ' + m[1] + '% retainage; the prime contract has none.');
    // billing
    var bm = t.match(/no later than the (\d{1,2})\s*(st|nd|rd|th)/i), want = parseInt(T.billing, 10);
    if (!/Monthly Progress Pay Requests/i.test(t)) warn('has no billing paragraph (Monthly Progress Pay Requests…).');
    else if (bm && want && Number(bm[1]) !== want) bad('sets billing by the ' + bm[1] + bm[2] + '; this job\'s Monthly Billing Date is the ' + T.billing + '.');
    // sales tax
    var taxItem = /does not include Sales Taxes/i.test(t) || /Sales and Use Tax Exemption/i.test(t), taxEx = /(^|\n)\s*([-•\d.)]*\s*)?Sales taxes\./m.test(t);
    if (!T.taxExempt && taxItem) bad('uses the owner\'s sales tax exemption, but this job is not tax exempt.');
    if (!T.taxExempt && taxEx) bad('excludes sales taxes, but this job is not tax exempt: the price must include sales tax on materials.');
    if (T.taxExempt && !taxItem && !/tax exempt/i.test(t)) warn('has no sales tax exemption item; this job is tax exempt.');
    // bonds
    var bondsEx = /(^|\n)\s*([-•\d.)]*\s*)?Bonds\./m.test(t);
    if (row.bond === 'Yes' && bondsEx) bad('excludes Bonds, but this commitment requires a P&P bond.');
    if (row.bond !== 'Yes' && !bondsEx) warn('does not exclude Bonds (no bond is required on this commitment).');
    // platform
    var proc = /Procore will be utilized/i.test(t);
    if (T.ownerPlatform && proc) warn('says Procore will be used; the owner requires its own platform.');
    if (!T.ownerPlatform && !proc) warn('has no Procore item (General Item 21).');
    // stored materials
    if (!/stored off(-| )?(the )?site/i.test(t)) warn('has no terms for materials stored off site.');
    // accepted alternates
    T.alternates.filter(function (a) { return /^accepted/i.test(a.status); }).forEach(function (a) {
      var k = a.label.match(/^(?:Alternate|Alt\.?)\s*#?\s*(\d+)/i);
      if (k && !new RegExp('Alt(?:ernate)?\\.?\\s*#?\\s*' + k[1] + '\\b', 'i').test(t)) warn('does not mention accepted Alternate ' + k[1] + '. Ignore if it is not in this trade.');
    });
    // vendor
    var v = S.ref && S.ref.vendors.find(function (x) { return String(x.id) === String(row.vendorId); });
    if (v && normName(t).indexOf(normName(v.name).slice(0, 10)) < 0) warn('does not name ' + v.name + '.');
    return out;
  }
  async function readOwnB(row, file) {
    if (!file) return;
    row.ownB = { name: file.name, reading: true }; drawRows();
    try {
      var bytes = new Uint8Array(await file.arrayBuffer()), info, pdf = /\.pdf$/i.test(file.name);
      if (/\.docx$/i.test(file.name)) info = await BHDocs.docxText(bytes);
      else if (pdf) { var pages = await pdfPages(bytes.slice(0)); info = { text: pages.join('\n'), scanned: pages.join('').trim().length < 200 }; }
      else throw new Error('use a Word (.docx) or PDF file');
      row.ownB = Object.assign({ name: file.name, bytes: bytes, type: pdf ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, info);
    } catch (e) { row.ownB = { name: file.name, error: e.message }; }
    drawRows();
  }
  /* PM sign-off: warnings don't block, but the PM has to accept them; a change to the warnings clears the sign-off */
  function warnKey(iss) { return iss.filter(function (x) { return !x.bad; }).map(function (x) { return x.t; }).join('|'); }
  function acked(row, iss) { var k = warnKey(iss || issues(row)); return !k || row.ackKey === k; }
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
    if (row.type === 'LF' && !row.bMode) bad('Attachment B: choose "Build it here" or "Upload my own".');
    else if (row.type === 'LF' && row.bMode === 'own') checkOwnB(row).forEach(function (x) { out.push(x); });
    else if (row.type === 'LF') {
      if (!row.scopeName.trim()) bad('Add the scope name for Attachment B.');
      if (!linesOf(row.inclusions).length) warn('No inclusions entered; Attachment B will have an empty scope list.');
      if (!row.submittalDate.trim()) bad('Add the submittal deadline for Attachment B.');
      var adds = parseAdds(row.adds);
      adds.forEach(function (a, k) { if (!a.amount) bad('Cost table line ' + (k + 1) + ': add the amount after a | (e.g. "Alternate 1 – Wood Panels | 142000").'); });
      var addSum = adds.reduce(function (s, a) { return s + a.amount; }, 0);
      if (adds.length && total - addSum <= 0) bad('Cost table: the alternates and adds (' + money(addSum) + ') are more than the commitment total (' + money(total) + '); the lines should include them.');
      if (!row.proposalDate.trim()) warn('No proposal date; the cost table will say "Base Bid" without a date.');
      // edits that move away from the prime contract or the BH standard
      var pcItems = T.items.map(function (x) { return x.trim().toLowerCase(); }), mine = linesOf(row.items).map(function (x) { return x.trim().toLowerCase(); });
      var dropped = T.items.filter(function (x, k) { return mine.indexOf(pcItems[k]) < 0; });
      if (String(row.items || '').trim() && dropped.length) warn('Section B leaves out ' + dropped.length + ' prime contract item(s): ' + dropped.map(function (x) { return '"' + x.slice(0, 50) + (x.length > 50 ? '…' : '') + '"'; }).join('; ') + '.');
      if (row.equipAvail) warn('General Item 33 changed: BH equipment will be available to this trade partner for loading/unloading.');
      var exl = linesOf(row.exclusions).join(' \n ');
      if (/\bbond/i.test(exl) && row.bond === 'Yes') bad('Exclusions list bonds, but this commitment requires a P&P bond.');
      if (/sales\s*tax/i.test(exl) && !T.taxExempt) bad('Exclusions list sales tax, but this job is not tax exempt: the price must include sales tax on materials.');
      if (/retainage|retention/i.test(exl)) warn('Exclusions mention retainage; retainage terms come from the prime contract (' + T.retPct + '% ' + T.retDuration + ').');
      if (/tax[- ]exempt|tax exemption/i.test(linesOf(row.inclusions).join(' ')) && !T.taxExempt) bad('Inclusions mention a tax exemption, but this job is not tax exempt.');
    }
    if (row.rider && row.rider.matches && row.rider.matches.length > 1 && !row.riderPick) bad('Several approved riders match this vendor; pick one.');
    if (!row.number) bad('No free commitment number left in the ' + TYPES[row.type].series + ' series.');
    return out;
  }
  function bindRows() {
    var host = $('#cm-rows');
    host.addEventListener('change', function (e) {
      if (e.target.hasAttribute('data-ack-all')) {
        S.rows.forEach(function (r) { if (r.done) return; var iss = issues(r); if (!iss.some(function (x) { return x.bad; })) r.ackKey = e.target.checked ? warnKey(iss) : ''; });
        drawRows(); return;
      }
      var rEl = e.target.closest('.cd-row'); if (!rEl) return;
      var row = S.rows[+rEl.dataset.i], f = e.target.dataset.f, l = e.target.dataset.l;
      if (e.target.hasAttribute('data-ack')) { row.ackKey = e.target.checked ? warnKey(issues(row)) : ''; drawRows(); return; }
      if (e.target.hasAttribute('data-own-b')) { readOwnB(row, e.target.files && e.target.files[0]); return; }
      if (f) { row[f] = e.target.type === 'checkbox' ? e.target.checked : e.target.value; if (f === 'vendorId' || f === 'type') { row.rider = null; row.riderPick = ''; if (row.vendorId) findRider(row); } if (f === 'vendorId') row.signatoryId = pickSignatory(row.vendorId, ''); }
      if (l) { var j = +e.target.closest('tr').dataset.j; row.lines[j][l] = e.target.value; }
      drawRows();
    });
    host.addEventListener('click', function (e) {
      var b = e.target.closest('[data-act]'); if (!b) return;
      var rEl = b.closest('.cd-row'), row = S.rows[+rEl.dataset.i];
      if (b.dataset.act === 'toggle') row.open = !row.open;
      if (b.dataset.act === 'del') S.rows.splice(+rEl.dataset.i, 1);
      if (b.dataset.act === 'addalt') row.adds = (String(row.adds || '').trim() ? String(row.adds).replace(/\s+$/, '') + '\n' : '') + b.dataset.alt + ' | ';
      if (b.dataset.act === 'al') row.lines.push({ costCodeId: '', typeId: '', desc: '', amount: '' });
      if (b.dataset.act === 'dl') { row.lines.splice(+b.closest('tr').dataset.j, 1); if (!row.lines.length) row.lines.push({ costCodeId: '', typeId: '', desc: '', amount: '' }); }
      drawRows();
    });
    $('#cm-add').onclick = function () { S.rows.push(newRow()); drawRows(); };
    $('#cm-paste').onclick = function () { $('#cm-xlsx').click(); };
    $('#cm-xlsx').onchange = function (e) { var f = e.target.files && e.target.files[0]; e.target.value = ''; if (f) importWorkbook(f); };
    $('#cm-bfiles').onclick = function () { $('#cm-bin').click(); };
    $('#cm-bin').onchange = function (e) { var fs = Array.prototype.slice.call(e.target.files || []); e.target.value = ''; addOwnBFiles(fs); };
    $('#cm-create').onclick = createReady;
  }
  /* ---- bulk workbook (Commitments + Lines sheets) ---- */
  function typeOfLabel(s) { s = String(s).toLowerCase(); return /long/.test(s) ? 'LF' : /short/.test(s) ? 'SF' : /psa|professional/.test(s) ? 'PSA' : /off/.test(s) ? 'POOFF' : /po|purchase/.test(s) ? 'POON' : 'LF'; }
  function xlDate(v, weekday) {   // Excel stores a typed date as a day number
    if (!/^\d+(\.\d+)?$/.test(v) || Number(v) < 20000) return v;
    var d = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(v)) * 864e5);
    return d.toLocaleDateString('en-US', Object.assign({ timeZone: 'UTC', month: 'long', day: 'numeric', year: 'numeric' }, weekday ? { weekday: 'long' } : {}));
  }
  function pickSignatory(vendorId, name) {
    if (!vendorId || !S.ref) return '';
    var uname = function (u) { return u.name || ((u.first_name || '') + ' ' + (u.last_name || '')); };
    var list = S.ref.users.filter(function (u) { return u.vendor && String(u.vendor.id) === String(vendorId); }), nm = normName(name);
    var hit = nm && (list.find(function (u) { return normName(uname(u)) === nm; }) || list.find(function (u) { return normName(uname(u)).indexOf(nm) >= 0 || String(u.email_address || '').toLowerCase() === String(name).toLowerCase(); }));
    if (hit) return String(hit.id);
    return list.length === 1 ? String(list[0].id) : '';
  }
  async function importWorkbook(file) {
    var R = S.ref, T = pcTerms();
    msg('Reading ' + file.name + '…');
    try {
      var wb = await BHDocs.readXlsx(new Uint8Array(await file.arrayBuffer()));
      var sheet = function (n) { var k = Object.keys(wb).find(function (x) { return x.trim().toLowerCase() === n; }); return k ? wb[k] : null; };
      var C = sheet('commitments'), L = sheet('lines');
      if (!C || !L) throw new Error('it needs a Commitments sheet and a Lines sheet. Download the bulk template from this page.');
      var hdr = function (rows) { var h = (rows[0] || []).map(function (x) { return String(x || '').trim().toLowerCase(); }); return function (row, name) { var i = h.indexOf(name); return i < 0 || !row ? '' : String(row[i] == null ? '' : row[i]).trim(); }; };
      var gc = hdr(C), gl = hdr(L), add = [], unknown = [], probs = [], lines = {};
      L.slice(1).forEach(function (r) { var ref = gl(r, 'ref'); if (ref) (lines[ref] = lines[ref] || []).push(r); });
      C.slice(1).forEach(function (r) {
        if (!gc(r, 'vendor') && !gc(r, 'title')) return;
        var ref = gc(r, 'ref'), vn = gc(r, 'vendor');
        var v = R.vendors.find(function (x) { return normName(x.name) === normName(vn); }) || R.vendors.find(function (x) { return normName(vn) && normName(x.name).indexOf(normName(vn)) === 0; });
        if (!v) unknown.push(vn || '(blank vendor)');
        var type = typeOfLabel(gc(r, 'type')), b = gc(r, 'attachment b').toLowerCase(), extra = gc(r, 'extra section b items');
        var row = newRow({ open: false, ref: ref, type: type, vendorId: v ? String(v.id) : '', title: gc(r, 'title'), bond: /^y/i.test(gc(r, 'p&p bond required')) ? 'Yes' : 'No',
          bMode: type === 'LF' ? (/^own/.test(b) ? 'own' : /^build/.test(b) ? 'build' : '') : '', ownBName: gc(r, 'own b file name'),
          scopeName: gc(r, 'scope name'), submittalDate: xlDate(gc(r, 'submittal deadline'), true), csi: gc(r, 'spec sections'), proposalDate: xlDate(gc(r, 'proposal date')),
          adds: gc(r, 'alternates and adds'), items: (T ? T.items.join('\n') : '') + (extra ? '\n' + extra : ''), inclusions: gc(r, 'scope inclusions'), exclusions: gc(r, 'accepted exclusions'),
          equipAvail: /^y/i.test(gc(r, 'bh equipment available')), description: gc(r, 'description'), lines: [] });
        row.signatoryId = pickSignatory(row.vendorId, gc(r, 'vendor signatory'));
        (lines[ref] || []).forEach(function (l) {
          var code = gl(l, 'cost code').replace(/\s/g, ''), tc = gl(l, 'cost type').toUpperCase().charAt(0);
          var cc = R.costCodes.find(function (x) { return String(x.full_code || x.code || '').replace(/\s/g, '') === code; });
          var tt = R.types.find(function (x) { return typeCode(x) === tc; });
          row.lines.push({ costCodeId: cc ? String(cc.id) : '', typeId: tt ? String(tt.id) : '', desc: gl(l, 'description') || row.title, amount: num(gl(l, 'amount')) });
        });
        if (!row.lines.length) { probs.push('Ref ' + (ref || '?') + ' (' + row.title + ') has no lines on the Lines sheet'); row.lines.push({ costCodeId: '', typeId: '', desc: '', amount: '' }); }
        add.push(row);
      });
      if (!add.length) throw new Error('no commitments on the Commitments sheet.');
      if (S.rows.length === 1 && !S.rows[0].vendorId && !S.rows[0].title) S.rows = [];
      S.rows = S.rows.concat(add);
      add.forEach(function (r) { if (r.vendorId) findRider(r); });
      drawRows();
      var note = add.length + ' commitment(s) loaded from ' + file.name + '. Open a row with Details to check or change it.', own = add.filter(function (r) { return r.bMode === 'own'; }).length;
      if (unknown.length) note += ' Not in the project directory, pick these vendors on their rows: ' + unknown.join(', ') + '.';
      if (probs.length) note += ' ' + probs.join('; ') + '.';
      if (own) note += ' ' + own + ' Long Form(s) use their own Attachment B: click Add Attachment B files.';
      msg(note, !!(unknown.length || probs.length));
    } catch (e) { msg('Could not read the workbook: ' + e.message, true); }
  }
  function addOwnBFiles(files) {
    var want = S.rows.filter(function (r) { return !r.done && r.type === 'LF' && r.bMode === 'own'; }), unmatched = [];
    if (!want.length) { msg('No Long Form is set to use its own Attachment B. Set Attachment B to "Upload my own" first.', true); return; }
    files.forEach(function (f) {
      var lower = f.name.toLowerCase(), key = normName(f.name.replace(/\.[^.]+$/, ''));
      var row = want.find(function (r) { return r.ownBName && r.ownBName.trim().toLowerCase() === lower; }) ||
        want.find(function (r) { var v = S.ref.vendors.find(function (x) { return String(x.id) === String(r.vendorId); }); return !r.ownBName && v && key.indexOf(normName(v.name).slice(0, 10)) >= 0; });
      if (row) readOwnB(row, f); else unmatched.push(f.name);
    });
    msg(unmatched.length ? 'Not matched to a Long Form set to its own B: ' + unmatched.join(', ') + '. Check the vendor name, or the Own B file name column.' : files.length + ' Attachment B file(s) matched and checked.', !!unmatched.length);
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
    var ready = S.rows.filter(function (r) { var iss = issues(r); return !r.done && !iss.some(function (x) { return x.bad; }) && acked(r, iss); });
    if (!ready.length) { msg('No rows are validated. Fix the items marked in red, and sign off the warnings, first.', true); return; }
    var withWarn = ready.filter(function (r) { return warnKey(issues(r)); }).map(function (r) { return r.number; });
    if (!confirm('Create ' + ready.length + ' Draft commitment(s) in Procore on ' + (S.project.project_number || S.project.name) + '?' + (withWarn.length ? '\n\n' + withWarn.join(', ') + ' differ from the prime contract or the BH standard, and you signed off the warnings. Continue?' : ''))) return;
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
    var ret = (row.type === 'PSA' || /^PO/.test(row.type)) ? 0 : T.retPct;   // PSAs and POs carry no retainage
    var cfs = {}, isPO = ep === 'purchase_order_contracts';
    var hasRider = row.riderPick && row.riderPick !== 'none';
    if (row.bhSignatoryId) cfs['custom_field_' + CF.bhSignatory] = Number(row.bhSignatoryId);
    if (row.signatoryId) cfs['custom_field_' + CF.subSignatory] = Number(row.signatoryId);
    if (lovId(CF.riderNeeded, hasRider ? 'Yes, attached' : 'No')) cfs['custom_field_' + CF.riderNeeded] = lovId(CF.riderNeeded, hasRider ? 'Yes, attached' : 'No');
    if (isPO) {
      // the PO fieldset has its own PDF Type and a Vendor Type; no billing date, bond or retainage duration
      if (lovId(CF.poPdfType, 'Purchase Order')) cfs['custom_field_' + CF.poPdfType] = lovId(CF.poPdfType, 'Purchase Order');
      var vt = row.type === 'POOFF' ? 'Third-Party Delivery' : 'On-Site';
      if (lovId(CF.vendorType, vt)) cfs['custom_field_' + CF.vendorType] = lovId(CF.vendorType, vt);
    } else {
      if (TY.pdf && lovId(CF.pdfType, TY.pdf)) cfs['custom_field_' + CF.pdfType] = lovId(CF.pdfType, TY.pdf);
      cfs['custom_field_' + CF.billing] = T.billing;
      if (lovId(CF.bond, row.bond)) cfs['custom_field_' + CF.bond] = lovId(CF.bond, row.bond);
      var durLabel = row.type === 'PSA' ? 'N/A' : T.retDuration;
      if (lovId(CF.retDuration, durLabel)) cfs['custom_field_' + CF.retDuration] = lovId(CF.retDuration, durLabel);
    }
    var need = [[CF.riderNeeded, 'Rider Needed?']];
    if (isPO) { need.push([CF.poPdfType, 'Purchase Order PDF Type']); need.push([CF.vendorType, 'Vendor Type']); }
    else { need.push([CF.retDuration, 'Retainage Duration']); if (TY.pdf) need.push([CF.pdfType, 'Subcontract PDF Type']); }
    var missing = need.filter(function (n) { return !cfs['custom_field_' + n[0]]; }).map(function (n) { return n[1]; });
    if (!row.bhSignatoryId || !row.signatoryId) missing.push('signatories');
    if (missing.length) throw new Error('Not created: the desk could not set ' + missing.join(', ') + ' (option lists did not load). Reload the page and try again.');
    // validator of record
    var today = new Date().toISOString().slice(0, 10), accepted = issues(row).filter(function (x) { return !x.bad; }).map(function (x) { return x.t; });
    var vNote = 'Attachment B: ' + (row.type !== 'LF' ? 'n/a' : row.bMode === 'own' ? 'PM upload (' + (row.ownB && row.ownB.name) + '), checked' : 'built on the desk') + '. ' + (accepted.length ? 'Warnings accepted by the PM: ' + accepted.join(' / ') : 'No warnings.');
    var vcf = {};   // written in a second call, so a fieldset without these fields can't stop the commitment's required fields saving
    if (CF.validatedBy && S.me) vcf['custom_field_' + CF.validatedBy] = Number(S.me.id);
    if (CF.validatedDate) vcf['custom_field_' + CF.validatedDate] = new Date().toISOString();
    if (CF.validationNotes) vcf['custom_field_' + CF.validationNotes] = vNote;
    log(row.number + ': validated by ' + (S.me ? S.me.name || S.me.login : '?') + ' on ' + today + '. ' + vNote);
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
    if (Object.keys(vcf).length) {
      try { await pc('PATCH', '/rest/v2.0/companies/' + CID + '/projects/' + pid + '/commitment_contracts/' + made.id, Object.assign({}, v2, vcf)); }
      catch (e) { log(row.number + ': could not record the validator in Procore (' + e.message + '). Check PC Verified By / Date are in this commitment\'s fieldset.'); vcf = {}; }
    }
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
    var miss = Object.keys(Object.assign({}, cfs, vcf)).filter(function (k) { var v = back.custom_fields && back.custom_fields[k]; v = v && v.value !== undefined ? v.value : v; return v == null || v === ''; });
    var nAtt = (back.attachments || []).length;
    log(row.number + ': checked in Procore — ' + nAtt + ' attachment(s), retainage ' + back.retainage_percent + '%' + (miss.length ? ', custom fields NOT saved: ' + miss.join(', ') : ', custom fields saved') + '.');
    row.done = 'id ' + made.id;
  }
  async function buildAttachments(row, T, vendor) {
    var TY = TYPES[row.type], out = [], pn = (S.project.project_number ? S.project.project_number + ' ' : '') + S.project.name;
    var city = [S.project.city, S.project.state_code].filter(Boolean).join(', ');
    var d = { projectNumName: pn, projectCity: city, projectNumber: S.project.project_number || '', commitmentNumber: row.number, subName: vendor ? vendor.name : '',
      ownerAI: T.ownerAI, taxExempt: T.taxExempt, ownerPlatform: T.ownerPlatform, bondRequired: row.bond === 'Yes', builderRiskOwner: T.builderRiskOwner, projectItems: String(row.items || '').trim() ? linesOf(row.items) : T.items, equipAvailable: !!row.equipAvail, alternates: T.alternates,
      scopeLines: row.lines.map(function (l) { return { desc: l.desc || row.title, amount: num(l.amount) }; }), submittalDate: row.submittalDate.trim(), scopeName: row.scopeName.trim(), proposalDate: (row.proposalDate || '').trim(), adds: parseAdds(row.adds),
      csi: linesOf(row.csi), inclusions: linesOf(row.inclusions), exclusions: linesOf(row.exclusions), submitDay: '10th', billingDate: T.billing, retPct: T.retPct, retDuration: T.retDuration };
    var DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    for (var i = 0; i < TY.att.length; i++) {
      var a = TY.att[i];
      try {
        if (a === 'G') {
          var g = 'Attachment G - Subs Pay App - ' + (T.retPct || 5) + (/50/.test(T.retDuration) ? ' thru 50' : '') + '.xlsx';
          out.push({ name: 'Attachment G - Subcontractors Application for Payment.xlsx', type: XLSX, bytes: await BHDocs.fillXlsx(await opsFile(TEMPLATE_DIR + '/' + g), d) });
        } else if (a === 'B' && row.bMode === 'own') {
          out.push({ name: 'Attachment B - Scope of Work (' + (d.subName || row.number) + ')' + (/\.pdf$/i.test(row.ownB.name) ? '.pdf' : '.docx'), type: row.ownB.type, bytes: row.ownB.bytes });
          log(row.number + ': Attachment B — your file ' + row.ownB.name + ' attached as uploaded.');
        } else if (a === 'B') {
          var b = await BHDocs.buildAttachmentB(await opsFile(TEMPLATE_DIR + '/' + ATT_FILES.B), d);
          out.push({ name: 'Attachment B - Scope of Work (' + (d.subName || row.number) + ').docx', type: DOCX, bytes: b.bytes });
          log(row.number + ': Attachment B — ' + b.log.join('; '));
        } else {
          out.push({ name: ATT_NAMES[a], type: DOCX, bytes: await BHDocs.fillDocx(await opsFile(TEMPLATE_DIR + '/' + ATT_FILES[a]), d) });
        }
      } catch (e) { log(row.number + ': could not build attachment ' + a + ' — ' + e.message); }
    }
    if (row.type === 'LF') out = (await jobAttachmentA(row)).concat(out);
    if (row.riderPick && row.riderPick !== 'none') {
      try { out.push({ name: row.riderPick, type: /\.pdf$/i.test(row.riderPick) ? 'application/pdf' : DOCX, bytes: await opsFile(RIDER_DIR + '/' + row.rider.folder + '/' + row.riderPick) }); }
      catch (e) { log(row.number + ': could not fetch the rider — ' + e.message); }
    }
    return out;
  }

  /* Attachment A (drawing and spec logs): job-wide, so copy it from the newest subcontract on the job that has one */
  var ATT_A = {};
  async function procoreFile(url) {
    var tk = await tokens();
    var r = await fetch(API + '/procore/file?url=' + encodeURIComponent(url), { method: 'POST', headers: { Authorization: 'Bearer ' + tk.idToken } });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return new Uint8Array(await r.arrayBuffer());
  }
  async function jobAttachmentA(row) {
    var pid = S.project.id;
    if (!ATT_A[pid]) ATT_A[pid] = (async function () {
      var list = (await pcAll('/rest/v1.0/work_order_contracts?project_id=' + pid)).sort(function (a, b) { return b.id - a.id; });
      for (var k = 0; k < list.length && k < 20; k++) {
        var c = await pc('GET', '/rest/v1.0/work_order_contracts/' + list[k].id + '?project_id=' + pid);
        var A = (c.attachments || []).filter(function (a) { return /^Attachment A\b/i.test(a.name || a.filename || ''); });
        if (A.length) return { from: c.number, files: A.map(function (a) { return { name: a.name || a.filename, url: a.url }; }) };
      }
      return { files: [] };
    })();
    var src = await ATT_A[pid], out = [];
    if (!src.files.length) { log(row.number + ': no Attachment A on this job\'s other subcontracts yet. Export the drawing and spec logs from Procore and attach them by hand.'); return out; }
    for (var i = 0; i < src.files.length; i++) {
      var f = src.files[i];
      try { out.push({ name: f.name, type: /\.pdf$/i.test(f.name) ? 'application/pdf' : 'application/octet-stream', bytes: await procoreFile(f.url) }); }
      catch (e) { log(row.number + ': could not copy ' + f.name + ' — ' + e.message); }
    }
    log(row.number + ': Attachment A copied from ' + src.from + ' (' + out.map(function (x) { return x.name; }).join(', ') + '). Check it is the latest log.');
    return out;
  }

  /* ---------------- navigation ---------------- */
  function bindNav() {
    var show = function (v) {
      $$('.cd-view').forEach(function (s) { s.hidden = s.id !== 'v-' + v; });
      $$('[data-view]').forEach(function (a) { if (a.dataset.view === v) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current'); });
      if (v === 'commitments' && S.project) renderRows();
      if (window.BHHub) window.BHHub.show(v);
    };
    $$('[data-view]').forEach(function (a) { a.addEventListener('click', function (e) { e.preventDefault(); history.replaceState(null, '', '#' + a.dataset.view); show(a.dataset.view); }); });
    var h = (location.hash.match(/^#(prime|commitments|pay|email)\b/) || [])[1];
    show(h || 'prime');
    $('#pc-read').onclick = readContract;
    $('#pc-save').onclick = savePrime;
    $('#pc-files').onchange = async function (e) {
      var fs = Array.prototype.slice.call(e.target.files || []);
      for (var i = 0; i < fs.length; i++) S.docs.push({ name: fs[i].name, bytes: new Uint8Array(await fs[i].arrayBuffer()), source: 'local' });
      renderDocs(); $('#pc-read').disabled = !S.pcRec; e.target.value = '';
    };
    bindRows();
  }

  /* shared with hub.js (Pay apps and Reminder email tabs) */
  window.BHDesk = { S: S, CF: CF, pc: pc, pcAll: pcAll, tokens: tokens, G: G, msg: msg, esc: esc, money: money, num: num, htmlToText: htmlToText, linesOf: linesOf,
    cfVal: cfVal, lovLabel: lovLabel, pcTerms: pcTerms, loadRef: loadRef, account: function () { return acct; } };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
