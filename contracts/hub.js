/* Project Hub (dashboard.baileyharris.com/contracts) — 24 Nov 2026.
   The Buyout list (with cost codes), Pay apps and Reminder email tabs. Uses window.BHDesk from contracts.js
   (sign-in, Procore through the bh-contracts Worker, the job and its verified prime contract).
   Pay apps come from the "AP Pay Applications" list on /sites/APDesk (the portal and payables@ intake write
   it). PM approval happens here; approving emails payables@ with the Attachment G, backup and approval
   history, and Drew enters it in ComputerEase as he does today. Nothing is written to ComputerEase.
   The reminder template and each job's settings are JSON files on the APDesk site's Documents library,
   under hub/, for send_billing_reminders.js to read. */
(function () {
  'use strict';
  var GRAPH = 'https://graph.microsoft.com/v1.0', SP_HOST = 'baileyharris.sharepoint.com', AP_SITE = '/sites/APDesk', AP_LIST = 'AP Pay Applications';
  var PAYABLES = 'payables@baileyharris.com', TPL_PATH = 'hub/reminder_template.json', JOB_DIR = 'hub/jobs/';
  var APPROVED = { 'ready for entry': 1, 'PM approved': 1, 'approved as noted': 1, 'approved and entered': 1 };
  var MAIL_ATTACH_MAX = 3 * 1024 * 1024;
  /* 24 Nov 2026: the portal makes fresh links for return emails; people files hold each PM's email setting */
  var PORTAL = 'https://pay.baileyharris.com', PEOPLE_DIR = 'hub/people/';
  /* same as the launcher's Pay App Queue (BH Dashboards\site\index.html PAYAPP_ALL / PAYAPP_DONE / PM_ROLE) - keep in step */
  var PAYAPP_ALL = ['tiffany.blackmon@baileyharris.com', 'bsims@baileyharris.com'];
  var PAYAPP_DONE = { 'ready for entry': 1, 'PM approved': 1, 'approved as noted': 1, 'approved and entered': 1, 'revise and resubmit': 1, 'duplicate archived': 1, 'email only': 1, 'handled outside desk': 1, 'no billing this period': 1 };
  var PM_ROLE = /^(project manager|project management team)$/i, RES_SITE = '/sites/Resources';
  /* one subject format per kind, so payables@ can sort them with an Outlook rule ("subject contains Pay App Approved") */
  var subjectFor = function (kind, f, vendor) { return kind + ' - ' + [f.JobNumber, vendor || f.Vendor, f.Subcontract].filter(Boolean).join(' - ') + (f.PayRequest ? ' - Pay App #' + f.PayRequest : ''); };
  var D = function () { return window.BHDesk; };
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var esc = function (s) { return D().esc(s); }, money = function (v) { return D().money(v); };
  var r2 = function (v) { return Math.round(Number(v || 0) * 100) / 100; };
  var H = { view: (location.hash.match(/^#(prime|commitments|pay|email)\b/) || [])[1] || 'prime', q: null, rows: null, period: null, sel: null, tpl: null, job: null, lines: {}, open: {}, busy: false };

  /* ---------------- Graph ---------------- */
  async function gcall(method, path, body, headers) {
    var t = await D().tokens();
    var h = Object.assign({ Authorization: 'Bearer ' + t.accessToken }, headers || {});
    if (body != null && !(body instanceof Uint8Array) && typeof body !== 'string') { h['Content-Type'] = 'application/json'; body = JSON.stringify(body); }
    var r = await fetch(path.indexOf('http') === 0 ? path : GRAPH + path, { method: method, headers: h, body: body });
    var txt = r.status === 204 ? '' : await r.text(), j = null; try { j = txt ? JSON.parse(txt) : null; } catch (_) { j = txt; }
    return { ok: r.ok, status: r.status, body: j };
  }
  async function gbytes(path) {
    var t = await D().tokens();
    var r = await fetch(GRAPH + path, { headers: { Authorization: 'Bearer ' + t.accessToken } });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return new Uint8Array(await r.arrayBuffer());
  }
  var encPath = function (p) { return p.split('/').map(encodeURIComponent).join('/'); };
  async function queue() {
    if (H.q) return H.q;
    var s = await gcall('GET', '/sites/' + SP_HOST + ':' + AP_SITE + '?$select=id');
    if (!s.ok) throw new Error('Could not reach the AP Desk SharePoint site (' + s.status + ').');
    var l = await gcall('GET', '/sites/' + s.body.id + '/lists/' + encodeURIComponent(AP_LIST) + '?$select=id');
    if (!l.ok) throw new Error('Could not open the "' + AP_LIST + '" list (' + l.status + ').');
    H.q = { sid: s.body.id, lid: l.body.id }; return H.q;
  }
  async function readJson(path, dflt) {
    var Q = await queue(), r = await gcall('GET', '/sites/' + Q.sid + '/drive/root:/' + encPath(path) + ':/content');
    if (r.status === 404) return dflt;
    if (!r.ok) throw new Error('Could not read ' + path + ' (' + r.status + ').');
    return typeof r.body === 'object' && r.body ? r.body : dflt;
  }
  async function writeJson(path, obj) {
    var Q = await queue(), r = await gcall('PUT', '/sites/' + Q.sid + '/drive/root:/' + encPath(path) + ':/content', JSON.stringify(obj, null, 2), { 'Content-Type': 'application/json' });
    if (!r.ok) throw new Error('Could not save ' + path + ' (' + r.status + (r.body && r.body.error ? ': ' + r.body.error.message : '') + ').');
  }
  function b64(bytes) { var s = '', CH = 0x8000; for (var i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH)); return btoa(s); }
  async function sendMail(to, subject, html, attachments, cc) {
    var m = { subject: subject, body: { contentType: 'HTML', content: html }, toRecipients: to.map(function (a) { return { emailAddress: { address: a } }; }) };
    if (cc && cc.length) m.ccRecipients = cc.map(function (a) { return { emailAddress: { address: a } }; });
    if (attachments && attachments.length) m.attachments = attachments.map(function (a) { return { '@odata.type': '#microsoft.graph.fileAttachment', name: a.name, contentType: a.type || 'application/octet-stream', contentBytes: b64(a.bytes) }; });
    var r = await gcall('POST', '/me/sendMail', { message: m, saveToSentItems: true });
    if (!r.ok) throw new Error((r.body && r.body.error && r.body.error.message) || ('HTTP ' + r.status) + (r.status === 403 ? ' (Mail.Send permission not granted yet)' : ''));
  }
  function me() { var a = D().account() || {}; return { email: String(a.username || '').toLowerCase(), name: a.name || a.username || 'PM' }; }

  /* ---------------- dates: due date = the job's Monthly Billing Date, next business day ---------------- */
  var iso = function (d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
  var nth = function (y, m, dow, n) { var d = new Date(y, m, 1); while (d.getDay() !== dow) d.setDate(d.getDate() + 1); d.setDate(d.getDate() + 7 * (n - 1)); return d; };
  function holidays(y) {
    var obs = function (d) { if (d.getDay() === 6) d.setDate(d.getDate() - 1); else if (d.getDay() === 0) d.setDate(d.getDate() + 1); return iso(d); };
    var last = new Date(y, 5, 0); while (last.getDay() !== 1) last.setDate(last.getDate() - 1);
    var h = {}; [obs(new Date(y, 0, 1)), iso(last), obs(new Date(y, 6, 4)), iso(nth(y, 8, 1, 1)), iso(nth(y, 10, 4, 4)), obs(new Date(y, 11, 25))].forEach(function (k) { h[k] = 1; });
    return h;
  }
  function nextBusiness(d) { while (d.getDay() === 0 || d.getDay() === 6 || holidays(d.getFullYear())[iso(d)]) d.setDate(d.getDate() + 1); return d; }
  function billingDay() { var T = D().pcTerms(), n = parseInt(String((T && T.billing) || '20'), 10); return n >= 1 && n <= 28 ? n : 20; }
  function leadDays() { var n = Number(H.job && H.job.leadDays); return n >= 1 && n <= 14 ? n : 7; }
  function periodInfo(key) {
    var y = +key.slice(0, 4), m = +key.slice(5, 7), day = billingDay();
    var due = nextBusiness(new Date(y, m - 1, day)), end = new Date(y, m - 1, 0), rem = new Date(due); rem.setDate(rem.getDate() - leadDays());
    while (rem.getDay() === 0 || rem.getDay() === 6 || holidays(rem.getFullYear())[iso(rem)]) rem.setDate(rem.getDate() - 1);   // same rule as the reminder script
    return { key: key, due: due, end: end, rem: rem, moved: due.getDate() !== day, day: day };
  }
  var ord = function (n) { var s = ['th', 'st', 'nd', 'rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); };
  var pkey = function (y, m) { return y + '-' + String(m).padStart(2, '0'); };
  function periodKeys() {
    var now = new Date(), y = now.getFullYear(), m = now.getMonth() + 1, today = new Date(y, now.getMonth(), now.getDate());
    if (periodInfo(pkey(y, m)).due < today) { m++; if (m > 12) { m = 1; y++; } }
    var out = []; for (var i = 0; i < 6; i++) { out.push(pkey(y, m)); m--; if (m < 1) { m = 12; y--; } }
    return out;
  }
  var fmtD = function (d, long) { return d.toLocaleDateString('en-US', long ? { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' } : { weekday: 'short', month: 'short', day: 'numeric' }); };
  var monthName = function (key) { return new Date(+key.slice(0, 4), +key.slice(5, 7) - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' }); };
  function parseD(v) {
    if (!v) return null; var s = String(v), m;
    if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return new Date(+m[3], +m[1] - 1, +m[2]);
    return null;
  }

  /* ---------------- job summary and tab captions ---------------- */
  function summary() {
    var S = D().S, r = S.pcRec, el = $('#hub-sum');
    $('#hub-title').textContent = S.project ? S.project.name : 'Pick a job';
    if (!r || !S.project) { el.hidden = true; $('#hub-kick').textContent = S.project ? String(S.project.project_number || '') : 'PROJECT HUB'; return; }
    var T = D().pcTerms(), val = r.grand_total || r.revised_contract_amount || r.original_contract_amount, owner = r.vendor && (r.vendor.name || r.vendor.company);
    $('#hub-kick').textContent = [S.project.project_number, owner].filter(Boolean).join(' · ').toUpperCase();
    var sc = r.substantial_completion_date ? new Date(r.substantial_completion_date + 'T12:00:00').toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : '—';
    var P = periodInfo(periodKeys()[0]);
    el.innerHTML = [['Prime contract', val ? money(val).replace(/\.00$/, '') : '—'], ['Retainage', T.retPct + '% ' + T.retDuration], ['Pay apps due', T.billing + ' monthly'], ['Tax exempt', T.taxExempt ? 'Yes' : 'No'], ['Substantial completion', sc], ['Status', T.verified ? 'Verified' : 'Not verified']]
      .map(function (x, i) { return '<div' + (i === 5 && !T.verified ? ' class="is-bad"' : '') + '><dt>' + esc(x[0]) + '</dt><dd>' + esc(x[1]) + '</dd></div>'; }).join('');
    el.hidden = false;
    $('#t-prime').textContent = T.verified ? 'Prime contract verified' : 'Verify the prime contract';
    $('#t-pay').textContent = monthName(P.key).split(' ')[0] + ' · due ' + fmtD(P.due);
    $('#t-email').textContent = 'Sends ' + fmtD(P.rem);
  }

  /* =====================================================================
     BUYOUT — commitments in Procore, click a row for its cost codes
     ===================================================================== */
  function typeOf(c, ep) {
    if (ep === 'purchase_order_contracts') return 'PO';
    var n = String(c.number || ''), d = n.charAt(0);
    var m = n.match(/^(\d)\d\d/) || n.match(/(\d)\d\d$/); d = m ? m[1] : d;
    return d === '4' ? 'Long Form' : d === '9' ? 'Short Form' : d === '1' ? 'PSA' : 'Subcontract';
  }
  function buyCheck(c, type) {
    var T = D().pcTerms(), rp = c.retainage_percent == null || c.retainage_percent === '' ? null : Number(c.retainage_percent);
    if (rp == null) return { k: 'mut', t: '—' };
    if ((type === 'PO' || type === 'PSA') && rp > 0) return { k: 'bad', t: 'Retainage ' + rp + '%: ' + type + 's carry none' };
    if ((type === 'Long Form' || type === 'Short Form') && T && rp !== Number(T.retPct)) return { k: 'warn', t: 'Retainage ' + rp + '%; prime contract ' + T.retPct + '%' };
    return { k: 'ok', t: 'Matches' };
  }
  var pcUrl = function (ep, id) { return 'https://app.procore.com/' + encodeURIComponent(D().S.project.id) + '/project/commitments/' + ep + '/' + encodeURIComponent(id); };
  async function renderBuy() {
    var S = D().S, el = $('#buy-list');
    if (!S.project) { el.innerHTML = ''; return; }
    if (!S.pcRec) { el.innerHTML = '<p class="bh-caption">No prime contract loaded for this job.</p>'; return; }
    el.innerHTML = '<p class="bh-caption">Loading commitments from Procore…</p>';
    var R = await D().loadRef();
    var all = (R.woc || []).map(function (c) { return { c: c, ep: 'work_order_contracts' }; }).concat((R.poc || []).map(function (c) { return { c: c, ep: 'purchase_order_contracts' }; }));
    all.sort(function (a, b) { return String(a.c.number).localeCompare(String(b.c.number)); });
    var total = 0, exec = 0, issues = 0;
    var rows = all.map(function (x) {
      var c = x.c, type = typeOf(c, x.ep), ck = buyCheck(c, type), v = Number(c.grand_total || 0), key = x.ep + ':' + c.id, open = !!H.open[key];
      total += v; if (c.executed) exec++; if (ck.k === 'bad' || ck.k === 'warn') issues++;
      var st = c.status === 'Approved' && c.executed ? ['ok', 'Executed'] : c.status === 'Draft' ? ['mut', 'Draft'] : c.status === 'Approved' ? ['warn', 'Not executed'] : ['warn', c.status || '—'];
      var lines = H.lines[key];
      return '<tr class="hub-buyrow' + (open ? ' is-open' : '') + '" data-key="' + esc(key) + '" tabindex="0" aria-expanded="' + open + '">' +
        '<td><span class="hub-caret" aria-hidden="true">' + (open ? '▾' : '▸') + '</span><a class="hub-pclink" href="' + esc(pcUrl(x.ep, c.id)) + '" target="_blank" rel="noopener" title="Open ' + esc(c.number) + ' in Procore"><b>' + esc(c.number) + '</b> <span aria-hidden="true">↗</span></a></td><td>' + esc((c.vendor && (c.vendor.company || c.vendor.name)) || '') + '</td><td class="mut">' + esc(c.title || '') + '</td><td>' + esc(type) + '</td>' +
        '<td class="num">' + money(v) + '</td><td><span class="hub-pill ' + st[0] + '">' + esc(st[1]) + '</span></td><td class="hub-ck ' + ck.k + '">' + esc(ck.t) + '</td></tr>' +
        (open ? '<tr class="hub-lines"><td colspan="7"><a class="hub-pclink small" href="' + esc(pcUrl(x.ep, c.id)) + '" target="_blank" rel="noopener">Open ' + esc(c.number) + ' in Procore ↗</a>' + (lines === 'loading' ? '<span class="bh-caption">Loading cost codes…</span>' : lines && lines.error ? '<span class="hub-ck bad">' + esc(lines.error) + '</span>' : linesTable(lines || [])) + '</td></tr>' : '');
    });
    el.innerHTML = '<div class="hub-kpis small">' + [['Commitments', all.length], ['Committed', money(total).replace(/\.\d\d$/, '')], ['Executed', exec + ' of ' + all.length], ['Prime contract issues', issues]].map(function (k) { return '<div><span>' + esc(k[0]) + '</span><b>' + esc(String(k[1])) + '</b></div>'; }).join('') + '</div>' +
      '<div class="bh-table-wrap"><table class="bh-table hub-table"><thead><tr><th>Number</th><th>Trade partner</th><th>Scope</th><th>Type</th><th class="num">Value</th><th>Status</th><th>Prime contract check</th></tr></thead><tbody>' + rows.join('') + '</tbody></table></div>';
    $('#t-commitments').textContent = all.length + ' commitments · ' + (all.length - exec) + ' not executed';
  }
  function linesTable(L) {
    if (!L.length) return '<span class="bh-caption">No line items on this commitment.</span>';
    var tot = 0;
    return '<table class="hub-subtable"><thead><tr><th>Cost code</th><th>Cost type</th><th>Description</th><th class="num">Amount</th></tr></thead><tbody>' + L.map(function (l) {
      var cc = l.cost_code || {}, ty = l.line_item_type || {}, a = Number(l.amount != null ? l.amount : (l.extended_amount || 0)); tot += a;
      return '<tr><td>' + esc((cc.full_code || cc.code || '') + (cc.name ? ' · ' + cc.name : '')) + '</td><td>' + esc(ty.code ? ty.code + ' · ' + ty.name : ty.name || '') + '</td><td>' + esc(l.description || '') + '</td><td class="num">' + money(a) + '</td></tr>';
    }).join('') + '<tr class="tot"><td colspan="3">Total</td><td class="num">' + money(tot) + '</td></tr></tbody></table>';
  }
  async function toggleBuy(key) {
    H.open[key] = !H.open[key];
    if (H.open[key] && !H.lines[key]) {
      H.lines[key] = 'loading'; renderBuy();
      var p = key.split(':'), pid = D().S.project.id;
      try { H.lines[key] = await D().pcAll('/rest/v1.0/' + p[0] + '/' + p[1] + '/line_items?project_id=' + pid); }
      catch (e) { H.lines[key] = { error: 'Could not load the cost codes: ' + e.message + (e.status === 403 ? ' (the Contracts service needs its update deployed)' : '') }; }
    }
    renderBuy();
  }

  /* =====================================================================
     PAY APPS
     ===================================================================== */
  function subs() {
    var R = D().S.ref || {};
    return (R.woc || []).filter(function (c) { return c.status === 'Approved' && c.executed; }).sort(function (a, b) { return String(a.number).localeCompare(String(b.number)); });
  }
  async function loadRows() {
    var Q = await queue(), pn = String(D().S.project.project_number || '').trim(), out = [], pref = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
    var base = '/sites/' + Q.sid + '/lists/' + Q.lid + '/items?$expand=fields&$top=500';
    var url = base + '&$filter=' + encodeURIComponent("fields/JobNumber eq '" + pn.replace(/'/g, "''") + "'");
    var first = await gcall('GET', url, null, pref);
    if (!first.ok) { url = base; first = null; }
    for (var p = 0; url && p < 40; p++) {
      var r = first || await gcall('GET', url, null, pref); first = null;
      if (!r.ok) throw new Error('Could not read the pay app list (' + r.status + ').');
      (r.body.value || []).forEach(function (it) { out.push({ id: it.id, f: it.fields || {} }); });
      url = r.body['@odata.nextLink'] || null;
    }
    var set = {}; subs().forEach(function (c) { set[String(c.number).toUpperCase()] = 1; });
    return out.filter(function (x) { return set[String(x.f.Subcontract || '').toUpperCase()] && x.f.DocType !== 'message' && x.f.DocType !== 'invoice'; });
  }
  function rowsOf(sub) { return (H.rows || []).filter(function (r) { return String(r.f.Subcontract || '').toUpperCase() === String(sub).toUpperCase(); }); }
  function answerFor(sub, P) {
    var same = function (a, b) { return a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); };
    var inP = rowsOf(sub).filter(function (r) { return r.f.BillingPeriod === P.key || (!r.f.BillingPeriod && same(parseD(r.f.PeriodEnding), P.end)); });
    inP.sort(function (a, b) { return String(b.f.ReceivedAt || b.f.Created || '').localeCompare(String(a.f.ReceivedAt || a.f.Created || '')); });
    return inP[0] || null;
  }
  function kindOf(row) {
    if (!row) return 'waiting';
    if (row.f.DocType === 'no billing') return 'nobill';
    if (APPROVED[row.f.Status]) return 'approved';
    if (row.f.Status === 'revise and resubmit') return 'returned';
    return 'submitted';
  }
  var KIND = { submitted: ['Submitted', 'info'], approved: ['Approved', 'ok'], returned: ['Returned', 'bad'], nobill: ['No billing', 'mut'], waiting: ['Waiting', 'warn'] };
  function parseLog(row) { try { var l = JSON.parse(row.f.StatusLog || '[]'); return Array.isArray(l) ? l : []; } catch (_) { return []; } }
  function checks(row, c) {
    var f = row.f, T = D().pcTerms(), out = [];
    var ok = function (t) { out.push({ k: 'ok', t: t }); }, warn = function (t) { out.push({ k: 'warn', t: t }); }, bad = function (t) { out.push({ k: 'bad', t: t }); };
    var adj = Number(f.AdjustedContract || 0), ctd = Number(f.CompletedToDate || 0), ret = Number(f.RetainageHeld || 0), prev = Number(f.PreviousPayments || 0), due = Number(f.BalanceDue || 0), pv = Number(c.grand_total || 0);
    if (!adj && !ctd) { warn('No Attachment G figures on this row (an emailed pay app not keyed yet). Open the PDF to review.'); return out; }
    if (pv && Math.abs(adj - pv) > 1) warn('Adjusted contract ' + money(adj) + '; Procore shows ' + money(pv) + '. Check change orders.'); else if (pv) ok('Contract matches Procore: ' + money(pv) + '.');
    if (ctd > adj + 1) bad('Completed to date ' + money(ctd) + ' is more than the adjusted contract ' + money(adj) + '.');
    var pct = c.retainage_percent != null && c.retainage_percent !== '' ? Number(c.retainage_percent) : T.retPct;
    var exp = r2(ctd * pct / 100); if (/50/.test(T.retDuration)) exp = Math.min(exp, r2(adj * 0.5 * pct / 100));
    if (Math.abs(ret - exp) > 1) bad('Retainage ' + money(ret) + '; expected ' + money(exp) + ' (' + pct + '% ' + T.retDuration + ').'); else ok('Retainage matches: ' + pct + '% ' + T.retDuration + '.');
    var calc = r2(ctd - ret - prev); if (Math.abs(calc - due) > 1) bad('Due ' + money(due) + ' does not add up: ' + money(ctd) + ' − ' + money(ret) + ' − ' + money(prev) + ' = ' + money(calc) + '.'); else ok('Attachment G figures add up.');
    var last = rowsOf(f.Subcontract).filter(function (r) { return r.id !== row.id && APPROVED[r.f.Status] && r.f.DocType !== 'no billing' && Number(r.f.PayRequest || 0) < Number(f.PayRequest || 1e9); })
      .sort(function (a, b) { return Number(b.f.PayRequest || 0) - Number(a.f.PayRequest || 0); })[0];
    if (last) { var lp = r2(Number(last.f.CompletedToDate || 0) - Number(last.f.RetainageHeld || 0)); if (Math.abs(lp - prev) > 1) warn('Previous approved applications ' + money(prev) + '; approved pay app #' + last.f.PayRequest + ' earned ' + money(lp) + '.'); else ok('Previous approved applications match pay app #' + last.f.PayRequest + '.'); }
    else warn('No earlier approved pay app on file to check previous payments against.');
    var K = H.apd ? compare(row) : null;
    if (K && K.disc) { if (/2% early-payment discount/i.test(f.StatusLog || '') || f.MatchBasis !== 'portal') ok('2% discount applies (' + K.discFrom + '): ' + money(K.disc) + ' off, net ' + money(K.net) + '.'); else warn('The subcontract has the 2% discount but the portal did not take it: ' + money(K.disc) + ' off, net ' + money(K.net) + '. Check the Attachment G.'); }
    else if (/2% early-payment discount/i.test(f.StatusLog || '')) warn('The portal took a 2% discount, but neither Procore nor CE shows one on this subcontract now.');
    if (K && K.ce && Math.abs(K.ce.prev - prev) > 1 && !last) warn('Previous approved applications ' + money(prev) + '; CE has ' + money(K.ce.prev) + '.');
    if (f.ParseNote) warn(f.ParseNote);
    return out;
  }
  async function loadPay(force) {
    var S = D().S, el = $('#pay-list');
    if (!S.project) { el.innerHTML = ''; return; }
    var sel = $('#pay-period'), keys = periodKeys();
    if (!H.period || keys.indexOf(H.period) < 0) H.period = keys[0];
    sel.innerHTML = keys.map(function (k) { var P = periodInfo(k); return '<option value="' + k + '"' + (k === H.period ? ' selected' : '') + '>' + monthName(k) + ' · work through ' + P.end.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) + '</option>'; }).join('');
    el.innerHTML = '<p class="bh-caption">Loading pay apps…</p>';
    try {
      await D().loadRef();
      if (!H.job) H.job = await readJson(JOB_DIR + jobKey() + '.json', { leadDays: 7, note: '' });
      if (force || !H.rows) H.rows = await loadRows();
      if (!H.apd) { try { await apData(); } catch (e) { H.apdErr = e.message; } }
    } catch (e) { el.innerHTML = '<div class="cd-msg is-bad">' + esc(e.message) + ' The Project Hub needs read and write access to the AP Desk site (the same grant the AP desk uses).</div>'; return; }
    applyPending();
    if (H.period) { var ps = $('#pay-period'); if (ps) ps.value = H.period; }
    renderPay();
  }
  var jobKey = function () { return String(D().S.project.project_number || D().S.project.id).replace(/[^A-Za-z0-9_-]/g, '_'); };
  function renderPay() {
    var P = periodInfo(H.period), list = subs(), counts = { submitted: 0, approved: 0, returned: 0, nobill: 0, waiting: 0 };
    $('#pay-when').innerHTML = 'Due <b>' + esc(fmtD(P.due)) + '</b>' + (P.moved ? ' <span class="mut">(the ' + ord(P.day) + ' is not a business day)</span>' : '') +
      ' · reminder <b>' + esc(fmtD(P.rem)) + '</b> <span class="mut">(' + leadDays() + ' days before)</span> · work through <b>' + esc(P.end.toLocaleDateString('en-US', { month: 'long', day: 'numeric' })) + '</b>';
    var rows = list.map(function (c) {
      var a = answerFor(c.number, P), k = kindOf(a); counts[k]++;
      var ck = a && k !== 'nobill' ? checks(a, c) : [], nb = ck.filter(function (x) { return x.k === 'bad'; }).length, nw = ck.filter(function (x) { return x.k === 'warn'; }).length;
      var ckText = k === 'waiting' ? 'No answer yet' : k === 'nobill' ? 'No billing this period' : nb ? nb + ' to fix' : nw ? nw + ' to check' : 'All checks pass';
      var ckK = k === 'waiting' ? 'warn' : k === 'nobill' ? 'mut' : nb ? 'bad' : nw ? 'warn' : 'ok';
      return '<tr' + (H.sel && a && H.sel.id === a.id ? ' class="is-sel"' : '') + '><td><b>' + esc((c.vendor && (c.vendor.company || c.vendor.name)) || '') + '</b><div class="mut">' + esc(c.number) + ' · ' + esc(c.title || '') + '</div></td>' +
        '<td class="num">' + money(c.grand_total) + '</td><td class="num"><b>' + (a && a.f.BalanceDue != null && k !== 'nobill' ? money(a.f.BalanceDue) : '—') + '</b></td>' +
        '<td><span class="hub-pill ' + KIND[k][1] + '">' + KIND[k][0] + '</span>' + (a && a.f.PayRequest ? '<div class="mut">Pay app ' + esc(a.f.PayRequest) + '</div>' : '') + '</td><td class="hub-ck ' + ckK + '">' + esc(ckText) + '</td>' +
        '<td class="num">' + (a && k !== 'nobill' ? '<button class="bh-btn-outline" data-review="' + esc(a.id) + '" data-sub="' + esc(c.number) + '">Review</button>' : '') + '</td></tr>';
    });
    $('#pay-kpis').innerHTML = [['Subcontracts', list.length, ''], ['Submitted', counts.submitted, 'info'], ['No billing', counts.nobill, 'mut'], ['Waiting', counts.waiting, 'warn'], ['Approved', counts.approved, 'ok'], ['Returned', counts.returned, 'bad']]
      .map(function (k) { return '<div class="' + k[2] + '"><span>' + k[0] + '</span><b>' + k[1] + '</b></div>'; }).join('');
    $('#pay-list').innerHTML = list.length ? '<div class="bh-table-wrap"><table class="bh-table hub-table"><thead><tr><th>Trade partner</th><th class="num">Contract</th><th class="num">Due this period</th><th>Answer</th><th>Checks</th><th><span class="vh">Action</span></th></tr></thead><tbody>' + rows.join('') + '</tbody></table></div>'
      : '<p class="bh-caption">No approved, executed subcontracts on this job.</p>';
    if (H.sel) renderReview(); else $('#pay-review').hidden = true;
  }
  async function fileItem(path) { var Q = await queue(); return gcall('GET', '/sites/' + Q.sid + '/drive/root:/' + encodeURI(path) + '?$select=webUrl,size,name'); }
  function filesOf(row) {
    var out = []; if (row.f.PdfPath) out.push({ name: String(row.f.PdfPath).split('/').pop(), path: row.f.PdfPath, main: true });
    parseLog(row).forEach(function (e) { (e.files || []).forEach(function (x) { if (x.path && !out.some(function (o) { return o.path === x.path; })) out.push({ name: x.name || x.path.split('/').pop(), path: x.path }); }); });
    return out;
  }
  function renderReview() {
    var row = H.sel, el = $('#pay-review'); if (!row) { el.hidden = true; return; }
    var c = subs().find(function (x) { return String(x.number).toUpperCase() === String(row.f.Subcontract).toUpperCase(); }) || {};
    var f = row.f, k = kindOf(row), ck = checks(row, c), log = parseLog(row), files = filesOf(row);
    var line = function (l, v, b) { return '<tr' + (b ? ' class="tot"' : '') + '><td>' + l + '</td><td class="num">' + (v == null || v === '' ? '—' : money(v)) + '</td></tr>'; };
    el.innerHTML = '<div class="hub-eyebrow"><span class="sq"></span>PM review</div>' +
      '<h3>' + esc(f.Vendor || (c.vendor && c.vendor.company) || '') + '</h3><div class="mut">' + esc(f.Subcontract) + (f.PayRequest ? ' · pay app ' + esc(f.PayRequest) : '') + ' · ' + esc(KIND[k][0].toLowerCase()) + (f.ReceivedAt ? ' ' + esc(new Date(f.ReceivedAt).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })) : '') + (f.MatchBasis === 'portal' ? ' through the portal' : '') + '</div>' +
      (files.some(function (x) { return x.main; }) ? '<div class="hub-actions"><button class="bh-btn-outline" data-file="' + files.findIndex(function (x) { return x.main; }) + '">Open the Attachment G</button></div>' : '') +
      compareTable(row, compare(row), false) + (H.apdErr ? '<p class="mut small">CE and Procore columns not loaded: ' + esc(H.apdErr) + '</p>' : '') +
      '<table class="hub-fig" hidden>' + line('Original contract', f.OriginalContract) + line('Change orders', f.ChangeOrders) + line('Adjusted contract', f.AdjustedContract, 1) + line('Completed and stored to date', f.CompletedToDate) + line('Retainage', f.RetainageHeld) + line('Previous payments', f.PreviousPayments) + line('Due this period', f.BalanceDue, 1) + '</table>' +
      '<div class="hub-eyebrow small">Checked against the commitment and prime contract</div><ul class="hub-checks">' + ck.map(function (x) { return '<li class="' + x.k + '">' + esc(x.t) + '</li>'; }).join('') + '</ul>' +
      '<div class="hub-files">' + (files.length ? files.map(function (x, i) { return '<button class="hub-link" data-file="' + i + '">' + esc(x.main ? 'Attachment G · ' + x.name : x.name) + '</button>'; }).join('') : '<span class="mut">No files on this row.</span>') + '</div>' +
      '<details class="hub-hist"><summary>Approval history (' + log.length + ')</summary><ol>' + log.map(function (e) { return '<li><b>' + esc(e.s || '') + '</b> · ' + esc(e.at ? new Date(e.at).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : '') + ' · ' + esc(e.by || '') + (e.note ? '<div>' + esc(e.note) + '</div>' : '') + '</li>'; }).join('') + '</ol></details>' +
      (k === 'submitted' ? '<label class="hub-note">Note (required to approve as noted or return)<textarea id="rv-note" rows="3"></textarea></label>' +
        '<div class="hub-actions"><button class="bh-btn" data-decide="approve">Approve and email payables@</button><button class="bh-btn-outline" data-decide="noted">Approve as noted</button><button class="bh-btn-outline hub-danger" data-decide="return">Return to trade partner</button></div>'
        : k === 'approved' ? '<div class="hub-actions"><button class="bh-btn-outline" data-decide="resend">Email payables@ again</button></div>' : '') +
      '<p class="mut small">Approving emails payables@ with the figures, the Attachment G and backup, and this approval history. Drew enters it in ComputerEase as he does today.</p>' +
      '<button class="hub-close" data-close aria-label="Close review">×</button>';
    el.hidden = false;
    el._files = files;
  }
  async function openFile(x) {
    try { var Q = await queue(); var bytes = await gbytes('/sites/' + Q.sid + '/drive/root:/' + encodeURI(x.path) + ':/content');
      var type = /\.pdf$/i.test(x.name) ? 'application/pdf' : /\.png$/i.test(x.name) ? 'image/png' : /\.jpe?g$/i.test(x.name) ? 'image/jpeg' : 'application/octet-stream';
      window.open(URL.createObjectURL(new Blob([bytes], { type: type })), '_blank'); }
    catch (e) { D().msg('Could not open ' + x.name + ': ' + e.message, true); }
  }
  async function packet(row) {   // attachments up to the mail size limit, links for the rest
    var atts = [], links = [], size = 0;
    for (var i = 0, fs = filesOf(row); i < fs.length; i++) {
      var x = fs[i], it = await fileItem(x.path);
      var sz = it.ok ? Number(it.body.size || 0) : 0;
      if (it.ok && size + sz <= MAIL_ATTACH_MAX) {
        try { var Q = await queue(); var bytes = await gbytes('/sites/' + Q.sid + '/drive/root:/' + encodeURI(x.path) + ':/content'); atts.push({ name: x.name, bytes: bytes, type: /\.pdf$/i.test(x.name) ? 'application/pdf' : undefined }); size += bytes.length; continue; } catch (_) {}
      }
      links.push({ name: x.name, url: it.ok ? it.body.webUrl : '' });
    }
    return { atts: atts, links: links };
  }
  function payablesHtml(row, c, decision, note) {
    var f = row.f, P = periodInfo(H.period), log = parseLog(row), ck = checks(row, c), pj = D().S.project;
    var tr = function (l, v) { return '<tr><td style="padding:3px 16px 3px 0;color:#484b4d">' + esc(l) + '</td><td style="padding:3px 0;text-align:right">' + (v == null || v === '' ? '—' : money(v)) + '</td></tr>'; };
    return '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#211e1e">' +
      '<p><b>' + esc(decision) + '</b> by ' + esc(me().name) + ' in the Project Hub. Ready to enter in ComputerEase.</p>' + (note ? '<p><b>PM note:</b> ' + esc(note) + '</p>' : '') +
      '<table style="border-collapse:collapse;font-size:14px">' + [['Job', (pj.project_number || '') + ' ' + pj.name], ['Trade partner', f.Vendor || (c.vendor && (c.vendor.company || c.vendor.name))], ['Subcontract', f.Subcontract], ['Pay app', f.PayRequest], ['Billing period', monthName(P.key) + ' (work through ' + P.end.toLocaleDateString('en-US') + ')'], ['AP list item', 'AP-' + row.id]]
        .map(function (x) { return '<tr><td style="padding:3px 16px 3px 0;color:#484b4d">' + esc(x[0]) + '</td><td>' + esc(x[1] || '—') + '</td></tr>'; }).join('') + '</table><br>' +
      compareTable(row, compare(row), true) +
      '<p><b>Checks</b></p><ul>' + ck.map(function (x) { return '<li>' + (x.k === 'ok' ? '' : x.k === 'bad' ? '<b>Issue:</b> ' : '<b>Check:</b> ') + esc(x.t) + '</li>'; }).join('') + '</ul>' +
      '<p><b>Approval history</b></p><ol>' + log.map(function (e) { return '<li>' + esc(e.s || '') + ' · ' + esc(e.at ? new Date(e.at).toLocaleString('en-US') : '') + ' · ' + esc(e.by || '') + (e.note ? '<br><span style="color:#484b4d">' + esc(e.note) + '</span>' : '') + '</li>'; }).join('') + '</ol></div>';
  }
  async function emailPayables(row, c, decision, note) {
    var pk = await packet(row);
    var html = payablesHtml(row, c, decision, note) + (pk.links.length ? '<p>Files too large to attach (open in SharePoint):</p><ul>' + pk.links.map(function (l) { return '<li>' + (l.url ? '<a href="' + esc(l.url) + '">' + esc(l.name) + '</a>' : esc(l.name) + ' (not found)') + '</li>'; }).join('') + '</ul>' : '');
    var subj = subjectFor('Pay App Approved', row.f, row.f.Vendor || (c.vendor && (c.vendor.company || c.vendor.name)) || '') + (/noted/i.test(decision) ? ' (as noted)' : '');
    await sendMail([PAYABLES], subj, html, pk.atts);
  }
  async function decide(kind) {
    if (H.busy) return;
    var row = H.sel, note = ($('#rv-note') || {}).value || ''; note = note.trim();
    var c = subs().find(function (x) { return String(x.number).toUpperCase() === String(row.f.Subcontract).toUpperCase(); }) || {};
    if ((kind === 'noted' || kind === 'return') && !note) { D().msg('Add a note first: it goes to ' + (kind === 'return' ? 'the trade partner' : 'payables@') + ' with the decision.', true); return; }
    var bad = checks(row, c).filter(function (x) { return x.k === 'bad'; });
    if (kind === 'approve' && bad.length && !confirm('This pay app has ' + bad.length + ' issue(s):\n\n' + bad.map(function (x) { return '• ' + x.t; }).join('\n') + '\n\nApprove anyway? Consider "Approve as noted" with a note instead.')) return;
    H.busy = true; D().msg('Saving…');
    try {
      var Q = await queue(), now = new Date().toISOString(), m = me();
      if (kind !== 'resend') {
        var st = kind === 'approve' ? 'PM approved' : kind === 'noted' ? 'approved as noted' : 'revise and resubmit';
        var log = parseLog(row);
        log.push({ s: st, at: now, by: m.name + ' <' + m.email + '>', note: (kind === 'approve' ? 'Approved in the Project Hub.' : kind === 'noted' ? 'Approved as noted in the Project Hub: ' + note : 'Returned to the trade partner from the Project Hub: ' + note) + (kind !== 'return' ? ' Emailed to ' + PAYABLES + '.' : '') + (kind === 'approve' && note ? ' Note: ' + note : '') });
        var fields = { Status: st, StatusLog: JSON.stringify(log) };
        if (kind !== 'return') { fields.ApprovedBy = m.email; fields.ApprovedAt = now; }
        else fields.AssignedApprover = m.email;   // the portal sends the resubmission straight back to this PM
        var r = await gcall('PATCH', '/sites/' + Q.sid + '/lists/' + Q.lid + '/items/' + encodeURIComponent(row.id) + '/fields', fields);
        if (!r.ok) throw new Error('Could not save the decision (' + r.status + (r.body && r.body.error ? ': ' + r.body.error.message : '') + ').');
        Object.assign(row.f, fields);
      }
      if (kind === 'return') {
        var to = (parseLog(row).map(function (e) { var x = String(e.by || '').match(/<([^>]+@[^>]+)>/); return x && x[1]; }).find(Boolean)) || '';
        if (to && !/baileyharris\.com$/i.test(to)) {
          var lk = null; try { lk = await portalLink(c, row.f.BillingPeriod); } catch (e) { console.warn('portal link', e); }
          var btn = lk ? '<p><a href="' + esc(lk.link) + '" style="display:inline-block;background:#211e1e;color:#ffffff;padding:10px 18px;text-decoration:none;font-weight:700">Open my pay apps</a></p><p style="font-size:12px;color:#6b6b6b">This private link shows every Bailey-Harris pay app you have due this period and works until ' + esc(new Date(lk.until + 'T12:00:00').toLocaleDateString('en-US', { month: 'long', day: 'numeric' })) + '. Please don\'t forward it outside your company.</p>'
            : '<p>Please resubmit it with the link from your last pay app reminder email.</p>';
          try { await sendMail([to], subjectFor('Pay App Returned', row.f, ''), '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#211e1e"><p>Your pay application for subcontract ' + esc(row.f.Subcontract) + (row.f.PayRequest ? ' (Pay App #' + esc(row.f.PayRequest) + ')' : '') + ' on ' + esc(row.f.JobNumber || '') + ' needs a correction before it can be paid:</p><p><b>' + esc(note) + '</b></p>' + btn + '<p>Questions? Reply to this email or contact ' + esc(m.name) + ' at ' + esc(m.email) + '.</p><p>Thank you,<br>Bailey-Harris Construction</p></div>', [], [m.email]);
            D().msg('Returned. ' + to + ' was emailed the note' + (lk ? ' and a new portal link.' : ', but a new portal link could not be made (pay.baileyharris.com not reachable), so they were told to use their last reminder link.'), !lk); }
          catch (e) { D().msg('Returned and saved, but the email to ' + to + ' failed: ' + e.message, true); }
        } else D().msg('Returned and saved. No trade partner email on this pay app: send them the note yourself.', true);
      } else {
        try { await emailPayables(row, c, kind === 'noted' ? 'Approved as noted' : 'Approved', note); D().msg('Approved and emailed to payables@.'); }
        catch (e) { D().msg('Approved and saved, but the email to payables@ failed: ' + e.message + '. Use "Email payables@ again".', true); }
      }
      renderPay();
      if (H.mine) { H.mine = H.mine.filter(function (x) { return String(x.id) !== String(row.id); }); renderMine(); }
    } catch (e) { D().msg(e.message, true); }
    finally { H.busy = false; }
  }

  /* =====================================================================
     REMINDER EMAIL — company template + this job's settings
     ===================================================================== */
  /* the email itself is built by reminder-email.js (window.BHReminder), the same file the reminder script uses */
  var R = function () { return window.BHReminder; };
  var DEFAULT_TPL = window.BHReminder ? window.BHReminder.DEFAULT_TPL : { subject: '', body: '' };
  var FIELDS = window.BHReminder ? window.BHReminder.FIELDS : [];
  async function loadEmail() {
    var S = D().S, ed = $('#em-edit');
    if (!S.project) { ed.innerHTML = ''; $('#em-preview').innerHTML = ''; return; }
    ed.innerHTML = '<p class="bh-caption">Loading the template…</p>';
    try {
      await D().loadRef();
      if (!H.tpl) H.tpl = await readJson(TPL_PATH, DEFAULT_TPL);
      if (!H.job) H.job = await readJson(JOB_DIR + jobKey() + '.json', { leadDays: 7, note: '' });
      if (!H.rows) { try { H.rows = await loadRows(); } catch (_) { H.rows = []; } }
    } catch (e) { ed.innerHTML = '<div class="cd-msg is-bad">' + esc(e.message) + '</div>'; return; }
    renderEmail();
  }
  function renderEmail() {
    var t = H.tpl, j = H.job, list = subs();
    $('#em-edit').innerHTML = '<div class="hub-card"><div class="hub-cardh"><h3>Company template</h3><span class="mut small">' + (t.updatedBy ? 'Last saved by ' + esc(t.updatedBy) + ' · ' + esc(new Date(t.updatedAt).toLocaleDateString('en-US')) : 'Default wording, not saved yet') + '</span></div>' +
      '<label class="hub-field">Subject<input id="em-subject" type="text" value="' + esc(t.subject) + '"></label>' +
      '<label class="hub-field">Message<textarea id="em-body" rows="16">' + esc(t.body) + '</textarea></label>' +
      '<div class="mut small">Insert a field at the cursor:</div><div class="hub-chips">' + FIELDS.map(function (f) { return '<button type="button" class="hub-chip" data-field="' + esc(f) + '">' + esc(f) + '</button>'; }).join('') + '</div>' +
      '<div class="hub-actions"><button class="bh-btn" id="em-save-tpl">Save company template</button><button class="bh-btn-outline" id="em-reset">Back to default wording</button></div></div>' +
      '<div class="hub-card"><div class="hub-cardh"><h3>This job · ' + esc(D().S.project.project_number || '') + '</h3><span class="mut small">' + (j.updatedBy ? 'Saved by ' + esc(j.updatedBy) : '') + '</span></div>' +
      '<label class="hub-field">Send the reminder<select id="em-lead">' + [5, 7, 10].map(function (n) { return '<option value="' + n + '"' + (Number(j.leadDays || 7) === n ? ' selected' : '') + '>' + n + ' days before the due date</option>'; }).join('') + '</select></label>' +
      '<label class="hub-field">Job note (added to every reminder on this job)<textarea id="em-note" rows="3">' + esc(j.note || '') + '</textarea></label>' +
      '<p class="mut small">Due date: the prime contract\'s Monthly Billing Date (' + esc(D().pcTerms().billing) + '), moved to the next business day. Sent to each subcontract\'s signer, then the vendor\'s primary contact.</p>' +
      '<div class="hub-actions"><button class="bh-btn" id="em-save-job">Save job settings</button></div></div>';
    $('#em-preview').innerHTML = '<div class="hub-cardh"><span class="hub-eyebrow">Preview</span><label class="hub-inline">as <select id="em-as">' + list.map(function (c, i) { return '<option value="' + i + '">' + esc((c.vendor && (c.vendor.company || c.vendor.name)) || '') + ' · ' + esc(c.number) + '</option>'; }).join('') + '</select></label>' +
      '<button class="bh-btn-outline" id="em-test">Send me a test</button></div><div id="em-render"></div>';
    renderPreview();
  }
  /* preview figures from Procore and the pay app list; the monthly email itself uses the ComputerEase records and Billy compliance */
  function previewData(c, note) {
    var P = periodInfo(periodKeys()[0]), pj = D().S.project, m = me(), T = D().pcTerms();
    var last = rowsOf(c.number).filter(function (r) { return APPROVED[r.f.Status] && r.f.DocType !== 'no billing'; }).sort(function (a, b) { return Number(b.f.PayRequest || 0) - Number(a.f.PayRequest || 0); })[0];
    var pct = c.retainage_percent != null && c.retainage_percent !== '' ? Number(c.retainage_percent) : T.retPct;
    var rec = last ? { pr: Number(last.f.PayRequest || 0), ctd: Number(last.f.CompletedToDate || 0), ret: Number(last.f.RetainageHeld || 0), prev: r2(Number(last.f.CompletedToDate || 0) - Number(last.f.RetainageHeld || 0)) } : null;
    var isoD = function (x) { return iso(x); }, until = new Date(P.due); until.setDate(until.getDate() + 31);
    return { vendor: (c.vendor && (c.vendor.company || c.vendor.name)) || '', firstName: '', job: { number: pj.project_number || '', name: pj.name }, due: isoD(P.due), periodEnd: isoD(P.end), linkUntil: isoD(until),
      link: 'https://pay.baileyharris.com/', note: note || '', pm: { name: m.name, email: m.email },
      subs: [{ number: c.number, title: c.title || '', nextPr: rec ? rec.pr + 1 : 1, orig: Number(c.grand_total || 0), earlier: 0, fresh: 0, adj: Number(c.grand_total || 0), cos: [], rec: rec,
        retText: pct + '% ' + (/50/.test(T.retDuration) ? 'through 50%' : 'throughout'), last: null, comp: null }] };
  }
  function render(tpl, c, note) { return R().build(tpl, previewData(c, note)); }
  function current() { return { subject: ($('#em-subject') || {}).value || H.tpl.subject, body: ($('#em-body') || {}).value || H.tpl.body }; }
  function renderPreview() {
    var list = subs(), c = list[Number(($('#em-as') || {}).value || 0)];
    if (!c) { $('#em-render').innerHTML = '<p class="bh-caption">No approved, executed subcontracts to preview with.</p>'; return; }
    var r = render(current(), c, ($('#em-note') || {}).value), asText = H.emText;
    $('#em-render').innerHTML = '<div class="hub-viewtoggle" role="group" aria-label="Preview format"><button type="button" class="hub-chip' + (asText ? '' : ' is-on') + '" data-emview="html" aria-pressed="' + !asText + '">HTML</button><button type="button" class="hub-chip' + (asText ? ' is-on' : '') + '" data-emview="text" aria-pressed="' + !!asText + '">Plain text</button></div>' +
      '<div class="mut small">To: ' + esc((c.vendor && (c.vendor.company || c.vendor.name)) || '') + ' (subcontract signer) · first name, change orders, our records, last check and Billy compliance come from ComputerEase and Billy when it sends</div><div class="hub-subj">' + esc(r.subject) + '</div>' +
      (asText ? '<pre class="hub-plain">' + esc(r.text) + '</pre>' : '<div class="hub-mail">' + r.inner + '</div>');
  }
  async function saveTpl(reset) {
    var m = me(), t = reset ? Object.assign({}, DEFAULT_TPL) : current();
    if (!reset && !confirm('Save this wording as the company template? It is used for every job\'s reminders.')) return;
    t.updatedBy = m.name; t.updatedAt = new Date().toISOString();
    try { await writeJson(TPL_PATH, t); H.tpl = t; renderEmail(); D().msg(reset ? 'Template set back to the default wording.' : 'Company template saved.'); } catch (e) { D().msg(e.message, true); }
  }
  async function saveJob() {
    var m = me(), j = { project: D().S.project.project_number || '', procoreProjectId: D().S.project.id, leadDays: Number($('#em-lead').value), note: $('#em-note').value.trim(), billing: D().pcTerms().billing, updatedBy: m.name, updatedAt: new Date().toISOString() };
    try { await writeJson(JOB_DIR + jobKey() + '.json', j); H.job = j; renderEmail(); summary(); D().msg('Job settings saved.'); } catch (e) { D().msg(e.message, true); }
  }
  async function sendTest() {
    var list = subs(), c = list[Number($('#em-as').value || 0)]; if (!c) return;
    var r = render(current(), c, $('#em-note').value), m = me();
    r.subject = '[TEST] ' + r.subject;
    try {
      var res = await gcall('POST', '/me/sendMail', R().mime(r, m.email, [m.email]), { 'Content-Type': 'text/plain' });   // MIME: HTML with a plain-text part
      if (!res.ok) throw new Error((res.body && res.body.error && res.body.error.message) || ('HTTP ' + res.status) + (res.status === 403 ? ' (Mail.Send permission not granted yet)' : ''));
      D().msg('Test sent to ' + m.email + ' (HTML with a plain-text version).');
    } catch (e) { D().msg('Test not sent: ' + e.message, true); }
  }
  function insertField(f) {
    var a = document.activeElement && (document.activeElement.id === 'em-subject' || document.activeElement.id === 'em-body') ? document.activeElement : (H.lastFocus || $('#em-body'));
    var s = a.selectionStart || a.value.length, e = a.selectionEnd || s;
    a.value = a.value.slice(0, s) + f + a.value.slice(e); a.focus(); a.selectionStart = a.selectionEnd = s + f.length; renderPreview();
  }

  /* ---------------- CE and Procore figures for the comparison (shared files, no Procore calls) ----------------
     ap_commitments.json (the Procore commitments pull) and ap_ce_invoices.json (CE Subcontractor Detail), both in
     APData on /sites/Dashboard - the same files the AP desk and the portal read. */
  async function apData() {
    if (H.apd) return H.apd;
    var s = await gcall('GET', '/sites/' + SP_HOST + ':/sites/Dashboard?$select=id'); if (!s.ok) throw new Error('Dashboard site ' + s.status);
    var d = await gcall('GET', '/sites/' + s.body.id + '/drives?$select=id,name'); var lib = ((d.body && d.body.value) || []).find(function (x) { return x.name === 'APData'; });
    if (!lib) throw new Error('APData library not found');
    var get = async function (f) { var r = await gcall('GET', '/drives/' + lib.id + '/root:/' + f + ':/content'); return r.ok && r.body && typeof r.body === 'object' ? r.body : null; };
    var cm = await get('ap_commitments.json'), ce = await get('ap_ce_invoices.json'), byNum = {};
    ((cm && cm.commitments) || []).forEach(function (c) { byNum[String(c.number || '').trim().toUpperCase()] = c; });
    H.apd = { cm: byNum, ce: (ce && ce.subcontracts) || {}, ceAsOf: ce && (ce.as_of || ce.generated_at), cmAsOf: cm && (cm.generated_at || cm.as_of) };
    return H.apd;
  }
  var nn = function (v) { return v == null || v === '' || isNaN(Number(v)) ? null : Number(v); };
  /* the figures, three ways: what the trade partner billed, what CE has, what Procore has (portal rules) */
  function compare(row) {
    var f = row.f, num = String(f.Subcontract || '').trim().toUpperCase(), A = H.apd || { cm: {}, ce: {} }, pc = A.cm[num] || null, ce = A.ce[num] || null;
    var per = String(f.PeriodEnding || '').slice(0, 10) || (f.BillingPeriod ? iso(periodInfo(f.BillingPeriod).end) : iso(new Date()));
    /* CE: base + RFCs approved by the period end; prior invoices; previous approved applications = what CE certified */
    var C = null;
    if (ce) {
      var cos = ce.change_orders || [], sum = function (a) { return r2(a.reduce(function (t, x) { return t + (Number(x.amount) || 0); }, 0)); };
      var inP = cos.filter(function (x) { return (!x.status || /approv/i.test(x.status)) && (!x.date_approved || String(x.date_approved).slice(0, 10) <= per); });
      var base = ce.base_amount != null ? Number(ce.base_amount) : r2((Number(ce.amount) || 0) - sum(cos));
      var inv = (ce.invoices || []).filter(function (x) { return !x.date || String(x.date).slice(0, 10) <= per; });
      var cert = function (x) { var p = nn(x.payable) || 0, pd = nn(x.paid) || 0; return p < 0 ? p : p + Math.max(0, pd - p); };
      C = { orig: base, cos: sum(inP), adj: r2(base + sum(inP)), ctd: r2(inv.reduce(function (t, x) { return t + (nn(x.total) || 0); }, 0)), ret: r2(inv.reduce(function (t, x) { return t + (nn(x.retention) || 0); }, 0)),
        prev: r2(inv.reduce(function (t, x) { return t + cert(x); }, 0)), lastPr: inv.length ? inv[inv.length - 1].number : '', n: inv.length, disc: /^y/i.test(String(ce.discount_2pct || '')) };
    }
    /* Procore (ap_commitments.json): original, COs approved by the period end, retainage rule on the billed amount */
    var P = null;
    if (pc) {
      var pcos = Array.isArray(pc.change_orders) ? r2(pc.change_orders.filter(function (x) { return /approved/i.test(x.state || x.status || '') && (!x.approved_date || String(x.approved_date).slice(0, 10) <= per); }).reduce(function (t, x) { return t + (Number(x.amount) || 0); }, 0)) : nn(pc.approved_change_orders);
      var orig = nn(pc.contract_value), adj = orig != null && pcos != null ? r2(orig + pcos) : nn(pc.revised_contract);
      var rp = nn(pc.retainage_percent), thru50 = /50/.test(pc.retainage_duration || '') && /thru|through/i.test(pc.retainage_duration || ''), ctd = nn(f.CompletedToDate);
      P = { orig: orig, cos: pcos, adj: adj, ret: rp != null && ctd != null ? r2((thru50 && adj ? Math.min(ctd, adj / 2) : ctd) * rp / 100) : null, rule: rp != null ? rp + '% ' + (thru50 ? 'thru 50%' : 'throughout') : '', disc: /^y/i.test(String(pc.discount_2pct || '')) };
    }
    /* previous approved applications on the desk: the latest earlier approved pay app's line 6 (portal rule) */
    var last = rowsOf(f.Subcontract).filter(function (r) { return r.id !== row.id && APPROVED[r.f.Status] && r.f.DocType !== 'no billing' && Number(r.f.PayRequest || 0) < Number(f.PayRequest || 1e9); })
      .sort(function (a, b) { return Number(b.f.PayRequest || 0) - Number(a.f.PayRequest || 0); })[0];
    var deskPrev = last ? { v: r2(Number(last.f.CompletedToDate || 0) - Number(last.f.RetainageHeld || 0)), pr: last.f.PayRequest } : null;
    var due = nn(f.BalanceDue), hasDisc = !!((P && P.disc) || (C && C.disc));
    var disc = hasDisc && due > 0 ? r2(due * 0.02) : 0;
    return { pc: P, ce: C, deskPrev: deskPrev, disc: disc, discFrom: hasDisc ? (P && P.disc ? 'Procore' : 'CE') + ' subcontract says 2% discount' : '', net: due != null ? r2(due - disc) : null };
  }
  function compareTable(row, K, forEmail) {
    var f = row.f, ctd = nn(f.CompletedToDate), ret = nn(f.RetainageHeld), prev = nn(f.PreviousPayments), earned = ctd != null && ret != null ? r2(ctd - ret) : null;
    var dif = function (a, b) { return a != null && b != null && Math.abs(a - b) > 1; };
    var cell = function (v, ref) { if (v == null) return forEmail ? '—' : '<span class="mut">—</span>'; var off = dif(ref, v);
      return money(v) + (forEmail ? (off ? ' *' : '') : (ref == null ? '' : off ? ' <span class="hub-pill bad">' + (ref - v > 0 ? '+' : '−') + money(Math.abs(r2(ref - v))) + '</span>' : ' <span class="hub-pill ok">✓</span>')); };
    var C = K.ce || {}, P = K.pc || {};
    var prevNote = K.deskPrev ? 'Previous approved pay app #' + K.deskPrev.pr + ': ' + money(K.deskPrev.v) : (K.ce ? 'CE approved applications' + (C.lastPr ? ' through ' + C.lastPr : '') : 'No earlier approved application on file');
    var rows = [
      ['1. Original contract', nn(f.OriginalContract), C.orig, P.orig],
      ['2. Approved change orders', nn(f.ChangeOrders), C.cos, P.cos],
      ['3. Adjusted contract', nn(f.AdjustedContract), C.adj, P.adj],
      ['4. Completed and stored to date', ctd, K.ce ? C.ctd : null, null, K.ce ? 'CE: billed before this pay app (' + C.n + ' invoice' + (C.n === 1 ? '' : 's') + ')' : ''],
      ['5. Retainage', ret, null, P.ret, P.rule ? 'Procore: ' + P.rule + ' on line 4' : ''],
      ['6. Earned less retainage', earned, null, null],
      ['7. Less previous approved applications for payment', prev, K.ce ? C.prev : null, K.deskPrev ? K.deskPrev.v : null, prevNote + (K.deskPrev ? ' (Procore column)' : '')],
      ['8. Balance due this period', nn(f.BalanceDue), null, null, earned != null && prev != null && dif(r2(earned - prev), nn(f.BalanceDue)) ? 'Lines 6 − 7 = ' + money(r2(earned - prev)) : '']
    ];
    if (K.disc) { rows.push(['Less 2% discount', -K.disc, null, null, K.discFrom]); rows.push(['Net payment after discount', K.net, null, null, '']); }
    var tdS = forEmail ? ' style="padding:4px 10px;border-bottom:1px solid #e8e6e7;text-align:right"' : ' class="num"', thS = forEmail ? ' style="padding:4px 10px;text-align:right;color:#484b4d"' : ' class="num"';
    var lab = forEmail ? ' style="padding:4px 10px 4px 0;border-bottom:1px solid #e8e6e7"' : '';
    return '<table class="hub-cmp"' + (forEmail ? ' style="border-collapse:collapse;font-size:13px"' : '') + '><thead><tr><th' + (forEmail ? ' style="text-align:left;color:#484b4d"' : '') + '>Attachment G</th><th' + thS + '>Pay app</th><th' + thS + '>CE</th><th' + thS + '>Procore</th></tr></thead><tbody>' +
      rows.map(function (x, i) { var b = i === 7 || x[0].indexOf('Net') === 0;
        return '<tr' + (b && !forEmail ? ' class="tot"' : '') + '><td' + lab + '>' + (b ? '<b>' + esc(x[0]) + '</b>' : esc(x[0])) + (x[4] ? '<div class="mut small"' + (forEmail ? ' style="color:#6b6b6b;font-size:11px"' : '') + '>' + esc(x[4]) + '</div>' : '') + '</td><td' + tdS + '>' + (x[1] == null ? '—' : (b ? '<b>' + money(x[1]) + '</b>' : money(x[1]))) + '</td><td' + tdS + '>' + cell(x[2], x[1]) + '</td><td' + tdS + '>' + cell(x[3], x[1]) + '</td></tr>'; }).join('') +
      '</tbody></table>' + (forEmail ? '<p style="font-size:11px;color:#6b6b6b">* differs from the pay app by more than $1. CE as of ' + esc(H.apd && H.apd.ceAsOf ? String(H.apd.ceAsOf).slice(0, 10) : '?') + '; Procore from the commitments pull.</p>' : '');
  }

  /* ---------------- portal link (return email) ---------------- */
  async function portalLink(c, period) {
    var vid = c && c.vendor && c.vendor.id; if (!vid || !period) return null;
    var t = await D().tokens();
    var r = await fetch(PORTAL + '/api/link', { method: 'POST', headers: { Authorization: 'Bearer ' + t.accessToken, 'Content-Type': 'application/json' }, body: JSON.stringify({ vendor_id: String(vid), period: period }) });
    var j = await r.json().catch(function () { return {}; });
    if (!r.ok || !j.link) throw new Error(j.error || ('HTTP ' + r.status));
    return j;
  }

  /* ---------------- waiting for your approval (every job) ---------------- */
  var normJob = function (v) { return String(v || '').trim().toUpperCase().replace(/\s+/g, ''); };
  async function myJobs(email) {
    var s = await gcall('GET', '/sites/' + SP_HOST + ':' + RES_SITE + '?$select=id'); if (!s.ok) return null;
    var get = function (f) { return gcall('GET', '/sites/' + s.body.id + '/drive/root:/' + encPath('Compliance Dashboard/' + f) + ':/content'); };
    var r = await get('jobs_public.json'); if (r.status === 404) r = await get('projects_data.json');
    if (!r.ok || !r.body || typeof r.body !== 'object') return null;
    var out = {}; (r.body.projects || []).forEach(function (p) { if ((p.team || []).some(function (t) { return String(t.email || '').trim().toLowerCase() === email && PM_ROLE.test(String(t.role || '').trim()); })) out[normJob(p.project_number)] = 1; });
    return out;
  }
  async function loadMine() {
    var el = $('#pay-mine'); if (!el) return;
    var m = me(); if (!m.email) return;
    try {
      var Q = await queue(), all = [], pref = { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' };
      var url = '/sites/' + Q.sid + '/lists/' + Q.lid + '/items?$expand=fields($select=Title,OriginalTitle,Revision,Status,DocType,MatchBasis,JobNumber,Vendor,Subcontract,PayRequest,BalanceDue,ReceivedAt,BillingPeriod,AssignedApprover)&$top=999';
      for (var p = 0; url && p < 40; p++) { var r = await gcall('GET', url, null, pref); if (!r.ok) throw new Error('Could not read the pay app list (' + r.status + ').');
        (r.body.value || []).forEach(function (it) { all.push({ id: it.id, f: it.fields || {} }); }); url = r.body['@odata.nextLink'] || null; }
      var g = {}; all.filter(function (x) { return x.f.MatchBasis === 'portal' && x.f.DocType !== 'no billing' && x.f.DocType !== 'invoice'; })
        .forEach(function (x) { var k = String(x.f.OriginalTitle || x.f.Title || '').trim(); (g[k] = g[k] || []).push(x); });
      var open = Object.keys(g).map(function (k) { return g[k].sort(function (a, b) { return ((Number(a.f.Revision) || 1) - (Number(b.f.Revision) || 1)) || (a.id - b.id); })[g[k].length - 1]; })
        .filter(function (x) { return !PAYAPP_DONE[String(x.f.Status || '')]; });
      var every = PAYAPP_ALL.indexOf(m.email) >= 0, jobs = every ? null : await myJobs(m.email);
      H.mine = open.filter(function (x) { return every || String(x.f.AssignedApprover || '').toLowerCase() === m.email || (jobs && jobs[normJob(x.f.JobNumber)]); })
        .sort(function (a, b) { return String(a.f.ReceivedAt || '').localeCompare(String(b.f.ReceivedAt || '')); });
      H.mineEvery = every; H.mineNoJobs = !every && !jobs;
      if (H.notify == null) { var pf = await readJson(PEOPLE_DIR + m.email + '.json', {}); H.notify = !(pf && pf.payAppEmails === false); }
    } catch (e) { el.innerHTML = '<div class="cd-msg is-bad">' + esc(e.message) + '</div>'; return; }
    renderMine();
  }
  function renderMine() {
    var el = $('#pay-mine'); if (!el || !H.mine) return;
    var rows = H.mine.map(function (x) {
      var rd = x.f.ReceivedAt ? new Date(x.f.ReceivedAt).toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) : '';
      return '<tr><td><b>' + esc(x.f.JobNumber || '') + '</b></td><td><b>' + esc(x.f.Vendor || '') + '</b><div class="mut">' + esc(x.f.Subcontract || '') + (Number(x.f.Revision) > 1 ? ' · resubmitted' : '') + '</div></td>' +
        '<td>' + (x.f.PayRequest ? 'Pay App #' + esc(x.f.PayRequest) : '') + '</td><td class="num"><b>' + money(x.f.BalanceDue) + '</b></td><td>' + esc(rd) + '</td>' +
        '<td class="num"><button class="bh-btn-outline" data-openpay="' + esc(x.id) + '">Review</button></td></tr>';
    });
    el.innerHTML = '<div class="hub-mine-head"><div class="bh-section-title">Waiting for your approval <span class="bh-count">' + H.mine.length + '</span>' + (H.mineEvery ? ' <span class="mut">every job</span>' : '') + '</div>' +
      '<label class="hub-inline"><input type="checkbox" id="pay-notify"' + (H.notify ? ' checked' : '') + '> Email me when a pay app comes in</label></div>' +
      (H.mineNoJobs ? '<p class="bh-caption">Your PM jobs could not be read, so this shows only pay apps returned by you.</p>' : '') +
      (rows.length ? '<div class="bh-table-wrap"><table class="bh-table hub-table"><thead><tr><th>Job</th><th>Trade partner</th><th>Pay app</th><th class="num">Balance due</th><th>Received</th><th><span class="vh">Action</span></th></tr></thead><tbody>' + rows.join('') + '</tbody></table></div>'
        : '<p class="bh-caption">Nothing waiting for you. Portal pay apps on your jobs show here until you approve or return them.</p>');
  }
  async function saveNotify(on) {
    var m = me();
    try { await writeJson(PEOPLE_DIR + m.email + '.json', { email: m.email, payAppEmails: !!on, updatedAt: new Date().toISOString() }); H.notify = !!on;
      D().msg(on ? 'You will get an email when a pay app comes in on your jobs.' : 'Pay app emails are off. Check the Pay App Queue on the launcher or this tab instead.'); }
    catch (e) { D().msg(e.message, true); renderMine(); }
  }
  /* open one pay app: pick its job, its billing period, and select it (also #pay/item/<id> from the PM email) */
  async function openPay(id) {
    try {
      var Q = await queue(), r = await gcall('GET', '/sites/' + Q.sid + '/lists/' + Q.lid + '/items/' + encodeURIComponent(id) + '?$expand=fields');
      if (!r.ok) { D().msg('That pay app (AP-' + id + ') could not be found.', true); return; }
      var f = r.body.fields || {}, S = D().S;
      if (!S.projects || !S.projects.length) { D().msg('Connect Procore (top bar) to open ' + (f.JobNumber || 'the job') + '.', true); H.pendingId = id; return; }
      var pj = S.projects.find(function (p) { return normJob(p.project_number) === normJob(f.JobNumber); });
      if (!pj) { D().msg('Job ' + (f.JobNumber || '?') + ' is not in your Procore project list.', true); return; }
      H.pending = { id: String(id), period: f.BillingPeriod || null }; H.pendingId = null;
      if (location.hash.indexOf('#pay') !== 0 || /item/.test(location.hash)) history.replaceState(null, '', '#pay');
      var nav = document.querySelector('[data-view="pay"]'); if (nav && H.view !== 'pay') nav.click();
      var sel = $('#job');
      if (!S.project || String(S.project.id) !== String(pj.id)) { sel.value = String(pj.id); if (sel.onchange) await sel.onchange(); }
      else await loadPay(true);
      var top = $('#pay-when'); if (top && top.scrollIntoView) top.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) { D().msg(e.message, true); }
  }
  function applyPending() {
    if (!H.pending || !H.rows) return;
    var p = H.pending; H.pending = null;
    if (p.period && periodKeys().indexOf(p.period) >= 0) H.period = p.period;
    H.sel = H.rows.find(function (r) { return String(r.id) === p.id; }) || null;
  }

  /* ---------------- wiring ---------------- */
  document.addEventListener('click', function (e) {
    var t = e.target;
    if (t.closest && t.closest('a.hub-pclink')) return;   // the commitment number opens Procore in a new tab
    var br = t.closest && t.closest('.hub-buyrow'); if (br) { toggleBuy(br.dataset.key); return; }
    if (t.dataset && t.dataset.review) { H.sel = (H.rows || []).find(function (r) { return String(r.id) === t.dataset.review; }) || null; renderPay(); var rv = $('#pay-review'); if (rv && rv.scrollIntoView && window.innerWidth < 1100) rv.scrollIntoView({ behavior: 'smooth' }); return; }
    if (t.dataset && t.dataset.decide) { decide(t.dataset.decide); return; }
    if (t.dataset && t.dataset.openpay) { openPay(t.dataset.openpay); return; }
    if (t.dataset && t.dataset.file != null) { var fs = $('#pay-review')._files || []; if (fs[+t.dataset.file]) openFile(fs[+t.dataset.file]); return; }
    if (t.hasAttribute && t.hasAttribute('data-close')) { H.sel = null; renderPay(); return; }
    if (t.dataset && t.dataset.field) { insertField(t.dataset.field); return; }
    if (t.dataset && t.dataset.emview) { H.emText = t.dataset.emview === 'text'; renderPreview(); return; }
    if (t.id === 'em-save-tpl') saveTpl(false); else if (t.id === 'em-reset') { if (confirm('Put the default wording back in the company template?')) saveTpl(true); }
    else if (t.id === 'em-save-job') saveJob(); else if (t.id === 'em-test') sendTest(); else if (t.id === 'pay-refresh') loadPay(true);
  });
  document.addEventListener('keydown', function (e) { if (e.target.closest && e.target.closest('a.hub-pclink')) return; var br = e.target.closest && e.target.closest('.hub-buyrow'); if (br && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); toggleBuy(br.dataset.key); } });
  document.addEventListener('input', function (e) { if (/^em-(subject|body|note)$/.test(e.target.id)) { H.lastFocus = e.target.id === 'em-note' ? H.lastFocus : e.target; renderPreview(); } });
  document.addEventListener('focusin', function (e) { if (e.target.id === 'em-subject' || e.target.id === 'em-body') H.lastFocus = e.target; });
  document.addEventListener('change', function (e) { if (e.target.id === 'pay-notify') saveNotify(e.target.checked); if (e.target.id === 'pay-period') { H.period = e.target.value; H.sel = null; renderPay(); } if (e.target.id === 'em-as') renderPreview(); });

  window.BHHub = {
    job: function () { H.rows = null; H.sel = null; H.job = null; H.lines = {}; H.open = {}; H.period = null; summary(); if (H.view !== 'prime') this.show(H.view); },
    summary: summary,
    signedIn: function () { if (H.view === 'pay') loadMine(); },
    ready: function () { var m = location.hash.match(/^#pay\/item\/(\d+)/); if (m || H.pendingId) openPay(m ? m[1] : H.pendingId); },
    show: function (v) { H.view = v; if (v === 'pay' && D() && D().account()) loadMine(); if (!D() || !D().S.project) return; if (v === 'commitments') renderBuy().catch(function (e) { D().msg(e.message, true); }); else if (v === 'pay') loadPay(false); else if (v === 'email') loadEmail(); }
  };
})();
