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
        '<td><span class="hub-caret" aria-hidden="true">' + (open ? '▾' : '▸') + '</span><b>' + esc(c.number) + '</b></td><td>' + esc((c.vendor && (c.vendor.company || c.vendor.name)) || '') + '</td><td class="mut">' + esc(c.title || '') + '</td><td>' + esc(type) + '</td>' +
        '<td class="num">' + money(v) + '</td><td><span class="hub-pill ' + st[0] + '">' + esc(st[1]) + '</span></td><td class="hub-ck ' + ck.k + '">' + esc(ck.t) + '</td></tr>' +
        (open ? '<tr class="hub-lines"><td colspan="7">' + (lines === 'loading' ? '<span class="bh-caption">Loading cost codes…</span>' : lines && lines.error ? '<span class="hub-ck bad">' + esc(lines.error) + '</span>' : linesTable(lines || [])) + '</td></tr>' : '');
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
    if (last) { var lp = r2(Number(last.f.CompletedToDate || 0) - Number(last.f.RetainageHeld || 0)); if (Math.abs(lp - prev) > 1) warn('Previous payments ' + money(prev) + '; approved pay app #' + last.f.PayRequest + ' earned ' + money(lp) + '.'); else ok('Previous payments match approved pay app #' + last.f.PayRequest + '.'); }
    else warn('No earlier approved pay app on file to check previous payments against.');
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
    } catch (e) { el.innerHTML = '<div class="cd-msg is-bad">' + esc(e.message) + ' The Project Hub needs read and write access to the AP Desk site (the same grant the AP desk uses).</div>'; return; }
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
      '<table class="hub-fig">' + line('Original contract', f.OriginalContract) + line('Change orders', f.ChangeOrders) + line('Adjusted contract', f.AdjustedContract, 1) + line('Completed and stored to date', f.CompletedToDate) + line('Retainage', f.RetainageHeld) + line('Previous payments', f.PreviousPayments) + line('Due this period', f.BalanceDue, 1) + '</table>' +
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
      '<table style="border-collapse:collapse;font-size:14px">' + tr('Original contract', f.OriginalContract) + tr('Change orders', f.ChangeOrders) + tr('Adjusted contract', f.AdjustedContract) + tr('Completed and stored to date', f.CompletedToDate) + tr('Retainage', f.RetainageHeld) + tr('Previous payments', f.PreviousPayments) + '<tr><td style="padding:3px 16px 3px 0"><b>Due this period</b></td><td style="text-align:right"><b>' + money(f.BalanceDue) + '</b></td></tr></table>' +
      '<p><b>Checks</b></p><ul>' + ck.map(function (x) { return '<li>' + (x.k === 'ok' ? '' : x.k === 'bad' ? '<b>Issue:</b> ' : '<b>Check:</b> ') + esc(x.t) + '</li>'; }).join('') + '</ul>' +
      '<p><b>Approval history</b></p><ol>' + log.map(function (e) { return '<li>' + esc(e.s || '') + ' · ' + esc(e.at ? new Date(e.at).toLocaleString('en-US') : '') + ' · ' + esc(e.by || '') + (e.note ? '<br><span style="color:#484b4d">' + esc(e.note) + '</span>' : '') + '</li>'; }).join('') + '</ol></div>';
  }
  async function emailPayables(row, c, decision, note) {
    var pk = await packet(row);
    var html = payablesHtml(row, c, decision, note) + (pk.links.length ? '<p>Files too large to attach (open in SharePoint):</p><ul>' + pk.links.map(function (l) { return '<li>' + (l.url ? '<a href="' + esc(l.url) + '">' + esc(l.name) + '</a>' : esc(l.name) + ' (not found)') + '</li>'; }).join('') + '</ul>' : '');
    var subj = 'PM approved: ' + (row.f.Vendor || (c.vendor && (c.vendor.company || c.vendor.name)) || '') + ' ' + row.f.Subcontract + (row.f.PayRequest ? ' pay app ' + row.f.PayRequest : '') + ' · ' + money(row.f.BalanceDue);
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
        var r = await gcall('PATCH', '/sites/' + Q.sid + '/lists/' + Q.lid + '/items/' + encodeURIComponent(row.id) + '/fields', fields);
        if (!r.ok) throw new Error('Could not save the decision (' + r.status + (r.body && r.body.error ? ': ' + r.body.error.message : '') + ').');
        Object.assign(row.f, fields);
      }
      if (kind === 'return') {
        var to = (parseLog(row).map(function (e) { var x = String(e.by || '').match(/<([^>]+@[^>]+)>/); return x && x[1]; }).find(Boolean)) || '';
        if (to && !/baileyharris\.com$/i.test(to)) {
          try { await sendMail([to], 'Pay app returned: ' + row.f.Subcontract + (row.f.PayRequest ? ' pay app ' + row.f.PayRequest : ''), '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px"><p>Your pay application for subcontract ' + esc(row.f.Subcontract) + ' needs a correction before it can be paid:</p><p><b>' + esc(note) + '</b></p><p>Please resubmit it with your pay app link, or reply to ' + esc(m.email) + ' with questions.</p><p>Thank you,<br>Bailey-Harris Construction</p></div>', [], [m.email]);
            D().msg('Returned. ' + to + ' was emailed the note.'); }
          catch (e) { D().msg('Returned and saved, but the email to ' + to + ' failed: ' + e.message, true); }
        } else D().msg('Returned and saved. No trade partner email on this pay app: send them the note yourself.', true);
      } else {
        try { await emailPayables(row, c, kind === 'noted' ? 'Approved as noted' : 'Approved', note); D().msg('Approved and emailed to payables@.'); }
        catch (e) { D().msg('Approved and saved, but the email to payables@ failed: ' + e.message + '. Use "Email payables@ again".', true); }
      }
      renderPay();
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

  /* ---------------- wiring ---------------- */
  document.addEventListener('click', function (e) {
    var t = e.target;
    var br = t.closest && t.closest('.hub-buyrow'); if (br) { toggleBuy(br.dataset.key); return; }
    if (t.dataset && t.dataset.review) { H.sel = (H.rows || []).find(function (r) { return String(r.id) === t.dataset.review; }) || null; renderPay(); var rv = $('#pay-review'); if (rv && rv.scrollIntoView && window.innerWidth < 1100) rv.scrollIntoView({ behavior: 'smooth' }); return; }
    if (t.dataset && t.dataset.decide) { decide(t.dataset.decide); return; }
    if (t.dataset && t.dataset.file != null) { var fs = $('#pay-review')._files || []; if (fs[+t.dataset.file]) openFile(fs[+t.dataset.file]); return; }
    if (t.hasAttribute && t.hasAttribute('data-close')) { H.sel = null; renderPay(); return; }
    if (t.dataset && t.dataset.field) { insertField(t.dataset.field); return; }
    if (t.dataset && t.dataset.emview) { H.emText = t.dataset.emview === 'text'; renderPreview(); return; }
    if (t.id === 'em-save-tpl') saveTpl(false); else if (t.id === 'em-reset') { if (confirm('Put the default wording back in the company template?')) saveTpl(true); }
    else if (t.id === 'em-save-job') saveJob(); else if (t.id === 'em-test') sendTest(); else if (t.id === 'pay-refresh') loadPay(true);
  });
  document.addEventListener('keydown', function (e) { var br = e.target.closest && e.target.closest('.hub-buyrow'); if (br && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); toggleBuy(br.dataset.key); } });
  document.addEventListener('input', function (e) { if (/^em-(subject|body|note)$/.test(e.target.id)) { H.lastFocus = e.target.id === 'em-note' ? H.lastFocus : e.target; renderPreview(); } });
  document.addEventListener('focusin', function (e) { if (e.target.id === 'em-subject' || e.target.id === 'em-body') H.lastFocus = e.target; });
  document.addEventListener('change', function (e) { if (e.target.id === 'pay-period') { H.period = e.target.value; H.sel = null; renderPay(); } if (e.target.id === 'em-as') renderPreview(); });

  window.BHHub = {
    job: function () { H.rows = null; H.sel = null; H.job = null; H.lines = {}; H.open = {}; H.period = null; summary(); if (H.view !== 'prime') this.show(H.view); },
    summary: summary,
    show: function (v) { H.view = v; if (!D() || !D().S.project) return; if (v === 'commitments') renderBuy().catch(function (e) { D().msg(e.message, true); }); else if (v === 'pay') loadPay(false); else if (v === 'email') loadEmail(); }
  };
})();
