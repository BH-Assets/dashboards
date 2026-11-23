/* cloud.js - puts John Barron's equipment app on the Equipment service (19 Nov 2026).
 *
 * Loaded AFTER app.js (John's code, lightly patched by port.py). It replaces only the parts that kept data
 * in one browser: loading, saving, sign-in, users, and the request send. Every screen is John's.
 *
 *   Load    - Microsoft sign-in (the BH Dashboard app), then GET /api/data from the Equipment Worker.
 *   Save    - John's save() is called after every change. Here it compares each record with the copy it
 *             loaded and sends only what changed (Admin only). Photos and files added in the app are
 *             uploaded first; the record keeps the file's path, not the picture itself.
 *   Roles   - from the Worker: Admin (John, Tiffany), Management (Clay), Super/PM (on a Procore job team),
 *             Viewer (everyone else). Super/PM, Management and Viewer use John's superintendent screens.
 *   Crew requests - field crews use /request/. Approvers see them under "Crew requests" here.
 */
(function () {
  'use strict';
  var API = 'https://equipment-api.baileyharris.com';
  var TENANT_ID = '848aa890-7982-4028-befa-3323d82583ce', CLIENT_ID = '1825a1c0-1ce3-4078-aeb7-a4714562c422', SCOPES = ['openid', 'profile', 'email'];
  var COLLS = { items: 'id', stock: 'id', inventory: 'id', orders: 'no', requests: 'id', customers: 'id', locations: 'id', maintenance: 'id', custody: 'id', bundles: 'id' };
  var CL = window.CL = { role: null, jobs: [], user: null, etags: {}, base: {}, jobFeed: [], saving: false, queued: false, timer: null, conflicts: [] };

  /* ---------------- sign-in ---------------- */
  var pca, acct;
  async function signIn() {
    pca = new msal.PublicClientApplication({ auth: { clientId: CLIENT_ID, authority: 'https://login.microsoftonline.com/' + TENANT_ID, redirectUri: location.origin + '/equipment/', navigateToLoginRequestUrl: false }, cache: { cacheLocation: 'localStorage' } });
    await pca.initialize(); await pca.handleRedirectPromise();
    acct = pca.getAllAccounts()[0];
    if (!acct) { await pca.loginRedirect({ scopes: SCOPES }); return false; }
    pca.setActiveAccount(acct); return true;
  }
  /* ID tokens last about an hour, and MSAL can hand back a cached one that has already expired, which the Worker
     then refuses ("Microsoft sign-in not accepted"). So: refresh it when it has under 5 minutes left; if the silent
     refresh fails, try single sign-on in the background; only if that fails too, ask the person to sign in again,
     with a button, so nothing on the screen is lost. (23 Nov 2026) */
  var signInWait = null;
  var expMs = function (r) { return r && r.idTokenClaims && r.idTokenClaims.exp ? r.idTokenClaims.exp * 1000 : 0; };
  async function idToken(force) {
    try {
      var r = await pca.acquireTokenSilent({ scopes: SCOPES, account: acct, forceRefresh: !!force });
      if (!force && expMs(r) && expMs(r) - Date.now() < 5 * 60e3) return idToken(true);
      if (expMs(r) && expMs(r) < Date.now() + 60e3) throw new Error('stale id token');
      return r.idToken;
    } catch (e) {
      try { var s = await pca.ssoSilent({ scopes: SCOPES, loginHint: acct && acct.username }); acct = s.account || acct; pca.setActiveAccount(acct); return s.idToken; }
      catch (e2) { return askToSignIn(); }
    }
  }
  function askToSignIn() {
    if (signInWait) return signInWait;
    signInWait = new Promise(function (resolve) {
      showBanner('Your Microsoft sign-in has expired. Sign in again to keep saving; your changes stay on this screen.', true);
      var b = document.querySelector('#cl-banner');
      b.innerHTML = b.innerHTML.replace(/<button[^>]*data-cl="reload"[^>]*>[^<]*<\/button>/, '') + ' <button class="btn sm primary" data-cl="signin">Sign in again</button>';
      CL.resumeSignIn = resolve;
    });
    return signInWait;
  }
  // the button: a popup keeps this page (and anything not saved yet) as it is
  document.addEventListener('click', function (e) {
    if (!e.target.closest('[data-cl="signin"]')) return;
    pca.acquireTokenPopup({ scopes: SCOPES, account: acct, loginHint: acct && acct.username }).then(function (r) {
      acct = r.account || acct; pca.setActiveAccount(acct); var b = document.querySelector('#cl-banner'); if (b) b.remove();
      var done = CL.resumeSignIn; signInWait = null; CL.resumeSignIn = null; if (done) done(r.idToken);
    }).catch(function () { toast('Sign-in did not finish. Allow pop-ups for dashboard.baileyharris.com, then try again.'); });
  });
  async function api(path, opt, retried) {
    opt = opt || {};
    var headers = Object.assign({ Authorization: 'Bearer ' + await idToken(!!retried) }, opt.json !== undefined ? { 'Content-Type': 'application/json' } : {}, opt.headers || {});
    var r = await fetch(API + path, { method: opt.method || 'GET', headers: headers, body: opt.json !== undefined ? JSON.stringify(opt.json) : opt.body });
    if (r.status === 401 && !retried) return api(path, opt, true);   // token refused: get a fresh one and try once more
    var t = await r.text(), j = null; try { j = JSON.parse(t); } catch (e) { }
    if (!r.ok && r.status !== 207) { var er = new Error((j && j.error) || ('The Equipment service answered ' + r.status)); er.status = r.status; er.body = j; throw er; }
    return j;
  }
  CL.api = api;

  /* ---------------- records <-> the lists ---------------- */
  function rows() {
    var out = {};
    Object.keys(COLLS).forEach(function (c) { (S[c] || []).forEach(function (r) { out[c + ':' + r[COLLS[c]]] = { coll: c, key: String(r[COLLS[c]]), rec: r }; }); });
    var st = Object.assign({}, S.settings); delete st.nextSR; delete st.nextRQ;
    out['settings:settings'] = { coll: 'settings', key: 'settings', rec: st };
    out['settings:groups'] = { coll: 'settings', key: 'groups', rec: S.groups };
    out['settings:counters'] = { coll: 'settings', key: 'counters', rec: { nextOrder: S.nextOrder, nextSR: S.settings.nextSR, nextCheckout: S.nextCheckout } };
    out['settings:flags'] = { coll: 'settings', key: 'flags', rec: { v: S.v, groupsByNumber: S.groupsByNumber, yardRepair1: S.yardRepair1, invToShop1: S.invToShop1 } };
    return out;
  }
  function snapshot() { var r = rows(), b = {}; Object.keys(r).forEach(function (k) { b[k] = JSON.stringify(r[k].rec); }); return b; }

  /* files: signed links from the Worker are dropped before saving (the record keeps .file / .photoFile);
     pictures added in the app (data: URIs) are uploaded and replaced by their path */
  var safe = function (s) { return String(s || 'file').replace(/[^A-Za-z0-9._-]+/g, '_').slice(0, 80); };
  var EXT = { 'image/jpeg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'image/gif': '.gif', 'image/svg+xml': '.svg', 'application/pdf': '.pdf' };
  async function upload(path, dataUrl) {
    var blob = dataUrlToBlob(dataUrl);   // not fetch(): the page CSP blocks fetching data: URIs
    if (!/\.[a-z0-9]{2,5}$/i.test(path)) path += EXT[blob.type] || '';
    var j = await api('/api/upload?path=' + encodeURIComponent(path), { method: 'POST', body: blob, headers: { 'Content-Type': 'application/octet-stream' } });
    return j.path;
  }
  var isLink = function (v) { return typeof v === 'string' && v.indexOf(API + '/f/') === 0; };
  var isData = function (v) { return typeof v === 'string' && v.indexOf('data:') === 0; };
  async function fileList(coll, id, arr) {
    for (var i = 0; arr && i < arr.length; i++) {
      var f = arr[i]; if (!f) continue;
      if (isLink(f.data)) delete f.data;
      else if (isData(f.data)) { f.file = await upload(coll + '/' + safe(id) + '/' + safe(f.id || Date.now()) + '-' + safe(f.name), f.data); delete f.data; }
    }
  }
  async function dehydrate(coll, rec) {
    var r = JSON.parse(JSON.stringify(rec));
    if (coll === 'items') { await fileList('items', r.id, r.photos); await fileList('items', r.id, r.files); }
    if (coll === 'maintenance') await fileList('maintenance', r.id, r.files);
    if (coll === 'stock' || coll === 'inventory') {
      if (isLink(r.photo)) delete r.photo;
      else if (isData(r.photo)) { r.photoFile = await upload(coll + '/' + safe(r.id) + '/photo-' + Date.now(), r.photo); delete r.photo; }
      else if (!r.photo) delete r.photoFile;
    }
    if (coll === 'settings' && rec === S.settings) { /* not used: settings row is a copy */ }
    if (coll === 'settings' && r && 'company' in r) {
      if (isLink(r.logo)) delete r.logo;
      else if (isData(r.logo)) { r.logoFile = await upload('settings/logo-' + Date.now(), r.logo); delete r.logo; }
      else if (!r.logo) delete r.logoFile;
    }
    return r;
  }

  async function flush() {
    if (CL.role !== 'Admin') return;
    if (CL.saving) { CL.queued = true; return; }
    CL.saving = true; setSaveState('Saving…');
    try {
      var cur = rows(), changes = [], seen = {};
      Object.keys(cur).forEach(function (k) {
        var str = JSON.stringify(cur[k].rec); seen[k] = 1;
        if (CL.base[k] !== str) changes.push({ k: k, str: str, coll: cur[k].coll, key: cur[k].key, op: 'upsert', rec: cur[k].rec });
      });
      Object.keys(CL.base).forEach(function (k) { if (!seen[k]) { var i = k.indexOf(':'); changes.push({ k: k, coll: k.slice(0, i), key: k.slice(i + 1), op: 'delete' }); } });
      for (var n = 0; n < changes.length; n += 40) {
        var chunk = changes.slice(n, n + 40), send = [];
        for (var c of chunk) send.push({ coll: c.coll, key: c.key, op: c.op, etag: CL.etags[c.k] || '', rec: c.op === 'delete' ? undefined : await dehydrate(c.coll, c.rec) });
        var res = await api('/api/save', { method: 'POST', json: { changes: send } });
        var failed = {}; (res.failed || []).forEach(function (f) { failed[f.coll + ':' + f.key] = f; });
        chunk.forEach(function (c) {
          if (failed[c.k]) return;
          if (c.op === 'delete') { delete CL.base[c.k]; delete CL.etags[c.k]; }
          else { CL.base[c.k] = c.str; if (res.etags && res.etags[c.k]) CL.etags[c.k] = res.etags[c.k]; }
        });
        var bad = res.failed || [];
        if (bad.length) {
          CL.conflicts = CL.conflicts.concat(bad);
          showBanner(bad.some(function (f) { return f.conflict; })
            ? 'Someone else saved ' + bad.length + ' of these records before you. Reload to get their changes; your change to those records was not saved.'
            : 'Some changes were not saved: ' + bad.map(function (f) { return f.error; }).join('; '), true);
        }
      }
      setSaveState(CL.conflicts.length ? 'Not all saved' : 'Saved');
    } catch (e) {
      setSaveState('Not saved'); showBanner('Could not save: ' + e.message + ' Your changes are still on this screen. Check your connection and make any change to try again.', true);
    } finally {
      CL.saving = false;
      if (CL.queued) { CL.queued = false; flush(); }
    }
  }

  /* ---------------- load ---------------- */
  async function load() {
    var d = await api('/api/data');
    CL.role = d.role; CL.jobs = d.jobs || []; CL.etags = d.etags || {}; CL.jobFeed = (d.jobFeed && d.jobFeed.jobs) || [];
    CL.user = { id: d.email, email: d.email, name: d.name, role: d.role === 'Admin' ? 'Admin' : 'Superintendent', cloudRole: d.role, active: true, title: '', phone: '' };
    var s = Object.assign({}, d.settings || {}), cn = d.counters || {};
    s.nextSR = Math.max(+cn.nextSR || 1001, 1 + Math.max.apply(null, [0].concat((d.maintenance || []).map(function (m) { return +m.no || 0; }))));
    s.nextRQ = +cn.nextRQ || 1001;
    CL.changelog = d.changelog || null;
    var f = d.flags || {};
    S = { v: f.v || 1, nextOrder: Math.max(+cn.nextOrder || 1001, 1 + Math.max.apply(null, [0].concat((d.orders || []).map(function (o) { return +o.no || 0; })))),
      nextCheckout: +cn.nextCheckout || 1001, settings: s, groups: d.groups || [], groupsByNumber: f.groupsByNumber, yardRepair1: f.yardRepair1, invToShop1: f.invToShop1,
      users: [], items: d.items || [], stock: d.stock || [], inventory: d.inventory || [], orders: d.orders || [], requests: d.requests || [], customers: d.customers || [],
      locations: d.locations || [], maintenance: d.maintenance || [], custody: d.custody || [], bundles: d.bundles || [] };
    // keep every "where it's kept" link and number rule of John's app
    migrate(S);
    S.users = [CL.user];
    ROLES.Superintendent.perms = CL.role === 'Viewer' ? [] : ['request'];
    ROLES.Superintendent.desc = { 'Super/PM': 'Superintendent or PM on an active job. Sees the whole catalog with rates, requests equipment, stock and inventory, and approves crew requests for their jobs.',
      Management: 'Self Perform management. Super/PM, plus approves Self Perform crew requests.', Viewer: 'Sees the catalog and rates. Requests are for superintendents and PMs on an active job.' }[CL.role] || ROLES.Superintendent.desc;
    CL.base = snapshot();
    CL.loadedAt = new Date();
    lockScreens();
  }
  CL.reload = async function () { await load(); render(); };

  /* ---------------- John's functions, re-pointed ---------------- */
  me = function () { return CL.user; };
  setSession = function (id) { if (!id && pca) pca.logoutRedirect({ account: acct, postLogoutRedirectUri: location.origin + '/' }); };
  authGate = function () {
    if (!CL.user) return false;
    var a = document.querySelector('#auth-root'); if (a) a.innerHTML = ''; document.querySelector('.app').hidden = false; return true;
  };
  idbSave = function () { };
  writeFile = async function () { };
  save = function () {
    (S.items || []).map(function (i) { return i.category; }).filter(function (c, i, a) { return a.indexOf(c) === i; })
      .forEach(function (c) { if (!S.groups.some(function (g) { return g.name === c; })) S.groups.push({ name: c, desc: seriesDesc(c) }); });
    updateBell();
    if (CL.role !== 'Admin') return;
    setSaveState('Saving…'); clearTimeout(CL.timer); CL.timer = setTimeout(flush, 350);
  };
  fileSave = function () { toast('Everything saves to SharePoint automatically.'); };
  fileOpen = function () { toast('Data now lives in SharePoint. Opening a data file is turned off.'); };
  var dl0 = downloadData;
  downloadData = function (name, data) {
    if (!isLink(data)) return dl0(name, data);
    fetch(data).then(function (r) { if (!r.ok) throw new Error(r.status); return r.blob(); }).then(function (b) { saveFile(name, b); })
      .catch(function () { toast('Could not open that file. Reload the page and try again.'); });
  };

  // requests: signed-in people send through the Worker (it routes and numbers them)
  reqSend = async function () {
    var fd = Object.fromEntries(new FormData(document.querySelector('#rqform')));
    if (!fd.job) { toast('Pick the job'); return; }
    if (!fd.delivery) { toast('Pick delivery or pickup'); return; }
    if (fd.returnBy && fd.needBy && (fd.returnBy + 'T' + (fd.returnTime || '23:59')) <= (fd.needBy + 'T' + (fd.needTime || '00:00'))) { toast("Return date is before the date it's needed"); return; }
    var btn = document.querySelector('[data-act="rq-send"]'); if (btn) btn.disabled = true;
    try {
      var j = await api('/api/request', { method: 'POST', json: { job: fd.job, needBy: fd.needBy, needTime: fd.needTime, returnBy: fd.returnBy, returnTime: fd.returnTime, delivery: fd.delivery, deliveryNotes: fd.deliveryNotes, notes: fd.notes, lines: ui.req.cart } });
      ui.req.cart = []; saveCart(); localStorage.setItem(KEY + '-lastjob', fd.job);
      await load();
      openModal('<header><h2>' + rqNo(j.request) + ' sent</h2><button class="x" data-act="close" aria-label="Close">×</button></header><div class="mbody"><p>Your request went straight to the Equipment Manager. You\'ll see his answer under <b>My requests</b>.</p></div><footer><button class="btn" data-act="close">Request more</button><button class="btn primary" data-act="goto" data-to="my-requests">See my requests</button></footer>', 'narrow');
      rerender();
    } catch (e) { toast(e.message); if (btn) btn.disabled = false; }
  };
  var ra0 = reqAction;
  reqAction = function (a, ds) {
    if (a === 'rq-cancel' && CL.role !== 'Admin') {
      api('/api/decide', { method: 'POST', json: { id: ds.id, etag: CL.etags['requests:' + ds.id], action: 'cancel' } })
        .then(function () { return CL.reload(); }).then(function () { toast('Request cancelled'); }).catch(function (e) { toast(e.message); });
      return true;
    }
    return ra0(a, ds);
  };

  /* ---------------- crew requests waiting on someone ---------------- */
  function canApprove(r) {
    if (r.status !== 'Awaiting approval') return false;
    if (CL.role === 'Admin') return true;
    var a = r.approvers || {};
    if (a.type === 'management') return CL.role === 'Management';
    return (a.emails || []).indexOf(CL.user.email) >= 0;
  }
  function waiting() { return (S.requests || []).filter(canApprove).sort(function (a, b) { return a.no - b.no; }); }
  CL.waiting = waiting;
  var AP = {};   // id -> working copy of the lines while someone edits
  function apLines(r) { return AP[r.id] || (AP[r.id] = r.lines.map(function (l) { return Object.assign({}, l); })); }
  function vApprovals() {
    var list = waiting(), all = (S.requests || []).filter(function (r) { return r.source === 'crew' && r.status !== 'Awaiting approval' && (CL.role === 'Admin' || (r.approvers && ((r.approvers.emails || []).indexOf(CL.user.email) >= 0 || (r.approvers.type === 'management' && CL.role === 'Management')))); })
      .sort(function (a, b) { return b.no - a.no; }).slice(0, 15);
    var card = function (r) {
      var L = locById(r.locationId), lines = apLines(r);
      return '<div class="rq-req" data-ap="' + r.id + '"><div class="rq-req-h"><b class="num">' + rqNo(r) + '</b><span class="chip neutral">Waiting on you</span><span><b>' + esc(r.byName) + '</b> · ' + esc(fmtPhone(r.byPhone)) + '</span>' + (r.selfPerform ? '<span class="chip neutral">Self Perform</span>' : '') + '<span class="muted small">' + fmtDate(r.createdAt) + '</span></div>' +
        '<div class="rq-req-meta"><span>' + ico('locations') + ' ' + esc(locLabel(L)) + '</span></div>' +
        '<div class="form" style="margin:8px 0"><div class="field"><label>Needed by</label><input type="date" data-apf="needBy" value="' + esc(r.needBy || '') + '"></div><div class="field"><label>Return by</label><input type="date" data-apf="returnBy" value="' + esc(r.returnBy || '') + '"></div>' +
        '<div class="field"><label>Delivery or pickup</label><select data-apf="delivery"><option ' + (r.delivery === 'Delivery' ? 'selected' : '') + '>Delivery</option><option ' + (r.delivery === 'Pickup' ? 'selected' : '') + '>Pickup</option></select></div>' +
        '<div class="field full"><label>Delivery details</label><input data-apf="deliveryNotes" value="' + esc(r.deliveryNotes || '') + '"></div>' +
        '<div class="field full"><label>Notes</label><textarea data-apf="notes">' + esc(r.notes || '') + '</textarea></div></div>' +
        '<div class="rq-lines">' + lines.map(function (l, i) { return '<div class="rq-line"><b>' + esc(reqLineName(l)) + '</b><input type="number" min="0.01" step="any" class="sm-input num" style="width:80px" data-apq="' + i + '" value="' + l.qty + '" aria-label="Quantity"> <span class="small muted">' + esc(reqLineUnit(l)) + '</span><span class="small ' + (reqLineAvail(l).ok ? 'muted' : '') + '">' + esc(reqLineAvail(l).txt) + '</span><button class="del" data-act="ap-line-del" data-id="' + r.id + '" data-i="' + i + '" aria-label="Remove">×</button></div>'; }).join('') + '</div>' +
        '<div class="field full" style="margin-top:8px"><label>Note to the crew (needed to decline)</label><input data-apf="note" placeholder="Why, or what to use instead"></div>' +
        '<div class="rq-req-f"><button class="btn" data-act="ap-deny" data-id="' + r.id + '">Decline</button><button class="btn" data-act="ap-save" data-id="' + r.id + '">Save changes</button><button class="btn primary" data-act="ap-approve" data-id="' + r.id + '">Approve and send to the Equipment Manager</button></div></div>';
    };
    return band('Crew requests', 'Stock and inventory requests from field crews on your jobs. Change what you need to, then approve or decline. Approved requests go to the Equipment Manager.', '') +
      '<div class="content"><div class="rq-list">' + (list.map(card).join('') || '<div class="empty"><strong>Nothing waiting on you</strong>When a crew on one of your jobs sends a request, it shows up here.</div>') + '</div>' +
      (all.length ? '<div class="subhead" style="margin-top:20px"><span>Recently decided</span></div><div class="rq-list">' + all.map(function (r) { return reqCardHtml(r, false); }).join('') + '</div>' : '') + '</div>';
  }
  var fmtPhone = function (p) { p = String(p || ''); return p.length === 10 ? '(' + p.slice(0, 3) + ') ' + p.slice(3, 6) + '-' + p.slice(6) : p; };
  VIEWS.approvals = vApprovals;
  SUPER_ROUTES.push('approvals');
  async function decide(id, action) {
    var box = document.querySelector('[data-ap="' + id + '"]'), r = (S.requests || []).find(function (x) { return x.id === id; }); if (!box || !r) return;
    var v = function (f) { var el = box.querySelector('[data-apf="' + f + '"]'); return el ? el.value : undefined; };
    var lines = apLines(r).filter(function (l) { return +l.qty > 0; });
    if (!lines.length && action !== 'deny') { toast('Leave at least one item, or decline the request'); return; }
    if (action === 'deny' && !String(v('note') || '').trim()) { toast('Add a short note so the crew knows why'); box.querySelector('[data-apf="note"]').focus(); return; }
    box.querySelectorAll('button').forEach(function (b) { b.disabled = true; });
    try {
      await api('/api/decide', { method: 'POST', json: { id: id, etag: CL.etags['requests:' + id], action: action, note: v('note'), lines: lines, needBy: v('needBy'), returnBy: v('returnBy'), delivery: v('delivery'), deliveryNotes: v('deliveryNotes'), notes: v('notes') } });
      delete AP[id]; await CL.reload();
      toast(rqNo(r) + (action === 'approve' ? ' approved and sent to the Equipment Manager' : action === 'deny' ? ' declined' : ' saved'));
    } catch (e) { toast(e.message); box.querySelectorAll('button').forEach(function (b) { b.disabled = false; }); if (e.status === 409) CL.reload(); }
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest('[data-act^="ap-"]'); if (!b) return;
    e.preventDefault(); e.stopPropagation();
    var id = b.dataset.id, r = (S.requests || []).find(function (x) { return x.id === id; });
    if (b.dataset.act === 'ap-line-del' && r) { apLines(r).splice(+b.dataset.i, 1); rerender(); return; }
    if (b.dataset.act === 'ap-approve') decide(id, 'approve');
    if (b.dataset.act === 'ap-deny') decide(id, 'deny');
    if (b.dataset.act === 'ap-save') decide(id, 'edit');
  }, true);
  document.addEventListener('input', function (e) {
    var q = e.target.closest('[data-apq]'); if (!q) return;
    var id = q.closest('[data-ap]').dataset.ap, r = (S.requests || []).find(function (x) { return x.id === id; });
    if (r) apLines(r)[+q.dataset.apq].qty = +q.value || 0;
  }, true);

  /* navigation: add "Crew requests" for anyone who can approve */
  var mayApprove = function () { return CL.role === 'Admin' || CL.role === 'Management' || CL.role === 'Super/PM'; };
  var route0 = route;
  route = function () { var h = location.hash.replace('#', ''); if (h === 'approvals' && mayApprove()) return 'approvals'; if (h === 'changelog' && CL.role === 'Admin') return 'changelog'; return route0(); };
  var nav0 = navList;
  navList = function () {
    var n = nav0();
    if (CL.role === 'Viewer') return [['request', 'Catalog']];
    if (!mayApprove()) return n;
    var w = waiting().length, item = ['approvals', 'Crew requests' + (w ? ' (' + w + ')' : '')];
    var i = n.findIndex(function (x) { return x[0] === 'requests' || x[0] === 'my-requests'; });
    n.splice(i < 0 ? n.length : i + 1, 0, item); return n;
  };
  var bell0 = bellCount;
  bellCount = function () { return bell0() + (mayApprove() ? waiting().length : 0); };
  var more0 = openMore;
  openMore = function () {
    more0();
    if (isSuper() && mayApprove()) {
      var g = document.querySelector('.m-sheet'); if (!g) return;
      var w = waiting().length, d = document.createElement('div'); d.className = 'm-group';
      d.innerHTML = '<a class="m-row" href="#approvals"><span class="m-row-ic">' + ico('request') + '</span><span class="m-row-txt"><b>Crew requests' + (w ? ' (' + w + ')' : '') + '</b><small>Approve or decline requests from field crews</small></span><span class="m-chev">›</span></a>';
      g.insertBefore(d, g.firstChild);
    }
  };

  /* profile menu: Microsoft account, no passwords */
  toggleProfileMenu = function () {
    var m = document.querySelector('#pmenu'); if (m) { m.remove(); return; }
    var u = me(); if (!u) return;
    m = document.createElement('div'); m.id = 'pmenu'; m.className = 'pmenu'; m.setAttribute('role', 'menu');
    m.innerHTML = '<div class="pm-head">' + avatarHtml(u, 'big') + '<div style="min-width:0"><b>' + esc(u.name) + '</b><div class="small muted trunc" style="max-width:190px">' + esc(u.email) + '</div>' + chipFor(CL.role) + '</div></div>' +
      (CL.role === 'Admin' ? '<button role="menuitem" data-act="goto" data-to="users">Manage users</button>' : '') +
      (CL.role === 'Admin' ? '<button role="menuitem" data-act="theme">Switch light / dark</button>' : '') + '<button role="menuitem" data-act="logout" class="pm-out">Sign out</button>';
    document.body.appendChild(m);
    var r = document.querySelector('#avatar-btn').getBoundingClientRect(), narrow = matchMedia('(max-width:760px)').matches;
    m.style.left = (narrow ? Math.max(8, r.left) : r.right + 8) + 'px'; m.style.top = (narrow ? r.bottom + 6 : r.top) + 'px';
  };
  var ACCESS = null, ACCESS_ERR = '';
  function loadAccess() {
    api('/api/access').then(function (a) { ACCESS = a; ACCESS_ERR = ''; if (route() === 'users') rerender(); })
      .catch(function (e) { ACCESS_ERR = e.message; if (route() === 'users') rerender(); });
  }
  VIEWS.users = function () {
    if (CL.role !== 'Admin') return band('Manage users', '', '') + '<div class="content"><div class="empty"><strong>Admins only</strong></div></div>';
    if (!ACCESS && !ACCESS_ERR) loadAccess();
    var nm = function (e) { var u = ACCESS && ACCESS.superPm.find(function (p) { return p.email === e; }); return u ? u.name : (CL.user && CL.user.email === e ? CL.user.name : ''); };
    var who = function (arr) { return arr && arr.length ? arr.map(function (e) { var n = nm(e); return (n ? '<b>' + esc(n) + '</b> ' : '') + '<span class="muted">' + esc(e) + '</span>'; }).join('<br>') : '<span class="muted">Nobody set</span>'; };
    var roles = '<table class="tbl"><thead><tr><th>Role</th><th>Who</th><th>What they can do</th></tr></thead><tbody>' +
      '<tr><td><b>Admin</b></td><td>' + (ACCESS ? who(ACCESS.admins) : '…') + '</td><td>Everything: catalog, rates, orders, check-outs, the shop, requests, setup.</td></tr>' +
      '<tr><td><b>Management</b></td><td>' + (ACCESS ? who(ACCESS.management) : '…') + '</td><td>Everything Super/PM can, and approves Self Perform crew requests.</td></tr>' +
      '<tr><td><b>Super/PM</b></td><td>' + (ACCESS ? ACCESS.superPm.length + ' people on active and warranty jobs in Procore (listed below)' + (ACCESS.extras.length ? ', plus ' + ACCESS.extras.length + ' added by hand' : '') : '…') + '</td><td>Sees the whole catalog with rates, requests equipment, stock and inventory (straight to the Equipment Manager), and approves crew requests for their jobs.</td></tr>' +
      '<tr><td><b>Viewer</b></td><td>Everyone else at Bailey-Harris</td><td>Sees the catalog, rates and what is available. Cannot request.</td></tr>' +
      '<tr><td><b>Field crews</b></td><td>No sign-in</td><td>Request stock and inventory at dashboard.baileyharris.com/request with a job number.</td></tr></tbody></table>';
    var people = !ACCESS ? '<p class="muted">' + (ACCESS_ERR ? esc(ACCESS_ERR) : 'Loading…') + '</p>' :
      '<table class="tbl"><thead><tr><th>Name</th><th>Email</th><th>Jobs</th></tr></thead><tbody>' +
      ACCESS.superPm.map(function (p) { return '<tr><td><b>' + esc(p.name) + '</b></td><td class="muted">' + esc(p.email) + '</td><td>' + p.jobs.map(function (j) { return esc(j.job) + ' <span class="muted small">(' + esc(j.as) + ')</span>'; }).join(', ') + '</td></tr>'; }).join('') +
      ACCESS.extras.map(function (e) { return '<tr><td>' + esc(nm(e) || e) + '</td><td class="muted">' + esc(e) + '</td><td class="muted">Added by hand (EqAccessExtras)</td></tr>'; }).join('') +
      '</tbody></table><p class="small muted" style="margin-top:8px">From Procore, refreshed nightly' + (ACCESS.feedAt ? ' (last ' + esc(new Date(ACCESS.feedAt).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' })) + ')' : '') + '.</p>';
    return band('Manage users', 'Managed by Microsoft 365. Everyone signs in with their Bailey-Harris Microsoft account; there are no passwords in this app.', '<button class="btn" data-cl-access="1">Refresh</button>') +
      '<div class="content"><div class="panel"><h3>Who has access to what</h3><div class="body">' + roles + '</div></div>' +
      '<div class="panel" style="margin-top:16px"><h3>Superintendents and PMs' + (ACCESS ? ' (' + ACCESS.superPm.length + ')' : '') + '</h3><div class="body">' + people + '</div></div>' +
      '<div class="panel" style="margin-top:16px"><h3>Making a change</h3><div class="body small">' +
      '<p><b>Super/PM</b> follows the Procore project team: add or remove someone on the job\'s team in Procore (Project Management Team, Field Management Team or a Superintendent role) and it changes overnight.</p>' +
      '<p><b>Someone not on a Procore team</b> (an estimator, the office): add their email to the EqAccessExtras list on the Equipment SharePoint site.</p>' +
      '<p><b>Admin and Management</b> are set on the Equipment service by Tiffany (ADMIN_EMAILS, MGMT_EMAILS).</p>' +
      '<p><b>Someone who leaves</b> loses access when their Microsoft 365 account is turned off.</p></div></div></div>';
  };
  document.addEventListener('click', function (e) { if (e.target.closest('[data-cl-access]')) { ACCESS = null; ACCESS_ERR = ''; rerender(); } });

  /* Only Admins (John, Tiffany) change anything or customise screens (Tiffany, 20 Nov 2026).
     Everyone else: no light/dark switch, no column resizing, no Setup button, no saved layouts.
     Admins keep their dashboard layout per person in this browser. */
  function lockScreens() {
    var admin = CL.role === 'Admin';
    document.body.classList.toggle('cl-locked', !admin);
    ACT_PERM.theme = admin ? undefined : 'admin';
    if (!admin) { delete document.documentElement.dataset.theme; try { localStorage.removeItem(KEY + '-theme'); } catch (e) { } }
    if (admin) { try { var d = JSON.parse(localStorage.getItem(KEY + '-dash-' + CL.user.email) || 'null'); if (Array.isArray(d)) CL.user.dash = d; } catch (e) { } }
  }
  CL.lockScreens = lockScreens;
  var resize0 = initResizable;
  initResizable = function (root) { if (CL.role === 'Admin') return resize0(root); };
  var dash0 = setDashLayout;
  setDashLayout = function (l) { dash0(l); try { localStorage.setItem(KEY + '-dash-' + CL.user.email, JSON.stringify(CL.user.dash)); } catch (e) { } };
  document.addEventListener('click', function (e) {
    if (CL.role !== 'Admin' && e.target.closest('[data-act="theme"],[data-to="setup"],[data-act="dw-edit"]')) { e.preventDefault(); e.stopPropagation(); }
  }, true);
  var st = document.createElement('style');
  st.textContent = 'body.cl-locked [data-act="theme"],body.cl-locked [data-to="setup"],body.cl-locked .col-grip,body.cl-locked [data-act="dw-edit"]{display:none!important}';
  document.head.appendChild(st);

  /* ---------------- Change log (Admins only), under Setup - Tiffany, 23 Nov 2026 ----------------
     Every bug or change John or Tiffany reports: what, who, status, when it went live.
     Kept in the EqSettings list as the "changelog" row; the Worker only gives it to Admins. */
  var CL_STATUSES = ['Not started', 'In progress', 'Ready to publish', 'Implemented', "Won't fix"];
  var CL_SEED = [
    { n: 1, reported: '2026-11-23', by: 'John', area: 'Maintenance', issue: "Parts on a task can't be edited after they are entered.",
      change: 'Each part has an Edit link: name, part #, vendor, PO, quantity and unit cost. Each edit is written to the service request history. For parts from inventory, changing the quantity adjusts stock by the difference.', status: 'Implemented', done: '2026-11-23', notes: '' },
    { n: 2, reported: '2026-11-23', by: 'John', area: 'Search', issue: 'Picking an inventory item from the top search opened Adjust quantity instead of the item.',
      change: 'Now opens the Inventory page filtered to that item, with its quantity, who has it, and Edit and Adjust.', status: 'Implemented', done: '2026-11-23', notes: '' }
  ];
  var logRows = function () { return (CL.changelog && CL.changelog.rows) || []; };
  var today = function () { return toLocalInput(now()).slice(0, 10); };
  var dStr = function (d) { return d ? fmtDate(d + 'T00:00') : ''; };
  var statusChip = function (s) { var c = s === 'Implemented' ? 'available' : (s === 'Ready to publish' || s === 'In progress') ? 'maint' : 'neutral'; return '<span class="chip ' + c + '">' + esc(s) + '</span>'; };
  async function saveLog(rows) {
    var rec = { rows: rows, updatedAt: new Date().toISOString(), updatedBy: CL.user.name };
    var res = await api('/api/save', { method: 'POST', json: { changes: [{ coll: 'settings', key: 'changelog', op: 'upsert', rec: rec, etag: CL.etags['settings:changelog'] || '' }] } });
    if ((res.failed || []).length) { await CL.reload(); throw new Error(res.failed[0].conflict ? 'Someone else changed the log first. It has been reloaded; make your change again.' : res.failed[0].error); }
    CL.etags['settings:changelog'] = res.etags['settings:changelog']; CL.changelog = rec;
  }
  CL.saveLog = saveLog;
  function vChangeLog() {
    if (CL.role !== 'Admin') return band('Change log', '', '') + '<div class="content"><div class="empty"><strong>Admins only</strong></div></div>';
    var rows = logRows().slice().sort(function (a, b) { return b.n - a.n; });
    var count = function (s) { return rows.filter(function (r) { return r.status === s; }).length; };
    return band('Change log', 'Bugs and changes John or Tiffany have reported, and whether each one is live yet. Newest first.', '<button class="btn" data-act="goto" data-to="setup">Back to Setup</button><button class="btn primary" data-cl-log="new">+ Add entry</button>') +
      '<div class="content"><div class="kpis">' + ['Implemented', 'Ready to publish', 'In progress', 'Not started'].map(function (s) { return '<div class="kpi"><span class="lbl">' + s + '</span><span class="val num">' + count(s) + '</span></div>'; }).join('') + '</div>' +
      '<div class="tablewrap" style="margin-top:14px"><table><thead><tr><th>#</th><th>Reported</th><th>By</th><th>Area</th><th>Issue or request</th><th>What changed</th><th>Status</th><th>Live since</th></tr></thead><tbody>' +
      (rows.map(function (r) { return '<tr data-cl-row="' + r.n + '" style="cursor:pointer"><td class="num">' + r.n + '</td><td class="num" style="white-space:nowrap">' + esc(dStr(r.reported)) + '</td><td>' + esc(r.by) + '</td><td>' + esc(r.area) + '</td><td>' + esc(r.issue) + (r.notes ? '<div class="small muted">' + esc(r.notes) + '</div>' : '') + '</td><td class="small">' + esc(r.change || '') + '</td><td>' + statusChip(r.status) + '</td><td class="num" style="white-space:nowrap">' + esc(dStr(r.done)) + '</td></tr>'; }).join('') ||
        '<tr><td colspan="8"><div class="empty"><strong>Nothing logged yet</strong></div></td></tr>') + '</tbody></table></div>' +
      (CL.changelog && CL.changelog.updatedAt ? '<p class="small muted" style="margin-top:8px">Last updated ' + esc(new Date(CL.changelog.updatedAt).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' })) + (CL.changelog.updatedBy ? ' by ' + esc(CL.changelog.updatedBy) : '') + '.</p>' : '') + '</div>';
  }
  function logForm(n) {
    var r = n ? logRows().find(function (x) { return x.n === n; }) : { reported: today(), by: (CL.user.name || '').split(' ')[0], area: '', issue: '', change: '', status: 'Not started', done: '', notes: '' };
    if (!r) return;
    var f = function (id, label, val, full, type) { return '<div class="field' + (full ? ' full' : '') + '"><label for="' + id + '">' + label + '</label>' + (type === 'area' ? '<textarea id="' + id + '">' + esc(val) + '</textarea>' : '<input id="' + id + '" type="' + (type || 'text') + '" value="' + esc(val) + '">') + '</div>'; };
    openModal('<header><h2>' + (n ? 'Change log #' + n : 'Add to the change log') + '</h2><button class="x" data-act="close" aria-label="Close">×</button></header><div class="mbody"><div class="form">' +
      f('cl-reported', 'Date reported', r.reported, false, 'date') + f('cl-by', 'Found by', r.by) + f('cl-area', 'Area', r.area, true) +
      f('cl-issue', 'Issue or request', r.issue, true, 'area') + f('cl-change', 'What changed', r.change, true, 'area') +
      '<div class="field"><label for="cl-status">Status</label><select id="cl-status">' + CL_STATUSES.map(function (s) { return '<option ' + (s === r.status ? 'selected' : '') + '>' + esc(s) + '</option>'; }).join('') + '</select></div>' +
      f('cl-done', 'Live since', r.done, false, 'date') + f('cl-notes', 'Notes', r.notes, true) +
      '<p class="small muted full">Choosing Implemented fills in today as Live since, if it is empty. You can change the date.</p></div></div>' +
      '<footer>' + (n ? '<button class="btn danger" data-cl-log="del" data-n="' + n + '">Delete</button>' : '') + '<button class="btn" data-act="close">Cancel</button><button class="btn primary" data-cl-log="save" data-n="' + (n || '') + '">Save</button></footer>', 'narrow');
  }
  async function logSave(n, del) {
    var g = function (id) { var el = document.getElementById(id); return el ? el.value.trim() : ''; };
    var rows = logRows().map(function (r) { return Object.assign({}, r); });
    if (del && !logSave.sure) { confirmBox('Delete change log #' + n + '?', 'This removes the entry for John and you.', 'Delete', function () { logSave.sure = true; logSave(n, true).finally(function () { logSave.sure = false; }); }); return; }
    if (del) rows = rows.filter(function (r) { return r.n !== n; });
    else {
      if (!g('cl-issue')) { toast('Describe the issue or request'); return; }
      var rec = { reported: g('cl-reported') || today(), by: g('cl-by'), area: g('cl-area'), issue: g('cl-issue'), change: g('cl-change'), status: g('cl-status'), done: g('cl-done'), notes: g('cl-notes') };
      if (rec.status === 'Implemented' && !rec.done) rec.done = today();
      if (n) { var i = rows.findIndex(function (r) { return r.n === n; }); rows[i] = Object.assign(rows[i], rec); }
      else { rec.n = 1 + Math.max.apply(null, [0].concat(rows.map(function (r) { return r.n; }))); rows.push(rec); }
    }
    try { await saveLog(rows); closeModal(); rerender(); toast(del ? 'Deleted' : 'Change log saved'); } catch (e) { toast(e.message); }
  }
  // picking Implemented fills in today's date straight away (if Live since is empty), so it shows before Save
  document.addEventListener('change', function (e) {
    if (e.target.id !== 'cl-status') return;
    var d = document.getElementById('cl-done'); if (d && e.target.value === 'Implemented' && !d.value) d.value = today();
  });
  VIEWS.changelog = vChangeLog;
  var setup0 = VIEWS.setup;
  VIEWS.setup = function () {
    var html = setup0(), n = logRows().filter(function (r) { return r.status !== 'Implemented' && r.status !== "Won't fix"; }).length;
    var panel = '<div class="panel" style="margin-bottom:14px"><h3>Change log <button class="btn sm" data-act="goto" data-to="changelog">Open</button></h3><div class="body small">Bugs and changes you or John have reported and whether each one is live. ' +
      logRows().length + ' logged' + (n ? ', ' + n + ' not live yet' : ', all live') + '.</div></div>';
    return CL.role === 'Admin' ? html.replace('<div class="content">', '<div class="content">' + panel) : html;
  };
  document.addEventListener('click', function (e) {
    var b = e.target.closest('[data-cl-log]'), row = e.target.closest('[data-cl-row]');
    if (b) { e.preventDefault(); var a = b.dataset.clLog, n = +b.dataset.n || 0; if (a === 'new') logForm(0); else if (a === 'save') logSave(n, false); else if (a === 'del') logSave(n, true); }
    else if (row && CL.role === 'Admin') logForm(+row.dataset.clRow);
  });
  // first time an Admin opens the app after this ships: start the log with John's two fixes
  CL.seedLog = async function () { if (CL.role === 'Admin' && !CL.changelog) { try { await saveLog(CL_SEED.map(function (r) { return Object.assign({}, r); })); } catch (e) { } } };

  /* bits the patched screens call */
  window.cloudJobSelect = function () {
    var last = localStorage.getItem(KEY + '-lastjob') || '', mine = CL.jobs || [];
    var jobs = CL.jobFeed.slice().sort(function (a, b) { var ma = mine.indexOf(String(a.job_number).toUpperCase()) >= 0 ? 0 : 1, mb = mine.indexOf(String(b.job_number).toUpperCase()) >= 0 ? 0 : 1; return ma - mb || String(a.job_number).localeCompare(String(b.job_number)); });
    var opt = function (j) { return '<option value="' + esc(j.job_number) + '" ' + (j.job_number === last ? 'selected' : '') + '>' + esc(j.job_number + ' - ' + j.name) + '</option>'; };
    var m = jobs.filter(function (j) { return mine.indexOf(String(j.job_number).toUpperCase()) >= 0; }), o = jobs.filter(function (j) { return mine.indexOf(String(j.job_number).toUpperCase()) < 0; });
    return '<div class="field full"><label for="rq-job">Job</label><select id="rq-job" name="job" required><option value="">Pick the job</option>' +
      (m.length ? '<optgroup label="Your jobs">' + m.map(opt).join('') + '</optgroup><optgroup label="All active jobs">' + o.map(opt).join('') + '</optgroup>' : jobs.map(opt).join('')) + '</select></div>';
  };
  window.cloudDeliveryFields = function () {
    return '<div class="field"><label for="rq-del">Delivery or pickup</label><select id="rq-del" name="delivery" required><option value="">Pick one</option><option>Delivery</option><option>Pickup</option></select></div>' +
      '<div class="field"><label for="rq-deln">Delivery details</label><input id="rq-deln" name="deliveryNotes" placeholder="Where on site, gate code, who to call"></div>';
  };
  window.cloudReqExtra = function (r, admin) {
    var bits = [];
    if (r.delivery) bits.push('<span><b>' + esc(r.delivery) + '</b>' + (r.deliveryNotes ? ': ' + esc(r.deliveryNotes) : '') + '</span>');
    if (r.source === 'crew') bits.push('<span>Field crew' + (r.selfPerform ? ' · Self Perform' : '') + (admin && r.byPhone ? ' · ' + esc(fmtPhone(r.byPhone)) : '') + '</span>');
    if (r.status === 'Awaiting approval' && r.approvers) bits.push('<span>Waiting on <b>' + esc(r.approvers.label) + '</b></span>');
    if (r.approvedBy) bits.push('<span>Approved by ' + esc(r.approvedBy) + '</span>');
    return bits.length ? '<div class="rq-req-meta">' + bits.join('') + '</div>' : '';
  };

  /* save status and problems */
  function setSaveState(t) { var el = document.querySelector('#cl-save'); if (el) el.textContent = t; }
  function showBanner(t, bad) {
    var b = document.querySelector('#cl-banner');
    if (!b) { b = document.createElement('div'); b.id = 'cl-banner'; b.setAttribute('role', 'alert'); document.body.prepend(b); }
    b.className = bad ? 'cl-banner bad' : 'cl-banner';
    b.innerHTML = esc(t) + ' <button class="btn sm" data-cl="reload">Reload</button> <button class="btn sm" data-cl="close">Dismiss</button>';
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest('[data-cl]'); if (!b) return;
    if (b.dataset.cl === 'reload') location.reload();
    else if (b.dataset.cl === 'close') { var bn = document.querySelector('#cl-banner'); if (bn) bn.remove(); CL.conflicts = []; }
  });
  window.addEventListener('beforeunload', function (e) { if (CL.saving || CL.timer && CL.base && JSON.stringify(snapshot()) !== JSON.stringify(CL.base)) { e.preventDefault(); e.returnValue = ''; } });

  /* ---------------- start ---------------- */
  (async function () {
    // MSAL's hidden sign-in frame and sign-in popup land back on this page with ?code= / #code=; the opening page reads it, so don't start the app there
    if ((window.self !== window.top || window.opener) && /[#&?](code|error)=/.test(location.hash + location.search)) return;
    var st = document.querySelector('#auth-root');
    var msg = function (t) { if (st) st.innerHTML = '<div class="auth"><div class="auth-card"><div class="logo auth-logo">' + brandHtml('auth') + '</div><p>' + t + '</p></div></div>'; };
    msg('Signing you in…');
    try {
      if (!(await signIn())) return;
      msg('Loading the equipment list…');
      await load();
      render();
      CL.seedLog && CL.seedLog();
      var bar = document.querySelector('.topbar'); if (bar && !document.querySelector('#cl-save')) { var s = document.createElement('span'); s.id = 'cl-save'; s.className = 'small muted'; s.style.marginLeft = '8px'; bar.appendChild(s); }
      setSaveState(CL.role === 'Admin' ? 'Saved' : '');
    } catch (e) {
      msg(esc(e.message || 'Something went wrong.') + '<br><br><button class="btn primary" data-cl="reload">Try again</button>');
    }
  })();
})();
