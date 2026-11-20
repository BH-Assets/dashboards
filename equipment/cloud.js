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
  async function idToken() {
    try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account: acct })).idToken; }
    catch (e) { await pca.acquireTokenRedirect({ scopes: SCOPES, account: acct }); throw e; }
  }
  async function api(path, opt) {
    opt = opt || {};
    var headers = Object.assign({ Authorization: 'Bearer ' + await idToken() }, opt.json !== undefined ? { 'Content-Type': 'application/json' } : {}, opt.headers || {});
    var r = await fetch(API + path, { method: opt.method || 'GET', headers: headers, body: opt.json !== undefined ? JSON.stringify(opt.json) : opt.body });
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
  route = function () { var h = location.hash.replace('#', ''); if (h === 'approvals' && mayApprove()) return 'approvals'; return route0(); };
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
      (CL.role === 'Admin' ? '<button role="menuitem" data-act="goto" data-to="users">Who has access</button>' : '') +
      '<button role="menuitem" data-act="theme">Switch light / dark</button><button role="menuitem" data-act="logout" class="pm-out">Sign out</button>';
    document.body.appendChild(m);
    var r = document.querySelector('#avatar-btn').getBoundingClientRect(), narrow = matchMedia('(max-width:760px)').matches;
    m.style.left = (narrow ? Math.max(8, r.left) : r.right + 8) + 'px'; m.style.top = (narrow ? r.bottom + 6 : r.top) + 'px';
  };
  VIEWS.users = function () {
    return band('Who has access', 'People sign in with their Bailey-Harris Microsoft account. Nobody has a separate password for this app.', '') +
      '<div class="content"><div class="panel"><div class="body"><table class="tbl"><thead><tr><th>Role</th><th>Who</th><th>What they can do</th></tr></thead><tbody>' +
      '<tr><td><b>Admin</b></td><td>John Barron, Tiffany Blackmon</td><td>Everything: catalog, rates, orders, check-outs, the shop, requests.</td></tr>' +
      '<tr><td><b>Management</b></td><td>Clay (Self Perform)</td><td>Everything Super/PM can, and approves Self Perform crew requests.</td></tr>' +
      '<tr><td><b>Super/PM</b></td><td>Anyone listed as a superintendent or on the PM team of an active or warranty job in Procore, plus the EqAccessExtras list</td><td>Sees the whole catalog with rates, requests equipment, stock and inventory (straight to John), and approves crew requests for their jobs.</td></tr>' +
      '<tr><td><b>Viewer</b></td><td>Everyone else at Bailey-Harris</td><td>Sees the catalog, rates and what is available.</td></tr>' +
      '<tr><td><b>Field crews</b></td><td>No sign-in</td><td>Request stock and inventory at dashboard.baileyharris.com/request with a job number.</td></tr>' +
      '</tbody></table><p class="small muted" style="margin-top:12px">Admin and Management are set on the Equipment service (ADMIN_EMAILS, MGMT_EMAILS). Super/PM comes from Procore every night. To give someone Super/PM who is not on a Procore team, add their email to the EqAccessExtras list on the Equipment SharePoint site.</p></div></div></div>';
  };

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
    else { document.querySelector('#cl-banner').remove(); CL.conflicts = []; }
  });
  window.addEventListener('beforeunload', function (e) { if (CL.saving || CL.timer && CL.base && JSON.stringify(snapshot()) !== JSON.stringify(CL.base)) { e.preventDefault(); e.returnValue = ''; } });

  /* ---------------- start ---------------- */
  (async function () {
    var st = document.querySelector('#auth-root');
    var msg = function (t) { if (st) st.innerHTML = '<div class="auth"><div class="auth-card"><div class="logo auth-logo">' + brandHtml('auth') + '</div><p>' + t + '</p></div></div>'; };
    msg('Signing you in…');
    try {
      if (!(await signIn())) return;
      msg('Loading the equipment list…');
      await load();
      render();
      var bar = document.querySelector('.topbar'); if (bar && !document.querySelector('#cl-save')) { var s = document.createElement('span'); s.id = 'cl-save'; s.className = 'small muted'; s.style.marginLeft = '8px'; bar.appendChild(s); }
      setSaveState(CL.role === 'Admin' ? 'Saved' : '');
    } catch (e) {
      msg(esc(e.message || 'Something went wrong.') + '<br><br><button class="btn primary" data-cl="reload">Try again</button>');
    }
  })();
})();
