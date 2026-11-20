/* Crew request page (dashboard.baileyharris.com/request/), 19 Nov 2026.
 * No sign-in. Field crews request stock and inventory (not equipment) with their job number.
 * The bh-equipment Worker checks the job, the Turnstile token and the rate limits, routes the request
 * to the approver (Self Perform -> Clay; else the job's super, then PM, then John) and returns a status link.
 * ?s=RQ-1042-<code> shows that request's status instead of the form. */
(function () {
  'use strict';
  var API = 'https://equipment-api.baileyharris.com';
  var TURNSTILE_SITEKEY = '0x4AAAAAAGKzG-DvaQXiirFQ';
  var $ = function (s) { return document.querySelector(s); };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var money = function (v) { return v == null ? '' : '$' + (+v).toLocaleString('en-US', { minimumFractionDigits: v % 1 ? 2 : 0, maximumFractionDigits: 2 }); };
  var main = $('#main'), CAT = null, cart = {}, q = '', job = null, tsId = null;
  var fmtDate = function (d) { return d ? new Date(d + 'T12:00').toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' }) : ''; };
  var today = function (plus) { var d = new Date(Date.now() + (plus || 0) * 864e5 - new Date().getTimezoneOffset() * 6e4); return d.toISOString().slice(0, 10); };
  async function get(path) { var r = await fetch(API + path), j = await r.json().catch(function () { return {}; }); if (!r.ok) throw new Error(j.error || 'Something went wrong (' + r.status + ').'); return j; }

  /* ---------------- status ---------------- */
  async function status(s) {
    main.innerHTML = '<p class="rq-loading">Looking up your request…</p>';
    try {
      var j = await get('/public/status?s=' + encodeURIComponent(s));
      main.innerHTML = '<section class="rq-sec"><h2>' + esc(j.number) + ' · Job ' + esc(j.job) + '</h2><div class="rq-status">' + esc(j.status) + '</div>' +
        (j.note ? '<div class="rq-err" style="background:#F2ECE1;color:#211E1E">Note: ' + esc(j.note) + '</div>' : '') +
        '<p>Needed ' + esc(fmtDate(j.needBy)) + (j.returnBy ? ', back by ' + esc(fmtDate(j.returnBy)) : '') + ' · ' + esc(j.delivery || '') + '</p>' +
        '<ul class="rq-lines">' + j.lines.map(function (l) { return '<li>' + esc(l.name) + ' <b>× ' + esc(l.qty) + '</b> ' + esc(l.unit) + '</li>'; }).join('') + '</ul></section>' +
        '<section class="rq-sec"><a class="rq-btn" style="display:block;text-align:center;text-decoration:none" href="./">New request</a></section>';
    } catch (e) { main.innerHTML = '<div class="rq-err">' + esc(e.message) + '</div><a href="./">New request</a>'; }
  }

  /* ---------------- form ---------------- */
  function itemRow(x) {
    var k = x.kind + ':' + x.id, n = cart[k] || 0, price = x.kind === 'inv' ? (x.price != null ? money(x.price) + ' per ' + x.unit : '') :
      [['day', 'day'], ['week', 'wk'], ['month', 'mo']].filter(function (p) { return x[p[0]] > 0; }).map(function (p) { return money(x[p[0]]) + '/' + p[1]; }).join(' · ');
    var img = x.photo ? '<img src="' + API + '/public/photo?kind=' + x.kind + '&id=' + encodeURIComponent(x.id) + '" alt="" loading="lazy">' : '<span class="ph"></span>';
    var ctl = x.available <= 0 && !n ? '<span class="rq-none">None available</span>' : n ? '<div class="rq-step"><button type="button" data-d="' + esc(k) + '" aria-label="One less">−</button><span>' + n + '</span><button type="button" data-i="' + esc(k) + '" aria-label="One more">+</button></div>' : '<button type="button" class="rq-add" data-i="' + esc(k) + '">Add</button>';
    return '<div class="rq-item">' + img + '<div><b>' + esc(x.name) + '</b><small>' + esc(x.available) + ' ' + esc(x.unit) + ' available' + (price ? ' · ' + esc(price) : '') + '</small></div>' + ctl + '</div>';
  }
  function list() {
    var all = CAT.stock.concat(CAT.inventory), s = q.toLowerCase();
    var hits = all.filter(function (x) { return !s || (x.name + ' ' + x.category).toLowerCase().indexOf(s) >= 0; }).sort(function (a, b) { return a.name.localeCompare(b.name); });
    $('#items').innerHTML = hits.map(itemRow).join('') || '<p class="rq-none">Nothing matches. Put what you need in the notes.</p>';
    var n = Object.keys(cart).length; $('#bar').hidden = !n; $('#bar span').textContent = n + ' item' + (n === 1 ? '' : 's') + ' added';
  }
  function form() {
    var saved = {}; try { saved = JSON.parse(localStorage.getItem('bh-crew') || '{}'); } catch (e) { }
    main.innerHTML =
      '<form id="f" novalidate>' +
      '<section class="rq-sec"><h2>Your job</h2>' +
      '<div class="rq-f"><label for="job">Job number</label><input id="job" name="job" autocomplete="off" autocapitalize="characters" placeholder="e.g. 25AU010" value="' + esc(saved.job || '') + '" required><span class="rq-hint" id="jobmsg">Your superintendent can give you this.</span></div>' +
      '<div class="rq-f"><label for="name">Your name</label><input id="name" name="name" autocomplete="name" value="' + esc(saved.name || '') + '" required></div>' +
      '<div class="rq-f"><label for="phone">Your cell number</label><input id="phone" name="phone" type="tel" inputmode="tel" autocomplete="tel" value="' + esc(saved.phone || '') + '" required></div>' +
      '<label class="rq-check"><input type="checkbox" id="sp" checked> <span>I\'m on a <b>Self Perform</b> crew<br><span class="rq-hint">Leave this checked unless you work for a superintendent on another crew.</span></span></label></section>' +
      '<section class="rq-sec"><h2>What you need</h2><div class="rq-search"><div class="rq-f" style="margin:0"><input id="q" type="search" placeholder="Search stock and inventory" aria-label="Search"></div></div><div id="items"></div>' +
      '<p class="rq-hint">Equipment (machines, lifts, trailers) is requested by your superintendent.</p></section>' +
      '<section class="rq-sec"><h2>When and where</h2>' +
      '<div class="rq-f"><label for="need">Needed by</label><input id="need" type="date" min="' + today() + '" value="' + today(1) + '" required></div>' +
      '<div class="rq-f"><label for="ret">Back by (if it comes back)</label><input id="ret" type="date" min="' + today() + '"></div>' +
      '<div class="rq-f"><label>Delivery or pickup</label><div class="rq-seg"><label><input type="radio" name="del" value="Delivery"> Delivery</label><label><input type="radio" name="del" value="Pickup"> Pickup</label></div></div>' +
      '<div class="rq-f"><label for="deln">Delivery details</label><input id="deln" placeholder="Where on site, gate code, who to call"></div>' +
      '<div class="rq-f"><label for="notes">Notes</label><textarea id="notes" placeholder="Anything else"></textarea></div></section>' +
      '<section class="rq-sec"><div id="ts"></div><div id="err"></div><button class="rq-btn" id="send" type="submit">Send request</button></section></form>' +
      '<div class="rq-bar" id="bar" hidden><span></span><button class="rq-btn" type="button" id="tosend">Review</button></div>';
    list();
    $('#q').addEventListener('input', function (e) { q = e.target.value; list(); });
    $('#items').addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b) return;
      if (b.dataset.i) cart[b.dataset.i] = (cart[b.dataset.i] || 0) + 1;
      if (b.dataset.d) { cart[b.dataset.d] -= 1; if (cart[b.dataset.d] <= 0) delete cart[b.dataset.d]; }
      list();
    });
    $('#tosend').addEventListener('click', function () { $('#need').scrollIntoView({ behavior: 'smooth' }); });
    $('#job').addEventListener('change', checkJob); if (saved.job) checkJob();
    $('#f').addEventListener('submit', send);
    var tryTs = function () { if (window.turnstile) tsId = turnstile.render('#ts', { sitekey: TURNSTILE_SITEKEY, appearance: 'interaction-only' }); else setTimeout(tryTs, 300); }; tryTs();
  }
  async function checkJob() {
    var v = $('#job').value.trim(), m = $('#jobmsg'); job = null; if (!v) return;
    m.className = 'rq-hint'; m.textContent = 'Checking…';
    try { var j = await get('/public/job?n=' + encodeURIComponent(v)); if (j.ok) { job = j.job; m.className = 'rq-ok'; m.textContent = j.job + ' · ' + j.name; } else { m.className = 'rq-bad'; m.textContent = 'That is not an active job number. Check it with your superintendent.'; } }
    catch (e) { m.className = 'rq-bad'; m.textContent = e.message; }
  }
  function err(t) { $('#err').innerHTML = t ? '<div class="rq-err">' + esc(t) + '</div>' : ''; if (t) $('#err').scrollIntoView({ behavior: 'smooth', block: 'center' }); }
  async function send(e) {
    e.preventDefault(); err('');
    var del = (document.querySelector('input[name="del"]:checked') || {}).value;
    var lines = Object.keys(cart).map(function (k) { var p = k.split(':'); return { kind: p[0], id: p.slice(1).join(':'), qty: cart[k] }; });
    if (!job) return err('Enter your job number.');
    if ($('#name').value.trim().length < 2) return err('Enter your name.');
    if ($('#phone').value.replace(/\D/g, '').length < 10) return err('Enter your 10-digit cell number.');
    if (!lines.length) return err('Add at least one item.');
    if (!del) return err('Pick delivery or pickup.');
    var tok = window.turnstile && tsId != null ? turnstile.getResponse(tsId) : '';
    $('#send').disabled = true; $('#send').textContent = 'Sending…';
    try {
      var r = await fetch(API + '/public/request', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({
        turnstile: tok, name: $('#name').value, phone: $('#phone').value, job: job, selfPerform: $('#sp').checked, needBy: $('#need').value, returnBy: $('#ret').value,
        delivery: del, deliveryNotes: $('#deln').value, notes: $('#notes').value, lines: lines }) });
      var j = await r.json().catch(function () { return {}; });
      if (!r.ok) throw new Error(j.error || 'Could not send (' + r.status + ').');
      try { localStorage.setItem('bh-crew', JSON.stringify({ job: job, name: $('#name').value, phone: $('#phone').value })); } catch (x) { }
      done(j);
    } catch (x) { err(x.message); $('#send').disabled = false; $('#send').textContent = 'Send request'; if (window.turnstile && tsId != null) turnstile.reset(tsId); }
  }
  function done(j) {
    var sms = 'sms:?&body=' + encodeURIComponent(j.number + ' status: ' + j.statusUrl);
    main.innerHTML = '<section class="rq-sec"><div class="rq-done"><div>Request sent</div><div class="num">' + esc(j.number) + '</div>' +
      '<p>It goes to <b>' + esc(j.approver) + '</b> to approve, then to the Equipment Manager.</p>' +
      '<p><b>Save this link</b> to check on it:</p><div class="rq-link">' + esc(j.statusUrl) + '</div>' +
      '<div class="rq-actions"><a href="' + esc(sms) + '">Text it to me</a><button type="button" id="copy">Copy link</button><a href="' + esc(j.statusUrl) + '">Open it</a></div></div></section>' +
      '<section class="rq-sec"><a class="rq-btn" style="display:block;text-align:center;text-decoration:none" href="./">New request</a></section>';
    $('#copy').addEventListener('click', function () { navigator.clipboard && navigator.clipboard.writeText(j.statusUrl).then(function () { $('#copy').textContent = 'Copied'; }); });
  }

  var s = new URLSearchParams(location.search).get('s');
  if (s) status(s);
  else get('/public/catalog').then(function (c) { CAT = c; form(); }).catch(function (e) { main.innerHTML = '<div class="rq-err">' + esc(e.message) + '</div>'; });
})();
