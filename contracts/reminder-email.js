/* reminder-email.js — the monthly pay application reminder, built once for both the Project Hub preview
   (browser: window.BHReminder) and send_billing_reminders.js (node: require). 24 Nov 2026.

   Layout follows the 22 Oct portal mockup ("1 · Monthly email"): black Bailey-Harris bar, PAY APPLICATION
   REMINDER heading, the pay app number and due date, the start button with the secure-link note, the
   Billy compliance warning, YOUR CONTRACT, CHANGE ORDERS (bill this period / pending), OUR RECORDS with
   retainage and the last check, "keep [subcontract] in the subject", and the company footer. Added since:
   the No billing this period button, the job note and the PM contact. (Lien waivers are not in the build.)

   Email-safe HTML: tables with role="presentation", inline styles, bgcolor, a 600 px column, bulletproof
   table buttons, a hidden preheader, no images, no web fonts, no flex or grid. Every email also carries a
   plain-text part (multipart/alternative) for clients that show text only.

   The wording around the figures is the Project Hub's company template: subject + body with merge fields.
   A paragraph that holds only block fields ({Heading}, {Submit button}, {No billing button},
   {Compliance}, {Contract summary}) becomes that block; other paragraphs are text with fields filled in. */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BHReminder = factory();
})(this, function () {
  'use strict';
  var INK = '#211e1e', MUTED = '#484b4d', LINE = '#e8e6e7', RED = '#ed2b30', MAROON = '#7f1b29', FONT = "'Segoe UI', Arial, Helvetica, sans-serif";
  var ADDRESS = 'Bailey-Harris Construction · 1600 Bailey Harris Drive, Auburn, AL 36830';

  var DEFAULT_TPL = {
    subject: '{Pay app} due {Due date short} · {Job} ({Subcontract})',
    body: '{Heading}\n\nHello {Contact first name}. Here is everything you need for this month\'s application on {Job}, for work completed through {Period end}, with no projection of work. The figures below are already filled in on your pay app.\n\n{Submit button}  {No billing button}\n\n{Compliance}\n\n{Contract summary}\n\n{Job note}\n\nQuestions? Reply to this email and keep [{Subcontract}] in the subject, or contact {PM name} at {PM email}. If the figures above don\'t match your records, tell us before you submit.'
  };
  var FIELDS = ['{Heading}', '{Pay app}', '{Trade partner}', '{Contact first name}', '{Job}', '{Subcontract}', '{Due date}', '{Due date short}', '{Period end}', '{Link expires}',
    '{Submit button}', '{No billing button}', '{Compliance}', '{Contract summary}', '{Job note}', '{PM name}', '{PM email}'];
  var BLOCKS = ['{Heading}', '{Submit button}', '{No billing button}', '{Compliance}', '{Contract summary}'];

  /* ---------- small helpers ---------- */
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var money = function (v) { return v == null || v === '' || isNaN(Number(v)) ? '-' : '$' + Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
  var d = function (iso) { var p = String(iso || '').slice(0, 10).split('-'); return p.length === 3 ? new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12)) : null; };
  var longDate = function (iso) { var x = d(iso); return x ? x.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', timeZone: 'UTC' }) : ''; };
  var shortDate = function (iso) { var x = d(iso); return x ? x.toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' }) : ''; };
  var us = function (iso) { var x = d(iso); return x ? (x.getUTCMonth() + 1) + '/' + x.getUTCDate() + '/' + x.getUTCFullYear() : ''; };
  var dots = function (label, value, w) { w = w || 64; label = String(label); value = String(value); var n = Math.max(2, w - label.length - value.length); return '  ' + label + ' ' + '.'.repeat(n - 2) + ' ' + value; };

  /* ---------- the facts for one email ----------
     data = { vendor, firstName, job: { number, name }, due, periodEnd, linkUntil, link, note, pm: { name, email },
              subs: [{ number, title, nextPr, disc, orig, earlier, fresh, adj, cos: [{ no, title, amount, bill, isNew, why }],
                       rec: { pr, ctd, ret, prev } | null, retText, last: { check, date, amount } | null, comp: { bad, text } | null }] } */
  function payAppLabel(data) { return data.subs.length === 1 ? 'Pay App #' + (data.subs[0].nextPr || 1) : 'Your pay apps'; }

  /* ---------- HTML blocks ---------- */
  function hHeading(data) {
    return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding:0 0 18px;font-family:' + FONT + '">' +
      '<div style="font-size:12px;line-height:16px;font-weight:700;letter-spacing:1px;color:' + MAROON + '">PAY APPLICATION REMINDER</div>' +
      '<h1 style="margin:8px 0 0;font-size:26px;line-height:32px;font-weight:700;color:' + INK + '">' + esc(payAppLabel(data)) + ' ' + (data.subs.length === 1 ? 'is' : 'are') + ' due by ' + esc(longDate(data.due)) + '</h1></td></tr></table>';
  }
  function hButtons(data, submit, nobill) {
    var b1 = '<td bgcolor="' + INK + '" style="background:' + INK + ';padding:14px 22px"><a href="' + esc(data.link) + '" style="font-family:' + FONT + ';font-size:14px;line-height:18px;font-weight:700;letter-spacing:1px;color:#ffffff;text-decoration:none;display:inline-block">' + esc(data.subs.length === 1 ? 'START PAY APP #' + (data.subs[0].nextPr || 1) : 'START YOUR PAY APPS') + '</a></td>';
    var b2 = '<td style="border:1px solid ' + INK + ';padding:13px 20px"><a href="' + esc(data.link) + '" style="font-family:' + FONT + ';font-size:14px;line-height:18px;font-weight:700;letter-spacing:1px;color:' + INK + ';text-decoration:none;display:inline-block">NO BILLING THIS PERIOD</a></td>';
    var cells = []; if (submit) cells.push(b1); if (submit && nobill) cells.push('<td width="10" style="width:10px;font-size:0;line-height:0">&nbsp;</td>'); if (nobill) cells.push(b2);
    return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 6px"><tr>' + cells.join('') + '</tr></table>' +
      '<p style="margin:0 0 20px;font-family:' + FONT + ';font-size:13px;line-height:19px;color:' + MUTED + '">Secure link for ' + esc(data.vendor) + ' · works until ' + esc(shortDate(data.linkUntil)) + '. Please don\'t forward it outside your company.</p>';
  }
  function hCompliance(data) {
    return data.subs.filter(function (s) { return s.comp; }).map(function (s) {
      var bad = s.comp.bad;
      return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 18px"><tr><td bgcolor="#fff4f1" style="background:#fff4f1;border:1px solid #f3c2bd;padding:14px 16px;font-family:' + FONT + ';font-size:14px;line-height:21px;color:' + INK + '">' +
        '<b>' + (bad ? 'You are not compliant in Billy.' : 'Insurance expiring soon.') + '</b> ' + esc(s.comp.text) + '. ' + (bad ? 'Upload a current certificate in Billy before the due date, or payment on this application may be held.' : 'Upload a renewal in Billy before it lapses.') +
        (data.subs.length > 1 ? ' <span style="color:' + MUTED + '">(' + esc(s.number) + ')</span>' : '') + '</td></tr></table>';
    }).join('');
  }
  var row = function (l, v, opt) { opt = opt || {}; var b = opt.bold ? 'font-weight:700;' : '', c = opt.muted ? 'color:' + MUTED + ';' : 'color:' + INK + ';', bd = opt.last ? '' : 'border-bottom:1px solid ' + LINE + ';';
    return '<tr><td style="padding:8px 0;' + bd + b + c + 'font-family:' + FONT + ';font-size:14px;line-height:20px">' + l + '</td><td align="right" style="padding:8px 0;' + bd + b + c + 'font-family:' + FONT + ';font-size:14px;line-height:20px;white-space:nowrap">' + v + '</td></tr>'; };
  var label = function (t) { return '<div style="margin:16px 0 4px;font-family:' + FONT + ';font-size:12px;line-height:16px;font-weight:700;letter-spacing:1px;color:' + INK + '">' + t + '</div>'; };
  function hSub(s, many) {
    var cos = (s.cos || []).filter(function (x) { return x.isNew || !x.bill; });
    return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 18px"><tr><td style="border-top:2px solid ' + INK + ';padding:12px 0 0;font-family:' + FONT + '">' +
      (many ? '<div style="font-size:16px;line-height:22px;font-weight:700;color:' + INK + '">Pay App #' + (s.nextPr || 1) + ' · subcontract ' + esc(s.number) + '</div>' : '<div style="font-size:16px;line-height:22px;font-weight:700;color:' + INK + '">Subcontract ' + esc(s.number) + '</div>') +
      '<div style="font-size:13px;line-height:19px;color:' + MUTED + '">' + esc(s.title || '') + (s.disc ? ' · 2% early-payment discount' : '') + '</div>' +
      label('YOUR CONTRACT') +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' + row('Original contract', money(s.orig)) + row('Approved change orders (on earlier pay apps)', money(s.earlier)) +
        (s.fresh ? row('New approved change orders', money(s.fresh)) : '') + row('Adjusted contract for Pay App #' + (s.nextPr || 1), money(s.adj), { bold: true, last: true }) + '</table>' +
      (cos.length ? label('CHANGE ORDERS') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' + cos.map(function (x) {
        var bg = x.isNew ? '#eef8f0' : '#f2ece1';
        return '<tr><td bgcolor="' + bg + '" style="background:' + bg + ';padding:9px 12px;font-family:' + FONT + ';font-size:14px;line-height:20px;color:' + INK + '"><b>CO #' + esc(x.no) + '</b>' + (x.title ? ' ' + esc(x.title) : '') + ' · ' + (x.isNew ? '<b>' + esc(x.why) + '</b>' : esc(x.why)) + '</td><td bgcolor="' + bg + '" align="right" style="background:' + bg + ';padding:9px 12px;font-family:' + FONT + ';font-size:14px;line-height:20px;color:' + INK + ';white-space:nowrap">' + money(x.amount) + '</td></tr><tr><td colspan="2" style="height:4px;font-size:0;line-height:0">&nbsp;</td></tr>';
      }).join('') + '</table>' : '') +
      (s.rec ? label('OUR RECORDS (THROUGH PAY APP #' + s.rec.pr + ')') + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' +
        row('Completed and stored to date', money(s.rec.ctd)) + row('Retainage held' + (s.retText ? ' · ' + esc(s.retText) : ''), money(s.rec.ret)) + row('Previous payments (line 7)', money(s.rec.prev), { last: !s.last }) +
        (s.last ? row('Last check · #' + esc(s.last.check) + ' on ' + esc(us(s.last.date)), money(s.last.amount), { muted: true, last: true }) : '') + '</table>'
        : '<div style="margin:12px 0 0;font-size:13px;line-height:19px;color:' + MUTED + '">No pay applications on our records yet: this will be Pay App #1.</div>') +
      '</td></tr></table>';
  }
  function hContracts(data) { return data.subs.map(function (s) { return hSub(s, data.subs.length > 1); }).join(''); }

  /* ---------- plain-text blocks ---------- */
  function tHeading(data) { return 'PAY APPLICATION REMINDER\n' + payAppLabel(data) + ' ' + (data.subs.length === 1 ? 'is' : 'are') + ' due by ' + longDate(data.due); }
  function tButtons(data, submit, nobill) {
    var out = [];
    if (submit) out.push((data.subs.length === 1 ? 'Start Pay App #' + (data.subs[0].nextPr || 1) : 'Start your pay apps') + ': ' + data.link);
    if (nobill) out.push('No billing this period? Tell us at the same link' + (submit ? '.' : ': ' + data.link));
    out.push('(Secure link for ' + data.vendor + ', works until ' + shortDate(data.linkUntil) + '. Please don\'t forward it outside your company.)');
    return out.join('\n');
  }
  function tCompliance(data) { return data.subs.filter(function (s) { return s.comp; }).map(function (s) { return '*** ' + (s.comp.bad ? 'YOU ARE NOT COMPLIANT IN BILLY' : 'INSURANCE EXPIRING SOON') + ' ***\n' + s.comp.text + '. ' + (s.comp.bad ? 'Upload a current certificate in Billy before the due date, or payment on this application may be held.' : 'Upload a renewal in Billy before it lapses.') + (data.subs.length > 1 ? ' (' + s.number + ')' : ''); }).join('\n\n'); }
  function tSub(s) {
    var L = [], cos = (s.cos || []).filter(function (x) { return x.isNew || !x.bill; });
    L.push('------------------------------------------------------------------');
    L.push('PAY APP #' + (s.nextPr || 1) + ' · SUBCONTRACT ' + s.number + (s.title ? ' · ' + s.title : '') + (s.disc ? ' · 2% discount' : ''));
    L.push('', 'YOUR CONTRACT', dots('Original contract', money(s.orig)), dots('Approved change orders (earlier pay apps)', money(s.earlier)));
    if (s.fresh) L.push(dots('New approved change orders', money(s.fresh)));
    L.push(dots('Adjusted contract for Pay App #' + (s.nextPr || 1), money(s.adj)));
    if (cos.length) { L.push('', 'CHANGE ORDERS'); cos.forEach(function (x) { L.push('  CO #' + x.no + (x.title ? ' ' + x.title : '') + ' · ' + x.why + ' · ' + money(x.amount)); }); }
    if (s.rec) {
      L.push('', 'OUR RECORDS (THROUGH PAY APP #' + s.rec.pr + ')', dots('Completed and stored to date', money(s.rec.ctd)), dots('Retainage held' + (s.retText ? ' (' + s.retText + ')' : ''), money(s.rec.ret)), dots('Previous payments (line 7)', money(s.rec.prev)));
      if (s.last) L.push(dots('Last check #' + s.last.check + ' on ' + us(s.last.date), money(s.last.amount)));
    } else L.push('', 'No pay applications on our records yet: this will be Pay App #1.');
    return L.join('\n');
  }
  function tContracts(data) { return data.subs.map(tSub).join('\n\n'); }

  /* ---------- build ---------- */
  function build(tpl, data) {
    tpl = tpl && tpl.body ? tpl : DEFAULT_TPL;
    var first = (String(data.firstName || '').trim().split(/\s+/)[0]) || 'there';
    var jobName = (data.job.number ? data.job.number + ' ' : '') + (data.job.name || '');
    var subsList = data.subs.map(function (s) { return s.number; }).join(', ');
    var text = { '{Pay app}': payAppLabel(data), '{Trade partner}': data.vendor || '', '{Contact first name}': first, '{Job}': jobName, '{Subcontract}': subsList,
      '{Due date}': longDate(data.due), '{Due date short}': shortDate(data.due), '{Period end}': longDate(data.periodEnd).replace(/^\w+, /, '') + (d(data.periodEnd) ? ', ' + d(data.periodEnd).getUTCFullYear() : ''), '{Link expires}': shortDate(data.linkUntil),
      '{Job note}': data.note || '', '{PM name}': (data.pm && data.pm.name) || 'Bailey-Harris Accounts Payable', '{PM email}': (data.pm && data.pm.email) || 'payables@baileyharris.com' };
    var fillText = function (s) { Object.keys(text).forEach(function (k) { s = s.split(k).join(text[k]); }); return s; };
    var subject = fillText(tpl.subject || DEFAULT_TPL.subject).replace(/\s+/g, ' ').replace(/\(\s*\)/g, '').trim();
    var paras = String(tpl.body || '').replace(/\r/g, '').split(/\n{2,}/);
    var html = [], plain = [];
    paras.forEach(function (p) {
      var t = p.trim(); if (!t) return;
      var bare = BLOCKS.reduce(function (s, k) { return s.split(k).join(''); }, t).trim();
      if (!bare) {   // a paragraph of block fields only
        var hasS = t.indexOf('{Submit button}') >= 0, hasN = t.indexOf('{No billing button}') >= 0;
        var order = t.match(/\{(Heading|Submit button|No billing button|Compliance|Contract summary)\}/g) || [], doneBtn = false;
        order.forEach(function (k) {
          if (k === '{Heading}') { html.push(hHeading(data)); plain.push(tHeading(data)); }
          else if (k === '{Compliance}') { var c = hCompliance(data); if (c) { html.push(c); plain.push(tCompliance(data)); } }
          else if (k === '{Contract summary}') { html.push(hContracts(data)); plain.push(tContracts(data)); }
          else if (!doneBtn) { doneBtn = true; html.push(hButtons(data, hasS, hasN)); plain.push(tButtons(data, hasS, hasN)); }
        });
        return;
      }
      var tx = fillText(BLOCKS.reduce(function (s, k) { return s.split(k).join(''); }, t)).trim();
      if (!tx) return;
      plain.push(tx);
      html.push('<p style="margin:0 0 16px;font-family:' + FONT + ';font-size:15px;line-height:23px;color:' + (/^Questions\?/.test(tx) ? MUTED : INK) + '">' + esc(tx).replace(/\n/g, '<br>') + '</p>');
    });
    var pre = payAppLabel(data) + ' due ' + longDate(data.due) + ' for ' + jobName + '. Submit it, or tell us there is no billing this period.';
    var inner = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="' + LINE + '" style="background:' + LINE + '"><tr><td align="center" style="padding:24px 12px">' +
      '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:600px">' +
      '<tr><td bgcolor="' + INK + '" style="background:' + INK + ';padding:16px 24px"><table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td width="16" height="16" bgcolor="' + RED + '" style="width:16px;height:16px;background:' + RED + ';font-size:0;line-height:0">&nbsp;</td><td style="padding-left:12px;font-family:' + FONT + ';font-size:14px;line-height:18px;font-weight:700;letter-spacing:1px;color:#ffffff">BAILEY-HARRIS CONSTRUCTION</td></tr></table></td></tr>' +
      '<tr><td bgcolor="#ffffff" style="background:#ffffff;padding:28px 28px 12px">' + html.join('') + '</td></tr>' +
      '<tr><td style="padding:14px 4px;font-family:' + FONT + ';font-size:12px;line-height:18px;color:' + MUTED + '">' + esc(ADDRESS) + '<br>Sent before each pay application due date to active trade partners.</td></tr>' +
      '</table></td></tr></table>';
    var doc = '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light"><title>' + esc(subject) + '</title></head>' +
      '<body style="margin:0;padding:0;background:' + LINE + '"><div style="display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all">' + esc(pre) + '</div>' + inner + '</body></html>';
    var txt = plain.join('\n\n') + '\n\n--\n' + ADDRESS + '\n';
    return { subject: subject, html: doc, inner: inner, text: txt, preheader: pre };
  }

  /* ---------- MIME (multipart/alternative) for Graph sendMail ---------- */
  function b64utf8(s) {
    if (typeof Buffer !== 'undefined') return Buffer.from(s, 'utf8').toString('base64');
    var bytes = new TextEncoder().encode(s), bin = '', CH = 0x8000;
    for (var i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return btoa(bin);
  }
  var wrap = function (b) { return b.replace(/.{1,76}/g, '$&\r\n'); };
  /* returns the base64 of a whole MIME message, which Graph sendMail takes with Content-Type: text/plain */
  function mime(built, from, to, cc) {
    var bd = 'bh-' + Math.random().toString(36).slice(2) + Date.now().toString(36);
    var hdr = function (n, v) { return n + ': ' + v + '\r\n'; };
    var m = hdr('From', from) + hdr('To', to.join(', ')) + (cc && cc.length ? hdr('Cc', cc.join(', ')) : '') + hdr('Subject', '=?UTF-8?B?' + b64utf8(built.subject) + '?=') + hdr('MIME-Version', '1.0') +
      hdr('Content-Type', 'multipart/alternative; boundary="' + bd + '"') + '\r\n' +
      '--' + bd + '\r\n' + hdr('Content-Type', 'text/plain; charset=UTF-8') + hdr('Content-Transfer-Encoding', 'base64') + '\r\n' + wrap(b64utf8(built.text)) +
      '--' + bd + '\r\n' + hdr('Content-Type', 'text/html; charset=UTF-8') + hdr('Content-Transfer-Encoding', 'base64') + '\r\n' + wrap(b64utf8(built.html)) +
      '--' + bd + '--\r\n';
    return b64utf8(m);
  }

  return { DEFAULT_TPL: DEFAULT_TPL, FIELDS: FIELDS, build: build, mime: mime };
});
