/* =========================================================
   StockSense F&B — Stock take for food & beverage outlets
   GitHub Pages + Power Automate + SharePoint
   ========================================================= */
'use strict';

/* ---------------- CONFIG (edit here) ---------------- */
var CONFIG = {
  COMPANY_NAME: 'Your Company',     // shown as logo text and on photo stamps
  COMPANY_SHORT: '',                // badge text, e.g. 'ACME'. Blank = initials
  CURRENCY: 'RM',
  AUTH_ENABLED: true,               // false = skip OTP (testing only)

  FLOW_SEND_OTP: 'https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/22/workflows/587bb4ec5a9141d4b2b61445d9841139/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dMgM0-vkKVd11vFfkQwMaUdsrnId4PhwMAI0ulkC9tE',
  FLOW_VERIFY_OTP: 'https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/31/workflows/296ff3060c1a449582e85c81e154c6df/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tdeev6xhU5HgYOU0zcHANo0FBK2kfDEqVrptgt1bkjQ',
  FLOW_CONFIG: 'PASTE_FB_GetConfig_HTTP_URL',     // returns session, zones, items
  FLOW_SUBMIT: 'PASTE_FB_SubmitCount_HTTP_URL',   // saves one count + photos
  FLOW_HISTORY: 'PASTE_FB_MyCounts_HTTP_URL',     // optional

  QR_LIBRARY: 'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js',
  UOMS: ['pcs', 'kg', 'g', 'L', 'ml', 'pack', 'box', 'carton', 'case', 'bottle', 'can', 'bag', 'tray', 'tub', 'roll', 'set', 'dozen', 'unit'],
  PHOTO_RULE: 'exceptions',         // 'always' or 'exceptions'
  MAX_PHOTOS: 5,
  NEAR_EXPIRY_DAYS: 3,
  RESEND_SECONDS: 60,               // fallback; flow expiresIn is used when present
  LOGIN_HOURS: 10,

  /* Categories offered when adding an item manually */
  CATEGORIES: ['Food – Fresh', 'Food – Frozen', 'Food – Dry goods', 'Beverage', 'Alcohol', 'Packaging & Disposables',
    'Cleaning & Chemicals', 'Smallwares & Utensils', 'Equipment', 'Linen & Uniforms', 'Stationery & Supplies', 'Other'],

  /* Storage areas used until FLOW_CONFIG returns your own.
     No items or business data are stored in this public file. */
  DEFAULT_ZONES: [
    { name: 'Chiller', tempMin: 0, tempMax: 5 },
    { name: 'Freezer', tempMin: -25, tempMax: -18 },
    { name: 'Dry Store' },
    { name: 'Bar' },
    { name: 'Kitchen' },
    { name: 'Back Store' }
  ]
};

/* ---------------- STATE ---------------- */
var S = {
  user: null, pendingEmail: '', busy: false, syncing: false,
  session: '', zones: [], items: [], zone: '',
  item: null, manual: false, condition: 'Good', photos: [],
  stream: null, facing: 'environment', stampTimer: null, scanner: null, scanDone: false,
  geo: null, log: [], resendTimer: null, flashTimer: null
};

/* ---------------- HELPERS ---------------- */
function $(id) { return document.getElementById(id); }
function on(id, ev, fn) { var e = $(id); if (e) { e.addEventListener(ev, fn); } }
function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function r3(v) { return Math.round(v * 1000) / 1000; }
function lc(s) { return String(s == null ? '' : s).toLowerCase(); }
function fmtQty(v) { return Number(r3(num(v))).toLocaleString('en-MY', { maximumFractionDigits: 3 }); }
function signed(v) { return (v > 0 ? '+' : (v < 0 ? '\u2212' : '')) + fmtQty(Math.abs(v)); }
function money(v) {
  var n = Number(v) || 0;
  return (n < 0 ? '\u2212' : '') + CONFIG.CURRENCY + ' ' + Math.abs(n).toLocaleString('en-MY', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function isSet(u) { return !!u && String(u).indexOf('PASTE_') !== 0; }
function uid() { return (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : 'c' + Date.now() + Math.random().toString(16).slice(2); }
function safeName(s) { return String(s || 'item').replace(/[^A-Za-z0-9-]/g, '_'); }
function buzz(p) { if (navigator.vibrate) { navigator.vibrate(p); } }
function fmtDT(iso) { return iso ? new Date(iso).toLocaleString('en-MY', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', hour12: false }) : ''; }
function el(tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) { e.className = cls; }
  if (text != null) { e.textContent = text; }
  return e;
}
function icon(name, cls) {
  var ns = 'http://www.w3.org/2000/svg';
  var s = document.createElementNS(ns, 'svg');
  s.setAttribute('class', 'ic' + (cls ? ' ' + cls : ''));
  var u = document.createElementNS(ns, 'use');
  u.setAttribute('href', '#i-' + name);
  s.appendChild(u);
  return s;
}
var store = {
  get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
  del: function (k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }
};

function toast(msg, type, ms) {
  var t = el('div', 'toast ' + (type || 'info'));
  t.appendChild(el('i'));
  t.appendChild(el('span', '', msg));
  var host = $('toastHost');
  while (host.children.length >= 2) { host.removeChild(host.firstChild); }
  host.appendChild(t);
  setTimeout(function () { t.style.opacity = '0'; setTimeout(function () { t.remove(); }, 300); }, ms || 3400);
}
function loading(onOff, text) {
  $('loaderText').textContent = text || 'Loading\u2026';
  $('loader').classList.toggle('hidden', !onOff);
}
function showView(id) {
  document.querySelectorAll('.view').forEach(function (v) { v.classList.toggle('active', v.id === id); });
  document.body.classList.toggle('on-count', id === 'viewApp' && $('tabCount').classList.contains('active'));
}
function loadScript(src) {
  return new Promise(function (ok, fail) {
    var s = document.createElement('script');
    s.src = src; s.onload = ok; s.onerror = function () { fail(new Error('Could not load ' + src)); };
    document.head.appendChild(s);
  });
}
async function api(url, body, timeoutMs) {
  if (!isSet(url)) { throw new Error('Flow URL is not configured'); }
  var ctrl = new AbortController();
  var timer = setTimeout(function () { ctrl.abort(); }, timeoutMs || 60000);
  try {
    var res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body), signal: ctrl.signal });
    var text = await res.text();
    var data = {};
    try { data = text ? JSON.parse(text) : {}; } catch (e) { data = { raw: text }; }
    if (!res.ok) { throw new Error(data.message || ('Request failed (' + res.status + ')')); }
    return data;
  } catch (e) {
    if (e.name === 'AbortError') { throw new Error('Request timed out'); }
    if (e instanceof TypeError) { throw new Error('Cannot reach the server. Check your connection.'); }
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

/* ---------------- THEME + BRANDING ---------------- */
function toggleTheme() {
  var t = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', t);
  var m = document.querySelector('meta[name="theme-color"]');
  if (m) { m.setAttribute('content', t === 'dark' ? '#0a0d12' : '#f3f4f6'); }
  store.set('ss_theme', t);
}
function applyBranding() {
  var name = CONFIG.COMPANY_NAME || 'Company';
  var short = CONFIG.COMPANY_SHORT || name.split(/\s+/).filter(Boolean).slice(0, 2).map(function (w) { return w.charAt(0); }).join('').toUpperCase();
  document.querySelectorAll('.js-company').forEach(function (e) { e.textContent = name; });
  document.querySelectorAll('.js-logo').forEach(function (e) { e.textContent = short; });
  document.title = name + ' \u00b7 StockSense';
  var sel = $('selManUom');
  CONFIG.UOMS.forEach(function (u) { sel.appendChild(new Option(u, u)); });
  var cat = $('selManCat');
  CONFIG.CATEGORIES.forEach(function (c) { cat.appendChild(new Option(c, c)); });
}
function greeting() {
  var h = new Date().getHours();
  return h < 12 ? 'Good morning' : (h < 18 ? 'Good afternoon' : 'Good evening');
}

/* ---------------- OFFLINE QUEUE ---------------- */
var DB = {
  db: null,
  open: function () {
    var self = this;
    return new Promise(function (ok, fail) {
      if (self.db) { return ok(self.db); }
      if (!window.indexedDB) { return fail(new Error('Offline storage not supported')); }
      var r = indexedDB.open('stocksense_fnb', 1);
      r.onupgradeneeded = function () { r.result.createObjectStore('queue', { keyPath: 'clientId' }); };
      r.onsuccess = function () { self.db = r.result; ok(self.db); };
      r.onerror = function () { fail(r.error); };
    });
  },
  run: async function (mode, fn) {
    var db = await this.open();
    return new Promise(function (ok, fail) {
      var tx = db.transaction('queue', mode);
      var req = fn(tx.objectStore('queue'));
      tx.oncomplete = function () { ok(req ? req.result : undefined); };
      tx.onerror = function () { fail(tx.error); };
    });
  },
  put: function (r) { return this.run('readwrite', function (s) { return s.put(r); }); },
  del: function (id) { return this.run('readwrite', function (s) { return s.delete(id); }); },
  all: function () { return this.run('readonly', function (s) { return s.getAll(); }); }
};

/* ---------------- LOCAL LOG ---------------- */
function logKey() { return 'ss_fnb_log_' + (S.session || 'none'); }
function loadLog() { try { S.log = JSON.parse(store.get(logKey()) || '[]'); } catch (e) { S.log = []; } }
function saveLog() { store.set(logKey(), JSON.stringify(S.log.slice(0, 3000))); }
function latest() {
  var seen = {}, out = [];
  S.log.forEach(function (x) {
    var k = (x.itemCode || x.itemName) + '|' + x.zone;
    if (!seen[k]) { seen[k] = true; out.push(x); }
  });
  return out;
}
function countedEntry(item, zone) {
  var list = latest();
  for (var i = 0; i !== list.length; i++) {
    var x = list[i];
    if (x.zone === zone && ((item.code && x.itemCode === item.code) || (!item.code && x.itemName === item.name))) { return x; }
  }
  return null;
}

/* ---------------- AUTH ---------------- */
function saveUser(email) {
  S.user = { email: email };
  store.set('ss_user', JSON.stringify({ email: email, exp: Date.now() + CONFIG.LOGIN_HOURS * 3600000 }));
}
function loadUser() {
  try { var u = JSON.parse(store.get('ss_user')); if (u && u.exp > Date.now()) { return u; } } catch (e) { /* ignore */ }
  store.del('ss_user');
  return null;
}
function startResendTimer(seconds) {
  var b = $('btnResend');
  var left = parseInt(seconds, 10) || CONFIG.RESEND_SECONDS;
  clearInterval(S.resendTimer);
  b.disabled = true;
  b.textContent = 'Code expires in ' + left + 's';
  S.resendTimer = setInterval(function () {
    left--;
    if (left > 0) { b.textContent = 'Code expires in ' + left + 's'; }
    else { clearInterval(S.resendTimer); b.disabled = false; b.textContent = 'Code expired \u00b7 Resend'; }
  }, 1000);
}
async function sendOtp(isResend) {
  if (S.busy) { return; }
  var email = isResend ? S.pendingEmail : $('inpEmail').value.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { return fail('inpEmail', 'Enter a valid email address'); }
  $('inpEmail').classList.remove('invalid');
  if (!CONFIG.AUTH_ENABLED) { saveUser(email); enterApp(); return; }
  S.busy = true;
  loading(true, 'Sending code\u2026');
  try {
    var r = await api(CONFIG.FLOW_SEND_OTP, { email: email });
    var st = lc(r.status);
    if (r.success === false || st === 'error' || st === 'failed' || st === 'blocked') { throw new Error(r.message || 'Unable to send code'); }
    S.pendingEmail = email;
    $('lblOtpEmail').textContent = email;
    $('stepEmail').classList.add('hidden');
    $('stepOtp').classList.remove('hidden');
    $('inpOtp').value = '';
    $('inpOtp').focus();
    startResendTimer(r.expiresIn);
    toast(isResend ? 'A new code has been sent' : 'Code sent. Check your inbox.', 'success');
  } catch (e) {
    toast(e.message, 'error', 6000);
  } finally {
    S.busy = false;
    loading(false);
  }
}
async function verifyOtp() {
  if (S.busy) { return; }
  var code = $('inpOtp').value.trim();
  if (!/^\d{6}$/.test(code)) { return fail('inpOtp', 'Enter the 6-digit code'); }
  S.busy = true;
  loading(true, 'Verifying\u2026');
  try {
    var r = await api(CONFIG.FLOW_VERIFY_OTP, { email: S.pendingEmail, code: code, otp: code });
    var st = lc(r.status);
    var ok = st === 'verified' || st === 'valid' || st === 'success' || r.success === true || r.verified === true || r.valid === true;
    if (!ok) {
      if (st === 'locked' || st === 'expired') {
        $('inpOtp').value = '';
        clearInterval(S.resendTimer);
        $('btnResend').disabled = false;
        $('btnResend').textContent = 'Resend code';
      }
      throw new Error(r.message || 'Invalid or expired code');
    }
    clearInterval(S.resendTimer);
    saveUser(S.pendingEmail);
    S.busy = false;
    loading(false);
    await enterApp();
  } catch (e) {
    S.busy = false;
    loading(false);
    fail('inpOtp', e.message);
  }
}
function logout() {
  stopCamera(); stopScan();
  store.del('ss_user');
  S.user = null;
  clearForm();
  $('inpEmail').value = '';
  $('stepOtp').classList.add('hidden');
  $('stepEmail').classList.remove('hidden');
  showView('viewLogin');
}

/* ---------------- BOOT ---------------- */
function normZones(list) {
  return (list || []).map(function (z) {
    if (typeof z === 'string') { return { name: z }; }
    return { name: z.name || z.Title || z.zone, tempMin: z.tempMin, tempMax: z.tempMax };
  }).filter(function (z) { return !!z.name; });
}
function normItems(list) {
  return (list || []).map(function (x) {
    if (Array.isArray(x)) { return { code: x[0], barcode: x[1], name: x[2], category: x[3], uom: x[4], zone: x[5], systemQty: x[6], unitCost: x[7] }; }
    return {
      code: String(x.code || x.itemCode || x.Title || ''),
      barcode: String(x.barcode || x.Barcode || ''),
      name: x.name || x.itemName || x.ItemName || '',
      category: x.category || x.Category || '',
      uom: x.uom || x.UOM || 'pcs',
      zone: x.zone || x.Zone || '',
      systemQty: (x.systemQty === '' || x.systemQty == null) ? null : num(x.systemQty),
      unitCost: num(x.unitCost)
    };
  }).filter(function (x) { return !!x.name; });
}
async function loadConfig() {
  var cached = null;
  try { cached = JSON.parse(store.get('ss_fnb_config') || 'null'); } catch (e) { /* ignore */ }
  function apply(c) { S.session = c.session || ''; S.zones = normZones(c.zones || c.locations); S.items = normItems(c.items); }

  if (!isSet(CONFIG.FLOW_CONFIG)) {
    apply({ session: 'ST-' + new Date().toISOString().slice(0, 7), zones: CONFIG.DEFAULT_ZONES, items: [] });
  } else if (navigator.onLine) {
    loading(true, 'Loading items\u2026');
    try {
      var r = await api(CONFIG.FLOW_CONFIG, { email: S.user.email });
      apply(r);
      store.set('ss_fnb_config', JSON.stringify({ session: S.session, zones: S.zones, items: S.items }));
    } catch (e) {
      if (cached) { apply(cached); toast('Using saved item list', 'warn'); }
      else { toast('Could not load items: ' + e.message, 'error', 6000); }
    } finally { loading(false); }
  } else if (cached) {
    apply(cached);
    toast('Offline. Using saved item list.', 'warn');
  }
  var pref = store.get('ss_fnb_zone');
  S.zone = (pref && zoneObj(pref)) ? pref : (S.zones[0] ? S.zones[0].name : '');
}
async function enterApp() {
  showView('viewApp');
  switchTab('tabCount');
  $('lblUser').textContent = S.user.email;
  $('lblGreeting').textContent = greeting();
  updateNet();
  await loadConfig();
  $('lblSession').textContent = S.session || 'No session';
  loadLog();
  renderZones();
  renderHero();
  renderPhotos();
  if (!S.items.length && !S.manual) { toggleManual(); $('btnManual').classList.add('hidden'); }
  getGeo();
  updateQueueBadge();
  updateAll();
  if (navigator.onLine) { syncQueue(true); }
}

/* ---------------- NETWORK ---------------- */
function updateNet() {
  var b = $('netBadge'), online = navigator.onLine;
  b.className = 'net ' + (online ? 'on' : 'off');
  b.querySelector('b').textContent = online ? 'Online' : 'Offline';
}

/* ---------------- ZONES ---------------- */
function zoneObj(name) { for (var i = 0; i !== S.zones.length; i++) { if (S.zones[i].name === name) { return S.zones[i]; } } return null; }
function zoneIconName(name) {
  var k = lc(name);
  if (k.indexOf('freez') !== -1) { return 'freeze'; }
  if (k.indexOf('chill') !== -1 || k.indexOf('cold') !== -1) { return 'snow'; }
  if (k.indexOf('bar') !== -1) { return 'glass'; }
  if (k.indexOf('kitchen') !== -1 || k.indexOf('line') !== -1) { return 'flame'; }
  if (k.indexOf('dry') !== -1 || k.indexOf('store') !== -1) { return 'grain'; }
  return 'box';
}
function zoneCounts(name) {
  return {
    total: S.items.filter(function (i) { return i.zone === name; }).length,
    done: latest().filter(function (x) { return x.zone === name; }).length
  };
}
function renderZones() {
  var box = $('zoneList');
  box.textContent = '';
  S.zones.forEach(function (z) {
    var c = zoneCounts(z.name);
    var b = el('button', 'zone' + (z.name === S.zone ? ' active' : ''));
    b.type = 'button';
    var top = el('span', 'zone-top');
    top.appendChild(icon(zoneIconName(z.name)));
    top.appendChild(el('b', '', z.name));
    b.appendChild(top);
    b.appendChild(el('span', 'zone-meta', c.total ? c.done + ' of ' + c.total + ' counted' : c.done + ' counted'));
    var tr = el('span', 'zone-track'), f = el('i');
    f.style.width = (c.total ? Math.min(100, Math.round(c.done / c.total * 100)) : 0) + '%';
    tr.appendChild(f);
    b.appendChild(tr);
    b.addEventListener('click', function () { setZone(z.name); });
    box.appendChild(b);
  });
  var cur = zoneCounts(S.zone);
  $('lblZone').textContent = cur.total ? (cur.total - cur.done) + ' left in ' + S.zone : (S.zone || '');
}
function setZone(name) {
  S.zone = name;
  store.set('ss_fnb_zone', name);
  if (S.item && S.item.zone && S.item.zone !== name) { clearItem(); }
  renderZones();
  doSearch();
  updateAll();
}

/* ---------------- ITEMS ---------------- */
function doSearch() {
  var q = lc($('inpSearch').value.trim());
  var box = $('searchResults');
  box.textContent = '';
  if (!q) { return; }
  var list = S.items.filter(function (i) {
    return [i.name, i.code, i.barcode, i.category].some(function (v) { return lc(v).indexOf(q) !== -1; });
  }).sort(function (a, b) { return (a.zone === S.zone ? 0 : 1) - (b.zone === S.zone ? 0 : 1); }).slice(0, 20);
  if (!list.length) {
    box.appendChild(el('div', 'res-empty', S.items.length ? 'No match. Use \u201cAdd manually\u201d below.' : 'No item list connected yet. Use \u201cAdd manually\u201d below.'));
    return;
  }
  list.forEach(function (it) {
    var b = el('button', 'res');
    b.type = 'button';
    var m = el('span');
    m.appendChild(el('b', '', it.name));
    m.appendChild(el('small', '', it.code + ' \u00b7 ' + it.uom + (it.category ? ' \u00b7 ' + it.category : '')));
    b.appendChild(m);
    var done = countedEntry(it, it.zone || S.zone);
    b.appendChild(el('span', 'res-tag' + (done ? ' done' : ''), done ? 'Counted' : (it.zone || '')));
    b.addEventListener('click', function () { selectItem(it); });
    box.appendChild(b);
  });
}
function selectItem(it) {
  S.item = it;
  S.manual = false;
  $('manualWrap').classList.add('hidden');
  $('btnManual').textContent = '+ Item not in list? Add manually';
  $('searchResults').textContent = '';
  $('inpSearch').value = '';
  if (it.zone && it.zone !== S.zone && zoneObj(it.zone)) {
    S.zone = it.zone;
    store.set('ss_fnb_zone', it.zone);
    renderZones();
    toast('Switched to ' + it.zone);
  }
  $('icCat').textContent = it.category || 'Item';
  $('icName').textContent = it.name;
  $('icCode').textContent = [it.code, it.barcode].filter(Boolean).join(' \u00b7 ');
  $('icExp').textContent = it.systemQty == null ? '\u2014' : fmtQty(it.systemQty) + ' ' + it.uom;
  $('icCost').textContent = it.unitCost ? money(it.unitCost) : '\u2014';
  $('icZone').textContent = S.zone;
  var prev = countedEntry(it, S.zone);
  if (prev) {
    $('icRecount').textContent = 'Already counted ' + fmtQty(prev.countedQty) + ' ' + prev.uom + ' at ' + fmtDT(prev.submittedAt) + '. Submitting again records a recount.';
    $('icRecount').classList.remove('hidden');
  } else {
    $('icRecount').classList.add('hidden');
  }
  $('itemCard').classList.remove('hidden');
  $('inpQty').value = '';
  buzz(15);
  updateAll();
  setTimeout(function () { $('inpQty').focus(); }, 120);
}
function clearItem() {
  S.item = null;
  $('itemCard').classList.add('hidden');
  updateAll();
}
function toggleManual() {
  S.manual = !S.manual;
  $('manualWrap').classList.toggle('hidden', !S.manual);
  $('btnManual').textContent = S.manual ? '\u2190 Back to item list' : '+ Item not in list? Add manually';
  if (S.manual) {
    clearItem();
    var q = $('inpSearch').value.trim();
    if (q) { if (/^\d{6,}$/.test(q)) { $('inpManCode').value = q; } else { $('inpManName').value = q; } }
    $('searchResults').textContent = '';
    $('inpManName').focus();
  }
  updateAll();
}
function currentItem() {
  if (S.manual) {
    var n = $('inpManName').value.trim();
    if (!n) { return null; }
    return { code: $('inpManCode').value.trim().toUpperCase(), barcode: '', name: n, category: $('selManCat').value, uom: $('selManUom').value, zone: S.zone, systemQty: null, unitCost: 0, manual: true };
  }
  return S.item;
}

/* ---------------- COUNT ---------------- */
function isWeight(u) { return ['kg', 'l'].indexOf(lc(u)) !== -1; }
function isSmall(u) { return ['g', 'ml'].indexOf(lc(u)) !== -1; }
function stepSize() { var it = currentItem(); if (!it) { return 1; } return isWeight(it.uom) ? 0.1 : (isSmall(it.uom) ? 50 : 1); }
function bump(d) {
  $('inpQty').value = Math.max(0, r3(num($('inpQty').value) + d));
  $('inpQty').classList.remove('invalid');
  buzz(6);
  updateAll();
}
function qtyValid() { var q = String($('inpQty').value).trim(); return q !== '' && !isNaN(q) && num(q) >= 0; }
function variance() {
  var it = currentItem();
  if (!it || it.systemQty == null || !qtyValid()) { return null; }
  var v = r3(num($('inpQty').value) - it.systemQty);
  return { qty: v, value: Math.round(v * (it.unitCost || 0) * 100) / 100 };
}
function renderChips() {
  var it = currentItem();
  var uom = it ? it.uom : 'unit';
  $('qtyUom').textContent = uom;
  var steps = isWeight(uom) ? [0.1, 0.5, 1, 5] : (isSmall(uom) ? [100, 250, 500, 1000] : [1, 5, 10, 24]);
  var box = $('qtyChips');
  box.textContent = '';
  steps.forEach(function (s) {
    var b = el('button', 'chip', '+' + s);
    b.type = 'button';
    b.addEventListener('click', function () { bump(s); });
    box.appendChild(b);
  });
  var r = el('button', 'chip', 'Clear');
  r.type = 'button';
  r.addEventListener('click', function () { $('inpQty').value = ''; updateAll(); });
  box.appendChild(r);
}
function renderVariance() {
  var p = $('varPanel'), it = currentItem();
  if (!it || it.systemQty == null) { p.classList.add('hidden'); return; }
  p.classList.remove('hidden');
  var v = variance();
  $('vpExp').textContent = fmtQty(it.systemQty) + ' ' + it.uom;
  $('vpCnt').textContent = qtyValid() ? fmtQty($('inpQty').value) + ' ' + it.uom : '\u2014';
  p.className = 'variance';
  if (!v) { $('vpVar').textContent = '\u2014'; $('vpVal').textContent = '\u2014'; return; }
  $('vpVar').textContent = v.qty === 0 ? 'Match' : signed(v.qty);
  $('vpVal').textContent = money(v.value);
  p.classList.add(v.qty === 0 ? 'ok' : (v.qty < 0 ? 'short' : 'over'));
}
function setCondition(c) {
  S.condition = c;
  document.querySelectorAll('#condGroup .seg').forEach(function (s) { s.classList.toggle('active', s.dataset.cond === c); });
  updateAll();
}
function onExpiry() {
  var v = $('inpExpiry').value, b = $('expBadge');
  if (!v) { b.classList.add('hidden'); updateAll(); return; }
  var t = new Date(); t.setHours(0, 0, 0, 0);
  var days = Math.round((new Date(v + 'T00:00:00') - t) / 86400000);
  b.classList.remove('hidden');
  if (days < 0) {
    b.className = 'pill bad';
    b.textContent = 'Expired ' + Math.abs(days) + 'd ago';
    if (S.condition === 'Good' || S.condition === 'Near Expiry') { setCondition('Expired'); toast('Condition set to Expired', 'warn'); }
  } else if (days <= CONFIG.NEAR_EXPIRY_DAYS) {
    b.className = 'pill warn';
    b.textContent = days === 0 ? 'Expires today' : 'Expires in ' + days + 'd';
    if (S.condition === 'Good') { setCondition('Near Expiry'); toast('Condition set to Near expiry', 'warn'); }
  } else {
    b.className = 'pill ok';
    b.textContent = days + ' days left';
    if (S.condition === 'Near Expiry' || S.condition === 'Expired') { setCondition('Good'); }
  }
  updateAll();
}
function tempRange() {
  var z = zoneObj(S.zone), k = lc(S.zone);
  if (z && (z.tempMin != null || z.tempMax != null)) { return [z.tempMin, z.tempMax]; }
  if (k.indexOf('freez') !== -1) { return [-25, -18]; }
  if (k.indexOf('chill') !== -1 || k.indexOf('cold') !== -1) { return [0, 5]; }
  return null;
}
function renderTemp() {
  var rg = tempRange();
  $('tempWrap').classList.toggle('hidden', !rg);
  if (!rg) { return; }
  var b = $('tempBadge'), v = $('inpTemp').value;
  if (v === '') { b.className = 'pill'; b.textContent = 'Target ' + rg[0] + ' to ' + rg[1] + '\u00b0C'; return; }
  var t = num(v), ok = t >= rg[0] && t <= rg[1];
  b.className = 'pill ' + (ok ? 'ok' : 'bad');
  b.textContent = ok ? 'In range' : 'Out of range';
}
function isException() { var v = variance(); return S.condition !== 'Good' || !!(v && v.qty !== 0); }
function remarkNeeded() { var v = variance(); return ['Expired', 'Spoiled', 'Damaged', 'Missing'].indexOf(S.condition) !== -1 || !!(v && v.qty !== 0); }
function photosNeeded() { return (CONFIG.PHOTO_RULE === 'always' || isException()) ? 1 : 0; }

/* ---------------- READINESS ---------------- */
function updateAll() {
  var it = currentItem();
  renderChips();
  renderVariance();
  renderTemp();
  var needP = photosNeeded(), needR = remarkNeeded();
  $('photoReq').textContent = needP ? 'At least one photo is required for this count.' : 'Optional for a matching count. Required for variances and problem stock.';
  $('photoReq').classList.toggle('on', !!needP);
  $('remarkReq').classList.toggle('hidden', !needR);

  var s1 = !!(S.zone && it);
  var s2 = qtyValid() && (!needR || !!$('inpRemarks').value.trim());
  var s3 = S.photos.length >= needP;
  [s1, s2, s3].forEach(function (ok, i) { $('ps' + (i + 1)).classList.toggle('done', ok); });
  var done = [s1, s2, s3].filter(Boolean).length;
  $('progBar').style.width = Math.round(done / 3 * 100) + '%';
  $('btnSubmit').classList.toggle('ready', done === 3);

  var v = variance();
  $('sumZone').textContent = S.zone || '\u2014';
  $('sumItem').textContent = it ? it.name : '\u2014';
  $('sumQty').textContent = (it && qtyValid() ? fmtQty($('inpQty').value) + ' ' + it.uom + ' \u00b7 ' : '') + S.condition;
  $('sumVar').textContent = v ? (v.qty === 0 ? 'Match' : signed(v.qty) + '  ' + money(v.value)) : '\u2014';
  $('sumPhotos').textContent = S.photos.length + ' / ' + CONFIG.MAX_PHOTOS;
}
function clearForm() {
  stopCamera();
  S.item = null; S.manual = false; S.photos = []; S.condition = 'Good';
  $('itemCard').classList.add('hidden');
  $('manualWrap').classList.add('hidden');
  $('btnManual').textContent = '+ Item not in list? Add manually';
  if (!S.items.length && S.user) { S.manual = true; $('manualWrap').classList.remove('hidden'); }
  ['inpSearch', 'inpManName', 'inpManCode', 'inpQty', 'inpExpiry', 'inpBatch', 'inpTemp', 'inpRemarks'].forEach(function (id) {
    $(id).value = ''; $(id).classList.remove('invalid');
  });
  $('searchResults').textContent = '';
  $('expBadge').classList.add('hidden');
  document.querySelectorAll('#condGroup .seg').forEach(function (s) { s.classList.toggle('active', s.dataset.cond === 'Good'); });
  renderPhotos();
}

/* ---------------- SCANNER ---------------- */
async function startScan() {
  stopCamera();
  await stopScan();
  if (!window.Html5Qrcode) {
    loading(true, 'Starting scanner\u2026');
    try { await loadScript(CONFIG.QR_LIBRARY); } catch (e) { loading(false); toast('Scanner could not load', 'error'); return; }
    loading(false);
  }
  $('scanWrap').classList.remove('hidden');
  S.scanDone = false;
  S.scanner = new window.Html5Qrcode('reader');
  try {
    await S.scanner.start({ facingMode: 'environment' },
      { fps: 12, qrbox: function (w, h) { return { width: Math.max(160, Math.floor(w * 0.85)), height: Math.max(90, Math.floor(h * 0.4)) }; } },
      onScan, function () {});
  } catch (e) {
    toast('Camera not available for scanning', 'error');
    stopScan();
  }
}
async function stopScan() {
  if (S.scanner) { try { if (S.scanner.isScanning) { await S.scanner.stop(); } S.scanner.clear(); } catch (e) { /* ignore */ } S.scanner = null; }
  $('scanWrap').classList.add('hidden');
}
async function onScan(text) {
  if (S.scanDone) { return; }
  S.scanDone = true;
  await stopScan();
  buzz(60);
  var code = String(text).trim();
  var hit = S.items.filter(function (i) { return i.barcode === code || lc(i.code) === lc(code); })[0];
  if (hit) { selectItem(hit); return; }
  toast('Barcode not in item list. Add it manually.', 'warn', 4500);
  if (!S.manual) { toggleManual(); }
  $('inpManCode').value = code;
  $('inpManName').focus();
}

/* ---------------- CAMERA ---------------- */
function stampLines() {
  var it = currentItem();
  var ts = new Date().toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
  var gps = S.geo ? '  ' + S.geo.lat.toFixed(5) + ', ' + S.geo.lng.toFixed(5) : '';
  return [
    (CONFIG.COMPANY_NAME || '').toUpperCase() + ' \u00b7 ' + (S.session || '-'),
    (S.zone || '-') + ' \u00b7 ' + (it ? (it.code ? it.code + ' ' : '') + it.name : '-'),
    (it && qtyValid() ? fmtQty($('inpQty').value) + ' ' + it.uom : '-') + ' \u00b7 ' + S.condition + ($('inpExpiry').value ? ' \u00b7 Exp ' + $('inpExpiry').value : ''),
    ((S.user && S.user.email) || '-') + ' \u00b7 ' + ts + gps
  ];
}
function updateStamp() { $('camStamp').textContent = stampLines().join('\n'); }
async function openCamera() {
  if (S.photos.length >= CONFIG.MAX_PHOTOS) { toast('Maximum ' + CONFIG.MAX_PHOTOS + ' photos', 'warn'); return; }
  await stopScan();
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) { $('fileCamera').click(); return; }
  try {
    stopCamera();
    S.stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: S.facing }, width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false });
    var v = $('camVideo');
    v.srcObject = S.stream;
    $('camWrap').classList.remove('hidden');
    await v.play();
    updateStamp();
    S.stampTimer = setInterval(updateStamp, 1000);
    $('camWrap').scrollIntoView({ behavior: 'smooth', block: 'center' });
  } catch (e) {
    stopCamera();
    toast('Camera blocked. Opening device camera instead.', 'warn');
    $('fileCamera').click();
  }
}
function stopCamera() {
  if (S.stream) { S.stream.getTracks().forEach(function (t) { t.stop(); }); S.stream = null; }
  if (S.stampTimer) { clearInterval(S.stampTimer); S.stampTimer = null; }
  var v = $('camVideo'); if (v) { v.srcObject = null; }
  $('camWrap').classList.add('hidden');
}
async function switchCamera() {
  S.facing = S.facing === 'environment' ? 'user' : 'environment';
  if (S.stream) { stopCamera(); await openCamera(); }
}
function processImage(src, w, h) {
  var scale = Math.min(1, 1600 / Math.max(w, h));
  var cw = Math.round(w * scale), ch = Math.round(h * scale);
  var c = $('workCanvas'); c.width = cw; c.height = ch;
  var ctx = c.getContext('2d');
  ctx.drawImage(src, 0, 0, cw, ch);
  var lines = stampLines();
  var fs = Math.max(13, Math.round(cw / 54)), pad = Math.round(fs * 0.8), lh = Math.round(fs * 1.45);
  var bh = pad * 2 + lh * lines.length;
  ctx.fillStyle = 'rgba(11,18,32,0.78)';
  ctx.fillRect(0, ch - bh, cw, bh);
  ctx.fillStyle = '#2ccf93';
  ctx.fillRect(0, ch - bh, Math.max(4, Math.round(fs / 3)), bh);
  ctx.font = '600 ' + fs + 'px "JetBrains Mono", Menlo, monospace';
  ctx.textBaseline = 'top';
  lines.forEach(function (t, i) {
    ctx.fillStyle = i === 0 ? '#2ccf93' : '#ffffff';
    ctx.fillText(t, pad + Math.round(fs / 2), ch - bh + pad + i * lh);
  });
  return c.toDataURL('image/jpeg', 0.75);
}
function capture() {
  var v = $('camVideo');
  if (!v.videoWidth) { toast('Camera is starting\u2026', 'warn'); return; }
  addPhoto(processImage(v, v.videoWidth, v.videoHeight));
  var f = $('camFlash'); f.classList.add('on'); setTimeout(function () { f.classList.remove('on'); }, 60);
  buzz(30);
  if (S.photos.length >= CONFIG.MAX_PHOTOS) { stopCamera(); toast('Photo limit reached'); }
}
function loadImg(file) {
  return new Promise(function (ok, bad) {
    var url = URL.createObjectURL(file), img = new Image();
    img.onload = function () { URL.revokeObjectURL(url); ok(img); };
    img.onerror = function () { URL.revokeObjectURL(url); bad(new Error('Unreadable image')); };
    img.src = url;
  });
}
async function onFiles(e) {
  var files = Array.prototype.slice.call(e.target.files || []);
  e.target.value = '';
  for (var i = 0; i !== files.length; i++) {
    if (S.photos.length >= CONFIG.MAX_PHOTOS) { toast('Maximum ' + CONFIG.MAX_PHOTOS + ' photos', 'warn'); break; }
    if (files[i].type.indexOf('image/') !== 0) { continue; }
    try { var img = await loadImg(files[i]); addPhoto(processImage(img, img.naturalWidth, img.naturalHeight)); }
    catch (err) { toast('Could not read ' + files[i].name, 'error'); }
  }
}
function addPhoto(data) { S.photos.push({ id: uid(), data: data }); renderPhotos(); }
function renderPhotos() {
  var g = $('photoGrid');
  g.textContent = '';
  S.photos.forEach(function (p) {
    var t = el('div', 'thumb');
    var img = document.createElement('img'); img.src = p.data; img.alt = 'Evidence photo';
    var x = el('button'); x.type = 'button'; x.setAttribute('aria-label', 'Remove photo'); x.appendChild(icon('x'));
    x.addEventListener('click', function () { S.photos = S.photos.filter(function (q) { return q.id !== p.id; }); renderPhotos(); });
    t.appendChild(img); t.appendChild(x);
    g.appendChild(t);
  });
  $('photoCounter').textContent = S.photos.length + '/' + CONFIG.MAX_PHOTOS;
  var full = S.photos.length >= CONFIG.MAX_PHOTOS;
  $('btnOpenCam').disabled = full;
  $('btnPickPhoto').disabled = full;
  updateAll();
}

/* ---------------- GPS ---------------- */
function getGeo() {
  var b = $('geoBadge');
  function setText(t, ok) { b.textContent = ''; b.appendChild(icon('pin', 'xs')); b.appendChild(document.createTextNode(t)); b.className = ok ? 'ok' : ''; }
  if (!navigator.geolocation) { setText('Location not supported'); return; }
  setText('Locating\u2026');
  navigator.geolocation.getCurrentPosition(function (p) {
    S.geo = { lat: p.coords.latitude, lng: p.coords.longitude, acc: Math.round(p.coords.accuracy) };
    setText('Location captured \u00b1' + S.geo.acc + ' m', true);
  }, function () { setText('Location unavailable (optional)'); }, { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 });
}

/* ---------------- SUBMIT ---------------- */
function fail(id, msg) {
  if (id && $(id)) { $(id).classList.add('invalid'); $(id).focus(); }
  toast(msg, 'error');
  return false;
}
function validate() {
  if (!S.session) { return fail(null, 'No active session. Contact your supervisor.'); }
  if (!S.zone) { return fail(null, 'Choose a storage zone'); }
  if (!currentItem()) { return S.manual ? fail('inpManName', 'Enter the item name') : fail('inpSearch', 'Search or scan an item first'); }
  if (!qtyValid()) { return fail('inpQty', 'Enter the quantity on hand'); }
  if (remarkNeeded() && !$('inpRemarks').value.trim()) { return fail('inpRemarks', 'Add a remark explaining the variance or condition'); }
  if (S.photos.length < photosNeeded()) { $('btnOpenCam').scrollIntoView({ behavior: 'smooth', block: 'center' }); return fail(null, 'Add at least one photo as evidence'); }
  return true;
}
function flash(title, sub, offline) {
  $('flashTitle').textContent = title;
  $('flashSub').textContent = sub;
  $('flashIc').classList.toggle('off', !!offline);
  var f = $('flash');
  f.classList.remove('show'); void f.offsetWidth; f.classList.add('show');
  clearTimeout(S.flashTimer);
  S.flashTimer = setTimeout(function () { f.classList.remove('show'); }, 2600);
}
async function submitCount() {
  if (S.busy || !validate()) { return; }
  S.busy = true;
  stopCamera();
  var it = currentItem(), v = variance(), stamp = Date.now();
  var rec = {
    clientId: uid(), session: S.session, zone: S.zone,
    itemCode: it.code || '', barcode: it.barcode || '', itemName: it.name, category: it.category || '', uom: it.uom,
    isManualItem: !!it.manual, systemQty: it.systemQty, countedQty: r3(num($('inpQty').value)),
    variance: v ? v.qty : null, unitCost: it.unitCost || 0, varianceValue: v ? v.value : null,
    condition: S.condition, expiryDate: $('inpExpiry').value || '', batchNo: $('inpBatch').value.trim(),
    temperature: $('inpTemp').value === '' ? null : num($('inpTemp').value),
    remarks: $('inpRemarks').value.trim(), submittedEmail: S.user.email, submittedAt: new Date().toISOString(),
    latitude: S.geo ? String(S.geo.lat) : '', longitude: S.geo ? String(S.geo.lng) : '',
    photos: S.photos.map(function (p, i) { return { fileName: safeName(it.code || it.name) + '_' + stamp + '_' + (i + 1) + '.jpg', content: p.data.split(',')[1] }; })
  };
  var sub = it.name + ' \u00b7 ' + fmtQty(rec.countedQty) + ' ' + rec.uom + (v && v.qty !== 0 ? ' \u00b7 ' + signed(v.qty) : '');

  function finish(id, status, offline) {
    var copy = {}; for (var k in rec) { if (k !== 'photos') { copy[k] = rec[k]; } }
    copy.id = id || ''; copy.status = status;
    S.log.unshift(copy);
    saveLog();
    clearForm();
    renderZones();
    renderHero();
    updateAll();
    flash(offline ? 'Saved offline' : 'Counted' + (id ? ' \u00b7 ' + id : ''), sub, offline);
    buzz(offline ? [40, 40, 40] : [25, 40, 25]);
    window.scrollTo({ top: 0, behavior: 'smooth' });
    setTimeout(function () { ($('manualWrap').classList.contains('hidden') ? $('inpSearch') : $('inpManName')).focus(); }, 450);
  }

  if (!isSet(CONFIG.FLOW_SUBMIT)) { S.busy = false; finish('TEST-' + String(S.log.length + 1).padStart(4, '0'), 'Test', false); return; }
  if (!navigator.onLine) {
    try { await DB.put(rec); finish('', 'Queued', true); } catch (e) { toast('Could not save offline: ' + e.message, 'error'); }
    S.busy = false; updateQueueBadge(); return;
  }
  loading(true, 'Uploading count\u2026');
  try {
    var r = await api(CONFIG.FLOW_SUBMIT, rec, 120000);
    if (r.success === false) { throw new Error(r.message || 'Submission failed'); }
    finish(r.stockTakeId || '', 'Submitted', false);
  } catch (e) {
    try { await DB.put(rec); finish('', 'Queued', true); } catch (err) { toast('Submission failed: ' + e.message, 'error', 6000); }
  } finally {
    S.busy = false; loading(false); updateQueueBadge();
  }
}

/* ---------------- SYNC + QUEUE ---------------- */
async function syncQueue(silent) {
  if (S.syncing || !isSet(CONFIG.FLOW_SUBMIT)) { if (!silent && !isSet(CONFIG.FLOW_SUBMIT)) { toast('Submit flow not connected yet', 'warn'); } return; }
  if (!navigator.onLine) { if (!silent) { toast('Still offline', 'warn'); } return; }
  var all = []; try { all = await DB.all(); } catch (e) { return; }
  if (!all.length) { if (!silent) { toast('Nothing to sync'); } return; }
  S.syncing = true;
  if (!silent) { loading(true, 'Syncing ' + all.length + ' count(s)\u2026'); }
  var ok = 0, bad = 0;
  for (var i = 0; i !== all.length; i++) {
    try {
      var r = await api(CONFIG.FLOW_SUBMIT, all[i], 120000);
      if (r.success === false) { throw new Error(r.message); }
      await DB.del(all[i].clientId);
      S.log.forEach(function (x) { if (x.clientId === all[i].clientId) { x.status = 'Submitted'; x.id = r.stockTakeId || ''; } });
      ok++;
    } catch (e) { bad++; }
  }
  saveLog();
  S.syncing = false;
  if (!silent) { loading(false); }
  updateQueueBadge();
  if ($('tabQueue').classList.contains('active')) { renderQueue(); }
  if (ok) { toast(ok + ' count(s) synced', 'success'); }
  if (bad) { toast(bad + ' still pending', 'warn'); }
}
function emptyState(title, text) { var e = el('div', 'empty'); e.appendChild(el('b', '', title)); e.appendChild(document.createTextNode(text)); return e; }
async function renderQueue() {
  var all = []; try { all = await DB.all(); } catch (e) { /* ignore */ }
  var box = $('queueList');
  box.textContent = '';
  if (!all.length) { box.appendChild(emptyState('All caught up', 'Nothing is waiting to sync.')); return; }
  all.forEach(function (r) {
    var row = el('div', 'row'), m = el('div', 'row-main');
    m.appendChild(el('b', '', r.itemName));
    m.appendChild(el('small', '', r.zone + ' \u00b7 ' + fmtQty(r.countedQty) + ' ' + r.uom + ' \u00b7 ' + r.photos.length + ' photo(s) \u00b7 ' + fmtDT(r.submittedAt)));
    var d = el('button', 'del', 'Discard'); d.type = 'button';
    d.addEventListener('click', async function () {
      if (!confirm('Discard this queued count? This cannot be undone.')) { return; }
      await DB.del(r.clientId); renderQueue(); updateQueueBadge();
    });
    row.appendChild(m); row.appendChild(d);
    box.appendChild(row);
  });
}
async function updateQueueBadge() {
  try {
    var n = (await DB.all()).length;
    $('queueBadge').textContent = n;
    $('queueBadge').classList.toggle('hidden', !n);
    $('kpiQueue').textContent = n;
  } catch (e) { /* ignore */ }
}

/* ---------------- HERO + PROGRESS ---------------- */
function renderHero() {
  var list = latest();
  var val = list.reduce(function (s, x) { return s + (x.varianceValue ? num(x.varianceValue) : 0); }, 0);
  var risk = list.filter(function (x) { return ['Near Expiry', 'Expired', 'Spoiled'].indexOf(x.condition) !== -1; }).length;
  $('hsCount').textContent = list.length;
  $('hsVar').textContent = money(val);
  $('hsRisk').textContent = risk;
}
function renderProgress() {
  var list = latest(), total = S.items.length, counted = list.length;
  var pct = total ? Math.min(100, Math.round(counted / total * 100)) : 0;
  var withExp = list.filter(function (x) { return x.variance != null; });
  var matched = withExp.filter(function (x) { return num(x.variance) === 0; }).length;
  var val = list.reduce(function (s, x) { return s + (x.varianceValue ? num(x.varianceValue) : 0); }, 0);
  var risk = list.filter(function (x) { return ['Near Expiry', 'Expired', 'Spoiled'].indexOf(x.condition) !== -1; }).length;

  $('ringPct').textContent = pct + '%';
  $('ringCap').textContent = total ? counted + ' of ' + total + ' listed items' : counted + ' items';
  var circ = 2 * Math.PI * 34;
  $('ringVal').setAttribute('stroke-dasharray', circ.toFixed(1));
  $('ringVal').setAttribute('stroke-dashoffset', (circ * (1 - pct / 100)).toFixed(1));
  $('kpiAcc').textContent = withExp.length ? Math.round(matched / withExp.length * 100) + '%' : '\u2014';
  $('kpiVal').textContent = money(val);
  $('kpiVal').className = 'mono ' + (val < 0 ? 'neg' : (val > 0 ? 'pos' : ''));
  $('kpiRisk').textContent = risk;
  $('kpiRisk').className = 'mono ' + (risk ? 'neg' : '');

  var zb = $('zoneBars');
  zb.textContent = '';
  S.zones.forEach(function (z) {
    var c = zoneCounts(z.name), p = c.total ? Math.min(100, Math.round(c.done / c.total * 100)) : 0;
    var row = el('div', 'zb'), top = el('div', 'zb-top');
    top.appendChild(el('b', '', z.name));
    top.appendChild(el('span', 'mono', c.total ? c.done + '/' + c.total + '  ' + p + '%' : c.done + ' counted'));
    var tr = el('div', 'zb-track'), f = el('i');
    tr.appendChild(f); row.appendChild(top); row.appendChild(tr); zb.appendChild(row);
    setTimeout(function () { f.style.width = p + '%'; }, 40);
  });

  var q = lc($('inpFilter').value.trim());
  var rows = q ? list.filter(function (x) { return [x.itemName, x.itemCode, x.zone].some(function (v) { return lc(v).indexOf(q) !== -1; }); }) : list;
  var box = $('historyList');
  box.textContent = '';
  if (!rows.length) { box.appendChild(list.length ? emptyState('No matches', 'Try another item, code or zone.') : emptyState('Nothing counted yet', 'Counts from this session will appear here.')); return; }
  rows.slice(0, 300).forEach(function (x) {
    var row = el('div', 'row'), m = el('div', 'row-main');
    m.appendChild(el('b', '', x.itemName));
    m.appendChild(el('small', '', [x.itemCode, x.zone, fmtDT(x.submittedAt), x.id].filter(Boolean).join(' \u00b7 ')));
    var end = el('div', 'row-end');
    var vq = x.variance == null ? null : num(x.variance);
    end.appendChild(el('span', 'mono ' + (vq == null ? '' : (vq < 0 ? 'neg' : (vq > 0 ? 'pos' : 'zero'))), fmtQty(x.countedQty) + ' ' + x.uom + (vq ? '  ' + signed(vq) : '')));
    var tagTxt = x.status === 'Queued' ? 'Queued' : x.condition;
    end.appendChild(el('span', 'tag ' + String(tagTxt).split(' ')[0], tagTxt));
    row.appendChild(m); row.appendChild(end);
    box.appendChild(row);
  });
}

/* ---------------- TABS ---------------- */
function switchTab(id) {
  document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.id === id); });
  document.querySelectorAll('.dock-btn').forEach(function (b) { b.classList.toggle('active', b.dataset.tab === id); });
  document.body.classList.toggle('on-count', id === 'tabCount');
  if (id !== 'tabCount') { stopCamera(); stopScan(); }
  if (id === 'tabProgress') { renderProgress(); updateQueueBadge(); }
  if (id === 'tabQueue') { renderQueue(); }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ---------------- EVENTS ---------------- */
function bind() {
  document.querySelectorAll('.js-theme').forEach(function (b) { b.addEventListener('click', toggleTheme); });

  on('btnSendOtp', 'click', function () { sendOtp(false); });
  on('btnResend', 'click', function () { sendOtp(true); });
  on('btnVerifyOtp', 'click', verifyOtp);
  on('btnBackEmail', 'click', function () { clearInterval(S.resendTimer); $('stepOtp').classList.add('hidden'); $('stepEmail').classList.remove('hidden'); });
  on('inpEmail', 'keydown', function (e) { if (e.key === 'Enter') { sendOtp(false); } });
  on('inpEmail', 'input', function () { $('inpEmail').classList.remove('invalid'); });
  on('inpOtp', 'input', function (e) {
    e.target.classList.remove('invalid');
    e.target.value = e.target.value.replace(/\D/g, '').slice(0, 6);
    if (e.target.value.length === 6) { verifyOtp(); }
  });
  on('btnLogout', 'click', function () { if (confirm('Sign out of StockSense?')) { logout(); } });

  on('inpSearch', 'input', function () { $('inpSearch').classList.remove('invalid'); doSearch(); });
  on('inpSearch', 'keydown', function (e) { if (e.key === 'Enter') { var f = $('searchResults').querySelector('.res'); if (f) { f.click(); } } });
  on('btnClearItem', 'click', function () { clearItem(); $('inpSearch').focus(); });
  on('btnManual', 'click', toggleManual);
  on('inpManName', 'input', function () { $('inpManName').classList.remove('invalid'); updateAll(); });
  on('inpManCode', 'input', updateAll);
  on('selManUom', 'change', updateAll);
  on('selManCat', 'change', updateAll);
  on('btnScan', 'click', startScan);
  on('btnStopScan', 'click', stopScan);

  on('btnMinus', 'click', function () { bump(-stepSize()); });
  on('btnPlus', 'click', function () { bump(stepSize()); });
  on('inpQty', 'input', function () { $('inpQty').classList.remove('invalid'); updateAll(); });
  on('inpQty', 'focus', function (e) { e.target.select(); });
  document.querySelectorAll('#condGroup .seg').forEach(function (s) { s.addEventListener('click', function () { setCondition(s.dataset.cond); }); });
  on('inpExpiry', 'change', onExpiry);
  on('inpTemp', 'input', renderTemp);
  on('inpRemarks', 'input', function () { $('inpRemarks').classList.remove('invalid'); updateAll(); });

  on('btnOpenCam', 'click', openCamera);
  on('btnCamClose', 'click', stopCamera);
  on('btnCapture', 'click', capture);
  on('btnCamSwitch', 'click', switchCamera);
  on('btnPickPhoto', 'click', function () { $('fileGallery').click(); });
  on('fileCamera', 'change', onFiles);
  on('fileGallery', 'change', onFiles);

  on('btnGeo', 'click', getGeo);
  on('btnSubmit', 'click', submitCount);
  on('btnClear', 'click', function () { if (confirm('Clear this count?')) { clearForm(); updateAll(); } });

  on('btnRefreshHistory', 'click', renderProgress);
  on('inpFilter', 'input', renderProgress);
  on('btnSync', 'click', function () { syncQueue(false); });

  document.querySelectorAll('.dock-btn').forEach(function (b) { b.addEventListener('click', function () { switchTab(b.dataset.tab); }); });

  window.addEventListener('online', function () { updateNet(); toast('Back online', 'success'); syncQueue(true); });
  window.addEventListener('offline', function () { updateNet(); toast('Offline. Counts will be saved on this device.', 'warn'); });
  document.addEventListener('visibilitychange', function () { if (document.hidden) { stopCamera(); stopScan(); } });
}

/* ---------------- INIT ---------------- */
(function init() {
  applyBranding();
  bind();
  if (!CONFIG.AUTH_ENABLED) { $('btnSendOtp').firstChild.textContent = 'Continue'; }
  window.STOCKSENSE_READY = true;
  var u = loadUser();
  if (u) { S.user = { email: u.email }; enterApp(); }
})();
