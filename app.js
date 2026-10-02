/* =========================================================
   StockSense · New Item Registration
   GitHub Pages front end → Power Automate → SharePoint
   ========================================================= */
'use strict';

/* ---------------- SETTINGS (edit here) ---------------- */
var CONFIG = {
  APP_NAME: 'StockSense',

  // Flow 1: ST - Verify Staff Code (receives { code }, returns { status, staffName, outlet })
  FLOW_VERIFY_STAFF_CODE: 'https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/23/workflows/18931167830b40868c07c08724a31fa2/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=S16ve41XH4Ry1gF43VxNDY1ZP9TjQ639-dYbwT4BFBY',

  // Flow 2: ST - Register New Item (receives the item + photos, returns { success, reference })
  FLOW_SUBMIT_ITEM: 'PASTE_ST_REGISTER_NEW_ITEM_HTTP_URL',

  CATEGORIES: ['Fresh food', 'Frozen food', 'Dry goods', 'Dairy', 'Beverage', 'Alcohol',
    'Sauces & condiments', 'Bakery', 'Packaging & disposables', 'Cleaning & chemicals',
    'Smallwares & utensils', 'Equipment', 'Other'],
  STORAGE_AREAS: ['Chiller', 'Freezer', 'Dry store', 'Bar', 'Kitchen', 'Back store', 'Front counter'],
  UOMS: ['pcs', 'kg', 'g', 'L', 'ml', 'pack', 'box', 'carton', 'case', 'bottle', 'can', 'bag',
    'tray', 'tub', 'roll', 'set', 'dozen'],

  MAX_PHOTOS: 4,
  PHOTO_MAX_SIDE: 1600,
  PHOTO_QUALITY: 0.75,
  NEAR_EXPIRY_DAYS: 3,
  SESSION_HOURS: 12,
  QR_LIBRARY: 'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js'
};

/* ---------------- STATE ---------------- */
var S = {
  staff: null,
  busy: false,
  syncing: false,
  category: '',
  area: '',
  photos: [],
  scanner: null,
  scanDone: false,
  history: []
};

/* ---------------- HELPERS ---------------- */
function $(id) { return document.getElementById(id); }
function on(id, ev, fn) { var e = $(id); if (e) { e.addEventListener(ev, fn); } }
function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function lc(s) { return String(s == null ? '' : s).toLowerCase(); }
function isSet(url) { return !!url && String(url).indexOf('PASTE_') !== 0; }
function uid() { return (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : 'c' + Date.now() + Math.random().toString(16).slice(2); }
function safeName(s) { return String(s || 'item').replace(/[^A-Za-z0-9-]/g, '_').slice(0, 40); }
function fmtQty(v) { return Number(Math.round(num(v) * 1000) / 1000).toLocaleString('en-MY', { maximumFractionDigits: 3 }); }
function fmtDateTime(iso) {
  return iso ? new Date(iso).toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true }) : '';
}
function el(tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) { e.className = cls; }
  if (text != null) { e.textContent = text; }
  return e;
}
function svgIcon(name) {
  var ns = 'http://www.w3.org/2000/svg';
  var s = document.createElementNS(ns, 'svg');
  s.setAttribute('class', 'ic');
  var u = document.createElementNS(ns, 'use');
  u.setAttribute('href', '#i-' + name);
  s.appendChild(u);
  return s;
}
var store = {
  get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* storage full or blocked */ } },
  del: function (k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }
};

function toast(msg, type, ms) {
  var host = $('toasts');
  while (host.children.length >= 2) { host.removeChild(host.firstChild); }
  var t = el('div', 'toast ' + (type || ''));
  t.appendChild(el('i'));
  t.appendChild(el('span', '', msg));
  host.appendChild(t);
  setTimeout(function () { t.style.opacity = '0'; setTimeout(function () { t.remove(); }, 300); }, ms || 3500);
}
function loading(show, text) {
  $('loaderText').textContent = text || 'Please wait…';
  $('loader').classList.toggle('hidden', !show);
}
function showView(id) {
  document.querySelectorAll('.view').forEach(function (v) { v.classList.toggle('active', v.id === id); });
  window.scrollTo(0, 0);
}
async function postJson(url, body, timeoutMs) {
  var ctrl = new AbortController();
  var timer = setTimeout(function () { ctrl.abort(); }, timeoutMs || 60000);
  try {
    var res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: ctrl.signal
    });
    var text = await res.text();
    var data = {};
    try { data = text ? JSON.parse(text) : {}; } catch (e) { data = {}; }
    if (!res.ok) { throw new Error(data.message || 'The server returned an error (' + res.status + '). Please try again.'); }
    return data;
  } catch (e) {
    if (e.name === 'AbortError') { throw new Error('The request took too long. Please try again.'); }
    if (e instanceof TypeError) { throw new Error('Cannot reach the server. Check your internet connection.'); }
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

/* ---------------- THEME ---------------- */
function toggleTheme() {
  var t = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', t);
  var m = document.querySelector('meta[name="theme-color"]');
  if (m) { m.setAttribute('content', t === 'dark' ? '#0b0c22' : '#f4f2ff'); }
  store.set('ni_theme', t);
}

/* ---------------- OFFLINE QUEUE (IndexedDB) ---------------- */
var DB = {
  db: null,
  open: function () {
    var self = this;
    return new Promise(function (ok, fail) {
      if (self.db) { return ok(self.db); }
      if (!window.indexedDB) { return fail(new Error('Offline storage is not supported on this device')); }
      var r = indexedDB.open('stocksense_newitems', 1);
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
  put: function (rec) { return this.run('readwrite', function (s) { return s.put(rec); }); },
  del: function (id) { return this.run('readwrite', function (s) { return s.delete(id); }); },
  all: function () { return this.run('readonly', function (s) { return s.getAll(); }); }
};

/* ---------------- LOCAL HISTORY ---------------- */
function historyKey() { return 'ni_history_' + lc(S.staff ? S.staff.staffName + '|' + S.staff.outlet : 'none'); }
function loadHistory() { try { S.history = JSON.parse(store.get(historyKey()) || '[]'); } catch (e) { S.history = []; } }
function saveHistory() {
  // keep the newest 300 entries; thumbnails are small
  S.history = S.history.slice(0, 300);
  store.set(historyKey(), JSON.stringify(S.history));
}

/* ---------------- SIGN IN ---------------- */
function saveStaff(staff) {
  S.staff = staff;
  store.set('ni_staff', JSON.stringify({ staffName: staff.staffName, outlet: staff.outlet, exp: Date.now() + CONFIG.SESSION_HOURS * 3600000 }));
}
function loadStaff() {
  try {
    var s = JSON.parse(store.get('ni_staff'));
    if (s && s.exp > Date.now() && s.staffName) { return { staffName: s.staffName, outlet: s.outlet || '' }; }
  } catch (e) { /* ignore */ }
  store.del('ni_staff');
  return null;
}
async function signIn() {
  if (S.busy) { return; }
  var input = $('inpStaffCode');
  var code = input.value.trim();
  if (!code) { input.classList.add('invalid'); input.focus(); toast('Please enter your staff code', 'error'); return; }
  if (!isSet(CONFIG.FLOW_VERIFY_STAFF_CODE)) {
    toast('Setup needed: paste the Verify Staff Code flow URL into app.js', 'error', 6000);
    return;
  }
  S.busy = true;
  loading(true, 'Checking your code…');
  try {
    var r = await postJson(CONFIG.FLOW_VERIFY_STAFF_CODE, { code: code });
    if (lc(r.status) !== 'verified') { throw new Error(r.message || 'That code was not recognised. Please check and try again.'); }
    saveStaff({ staffName: String(r.staffName || 'Team member'), outlet: String(r.outlet || '') });
    input.value = '';
    enterApp(true);
  } catch (e) {
    input.classList.add('invalid');
    toast(e.message, 'error', 5000);
  } finally {
    S.busy = false;
    loading(false);
  }
}
function signOut() {
  if (!confirm('Sign out of StockSense?')) { return; }
  stopScan();
  store.del('ni_staff');
  S.staff = null;
  resetForm(true);
  showView('viewLogin');
  setTimeout(function () { $('inpStaffCode').focus(); }, 150);
}

/* ---------------- GREETING ---------------- */
function firstName(full) {
  var w = String(full || '').trim().split(/\s+/)[0] || 'there';
  return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
}
function greetingWord() {
  var h = new Date().getHours();
  if (h < 12) { return 'Good morning'; }
  if (h < 18) { return 'Good afternoon'; }
  return 'Good evening';
}
function heroLine(todayCount) {
  if (todayCount === 0) { return 'Ready to register today\u2019s new items?'; }
  if (todayCount === 1) { return 'Great start. 1 item registered today.'; }
  return 'Nice work. ' + todayCount + ' items registered today.';
}
function renderGreeting() {
  $('lblGreeting').textContent = greetingWord();
  $('lblFirstName').textContent = firstName(S.staff.staffName);
  $('lblOutlet').textContent = S.staff.outlet || 'Outlet not set';
  $('lblDate').textContent = new Date().toLocaleDateString('en-MY', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
}
function renderStats() {
  var today = new Date().toDateString();
  var t = S.history.filter(function (h) { return new Date(h.submittedAt).toDateString() === today; }).length;
  $('statToday').textContent = t;
  $('statTotal').textContent = S.history.length;
  $('lblHeroSub').textContent = heroLine(t);
  updatePendingBadge();
}

/* ---------------- ENTER APP ---------------- */
function enterApp(fresh) {
  showView('viewApp');
  loadHistory();
  renderGreeting();
  renderStats();
  updateNet();
  switchTab('tabRegister');
  updateForm();
  if (fresh) { toast('Welcome, ' + firstName(S.staff.staffName) + '! You are signed in for ' + (S.staff.outlet || 'your outlet') + '.', 'success', 4000); }
  if (navigator.onLine) { syncQueue(true); }
}

/* ---------------- FORM: CHIPS ---------------- */
function renderChips(boxId, list, key) {
  var box = $(boxId);
  box.textContent = '';
  list.forEach(function (label) {
    var b = el('button', 'chip', label);
    b.type = 'button';
    b.setAttribute('role', 'radio');
    b.setAttribute('aria-checked', 'false');
    b.addEventListener('click', function () {
      S[key] = label;
      box.classList.remove('invalid');
      box.querySelectorAll('.chip').forEach(function (c) {
        var a = c === b;
        c.classList.toggle('active', a);
        c.setAttribute('aria-checked', a ? 'true' : 'false');
      });
      updateForm();
    });
    box.appendChild(b);
  });
}
function clearChips(boxId) {
  $(boxId).querySelectorAll('.chip').forEach(function (c) { c.classList.remove('active'); c.setAttribute('aria-checked', 'false'); });
  $(boxId).classList.remove('invalid');
}

/* ---------------- FORM: QUANTITY ---------------- */
function stepSize() {
  var u = $('selUom').value;
  if (u === 'kg' || u === 'L') { return 0.5; }
  if (u === 'g' || u === 'ml') { return 100; }
  return 1;
}
function bumpQty(dir) {
  var v = Math.max(0, Math.round((num($('inpQty').value) + dir * stepSize()) * 1000) / 1000);
  $('inpQty').value = v;
  $('inpQty').classList.remove('invalid');
  updateForm();
}

/* ---------------- FORM: EXPIRY ---------------- */
function onExpiry() {
  var v = $('inpExpiry').value, b = $('expiryHint');
  if (!v) { b.classList.add('hidden'); return; }
  var t = new Date(); t.setHours(0, 0, 0, 0);
  var days = Math.round((new Date(v + 'T00:00:00') - t) / 86400000);
  b.classList.remove('hidden');
  if (days < 0) { b.className = 'pill bad'; b.textContent = 'Already expired'; }
  else if (days <= CONFIG.NEAR_EXPIRY_DAYS) { b.className = 'pill warn'; b.textContent = days === 0 ? 'Expires today' : 'Expires in ' + days + ' day' + (days === 1 ? '' : 's'); }
  else { b.className = 'pill ok'; b.textContent = days + ' days of shelf life'; }
}

/* ---------------- FORM: DUPLICATE CHECK ---------------- */
function checkDuplicate() {
  var bc = $('inpBarcode').value.trim();
  var nm = lc($('inpName').value.trim());
  var hit = null;
  for (var i = 0; i < S.history.length; i++) {
    var h = S.history[i];
    if ((bc && h.barcode && h.barcode === bc) || (nm && lc(h.itemName) === nm)) { hit = h; break; }
  }
  var w = $('dupWarn');
  if (hit) {
    w.textContent = 'You already registered \u201c' + hit.itemName + '\u201d on ' + fmtDateTime(hit.submittedAt) + '. Only submit again if this is a different item.';
    w.classList.remove('hidden');
  } else {
    w.classList.add('hidden');
  }
}

/* ---------------- FORM: READINESS ---------------- */
function readiness() {
  return {
    name: !!$('inpName').value.trim(),
    cat: !!S.category,
    area: !!S.area,
    qty: $('inpQty').value.trim() !== '' && num($('inpQty').value) >= 0 && !!$('selUom').value,
    photo: S.photos.length > 0
  };
}
function updateForm() {
  var r = readiness();
  var keys = Object.keys(r);
  var done = keys.filter(function (k) { return r[k]; }).length;
  $('progText').textContent = done === keys.length ? 'All set. Ready to submit' : done + ' of ' + keys.length + ' required';
  $('progBar').style.width = Math.round(done / keys.length * 100) + '%';
  document.querySelectorAll('#checklist li').forEach(function (li) { li.classList.toggle('done', !!r[li.dataset.key]); });

  var qty = $('inpQty').value.trim();
  $('sumItem').textContent = $('inpName').value.trim() || '\u2014';
  $('sumBarcode').textContent = $('inpBarcode').value.trim() || 'None';
  $('sumCat').textContent = S.category || '\u2014';
  $('sumArea').textContent = S.area || '\u2014';
  $('sumQty').textContent = qty !== '' ? fmtQty(qty) + ' ' + $('selUom').value : '\u2014';
  $('sumPhotos').textContent = S.photos.length + ' of ' + CONFIG.MAX_PHOTOS;
}
function resetForm(silent) {
  ['inpBarcode', 'inpName', 'inpBrand', 'inpPack', 'inpQty', 'inpCost', 'inpExpiry', 'inpBatch', 'inpRemarks'].forEach(function (id) {
    $(id).value = '';
    $(id).classList.remove('invalid');
  });
  $('selUom').value = CONFIG.UOMS[0];
  S.category = ''; S.area = ''; S.photos = [];
  clearChips('catGroup'); clearChips('areaGroup');
  $('expiryHint').classList.add('hidden');
  $('dupWarn').classList.add('hidden');
  stopScan();
  renderPhotos();
  updateForm();
  if (!silent) { window.scrollTo({ top: 0, behavior: 'smooth' }); }
}

/* ---------------- BARCODE SCANNER (captures barcode of the new item) ---------------- */
function loadScript(src) {
  return new Promise(function (ok, fail) {
    var s = document.createElement('script');
    s.src = src; s.onload = ok;
    s.onerror = function () { fail(new Error('Scanner could not load')); };
    document.head.appendChild(s);
  });
}
async function startScan() {
  await stopScan();
  if (!window.Html5Qrcode) {
    loading(true, 'Starting camera…');
    try { await loadScript(CONFIG.QR_LIBRARY); }
    catch (e) { loading(false); toast('Scanner could not load. Type the barcode instead.', 'error'); return; }
    loading(false);
  }
  $('scanBox').classList.remove('hidden');
  S.scanDone = false;
  S.scanner = new window.Html5Qrcode('reader');
  try {
    await S.scanner.start(
      { facingMode: 'environment' },
      { fps: 12, qrbox: function (w, h) { return { width: Math.floor(w * 0.85), height: Math.max(90, Math.floor(h * 0.4)) }; } },
      onScanned,
      function () { /* ignore frames without a code */ }
    );
    $('scanBox').scrollIntoView({ behavior: 'smooth', block: 'center' });
  } catch (e) {
    toast('Camera not available. Allow camera access or type the barcode.', 'error', 5000);
    stopScan();
  }
}
async function stopScan() {
  if (S.scanner) {
    try { if (S.scanner.isScanning) { await S.scanner.stop(); } S.scanner.clear(); } catch (e) { /* ignore */ }
    S.scanner = null;
  }
  $('scanBox').classList.add('hidden');
}
async function onScanned(text) {
  if (S.scanDone) { return; }
  S.scanDone = true;
  await stopScan();
  if (navigator.vibrate) { navigator.vibrate(60); }
  $('inpBarcode').value = String(text).trim();
  toast('Barcode captured', 'success');
  checkDuplicate();
  updateForm();
  $('inpName').focus();
}

/* ---------------- PHOTOS ---------------- */
function stampLines() {
  var ts = new Date().toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true });
  return [
    (CONFIG.APP_NAME + ' \u00b7 New item').toUpperCase(),
    ($('inpName').value.trim() || 'Item') + ($('inpBarcode').value.trim() ? '  \u00b7  ' + $('inpBarcode').value.trim() : ''),
    (S.staff ? S.staff.staffName + '  \u00b7  ' + S.staff.outlet : '') + '  \u00b7  ' + ts
  ];
}
function processImage(img, w, h) {
  var scale = Math.min(1, CONFIG.PHOTO_MAX_SIDE / Math.max(w, h));
  var cw = Math.round(w * scale), ch = Math.round(h * scale);
  var c = $('workCanvas'); c.width = cw; c.height = ch;
  var ctx = c.getContext('2d');
  ctx.drawImage(img, 0, 0, cw, ch);
  var lines = stampLines();
  var fs = Math.max(13, Math.round(cw / 50)), pad = Math.round(fs * 0.8), lh = Math.round(fs * 1.45);
  var bh = pad * 2 + lh * lines.length;
  ctx.fillStyle = 'rgba(16,14,46,0.74)';
  ctx.fillRect(0, ch - bh, cw, bh);
  var g = ctx.createLinearGradient(0, ch - bh, 0, ch);
  g.addColorStop(0, '#14b8d4'); g.addColorStop(1, '#6d4dff');
  ctx.fillStyle = g;
  ctx.fillRect(0, ch - bh, Math.max(4, Math.round(fs / 3)), bh);
  ctx.font = '600 ' + fs + 'px Arial, sans-serif';
  ctx.textBaseline = 'top';
  lines.forEach(function (t, i) {
    ctx.fillStyle = i === 0 ? '#7ee3f2' : '#ffffff';
    ctx.fillText(t, pad + Math.round(fs / 2), ch - bh + pad + i * lh);
  });
  return c.toDataURL('image/jpeg', CONFIG.PHOTO_QUALITY);
}
function thumbnail(dataUrl) {
  return new Promise(function (ok) {
    var img = new Image();
    img.onload = function () {
      var c = document.createElement('canvas'), s = 120;
      var r = Math.max(s / img.width, s / img.height);
      c.width = s; c.height = s;
      c.getContext('2d').drawImage(img, (s - img.width * r) / 2, (s - img.height * r) / 2, img.width * r, img.height * r);
      ok(c.toDataURL('image/jpeg', 0.6));
    };
    img.onerror = function () { ok(''); };
    img.src = dataUrl;
  });
}
function readImage(file) {
  return new Promise(function (ok, fail) {
    var url = URL.createObjectURL(file), img = new Image();
    img.onload = function () { URL.revokeObjectURL(url); ok(img); };
    img.onerror = function () { URL.revokeObjectURL(url); fail(new Error('That image could not be read')); };
    img.src = url;
  });
}
async function onFiles(e) {
  var files = Array.prototype.slice.call(e.target.files || []);
  e.target.value = '';
  for (var i = 0; i < files.length; i++) {
    if (S.photos.length >= CONFIG.MAX_PHOTOS) { toast('Maximum ' + CONFIG.MAX_PHOTOS + ' photos per item', 'warn'); break; }
    if (files[i].type.indexOf('image/') !== 0) { continue; }
    try {
      var img = await readImage(files[i]);
      S.photos.push({ id: uid(), data: processImage(img, img.naturalWidth, img.naturalHeight) });
    } catch (err) { toast(err.message, 'error'); }
  }
  renderPhotos();
}
function renderPhotos() {
  var g = $('photoGrid');
  g.textContent = '';
  S.photos.forEach(function (p, i) {
    var t = el('div', 'thumb');
    var img = document.createElement('img'); img.src = p.data; img.alt = 'Item photo ' + (i + 1);
    var x = el('button'); x.type = 'button'; x.setAttribute('aria-label', 'Remove photo'); x.appendChild(svgIcon('x'));
    x.addEventListener('click', function () { S.photos = S.photos.filter(function (q) { return q.id !== p.id; }); renderPhotos(); });
    t.appendChild(img); t.appendChild(el('span', '', i === 0 ? 'Main' : '#' + (i + 1))); t.appendChild(x);
    g.appendChild(t);
  });
  $('photoCounter').textContent = S.photos.length + '/' + CONFIG.MAX_PHOTOS;
  var full = S.photos.length >= CONFIG.MAX_PHOTOS;
  $('btnTakePhoto').disabled = full;
  $('btnPickPhoto').disabled = full;
  updateForm();
}

/* ---------------- SUBMIT ---------------- */
function validate() {
  var r = readiness();
  if (!r.name) { $('inpName').classList.add('invalid'); $('inpName').focus(); toast('Please enter the item name', 'error'); return false; }
  if (!r.cat) { $('catGroup').classList.add('invalid'); $('catGroup').scrollIntoView({ behavior: 'smooth', block: 'center' }); toast('Please choose a category', 'error'); return false; }
  if (!r.area) { $('areaGroup').classList.add('invalid'); $('areaGroup').scrollIntoView({ behavior: 'smooth', block: 'center' }); toast('Please choose the storage area', 'error'); return false; }
  if (!r.qty) { $('inpQty').classList.add('invalid'); $('inpQty').focus(); toast('Please enter the quantity on hand', 'error'); return false; }
  if (!r.photo) { $('btnTakePhoto').scrollIntoView({ behavior: 'smooth', block: 'center' }); toast('Please add at least one photo of the item', 'error'); return false; }
  return true;
}
function buildRecord() {
  var name = $('inpName').value.trim(), stamp = Date.now();
  return {
    clientId: uid(),
    staffName: S.staff.staffName,
    outlet: S.staff.outlet,
    barcode: $('inpBarcode').value.trim(),
    itemName: name,
    brand: $('inpBrand').value.trim(),
    packSize: $('inpPack').value.trim(),
    category: S.category,
    storageArea: S.area,
    quantity: Math.round(num($('inpQty').value) * 1000) / 1000,
    uom: $('selUom').value,
    unitCost: $('inpCost').value === '' ? null : Math.round(num($('inpCost').value) * 100) / 100,
    expiryDate: $('inpExpiry').value || '',
    batchNo: $('inpBatch').value.trim(),
    remarks: $('inpRemarks').value.trim(),
    submittedAt: new Date().toISOString(),
    photos: S.photos.map(function (p, i) {
      return { fileName: safeName(name) + '_' + stamp + '_' + (i + 1) + '.jpg', content: p.data.split(',')[1] };
    })
  };
}
async function addToHistory(rec, reference, status) {
  var thumb = S.photos.length ? await thumbnail(S.photos[0].data) : '';
  S.history.unshift({
    clientId: rec.clientId, reference: reference || '', status: status,
    itemName: rec.itemName, barcode: rec.barcode, category: rec.category, storageArea: rec.storageArea,
    quantity: rec.quantity, uom: rec.uom, submittedAt: rec.submittedAt, thumb: thumb
  });
  saveHistory();
}
function showSuccess(reference, offline) {
  $('sheetIcon').classList.toggle('off', !!offline);
  $('sheetTitle').textContent = offline ? 'Saved on this device' : 'Item registered';
  $('sheetText').textContent = offline
    ? 'No connection right now. It will be sent automatically when you are back online.'
    : 'Thank you, ' + firstName(S.staff.staffName) + '. Your submission has been received.';
  $('sheetRef').textContent = reference || 'Pending sync';
  $('sheet').classList.remove('hidden');
  if (navigator.vibrate) { navigator.vibrate([30, 40, 30]); }
}
async function submitItem(e) {
  if (e) { e.preventDefault(); }
  if (S.busy || !validate()) { return; }
  S.busy = true;
  var rec = buildRecord();
  try {
    if (!isSet(CONFIG.FLOW_SUBMIT_ITEM)) {
      // Test mode: flow not connected yet. Nothing leaves the device.
      var ref = 'TEST-' + String(S.history.length + 1).padStart(4, '0');
      await addToHistory(rec, ref, 'Test');
      showSuccess(ref, false);
      toast('Test mode: connect FLOW_SUBMIT_ITEM to save to SharePoint', 'warn', 5000);
    } else if (!navigator.onLine) {
      await DB.put(rec);
      await addToHistory(rec, '', 'Pending');
      showSuccess('', true);
    } else {
      loading(true, 'Submitting item and photos…');
      try {
        var r = await postJson(CONFIG.FLOW_SUBMIT_ITEM, rec, 120000);
        if (r.success === false) { throw new Error(r.message || 'Submission was not accepted'); }
        await addToHistory(rec, r.reference || r.stockTakeId || '', 'Submitted');
        showSuccess(r.reference || r.stockTakeId || 'Received', false);
      } catch (err) {
        await DB.put(rec);
        await addToHistory(rec, '', 'Pending');
        showSuccess('', true);
        toast(err.message, 'warn', 5000);
      } finally {
        loading(false);
      }
    }
    resetForm(true);
    renderStats();
  } catch (err) {
    toast('Could not save: ' + err.message, 'error', 6000);
  } finally {
    S.busy = false;
  }
}

/* ---------------- SYNC ---------------- */
async function updatePendingBadge() {
  var n = 0;
  try { n = (await DB.all()).length; } catch (e) { /* ignore */ }
  $('statPending').textContent = n;
  $('badgePending').textContent = n;
  $('badgePending').classList.toggle('hidden', !n);
}
async function syncQueue(silent) {
  if (S.syncing) { return; }
  if (!isSet(CONFIG.FLOW_SUBMIT_ITEM)) { if (!silent) { toast('The submit flow is not connected yet', 'warn'); } return; }
  if (!navigator.onLine) { if (!silent) { toast('You are offline. Try again when connected.', 'warn'); } return; }
  S.syncing = true;
  var all = [];
  try { all = await DB.all(); } catch (e) { S.syncing = false; return; }
  if (!all.length) { S.syncing = false; if (!silent) { toast('Everything is up to date', 'success'); } return; }
  if (!silent) { loading(true, 'Sending ' + all.length + ' saved item(s)…'); }
  var ok = 0, failed = 0;
  for (var i = 0; i < all.length; i++) {
    try {
      var r = await postJson(CONFIG.FLOW_SUBMIT_ITEM, all[i], 120000);
      if (r.success === false) { throw new Error(r.message); }
      await DB.del(all[i].clientId);
      var id = all[i].clientId;
      S.history.forEach(function (h) { if (h.clientId === id) { h.status = 'Submitted'; h.reference = r.reference || r.stockTakeId || ''; } });
      ok++;
    } catch (e) { failed++; }
  }
  saveHistory();
  S.syncing = false;
  if (!silent) { loading(false); }
  renderStats();
  renderHistory();
  if (ok) { toast(ok + ' saved item(s) sent successfully', 'success'); }
  if (failed && !silent) { toast(failed + ' item(s) could not be sent yet', 'warn'); }
}

/* ---------------- HISTORY VIEW ---------------- */
function renderHistory() {
  var q = lc($('inpFilter').value.trim());
  var list = q ? S.history.filter(function (h) { return [h.itemName, h.barcode, h.category, h.storageArea, h.reference].some(function (v) { return lc(v).indexOf(q) !== -1; }); }) : S.history;
  var box = $('historyList');
  box.textContent = '';
  if (!list.length) {
    var e = el('div', 'empty');
    e.appendChild(el('b', '', S.history.length ? 'No matching items' : 'No submissions yet'));
    e.appendChild(document.createTextNode(S.history.length ? 'Try a different search.' : 'Items you register will appear here.'));
    box.appendChild(e);
    return;
  }
  list.forEach(function (h) {
    var row = el('article', 'h-item');
    if (h.thumb) { var img = document.createElement('img'); img.src = h.thumb; img.alt = ''; row.appendChild(img); }
    else { var ph = el('div', 'h-noimg'); ph.appendChild(svgIcon('box')); row.appendChild(ph); }
    var m = el('div', 'h-main');
    var top = el('div', 'h-top');
    top.appendChild(el('b', '', h.itemName));
    top.appendChild(el('span', 'status ' + h.status, h.status === 'Pending' ? 'Waiting to sync' : h.status));
    m.appendChild(top);
    m.appendChild(el('p', 'h-meta', fmtQty(h.quantity) + ' ' + h.uom + '  \u00b7  ' + h.storageArea + '  \u00b7  ' + h.category));
    m.appendChild(el('p', 'h-meta', (h.reference ? h.reference + '  \u00b7  ' : '') + fmtDateTime(h.submittedAt)));
    row.appendChild(m);
    box.appendChild(row);
  });
}

/* ---------------- TABS + NETWORK ---------------- */
function switchTab(id) {
  document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.id === id); });
  document.querySelectorAll('.tab-btn').forEach(function (b) { b.classList.toggle('active', b.dataset.tab === id); });
  if (id === 'tabHistory') { stopScan(); renderHistory(); }
}
function updateNet() {
  var b = $('netBadge'), online = navigator.onLine;
  b.className = 'net ' + (online ? 'on' : 'off');
  b.querySelector('b').textContent = online ? 'Online' : 'Offline';
}

/* ---------------- EVENTS ---------------- */
function bind() {
  document.querySelectorAll('.js-theme').forEach(function (b) { b.addEventListener('click', toggleTheme); });

  on('btnLogin', 'click', signIn);
  on('inpStaffCode', 'keydown', function (e) { if (e.key === 'Enter') { signIn(); } });
  on('inpStaffCode', 'input', function () { $('inpStaffCode').classList.remove('invalid'); });
  on('btnToggleCode', 'click', function () {
    var i = $('inpStaffCode');
    i.type = i.type === 'password' ? 'text' : 'password';
  });
  on('btnLogout', 'click', signOut);

  on('btnScan', 'click', startScan);
  on('btnStopScan', 'click', stopScan);
  on('inpBarcode', 'change', function () { checkDuplicate(); updateForm(); });
  on('inpBarcode', 'input', updateForm);
  on('inpName', 'input', function () { $('inpName').classList.remove('invalid'); updateForm(); });
  on('inpName', 'blur', checkDuplicate);

  on('btnMinus', 'click', function () { bumpQty(-1); });
  on('btnPlus', 'click', function () { bumpQty(1); });
  on('inpQty', 'input', function () { $('inpQty').classList.remove('invalid'); updateForm(); });
  on('selUom', 'change', updateForm);
  on('inpExpiry', 'change', onExpiry);

  on('btnTakePhoto', 'click', function () { $('fileCamera').click(); });
  on('btnPickPhoto', 'click', function () { $('fileGallery').click(); });
  on('fileCamera', 'change', onFiles);
  on('fileGallery', 'change', onFiles);

  on('itemForm', 'submit', submitItem);
  on('btnReset', 'click', function () { if (confirm('Clear everything on this form?')) { resetForm(); } });

  on('btnAnother', 'click', function () {
    $('sheet').classList.add('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    setTimeout(function () { $('inpBarcode').focus(); }, 300);
  });
  on('btnViewHistory', 'click', function () { $('sheet').classList.add('hidden'); switchTab('tabHistory'); window.scrollTo({ top: 0, behavior: 'smooth' }); });

  on('inpFilter', 'input', renderHistory);
  on('btnSync', 'click', function () { syncQueue(false); });
  document.querySelectorAll('.tab-btn').forEach(function (b) { b.addEventListener('click', function () { switchTab(b.dataset.tab); }); });

  window.addEventListener('online', function () { updateNet(); syncQueue(true); });
  window.addEventListener('offline', function () { updateNet(); toast('You are offline. New items will be saved on this device.', 'warn'); });
  document.addEventListener('visibilitychange', function () { if (document.hidden) { stopScan(); } });
}

/* ---------------- START ---------------- */
(function init() {
  renderChips('catGroup', CONFIG.CATEGORIES, 'category');
  renderChips('areaGroup', CONFIG.STORAGE_AREAS, 'area');
  CONFIG.UOMS.forEach(function (u) { $('selUom').appendChild(new Option(u, u)); });
  bind();
  renderPhotos();
  window.STOCKSENSE_READY = true;
  var staff = loadStaff();
  if (staff) { S.staff = staff; enterApp(false); }
})();
