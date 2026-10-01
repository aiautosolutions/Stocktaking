/* =========================================================
   StockSense — Operations Stock Take (single SharePoint list)
   ========================================================= */
'use strict';

/* ---------------- CONFIG ---------------- */
var CONFIG = {
  COMPANY_NAME: 'Your Company',   // logo text + photo stamp
  COMPANY_SHORT: '',              // optional badge text e.g. 'ACME'. Blank = initials
  AUTH_ENABLED: true,             // false = skip OTP (testing only)

  FLOW_SEND_OTP: 'https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/22/workflows/587bb4ec5a9141d4b2b61445d9841139/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dMgM0-vkKVd11vFfkQwMaUdsrnId4PhwMAI0ulkC9tE',
  FLOW_VERIFY_OTP: 'https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/31/workflows/296ff3060c1a449582e85c81e154c6df/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tdeev6xhU5HgYOU0zcHANo0FBK2kfDEqVrptgt1bkjQ',
  FLOW_CONFIG: 'PASTE_OPS_ST_GetConfig_HTTP_URL',
  FLOW_SUBMIT: 'PASTE_OPS_ST_SubmitCount_HTTP_URL',
  FLOW_HISTORY: 'PASTE_OPS_ST_MySubmissions_HTTP_URL',

  QR_LIBRARY: 'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js',
  DEMO_LOCATIONS: ['Head Office', 'Store Room', 'Archive Room', 'Warehouse'],
  REMARK_REQUIRED_FOR: ['Damaged', 'Missing', 'Obsolete'],
  MIN_PHOTOS: 1,
  MAX_PHOTOS: 5,
  PHOTO_MAX_DIM: 1600,
  PHOTO_QUALITY: 0.72,
  LOGIN_HOURS: 10,
  RESEND_SECONDS: 60              // fallback only; flow's expiresIn is used when provided
};

/* ---------------- STATE ---------------- */
var S = {
  user: null, pendingEmail: '', busy: false, syncing: false,
  session: '', locations: [],
  condition: 'Good', photos: [],
  stream: null, facing: 'environment', stampTimer: null,
  scanner: null, scanHandled: false,
  geo: null, history: [], resendTimer: null
};

/* ---------------- HELPERS ---------------- */
function $(id) { return document.getElementById(id); }
function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function safe(s) { return String(s == null ? '' : s).replace(/[^A-Za-z0-9-]/g, '_'); }
function uid() { return (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : 'c' + Date.now() + Math.random().toString(16).slice(2); }
function buzz(ms) { if (navigator.vibrate) { navigator.vibrate(ms); } }
function isSet(url) { return !!url && String(url).indexOf('PASTE_') !== 0; }
function fmtDT(iso) {
  return iso ? new Date(iso).toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false }) : '—';
}

var store = {
  get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
  del: function (k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }
};

function el(tag, cls, text) {
  var e = document.createElement(tag);
  if (cls) { e.className = cls; }
  if (text !== undefined && text !== null) { e.textContent = text; }
  return e;
}
function toast(msg, type, ms) {
  type = type || 'info';
  ms = ms || 3400;
  var t = el('div', 'toast ' + type);
  var icon = type === 'error' ? '!' : (type === 'warn' ? '•' : '✓');
  t.appendChild(el('span', 'tdot', icon));
  t.appendChild(el('span', '', msg));
  $('toastHost').appendChild(t);
  setTimeout(function () { t.style.opacity = '0'; setTimeout(function () { t.remove(); }, 300); }, ms);
}
function loading(on, text) {
  $('loaderText').textContent = text || 'Loading…';
  $('loader').classList.toggle('hidden', !on);
}
function showView(id) {
  document.querySelectorAll('.view').forEach(function (v) { v.classList.toggle('active', v.id === id); });
}
function loadScript(src) {
  return new Promise(function (resolve, reject) {
    var s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = function () { reject(new Error('Failed to load ' + src)); };
    document.head.appendChild(s);
  });
}
async function api(url, body, timeoutMs) {
  if (!isSet(url)) { throw new Error('Flow URL is not configured in app.js'); }
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
    try { data = text ? JSON.parse(text) : {}; } catch (e) { data = { raw: text }; }
    if (!res.ok) { throw new Error(data.message || ('Request failed (' + res.status + ')')); }
    return data;
  } catch (e) {
    if (e.name === 'AbortError') { throw new Error('Request timed out'); }
    if (e instanceof TypeError) { throw new Error('Cannot reach the flow. Check the HTTP trigger is set to Anyone and the Response has header Access-Control-Allow-Origin = *'); }
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

/* ---------------- BRANDING ---------------- */
function applyBranding() {
  var name = CONFIG.COMPANY_NAME || 'Company';
  var short = CONFIG.COMPANY_SHORT || name.split(/\s+/).filter(Boolean).slice(0, 2).map(function (w) { return w.charAt(0); }).join('').toUpperCase();
  document.querySelectorAll('.js-company').forEach(function (e) { e.textContent = name; });
  document.querySelectorAll('.js-logo').forEach(function (e) { e.textContent = short; });
  document.querySelectorAll('.js-mark').forEach(function (e) { e.textContent = name.toUpperCase(); });
  document.title = name + ' · StockSense';
}
function greeting() {
  var h = new Date().getHours();
  if (h >= 18) { return 'Good evening'; }
  if (h >= 12) { return 'Good afternoon'; }
  return 'Good morning';
}
function todayKey() { return 'ops_st_today_' + new Date().toISOString().slice(0, 10); }
function renderToday() { $('lblToday').textContent = num(store.get(todayKey())); }
function bumpToday() { store.set(todayKey(), String(num(store.get(todayKey())) + 1)); renderToday(); }

/* ---------------- OFFLINE QUEUE (IndexedDB) ---------------- */
var DB = {
  db: null,
  open: function () {
    var self = this;
    return new Promise(function (resolve, reject) {
      if (self.db) { return resolve(self.db); }
      if (!window.indexedDB) { return reject(new Error('Offline storage not supported')); }
      var r = indexedDB.open('stocksense_ops', 1);
      r.onupgradeneeded = function () { r.result.createObjectStore('queue', { keyPath: 'clientId' }); };
      r.onsuccess = function () { self.db = r.result; resolve(self.db); };
      r.onerror = function () { reject(r.error); };
    });
  },
  run: async function (mode, fn) {
    var db = await this.open();
    return new Promise(function (resolve, reject) {
      var tx = db.transaction('queue', mode);
      var req = fn(tx.objectStore('queue'));
      tx.oncomplete = function () { resolve(req ? req.result : undefined); };
      tx.onerror = function () { reject(tx.error); };
    });
  },
  put: function (rec) { return this.run('readwrite', function (s) { return s.put(rec); }); },
  del: function (id) { return this.run('readwrite', function (s) { return s.delete(id); }); },
  all: function () { return this.run('readonly', function (s) { return s.getAll(); }); }
};

/* ---------------- AUTH ---------------- */
function saveUser(u) {
  S.user = u;
  store.set('ops_st_user', JSON.stringify({ email: u.email, exp: Date.now() + CONFIG.LOGIN_HOURS * 3600000 }));
}
function loadUser() {
  try {
    var u = JSON.parse(store.get('ops_st_user'));
    if (u && u.exp > Date.now()) { return u; }
  } catch (e) { /* ignore */ }
  store.del('ops_st_user');
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
    if (left > 0) {
      b.textContent = 'Code expires in ' + left + 's';
    } else {
      clearInterval(S.resendTimer);
      b.disabled = false;
      b.textContent = 'Code expired · Resend';
    }
  }, 1000);
}
async function sendOtp(isResend) {
  if (S.busy) { return; }
  var email = isResend ? S.pendingEmail : $('inpEmail').value.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    $('inpEmail').classList.add('invalid');
    toast('Enter a valid email address', 'error');
    return;
  }
  $('inpEmail').classList.remove('invalid');

  if (!CONFIG.AUTH_ENABLED) { saveUser({ email: email }); enterApp(); return; }

  S.busy = true;
  loading(true, 'Sending verification code…');
  try {
    var r = await api(CONFIG.FLOW_SEND_OTP, { email: email });
    var sst = String(r.status || '').toLowerCase();
    if (r.success === false || sst === 'error' || sst === 'failed' || sst === 'blocked') { throw new Error(r.message || 'Unable to send code'); }
    S.pendingEmail = email;
    $('lblOtpEmail').textContent = email;
    $('stepEmail').classList.add('hidden');
    $('stepOtp').classList.remove('hidden');
    $('inpOtp').value = '';
    $('inpOtp').focus();
    startResendTimer(r.expiresIn);
    toast(isResend ? 'A new code has been sent' : 'Verification code sent', 'success');
  } catch (e) {
    toast(e.message, 'error', 6000);
  } finally {
    S.busy = false;
    loading(false);
  }
}
async function verifyOtp() {
  if (S.busy) { return; }
  var otp = $('inpOtp').value.trim();
  if (!/^\d{6}$/.test(otp)) { toast('Enter the 6-digit code', 'error'); return; }
  S.busy = true;
  loading(true, 'Verifying…');
  try {
    // Flow PC_VerifyOTP reads triggerBody()?['code']; otp sent too for safety
    var r = await api(CONFIG.FLOW_VERIFY_OTP, { email: S.pendingEmail, code: otp, otp: otp });
    var st = String(r.status || '').toLowerCase();
    var ok = st === 'verified' || st === 'valid' || st === 'success' ||
             r.success === true || r.verified === true || r.valid === true;
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
    saveUser({ email: S.pendingEmail });
    S.busy = false;
    loading(false);
    toast(r.message || 'Identity verified', 'success');
    await enterApp();
  } catch (e) {
    S.busy = false;
    loading(false);
    $('inpOtp').classList.add('invalid');
    toast(e.message, 'error', 5000);
  }
}
function logout() {
  stopCamera();
  stopScan();
  store.del('ops_st_user');
  S.user = null;
  S.history = [];
  clearForm();
  $('inpEmail').value = '';
  $('stepOtp').classList.add('hidden');
  $('stepEmail').classList.remove('hidden');
  showView('viewLogin');
}

/* ---------------- BOOT ---------------- */
async function enterApp() {
  showView('viewApp');
  $('lblGreeting').textContent = greeting();
  $('lblUser').textContent = S.user.email;
  $('avatar').textContent = S.user.email.slice(0, 2).toUpperCase();
  renderToday();
  updateNet();
  await loadConfig();
  getGeo();
  updateQueueBadge();
  updateProgress();
  if (navigator.onLine) { syncQueue(true); }
}
async function loadConfig() {
  var cached = null;
  try { cached = JSON.parse(store.get('ops_st_config') || 'null'); } catch (e) { /* ignore */ }

  if (!isSet(CONFIG.FLOW_CONFIG)) {
    S.session = 'ST-DEMO';
    S.locations = CONFIG.DEMO_LOCATIONS.slice();
    toast('Demo mode: Get Config flow URL not set yet', 'warn', 4500);
  } else if (navigator.onLine) {
    loading(true, 'Preparing your workspace…');
    try {
      var r = await api(CONFIG.FLOW_CONFIG, { email: S.user.email });
      S.session = r.session || '';
      S.locations = Array.isArray(r.locations) ? r.locations : [];
      store.set('ops_st_config', JSON.stringify({ session: S.session, locations: S.locations }));
    } catch (e) {
      if (cached) {
        S.session = cached.session;
        S.locations = cached.locations;
        toast('Using saved configuration', 'warn');
      } else {
        toast('Unable to load configuration: ' + e.message, 'error', 6000);
      }
    } finally {
      loading(false);
    }
  } else if (cached) {
    S.session = cached.session;
    S.locations = cached.locations;
    toast('Offline — using saved configuration', 'warn');
  }

  $('lblSession').textContent = S.session || 'Not configured';
  var sel = $('selLocation');
  sel.innerHTML = '';
  sel.appendChild(new Option('Select location…', ''));
  S.locations.forEach(function (l) { sel.appendChild(new Option(l, l)); });
  var pref = store.get('ops_st_location');
  if (pref && S.locations.indexOf(pref) !== -1) { sel.value = pref; }
}

/* ---------------- NETWORK ---------------- */
function updateNet() {
  var on = navigator.onLine;
  $('netBadge').className = 'net ' + (on ? 'online' : 'offline');
  $('netBadge').querySelector('b').textContent = on ? 'Online' : 'Offline';
}

/* ---------------- FORM ---------------- */
function qtyValid() {
  var q = String($('inpQty').value).trim();
  return q !== '' && !isNaN(q) && num(q) >= 0 && Number.isInteger(num(q));
}
function needsRemark() { return CONFIG.REMARK_REQUIRED_FOR.indexOf(S.condition) !== -1; }
function updateProgress() {
  var loc = $('selLocation').value;
  var code = $('inpAssetCode').value.trim();
  var name = $('inpAssetName').value.trim();
  var qOk = qtyValid();
  var remOk = !needsRemark() || !!$('inpRemarks').value.trim();
  var steps = [!!(loc && code && name), qOk && remOk, S.photos.length >= CONFIG.MIN_PHOTOS];
  var done = steps.filter(Boolean).length;

  steps.forEach(function (ok, i) { $('ps' + (i + 1)).classList.toggle('done', ok); });
  $('progText').textContent = done + ' / 3';
  $('progBar').style.width = Math.round(done / 3 * 100) + '%';
  $('btnSubmit').classList.toggle('ready', done === 3);

  $('sumLoc').textContent = loc || '—';
  $('sumAsset').textContent = code ? code + (name ? ' · ' + name : '') : '—';
  $('sumQty').textContent = (qOk ? num($('inpQty').value) : '—') + ' · ' + S.condition;
  $('sumPhotos').textContent = S.photos.length + ' / ' + CONFIG.MAX_PHOTOS;
}
function setCondition(c) {
  S.condition = c;
  document.querySelectorAll('.seg').forEach(function (s) { s.classList.toggle('active', s.dataset.cond === c); });
  $('remarkReq').classList.toggle('hidden', !needsRemark());
  if (c === 'Missing') { $('inpQty').value = 0; }
  updateProgress();
}
function clearForm() {
  stopCamera();
  $('inpAssetCode').value = '';
  $('inpAssetName').value = '';
  $('inpQty').value = 1;
  $('inpRemarks').value = '';
  ['inpAssetCode', 'inpAssetName', 'selLocation', 'inpRemarks'].forEach(function (id) { $(id).classList.remove('invalid'); });
  S.photos = [];
  setCondition('Good');
  renderPhotos();
}

/* ---------------- BARCODE SCANNER ---------------- */
async function startScan() {
  stopCamera();
  await stopScan();
  if (!window.Html5Qrcode) {
    loading(true, 'Starting scanner…');
    try {
      await loadScript(CONFIG.QR_LIBRARY);
    } catch (e) {
      loading(false);
      toast('Scanner could not load. Check your connection.', 'error');
      return;
    }
    loading(false);
  }
  $('scanWrap').classList.remove('hidden');
  S.scanHandled = false;
  S.scanner = new window.Html5Qrcode('reader');
  try {
    await S.scanner.start(
      { facingMode: 'environment' },
      { fps: 12, qrbox: function (w, h) { return { width: Math.max(140, Math.floor(w * 0.8)), height: Math.max(90, Math.floor(h * 0.4)) }; } },
      onScan,
      function () {}
    );
  } catch (e) {
    toast('Camera not available for scanning', 'error');
    stopScan();
  }
}
async function stopScan() {
  if (S.scanner) {
    try {
      if (S.scanner.isScanning) { await S.scanner.stop(); }
      S.scanner.clear();
    } catch (e) { /* ignore */ }
    S.scanner = null;
  }
  $('scanWrap').classList.add('hidden');
}
async function onScan(text) {
  if (S.scanHandled) { return; }
  S.scanHandled = true;
  await stopScan();
  buzz(70);
  $('inpAssetCode').value = String(text).trim().toUpperCase();
  $('inpAssetCode').classList.remove('invalid');
  toast('Scanned ' + $('inpAssetCode').value, 'success');
  $('inpAssetName').focus();
  updateProgress();
}

/* ---------------- CAMERA & PHOTOS ---------------- */
function stampLines() {
  var ts = new Date().toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
  var gps = S.geo ? ' · ' + S.geo.lat.toFixed(5) + ', ' + S.geo.lng.toFixed(5) : '';
  return [
    (CONFIG.COMPANY_NAME || '').toUpperCase() + ' · ' + (S.session || '-'),
    ($('selLocation').value || '-') + ' · ' + ($('inpAssetCode').value.trim() || '-'),
    (S.user && S.user.email) || '-',
    ts + gps
  ];
}
function updateStamp() {
  var box = $('camStamp');
  box.textContent = '';
  stampLines().forEach(function (l, i) {
    if (i) { box.appendChild(document.createElement('br')); }
    box.appendChild(document.createTextNode(l));
  });
}
async function openCamera() {
  if (S.photos.length >= CONFIG.MAX_PHOTOS) { toast('Maximum ' + CONFIG.MAX_PHOTOS + ' photos', 'warn'); return; }
  await stopScan();
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) { $('fileCamera').click(); return; }
  try {
    stopCamera();
    S.stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: S.facing }, width: { ideal: 1920 }, height: { ideal: 1080 } },
      audio: false
    });
    var v = $('camVideo');
    v.srcObject = S.stream;
    $('camWrap').classList.remove('hidden');
    await v.play();
    updateStamp();
    S.stampTimer = setInterval(updateStamp, 1000);
    $('camWrap').scrollIntoView({ behavior: 'smooth', block: 'center' });
  } catch (e) {
    stopCamera();
    toast('Camera permission blocked — opening device camera', 'warn');
    $('fileCamera').click();
  }
}
function stopCamera() {
  if (S.stream) { S.stream.getTracks().forEach(function (t) { t.stop(); }); S.stream = null; }
  if (S.stampTimer) { clearInterval(S.stampTimer); S.stampTimer = null; }
  $('camVideo').srcObject = null;
  $('camWrap').classList.add('hidden');
}
async function switchCamera() {
  S.facing = S.facing === 'environment' ? 'user' : 'environment';
  if (S.stream) { stopCamera(); await openCamera(); }
}
function processImage(src, w, h) {
  var scale = Math.min(1, CONFIG.PHOTO_MAX_DIM / Math.max(w, h));
  var cw = Math.round(w * scale);
  var ch = Math.round(h * scale);
  var c = $('workCanvas');
  c.width = cw;
  c.height = ch;
  var ctx = c.getContext('2d');
  ctx.drawImage(src, 0, 0, cw, ch);

  var lines = stampLines();
  var fs = Math.max(14, Math.round(cw / 48));
  var pad = Math.round(fs * 0.8);
  var lh = Math.round(fs * 1.4);
  var boxH = pad * 2 + lh * lines.length;

  ctx.fillStyle = 'rgba(12,10,40,0.72)';
  ctx.fillRect(0, ch - boxH, cw, boxH);
  var g = ctx.createLinearGradient(0, ch - boxH, 0, ch);
  g.addColorStop(0, '#7c5cff');
  g.addColorStop(1, '#22d3ee');
  ctx.fillStyle = g;
  ctx.fillRect(0, ch - boxH, Math.max(4, Math.round(fs / 3)), boxH);

  ctx.font = '600 ' + fs + 'px "Space Grotesk", Arial, sans-serif';
  ctx.textBaseline = 'top';
  lines.forEach(function (t, i) {
    ctx.fillStyle = i === 0 ? '#67e8f9' : '#FFFFFF';
    ctx.fillText(t, pad + Math.round(fs / 2), ch - boxH + pad + i * lh);
  });
  return c.toDataURL('image/jpeg', CONFIG.PHOTO_QUALITY);
}
function capture() {
  var v = $('camVideo');
  if (!v.videoWidth) { toast('Camera is starting…', 'warn'); return; }
  addPhoto(processImage(v, v.videoWidth, v.videoHeight));
  var f = $('camFlash');
  f.classList.add('on');
  setTimeout(function () { f.classList.remove('on'); }, 60);
  buzz(35);
  if (S.photos.length >= CONFIG.MAX_PHOTOS) { stopCamera(); toast('Photo limit reached'); }
}
function loadImg(file) {
  return new Promise(function (resolve, reject) {
    var url = URL.createObjectURL(file);
    var img = new Image();
    img.onload = function () { URL.revokeObjectURL(url); resolve(img); };
    img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('Unreadable image')); };
    img.src = url;
  });
}
async function onFilePick(e) {
  var files = Array.prototype.slice.call(e.target.files || []);
  e.target.value = '';
  for (var i = 0; i !== files.length; i++) {
    var f = files[i];
    if (S.photos.length >= CONFIG.MAX_PHOTOS) { toast('Maximum ' + CONFIG.MAX_PHOTOS + ' photos', 'warn'); break; }
    if (f.type.indexOf('image/') !== 0) { continue; }
    try {
      var img = await loadImg(f);
      addPhoto(processImage(img, img.naturalWidth, img.naturalHeight));
    } catch (err) {
      toast('Could not read ' + f.name, 'error');
    }
  }
}
function addPhoto(dataUrl) {
  S.photos.push({ id: uid(), dataUrl: dataUrl });
  renderPhotos();
}
function renderPhotos() {
  var grid = $('photoGrid');
  grid.textContent = '';
  S.photos.forEach(function (p, i) {
    var wrap = el('div', 'thumb');
    var img = document.createElement('img');
    img.src = p.dataUrl;
    img.alt = 'Photo ' + (i + 1);
    var no = el('span', 'no', '#' + (i + 1));
    var rm = el('button', 'rm', '✕');
    rm.type = 'button';
    rm.setAttribute('aria-label', 'Remove photo');
    rm.onclick = function () { S.photos = S.photos.filter(function (x) { return x.id !== p.id; }); renderPhotos(); };
    wrap.appendChild(img);
    wrap.appendChild(no);
    wrap.appendChild(rm);
    grid.appendChild(wrap);
  });
  $('photoCounter').textContent = S.photos.length + ' / ' + CONFIG.MAX_PHOTOS;
  var full = S.photos.length >= CONFIG.MAX_PHOTOS;
  $('btnOpenCam').disabled = full;
  $('btnPickPhoto').disabled = full;
  updateProgress();
}

/* ---------------- GPS ---------------- */
function getGeo() {
  return new Promise(function (resolve) {
    var b = $('geoBadge');
    if (!navigator.geolocation) { b.textContent = '📍 Location not supported'; resolve(null); return; }
    b.textContent = '📍 Locating…';
    b.className = 'geo';
    navigator.geolocation.getCurrentPosition(function (p) {
      S.geo = { lat: p.coords.latitude, lng: p.coords.longitude, acc: Math.round(p.coords.accuracy) };
      b.textContent = '📍 Location captured (±' + S.geo.acc + ' m)';
      b.className = 'geo ok';
      resolve(S.geo);
    }, function () {
      b.textContent = '📍 Location unavailable (optional)';
      resolve(null);
    }, { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 });
  });
}

/* ---------------- SUBMIT ---------------- */
function validate() {
  var checks = [
    ['selLocation', !$('selLocation').value, 'Select a location'],
    ['inpAssetCode', !$('inpAssetCode').value.trim(), 'Enter the asset code'],
    ['inpAssetName', !$('inpAssetName').value.trim(), 'Enter the asset name']
  ];
  for (var i = 0; i !== checks.length; i++) {
    var c = checks[i];
    $(c[0]).classList.toggle('invalid', c[1]);
    if (c[1]) { $(c[0]).focus(); toast(c[2], 'error'); return false; }
  }
  if (!qtyValid()) { toast('Enter a valid whole-number quantity', 'error'); return false; }
  if (needsRemark() && !$('inpRemarks').value.trim()) {
    $('inpRemarks').classList.add('invalid');
    $('inpRemarks').focus();
    toast('Remarks are required for "' + S.condition + '"', 'error');
    return false;
  }
  $('inpRemarks').classList.remove('invalid');
  if (S.photos.length < CONFIG.MIN_PHOTOS) { toast('At least ' + CONFIG.MIN_PHOTOS + ' photo is required', 'error'); return false; }
  if (!S.session) { toast('No active session. Contact the administrator.', 'error'); return false; }
  if (!isSet(CONFIG.FLOW_SUBMIT)) { toast('Submit flow URL is not set in app.js yet', 'error', 5000); return false; }
  return true;
}
function showSuccess(id, offline) {
  $('sucIcon').textContent = offline ? '☁' : '✓';
  $('sucIcon').classList.toggle('offline', !!offline);
  $('sucWord').textContent = offline ? 'saved offline' : 'submitted';
  $('sucSub').textContent = offline
    ? 'No connection right now. This count is safely queued and will sync automatically.'
    : 'Your count and photo evidence are saved in SharePoint.';
  $('sucId').textContent = id || (offline ? 'Queued' : 'Saved');
  $('successSheet').classList.add('open');
}
function closeSuccess() {
  $('successSheet').classList.remove('open');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
async function submitCount() {
  if (S.busy || !validate()) { return; }
  S.busy = true;
  stopCamera();

  var code = $('inpAssetCode').value.trim().toUpperCase();
  var stamp = Date.now();
  var rec = {
    clientId: uid(),
    location: $('selLocation').value,
    assetCode: code,
    assetName: $('inpAssetName').value.trim(),
    physicalQty: parseInt($('inpQty').value, 10),
    condition: S.condition,
    remarks: $('inpRemarks').value.trim(),
    submittedEmail: S.user.email,
    submittedAt: new Date().toISOString(),
    photos: S.photos.map(function (p, i) {
      return { fileName: safe(code) + '_' + stamp + '_' + (i + 1) + '.jpg', content: p.dataUrl.split(',')[1] };
    })
  };

  if (!navigator.onLine) {
    try { await DB.put(rec); } catch (e) { S.busy = false; toast('Could not save offline: ' + e.message, 'error'); return; }
    bumpToday();
    S.busy = false;
    clearForm();
    updateQueueBadge();
    showSuccess(null, true);
    return;
  }

  loading(true, 'Submitting count and photo evidence…');
  try {
    var r = await api(CONFIG.FLOW_SUBMIT, rec, 120000);
    if (r.success === false) { throw new Error(r.message || 'Submission failed'); }
    bumpToday();
    buzz([30, 50, 30]);
    clearForm();
    showSuccess(r.stockTakeId, false);
  } catch (e) {
    try {
      await DB.put(rec);
      bumpToday();
      clearForm();
      showSuccess(null, true);
    } catch (err) {
      toast('Submission failed: ' + e.message, 'error', 6000);
    }
  } finally {
    S.busy = false;
    loading(false);
    updateQueueBadge();
  }
}

/* ---------------- SYNC ---------------- */
async function syncQueue(silent) {
  if (S.syncing || !isSet(CONFIG.FLOW_SUBMIT)) { return; }
  if (!navigator.onLine) { if (!silent) { toast('Still offline', 'warn'); } return; }
  var all = [];
  try { all = await DB.all(); } catch (e) { return; }
  if (!all.length) { if (!silent) { toast('Nothing to sync'); } return; }

  S.syncing = true;
  if (!silent) { loading(true, 'Syncing ' + all.length + ' record(s)…'); }
  var ok = 0;
  var fail = 0;
  for (var i = 0; i !== all.length; i++) {
    try {
      var r = await api(CONFIG.FLOW_SUBMIT, all[i], 120000);
      if (r.success === false) { throw new Error(r.message); }
      await DB.del(all[i].clientId);
      ok++;
    } catch (e) {
      fail++;
    }
  }
  S.syncing = false;
  if (!silent) { loading(false); }
  updateQueueBadge();
  if ($('tabQueue').classList.contains('active')) { renderQueue(); }
  if (ok) { toast(ok + ' queued record(s) synced', 'success'); }
  if (fail) { toast(fail + ' record(s) still pending', 'warn'); }
}
function emptyState(icon, title, copy) {
  var e = el('div', 'empty');
  e.appendChild(el('div', 'eic', icon));
  e.appendChild(el('h3', 'etitle', title));
  e.appendChild(el('p', 'ecopy', copy));
  return e;
}
function itemRow(icon, iconCls, code, name, meta, right) {
  var row = el('div', 'aitem');
  var main = el('div', 'amain');
  var t = el('p', 'at');
  t.appendChild(el('span', 'code', code));
  t.appendChild(document.createTextNode(name || ''));
  main.appendChild(t);
  main.appendChild(el('p', 'am', meta));
  row.appendChild(el('div', 'aic ' + iconCls, icon));
  row.appendChild(main);
  row.appendChild(right);
  return row;
}
async function renderQueue() {
  var all = [];
  try { all = await DB.all(); } catch (e) { /* ignore */ }
  var list = $('queueList');
  list.textContent = '';
  if (!all.length) { list.appendChild(emptyState('☁', 'All caught up', 'Nothing is waiting to sync.')); return; }
  all.forEach(function (r) {
    var del = el('button', 'del', '✕');
    del.type = 'button';
    del.setAttribute('aria-label', 'Discard');
    del.onclick = async function () {
      if (!confirm('Discard this queued record? This cannot be undone.')) { return; }
      await DB.del(r.clientId);
      renderQueue();
      updateQueueBadge();
    };
    var meta = r.location + ' · Qty ' + r.physicalQty + ' · ' + r.photos.length + ' photo(s) · ' + fmtDT(r.submittedAt);
    list.appendChild(itemRow('⇪', 'queued', r.assetCode, r.assetName, meta, del));
  });
}
async function updateQueueBadge() {
  try {
    var n = (await DB.all()).length;
    $('queueBadge').textContent = n;
    $('queueBadge').classList.toggle('hidden', !n);
  } catch (e) { /* ignore */ }
}

/* ---------------- HISTORY ---------------- */
async function loadHistory() {
  if (!isSet(CONFIG.FLOW_HISTORY)) { toast('My Submissions flow URL is not set yet', 'warn'); renderHistory(); return; }
  if (!navigator.onLine) { toast('Offline — history unavailable', 'warn'); renderHistory(); return; }
  loading(true, 'Loading your submissions…');
  try {
    var r = await api(CONFIG.FLOW_HISTORY, { email: S.user.email });
    S.history = (r.items || []).sort(function (a, b) { return new Date(b.submittedDate) - new Date(a.submittedDate); });
  } catch (e) {
    toast('Unable to load history: ' + e.message, 'error');
  } finally {
    loading(false);
    renderHistory();
  }
}
function renderHistory() {
  var h = S.history;
  var good = h.filter(function (x) { return (x.condition || 'Good') === 'Good'; }).length;
  var pct = h.length ? Math.round(good / h.length * 100) : 0;
  $('kpiTotal').textContent = h.length;
  $('kpiQty').textContent = h.reduce(function (s, x) { return s + num(x.physicalQty); }, 0).toLocaleString('en-MY');
  $('kpiExc').textContent = h.length - good;
  $('kpiVer').textContent = h.filter(function (x) { return x.status === 'Verified'; }).length + ' verified';
  $('ringPct').textContent = pct + '%';
  $('ringCap').textContent = h.length ? good + ' of ' + h.length + ' assets in good condition' : 'No records yet';
  var circ = 2 * Math.PI * 31;
  $('ringVal').setAttribute('stroke-dasharray', circ.toFixed(1));
  $('ringVal').setAttribute('stroke-dashoffset', (circ * (1 - pct / 100)).toFixed(1));

  var q = $('inpFilter').value.trim().toLowerCase();
  var rows = q ? h.filter(function (x) {
    return [x.assetCode, x.assetName, x.location].some(function (v) { return String(v || '').toLowerCase().indexOf(q) !== -1; });
  }) : h;
  var list = $('historyList');
  list.textContent = '';
  if (!rows.length) {
    list.appendChild(h.length
      ? emptyState('⌕', 'No matches', 'Try a different asset code, name or location.')
      : emptyState('✦', 'Nothing yet', 'Your submitted counts for this session will appear here.'));
    return;
  }
  var icons = { Good: '✓', Damaged: '!', Missing: '?', Obsolete: '⌛' };
  rows.forEach(function (x) {
    var cond = x.condition || 'Good';
    var status = x.status || 'Submitted';
    var right = el('div', 'aright');
    right.appendChild(el('span', 'ap', '×' + (x.physicalQty == null ? 0 : x.physicalQty)));
    right.appendChild(el('span', 'spill ' + status.split(' ')[0], status));
    var meta = (x.stockTakeId || '') + ' · ' + (x.location || '') + ' · ' + fmtDT(x.submittedDate);
    list.appendChild(itemRow(icons[cond] || '•', cond, x.assetCode, x.assetName, meta, right));
  });
}

/* ---------------- TABS ---------------- */
function switchTab(id) {
  document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.id === id); });
  document.querySelectorAll('.tab-btn').forEach(function (b) { b.classList.toggle('active', b.dataset.tab === id); });
  if (id !== 'tabCount') { stopCamera(); stopScan(); }
  if (id === 'tabHistory') { loadHistory(); }
  if (id === 'tabQueue') { renderQueue(); }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ---------------- EVENTS ---------------- */
function bind() {
  // Login
  $('btnSendOtp').addEventListener('click', function () { sendOtp(false); });
  $('btnResend').addEventListener('click', function () { sendOtp(true); });
  $('btnVerifyOtp').addEventListener('click', verifyOtp);
  $('btnBackEmail').addEventListener('click', function () {
    clearInterval(S.resendTimer);
    $('stepOtp').classList.add('hidden');
    $('stepEmail').classList.remove('hidden');
  });
  $('inpEmail').addEventListener('keydown', function (e) { if (e.key === 'Enter') { sendOtp(false); } });
  $('inpOtp').addEventListener('input', function (e) {
    e.target.classList.remove('invalid');
    e.target.value = e.target.value.replace(/\D/g, '').slice(0, 6);
    if (e.target.value.length === 6) { verifyOtp(); }
  });

  // Logout (theme toggle is handled inside index.html)
  $('btnLogout').addEventListener('click', function () { if (confirm('Sign out of StockSense?')) { logout(); } });

  // Asset
  $('selLocation').addEventListener('change', function () {
    store.set('ops_st_location', $('selLocation').value);
    $('selLocation').classList.remove('invalid');
    updateProgress();
  });
  ['inpAssetCode', 'inpAssetName', 'inpRemarks', 'inpQty'].forEach(function (id) {
    $(id).addEventListener('input', function () { $(id).classList.remove('invalid'); updateProgress(); });
  });
  $('inpAssetCode').addEventListener('blur', function (e) { e.target.value = e.target.value.trim().toUpperCase(); updateProgress(); });
  $('btnScan').addEventListener('click', startScan);
  $('btnStopScan').addEventListener('click', stopScan);

  // Count
  document.querySelectorAll('[data-delta]').forEach(function (b) {
    b.addEventListener('click', function () {
      $('inpQty').value = Math.max(0, (parseInt($('inpQty').value, 10) || 0) + parseInt(b.dataset.delta, 10));
      buzz(8);
      updateProgress();
    });
  });
  $('btnQtyReset').addEventListener('click', function () { $('inpQty').value = 0; updateProgress(); });
  $('inpQty').addEventListener('focus', function (e) { e.target.select(); });
  document.querySelectorAll('.seg').forEach(function (s) {
    s.addEventListener('click', function () { setCondition(s.dataset.cond); });
  });

  // Photos
  $('btnOpenCam').addEventListener('click', openCamera);
  $('btnCamClose').addEventListener('click', stopCamera);
  $('btnCapture').addEventListener('click', capture);
  $('btnCamSwitch').addEventListener('click', switchCamera);
  $('btnPickPhoto').addEventListener('click', function () { $('fileGallery').click(); });
  $('fileCamera').addEventListener('change', onFilePick);
  $('fileGallery').addEventListener('change', onFilePick);

  // Submit
  $('btnGeo').addEventListener('click', getGeo);
  $('btnSubmit').addEventListener('click', submitCount);
  $('btnClear').addEventListener('click', function () { if (confirm('Clear all entered details and photos?')) { clearForm(); } });
  $('btnNext').addEventListener('click', closeSuccess);
  $('successSheet').addEventListener('click', function (e) { if (e.target.id === 'successSheet') { closeSuccess(); } });

  // History + queue
  $('btnRefreshHistory').addEventListener('click', loadHistory);
  $('inpFilter').addEventListener('input', renderHistory);
  $('btnSync').addEventListener('click', function () { syncQueue(false); });

  // Navigation
  document.querySelectorAll('.tab-btn').forEach(function (b) {
    b.addEventListener('click', function () { switchTab(b.dataset.tab); });
  });

  // Network + lifecycle
  window.addEventListener('online', function () { updateNet(); toast('Back online — syncing', 'success'); syncQueue(true); });
  window.addEventListener('offline', function () { updateNet(); toast('You are offline. Submissions will be queued.', 'warn'); });
  document.
