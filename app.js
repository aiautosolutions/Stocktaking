/* =========================================================
   StockSense — Operations Stock Take (single SharePoint list)
   ========================================================= */
'use strict';

/* ---------------- CONFIG ---------------- */
const CONFIG = {
  COMPANY_NAME: 'Your Company',   // logo text + photo stamp
  COMPANY_SHORT: '',              // optional badge text e.g. 'ACME'. Blank = initials
  AUTH_ENABLED: true,             // false = skip OTP (testing only)

  FLOW_SEND_OTP:   'https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/22/workflows/587bb4ec5a9141d4b2b61445d9841139/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=dMgM0-vkKVd11vFfkQwMaUdsrnId4PhwMAI0ulkC9tE',
  FLOW_VERIFY_OTP: 'https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/31/workflows/296ff3060c1a449582e85c81e154c6df/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tdeev6xhU5HgYOU0zcHANo0FBK2kfDEqVrptgt1bkjQ',
  FLOW_CONFIG:     'PASTE_OPS_ST_GetConfig_HTTP_URL',
  FLOW_SUBMIT:     'PASTE_OPS_ST_SubmitCount_HTTP_URL',
  FLOW_HISTORY:    'PASTE_OPS_ST_MySubmissions_HTTP_URL',

  QR_LIBRARY: 'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js',
  DEMO_LOCATIONS: ['Head Office', 'Store Room', 'Archive Room', 'Warehouse'],
  REMARK_REQUIRED_FOR: ['Damaged', 'Missing', 'Obsolete'],
  MIN_PHOTOS: 1,
  MAX_PHOTOS: 5,
  PHOTO_MAX_DIM: 1600,
  PHOTO_QUALITY: 0.72,
  LOGIN_HOURS: 10,
  RESEND_SECONDS: 30
};

/* ---------------- STATE ---------------- */
const S = {
  user: null, pendingEmail: '', busy: false, syncing: false,
  session: '', locations: [],
  condition: 'Good', photos: [],
  stream: null, facing: 'environment', stampTimer: null,
  scanner: null, scanHandled: false,
  geo: null, history: [], resendTimer: null
};

/* ---------------- HELPERS ---------------- */
const $ = id => document.getElementById(id);
const num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const safe = s => String(s ?? '').replace(/[^A-Za-z0-9-]/g, '_');
const uid = () => (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : 'c' + Date.now() + Math.random().toString(16).slice(2);
const buzz = ms => { if (navigator.vibrate) navigator.vibrate(ms); };
const isSet = url => !!url && !String(url).startsWith('PASTE_');
const fmtDT = iso => iso ? new Date(iso).toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false }) : '—';

const store = {
  get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } }
};

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined && text !== null) e.textContent = text;
  return e;
}
function toast(msg, type = 'info', ms = 3400) {
  const t = el('div', 'toast ' + type);
  const icon = type === 'error' ? '!' : (type === 'warn' ? '•' : '✓');
  t.append(el('span', 'tdot', icon), el('span', '', msg));
  $('toastHost').appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; setTimeout(() => t.remove(), 300); }, ms);
}
function loading(on, text = 'Loading…') {
  $('loaderText').textContent = text;
  $('loader').classList.toggle('hidden', !on);
}
function showView(id) {
  document.querySelectorAll('.view').forEach(v => v.classList.toggle('active', v.id === id));
}
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error('Failed to load ' + src));
    document.head.appendChild(s);
  });
}
async function api(url, body, timeoutMs = 60000) {
  if (!isSet(url)) throw new Error('Flow URL is not configured in app.js');
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: ctrl.signal
    });
    const text = await res.text();
    let data = {};
    try { data = text ? JSON.parse(text) : {}; } catch (e) { data = { raw: text }; }
    if (!res.ok) throw new Error(data.message || ('Request failed (' + res.status + ')'));
    return data;
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Request timed out');
    if (e instanceof TypeError) throw new Error('Cannot reach the server. Check connection or flow CORS settings.');
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

/* ---------------- BRANDING + THEME ---------------- */
function applyBranding() {
  const name = CONFIG.COMPANY_NAME || 'Company';
  const short = CONFIG.COMPANY_SHORT || name.split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
  document.querySelectorAll('.js-company').forEach(e => { e.textContent = name; });
  document.querySelectorAll('.js-logo').forEach(e => { e.textContent = short; });
  document.querySelectorAll('.js-mark').forEach(e => { e.textContent = name.toUpperCase(); });
  document.title = name + ' · StockSense';
}
function applyTheme(t) {
  document.documentElement.setAttribute('data-theme', t);
  const m = document.querySelector('meta[name="theme-color"]');
  if (m) m.setAttribute('content', t === 'dark' ? '#07071a' : '#f5f6ff');
  store.set('ops_st_theme', t);
}
function toggleTheme() {
  applyTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
}
function greeting() {
  const h = new Date().getHours();
  return h < 12 ? 'Good morning' : (h < 18 ? 'Good afternoon' : 'Good evening');
}
function todayKey() { return 'ops_st_today_' + new Date().toISOString().slice(0, 10); }
function renderToday() { $('lblToday').textContent = num(store.get(todayKey())); }
function bumpToday() { store.set(todayKey(), String(num(store.get(todayKey())) + 1)); renderToday(); }

/* ---------------- OFFLINE QUEUE (IndexedDB) ---------------- */
const DB = {
  db: null,
  open() {
    return new Promise((resolve, reject) => {
      if (this.db) return resolve(this.db);
      if (!window.indexedDB) return reject(new Error('Offline storage not supported'));
      const r = indexedDB.open('stocksense_ops', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('queue', { keyPath: 'clientId' });
      r.onsuccess = () => { this.db = r.result; resolve(this.db); };
      r.onerror = () => reject(r.error);
    });
  },
  async run(mode, fn) {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('queue', mode);
      const req = fn(tx.objectStore('queue'));
      tx.oncomplete = () => resolve(req ? req.result : undefined);
      tx.onerror = () => reject(tx.error);
    });
  },
  put(rec) { return this.run('readwrite', s => s.put(rec)); },
  del(id) { return this.run('readwrite', s => s.delete(id)); },
  all() { return this.run('readonly', s => s.getAll()); }
};

/* ---------------- AUTH ---------------- */
function saveUser(u) {
  S.user = u;
  store.set('ops_st_user', JSON.stringify(Object.assign({}, u, { exp: Date.now() + CONFIG.LOGIN_HOURS * 3600e3 })));
}
function loadUser() {
  try {
    const u = JSON.parse(store.get('ops_st_user'));
    if (u && u.exp > Date.now()) return u;
  } catch (e) { /* ignore */ }
  store.del('ops_st_user');
  return null;
}
function startResendTimer() {
  const b = $('btnResend');
  let left = CONFIG.RESEND_SECONDS;
  clearInterval(S.resendTimer);
  b.disabled = true;
  b.textContent = 'Resend in ' + left + 's';
  S.resendTimer = setInterval(() => {
    left--;
    if (left <= 0) {
      clearInterval(S.resendTimer);
      b.disabled = false;
      b.textContent = 'Resend code';
    } else {
      b.textContent = 'Resend in ' + left + 's';
    }
  }, 1000);
}
async function sendOtp(isResend) {
  if (S.busy) return;
  const email = isResend ? S.pendingEmail : $('inpEmail').value.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    $('inpEmail').classList.add('invalid');
    return toast('Enter a valid email address', 'error');
  }
  $('inpEmail').classList.remove('invalid');

  if (!CONFIG.AUTH_ENABLED) { saveUser({ email }); return enterApp(); }

  S.busy = true;
  loading(true, 'Sending verification code…');
  try {
    const r = await api(CONFIG.FLOW_SEND_OTP, { email });
    if (r.success === false) throw new Error(r.message || 'Unable to send code');
    S.pendingEmail = email;
    $('lblOtpEmail').textContent = email;
    $('stepEmail').classList.add('hidden');
    $('stepOtp').classList.remove('hidden');
    $('inpOtp').value = '';
    $('inpOtp').focus();
    startResendTimer();
    toast(isResend ? 'A new code has been sent' : 'Verification code sent', 'success');
  } catch (e) {
    toast(e.message, 'error', 5000);
  } finally {
    S.busy = false;
    loading(false);
  }
}
async function verifyOtp() {
  if (S.busy) return;
  const otp = $('inpOtp').value.trim();
  if (!/^\d{6}$/.test(otp)) return toast('Enter the 6-digit code', 'error');
  S.busy = true;
  loading(true, 'Verifying…');
  try {
    const r = await api(CONFIG.FLOW_VERIFY_OTP, { email: S.pendingEmail, otp });
    const ok = r.success === true || r.verified === true || r.valid === true;
    if (!ok) throw new Error(r.message || 'Invalid or expired code');
    clearInterval(S.resendTimer);
    saveUser({ email: S.pendingEmail });
    S.busy = false;
    loading(false);
    enterApp();
  } catch (e) {
    S.busy = false;
    loading(false);
    $('inpOtp').classList.add('invalid');
    toast(e.message, 'error');
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
  if (navigator.onLine) syncQueue(true);
}
async function loadConfig() {
  let cached = null;
  try { cached = JSON.parse(store.get('ops_st_config') || 'null'); } catch (e) { /* ignore */ }

  if (!isSet(CONFIG.FLOW_CONFIG)) {
    S.session = 'ST-DEMO';
    S.locations = CONFIG.DEMO_LOCATIONS.slice();
    toast('Demo mode: Get Config flow URL not set yet', 'warn', 4500);
  } else if (navigator.onLine) {
    loading(true, 'Preparing your workspace…');
    try {
      const r = await api(CONFIG.FLOW_CONFIG, { email: S.user.email });
      S.session = r.session || '';
      S.locations = Array.isArray(r.locations) ? r.locations : [];
      store.set('ops_st_config', JSON.stringify({ session: S.session, locations: S.locations }));
    } catch (e) {
      if (cached) {
        S.session = cached.session;
        S.locations = cached.locations;
        toast('Using saved configuration', 'warn');
      } else {
        toast('Unable to load configuration: ' + e.message, 'error', 5000);
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
  const sel = $('selLocation');
  sel.innerHTML = '';
  sel.appendChild(new Option('Select location…', ''));
  S.locations.forEach(l => sel.appendChild(new Option(l, l)));
  const pref = store.get('ops_st_location');
  if (pref && S.locations.includes(pref)) sel.value = pref;
}

/* ---------------- NETWORK ---------------- */
function updateNet() {
  const on = navigator.onLine;
  $('netBadge').className = 'net ' + (on ? 'online' : 'offline');
  $('netBadge').querySelector('b').textContent = on ? 'Online' : 'Offline';
}

/* ---------------- FORM ---------------- */
function qtyValid() {
  const q = String($('inpQty').value).trim();
  return q !== '' && !isNaN(q) && num(q) >= 0 && Number.isInteger(num(q));
}
function updateProgress() {
  const loc = $('selLocation').value;
  const code = $('inpAssetCode').value.trim();
  const name = $('inpAssetName').value.trim();
  const qOk = qtyValid();
  const remOk = !CONFIG.REMARK_REQUIRED_FOR.includes(S.condition) || !!$('inpRemarks').value.trim();
  const steps = [!!(loc && code && name), qOk && remOk, S.photos.length >= CONFIG.MIN_PHOTOS];
  const done = steps.filter(Boolean).length;

  steps.forEach((ok, i) => $('ps' + (i + 1)).classList.toggle('done', ok));
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
  document.querySelectorAll('.seg').forEach(s => s.classList.toggle('active', s.dataset.cond === c));
  $('remarkReq').classList.toggle('hidden', !CONFIG.REMARK_REQUIRED_FOR.includes(c));
  if (c === 'Missing') $('inpQty').value = 0;
  updateProgress();
}
function clearForm() {
  stopCamera();
  $('inpAssetCode').value = '';
  $('inpAssetName').value = '';
  $('inpQty').value = 1;
  $('inpRemarks').value = '';
  ['inpAssetCode', 'inpAssetName', 'selLocation', 'inpRemarks'].forEach(id => $(id).classList.remove('invalid'));
  S.photos = [];
  setCondition('Good');
  renderPhotos();
}

/* ---------------- BARCODE SCANNER ---------------- */
async function ensureQrLib() {
  if (window.Html5Qrcode) return;
  await loadScript(CONFIG.QR_LIBRARY);
}
async function startScan() {
  stopCamera();
  await stopScan();
  try {
    loading(true, 'Starting scanner…');
    await ensureQrLib();
  } catch (e) {
    loading(false);
    return toast('Scanner could not load. Check your connection.', 'error');
  }
  loading(false);
  $('scanWrap').classList.remove('hidden');
  S.scanHandled = false;
  S.scanner = new Html5Qrcode('reader');
  try {
    await S.scanner.start(
      { facingMode: 'environment' },
      { fps: 12, qrbox: (w, h) => ({ width: Math.max(140, Math.floor(w * 0.8)), height: Math.max(90, Math.floor(h * 0.4)) }) },
      onScan,
      () => {}
    );
  } catch (e) {
    toast('Camera not available for scanning', 'error');
    stopScan();
  }
}
async function stopScan() {
  if (S.scanner) {
    try {
      if (S.scanner.isScanning) await S.scanner.stop();
      S.scanner.clear();
    } catch (e) { /* ignore */ }
    S.scanner = null;
  }
  $('scanWrap').classList.add('hidden');
}
async function onScan(text) {
  if (S.scanHandled) return;
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
  const ts = new Date().toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
  const gps = S.geo ? ' · ' + S.geo.lat.toFixed(5) + ', ' + S.geo.lng.toFixed(5) : '';
  return [
    (CONFIG.COMPANY_NAME || '').toUpperCase() + ' · ' + (S.session || '-'),
    ($('selLocation').value || '-') + ' · ' + ($('inpAssetCode').value.trim() || '-'),
    (S.user && S.user.email) || '-',
    ts + gps
  ];
}
function updateStamp() {
  const box = $('camStamp');
  box.textContent = '';
  stampLines().forEach((l, i) => {
    if (i) box.appendChild(document.createElement('br'));
    box.appendChild(document.createTextNode(l));
  });
}
async function openCamera() {
  if (S.photos.length >= CONFIG.MAX_PHOTOS) return toast('Maximum ' + CONFIG.MAX_PHOTOS + ' photos', 'warn');
  await stopScan();
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) { $('fileCamera').click(); return; }
  try {
    stopCamera();
    S.stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: S.facing }, width: { ideal: 1920 }, height: { ideal: 1080 } },
      audio: false
    });
    const v = $('camVideo');
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
  if (S.stream) { S.stream.getTracks().forEach(t => t.stop()); S.stream = null; }
  if (S.stampTimer) { clearInterval(S.stampTimer); S.stampTimer = null; }
  $('camVideo').srcObject = null;
  $('camWrap').classList.add('hidden');
}
async function switchCamera() {
  S.facing = S.facing === 'environment' ? 'user' : 'environment';
  if (S.stream) { stopCamera(); await openCamera(); }
}
function processImage(src, w, h) {
  const scale = Math.min(1, CONFIG.PHOTO_MAX_DIM / Math.max(w, h));
  const cw = Math.round(w * scale);
  const ch = Math.round(h * scale);
  const c = $('workCanvas');
  c.width = cw;
  c.height = ch;
  const ctx = c.getContext('2d');
  ctx.drawImage(src, 0, 0, cw, ch);

  const lines = stampLines();
  const fs = Math.max(14, Math.round(cw / 48));
  const pad = Math.round(fs * 0.8);
  const lh = Math.round(fs * 1.4);
  const boxH = pad * 2 + lh * lines.length;

  ctx.fillStyle = 'rgba(12,10,40,0.72)';
  ctx.fillRect(0, ch - boxH, cw, boxH);
  const g = ctx.createLinearGradient(0, ch - boxH, 0, ch);
  g.addColorStop(0, '#7c5cff');
  g.addColorStop(1, '#22d3ee');
  ctx.fillStyle = g;
  ctx.fillRect(0, ch - boxH, Math.max(4, Math.round(fs / 3)), boxH);

  ctx.font = '600 ' + fs + 'px "Space Grotesk", Arial, sans-serif';
  ctx.textBaseline = 'top';
  lines.forEach((t, i) => {
    ctx.fillStyle = i === 0 ? '#67e8f9' : '#FFFFFF';
    ctx.fillText(t, pad + Math.round(fs / 2), ch - boxH + pad + i * lh);
  });
  return c.toDataURL('image/jpeg', CONFIG.PHOTO_QUALITY);
}
function capture() {
  const v = $('camVideo');
  if (!v.videoWidth) return toast('Camera is starting…', 'warn');
  addPhoto(processImage(v, v.videoWidth, v.videoHeight));
  const f = $('camFlash');
  f.classList.add('on');
  setTimeout(() => f.classList.remove('on'), 60);
  buzz(35);
  if (S.photos.length >= CONFIG.MAX_PHOTOS) { stopCamera(); toast('Photo limit reached'); }
}
function loadImg(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Unreadable image')); };
    img.src = url;
  });
}
async function onFilePick(e) {
  const files = Array.from(e.target.files || []);
  e.target.value = '';
  for (const f of files) {
    if (S.photos.length >= CONFIG.MAX_PHOTOS) { toast('Maximum ' + CONFIG.MAX_PHOTOS + ' photos', 'warn'); break; }
    if (!f.type.startsWith('image/')) continue;
    try {
      const img = await loadImg(f);
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
  const grid = $('photoGrid');
  grid.textContent = '';
  S.photos.forEach((p, i) => {
    const wrap = el('div', 'thumb');
    const img = document.createElement('img');
    img.src = p.dataUrl;
    img.alt = 'Photo ' + (i + 1);
    const no = el('span', 'no', '#' + (i + 1));
    const rm = el('button', 'rm', '✕');
    rm.type = 'button';
    rm.setAttribute('aria-label', 'Remove photo');
    rm.onclick = () => { S.photos = S.photos.filter(x => x.id !== p.id); renderPhotos(); };
    wrap.append(img, no, rm);
    grid.appendChild(wrap);
  });
  $('photoCounter').textContent = S.photos.length + ' / ' + CONFIG.MAX_PHOTOS;
  const full = S.photos.length >= CONFIG.MAX_PHOTOS;
  $('btnOpenCam').disabled = full;
  $('btnPickPhoto').disabled = full;
  updateProgress();
}

/* ---------------- GPS ---------------- */
function getGeo() {
  return new Promise(resolve => {
    const b = $('geoBadge');
    if (!navigator.geolocation) { b.textContent = '📍 Location not supported'; return resolve(null); }
    b.textContent = '📍 Locating…';
    b.className = 'geo';
    navigator.geolocation.getCurrentPosition(p => {
      S.geo = { lat: p.coords.latitude, lng: p.coords.longitude, acc: Math.round(p.coords.accuracy) };
      b.textContent = '📍 Location captured (±' + S.geo.acc + ' m)';
      b.className = 'geo ok';
      resolve(S.geo);
    }, () => {
      b.textContent = '📍 Location unavailable (optional)';
      resolve(null);
    }, { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 });
  });
}

/* ---------------- SUBMIT ---------------- */
function validate() {
  const checks = [
    ['selLocation', !$('selLocation').value, 'Select a location'],
    ['inpAssetCode', !$('inpAssetCode').value.trim(), 'Enter the asset code'],
    ['inpAssetName', !$('inpAssetName').value.trim(), 'Enter the asset name']
  ];
  for (const c of checks) {
    $(c[0]).classList.toggle('invalid', c[1]);
    if (c[1]) { $(c[0]).focus(); toast(c[2], 'error'); return false; }
  }
  if (!qtyValid()) { toast('Enter a valid whole-number quantity', 'error'); return false; }
  if (CONFIG.REMARK_REQUIRED_FOR.includes(S.condition) && !$('inpRemarks').value.trim()) {
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
  if (S.busy || !validate()) return;
  S.busy = true;
  stopCamera();

  const code = $('inpAssetCode').value.trim().toUpperCase();
  const stamp = Date.now();
  const rec = {
    clientId: uid(),
    location: $('selLocation').value,
    assetCode: code,
    assetName: $('inpAssetName').value.trim(),
    physicalQty: parseInt($('inpQty').value, 10),
    condition: S.condition,
    remarks: $('inpRemarks').value.trim(),
    submittedEmail: S.user.email,
    submittedAt: new Date().toISOString(),
    photos: S.photos.map((p, i) => ({
      fileName: safe(code) + '_' + stamp + '_' + (i + 1) + '.jpg',
      content: p.dataUrl.split(',')[1]
    }))
  };

  if (!navigator.onLine) {
    try { await DB.put(rec); } catch (e) { S.busy = false; return toast('Could not save offline: ' + e.message, 'error'); }
    bumpToday();
    S.busy = false;
    clearForm();
    updateQueueBadge();
    showSuccess(null, true);
    return;
  }

  loading(true, 'Submitting count and photo evidence…');
  try {
    const r = await api(CONFIG.FLOW_SUBMIT, rec, 120000);
    if (r.success === false) throw new Error(r.message || 'Submission failed');
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
      toast('Submission failed: ' + e.message, 'error', 5000);
    }
  } finally {
    S.busy = false;
    loading(false);
    updateQueueBadge();
  }
}

/* ---------------- SYNC ---------------- */
async function syncQueue(silent) {
  if (S.syncing || !isSet(CONFIG.FLOW_SUBMIT)) return;
  if (!navigator.onLine) { if (!silent) toast('Still offline', 'warn'); return; }
  let all = [];
  try { all = await DB.all(); } catch (e) { return; }
  if (!all.length) { if (!silent) toast('Nothing to sync'); return; }

  S.syncing = true;
  if (!silent) loading(true, 'Syncing ' + all.length + ' record(s)…');
  let ok = 0;
  let fail = 0;
  for (const rec of all) {
    try {
      const r = await api(CONFIG.FLOW_SUBMIT, rec, 120000);
      if (r.success === false) throw new Error(r.message);
      await DB.del(rec.clientId);
      ok++;
    } catch (e) {
      fail++;
    }
  }
  S.syncing = false;
  if (!silent) loading(false);
  updateQueueBadge();
  if ($('tabQueue').classList.contains('active')) renderQueue();
  if (ok) toast(ok + ' queued record(s) synced', 'success');
  if (fail) toast(fail + ' record(s) still pending', 'warn');
}
function emptyState(icon, title, copy) {
  const e = el('div', 'empty');
  e.append(el('div', 'eic', icon), el('h3', 'etitle', title), el('p', 'ecopy', copy));
  return e;
}
function itemRow(icon, iconCls, code, name, meta, right) {
  const row = el('div', 'aitem');
  const main = el('div', 'amain');
  const t = el('p', 'at');
  t.append(el('span', 'code', code), document.createTextNode(name || ''));
  main.append(t, el('p', 'am', meta));
  row.append(el('div', 'aic ' + iconCls, icon), main, right);
  return row;
}
async function renderQueue() {
  let all = [];
  try { all = await DB.all(); } catch (e) { /* ignore */ }
  const list = $('queueList');
  list.textContent = '';
  if (!all.length) { list.appendChild(emptyState('☁', 'All caught up', 'Nothing is waiting to sync.')); return; }
  all.forEach(r => {
    const del = el('button', 'del', '✕');
    del.type = 'button';
    del.setAttribute('aria-label', 'Discard');
    del.onclick = async () => {
      if (!confirm('Discard this queued record? This cannot be undone.')) return;
      await DB.del(r.clientId);
      renderQueue();
      updateQueueBadge();
    };
    const meta = r.location + ' · Qty ' + r.physicalQty + ' · ' + r.photos.length + ' photo(s) · ' + fmtDT(r.submittedAt);
    list.appendChild(itemRow('⇪', 'queued', r.assetCode, r.assetName, meta, del));
  });
}
async function updateQueueBadge() {
  try {
    const n = (await DB.all()).length;
    $('queueBadge').textContent = n;
    $('queueBadge').classList.toggle('hidden', !n);
  } catch (e) { /* ignore */ }
}

/* ---------------- HISTORY ---------------- */
async function loadHistory() {
  if (!isSet(CONFIG.FLOW_HISTORY)) { toast('My Submissions flow URL is not set yet', 'warn'); return renderHistory(); }
  if (!navigator.onLine) { toast('Offline — history unavailable', 'warn'); return renderHistory(); }
  loading(true, 'Loading your submissions…');
  try {
    const r = await api(CONFIG.FLOW_HISTORY, { email: S.user.email });
    S.history = (r.items || []).sort((a, b) => new Date(b.submittedDate) - new Date(a.submittedDate));
  } catch (e) {
    toast('Unable to load history: ' + e.message, 'error');
  } finally {
    loading(false);
    renderHistory();
  }
}
function renderHistory() {
  const h = S.history;
  const good = h.filter(x => (x.condition || 'Good') === 'Good').length;
  const pct = h.length ? Math.round(good / h.length * 100) : 0;
  $('kpiTotal').textContent = h.length;
  $('kpiQty').textContent = h.reduce((s, x) => s + num(x.physicalQty), 0).toLocaleString('en-MY');
  $('kpiExc').textContent = h.length - good;
  $('kpiVer').textContent = h.filter(x => x.status === 'Verified').length + ' verified';
  $('ringPct').textContent = pct + '%';
  $('ringCap').textContent = h.length ? good + ' of ' + h.length + ' assets in good condition' : 'No records yet';
  const circ = 2 * Math.PI * 31;
  $('ringVal').setAttribute('stroke-dasharray', circ.toFixed(1));
  $('ringVal').setAttribute('stroke-dashoffset', (circ * (1 - pct / 100)).toFixed(1));

  const q = $('inpFilter').value.trim().toLowerCase();
  const rows = q
    ? h.filter(x => [x.assetCode, x.assetName, x.location].some(v => String(v || '').toLowerCase().includes(q)))
    : h;
  const list = $('historyList');
  list.textContent = '';
  if (!rows.length) {
    list.appendChild(h.length
      ? emptyState('⌕', 'No matches', 'Try a different asset code, name or location.')
      : emptyState('✦', 'Nothing yet', 'Your submitted counts for this session will appear here.'));
    return;
  }
  const icons = { Good: '✓', Damaged: '!', Missing: '?', Obsolete: '⌛' };
  rows.forEach(x => {
    const cond = x.condition || 'Good';
    const status = x.status || 'Submitted';
    const right = el('div', 'aright');
    right.append(
      el('span', 'ap', '×' + (x.physicalQty == null ? 0 : x.physicalQty)),
      el('span', 'spill ' + status.split(' ')[0], status)
    );
    const meta = (x.stockTakeId || '') + ' · ' + (x.location || '') + ' · ' + fmtDT(x.submittedDate);
    list.appendChild(itemRow(icons[cond] || '•', cond, x.assetCode, x.assetName, meta, right));
  });
}

/* ---------------- TABS ---------------- */
function switchTab(id) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.id === id));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === id));
 
