/* =========================================================
   CGC StockSense — Operations Stock Take (single SharePoint list)
   ========================================================= */
'use strict';

/* ---------------- CONFIG ---------------- */
const CONFIG = {
  AUTH_ENABLED: true,                         // false = skip OTP (testing only)
  FLOW_SEND_OTP:   'PASTE_PC_SendOTP_HTTP_URL',
  FLOW_VERIFY_OTP: 'PASTE_PC_VerifyOTP_HTTP_URL',
  FLOW_CONFIG:     'PASTE_OPS_ST_GetConfig_HTTP_URL',
  FLOW_SUBMIT:     'PASTE_OPS_ST_SubmitCount_HTTP_URL',
  FLOW_HISTORY:    'PASTE_OPS_ST_MySubmissions_HTTP_URL',
  REMARK_REQUIRED_FOR: ['Damaged', 'Missing', 'Obsolete'],
  MIN_PHOTOS: 1,
  MAX_PHOTOS: 5,
  PHOTO_MAX_DIM: 1600,
  PHOTO_QUALITY: 0.72,
  LOGIN_HOURS: 10
};

/* ---------------- STATE ---------------- */
const S = {
  user: null, pendingEmail: '', busy: false, syncing: false,
  session: '', locations: [],
  condition: 'Good', photos: [],
  stream: null, facing: 'environment', stampTimer: null,
  scanner: null, scanHandled: false,
  geo: null, history: []
};

/* ---------------- HELPERS ---------------- */
const $ = id => document.getElementById(id);
const num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const safe = s => String(s ?? '').replace(/[^A-Za-z0-9-]/g, '_');
const uid = () => (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : `c${Date.now()}${Math.random().toString(16).slice(2)}`;
const buzz = ms => { if (navigator.vibrate) navigator.vibrate(ms); };
const fmtDT = iso => iso ? new Date(iso).toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false }) : '—';

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined && text !== null) e.textContent = text;
  return e;
}
function toast(msg, type = 'info', ms = 3400) {
  const t = el('div', `toast ${type}`, msg);
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
async function api(url, body, timeoutMs = 60000) {
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
    try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }
    if (!res.ok) throw new Error(data.message || `Request failed (${res.status})`);
    return data;
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Request timed out');
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

/* ---------------- OFFLINE QUEUE (IndexedDB) ---------------- */
const DB = {
  db: null,
  open() {
    return new Promise((resolve, reject) => {
      if (this.db) return resolve(this.db);
      const r = indexedDB.open('cgc_stocksense_ops', 1);
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
  localStorage.setItem('ops_st_user', JSON.stringify({ ...u, exp: Date.now() + CONFIG.LOGIN_HOURS * 3600e3 }));
}
function loadUser() {
  try {
    const u = JSON.parse(localStorage.getItem('ops_st_user'));
    if (u && u.exp > Date.now()) return u;
  } catch { /* ignore */ }
  localStorage.removeItem('ops_st_user');
  return null;
}
async function sendOtp() {
  if (S.busy) return;
  const email = $('inpEmail').value.trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { $('inpEmail').classList.add('invalid'); return toast('Enter a valid email address', 'error'); }
  $('inpEmail').classList.remove('invalid');

  if (!CONFIG.AUTH_ENABLED) { saveUser({ email }); return enterApp(); }

  S.busy = true; loading(true, 'Sending verification code…');
  try {
    const r = await api(CONFIG.FLOW_SEND_OTP, { email });
    if (r.success === false) throw new Error(r.message || 'Unable to send code');
    S.pendingEmail = email;
    $('lblOtpEmail').textContent = email;
    $('stepEmail').classList.add('hidden');
    $('stepOtp').classList.remove('hidden');
    $('inpOtp').value = '';
    $('inpOtp').focus();
    toast('Verification code sent', 'success');
  } catch (e) {
    toast(e.message, 'error');
  } finally {
    S.busy = false; loading(false);
  }
}
async function verifyOtp() {
  if (S.busy) return;
  const otp = $('inpOtp').value.trim();
  if (!/^\d{6}$/.test(otp)) return toast('Enter the 6-digit code', 'error');
  S.busy = true; loading(true, 'Verifying…');
  try {
    const r = await api(CONFIG.FLOW_VERIFY_OTP, { email: S.pendingEmail, otp });
    const ok = r.success === true || r.verified === true || r.valid === true;
    if (!ok) throw new Error(r.message || 'Invalid or expired code');
    saveUser({ email: S.pendingEmail });
    S.busy = false; loading(false);
    enterApp();
  } catch (e) {
    S.busy = false; loading(false);
    toast(e.message, 'error');
  }
}
function logout() {
  stopCamera(); stopScan();
  localStorage.removeItem('ops_st_user');
  S.user = null; S.history = [];
  clearForm();
  $('stepOtp').classList.add('hidden');
  $('stepEmail').classList.remove('hidden');
  showView('viewLogin');
}

/* ---------------- BOOT ---------------- */
async function enterApp() {
  showView('viewApp');
  $('lblUser').textContent = S.user.email;
  updateNet();
  await loadConfig();
  getGeo();
  updateQueueBadge();
  if (navigator.onLine) syncQueue(true);
}
async function loadConfig() {
  let cached = null;
  try { cached = JSON.parse(localStorage.getItem('ops_st_config') || 'null'); } catch { /* ignore */ }

  if (navigator.onLine) {
    loading(true, 'Preparing your workspace…');
    try {
      const r = await api(CONFIG.FLOW_CONFIG, { email: S.user.email });
      S.session = r.session || '';
      S.locations = Array.isArray(r.locations) ? r.locations : [];
      localStorage.setItem('ops_st_config', JSON.stringify({ session: S.session, locations: S.locations }));
    } catch (e) {
      if (cached) { S.session = cached.session; S.locations = cached.locations; toast('Using saved configuration', 'warn'); }
      else toast(`Unable to load configuration: ${e.message}`, 'error');
    } finally {
      loading(false);
    }
  } else if (cached) {
    S.session = cached.session; S.locations = cached.locations;
    toast('Offline — using saved configuration', 'warn');
  }

  $('lblSession').textContent = S.session || 'Not configured';
  const sel = $('selLocation');
  sel.innerHTML = '';
  sel.appendChild(new Option('Select location…', ''));
  S.locations.forEach(l => sel.appendChild(new Option(l, l)));
  const pref = localStorage.getItem('ops_st_location');
  if (pref && S.locations.includes(pref)) sel.value = pref;
}

/* ---------------- NETWORK ---------------- */
function updateNet() {
  const on = navigator.onLine;
  $('netBadge').className = `net ${on ? 'online' : 'offline'}`;
  $('netBadge').querySelector('b').textContent = on ? 'Online' : 'Offline';
}

/* ---------------- FORM ---------------- */
function setCondition(c) {
  S.condition = c;
  document.querySelectorAll('.seg').forEach(s => s.classList.toggle('active', s.dataset.cond === c));
  $('remarkReq').classList.toggle('hidden', !CONFIG.REMARK_REQUIRED_FOR.includes(c));
  if (c === 'Missing') $('inpQty').value = 0;
}
function clearForm() {
  stopCamera();
  $('inpAssetCode').value = '';
  $('inpAssetName').value = '';
  $('inpQty').value = 0;
  $('inpRemarks').value = '';
  ['inpAssetCode', 'inpAssetName', 'selLocation', 'inpRemarks'].forEach(id => $(id).classList.remove('invalid'));
  setCondition('Good');
  S.photos = [];
  renderPhotos();
}

/* ---------------- BARCODE SCANNER ---------------- */
async function startScan() {
  if (typeof Html5Qrcode === 'undefined') return toast('Scanner unavailable. Check your connection.', 'error');
  stopCamera();
  await stopScan();
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
  } catch {
    toast('Camera not available for scanning', 'error');
    stopScan();
  }
}
async function stopScan() {
  if (S.scanner) {
    try { if (S.scanner.isScanning) await S.scanner.stop(); S.scanner.clear(); } catch { /* ignore */ }
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
  toast(`Scanned ${$('inpAssetCode').value}`, 'success');
  $('inpAssetName').focus();
}

/* ---------------- CAMERA & PHOTOS ---------------- */
function stampLines() {
  const ts = new Date().toLocaleString('en-MY', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
  return [
    `CGC STOCKSENSE · ${S.session || '-'}`,
    `${$('selLocation').value || '-'} · ${$('inpAssetCode').value.trim() || '-'}`,
    `${S.user?.email || '-'}`,
    `${ts}${S.geo ? ` · ${S.geo.lat.toFixed(5)}, ${S.geo.lng.toFixed(5)}` : ''}`
  ];
}
function updateStamp() {
  const box = $('camStamp');
  box.textContent = '';
  stampLines().forEach((l, i) => { if (i) box.appendChild(document.createElement('br')); box.appendChild(document.createTextNode(l)); });
}
async function openCamera() {
  if (S.photos.length >= CONFIG.MAX_PHOTOS) return toast(`Maximum ${CONFIG.MAX_PHOTOS} photos`, 'warn');
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
  } catch {
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
  const cw = Math.round(w * scale), ch = Math.round(h * scale);
  const c = $('workCanvas');
  c.width = cw; c.height = ch;
  const ctx = c.getContext('2d');
  ctx.drawImage(src, 0, 0, cw, ch);

  const lines = stampLines();
  const fs = Math.max(14, Math.round(cw / 48));
  const pad = Math.round(fs * 0.8), lh = Math.round(fs * 1.4);
  const boxH = pad * 2 + lh * lines.length;

  ctx.fillStyle = 'rgba(14,26,43,0.78)';
  ctx.fillRect(0, ch - boxH, cw, boxH);
  ctx.fillStyle = '#A8834A';
  ctx.fillRect(0, ch - boxH, Math.max(4, Math.round(fs / 3)), boxH);

  ctx.font = `500 ${fs}px "JetBrains Mono", Menlo, monospace`;
  ctx.textBaseline = 'top';
  lines.forEach((t, i) => {
    ctx.fillStyle = i === 0 ? '#D9BF8E' : '#FFFFFF';
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
  const files = [...e.target.files];
  e.target.value = '';
  for (const f of files) {
    if (S.photos.length >= CONFIG.MAX_PHOTOS) { toast(`Maximum ${CONFIG.MAX_PHOTOS} photos`, 'warn'); break; }
    if (!f.type.startsWith('image/')) continue;
    try {
      const img = await loadImg(f);
      addPhoto(processImage(img, img.naturalWidth, img.naturalHeight));
    } catch {
      toast(`Could not read ${f.name}`, 'error');
    }
  }
}
function addPhoto(dataUrl) {
  S.photos.push({ id: uid(), dataUrl });
  renderPhotos();
}
function renderPhotos() {
  const grid = $('photoGrid');
  grid.textContent = '';
  S.photos.forEach((p, i) => {
    const wrap = el('div', 'thumb');
    const img = document.createElement('img');
    img.src = p.dataUrl;
    img.alt = `Photo ${i + 1}`;
    const no = el('span', 'no', `#${i + 1}`);
    const rm = el('button', 'rm', '✕');
    rm.setAttribute('aria-label', 'Remove photo');
    rm.onclick = () => { S.photos = S.photos.filter(x => x.id !== p.id); renderPhotos(); };
    wrap.append(img, no, rm);
    grid.appendChild(wrap);
  });
  $('photoCounter').textContent = `${S.photos.length} / ${CONFIG.MAX_PHOTOS}`;
  const full = S.photos.length >= CONFIG.MAX_PHOTOS;
  $('btnOpenCam').disabled = full;
  $('btnPickPhoto').disabled = full;
}

/* ---------------- GPS ---------------- */
function getGeo() {
  return new Promise(resolve => {
    const b = $('geoBadge');
    if (!navigator.geolocation) { b.textContent = 'Location not supported'; return resolve(null); }
    b.textContent = 'Locating…';
    b.className = 'geo';
    navigator.geolocation.getCurrentPosition(p => {
      S.geo = { lat: p.coords.latitude, lng: p.coords.longitude, acc: Math.round(p.coords.accuracy) };
      b.textContent = `Location captured (±${S.geo.acc} m)`;
      b.className = 'geo ok';
      resolve(S.geo);
    }, () => {
      b.textContent = 'Location unavailable (optional)';
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
  for (const [id, bad, msg] of checks) {
    $(id).classList.toggle('invalid', bad);
    if (bad) { $(id).focus(); toast(msg, 'error'); return false; }
  }
  const q = String($('inpQty').value).trim();
  if (q === '' || isNaN(q) || num(q) < 0 || !Number.isInteger(num(q))) { toast('Enter a valid whole-number quantity', 'error'); return false; }
  if (CONFIG.REMARK_REQUIRED_FOR.includes(S.condition) && !$('inpRemarks').value.trim()) {
    $('inpRemarks').classList.add('invalid'); $('inpRemarks').focus();
    toast(`Remarks are required for "${S.condition}"`, 'error'); return false;
  }
  $('inpRemarks').classList.remove('invalid');
  if (S.photos.length < CONFIG.MIN_PHOTOS) { toast(`At least ${CONFIG.MIN_PHOTOS} photo is required`, 'error'); return false; }
  if (!S.session) { toast('No active session. Contact the administrator.', 'error'); return false; }
  return true;
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
      fileName: `${safe(code)}_${stamp}_${i + 1}.jpg`,
      content: p.dataUrl.split(',')[1]
    }))
  };

  if (!navigator.onLine) {
    await DB.put(rec);
    toast('Offline — saved to queue. It will sync automatically.', 'warn', 4200);
    S.busy = false;
    clearForm();
    updateQueueBadge();
    return;
  }

  loading(true, 'Submitting count and photo evidence…');
  try {
    const r = await api(CONFIG.FLOW_SUBMIT, rec, 120000);
    if (r.success === false) throw new Error(r.message || 'Submission failed');
    toast(`Submitted${r.stockTakeId ? ' · ' + r.stockTakeId : ''}`, 'success');
    buzz([30, 50, 30]);
    clearForm();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  } catch {
    await DB.put(rec);
    toast('Connection issue — saved to offline queue', 'warn');
    clearForm();
  } finally {
    S.busy = false;
    loading(false);
    updateQueueBadge();
  }
}

/* ---------------- SYNC ---------------- */
async function syncQueue(silent = false) {
  if (S.syncing) return;
  if (!navigator.onLine) { if (!silent) toast('Still offline', 'warn'); return; }
  const all = await DB.all();
  if (!all.length) { if (!silent) toast('Nothing to sync'); return; }

  S.syncing = true;
  if (!silent) loading(true, `Syncing ${all.length} record(s)…`);
  let ok = 0, fail = 0;
  for (const rec of all) {
    try {
      const r = await api(CONFIG.FLOW_SUBMIT, rec, 120000);
      if (r.success === false) throw new Error(r.message);
      await DB.del(rec.clientId);
      ok++;
    } catch { fail++; }
  }
  S.syncing = false;
  if (!silent) loading(false);
  updateQueueBadge();
  if ($('tabQueue').classList.contains('active')) renderQueue();
  if (ok) toast(`${ok} queued record(s) synced`, 'success');
  if (fail) toast(`${fail} record(s) still pending`, 'warn');
}
async function renderQueue() {
  const all = await DB.all();
  const list = $('queueList');
  list.textContent = '';
  if (!all.length) { list.appendChild(el('div', 'empty', 'All caught up — nothing pending.')); return; }
  all.forEach(r => {
    const row = el('div', 'row');
    const main = el('div', 'row-main');
    const title = el('div', 'row-title');
    title.append(el('span', 'mono', r.assetCode), document.createTextNode(r.assetName));
    main.append(title, el('div', 'row-sub', `${r.location} · Qty ${r.physicalQty} · ${r.photos.length} photo(s) · ${fmtDT(r.submittedAt)}`));
    const del = el('button', 'icon-del', '✕');
    del.setAttribute('aria-label', 'Discard');
    del.onclick = async () => {
      if (!confirm('Discard this queued record? This cannot be undone.')) return;
      await DB.del(r.clientId);
      renderQueue(); updateQueueBadge();
    };
    row.append(main, del);
    list.appendChild(row);
  });
}
async function updateQueueBadge() {
  try {
    const n = (await DB.all()).length;
    $('queueBadge').textContent = n;
    $('queueBadge').classList.toggle('hidden', !n);
  } catch { /* ignore */ }
}

/* ---------------- HISTORY ---------------- */
async function loadHistory() {
  if (!navigator.onLine) { toast('Offline — history unavailable', 'warn'); return renderHistory(); }
  loading(true, 'Loading your submissions…');
  try {
    const r = await api(CONFIG.FLOW_HISTORY, { email: S.user.email });
    S.history = (r.items || []).sort((a, b) => new Date(b.submittedDate) - new Date(a.submittedDate));
  } catch (e) {
    toast(`Unable to load history: ${e.message}`, 'error');
  } finally {
    loading(false);
    renderHistory();
  }
}
function renderHistory() {
  const h = S.history;
  $('kpiTotal').textContent = h.length;
  $('kpiQty').textContent = h.reduce((s, x) => s + num(x.physicalQty), 0).toLocaleString('en-MY');
  $('kpiExc').textContent = h.filter(x => x.condition && x.condition !== 'Good').length;
  $('kpiVer').textContent = h.filter(x => x.status === 'Verified').length;

  const q = $('inpFilter').value.trim().toLowerCase();
  const rows = q ? h.filter(x => [x.assetCode, x.assetName, x.location].some(v => String(v || '').toLowerCase().includes(q))) : h;

  const list = $('historyList');
  list.textContent = '';
  if (!rows.length) { list.appendChild(el('div', 'empty', h.length ? 'No matching records.' : 'No submissions yet for this session.')); return; }
  rows.forEach(x => {
    const row = el('div', 'row');
    const main = el('div', 'row-main');
    const title = el('div', 'row-title');
    title.append(el('span', 'mono', x.assetCode), document.createTextNode(x.assetName || ''));
    main.append(title, el('div', 'row-sub', `${x.stockTakeId || ''} · ${x.location || ''} · ${fmtDT(x.submittedDate)}`));
    const right = el('div', 'row-right');
    right.append(el('span', 'row-qty', `Qty ${x.physicalQty ?? 0}`));
    const pills = el('div');
    if (x.condition && x.condition !== 'Good') pills.append(el('span', `pill ${x.condition}`, x.condition), document.createTextNode(' '));
    pills.append(el('span', `pill ${String(x.status || '').split(' ')[0]}`, x.status || 'Submitted'));
    right.append(pills);
    row.append(main, right);
    list.appendChild(row);
  });
}

/* ---------------- TABS ---------------- */
function switchTab(id) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.id === id));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === id));
  if (id !== 'tabCount') { stopCamera(); stopScan(); }
  if (id === 'tabHistory') loadHistory();
  if (id === 'tabQueue') renderQueue();
}

/* ---------------- EVENTS ---------------- */
function bind() {
  $('btnSendOtp').onclick = sendOtp;
  $('btnVerifyOtp').onclick = verifyOtp;
  $('btnBackEmail').onclick = () => { $('stepOtp').classList.add('hidden'); $('stepEmail').classList.remove('hidden'); };
  $('inpEmail').addEventListener('keydown', e => { if (e.key === 'Enter') sendOtp(); });
  $('inpOtp').addEventListener('input', e => {
    e.target.value = e.target.value.replace(/\D/g, '').slice(0, 6);
    if (e.target.value.length === 6) verifyOtp();
  });
  $('btnLogout').onclick = () => { if (confirm('Sign out of StockSense?')) logout(); };

  $('selLocation').onchange = () => { localStorage.setItem('ops_st_location', $('selLocation').value); $('selLocation').classList.remove('invalid'); };
  ['inpAssetCode', 'inpAssetName'].forEach(id => $(id).addEventListener('input', () => $(id).classList.remove('invalid')));
  $('inpAssetCode').addEventListener('blur', e => { e.target.value = e.target.value.trim().toUpperCase(); });

  $('btnScan').onclick = startScan;
  $('btnStopScan').onclick = stopScan;

  document.querySelectorAll('[data-delta]').forEach(b => b.addEventListener('click', () => {
    $('inpQty').value = Math.max(0, parseInt($('inpQty').value || 0, 10) + parseInt(b.dataset.delta, 10));
    buzz(8);
  }));
  document.querySelector('[data-reset]').onclick = () => { $('inpQty').value = 0; };
  $('inpQty').addEventListener('focus', e => e.target.select());
  document.querySelectorAll('.seg').forEach(s => { s.onclick = () => setCondition(s.dataset.cond); });

  $('btnOpenCam').onclick = openCamera;
  $('btnCamClose').onclick = stopCamera;
  $('btnCapture').onclick = capture;
  $('btnCamSwitch').onclick = switchCamera;
  $('btnPickPhoto').onclick = () => $('fileGallery').click();
  $('fileCamera').onchange = onFilePick;
  $('fileGallery').onchange = onFilePick;

  $('btnGeo').onclick = getGeo;
  $('btnSubmit').onclick = submitCount;
  $('btnClear').onclick = () => { if (confirm('Clear all entered details and photos?')) clearForm(); };

  $('btnRefreshHistory').onclick = loadHistory;
  $('inpFilter').addEventListener('input', renderHistory);
  $('btnSync').onclick = () => syncQueue(false);

  document.querySelectorAll('.tab-btn').forEach(b => { b.onclick = () => switchTab(b.dataset.tab); });

  window.addEventListener('online', () => { updateNet(); toast('Back online — syncing', 'success'); syncQueue(true); });
  window.addEventListener('offline', () => { updateNet(); toast('You are offline. Submissions will be queued.', 'warn'); });
  document.addEventListener('visibilitychange', () => { if (document.hidden) { stopCamera(); stopScan(); } });
}

/* ---------------- INIT ---------------- */
function init() {
  bind();
  renderPhotos();
  if (!CONFIG.AUTH_ENABLED) $('btnSendOtp').textContent = 'Continue';
  const u = loadUser();
  if (u) { S.user = u; enterApp(); }
}
init();
