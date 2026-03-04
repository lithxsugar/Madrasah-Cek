// ==================== CONSTANTS ====================
const DB_NAME = 'SurveiMadrasahDB';
const DB_VERSION = 1;
const STORE_NAME = 'surveys';

const FACILITIES = [
  { key: 'lapangan', label: 'Lapangan Olahraga', icon: '⚽' },
  { key: 'toilet', label: 'Toilet / WC', icon: '🚻', hasCount: true },
];

const MEUBELAIR_ITEMS = ['Meja Siswa', 'Kursi Siswa', 'Meja Guru', 'Kursi Guru', 'Papan Tulis', 'Lemari'];

const PHOTO_CATEGORIES = [
  { key: 'gedungDepan', label: 'Foto Tampak Depan Gedung', icon: '🏫' },
  { key: 'gedungSamping', label: 'Foto Tampak Samping Gedung', icon: '🏗️' },
  { key: 'gedungBelakang', label: 'Foto Tampak Belakang Gedung', icon: '🏘️' },
  { key: 'papanNama', label: 'Foto Papan Nama Sekolah', icon: '📋' },
  { key: 'kerusakan', label: 'Foto Detail Kerusakan', icon: '⚠️', hasDesc: true },
];

// ==================== STATE ====================
let db = null;
let currentSurveyId = null;
let surveyData = {};
let leafletMap = null;
let leafletMarker = null;

// ==================== IndexedDB ====================
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = e => {
      const d = e.target.result;
      if (!d.objectStoreNames.contains(STORE_NAME)) d.createObjectStore(STORE_NAME, { keyPath: 'id' });
    };
    req.onsuccess = e => { db = e.target.result; resolve(db); };
    req.onerror = e => reject(e.target.error);
  });
}
function dbPut(data) { return new Promise((res, rej) => { const tx = db.transaction(STORE_NAME, 'readwrite'); tx.objectStore(STORE_NAME).put(data); tx.oncomplete = () => res(); tx.onerror = e => rej(e.target.error); }); }
function dbGetAll() { return new Promise((res, rej) => { const tx = db.transaction(STORE_NAME, 'readonly'); const r = tx.objectStore(STORE_NAME).getAll(); r.onsuccess = () => res(r.result); r.onerror = e => rej(e.target.error); }); }
function dbGet(id) { return new Promise((res, rej) => { const tx = db.transaction(STORE_NAME, 'readonly'); const r = tx.objectStore(STORE_NAME).get(id); r.onsuccess = () => res(r.result); r.onerror = e => rej(e.target.error); }); }
function dbDelete(id) { return new Promise((res, rej) => { const tx = db.transaction(STORE_NAME, 'readwrite'); tx.objectStore(STORE_NAME).delete(id); tx.oncomplete = () => res(); tx.onerror = e => rej(e.target.error); }); }

// ==================== UTILITIES ====================
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
function showToast(msg) { const t = document.getElementById('toast'); t.textContent = msg; t.classList.add('show'); setTimeout(() => t.classList.remove('show'), 2500); }
function formatDate(iso) { if (!iso) return '-'; return new Date(iso).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }); }
function formatCurrency(n) { return new Intl.NumberFormat('id-ID').format(n || 0); }
function gv(id) { return document.getElementById(id)?.value || ''; }
function gvn(id) { return parseInt(document.getElementById(id)?.value) || 0; }
function sv(id, val) { const el = document.getElementById(id); if (el && val !== undefined && val !== null) el.value = val; }

function compressImage(file, maxW = 1200, quality = 0.7) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let w = img.width, h = img.height;
        if (w > maxW) { h = (maxW / w) * h; w = maxW; }
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

// ==================== TOGGLE FUNCTIONS ====================
function toggleLainnya(selectId, boxId) {
  const sel = document.getElementById(selectId);
  const box = document.getElementById(boxId);
  box.style.display = sel.value === 'lainnya' ? '' : 'none';
}

function toggleDetail(boxId) {
  const box = document.getElementById(boxId);
  box.style.display = box.style.display === 'none' ? '' : 'none';
}

function toggleFacility(key) {
  const chk = document.getElementById('fac-chk-' + key);
  const det = document.getElementById('fac-det-' + key);
  const item = document.getElementById('fac-' + key);
  det.style.display = chk.checked ? 'grid' : 'none';
  item.classList.toggle('checked', chk.checked);
}

// ==================== INIT ====================
async function init() {
  await openDB();
  buildFacilities();
  buildMeubelair();
  buildPhotoSections();
  bindEvents();
  showHome();
}

function bindEvents() {
  document.getElementById('newSurveyBtn').addEventListener('click', startNewSurvey);
  document.getElementById('backToHomeBtn').addEventListener('click', backToHome);
  document.getElementById('addBantuanBtn').addEventListener('click', () => addBantuanRow());
  document.getElementById('getGPSBtn').addEventListener('click', getGPS);
  document.getElementById('exportWordBtn').addEventListener('click', exportWord);
  document.getElementById('exportExcelBtn').addEventListener('click', exportExcel);
  document.getElementById('saveCloseBtn').addEventListener('click', async () => { await saveSurvey(); backToHome(); });
}

// ==================== BUILD UI ====================
function buildFacilities() {
  document.getElementById('kelengkapanList').innerHTML = FACILITIES.map(f => `
    <div class="facility-item" id="fac-${f.key}">
      <div class="facility-name"><input type="checkbox" id="fac-chk-${f.key}" onchange="toggleFacility('${f.key}')"><span>${f.icon} ${f.label}</span></div>
      <div class="facility-details" style="display:none" id="fac-det-${f.key}">
        ${f.hasCount ? `<label>Jumlah</label><input type="number" id="fac-cnt-${f.key}" min="0" placeholder="0" inputmode="numeric">` : ''}
        <label>Kondisi</label><select id="fac-kon-${f.key}"><option value="baik">Baik</option><option value="rusak_ringan">Rusak Ringan</option><option value="rusak_berat">Rusak Berat</option></select>
      </div>
    </div>`).join('');
}

function buildMeubelair() {
  document.getElementById('meubelairBody').innerHTML = MEUBELAIR_ITEMS.map((item, i) => `
    <tr><td style="font-weight:600">${item}</td>
    <td><input type="number" id="mbl-baik-${i}" min="0" placeholder="0" inputmode="numeric"></td>
    <td><input type="number" id="mbl-rr-${i}" min="0" placeholder="0" inputmode="numeric"></td>
    <td><input type="number" id="mbl-rb-${i}" min="0" placeholder="0" inputmode="numeric"></td></tr>`).join('');
}

function buildPhotoSections() {
  document.getElementById('photoSections').innerHTML = PHOTO_CATEGORIES.map(cat => `
    <div class="photo-category" id="cat-${cat.key}">
      <h3>${cat.icon} ${cat.label}</h3>
      <div class="photo-actions">
        <label class="btn-outline btn-sm" style="cursor:pointer">📷 Kamera<input type="file" accept="image/*" capture="environment" onchange="handlePhoto(event,'${cat.key}')" hidden></label>
        <label class="btn-outline btn-sm" style="cursor:pointer">📁 Pilih File<input type="file" accept="image/*" multiple onchange="handlePhoto(event,'${cat.key}')" hidden></label>
      </div>
      <div class="photo-grid" id="photos-${cat.key}"></div>
    </div>`).join('');
}

// ==================== NAVIGATION ====================
async function showHome() {
  document.getElementById('homeScreen').style.display = '';
  document.getElementById('surveyForm').style.display = 'none';
  document.getElementById('backToHomeBtn').style.display = 'none';
  document.getElementById('headerSubtitle').textContent = 'Pendataan Kerusakan Bangunan';
  await renderSurveyList();
}

function showForm() {
  document.getElementById('homeScreen').style.display = 'none';
  document.getElementById('surveyForm').style.display = '';
  document.getElementById('backToHomeBtn').style.display = '';
  initMap();
}

async function backToHome() { await saveSurvey(); showHome(); }

// ==================== SURVEY CRUD ====================
function startNewSurvey() {
  currentSurveyId = uid();
  surveyData = createEmptySurvey(currentSurveyId);
  populateForm(surveyData);
  document.getElementById('headerSubtitle').textContent = 'Survei Baru';
  showForm();
}

function createEmptySurvey(id) {
  return {
    id, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
    identitas: {}, lahan: {}, dataPokok: {}, kontak: {},
    bantuan: [], kelengkapan: {},
    meubelair: MEUBELAIR_ITEMS.map(j => ({ jenis: j, baik: 0, rusakRingan: 0, rusakBerat: 0 })),
    lembarBelakang: '', kesimpulanSurveyor: '', namaSurveyor: '', tanggalSurvei: '',
    fotos: { gedungDepan: [], gedungSamping: [], gedungBelakang: [], papanNama: [], kerusakan: [] },
    koordinat: { latitude: '', longitude: '' }
  };
}

async function openSurvey(id) {
  const data = await dbGet(id);
  if (!data) { showToast('Data tidak ditemukan'); return; }
  currentSurveyId = id;
  surveyData = data;
  populateForm(surveyData);
  document.getElementById('headerSubtitle').textContent = data.identitas?.namaMadrasah || 'Edit Survei';
  showForm();
}

async function deleteSurvey(id) {
  if (!confirm('Yakin ingin menghapus survei ini?')) return;
  await dbDelete(id); showToast('Survei dihapus'); renderSurveyList();
}

// ==================== COLLECT & POPULATE ====================
function collectFormData() {
  const d = surveyData;
  d.updatedAt = new Date().toISOString();
  d.identitas = { namaMadrasah: gv('namaMadrasah'), npsn: gv('npsn'), nsm: gv('nsm'), alamat: gv('alamat'), kecamatan: gv('kecamatan'), kabupaten: gv('kabupaten'), provinsi: gv('provinsi'), jenjang: gv('jenjang') };

  const sertVal = gv('statusSertifikat');
  const tanahVal = gv('statusTanah');
  d.lahan = {
    statusSertifikat: sertVal === 'lainnya' ? gv('statusSertifikatCustom') : sertVal,
    nomorSertifikat: gv('nomorSertifikat'), ketSertifikat: gv('ketSertifikat'),
    luasLahan: gv('luasLahan'), luasBangunan: gv('luasBangunan'),
    statusTanah: tanahVal === 'lainnya' ? gv('statusTanahCustom') : tanahVal,
    ketTanah: gv('ketTanah'), tahunBerdiri: gv('tahunBerdiri')
  };

  d.dataPokok = {
    jumlahSiswa: gvn('jumlahSiswa'), siswaL: gvn('siswaL'), siswaP: gvn('siswaP'),
    jumlahGuru: gvn('jumlahGuru'), guruPNS: gvn('guruPNS'), guruNonPNS: gvn('guruNonPNS'),
    rombel: gvn('rombel'), ruangKelas: gvn('ruangKelas'),
    ruangKelasBaik: gvn('ruangKelasBaik'), ruangKelasRusakRingan: gvn('ruangKelasRusakRingan'), ruangKelasRusakBerat: gvn('ruangKelasRusakBerat')
  };

  d.kontak = { namaKepsek: gv('namaKepsek'), hpKepsek: gv('hpKepsek'), namaGuru: gv('namaGuru'), hpGuru: gv('hpGuru') };
  d.bantuan = collectBantuan();

  d.kelengkapan = {};
  FACILITIES.forEach(f => {
    const chk = document.getElementById('fac-chk-' + f.key);
    d.kelengkapan[f.key] = { ada: chk.checked, kondisi: document.getElementById('fac-kon-' + f.key)?.value || '', jumlah: f.hasCount ? (parseInt(document.getElementById('fac-cnt-' + f.key)?.value) || 0) : null };
  });

  d.meubelair = MEUBELAIR_ITEMS.map((item, i) => ({ jenis: item, baik: parseInt(document.getElementById('mbl-baik-' + i)?.value) || 0, rusakRingan: parseInt(document.getElementById('mbl-rr-' + i)?.value) || 0, rusakBerat: parseInt(document.getElementById('mbl-rb-' + i)?.value) || 0 }));

  d.lembarBelakang = gv('lembarBelakang');
  d.kesimpulanSurveyor = gv('kesimpulanSurveyor');
  d.namaSurveyor = gv('namaSurveyor');
  d.tanggalSurvei = gv('tanggalSurvei');
  d.koordinat = { latitude: gv('latitude'), longitude: gv('longitude') };
  return d;
}

function populateForm(d) {
  sv('namaMadrasah', d.identitas?.namaMadrasah); sv('npsn', d.identitas?.npsn); sv('nsm', d.identitas?.nsm);
  sv('alamat', d.identitas?.alamat); sv('kecamatan', d.identitas?.kecamatan); sv('kabupaten', d.identitas?.kabupaten);
  sv('provinsi', d.identitas?.provinsi); sv('jenjang', d.identitas?.jenjang);

  // Lahan - check if custom value
  const sertOpts = ['', 'ada', 'proses', 'belum', 'lainnya'];
  const sertVal = d.lahan?.statusSertifikat || '';
  if (sertOpts.includes(sertVal)) { sv('statusSertifikat', sertVal); } else { sv('statusSertifikat', 'lainnya'); sv('statusSertifikatCustom', sertVal); toggleLainnya('statusSertifikat', 'statusSertifikatBox'); }
  sv('nomorSertifikat', d.lahan?.nomorSertifikat); sv('ketSertifikat', d.lahan?.ketSertifikat);
  sv('luasLahan', d.lahan?.luasLahan); sv('luasBangunan', d.lahan?.luasBangunan);
  const tanahOpts = ['', 'milik', 'wakaf', 'sewa', 'pemerintah', 'lainnya'];
  const tanahVal = d.lahan?.statusTanah || '';
  if (tanahOpts.includes(tanahVal)) { sv('statusTanah', tanahVal); } else { sv('statusTanah', 'lainnya'); sv('statusTanahCustom', tanahVal); toggleLainnya('statusTanah', 'statusTanahBox'); }
  sv('ketTanah', d.lahan?.ketTanah); sv('tahunBerdiri', d.lahan?.tahunBerdiri);

  sv('jumlahSiswa', d.dataPokok?.jumlahSiswa); sv('siswaL', d.dataPokok?.siswaL); sv('siswaP', d.dataPokok?.siswaP);
  if (d.dataPokok?.siswaL || d.dataPokok?.siswaP) { document.getElementById('showSiswaDetail').checked = true; toggleDetail('siswaDetailBox'); }
  sv('jumlahGuru', d.dataPokok?.jumlahGuru); sv('guruPNS', d.dataPokok?.guruPNS); sv('guruNonPNS', d.dataPokok?.guruNonPNS);
  if (d.dataPokok?.guruPNS || d.dataPokok?.guruNonPNS) { document.getElementById('showGuruDetail').checked = true; toggleDetail('guruDetailBox'); }
  sv('rombel', d.dataPokok?.rombel); sv('ruangKelas', d.dataPokok?.ruangKelas);
  sv('ruangKelasBaik', d.dataPokok?.ruangKelasBaik); sv('ruangKelasRusakRingan', d.dataPokok?.ruangKelasRusakRingan); sv('ruangKelasRusakBerat', d.dataPokok?.ruangKelasRusakBerat);

  sv('namaKepsek', d.kontak?.namaKepsek); sv('hpKepsek', d.kontak?.hpKepsek);
  sv('namaGuru', d.kontak?.namaGuru); sv('hpGuru', d.kontak?.hpGuru);

  populateBantuan(d.bantuan || []);

  FACILITIES.forEach(f => {
    const fac = d.kelengkapan?.[f.key];
    const chk = document.getElementById('fac-chk-' + f.key);
    if (chk && fac) {
      chk.checked = !!fac.ada; toggleFacility(f.key);
      if (fac.kondisi) document.getElementById('fac-kon-' + f.key).value = fac.kondisi;
      if (f.hasCount && fac.jumlah != null) document.getElementById('fac-cnt-' + f.key).value = fac.jumlah;
    }
  });

  (d.meubelair || []).forEach((m, i) => { sv('mbl-baik-' + i, m.baik); sv('mbl-rr-' + i, m.rusakRingan); sv('mbl-rb-' + i, m.rusakBerat); });

  sv('lembarBelakang', d.lembarBelakang); sv('kesimpulanSurveyor', d.kesimpulanSurveyor);
  sv('namaSurveyor', d.namaSurveyor); sv('tanggalSurvei', d.tanggalSurvei);

  PHOTO_CATEGORIES.forEach(cat => renderPhotos(cat.key));
  sv('latitude', d.koordinat?.latitude); sv('longitude', d.koordinat?.longitude);
}

// ==================== BANTUAN ====================
function addBantuanRow(data = {}) {
  const tbody = document.getElementById('bantuanBody');
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input type="number" class="b-tahun" value="${data.tahun || ''}" placeholder="2024" inputmode="numeric"></td>
    <td><input type="text" class="b-sumber" value="${data.sumber || ''}" placeholder="BOS/DAK/dll"></td>
    <td><input type="text" class="b-jenis" value="${data.jenis || ''}" placeholder="Rehabilitasi/dll"></td>
    <td><input type="number" class="b-nominal" value="${data.nominal || ''}" placeholder="0" inputmode="numeric"></td>
    <td><button type="button" class="btn-danger" onclick="this.closest('tr').remove()">✕</button></td>`;
  tbody.appendChild(tr);
}
function collectBantuan() {
  return Array.from(document.querySelectorAll('#bantuanBody tr')).map(r => ({
    tahun: r.querySelector('.b-tahun')?.value || '', sumber: r.querySelector('.b-sumber')?.value || '',
    jenis: r.querySelector('.b-jenis')?.value || '', nominal: parseInt(r.querySelector('.b-nominal')?.value) || 0
  })).filter(b => b.tahun || b.sumber || b.jenis || b.nominal);
}
function populateBantuan(list) { document.getElementById('bantuanBody').innerHTML = ''; list.forEach(b => addBantuanRow(b)); }

// ==================== PHOTOS ====================
async function handlePhoto(event, category) {
  const files = event.target.files; if (!files.length) return;
  for (const file of files) { const dataUrl = await compressImage(file); surveyData.fotos[category].push({ dataUrl, timestamp: new Date().toISOString(), desc: '' }); }
  renderPhotos(category); event.target.value = ''; showToast(`${files.length} foto ditambahkan`);
}
function deletePhoto(category, index) { surveyData.fotos[category].splice(index, 1); renderPhotos(category); }
function updatePhotoDesc(category, index, desc) { if (surveyData.fotos[category][index]) surveyData.fotos[category][index].desc = desc; }
function renderPhotos(category) {
  const grid = document.getElementById('photos-' + category); if (!grid) return;
  const photos = surveyData.fotos?.[category] || [];
  const cat = PHOTO_CATEGORIES.find(c => c.key === category);
  grid.innerHTML = photos.map((p, i) => `<div><div class="photo-thumb"><img src="${p.dataUrl}" alt="Foto ${i + 1}" loading="lazy"><button class="photo-delete" onclick="deletePhoto('${category}',${i})">✕</button></div>${cat?.hasDesc ? `<input class="photo-desc-input" placeholder="Keterangan..." value="${p.desc || ''}" onchange="updatePhotoDesc('${category}',${i},this.value)">` : ''}</div>`).join('');
}

// ==================== GPS ====================
function getGPS() {
  const status = document.getElementById('gpsStatus');
  if (!navigator.geolocation) { status.textContent = '❌ Geolocation tidak didukung'; return; }
  status.textContent = '📡 Mengambil koordinat GPS...';
  navigator.geolocation.getCurrentPosition(pos => {
    const lat = pos.coords.latitude.toFixed(6), lng = pos.coords.longitude.toFixed(6);
    document.getElementById('latitude').value = lat; document.getElementById('longitude').value = lng;
    status.textContent = `✅ Koordinat berhasil (akurasi: ${Math.round(pos.coords.accuracy)}m)`;
    surveyData.koordinat = { latitude: lat, longitude: lng };
    updateMap(parseFloat(lat), parseFloat(lng));
  }, err => { status.textContent = `❌ Gagal: ${err.message}`; }, { enableHighAccuracy: true, timeout: 15000 });
}

function initMap() {
  setTimeout(() => {
    const c = document.getElementById('mapContainer');
    if (!c || c.dataset.init === '1') { if (leafletMap) leafletMap.invalidateSize(); return; }
    const lat = parseFloat(gv('latitude')) || -6.2, lng = parseFloat(gv('longitude')) || 106.8;
    try { leafletMap = L.map(c).setView([lat, lng], 15); L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { attribution: '© OpenStreetMap' }).addTo(leafletMap); leafletMarker = L.marker([lat, lng]).addTo(leafletMap); c.dataset.init = '1'; } catch (e) { }
  }, 500);
}
function updateMap(lat, lng) { if (leafletMap && leafletMarker) { leafletMap.setView([lat, lng], 16); leafletMarker.setLatLng([lat, lng]); } }

// ==================== SAVE ====================
async function saveSurvey() { collectFormData(); await dbPut(surveyData); showToast('✅ Data tersimpan'); }

// ==================== SURVEY LIST ====================
async function renderSurveyList() {
  const list = document.getElementById('surveyList');
  const surveys = await dbGetAll();
  if (!surveys.length) { list.innerHTML = '<p style="text-align:center;color:var(--muted);padding:20px">Belum ada data survei.</p>'; return; }
  surveys.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
  list.innerHTML = surveys.map(s => `<div class="survey-card" onclick="openSurvey('${s.id}')"><div><h3>${s.identitas?.namaMadrasah || 'Tanpa Nama'}</h3><p class="meta">${s.identitas?.alamat || '-'}</p><p class="meta">Diperbarui: ${formatDate(s.updatedAt)}</p></div><div class="survey-card-actions" onclick="event.stopPropagation()"><button class="btn-outline btn-sm" onclick="exportWordById('${s.id}')">📄 Word</button><button class="btn-outline btn-sm" onclick="exportExcelById('${s.id}')">📊 Excel</button><button class="btn-danger" onclick="deleteSurvey('${s.id}')">🗑️ Hapus</button></div></div>`).join('');
}
async function exportWordById(id) { const d = await dbGet(id); if (d) { surveyData = d; exportWord(); } }
async function exportExcelById(id) { const d = await dbGet(id); if (d) { surveyData = d; exportExcel(); } }

// ==================== EXPORT WORD ====================
async function exportWord() {
  try {
    collectFormData();
    const d = surveyData, nama = d.identitas?.namaMadrasah || 'Madrasah';
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, WidthType, BorderStyle, ImageRun } = docx;
    const brd = { top: { style: BorderStyle.SINGLE, size: 1, color: '000000' }, bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' }, left: { style: BorderStyle.SINGLE, size: 1, color: '000000' }, right: { style: BorderStyle.SINGLE, size: 1, color: '000000' } };

    function c(text, opts = {}) { return new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(text || '-'), bold: opts.bold, size: opts.size || 22, font: 'Calibri' })], alignment: opts.align, spacing: { before: 40, after: 40 } })], borders: brd, width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined, shading: opts.shading ? { fill: opts.shading } : undefined, columnSpan: opts.colSpan, verticalAlign: docx.VerticalAlign?.CENTER }); }
    function hc(text, opts = {}) { return c(text, { bold: true, ...opts }); }
    function row(label, value) { return new TableRow({ children: [hc(label, { width: 40 }), c(value, { width: 60 })] }); }

    const ch = [];
    ch.push(new Paragraph({ children: [new TextRun({ text: 'LAPORAN SURVEI KERUSAKAN MADRASAH', bold: true, size: 32, font: 'Calibri' })], alignment: AlignmentType.CENTER, spacing: { after: 100 } }));
    ch.push(new Paragraph({ children: [new TextRun({ text: nama.toUpperCase(), bold: true, size: 28, font: 'Calibri' })], alignment: AlignmentType.CENTER, spacing: { after: 60 } }));
    if (d.tanggalSurvei) ch.push(new Paragraph({ children: [new TextRun({ text: 'Tanggal Survei: ' + d.tanggalSurvei, size: 22, font: 'Calibri' })], alignment: AlignmentType.CENTER, spacing: { after: 300 } }));

    // A. Identitas
    ch.push(new Paragraph({ text: 'A. IDENTITAS MADRASAH', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    ch.push(new Table({ rows: [row('Nama Madrasah', nama), row('NPSN', d.identitas?.npsn), row('NSM', d.identitas?.nsm), row('Jenjang', d.identitas?.jenjang), row('Alamat', d.identitas?.alamat), row('Kecamatan', d.identitas?.kecamatan), row('Kabupaten/Kota', d.identitas?.kabupaten), row('Provinsi', d.identitas?.provinsi)], width: { size: 100, type: WidthType.PERCENTAGE } }));

    // B. Lahan
    const lahanLabels = { ada: 'Ada/Bersertifikat', proses: 'Dalam Proses', belum: 'Belum Ada', milik: 'Milik Sendiri', wakaf: 'Wakaf', sewa: 'Sewa/Pinjam', pemerintah: 'Milik Pemerintah' };
    ch.push(new Paragraph({ text: 'B. DATA LAHAN', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    const lahanRows = [row('Status Sertifikat', lahanLabels[d.lahan?.statusSertifikat] || d.lahan?.statusSertifikat), row('Nomor Sertifikat', d.lahan?.nomorSertifikat)];
    if (d.lahan?.ketSertifikat) lahanRows.push(row('Keterangan Sertifikat', d.lahan.ketSertifikat));
    lahanRows.push(row('Luas Lahan (m²)', d.lahan?.luasLahan), row('Luas Bangunan (m²)', d.lahan?.luasBangunan), row('Status Kepemilikan', lahanLabels[d.lahan?.statusTanah] || d.lahan?.statusTanah));
    if (d.lahan?.ketTanah) lahanRows.push(row('Keterangan Kepemilikan', d.lahan.ketTanah));
    lahanRows.push(row('Tahun Berdiri', d.lahan?.tahunBerdiri));
    ch.push(new Table({ rows: lahanRows, width: { size: 100, type: WidthType.PERCENTAGE } }));

    // C. Data Pokok
    ch.push(new Paragraph({ text: 'C. DATA POKOK', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    const dpRows = [row('Jumlah Siswa', d.dataPokok?.jumlahSiswa)];
    if (d.dataPokok?.siswaL || d.dataPokok?.siswaP) dpRows.push(row('  - Laki-laki', d.dataPokok.siswaL), row('  - Perempuan', d.dataPokok.siswaP));
    dpRows.push(row('Jumlah Guru', d.dataPokok?.jumlahGuru));
    if (d.dataPokok?.guruPNS || d.dataPokok?.guruNonPNS) dpRows.push(row('  - PNS', d.dataPokok.guruPNS), row('  - Non-PNS', d.dataPokok.guruNonPNS));
    dpRows.push(row('Jumlah Rombel', d.dataPokok?.rombel), row('Jumlah Ruang Kelas', d.dataPokok?.ruangKelas), row('Ruang Kelas Baik', d.dataPokok?.ruangKelasBaik), row('Rusak Ringan', d.dataPokok?.ruangKelasRusakRingan), row('Rusak Berat', d.dataPokok?.ruangKelasRusakBerat));
    ch.push(new Table({ rows: dpRows, width: { size: 100, type: WidthType.PERCENTAGE } }));

    // D. Kontak
    ch.push(new Paragraph({ text: 'D. DATA KONTAK', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    ch.push(new Table({ rows: [row('Nama Kepala Sekolah', d.kontak?.namaKepsek), row('No. HP Kepsek', d.kontak?.hpKepsek), row('Nama Guru Pendamping', d.kontak?.namaGuru), row('No. HP Guru', d.kontak?.hpGuru)], width: { size: 100, type: WidthType.PERCENTAGE } }));

    // E. Bantuan
    if ((d.bantuan || []).length) {
      ch.push(new Paragraph({ text: 'E. BANTUAN YANG SUDAH MASUK', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
      const bRows = [new TableRow({ children: [hc('No'), hc('Tahun'), hc('Sumber'), hc('Jenis'), hc('Nominal (Rp)')] })];
      d.bantuan.forEach((b, i) => bRows.push(new TableRow({ children: [c(i + 1), c(b.tahun), c(b.sumber), c(b.jenis), c(formatCurrency(b.nominal))] })));
      ch.push(new Table({ rows: bRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
    }

    // F. Kelengkapan
    ch.push(new Paragraph({ text: 'F. KELENGKAPAN SEKOLAH', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    const kondisiMap = { baik: 'Baik', rusak_ringan: 'Rusak Ringan', rusak_berat: 'Rusak Berat' };
    const kRows = [new TableRow({ children: [hc('No'), hc('Fasilitas'), hc('Tersedia'), hc('Kondisi'), hc('Jumlah')] })];
    FACILITIES.forEach((f, i) => { const fac = d.kelengkapan?.[f.key] || {}; kRows.push(new TableRow({ children: [c(i + 1), c(f.label), c(fac.ada ? 'Ya' : 'Tidak'), c(fac.ada ? (kondisiMap[fac.kondisi] || '-') : '-'), c(fac.jumlah != null && fac.ada ? fac.jumlah : '-')] })); });
    ch.push(new Table({ rows: kRows, width: { size: 100, type: WidthType.PERCENTAGE } }));

    // G. Meubelair
    ch.push(new Paragraph({ text: 'G. DATA MEUBELAIR', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    const mRows = [new TableRow({ children: [hc('No'), hc('Jenis'), hc('Baik'), hc('Rusak Ringan'), hc('Rusak Berat'), hc('Total')] })];
    (d.meubelair || []).forEach((m, i) => { const t = (m.baik || 0) + (m.rusakRingan || 0) + (m.rusakBerat || 0); mRows.push(new TableRow({ children: [c(i + 1), c(m.jenis), c(m.baik), c(m.rusakRingan), c(m.rusakBerat), c(t)] })); });
    ch.push(new Table({ rows: mRows, width: { size: 100, type: WidthType.PERCENTAGE } }));

    // H. Catatan
    ch.push(new Paragraph({ text: 'H. CATATAN / LEMBAR BELAKANG', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    ch.push(new Paragraph({ children: [new TextRun({ text: d.lembarBelakang || '-', size: 22, font: 'Calibri' })], spacing: { after: 100 } }));
    if (d.kesimpulanSurveyor) {
      ch.push(new Paragraph({ children: [new TextRun({ text: 'Kesimpulan Surveyor:', bold: true, size: 22, font: 'Calibri' })], spacing: { before: 100 } }));
      ch.push(new Paragraph({ children: [new TextRun({ text: d.kesimpulanSurveyor, size: 22, font: 'Calibri' })] }));
    }

    // I. Koordinat
    ch.push(new Paragraph({ text: 'I. TITIK KOORDINAT', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
    ch.push(new Paragraph({ children: [new TextRun({ text: `Latitude: ${d.koordinat?.latitude || '-'}, Longitude: ${d.koordinat?.longitude || '-'}`, size: 22, font: 'Calibri' })] }));

    // J. Foto
    const allPhotos = [];
    PHOTO_CATEGORIES.forEach(cat => (d.fotos?.[cat.key] || []).forEach((p, i) => allPhotos.push({ ...p, catLabel: cat.label, index: i })));
    if (allPhotos.length) {
      ch.push(new Paragraph({ text: 'J. DOKUMENTASI FOTO', heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
      for (const photo of allPhotos) {
        try {
          const bin = atob(photo.dataUrl.split(',')[1]); const arr = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
          ch.push(new Paragraph({ children: [new TextRun({ text: `${photo.catLabel}${photo.desc ? ' - ' + photo.desc : ''}`, bold: true, size: 20, font: 'Calibri' })], spacing: { before: 200 } }));
          ch.push(new Paragraph({ children: [new ImageRun({ data: arr, transformation: { width: 450, height: 340 }, type: 'jpg' })], spacing: { after: 100 } }));
        } catch (e) { console.warn('Foto skip:', e); }
      }
    }

    // Signature
    ch.push(new Paragraph({ text: '', spacing: { before: 600 } }));
    ch.push(new Paragraph({ children: [new TextRun({ text: 'Surveyor,', size: 22, font: 'Calibri' })], alignment: AlignmentType.RIGHT, spacing: { before: 200 } }));
    ch.push(new Paragraph({ text: '', spacing: { before: 600 } }));
    ch.push(new Paragraph({ children: [new TextRun({ text: d.namaSurveyor || '(........................)', size: 22, underline: {}, font: 'Calibri' })], alignment: AlignmentType.RIGHT }));

    const doc = new Document({ sections: [{ properties: { page: { margin: { top: 1000, bottom: 1000, left: 1200, right: 1000 } } }, children: ch }] });
    const blob = await Packer.toBlob(doc);
    const fn = prompt('Nama file Word:', `Survei_${nama.replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}`);
    if (fn) saveAs(blob, fn + '.docx');
    showToast('✅ Word berhasil di-export');
  } catch (e) { console.error('Word error:', e); showToast('❌ Gagal export Word: ' + e.message); }
}

// ==================== EXPORT EXCEL ====================
async function exportExcel() {
  try {
    collectFormData();
    const d = surveyData, nama = d.identitas?.namaMadrasah || 'Madrasah';
    const wb = new ExcelJS.Workbook(); wb.creator = 'Survei Madrasah PWA';
    const bdr = () => ({ top: { style: 'thin', color: { argb: 'FF94A3B8' } }, bottom: { style: 'thin', color: { argb: 'FF94A3B8' } }, left: { style: 'thin', color: { argb: 'FF94A3B8' } }, right: { style: 'thin', color: { argb: 'FF94A3B8' } } });
    const hStyle = { font: { bold: true, color: { argb: 'FF065F46' }, size: 11 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } }, border: bdr(), alignment: { vertical: 'middle', wrapText: true } };

    // Sheet 1: Data Pokok
    const ws1 = wb.addWorksheet('Data Pokok');
    ws1.columns = [{ width: 32 }, { width: 42 }];

    function addTitle(ws, text, row) {
      ws.mergeCells(row, 1, row, 2);
      const r = ws.getRow(row); r.getCell(1).value = text;
      r.getCell(1).font = { bold: true, size: 14, color: { argb: 'FF065F46' } }; r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } }; r.height = 28;
    }

    function addSection(ws, text, rowNum) {
      ws.mergeCells(rowNum, 1, rowNum, 2);
      const r = ws.getRow(rowNum); r.getCell(1).value = text;
      r.getCell(1).font = { bold: true, size: 12, color: { argb: 'FF047857' } }; r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
      r.getCell(1).border = bdr(); r.height = 24;
    }

    function addPair(ws, label, value, rowNum) {
      const r = ws.getRow(rowNum);
      r.getCell(1).value = label; r.getCell(1).font = { bold: true, size: 10 }; r.getCell(1).border = bdr(); r.getCell(1).alignment = { vertical: 'middle', wrapText: true };
      r.getCell(2).value = value ?? '-'; r.getCell(2).border = bdr(); r.getCell(2).alignment = { vertical: 'middle', wrapText: true };
      r.height = 20;
    }

    addTitle(ws1, 'LAPORAN SURVEI KERUSAKAN MADRASAH', 1);
    addTitle(ws1, nama.toUpperCase(), 2);

    let rn = 4;
    addSection(ws1, 'IDENTITAS MADRASAH', rn++);
    const idPairs = [['Nama Madrasah', nama], ['NPSN', d.identitas?.npsn], ['NSM', d.identitas?.nsm], ['Jenjang', d.identitas?.jenjang], ['Alamat', d.identitas?.alamat], ['Kecamatan', d.identitas?.kecamatan], ['Kabupaten/Kota', d.identitas?.kabupaten], ['Provinsi', d.identitas?.provinsi]];
    idPairs.forEach(p => addPair(ws1, p[0], p[1], rn++));

    rn++;
    const lahanLabels = { ada: 'Ada/Bersertifikat', proses: 'Dalam Proses', belum: 'Belum Ada', milik: 'Milik Sendiri', wakaf: 'Wakaf', sewa: 'Sewa/Pinjam', pemerintah: 'Milik Pemerintah' };
    addSection(ws1, 'DATA LAHAN', rn++);
    [['Status Sertifikat', lahanLabels[d.lahan?.statusSertifikat] || d.lahan?.statusSertifikat], ['Nomor Sertifikat', d.lahan?.nomorSertifikat], ['Keterangan Sertifikat', d.lahan?.ketSertifikat], ['Luas Lahan (m²)', d.lahan?.luasLahan], ['Luas Bangunan (m²)', d.lahan?.luasBangunan], ['Status Kepemilikan', lahanLabels[d.lahan?.statusTanah] || d.lahan?.statusTanah], ['Keterangan Kepemilikan', d.lahan?.ketTanah], ['Tahun Berdiri', d.lahan?.tahunBerdiri]].forEach(p => addPair(ws1, p[0], p[1], rn++));

    rn++;
    addSection(ws1, 'DATA POKOK', rn++);
    const dpPairs = [['Jumlah Siswa', d.dataPokok?.jumlahSiswa]];
    if (d.dataPokok?.siswaL || d.dataPokok?.siswaP) dpPairs.push(['  - Laki-laki', d.dataPokok.siswaL], ['  - Perempuan', d.dataPokok.siswaP]);
    dpPairs.push(['Jumlah Guru', d.dataPokok?.jumlahGuru]);
    if (d.dataPokok?.guruPNS || d.dataPokok?.guruNonPNS) dpPairs.push(['  - PNS', d.dataPokok.guruPNS], ['  - Non-PNS', d.dataPokok.guruNonPNS]);
    dpPairs.push(['Jumlah Rombel', d.dataPokok?.rombel], ['Jumlah Ruang Kelas', d.dataPokok?.ruangKelas], ['R. Kelas Baik', d.dataPokok?.ruangKelasBaik], ['R. Kelas Rusak Ringan', d.dataPokok?.ruangKelasRusakRingan], ['R. Kelas Rusak Berat', d.dataPokok?.ruangKelasRusakBerat]);
    dpPairs.forEach(p => addPair(ws1, p[0], p[1], rn++));

    rn++;
    addSection(ws1, 'KONTAK', rn++);
    [['Nama Kepala Sekolah', d.kontak?.namaKepsek], ['HP Kepsek', d.kontak?.hpKepsek], ['Nama Guru Pendamping', d.kontak?.namaGuru], ['HP Guru', d.kontak?.hpGuru]].forEach(p => addPair(ws1, p[0], p[1], rn++));

    rn++;
    addSection(ws1, 'KOORDINAT', rn++);
    addPair(ws1, 'Latitude', d.koordinat?.latitude, rn++);
    addPair(ws1, 'Longitude', d.koordinat?.longitude, rn++);

    rn++;
    addSection(ws1, 'CATATAN', rn++);
    addPair(ws1, 'Surveyor', d.namaSurveyor, rn++);
    addPair(ws1, 'Tanggal', d.tanggalSurvei, rn++);
    addPair(ws1, 'Catatan', d.lembarBelakang, rn++);
    addPair(ws1, 'Kesimpulan', d.kesimpulanSurveyor, rn++);

    // Sheet 2: Bantuan
    if ((d.bantuan || []).length) {
      const ws2 = wb.addWorksheet('Bantuan');
      ws2.columns = [{ width: 6 }, { width: 10 }, { width: 25 }, { width: 25 }, { width: 20 }];
      const hRow = ws2.addRow(['No', 'Tahun', 'Sumber', 'Jenis Bantuan', 'Nominal (Rp)']);
      hRow.eachCell(c => { c.font = hStyle.font; c.fill = hStyle.fill; c.border = bdr(); c.alignment = { vertical: 'middle' }; });
      d.bantuan.forEach((b, i) => {
        const r = ws2.addRow([i + 1, b.tahun, b.sumber, b.jenis, b.nominal || 0]);
        r.eachCell(c => { c.border = bdr(); c.alignment = { vertical: 'middle' }; });
      });
    }

    // Sheet 3: Meubelair
    const ws3 = wb.addWorksheet('Meubelair');
    ws3.columns = [{ width: 6 }, { width: 20 }, { width: 12 }, { width: 14 }, { width: 14 }, { width: 10 }];
    const mhRow = ws3.addRow(['No', 'Jenis', 'Baik', 'Rusak Ringan', 'Rusak Berat', 'Total']);
    mhRow.eachCell(c => { c.font = hStyle.font; c.fill = hStyle.fill; c.border = bdr(); c.alignment = { vertical: 'middle' }; });
    (d.meubelair || []).forEach((m, i) => {
      const t = (m.baik || 0) + (m.rusakRingan || 0) + (m.rusakBerat || 0);
      const r = ws3.addRow([i + 1, m.jenis, m.baik, m.rusakRingan, m.rusakBerat, t]);
      r.eachCell(c => { c.border = bdr(); c.alignment = { vertical: 'middle' }; });
    });

    // Sheet 4: Kelengkapan
    const ws4 = wb.addWorksheet('Kelengkapan');
    ws4.columns = [{ width: 6 }, { width: 25 }, { width: 12 }, { width: 16 }, { width: 10 }];
    const khRow = ws4.addRow(['No', 'Fasilitas', 'Tersedia', 'Kondisi', 'Jumlah']);
    khRow.eachCell(c => { c.font = hStyle.font; c.fill = hStyle.fill; c.border = bdr(); c.alignment = { vertical: 'middle' }; });
    const kondisiMap = { baik: 'Baik', rusak_ringan: 'Rusak Ringan', rusak_berat: 'Rusak Berat' };
    FACILITIES.forEach((f, i) => {
      const fac = d.kelengkapan?.[f.key] || {};
      const r = ws4.addRow([i + 1, f.label, fac.ada ? 'Ya' : 'Tidak', fac.ada ? (kondisiMap[fac.kondisi] || '-') : '-', fac.jumlah ?? '-']);
      r.eachCell(c => { c.border = bdr(); c.alignment = { vertical: 'middle' }; });
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const fn = prompt('Nama file Excel:', `Survei_${nama.replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}`);
    if (fn) saveAs(blob, fn + '.xlsx');
    showToast('✅ Excel berhasil di-export');
  } catch (e) { console.error('Excel error:', e); showToast('❌ Gagal export Excel: ' + e.message); }
}

// ==================== INIT ====================
init();
